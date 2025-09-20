"""
Functions to handle PDF files:
- download missing pdfs from canvas
- extract text from pdfs
- deduplicate pdfs
- store extracted text
- ...
"""

from pathlib import Path

import numpy as np
from loguru import logger

from easy_access.settings import SETTINGS, DirSetting

TIMEOUT = 20  # set timeout for functions that might hang, e.g. text extraction
pdf_dir = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]


class TimeoutException(Exception):  # Custom exception class
    pass


def timeout_handler(signum, frame):  # Custom signal handler
    raise TimeoutException


def ocr_with_paddle(
    pdf_dir: Path,
    PAGE_NUM: int | None = None,
    outputdir: Path | None = None,
    suffix: str = "_paddleocr",
):
    """
    This function uses PaddleOCR to perform OCR on the given PDF files.
    It converts each page of the PDF to an image, processes the image with PaddleOCR,
    and saves the extracted text to a .txt file.

    Parameters:
        - PAGE_NUM: The number of pages to process from each PDF file. Default is 15.
        - pdfs: a generator that returns Paths for the PDF files to parse. Use Path().glob(*.pdf) for a dir of pdf files for example.
        - outputdir: The directory where the output text files will be saved. Default is the current working directory + /paddle_output.
        - suffix: The suffix to add to the output text files. Default is "_paddleocr". Will create files like {pdf_name}_paddleocr.txt.

    Requires a CUDA compatible gpu for PaddleOCR to work at a decent speed.
    """
    import cv2
    import fitz
    from paddleocr import PaddleOCR
    from PIL import Image

    if not PAGE_NUM:
        PAGE_NUM = 15
    if not pdf_dir or not pdf_dir.exists():
        pdf_dir = Path().cwd() / "pdfs"
    if not pdf_dir.exists():
        logger.warning(f"[red]Directory {pdf_dir} does not exist.[/red]")
        return
    if not outputdir or not outputdir.exists():
        outputdir = Path().cwd() / "paddle_output"

    pdfs = pdf_dir.glob("*.pdf")
    ocr = PaddleOCR(use_angle_cls=True, lang="en", page_num=PAGE_NUM, use_gpu=True)

    for pdf in pdfs:
        pdf_path = pdf
        pdf_name = pdf_path.stem

        if (outputdir / "{pdf_name}_paddle.txt").exists():
            logger.warning(f"[red]skipping {pdf_name}[/red]")
            continue
        imgs: list[Image.Image] = []
        full_text = []
        try:
            with fitz.open(pdf_path) as pdf:
                for pg in range(0, PAGE_NUM):
                    try:
                        page = pdf[pg]
                        mat = fitz.Matrix(2, 2)
                        pm = page.get_pixmap(matrix=mat, alpha=False)
                        if pm.width > 2000 or pm.height > 2000:
                            pm = page.get_pixmap(matrix=fitz.Matrix(1, 1), alpha=False)
                        img = Image.frombytes("RGB", [pm.width, pm.height], pm.samples)
                        img = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
                        imgs.append(img)
                    except Exception:
                        continue
        except Exception as e:
            logger.error(f"[red]Error processing {pdf_name}: {e}[/red]")
            continue
        if not imgs:
            continue
        for img in imgs:
            result = ocr.ocr(img, cls=True)

            if result is None:
                continue

            for residx in range(len(result)):
                res = result[residx]
                if res is None:
                    continue

                txts = [line[1][0] for line in res]
                full_text.extend(txts)

        full_text = " ".join(full_text)
        full_text = full_text.replace("  ", " ")
        with open(
            outputdir / (f"{pdf_name}" + suffix + ".txt"), "w", encoding="utf-8"
        ) as f:
            f.write(full_text)
