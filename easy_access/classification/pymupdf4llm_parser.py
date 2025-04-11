import asyncio

import pymupdf
import pymupdf4llm
from kreuzberg import ExtractionConfig, PSMMode, TesseractConfig, extract_file

from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import Directory, File

pdf_dir = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]


def extract_md(pdf) -> str:
    md_text = pymupdf4llm.to_markdown(
        pdf,
        pages=range(0, 15),
        write_images=False,
        ignore_graphics=True,
        ignore_code=True,
        force_text=True,
    )
    return md_text


async def extract_text_with_ocr(pdf: File):
    result = await extract_file(
        pdf.path,
        mime_type="application/pdf",
        config=ExtractionConfig(
            force_ocr=True,
            ocr_config=TesseractConfig(language="eng+nl", psm=PSMMode.AUTO),
        ),
    )
    return result


async def extract_list_with_ocr(pdfs: list[File]):
    for counter, pdf in enumerate(pdfs):
        try:
            print(f"Processing {counter + 1}/{len(pdfs)}")
            result = await extract_text_with_ocr(pdf)
            if result.content:
                if result.mime_type == "text/markdown":
                    extension = ".md"
                elif result.mime_type == "text/plain":
                    extension = ".txt"
                else:
                    extension = ".txt"
                filename: str = pdf.name.replace(".pdf", extension)
                save_path = (
                    SETTINGS.dirs[DirSetting.SCRIPT_DATA].full
                    / "parsed_pdf_text"
                    / filename
                )
                print(f"Saving {filename} to {save_path}")
                with open(save_path, "w", encoding="utf-8") as f:
                    f.write(result.content)
        except Exception as e:
            print(f"Error processing {pdf}: {e}")


def parse_pdfs():
    all_pdfs = [f for f in pdf_dir.files if f.extension == ".pdf"]
    existing_parsed_files = [
        f
        for f in Directory(
            SETTINGS.dirs[DirSetting.SCRIPT_DATA].full / "parsed_pdf_text"
        ).files
        if f.extension in [".md", ".txt"]
    ]
    pdf_ids = [pdf.name.replace(".pdf", "") for pdf in all_pdfs]
    existing_ids = [
        f.name.replace(".md", "").replace(".txt", "") for f in existing_parsed_files
    ]
    pdf_ids = list(set(pdf_ids) - set(existing_ids))

    pdfs = [pdf for pdf in all_pdfs if pdf.name.replace(".pdf", "") in pdf_ids]
    retry_list = []
    for counter, pdf in enumerate(pdfs):
        try:
            print(f"Processing {counter + 1}/{len(pdfs)}")
            print(pdf.path)
            doc = pymupdf.open(pdf.path)
            text = extract_md(doc)
            if not text:
                retry_list.append(pdf)

            markdown_file: str = pdf.name.replace(".pdf", ".md")
            save_path = (
                SETTINGS.dirs[DirSetting.SCRIPT_DATA].full
                / "parsed_pdf_text"
                / markdown_file
            )
            print(f"Saving {markdown_file} to {save_path}")
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(text)
        except Exception as e:
            print(f"Error processing {pdf}: {e}")
            retry_list.append(pdf)

    if retry_list:
        print(f"Trying OCR for {len(retry_list)} files that failed to parse initially.")
        asyncio.run(extract_list_with_ocr(retry_list))
