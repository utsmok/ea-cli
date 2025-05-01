"""
Functions to handle PDF files:
- download missing pdfs from canvas
- extract text from pdfs
- deduplicate pdfs
- store extracted text
- ...
"""

import asyncio
import contextlib
import traceback
from collections import defaultdict
from itertools import batched
from pathlib import Path
from types import CoroutineType

import kreuzberg
import numpy as np
import pikepdf
from kreuzberg import ExtractionResult, batch_extract_file, extract_file
from pydantic import BaseModel
from tortoise.queryset import QuerySet
from xxhash import xxh64

from easy_access.db.base import init
from easy_access.db.ingest import load_pdfs
from easy_access.db.models import PDF
from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import File, cool, info, warn

TIMEOUT = 20  # set timeout for functions that might hang, e.g. text extraction
pdf_dir = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]


class TimeoutException(Exception):  # Custom exception class
    pass


def timeout_handler(signum, frame):  # Custom signal handler
    raise TimeoutException


async def batch_extract_pdf_text(
    pdfs: list[PDF], max_pages: int | None = 15, str_limit: int | None = 50000
) -> list[PDF]:
    pdfs = [p for p in pdfs if p.path.exists() and not p.extracted_text]
    total = 0
    for batch in batched(pdfs, 20):
        total += len(batch)
        file_paths = [pdf.path for pdf in batch]
        results = await batch_extract_file(file_paths)
        for pdf, result in zip(batch, results, strict=False):
            content = result.content
            metadata = result.metadata
            if content:
                if len(content) > str_limit:
                    content = content[:str_limit]
                pdf.extracted_text = content
                pdf.extracted_text_max_length = str_limit
                # write extracted text to file
                with open(pdf_dir.full / f"{pdf.material_id}.md", "w") as f:
                    f.write(content)
                print(f"Extracted text for {pdf.material_id}:\n")
                print(content)

            if metadata:
                title = metadata.get("title")
                if metadata.get("subtitle"):
                    title += " - " + metadata.get("subtitle")

                author = metadata.get("authors")
                if isinstance(author, list):
                    author = ", ".join(author)
                subject = metadata.get("subject")
                creator = metadata.get("creator")
                producer = (
                    metadata.get("/Producer") if metadata.get("/Producer") else None
                )

                metadata = {
                    "title": str(title) if title else None,
                    "author": str(author) if author else None,
                    "subject": str(subject) if subject else None,
                    "creator": str(creator) if creator else None,
                    "producer": str(producer) if producer else None,
                }

                pdf = pdf.update_from_dict(metadata)

        info(
            f"{total}/{len(file_paths)} ({total / len(file_paths):.2%}) pdfs processed"
        )
        await PDF.bulk_update(
            batch,
            fields=[
                "extracted_text",
                "extracted_text_max_length",
                "title",
                "author",
                "subject",
                "creator",
                "producer",
            ],
        )


async def extract_pdf_text(
    pdf: PDF,
    max_pages: int | None = 15,
    str_limit: int | None = 50000,
    skip_failed: bool = True,
) -> PDF:
    """
    Extracts text from a PDF file using kreuzberg
    Limit the extraction length by number of pdf pages or output string length.
    Parameters:
        pdf (PDF): A PDF object (tortoise orm model)
        max_pages (int, optional): Max num of pages to process. Defaults to None (all pages).
        str_limit (int, optional): Truncate the result to this length. Defaults to None (no limit).
        skip_failed (bool, optional): Skip pdfs that have failed parsing previously. Defaults to True.
    Returns:
        The updated PDF object with extracted text and metadata (if successful).
    """
    path = pdf.path
    if not path.exists() or not path.is_file() or path.suffix.lower() != ".pdf":
        warn(f"Invalid/Not existing PDF file. Not parsing {path}")
        return pdf

    if pdf.parsing_failed and skip_failed:
        warn(f"PDF parsing failed previously. Not parsing {path}")
        return pdf
    try:
        cur_len = len(pdf.extracted_text) if pdf.extracted_text else 0
        cur_max_len = (
            pdf.extracted_text_max_length if pdf.extracted_text_max_length else 0
        )
        if cur_max_len and str_limit and cur_max_len >= str_limit:
            info(f"pdf already has extracted text of length {cur_max_len}")
            return pdf
    except Exception:
        ...

    result: None | ExtractionResult | CoroutineType = None
    try:
        result = await asyncio.wait_for(
            asyncio.to_thread(extract_file, file_path=path),
            TIMEOUT,
        )

        if isinstance(result, CoroutineType):
            result = await result

    except Exception:
        result = None

    if not result:
        try:
            info(f"trying OCR for {path}")
            result: ExtractionResult | CoroutineType = await asyncio.wait_for(
                asyncio.to_thread(
                    extract_file,
                    file_path=path,
                    config=kreuzberg.ExtractionConfig(
                        force_ocr=True,
                        ocr_backend="paddleocr",
                        ocr_config=kreuzberg.PaddleOCRConfig(
                            language="en",
                            use_gpu=True,  # Enable GPU acceleration if paddlepaddle-gpu is available (experimental)
                        ),
                    ),
                ),
                TIMEOUT,
            )

            if isinstance(result, CoroutineType):
                result = await result

        except TimeoutError:
            warn(
                f"Timeout extracting text from PDF: {path}\n    Setting parsing_failed to True for material id {pdf.material_id}"
            )
            pdf.parsing_failed = True
            await pdf.save()
            return pdf
        except Exception as e:
            warn(
                f"Error extracting text from PDF: {e}.\n    Setting parsing_failed to True for material id {pdf.material_id}"
            )
            pdf.parsing_failed = True
            await pdf.save()
            return pdf

    if not result:
        warn(
            f"Failed to extract text from PDF: {path}\n    Setting parsing_failed to True for material id {pdf.material_id}"
        )

        pdf.parsing_failed = True
        await pdf.save()
        return pdf
    pdf_text = result.content
    metadata = result.metadata

    if str_limit and len(pdf_text) > str_limit:
        pdf_text = pdf_text[:str_limit]

    if pdf_text and cur_len:
        if len(str(pdf_text)) <= int(cur_len):
            warn(
                f"Extracted text is shorter or equal to current text: {len(str(pdf_text))} <= {cur_len}. Not updating."
            )
            return pdf

    update_dict = {
        "extracted_text": pdf_text,
        "extracted_text_max_length": str_limit,
        "extracted_text_max_pages": max_pages,
        "parsing_failed": False,
    }

    if metadata:
        title = metadata.get("title")
        if metadata.get("subtitle"):
            title += " - " + metadata.get("subtitle")

        author = metadata.get("authors")
        if isinstance(author, list):
            author = ", ".join(author)
        subject = metadata.get("subject")
        creator = metadata.get("creator")
        producer = metadata.get("/Producer") if metadata.get("/Producer") else None

        update_dict.update(
            {
                "title": str(title) if title else None,
                "author": str(author) if author else None,
                "subject": str(subject) if subject else None,
                "creator": str(creator) if creator else None,
                "producer": str(producer) if producer else None,
            }
        )

    # write extracted text to file

    pdf = pdf.update_from_dict(update_dict)
    await pdf.save()
    with open(pdf_dir.full / f"{pdf.material_id}.md", "w") as f:
        f.write(pdf_text)
    cool(f"Extracted text with len {len(pdf_text)} from {path}")

    return pdf


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
    import numpy as np
    from paddleocr import PaddleOCR
    from PIL import Image

    if not PAGE_NUM:
        PAGE_NUM = 15
    if not pdf_dir or not pdf_dir.exists():
        pdf_dir = Path().cwd() / "pdfs"
    if not pdf_dir.exists():
        print(f"[red]Directory {pdf_dir} does not exist.[/red]")
        return
    if not outputdir or not outputdir.exists():
        outputdir = Path().cwd() / "paddle_output"

    pdfs = pdf_dir.glob("*.pdf")
    ocr = PaddleOCR(use_angle_cls=True, lang="en", page_num=PAGE_NUM, use_gpu=True)

    for pdf in pdfs:
        pdf_path = pdf
        pdf_name = pdf_path.stem

        if (outputdir / "{pdf_name}_paddle.txt").exists():
            print(f"[red]skipping {pdf_name}[/red]")
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
            print(f"[red]Error processing {pdf_name}: {e}[/red]")
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
            outputdir / f"{pdf_name}" + suffix + ".txt", "w", encoding="utf-8"
        ) as f:
            f.write(full_text)


async def extract_metadata(pdf: PDF) -> PDF:
    """
    Extracts metadata from a PDF file and updates the PDF object with the extracted information.
    This function reads various metadata fields from the PDF file including title, author,
    subject, creator, producer, and dates. If any of these fields contain values, they are
    set as attributes on the PDF object.
    Args:
        pdf (PDF): A PDF object containing the path to the PDF file and other attributes.
    Returns:
        PDF: The updated PDF object with extracted metadata fields set as attributes.
    Note:
        The function will only set attributes for metadata fields that contain values.
        The PDF object is saved to persistence storage after metadata extraction.
    """
    print(f"Extracting metadata from {pdf.path}")
    try:
        file_data = await asyncio.wait_for(
            asyncio.to_thread(pikepdf.open, pdf.path), TIMEOUT
        )
        metadata = file_data.docinfo
    except TimeoutError:
        warn(f"Timeout extracting metadata from PDF: {pdf.path}")
        return pdf
    except (
        pikepdf.PdfError,
        TypeError,
        FileNotFoundError,
        pikepdf.PasswordError,
        pikepdf.DataDecodingError,
    ) as e:
        warn(f"Error extracting metadata from PDF: {e}")
        warn("Deleting pdf file and db entry.")
        # delete pdf
        with contextlib.suppress(Exception):
            File(pdf.path).delete()
        with contextlib.suppress(Exception):
            await pdf.delete()

        return None
    except Exception as e:
        warn(f"Error extracting metadata from PDF: {e}")
        print(traceback.format_exc())
        input("Press enter to continue...")
        return pdf
    if not metadata:
        return pdf

    try:
        title = metadata.get("/Title") if metadata.get("/Title") else None
        author = metadata.get("/Author") if metadata.get("/Author") else None
        subject = metadata.get("/Subject") if metadata.get("/Subject") else None
        creator = metadata.get("/Creator") if metadata.get("/Creator") else None
        producer = metadata.get("/Producer") if metadata.get("/Producer") else None
        metadata = {
            "title": str(title) if title else None,
            "author": str(author) if author else None,
            "subject": str(subject) if subject else None,
            "creator": str(creator) if creator else None,
            "producer": str(producer) if producer else None,
        }

        pdf = pdf.update_from_dict(metadata)
        await pdf.save()
        """        date_data = {
            'file_creation_date': datetime.strptime(str(file_creation_date), "%Y%m%d%H%M%S") if file_creation_date else None,
            'file_modification_date': datetime.strptime(str(file_modification_date), "%Y%m%d%H%M%S") if file_modification_date else None
        }
        pdf = pdf.update_from_dict(date_data)
        await pdf.save()
        """
    except Exception as e:
        warn(f"Error extracting metadata from {pdf.path}: {e}")

    return pdf


async def enrich_pdfs(
    pdfs: list[PDF] | None = None,
    input_mat_ids: list[int] | list[str] | None = None,
    max_pages: int | None = 15,
    str_limit: int | None = 20000,
) -> list[PDF]:
    """
    Deduplicates, and then extracts text & metadata from a list of PDF files, and updates the objects accordingly.
    If no input pdfs are provided, it will first load all pdfs into the DB from disk, then retrieve PDF objects from the db
    Parameters:
        pdfs (list[PDF]): A list of PDF objects (tortoise orm models).
        max_pages (int, optional): Max num of pages to process. Defaults to None (all pages).
        str_limit (int, optional): Truncate the result to this length. Defaults to None (no limit).
    Returns:
        list[PDF]: The same list of PDF objects, but updated with deduplication info, metadata & extracted text.
    """
    await init()
    if pdfs:
        input_mat_ids = [pdf.material_id for pdf in pdfs]

    if not pdfs:
        if input_mat_ids:
            input_mat_ids = [int(mat_id) for mat_id in input_mat_ids]
            pdfs = await PDF.filter(material_id__in=input_mat_ids).all()
        else:
            pdfs = await PDF.all()
            input_mat_ids = [pdf.material_id for pdf in pdfs]

    if input_mat_ids:
        pdfs = await PDF.filter(material_id__in=input_mat_ids).all()
    else:
        pdfs = await PDF.all()
    await load_pdfs()
    pdfs = await PDF.all()

    pdfs_without_metadata = [
        pdf
        for pdf in pdfs
        if not any(
            [
                getattr(pdf, field)
                for field in [
                    "title",
                    "author",
                    "subject",
                    "creator",
                    "producer",
                    "file_creation_date",
                    "file_modification_date",
                ]
            ]
        )
    ]
    # metadata_extracted = [await extract_metadata(pdf) for pdf in pdfs_without_metadata]

    # pdfs = await PDF.all()
    # pdfs_without_metadata = [
    #    pdf
    #    for pdf in pdfs
    #    if not any(
    #        [
    #            getattr(pdf, field)
    #            for field in [
    #                "title",
    #                "author",
    #                "subject",
    #                "creator",
    #                "producer",
    #                "file_creation_date",
    #                "file_modification_date",
    #            ]
    #        ]
    #    )
    # ]
    # if pdf is in pdfs_without_metadata, remove from pdfs
    pdfs = [pdf for pdf in pdfs if pdf not in pdfs_without_metadata]
    pdfs_missing_text = [pdf for pdf in pdfs if not pdf.extracted_text]
    # info(
    #    f"Found {len(pdfs_missing_text)} pdfs without extracted text. First trying batch extract with kreuzberg"
    # )
    # await batch_extract_pdf_text(
    #    pdfs_missing_text, max_pages=max_pages, str_limit=str_limit
    # )
    # input("Press enter to continue...")
    # pdfs_missing_text = [pdf for pdf in pdfs if not pdf.extracted_text]
    info(f"found {len(pdfs_missing_text)} pdfs without extracted text.")

    for pdf in pdfs_missing_text:
        try:
            await extract_pdf_text(pdf, max_pages=max_pages, str_limit=str_limit)
        except Exception as e:
            warn(f"Error extracting text from {pdf.path}: {e}")
            continue

    pdfs_missing_metadata = [
        pdf
        for pdf in pdfs
        if not any(
            [
                getattr(pdf, field)
                for field in [
                    "title",
                    "author",
                    "subject",
                    "creator",
                    "producer",
                    "file_creation_date",
                    "file_modification_date",
                ]
            ]
        )
    ]
    info(f"Found {len(pdfs_missing_metadata)} pdfs without metadata.")

    if input_mat_ids:
        pdfs = await PDF.filter(material_id__in=input_mat_ids).all()
    else:
        pdfs = await PDF.all()

    # info(f"Deduplicating {len(pdfs)} pdfs.")
    # deduplicated_pdfs: list[PDF] = await deduplicate_pdfs(pdfs, False)
    # do something with deduplicated_pdfs


async def deduplicate_pdfs(pdfs: list[PDF], compare_with_db=True) -> list[PDF]:
    """
    Deduplicate a list of PDF files. Uses various techniques to compare the files.
    When duplicates are found, keep the -oldest- (?) file and make the others point to it.
    Parameters:
        pdfs (list[PDF]): A list of pdfs from the db to deduplicate.
    Returns:
        list[PDF]: The same list of pdfs, but updated where possible.
    """
    from qdrant_client import QdrantClient, models

    client = QdrantClient(path="qdrant.db")
    ids_per_hash = defaultdict(list)
    new_mapping: dict[int, int] = {}  # store duplicates as {id_to_replace: target_id}
    replaced_ids = set()  # store ids for pdfs that have been replaced: they do not need to be processed any further

    if compare_with_db:
        pdfs_in_db = (
            await PDF.all()
            .prefetch_related("replace_with")
            .values("material_id", "replace_with__material_id")
        )
    else:
        input_mat_ids = [pdf.material_id for pdf in pdfs]
        pdfs_in_db = (
            await PDF.filter(material_id__in=input_mat_ids)
            .prefetch_related("replace_with")
            .values("material_id", "replace_with__material_id")
        )

    pdfs_with_replace_with = [
        pdf for pdf in pdfs_in_db if pdf.get("replace_with__material_id")
    ]
    replaced_ids.update({pdf.get("material_id") for pdf in pdfs_with_replace_with})
    info(f"Found {len(replaced_ids)} pdfs that are already replaced in db.")
    # first sort the pdfs on their 'age' property
    pdfs.sort(key=lambda pdf: pdf.age)

    # now use list comprehension to populate ids_per_hash.
    # should result in a dict with hash as key and list of ids as value
    # if len(ids_per_hash[hash]) > 1, we have duplicates for that hash!

    [ids_per_hash[get_hash(pdf.path)].append(pdf.material_id) for pdf in pdfs]
    for file_hash, ids in ids_per_hash.items():
        if len(ids) > 1:
            new_mapping.update({id_to_replace: ids[0] for id_to_replace in ids[1:]})
            replaced_ids.union(set(ids[1:]))
    info(f"Found {len(new_mapping)} duplicates by hash.")
    by_hash = len(new_mapping)
    replaced = 0
    already_replaced = 0
    for old_id, new_id in new_mapping.items():
        replace_pdf = await PDF.get_or_none(material_id=new_id)
        if replace_pdf:
            pdf = await PDF.get_or_none(material_id=old_id)
            if pdf:
                current_replacement = await pdf.replace_with
                if current_replacement:
                    already_replaced += 1
                    continue
                pdf.replace_with = replace_pdf
                replaced += 1
                await pdf.save()

    info(
        f"Replaced {replaced} duplicate pdfs with alternative material_id.\n {already_replaced} pdfs already had a replacement."
    )

    # step 2: load extracted text for pdfs into qdrant
    new_mapping = {}
    TRESHOLD = 0.98  # treshold for similarity
    all_embedding_pdfs = [
        pdf
        for pdf in pdfs
        if pdf.extracted_text and pdf.material_id not in replaced_ids
    ]
    info("!!!! \n       Disabled embedding deduplication for now\n!!!!!!")
    if False:
        for embedding_pdfs in batched(all_embedding_pdfs, 20):
            selected: list[PDF] = []
            for pdf in embedding_pdfs:
                result = client.retrieve(
                    collection_name="pdfs",
                    ids=[pdf.material_id],
                    with_payload=False,
                    with_vectors=False,
                )
                if len(result) > 1:
                    continue
                selected.append(pdf)
            if not selected:
                continue
            docs = [pdf.extracted_text for pdf in selected]
            ids = [pdf.material_id for pdf in selected]
            client.add(collection_name="pdfs", documents=docs, ids=ids)
        # once stored, we can query for duplicates
        # determine this by looping over all pdfs & searching for the extracted text in the qdrant db
        # if QueryResponse has a score above treshold: mark as match (add to new_mapping etc).
        info(f"deduplicating {len(all_embedding_pdfs)} by embedding")
        for pdf in all_embedding_pdfs:
            query = pdf.extracted_text
            rep_w: QuerySet[PDF] = pdf.replace_with
            if rep_w:
                rep_pdf: PDF | None = await rep_w.first()
                if rep_pdf:
                    mat_id = rep_pdf.material_id
                    if mat_id:
                        info(f"Item is replaced by {mat_id}, skipping")
                        continue

            # exclude the pdf itself from the query: so id in qdrant != pdf.material_id
            response: list[BaseModel] = client.query(
                collection_name="pdfs",
                query_text=query,
                query_filter=models.Filter(
                    must_not=[
                        models.HasIdCondition(has_id=[pdf.material_id]),
                    ],
                ),
                limit=5,
            )
            for result in response:
                if result.score > TRESHOLD:
                    new_mapping[pdf.material_id] = result.id
                    break

        info(f"Found {len(new_mapping)} duplicates by embedding.")

    if not new_mapping:
        info("No duplicates found!")
        return pdfs

    replaced = 0
    already_replaced = 0
    for old_id, new_id in new_mapping.items():
        replace_pdf = await PDF.get_or_none(material_id=new_id)
        if replace_pdf:
            pdf = await PDF.get_or_none(material_id=old_id)
            if pdf:
                current_replacement = await pdf.replace_with
                if current_replacement:
                    already_replaced += 1
                    continue
                pdf.replace_with = replace_pdf
                replaced += 1
                await pdf.save()

    info(
        f"Replaced {replaced} duplicate pdfs with alternative material_id.\n {already_replaced} pdfs already had a replacement."
    )
    return pdfs


def get_hash(file_path: Path) -> str | None:
    """
    Calculate the XXH64 hash of a file and return it as a hexadecimal string.
    Args:
        file_path (Path): Path to the file to be hashed.
    Returns:
        str or None: Hexadecimal string representation of the file's XXH64 hash; None if an error occurs.
    """
    try:
        with open(file_path, "rb") as f:
            contents = f.read()
        return xxh64(contents).hexdigest()

    except Exception as e:
        warn(f"Error hashing {file_path}: {e}")
        return None


def get_embedding(pdf: PDF) -> list[np.ndarray] | None:
    """
    Calculate the embedding of the extracted pdf text using a pre-trained model.
    Args:
        pdf (PDF): PDF object with extracted text.
    Returns:
        list[float] or None: List of floats representing the file's embedding; None if an error occurs.
    """
    from fastembed import TextEmbedding

    embedding_model = TextEmbedding()
    return embedding_model.embed(pdf.extracted_text)

    pass
