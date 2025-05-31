"""
This module provides functions to handle PDF files for the classification workflow.
Its capabilities include:
- Extracting text and metadata from PDF files using the 'kreuzberg' library.
- Batch processing for text extraction.
- OCR (Optical Character Recognition) using PaddleOCR as a fallback or standalone utility.
- Deduplication of PDFs based on content hashes (embedding-based deduplication is currently disabled).
- Storing extracted information and linking it to PDF records in the database.
"""

import asyncio
import contextlib
import logging # Added
import traceback
from collections import defaultdict
from itertools import batched # Python 3.12+ feature, ensure compatibility or provide alternative
from pathlib import Path
from typing import Any, List, Coroutine, Optional # Used Coroutine instead of CoroutineType

import kreuzberg # For PDF text/metadata extraction
import numpy as np # For embedding, if re-enabled
import pikepdf # For PDF metadata extraction and manipulation
from kreuzberg import ExtractionResult, batch_extract_file, extract_file
# from pydantic import BaseModel # Not used directly in this file's current code
from tortoise.queryset import QuerySet # For type hint if needed

from easy_access.db.base import init as init_tortoise_orm
from easy_access.db.ingest import load_pdfs # To ensure PDFs are in DB before processing
from easy_access.db.models import PDF # The ORM model for PDFs
from easy_access.settings import SETTINGS, DirSetting
# from easy_access.utils import File # File objects not directly used here, paths are used.

logger = logging.getLogger(__name__)

# --- Configuration Candidates (Consider moving to settings.yaml) ---
PDF_PROCESSING_TIMEOUT: int = 30  # Timeout in seconds for potentially long operations like text extraction
PDF_TEXT_STR_LIMIT: int = 50000    # Character limit for extracted text to store
PDF_TEXT_MAX_PAGES: int = 15       # Max pages to process for text extraction by default
PDF_BATCH_SIZE: int = 20           # Batch size for processing multiple PDFs
OCR_LANGUAGE: str = "en"           # Default OCR language
OCR_USE_GPU: bool = True           # Whether to attempt using GPU for OCR
QDRANT_DB_PATH: str = "qdrant.db"  # Path for Qdrant vector database
DEDUPLICATION_SIMILARITY_THRESHOLD: float = 0.98
# --- End Configuration Candidates ---

# pdf_dir is already correctly fetched from SETTINGS later where needed.
# pdf_dir: Path = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full # Global might be problematic if SETTINGS not loaded


async def batch_extract_pdf_text(
    pdfs_to_process: List[PDF],
    max_pages: Optional[int] = PDF_TEXT_MAX_PAGES,
    str_limit: Optional[int] = PDF_TEXT_STR_LIMIT
) -> List[PDF]:
    """
    Extracts text and metadata in batches from a list of PDF model instances.

    Uses `kreuzberg.batch_extract_file` for efficient batch processing.
    Updates PDF instances with extracted text (truncated to `str_limit`) and metadata.
    Saves extracted text to individual .md files and updates PDF records in the database.

    Args:
        pdfs_to_process (List[PDF]): A list of PDF ORM objects to process.
        max_pages (Optional[int]): Maximum number of pages to extract text from.
                                   Defaults to PDF_TEXT_MAX_PAGES.
        str_limit (Optional[int]): Maximum length of the extracted text to store.
                                   Defaults to PDF_TEXT_STR_LIMIT.

    Returns:
        List[PDF]: The list of processed PDF objects (some may be unchanged if processing failed).
    """
    # Filter for PDFs that exist and don't already have sufficient extracted text
    valid_pdfs_for_extraction: List[PDF] = [
        p for p in pdfs_to_process
        if p.path.exists() and not (p.extracted_text and len(p.extracted_text) >= (str_limit or float('inf')))
    ]

    if not valid_pdfs_for_extraction:
        logger.info("No valid PDFs found requiring text extraction in this batch.")
        return pdfs_to_process # Return original list if nothing to process

    logger.info(f"Starting batch text extraction for {len(valid_pdfs_for_extraction)} PDFs.")

    processed_count: int = 0
    pdfs_updated_in_db: List[PDF] = []

    pdf_download_dir: Path = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full

    for pdf_batch in batched(valid_pdfs_for_extraction, PDF_BATCH_SIZE): # Use configured batch size
        batch_file_paths: List[Path] = [p.path for p in pdf_batch]
        extraction_results: List[ExtractionResult] = []
        try:
            # Assuming batch_extract_file can take an ExtractionConfig, if not, max_pages needs to be handled differently
            # For now, kreuzberg's batch_extract_file doesn't directly support max_pages in the same way as single extract_file.
            # This means max_pages here is only for record-keeping unless individual extraction is used.
            extraction_results = await batch_extract_file(batch_file_paths)
        except Exception as e_batch:
            logger.error(f"Error during kreuzberg.batch_extract_file: {e_batch}")
            # Continue to next batch or mark these as failed
            for pdf_in_failed_batch in pdf_batch:
                pdf_in_failed_batch.parsing_failed = True
                pdfs_updated_in_db.append(pdf_in_failed_batch) # To save parsing_failed status
            continue

        for pdf_obj, result in zip(pdf_batch, extraction_results):
            if result and result.content:
                content: str = result.content
                if str_limit and len(content) > str_limit:
                    content = content[:str_limit]

                pdf_obj.extracted_text = content
                pdf_obj.extracted_text_max_length = str_limit
                pdf_obj.extracted_text_max_pages = max_pages # Record intended max_pages
                pdf_obj.parsing_failed = False

                # Save extracted text to a .md file
                try:
                    md_file_path = pdf_download_dir / f"{pdf_obj.material_id}.md"
                    with open(md_file_path, "w", encoding="utf-8") as f_md:
                        f_md.write(content)
                    logger.debug(f"Extracted text for PDF {pdf_obj.material_id} saved to {md_file_path.name}")
                except OSError as e_io:
                    logger.warning(f"Could not write extracted text to .md file for {pdf_obj.material_id}: {e_io}")

            if result and result.metadata:
                metadata_dict = result.metadata
                title = str(metadata_dict.get("title", "") or "")
                if metadata_dict.get("subtitle"): title += f" - {metadata_dict['subtitle']}"

                authors_list = metadata_dict.get("authors", [])
                author_str = ", ".join(authors_list) if isinstance(authors_list, list) else str(authors_list or "")

                pdf_obj.title = title or None
                pdf_obj.author = author_str or None
                pdf_obj.subject = str(metadata_dict.get("subject", "") or "") or None
                pdf_obj.creator = str(metadata_dict.get("creator", "") or "") or None
                # Kreuzberg might use '/Producer', pikepdf uses 'Producer'
                producer_val = metadata_dict.get("Producer") or metadata_dict.get("/Producer")
                pdf_obj.producer = str(producer_val or "") or None

            pdfs_updated_in_db.append(pdf_obj)

        processed_count += len(pdf_batch)
        logger.info(f"{processed_count}/{len(valid_pdfs_for_extraction)} ({processed_count / len(valid_pdfs_for_extraction):.0%}) PDFs processed in batches.")

    if pdfs_updated_in_db:
        try:
            await PDF.bulk_update(
                pdfs_updated_in_db,
                fields=[
                    "extracted_text", "extracted_text_max_length", "extracted_text_max_pages", "parsing_failed",
                    "title", "author", "subject", "creator", "producer", "modified_at"
                ],
            )
            logger.info(f"Bulk updated {len(pdfs_updated_in_db)} PDF records in DB after batch extraction.")
        except Exception as e_db_update:
            logger.error(f"Error bulk updating PDFs in DB after batch extraction: {e_db_update}")
            # Optionally, try individual saves or log failures for manual review

    return pdfs_to_process # Return the original list, with objects modified in place


async def extract_pdf_text(
    pdf: PDF,
    max_pages: Optional[int] = PDF_TEXT_MAX_PAGES,
    str_limit: Optional[int] = PDF_TEXT_STR_LIMIT,
    skip_failed: bool = True,
) -> PDF:
    """
    Extracts text and metadata from a single PDF file using `kreuzberg`.
    Includes OCR fallback with PaddleOCR if direct extraction yields no text.
    Updates the PDF ORM object with extracted information and saves it.

    Args:
        pdf (PDF): The PDF ORM object to process.
        max_pages (Optional[int]): Max number of pages for text extraction.
        str_limit (Optional[int]): Max character length for stored text.
        skip_failed (bool): If True, skips PDFs previously marked as parsing_failed.

    Returns:
        PDF: The updated (or original if failed/skipped) PDF ORM object.
    """
    pdf_path: Path = pdf.path
    if not pdf_path.exists() or not pdf_path.is_file() or pdf_path.suffix.lower() != ".pdf":
        logger.warning(f"Invalid or non-existent PDF file path: {pdf_path}. Cannot parse.")
        return pdf

    if pdf.parsing_failed and skip_failed:
        logger.info(f"Skipping PDF {pdf_path.name} as it previously failed parsing.")
        return pdf

    # Check if text already extracted is sufficient
    current_text_len = len(pdf.extracted_text) if pdf.extracted_text else 0
    current_max_len_setting = pdf.extracted_text_max_length if pdf.extracted_text_max_length else 0
    if str_limit and current_text_len >= str_limit and current_max_len_setting >= str_limit:
        logger.info(f"PDF {pdf_path.name} already has extracted text of sufficient length ({current_text_len}). Skipping.")
        return pdf

    extraction_result: Optional[ExtractionResult] = None
    extraction_method_used = "direct" # For logging

    try:
        logger.debug(f"Attempting direct text extraction for {pdf_path.name} (max_pages: {max_pages}).")
        # kreuzberg.extract_file might not support max_pages directly in its config for all backends.
        # If max_pages is critical, pre-processing the PDF or using a library that supports it per page is needed.
        # For now, it's passed as a record of intent.
        extract_config = kreuzberg.ExtractionConfig(page_limit=max_pages if max_pages else 0) # page_limit=0 means all pages for some backends

        extraction_result = await asyncio.wait_for(
            asyncio.to_thread(extract_file, file_path=pdf_path, config=extract_config),
            timeout=PDF_PROCESSING_TIMEOUT,
        )
    except asyncio.TimeoutError:
        logger.warning(f"Timeout during direct text extraction for {pdf_path.name}.")
    except Exception as e_direct:
        logger.warning(f"Direct text extraction failed for {pdf_path.name}: {e_direct}")

    # Fallback to OCR if direct extraction failed or yielded no content
    if not extraction_result or not extraction_result.content:
        logger.info(f"Direct extraction yielded no content for {pdf_path.name}. Attempting OCR.")
        extraction_method_used = "ocr"
        try:
            ocr_config_paddle = kreuzberg.PaddleOCRConfig(language=OCR_LANGUAGE, use_gpu=OCR_USE_GPU)
            ocr_extract_config = kreuzberg.ExtractionConfig(
                force_ocr=True, ocr_backend="paddleocr", ocr_config=ocr_config_paddle, page_limit=max_pages if max_pages else 0
            )
            extraction_result = await asyncio.wait_for(
                asyncio.to_thread(extract_file, file_path=pdf_path, config=ocr_extract_config),
                timeout=PDF_PROCESSING_TIMEOUT * 2 # Allow more time for OCR
            )
        except asyncio.TimeoutError:
            logger.warning(f"Timeout during OCR extraction for {pdf_path.name}.")
            pdf.parsing_failed = True
        except Exception as e_ocr:
            logger.warning(f"OCR extraction failed for {pdf_path.name}: {e_ocr}")
            pdf.parsing_failed = True

    if not extraction_result or not extraction_result.content:
        logger.warning(f"All extraction attempts (direct and OCR) failed for {pdf_path.name}.")
        pdf.parsing_failed = True # Mark as failed if all attempts yield nothing
    else:
        pdf_text_content: str = extraction_result.content
        if str_limit and len(pdf_text_content) > str_limit:
            pdf_text_content = pdf_text_content[:str_limit]

        pdf.extracted_text = pdf_text_content
        pdf.extracted_text_max_length = str_limit
        pdf.extracted_text_max_pages = max_pages # Record what was attempted
        pdf.parsing_failed = False # Mark as success
        logger.info(f"Successfully extracted text (length {len(pdf_text_content)}, method: {extraction_method_used}) from {pdf_path.name}.")

        # Update metadata if available from the successful extraction
        if extraction_result.metadata:
            meta = extraction_result.metadata
            pdf.title = str(meta.get("title", pdf.title) or "") or None # Keep existing if new is empty
            authors = meta.get("authors", [])
            pdf.author = ", ".join(authors) if isinstance(authors, list) else (str(authors) if authors else None)
            pdf.subject = str(meta.get("subject", pdf.subject) or "") or None
            pdf.creator = str(meta.get("creator", pdf.creator) or "") or None
            producer_val = meta.get("Producer") or meta.get("/Producer") # Check both keys
            pdf.producer = str(producer_val or pdf.producer or "") or None

        # Save extracted text to a .md file
        pdf_download_dir: Path = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full
        md_file_path = pdf_download_dir / f"{pdf.material_id}.md"
        try:
            with open(md_file_path, "w", encoding="utf-8") as f_md:
                f_md.write(pdf.extracted_text or "")
            logger.debug(f"Extracted text for PDF {pdf.material_id} also saved to {md_file_path.name}")
        except OSError as e_io_md:
            logger.warning(f"Could not write extracted text to .md file for {pdf.material_id}: {e_io_md}")

    try:
        await pdf.save() # Save changes (text, metadata, parsing_failed status)
    except Exception as e_save:
        logger.error(f"Failed to save PDF model instance {pdf.material_id} after text extraction: {e_save}")
    return pdf


def ocr_with_paddle( # This function appears to be a standalone utility, not directly integrated.
    pdf_input_dir: Path,
    num_pages_to_ocr: Optional[int] = PDF_TEXT_MAX_PAGES, # Use configured default
    ocr_output_dir: Optional[Path] = None,
    output_suffix: str = "_paddleocr",
) -> None:
    """
    Performs OCR on PDF files in a directory using PaddleOCR via `fitz` (PyMuPDF) for image conversion.
    Saves extracted text to .txt files. This is a standalone utility.

    Args:
        pdf_input_dir (Path): Directory containing PDF files to OCR.
        num_pages_to_ocr (Optional[int]): Number of pages to process from each PDF. Defaults to PDF_TEXT_MAX_PAGES.
        ocr_output_dir (Optional[Path]): Directory to save output .txt files.
                                     Defaults to 'paddle_output' in the current working directory.
        output_suffix (str): Suffix to append to output filenames (before .txt).
    """
    try:
        import cv2
        import fitz # PyMuPDF
        import numpy as np_cv # Alias to avoid clash with main np
        from paddleocr import PaddleOCR # Heavy dependency
        from PIL import Image
    except ImportError as e_import:
        logger.error(f"Missing dependencies for ocr_with_paddle: {e_import}. Please install PyMuPDF, opencv-python, paddleocr, and Pillow.")
        return

    if not pdf_input_dir.exists() or not pdf_input_dir.is_dir():
        logger.error(f"PDF input directory {pdf_input_dir} does not exist or is not a directory.")
        return

    if ocr_output_dir is None:
        ocr_output_dir = Path.cwd() / "paddle_output"
    ocr_output_dir.mkdir(parents=True, exist_ok=True)

    logger.info(f"Initializing PaddleOCR (lang='{OCR_LANGUAGE}', use_gpu={OCR_USE_GPU}). This may take a moment...")
    try:
        ocr_engine = PaddleOCR(use_angle_cls=True, lang=OCR_LANGUAGE, page_num=num_pages_to_ocr, use_gpu=OCR_USE_GPU)
    except Exception as e_paddle_init: # PaddleOCR can raise various errors on init
        logger.error(f"Failed to initialize PaddleOCR engine: {e_paddle_init}")
        return

    for pdf_path_obj in pdf_input_dir.glob("*.pdf"):
        pdf_filename_stem = pdf_path_obj.stem
        output_txt_path = ocr_output_dir / f"{pdf_filename_stem}{output_suffix}.txt"

        if output_txt_path.exists():
            logger.info(f"Skipping {pdf_filename_stem}, output already exists: {output_txt_path.name}")
            continue

        logger.info(f"Processing {pdf_filename_stem} with PaddleOCR...")
        extracted_text_parts: List[str] = []
        try:
            with fitz.open(pdf_path_obj) as pdf_doc:
                pages_to_process = min(num_pages_to_ocr or len(pdf_doc), len(pdf_doc))
                for page_num in range(pages_to_process):
                    page = pdf_doc.load_page(page_num)
                    # Convert to image (PNG) for OCR
                    # Higher DPI can improve OCR but increases processing time
                    pix = page.get_pixmap(dpi=200, alpha=False)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    img_cv = cv2.cvtColor(np_cv.array(img), cv2.COLOR_RGB2BGR)

                    ocr_result = ocr_engine.ocr(img_cv, cls=True)
                    if ocr_result and ocr_result[0] is not None: # Check if result is not None
                        for line_data in ocr_result[0]: # Result is a list of lists of lines
                            if line_data and len(line_data) >= 2 and isinstance(line_data[1], tuple):
                                extracted_text_parts.append(line_data[1][0]) # Text is in line[1][0]

            full_extracted_text = " ".join(extracted_text_parts).replace("  ", " ")
            with open(output_txt_path, "w", encoding="utf-8") as f_out:
                f_out.write(full_extracted_text)
            logger.info(f"Successfully OCR'd {pdf_filename_stem} to {output_txt_path.name}")

        except Exception as e_ocr_process:
            logger.error(f"Error processing {pdf_filename_stem} with PaddleOCR: {e_ocr_process}")


async def extract_metadata(pdf: PDF) -> Optional[PDF]:
    """
    Extracts metadata from a PDF file using `pikepdf` and updates the PDF ORM object.

    Args:
        pdf (PDF): The PDF ORM object to process.

    Returns:
        Optional[PDF]: The updated PDF object if successful, or None if the PDF
                       was deleted due to critical errors (e.g., password protected).
                       The original PDF object is returned if metadata extraction times out
                       or non-critical errors occur.
    """
    logger.debug(f"Attempting metadata extraction from {pdf.path}")
    try:
        # Use asyncio.to_thread for the blocking pikepdf.open call
        pdf_document: pikepdf.Pdf = await asyncio.wait_for(
            asyncio.to_thread(pikepdf.open, pdf.path, allow_overwriting_input=False), # type: ignore
            timeout=PDF_PROCESSING_TIMEOUT
        )
        # Accessing docinfo might also block if it parses parts of the PDF
        docinfo = await asyncio.to_thread(lambda: dict(pdf_document.docinfo)) if pdf_document.docinfo else {} # type: ignore

        # Ensure pdf_document is closed
        await asyncio.to_thread(pdf_document.close)

    except asyncio.TimeoutError:
        logger.warning(f"Timeout extracting metadata from PDF: {pdf.path}")
        return pdf # Return original object
    except pikepdf.PasswordError:
        logger.warning(f"PDF {pdf.path} is password protected. Marking as parsing_failed and deleting file.")
        pdf.parsing_failed = True
        await pdf.save(update_fields=["parsing_failed", "modified_at"])
        with contextlib.suppress(OSError): pdf.path.unlink(missing_ok=True) # Delete the file
        return None # Indicate file was problematic and handled (deleted)
    except (pikepdf.PdfError, TypeError, FileNotFoundError, pikepdf.DataDecodingError) as e_pikepdf:
        logger.warning(f"Pikepdf error extracting metadata from {pdf.path}: {e_pikepdf}. Marking as parsing_failed.")
        pdf.parsing_failed = True
        await pdf.save(update_fields=["parsing_failed", "modified_at"])
        return pdf # Return original object, but marked
    except Exception as e_generic:
        logger.error(f"Unexpected error extracting metadata from {pdf.path}: {e_generic}")
        logger.debug(traceback.format_exc())
        return pdf

    if not docinfo:
        logger.info(f"No standard docinfo metadata found in {pdf.path}.")
        return pdf

    update_dict: Dict[str, Any] = {}
    # Helper to safely convert PdfType objects to string
    def _to_str_if_pdf_type(val: Any) -> Optional[str]:
        if isinstance(val, (pikepdf.String, pikepdf.Name, pikepdf.Date)): return str(val)
        if val is None: return None
        return str(val) # Fallback for other types, though usually they are PdfType

    title = _to_str_if_pdf_type(docinfo.get("/Title"))
    if title: update_dict["title"] = title

    author = _to_str_if_pdf_type(docinfo.get("/Author"))
    if author: update_dict["author"] = author

    subject = _to_str_if_pdf_type(docinfo.get("/Subject"))
    if subject: update_dict["subject"] = subject

    creator = _to_str_if_pdf_type(docinfo.get("/Creator"))
    if creator: update_dict["creator"] = creator

    producer = _to_str_if_pdf_type(docinfo.get("/Producer"))
    if producer: update_dict["producer"] = producer

    # Date parsing for pikepdf dates (format D:YYYYMMDDHHMMSSOHH'mm')
    def parse_pdf_date(pdf_date_str: Optional[str]) -> Optional[datetime]:
        if not pdf_date_str: return None
        pdf_date_str = str(pdf_date_str) # Ensure it's a string
        if pdf_date_str.startswith("D:"):
            pdf_date_str = pdf_date_str[2:]
            # Basic parsing, can be extended for timezone
            formats_to_try = ["%Y%m%d%H%M%S%z", "%Y%m%d%H%M%S", "%Y%m%d"]
            for fmt in formats_to_try:
                try: return datetime.strptime(pdf_date_str[:len(fmt)], fmt) # Match length of format
                except ValueError: continue
        logger.debug(f"Could not parse PDF date string: {pdf_date_str}")
        return None

    creation_date_str = _to_str_if_pdf_type(docinfo.get("/CreationDate"))
    mod_date_str = _to_str_if_pdf_type(docinfo.get("/ModDate"))

    if creation_date_str: update_dict["file_creation_date"] = parse_pdf_date(creation_date_str)
    if mod_date_str: update_dict["file_modification_date"] = parse_pdf_date(mod_date_str)

    if update_dict:
        logger.info(f"Extracted metadata for {pdf.path}: {list(update_dict.keys())}")
        for key, value in update_dict.items():
            setattr(pdf, key, value)
        pdf.parsing_failed = False # Successfully parsed metadata
        await pdf.save()
    else:
        logger.info(f"No new standard metadata fields extracted for {pdf.path}.")

    return pdf


async def enrich_pdfs(
    pdfs_queryset: Optional[QuerySet[PDF]] = None, # Allow passing a queryset
    input_mat_ids: Optional[List[int]] = None,
    max_pages_text: Optional[int] = PDF_TEXT_MAX_PAGES,
    char_limit_text: Optional[int] = PDF_TEXT_STR_LIMIT,
) -> List[PDF]:
    """
    Enriches a list of PDF ORM objects by extracting metadata and text.
    If no PDFs are provided, it fetches them based on `input_mat_ids` or all PDFs from the DB.
    It prioritizes PDFs that haven't had metadata or text extracted yet.

    Args:
        pdfs_queryset (Optional[QuerySet[PDF]]): A Tortoise ORM queryset of PDFs to process.
        input_mat_ids (Optional[List[int]]): A list of material IDs to filter PDFs if `pdfs_queryset` is not given.
        max_pages_text (Optional[int]): Max pages for text extraction.
        char_limit_text (Optional[int]): Character limit for extracted text.

    Returns:
        List[PDF]: The list of processed PDF objects, updated with extracted information.
    """
    await init_tortoise_orm()

    processed_pdfs: List[PDF] = []
    pdfs_to_process: List[PDF]

    if pdfs_queryset is not None:
        pdfs_to_process = await pdfs_queryset
        logger.info(f"Received {len(pdfs_to_process)} PDFs from queryset for enrichment.")
    elif input_mat_ids:
        sanitized_mat_ids = [int(mid) for mid in input_mat_ids if str(mid).isdigit()]
        pdfs_to_process = await PDF.filter(material_id__in=sanitized_mat_ids).all()
        logger.info(f"Fetched {len(pdfs_to_process)} PDFs based on input material IDs for enrichment.")
    else:
        logger.info("No specific PDFs or material_ids provided. Fetching all PDFs from DB for enrichment check.")
        await load_pdfs() # Ensure PDF table is populated from disk if not already
        pdfs_to_process = await PDF.all()
        logger.info(f"Fetched all {len(pdfs_to_process)} PDFs from DB for enrichment.")

    if not pdfs_to_process:
        logger.info("No PDFs found to enrich.")
        await Tortoise.close_connections()
        return []

    # Phase 1: Extract metadata for those missing it
    logger.info("Starting metadata extraction phase.")
    pdfs_needing_metadata = [
        p for p in pdfs_to_process
        if not (p.title or p.author or p.subject or p.creator or p.producer or p.file_creation_date or p.file_modification_date)
    ]
    logger.info(f"Found {len(pdfs_needing_metadata)} PDFs potentially needing metadata extraction.")
    metadata_tasks = [extract_metadata(pdf) for pdf in pdfs_needing_metadata]
    await asyncio.gather(*metadata_tasks, return_exceptions=True) # Results are PDF objects or None

    # Refresh from DB or update in-memory objects if extract_metadata saves individually (it does)
    # For simplicity, assume objects in pdfs_to_process are updated if extract_metadata saves.

    # Phase 2: Extract text for those missing it (can use batch for efficiency)
    logger.info("Starting text extraction phase.")
    pdfs_needing_text = [
        p for p in pdfs_to_process # Use the full list again, in case metadata step marked some as parsing_failed
        if (not p.extracted_text or (char_limit_text and len(p.extracted_text) < char_limit_text)) and not p.parsing_failed
    ]
    logger.info(f"Found {len(pdfs_needing_text)} PDFs potentially needing text extraction.")
    if pdfs_needing_text:
        # Using batch_extract_pdf_text (assuming it's more efficient for many files)
        await batch_extract_pdf_text(pdfs_needing_text, max_pages=max_pages_text, str_limit=char_limit_text)
        # If some still need text (e.g., batch failed for some, or individual retry logic is desired):
        # remaining_needing_text = [p for p in pdfs_needing_text if not p.extracted_text and not p.parsing_failed]
        # if remaining_needing_text:
        #     logger.info(f"Retrying text extraction individually for {len(remaining_needing_text)} PDFs.")
        #     text_tasks = [extract_pdf_text(pdf, max_pages=max_pages_text, str_limit=char_limit_text, skip_failed=True) for pdf in remaining_needing_text]
        #     await asyncio.gather(*text_tasks, return_exceptions=True)


    # Phase 3: Deduplication (currently disabled for embeddings)
    # logger.info(f"Starting deduplication for {len(pdfs_to_process)} PDFs.")
    # processed_pdfs = await deduplicate_pdfs(pdfs_to_process, compare_with_db=True) # compare_with_db=True is default

    # For now, returning the list that went through metadata and text extraction.
    # If deduplication were active and modified relationships, a final fetch might be needed.
    processed_pdfs = pdfs_to_process

    logger.info("PDF enrichment process (metadata, text) completed.")
    await Tortoise.close_connections()
    return processed_pdfs


async def deduplicate_pdfs(pdfs: List[PDF], compare_with_db: bool = True) -> List[PDF]:
    """
    Deduplicates a list of PDF files based on content hashes.
    Updates `replace_with` field in PDF ORM objects for duplicates.
    The embedding-based deduplication part is currently disabled.

    Args:
        pdfs (List[PDF]): A list of PDF ORM objects to deduplicate.
        compare_with_db (bool): If True, considers existing `replace_with` relations from DB.

    Returns:
        List[PDF]: The list of PDF objects, potentially updated with deduplication info.
    """
    try:
        from qdrant_client import QdrantClient, models as qdrant_models # Renamed models
    except ImportError:
        logger.warning("qdrant_client not installed. Embedding-based deduplication will be skipped.")
        # Fallback to hash-based only if Qdrant is not available
        # For now, the embedding part is 'if False' anyway.

    logger.info(f"Starting deduplication process for {len(pdfs)} PDFs. Compare with DB: {compare_with_db}")
    await init_tortoise_orm()

    ids_per_hash: defaultdict[str, list[int]] = defaultdict(list)
    new_hash_mappings: dict[int, int] = {}  # Maps: duplicate_id -> original_id
    already_replaced_ids: set[int] = set()

    if compare_with_db:
        db_pdfs_info = await PDF.all().values("material_id", "replace_with_id")
        for pdf_info in db_pdfs_info:
            if pdf_info["replace_with_id"] is not None:
                already_replaced_ids.add(int(pdf_info["material_id"]))
        logger.info(f"Found {len(already_replaced_ids)} PDFs already marked as duplicates in DB.")

    # Sort by age to prefer keeping the oldest PDF as the original
    # Requires PDF objects to have a reliable 'age' or 'created' attribute.
    # Assuming PDF model has 'created_at' from TimestampMixin.
    pdfs_for_hashing = sorted([p for p in pdfs if p.material_id not in already_replaced_ids], key=lambda p: p.created_at or datetime.min)

    for pdf_obj in pdfs_for_hashing:
        if not pdf_obj.path.exists():
            logger.warning(f"PDF file not found for hashing: {pdf_obj.path}. Skipping.")
            continue
        file_hash = get_hash(pdf_obj.path)
        if file_hash:
            ids_per_hash[file_hash].append(pdf_obj.material_id)

    updated_count_by_hash: int = 0
    for file_hash_val, mat_ids_list in ids_per_hash.items():
        if len(mat_ids_list) > 1:
            original_id = mat_ids_list[0] # Oldest one due to sort
            for duplicate_id in mat_ids_list[1:]:
                if duplicate_id != original_id: # Should always be true here
                    new_hash_mappings[duplicate_id] = original_id

    if new_hash_mappings:
        logger.info(f"Found {len(new_hash_mappings)} potential new duplicates by hash.")
        for duplicate_id, original_id in new_hash_mappings.items():
            try:
                pdf_to_update = await PDF.get_or_none(material_id=duplicate_id)
                original_pdf_ref = await PDF.get_or_none(material_id=original_id)
                if pdf_to_update and original_pdf_ref and pdf_to_update.replace_with_id is None:
                    pdf_to_update.replace_with = original_pdf_ref
                    await pdf_to_update.save(update_fields=["replace_with_id", "modified_at"])
                    updated_count_by_hash += 1
                    already_replaced_ids.add(duplicate_id) # Add to set to exclude from embedding part
            except Exception as e_save_hash:
                logger.error(f"Error saving hash-based duplicate link for {duplicate_id} -> {original_id}: {e_save_hash}")
    logger.info(f"Updated {updated_count_by_hash} PDFs as duplicates based on hash.")

    # Embedding-based deduplication (currently disabled)
    if False:  # pragma: no cover
        logger.info("Embedding-based deduplication is currently disabled.")
        # ... (embedding logic would go here if re-enabled) ...
        # Remember to handle already_replaced_ids for this part too.

    await Tortoise.close_connections()
    # Return the original list of PDF objects; their 'replace_with' may have been updated.
    return pdfs


def get_hash(file_path: Path) -> Optional[str]:
    """
    Calculates the XXH64 hash of a file.

    Args:
        file_path (Path): Path to the file.

    Returns:
        Optional[str]: Hexadecimal string of the hash, or None if an error occurs.
    """
    try:
        with open(file_path, "rb") as f:
            contents = f.read()
        return xxh64(contents).hexdigest()
    except FileNotFoundError:
        logger.warning(f"File not found for hashing: {file_path}")
    except OSError as e_os: # Catch other I/O errors
        logger.warning(f"OS error hashing file {file_path}: {e_os}")
    except Exception as e_generic:
        logger.error(f"Unexpected error hashing file {file_path}: {e_generic}")
    return None


def get_embedding(pdf: PDF) -> Optional[List[np.ndarray]]: # Type for list of numpy arrays
    """
    Calculates the embedding of the extracted text of a PDF using FastEmbed.
    (Note: This function is currently not actively used as embedding deduplication is disabled).

    Args:
        pdf (PDF): PDF ORM object with `extracted_text`.

    Returns:
        Optional[List[np.ndarray]]: List of embeddings (typically one for the whole text,
                                     but FastEmbed can return per-passage). None on error.
    """
    if not pdf.extracted_text:
        logger.debug(f"No extracted text for PDF {pdf.material_id}, cannot get embedding.")
        return None
    try:
        from fastembed import TextEmbedding # Local import as it's a heavy dependency
        embedding_model = TextEmbedding() # Initialize model (could be slow if done repeatedly)
        # embed() returns List[np.ndarray]
        embeddings: List[np.ndarray] = list(embedding_model.embed(pdf.extracted_text))
        return embeddings
    except ImportError:
        logger.warning("FastEmbed library not installed. Cannot generate embeddings.")
    except Exception as e:
        logger.error(f"Error generating embedding for PDF {pdf.material_id}: {e}")
    return None
