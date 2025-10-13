import datetime
from pathlib import Path

from kreuzberg import (
    ExtractionConfig,
    ExtractionResult,
    extract_file,
)
from loguru import logger
from sqlalchemy import or_, select
from xxhash import xxh3_64_hexdigest

from easy_access.db.compat import save_instance
from easy_access.db.sa_models import PDF as SAPDF
from easy_access.db.sa_models import PDFText as SAPDFText
from easy_access.db.session import get_session


async def parse_pdfs(
    filter_ids: list[int] | None = None, parse_text: bool = True
) -> None:
    """Parses all PDFs that have not yet been attempted for text extraction."""

    async for session in get_session():
        stmt = select(SAPDF).where(SAPDF.extraction_attempted == False)
        if filter_ids:
            stmt = stmt.where(
                or_(
                    SAPDF.copyright_item_id.in_(filter_ids),
                    SAPDF.v1_copyright_item_id.in_(filter_ids),
                )
            )
        result = await session.execute(stmt)
        pdfs = list(result.scalars().all())

    if not pdfs:
        logger.info("No PDFs found without extraction attempts.")
        return
    logger.info(f"Found {len(pdfs)} PDFs to process")
    if not parse_text:
        logger.warning(
            "Skipping text extraction as parse_text is False -- only hashing PDFs"
        )

    async def process_pdf(pdf: SAPDF):
        updatefields = []
        try:
            if hash := hash_pdf(pdf.path):
                pdf.filehash = hash
                updatefields.append("filehash")
        except Exception as e:
            logger.error(f"Error hashing PDF id={pdf.id}, path={pdf.path}: {e}")
        if not parse_text:
            return pdf, updatefields
        try:
            pdf.extraction_attempted = True
            result = await extract_text(pdf.path)
        except Exception as e:
            logger.error(
                f"Error extracting text from PDF id={pdf.id}, path={pdf.path}: {e}"
            )
            result = None

        updatefields.extend(["extraction_attempted", "extraction_successful"])

        if not result or len(result.content or "") < 1:
            pdf.extraction_successful = False

            return pdf, updatefields

        pdf.extraction_successful = True
        extracted_text = result.content

        num_pages = 0
        summary = result.metadata.get("summary")
        if summary and "PDF document with" in summary:
            try:
                num_pages_str = (
                    summary.split("PDF document with")[1].split("pages")[0].strip()
                )
                num_pages = int(num_pages_str)
            except Exception as e:
                logger.error(
                    f"Error parsing number of pages from summary for PDF id={pdf.id}, path={pdf.path}: {e}"
                )
        try:
            pdf_text = SAPDFText(extracted_text=extracted_text, num_pages=num_pages)
            # Save PDFText first to get ID
            await save_instance(pdf_text)
            pdf.extracted_text_id = pdf_text.id
        except Exception as e:
            logger.error(
                f"Error saving extracted text to PDFText for PDF id={pdf.id}, path={pdf.path}: {e}"
            )

        try:
            parsed_metadata = get_pdf_metadata(result)
            for key, value in parsed_metadata.items():
                setattr(pdf, key, value)
        except Exception as e:
            logger.error(
                f"Error adding metadata to PDF id={pdf.id}, path={pdf.path}: {e}"
            )

        updatefields.extend(parsed_metadata.keys())
        updatefields.append("extracted_text_id")
        updatefields.append("num_pages")
        return pdf, updatefields

    for pdf in pdfs:
        pdf, updatefields = await process_pdf(pdf)

        try:
            await save_instance(pdf)
        except Exception as e:
            logger.error(f"Error saving PDF id={pdf.id}, path={pdf.path}: {e}")

    print("Done extracting text from PDFs")


def hash_pdf(file: Path) -> str | None:
    """Calculates the hash of the PDF file and returns it."""
    try:
        hash = xxh3_64_hexdigest(file.read_bytes())
    except Exception as e:
        logger.error(f"Error hashing PDF w/ path={file}: {e}")
        return
    return hash


async def extract_text(path: Path) -> ExtractionResult:
    """
    Extracts text from the PDF file using kreuzberg's extract_file function.
    Currently just a wrapper around kreuzberg.extract_file with the most basic config for PDFs.
    """
    return await extract_file(
        file_path=path,
        mime_type="application/pdf",
        config=ExtractionConfig(
            ocr_backend=None,
        ),
    )


async def ocr_pdfs(max_pages: int = 5) -> None:
    """Runs OCR on all PDFs that have been attempted for text extraction but were not successful."""
    logger.debug("OCR is currently disabled.")
    return


def get_pdf_metadata(result: ExtractionResult) -> dict:
    """Adds metadata from the ExtractionResult to the PDF object."""
    metadata = result.metadata
    pdf = dict()
    if "title" in metadata and metadata["title"]:
        pdf["title"] = metadata["title"]
    if (
        "authors" in metadata
        and metadata.get("authors", "")
        and isinstance(metadata.get("authors"), list)
    ):
        if len(metadata["authors"]) == 1:
            pdf["author"] = metadata["authors"][0]
        else:
            pdf["author"] = ", ".join(metadata["authors"])

    if "created_by" in metadata and metadata["created_by"]:
        pdf["creator"] = metadata["created_by"]
    if "created_at" in metadata and metadata["created_at"]:
        try:
            pdf["creation_date"] = datetime.datetime.fromisoformat(
                metadata["created_at"]
            )
        except ValueError:
            pdf["creation_date"] = None
    if "modified_at" in metadata and metadata["modified_at"]:
        try:
            pdf["mod_date"] = datetime.datetime.fromisoformat(metadata["modified_at"])
        except ValueError:
            pdf["mod_date"] = None
    if "keywords" in metadata and metadata["keywords"]:
        pdf["keywords"] = metadata["keywords"]
    if "subject" in metadata and metadata["subject"]:
        pdf["subject"] = metadata["subject"]
    if "description" in metadata and metadata["description"]:
        pdf["description"] = metadata["description"]
    if "summary" in metadata and metadata["summary"]:
        pdf["summary"] = metadata["summary"]

    if result.keywords:
        # should be list of tuples (keyword, confidence)
        # transform into dict with key = keyword, value = confidence
        pdf["extracted_entities"] = {k: v for k, v in result.keywords}

    return pdf
