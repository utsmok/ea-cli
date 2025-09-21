import datetime
from pathlib import Path

from kreuzberg import ExtractionConfig, ExtractionResult, extract_file
from loguru import logger
from tortoise.expressions import Q
from xxhash import xxh3_64_hexdigest

from easy_access.db.models import PDF, PDFText


async def parse_pdfs(
    filter_ids: list[int] | None = None, parse_text: bool = False
) -> None:
    """Parses all PDFs that have not yet been attempted for text extraction."""

    pdfs = PDF.filter(extraction_successful=False)
    if filter_ids:
        pdfs = pdfs.filter(
            Q(copyright_item_id__in=filter_ids) | Q(v1_copyright_item_id__in=filter_ids)
        )
    pdfs = await pdfs.all()

    logger.info(f"Found {len(pdfs)} PDFs to process")
    if not parse_text:
        logger.warning(
            "Skipping text extraction as parse_text is False -- only hashing PDFs"
        )

    async def process_pdf(pdf: PDF):
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
            pdf_text = await PDFText.create(
                extracted_text=extracted_text, num_pages=num_pages
            )
            pdf.extracted_text = pdf_text
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
        updatefields.append("extracted_text")
        updatefields.append("num_pages")
        return pdf, updatefields

    for pdf in pdfs:
        pdf, updatefields = await process_pdf(pdf)

        try:
            await pdf.save(update_fields=updatefields)
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
    result = await extract_file(
        file_path=path,
        mime_type="application/pdf",
        config=ExtractionConfig(
            ocr_backend=None,
        ),
    )

    if not result.content:
        # no text extracted, try with OCR
        ...
        # skipping for now

    return result


def get_pdf_metadata(result: ExtractionResult) -> dict:
    """Adds metadata from the ExtractionResult to the PDF object."""
    metadata = result.metadata
    pdf = dict()
    if "title" in metadata and metadata["title"]:
        pdf["title"] = metadata["title"]
    if "authors" in metadata and metadata["authors"]:
        if isinstance(metadata["authors"], list):
            if len(metadata["authors"]) == 1:
                pdf["author"] = metadata["authors"][0]
            else:
                pdf["author"] = ", ".join(metadata["authors"])

    if "created_by" in metadata and metadata["created_by"]:
        pdf["created_by"] = metadata["created_by"]
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
