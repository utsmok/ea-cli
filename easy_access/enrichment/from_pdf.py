# Enrich items with data extracted from PDFs

from typing import Any

from loguru import logger

from easy_access.db.models import CopyrightItem
from easy_access.db.retrieve import retrieve_pdfs
from easy_access.settings import Settings


async def enrich_copyright_items_from_pdfs(settings: Settings):
    """
    Enriches copyright items with data extracted from their associated PDFs.
    This includes text content, metadata, and extracted entities.

    Args:
        settings: The application settings.
    """
    logger.info("Starting enrichment of copyright items from PDFs.")

    # Retrieve all PDFs with successful text extraction
    pdfs = await retrieve_pdfs(settings, only_successful_extracts=True)
    logger.info(f"Retrieved {len(pdfs)} PDFs with successful text extraction.")

    #
    # Enrichment logic goes here
    #

    logger.info("Completed enrichment of copyright items from PDFs.")


async def visualize_entities_on_pdf(settings: Settings, lookup: dict[str, Any]):
    """
    Visualizes extracted entities on a PDF document.

    Args:
        settings: The application settings.
        lookup: A dictionary to identify the PDF (e.g., {"id": pdf_id}). Must contain exactly one key-value pair.
        the function will deduce the appropriate filter based on the key and/or the value type.
        Valid key/value pairs include:
            - "id": int
            - "filename": str
            - [any string]: CopyrightItem
            - "url": str
            - "filehash" or "file_hash" or "hash": str

    """
    input_key = list(lookup.keys())[0]
    input_value = lookup[input_key]

    match (input_key, input_value):
        case ("id", int() as pdf_id) | ("id", str() as pdf_id):
            pdf_id = input_value
            lookup_filter = {"id": pdf_id}
        case ("filename", str() as filename):
            filename = input_value
            lookup_filter = {"filename": filename}
        case (_, CopyrightItem() as ci):
            lookup_filter = {"parent": input_value}
        case ("url" | "link", str() as url):
            url = input_value
            lookup_filter = {"url": url}
        case ("filehash" | "file_hash" | "hash", str() as file_hash):
            file_hash = input_value
            lookup_filter = {"file_hash": file_hash}
        case _:
            raise ValueError(
                f"Unrecognized lookup key/value pair: {input_key=}, {input_value=} ({type(input_value)=})"
            )

    # Retrieve the specific PDF
    pdfs = await retrieve_pdfs(settings, filter=lookup_filter)
    if not pdfs:
        logger.error(f"No PDF found with ID {pdf_id}.")
        return

    pdf = pdfs[0]
    if not pdf.extracted_entities:
        logger.warning(f"No extracted entities found for PDF ID {pdf_id}.")
        return

    #
    # Visualization logic goes here
    #

    logger.info(f"Completed visualization for PDF ID {pdf_id}.")
