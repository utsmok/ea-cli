"""
Functions to handle PDF files:
- download missing pdfs from canvas
- extract text from pdfs
- deduplicate pdfs
- store extracted text
- ...
"""
from pdfminer.high_level import extract_text
from easy_access.utils import info, warn, cool, File, Directory
from easy_access.settings import SETTINGS
from pathlib import Path
from pypdf import PdfReader
from easy_access.db.models import PDF

def extract_pdf_text(pdf: PDF, max_pages: int | None = None, str_limit: int | None = None) -> PDF:
    """
    Extracts text from a PDF file using pdfminer and pypdf.
    Limit the extraction length by number of pdf pages or output string length.
    Parameters:
        pdf (PDF): A PDF object (tortoise orm model)
        max_pages (int, optional): Max num of pages to process. Defaults to None (all pages).
        str_limit (int, optional): Truncate the result to this length. Defaults to None (no limit).
    Returns:
        The updated PDF object with extracted text and metadata (if successful).
    """
    path = pdf.path

    if not path.exists() or not path.is_file() or not path.suffix.lower() == '.pdf':
        warn(f"Invalid/Not existing PDF file: {path}")
        return pdf

    pdf_text: str = extract_text(pdf_file=path, maxpages=max_pages, codec='utf-8')
    if str_limit:
        if len(pdf_text)> str_limit:
            pdf_text = pdf_text[:10_000]

    if not pdf_text:
        warn(f"Failed to extract text from PDF: {path}")
        return pdf

    pdf.extracted_text = pdf_text
    pdf.extracted_text_max_length = str_limit
    pdf.extracted_text_max_pages = max_pages
    reader = PdfReader(path)
    meta = reader.metadata
    metadata = {
        'title': meta.title,
        'author': meta.author,
        'subject': meta.subject,
        'creator': meta.creator,
        'producer': meta.producer,
        'file_creation_date': meta.creation_date,
        'file_modification_date': meta.modification_date
    }
    for key, value in metadata.items():
        if value:
            setattr(pdf, key, value)
    return pdf

async def bulk_extract_text(pdfs: list[PDF], max_pages: int | None = None, str_limit: int | None = None) -> list[PDF]:
    """
    Extract text from a list of PDF files. Update the objects with extracted text and metadata.
    Parameters:
        pdfs (list[PDF]): A list of PDF objects (tortoise orm models).
        max_pages (int, optional): Max num of pages to process. Defaults to None (all pages).
        str_limit (int, optional): Truncate the result to this length. Defaults to None (no limit).
    Returns:
        list[PDF]: The same list of PDF objects, but updated with extracted text and metadata.
    """
    for pdf in pdfs:
        pdf = extract_pdf_text(pdf, max_pages=max_pages, str_limit=str_limit)
    await PDF.bulk_update(pdfs, update_fields=['modified_at', 'extracted_text', 'extracted_text_max_length', 'extracted_text_max_pages', 'title', 'author', 'subject', 'creator', 'producer', 'file_creation_date', 'file_modification_date'])
    return pdfs

async def deduplicate_pdfs(pdfs: list[PDF]) -> list[PDF]:
    """
    Deduplicate a list of PDF files. Uses various techniques to compare the files.
    Parameters:
        pdfs (list[PDF]): A list of pdfs from the db to deduplicate.
    Returns:
        list[PDF]: The same list of pdfs, but updated where possible.
    """
    for pdf in pdfs:
        replacement_id = None
        # do stuff to determine if pdf is a duplicate
        # if pdf is a duplicate, set replacement_id to the id of the pdf to replace
        if replacement_id:
            replace_pdf = await PDF.get_or_none(material_id=replacement_id)
            if replace_pdf:
                pdf.replace_with(replace_pdf)
                await pdf.save()
    return pdfs
