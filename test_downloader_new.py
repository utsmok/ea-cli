from easy_access.classification.entity_recognition import process_items
from easy_access.classification.httpx_downloader import download_pdfs
from easy_access.classification.pymupdf4llm_parser import parse_pdfs

if False:
    download_pdfs()

process_items()

try:
    parse_pdfs()
except Exception as e:
    print(f"An error occurred: {e}")
