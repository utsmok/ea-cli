from easy_access.classifier_api import main, delete_files
from easy_access.downloader import Downloader
from easy_access.pdf_parser import deduplicate_pdfs
import easy_access.settings
import asyncio

if __name__ == "__main__":
    delete_files()
    downloader = Downloader()
    downloader.download_pdfs(subset=[])
    deduplicate_pdfs()


