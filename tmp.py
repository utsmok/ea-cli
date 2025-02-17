from easy_access import downloader
from easy_access.classifier_api import main, delete_files
from easy_access.downloader import Downloader
from easy_access.pdf_parser import deduplicate_pdfs
import easy_access.settings
import asyncio

if __name__ == "__main__":
    downloader = Downloader()
    downloader.rename_pdfs()
    downloader.reset_chrome()
