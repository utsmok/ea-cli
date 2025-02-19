from easy_access.classification import downloader
from easy_access.classification.classifier_api import main, delete_files
from easy_access.classification.downloader import Downloader
from easy_access.classification.pdf_parser import deduplicate_pdfs
import easy_access.settings
import asyncio

if __name__ == "__main__":
    asyncio.run(main())
