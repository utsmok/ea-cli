from easy_access.classification import downloader
from easy_access.classification.classifier_api import main, delete_files
from easy_access.classification.downloader import Downloader
from easy_access.classification.pdf_parser import deduplicate_pdfs
from easy_access.orm.db import init, load_base_data, load_raw_items
import easy_access.settings
import asyncio

async def start_db():
    await init()
    await load_base_data()
    await load_raw_items()

if __name__ == "__main__":
    asyncio.run(start_db())
