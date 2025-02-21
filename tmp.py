from easy_access.classification import downloader
from easy_access.classification.classifier_api import main, delete_files
from easy_access.classification.downloader import Downloader
#from easy_access.classification.pdf_parser import deduplicate_pdfs
import easy_access.settings
from easy_access.settings import DirSetting, SETTINGS
import asyncio
from easy_access.orm.db import init, create, load_base_data, load_raw_items

async def dbfuncs():
    await init()
    await create()
    await load_base_data()
    await load_raw_items()

if __name__ == "__main__":
    asyncio.run(dbfuncs())