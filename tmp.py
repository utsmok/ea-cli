from easy_access.classification import downloader
from easy_access.classification.classifier_api import main, delete_files
from easy_access.classification.downloader import Downloader
#from easy_access.classification.pdf_parser import deduplicate_pdfs
import easy_access.settings
from easy_access.settings import DirSetting, SETTINGS
import asyncio
from easy_access.orm.db import update_copyright_relations, init, create, load_base_data, load_raw_items, update_copyright_items, load_new_llm_classifications, link_llm_classifications_to_copyright_items
from easy_access.utils import File

async def dbfuncs():
    #await init()
    #await create()
    #await load_base_data()

    await update_copyright_relations()

if __name__ == "__main__":
    asyncio.run(dbfuncs())