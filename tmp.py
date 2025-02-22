from easy_access.classification import downloader
from easy_access.classification.classifier_api import main, delete_files
from easy_access.classification.downloader import Downloader
#from easy_access.classification.pdf_parser import deduplicate_pdfs
import easy_access.settings
from easy_access.settings import DirSetting, SETTINGS
import asyncio
from easy_access.orm.db import update_copyright_relations, init, create, load_base_data, load_raw_items, update_copyright_items, load_new_llm_classifications, link_llm_classifications_to_copyright_items, retrieve_full_data
from easy_access.utils import File
import polars as pl

async def dbfuncs():

    await load_base_data()
    await load_raw_items(SETTINGS.dirs[DirSetting.RAW_COPYRIGHT_DATA_FULL].files[0])
    await update_copyright_relations()
    df =  retrieve_full_data()
    df = df.filter(
            df['period'].is_in(["2024-1A", "2024-1B"])
        )
    df.write_excel("all_items_sem_1_2025.xlsx")

if __name__ == "__main__":
    asyncio.run(dbfuncs())
