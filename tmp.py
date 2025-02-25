from easy_access.classification import downloader
from easy_access.classification.classifier_api import main, delete_files
from easy_access.classification.downloader import Downloader
#from easy_access.classification.pdf_parser import deduplicate_pdfs
from easy_access.settings import DirSetting, SETTINGS
import asyncio
from easy_access.utils import File
import polars as pl
from easy_access.db.base import init, create
from easy_access.db.ingest import load_pdfs
from easy_access.classification.pdf_handling import enrich_pdfs


async def run():
    # first download pdfs
    # then ingest into db
    # then enrich
    await init()
    print(f'downloading pdfs')
    downloader = Downloader()
    await downloader.download_pdfs(None)
    print(f'done, now ingesting into db')
    await load_pdfs()
    await enrich_pdfs(max_pages=20, str_limit=50000)


asyncio.run(run())
