from easy_access.classification import downloader
from easy_access.classification.classifier_api import main, delete_files
from easy_access.classification.downloader import Downloader
#from easy_access.classification.pdf_parser import deduplicate_pdfs
import easy_access.settings
from easy_access.settings import DirSetting, SETTINGS
import asyncio

if __name__ == "__main__":
    all_files = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].files
    pdfs = [f for f in all_files if f.extension == '.pdf']
    pdfs_found = len(pdfs)

    pdfs_for_mat_ids = [f for f in pdfs if "_" in f.name]
    pdf_material_ids: dict = {f.name.split(sep='_')[0]:f for f in pdfs_for_mat_ids}

    existing_classifications = [f.name.rstrip('.json') for f in SETTINGS.dirs[DirSetting.CLASSIFICATIONS].files if f.extension == '.json' and "_old" not in f.name]

    if existing_classifications:
        pdfs = [pdf_material_ids.get(f) for f in pdf_material_ids if f not in existing_classifications]

    print(f'{pdfs_found} files found, with {len(existing_classifications)} already classified. {len(pdfs)} files remaining.')
    if len(pdfs) < 30:
        for pdf in pdfs:
            print(pdf.name)