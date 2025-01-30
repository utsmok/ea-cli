import json
from pathlib import Path
import os
import dotenv
from utils import Directory, warn
dotenv.load_dotenv("settings.env")

DEPARTMENT_MAPPING = json.load(open(Path("department_mapping.json")))
COURSE_MAPPING = json.load(open(Path("course_mapping.json")))
FINE_AMOUNT = 0.3
DIRS =  {
            "copyright_export": Directory(os.getenv("COPYRIGHT_EXPORT_DIR")),
            "copyright_import": Directory(os.getenv("COPYRIGHT_IMPORT_DIR")),
            "faculties": Directory(os.getenv("FACULTIES_DIR")),
            "all_items": Directory(os.getenv("ALL_ITEMS_DIR")),
            "overviews_backup": Directory(os.getenv("OVERVIEWS_BACKUP_DIR")),
        }
try:
    OSIRIS_DATA:dict[str,dict] = json.load(open("osiris_data_w_contacts.json"))
except Exception as e:
    print(e)
    warn(
        f"{os.getcwd()}/osiris_data_w_contacts.json not found or unreadable. OSIRIS data enrichment will not be possible.\nPlease run the cli again with the refresh_osiris_data flag set to True to retrieve the required data."
    )
