import json
from utils import warn
from settings import SETTINGS, FileSetting
from rich import print

DEPARTMENT_MAPPING = SETTINGS.university_settings.department_mapping
COURSE_MAPPING = SETTINGS.university_settings.course_mapping
FINE_AMOUNT = SETTINGS.fine_amount

try:
    print(SETTINGS.files)
    OSIRIS_DATA:dict[str,dict] = json.load(open(SETTINGS.files[FileSetting.OSIRIS_DATA_W_CONTACTS].path))
except Exception as e:
    print(e)
    warn(
        f"{SETTINGS.files[FileSetting.OSIRIS_DATA_W_CONTACTS].path} not found or unreadable. OSIRIS data enrichment will not be possible.\nPlease run the cli again with the refresh_osiris_data flag set to True to retrieve the required data."
    )
