"""
A new .py module that handles the following:

- Receive/Read copyright data for items from all possible sources:
    - Faculty sheets
        - Weekly
        - Overview
    - Raw data exported from Qlik
    - Direct input from web dashboard
    - Direct input from the script
- Normalize, clean, standardize, validate this data
- Update the database with this data



We'll do the work with polars dataframes as much as possible: preferably, use polars methods to read the data.
Use enums and dataclasses where possible to prevent 'magic' strings/dicts/lists/...
Import as much of these enums/dataclasses from other modules as possible to prevent mismatches and to keep the code DRY.

Where possible, use constants defined in the settings module.

"""

from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import Directory, File, cool, info, warn


def sync_sheets_with_db() -> None:
    """
    Sync currently existing sheets with the database: raw data, faculty weekly+overview sheets.

    Reads in data from all sheets, compares it with the db and updates it.
    Then creates a new overview sheet.
    """
    raw_data_dir: Directory = SETTINGS.dirs[DirSetting.RAW_COPYRIGHT_DATA]
    faculty_overview_sheets: list[File] = [
        f
        for f in SETTINGS.dirs[DirSetting.FACULTIES_DIR].files_r
        if "overview" in f.name.lower()
    ]
    faculties = SETTINGS.university_settings.faculty_abbreviations
    faculty_dirs = SETTINGS.dirs[DirSetting.FACULTIES_DIR].dirs
    sheets_per_faculty: dict[str, list[File]] = {
        faculty_dir.name: faculty_dir.files
        for faculty_dir in faculty_dirs
        if faculty_dir.name in faculties
    }

    # remove 'overview' from the list of sheets per faculty
    for faculty, sheets in sheets_per_faculty.items():
        sheets_per_faculty[faculty] = [
            sheet for sheet in sheets if "overview" not in sheet.name.lower()
        ]

    if not raw_data_dir.exists:
        warn(
            f"Raw data directory {raw_data_dir} does not exist. Please check your settings."
        )

    if not faculty_overview_sheets:
        warn(
            f"No faculty overview sheets found in {SETTINGS.dirs[DirSetting.FACULTIES_DIR]}. Please check your settings."
        )

    cool(
        f"Found {len(faculty_overview_sheets)} faculty overview sheets in {SETTINGS.dirs[DirSetting.FACULTIES_DIR]}"
    )
    info(f"Latest file in raw data directory: {raw_data_dir.newest_file()}")
    info(
        f"Found {len(sheets_per_faculty)} faculty directories in {SETTINGS.dirs[DirSetting.FACULTIES_DIR]}"
    )

    # normalize, standardize, validate, clean all data

    # now do comparisons and updates

    # then remove the old overview sheets and create new ones, plus create new weekly sheets if required
