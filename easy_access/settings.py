import yaml
from dataclasses import dataclass, field
from easy_access.utils import File, Directory, print, warn
from enum import Enum
from typing import Literal
import logging
import os
import sys
from pathlib import Path
import json
from rich.traceback import install
from loguru import logger

"""
This script reads settings from settings.yaml and parses all containg info into a Settings dataclass,
with nested dataclasses for the different settings.

Used throughout the app to access settings (file/dirnames, university data, mappings, etc) in a structured way.
"""

def configure_logger():
    log_dir = Directory("logs")

    def console_formatter(record):
        level = record["level"].name
        if level == "INFO":
            return "<level>{level} |> </level>{message}\n"
        elif level == "WARNING":
            return "<level>{level} |> </level>{message}\n"
        elif level == "SUCCESS":  # We'll use SUCCESS level for 'cool' messages
            return "<level>{level} |> </level>{message}\n"
        else:
            return "<level>{level}: |> </level>{message}\n"

    logger.add(
        sys.stderr,
        colorize=True,
        format=console_formatter,
        level="TRACE",
        enqueue=True,
    )

    logger.add(
        log_dir.full / "app_{time}.log",
        rotation="1 month",
        format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {message}",
        level="TRACE",
        enqueue=True,
        colorize=False,
    )



@dataclass
class ColInfo:
    """
    contains the info for a single col used in a DataEntrySheet
    """

    name: str  # the colname as included in the sheet (e.g. 'manual_classification')
    dropdown_options: str = (
        ""  # the options for the dropdown; if not applicable, an empty str
    )
    is_url: bool = False  # format as url or not?
    is_new: bool = (
        False  # if True, this col is not present in the original data
    )
    is_editable: bool = False  # if True, this col can be edited
    new_name: str = ""  # if not empty, this col will be renamed to this name
    default_val: str = (
        ""  # if 'is_new' is True, use this as the default value for the new col
    )
    max_width: int = 8  # the max length of any value present in this col, to be set while processing. Min width is this initial number.
    count_max_width_over_40: int = 0  # the number of items in this col that are longer than 40 chars, to be set while processing

    @property
    def has_dropdown(self) -> bool:
        return len(self.dropdown_options) > 0


class DirSetting(Enum):
    """Enum for directories expected by the script"""
    RAW_COPYRIGHT_DATA = "raw_copyright_data"
    EXPORT_TO_SURF = "export_to_surf"
    FACULTIES_DIR = "faculties_dir"
    ALL_ITEMS_DIR = "all_items_dir"
    OVERVIEWS_BACKUP = "overviews_backup"
    SCRIPT_DATA = "script_data"

class FileSetting(Enum):
    """Enum for files expected by the script"""
    FULL_DATA_CSV = "full_data_csv"
    FULL_DATA_PARQUET = "full_data_parquet"
    OSIRIS_DATA = "osiris_data"
    OSIRIS_DATA_W_CONTACTS = "osiris_data_w_contacts"
    PERSON_DATA = "person_data"


class SheetSetting(Enum):
    """Enum for sheet settings expected by the script"""
    RAW_DATA_COL_ORDER = "raw_data_col_order"
    COMPLETE_DATA_NAME = "complete_data_name"
    COMPLETE_DATA_COLS = "complete_data_cols"
    DATA_ENTRY_NAME = "data_entry_name"
    DATA_ENTRY_COLS = "data_entry_cols"
    NEW_FIELDS = "new_fields"


class Functions(str, Enum):
    """
    CLI option for picking which functions to run, see easy_access_cli.cli()
    """

    both = "both"
    read = "read"
    export = "export"



@dataclass
class EasyAccessSettings:
    """Configuration settings for the Easy Access Tool."""
    functions: Functions
    only_changes: bool = True
    save_files: bool = True
    refresh_osiris_data: bool = False
    retrieve_all: bool = True
    other_sheet: Path | None = None
    enrich_with_osiris_data: bool = True
    dirs: dict[DirSetting, Directory] = field(default_factory=dict)
    disable_writes: bool = False

    @classmethod
    def from_env(cls, **kwargs) -> "EasyAccessSettings":
        """Create settings from environment variables and override with kwargs."""
        dirs = SETTINGS.dirs
        return cls(dirs=dirs, **kwargs)


@dataclass(frozen=True)
class Programme:
    name: str | None = None
    abbreviation: str | None = None
    programme_type: Literal['b', 'm', 'o'] = 'o'
    cluster: str | None = None
    faculty_name: str | None = None
    faculty_abbreviation: str | None = None


@dataclass
class Faculty:
    name: str
    abbreviation: str = ""
    programmes: list[Programme] = field(default_factory=list)


@dataclass
class DataSettings:
    data_entry_cols: list[ColInfo] = field(default_factory=list, init=False)
    complete_data_cols: list[ColInfo] = field(default_factory=list, init=False)
    complete_data_name: str = field(default="Complete Data", init=False)
    data_entry_name: str = field(default="Data Entry", init=False)
    raw_data_col_order: list[str] = field(default_factory=list, init=False)
    new_fields: dict[str, dict[str, str|list]] = field(default_factory=list, init=False)

@dataclass
class UniversitySettings:
    name: str = field(default="", init=False)
    abbreviation: str = field(default="", init=False)
    lms: dict[str, str] = field(default_factory=dict, init=False)
    course_catalogue: dict[str, str] = field(default_factory=dict, init=False)
    employee_catalogue: dict[str, str] = field(default_factory=dict, init=False)
    faculties: list[Faculty] = field(default_factory=list, init=False)
    programmes: set[Programme] = field(default_factory=set, init=False)

    def make_programme_set(self):
        if self.faculties:
            for faculty in self.faculties:
                programmes = faculty.programmes
                if not programmes:
                    continue
                for programme in programmes:
                    prog_dict = programme.__dict__
                    prog_dict['faculty_name'] = faculty.name
                    prog_dict['faculty_abbreviation'] = faculty.abbreviation
                    self.programmes.add(Programme(**prog_dict))

    @property
    def faculty_abbreviations(self):
        return {faculty.abbreviation for faculty in self.faculties}
    @property
    def department_mapping(self):
        if not self.programmes:
            self.make_programme_set()
        return {f"{programme.abbreviation+": " if programme.abbreviation else ""}{programme.name}": programme.faculty_abbreviation if programme.faculty_abbreviation else "" for programme in self.programmes}

    @property
    def course_mapping(self):
        if not self.programmes:
            self.make_programme_set()
        course_mapping_dict = {}
        for faculty in self.faculties:
            faculty_name = faculty.abbreviation
            data = dict()
            for programme in faculty.programmes:
                if programme.cluster:
                    data[f"{programme.abbreviation+": " if programme.abbreviation else ""}{programme.name}"] = programme.cluster
            if data:
                course_mapping_dict[faculty_name] = data
        return course_mapping_dict

@dataclass
class Settings:
    """Dataclass holding the app settings"""

    input_file_path: str = "settings.yaml"
    settings_file: File = field(init=False)
    raw_settings: dict = field(default_factory=dict, init=False, repr=False)
    dirs: dict[DirSetting, Directory] = field(default_factory=dict, init=False)
    files: dict[FileSetting, File] = field(default_factory=dict, init=False)
    fine_amount: float = field(default=0.3, init=False, repr=False)
    data_settings: DataSettings = field(default_factory=DataSettings, init=False)
    university_settings: UniversitySettings = field(default_factory=UniversitySettings, init=False)

    def __post_init__(self):
        self.settings_file = File(self.input_file_path)
        self.load()
        if self.raw_settings:
            self.parse_settings()

    def load(self) -> None:
        """Load the settings from the settings.yaml file"""
        try:
            with open(self.settings_file.path) as f:
                self.raw_settings = yaml.load(f, Loader=yaml.FullLoader)
        except Exception as e:
            logger.error(f"Error while loading settings from {self.settings_file}: {e}")
            self.raw_settings = {}

    def parse_settings(self) -> None:
        """Parse the settings into the dataclass"""
        for key, value in self.raw_settings.items():
            if key in self.KEY_TO_PARSER_MAPPING:
                self.KEY_TO_PARSER_MAPPING[key](self,value)
            else:
                logger.error(f'Unrecognized key: {key}. Directly setting value as attribute.')
                setattr(self, key, value)

    def parse_university(self, value):
        self.university_settings.name = value.get('name', "")
        self.university_settings.abbreviation = value.get('abbreviation', "")
        self.university_settings.lms = value.get('lms', {})
        self.university_settings.course_catalogue = value.get('course_catalogue', {})
        self.university_settings.employee_catalogue = value.get('employee_catalogue', {})
        for faculty in value.get('faculties', []):
            name = faculty.get('name', "")
            abbreviation = faculty.get('abbreviation', "")
            programmes = [Programme(**programme) for programme in faculty.get('programmes', [])]
            self.university_settings.faculties.append(Faculty(name, abbreviation, programmes))

        self.university_settings.make_programme_set()

    def parse_data_settings(self, data_settings: dict[str, str|list|dict]):
        for key, value in data_settings.items():
            try:
                key = SheetSetting(key)
            except ValueError:
                logger.error(f"Unrecognized data setting {key} (with value: {value}). Skipping.")
                continue
            match key:
                case SheetSetting.DATA_ENTRY_COLS:
                    self.data_settings.data_entry_cols = [ColInfo(**col_info) for col_info in value]
                case SheetSetting.COMPLETE_DATA_COLS:
                    warn('complete data cols setting not implemented yet')
                case SheetSetting.COMPLETE_DATA_NAME:
                    self.data_settings.complete_data_name = value
                case SheetSetting.DATA_ENTRY_NAME:
                    self.data_settings.data_entry_name = value
                case SheetSetting.RAW_DATA_COL_ORDER:
                    self.data_settings.raw_data_col_order = value
                case SheetSetting.NEW_FIELDS:
                    new_fields = {}
                    for colname, settings in value.items():
                        new_field_dict = {}
                        if 'values' in settings:
                            new_field_dict['values'] = settings['values']
                        if 'default' in settings:
                            new_field_dict['default'] = settings['default']
                        new_fields[colname] = new_field_dict
                    self.data_settings.new_fields = new_fields
                case _:
                    warn(f"Unrecognized data setting {key} (with value: {value}). Skipping.")



    def parse_directories(self, raw_dir_strs:dict[str,str]):
        for key, path in raw_dir_strs.items():
            # turn str key into DefaultDirs enum
            try:
                key = DirSetting(key)
            except ValueError:
                logger.error(f"Unrecognized directory type {key} (with path: {path}). Skipping.")
                continue
            try:
                self.dirs[key] = Directory(path)

            except Exception as e:
                logger.error(f"Error while creating directory {key} with path {path}: {e}")

    def parse_files(self, raw_file_strs:dict[str,str]):
        if 'file_folder' in raw_file_strs:
            self.dirs[DirSetting.SCRIPT_DATA] = Directory(raw_file_strs['file_folder'])
        else:
            self.dirs[DirSetting.SCRIPT_DATA] = Directory(os.getcwd())

        file_dir_path = self.dirs[DirSetting.SCRIPT_DATA].full
        for key, path in raw_file_strs.items():
            if key == 'file_folder':
                continue
            try:
                key = FileSetting(key)
            except ValueError:
                logger.error(f"Unrecognized file type {key} (with path: {path}). Skipping.")
                continue
            try:
                self.files[key] = File(file_dir_path / path)
            except Exception as e:
                logger.error(f"Error while adding file {key} with path {path}: {e}")

    def parse_unsorted(self, rest_values):
        """
        Settings in the unsorted key are set as-is as attributes for this Settings instance
        """
        for key, value in rest_values.items():
            setattr(self, key, value)


    KEY_TO_PARSER_MAPPING = {
        'university': parse_university,
        'data_settings': parse_data_settings,
        'directories': parse_directories,
        'files': parse_files,
        'unsorted': parse_unsorted
    }


def load_osiris_data() -> dict[str,dict]:

    # load osiris data from JSON if available, else warn user to refresh data
    try:
        return json.load(open(SETTINGS.files[FileSetting.OSIRIS_DATA_W_CONTACTS].path))
    except Exception as e:
        print(e)
        logger.error(
            f"{SETTINGS.files[FileSetting.OSIRIS_DATA_W_CONTACTS].path} not found or unreadable. OSIRIS data enrichment will not be possible.\nPlease run the cli again with the refresh_osiris_data flag set to True to retrieve the required data."
        )

# set up logging
logger.remove()
configure_logger()

# install rich traceback as default
install(show_locals=True)

# suppress some annoying warnings when reading excel files
logging.getLogger("fastexcel.types.dtype").setLevel(logging.ERROR)

# initialize settings from (default: read from 'settings.yaml')
SETTINGS = Settings()

# create global variables from certain settings
DEPARTMENT_MAPPING = SETTINGS.university_settings.department_mapping
COURSE_MAPPING = SETTINGS.university_settings.course_mapping
FINE_AMOUNT = SETTINGS.fine_amount
OSIRIS_DATA = load_osiris_data()
