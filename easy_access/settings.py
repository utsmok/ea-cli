import json
import os
import sys
from collections.abc import Callable
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Literal

import yaml
from loguru import logger
from rich.traceback import install

from easy_access.utils import Directory, File, safe_float

"""Manages application settings, loaded from YAML configuration files.

This module defines dataclasses for structuring settings and provides
functionality to load and parse them from 'settings.yaml' and 'sample.yaml'.
It configures logging and establishes global setting constants for use
throughout the application.
"""


def configure_logger() -> None:
    """Configures the Loguru logger for console and file output."""
    log_dir = Directory(path="logs")

    def console_formatter(record) -> str:
        """Formats log messages for console output.

        Args:
            record: The Loguru log record.

        Returns:
            The formatted log string.
        """
        return f"<level>{record['level'].name} | </level>{{message}}\n"

    logger.add(
        sink=sys.stderr,
        colorize=True,
        format=console_formatter,
        level="TRACE",
        enqueue=True,
    )

    logger.add(
        sink=log_dir.full / "app_{time}.log",
        rotation="1 month",
        format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {message}",
        level="TRACE",
        enqueue=True,
        colorize=False,
    )


@dataclass
class ColInfo:
    """Contains the info for a single column used in a DataEntrySheet.

    Attributes:
        name: The column name as included in the sheet (e.g., 'manual_classification').
        dropdown_options: Comma-separated string of options for a dropdown; empty if not applicable.
        is_url: Whether to format the column content as a URL.
        is_new: If True, this column is not present in the original data.
        is_editable: If True, this column can be edited.
        new_name: If not empty, this column will be renamed to this name.
        default_val: If 'is_new' is True, use this as the default value for the new column.
        max_width: The maximum length of any value present in this column, to be set during processing.
                   Minimum width is this initial number.
        count_max_width_over_40: The number of items in this column longer than 40 characters,
                                 to be set during processing.
    """

    name: str
    dropdown_options: str = ""
    is_url: bool = False
    is_new: bool = False
    is_editable: bool = False
    new_name: str = ""
    default_val: str = ""
    max_width: int = 8
    count_max_width_over_40: int = 0

    @property
    def has_dropdown(self) -> bool:
        """True if dropdown_options are specified, False otherwise."""
        return len(self.dropdown_options) > 0


class DirSetting(Enum):
    """Enum for directory settings keys used in settings.yaml."""

    RAW_COPYRIGHT_DATA = "raw_copyright_data"
    RAW_COPYRIGHT_DATA_FULL = "full_data"
    EXPORT_TO_SURF = "export_to_surf"
    FACULTIES_DIR = "faculties_dir"
    ALL_ITEMS_DIR = "all_items_dir"
    OVERVIEWS_BACKUP = "overviews_backup"
    SCRIPT_DATA = "script_data"
    FULL_BACKUPS = "full_backups"
    PDF_DOWNLOADS = "pdf_downloads"
    CLASSIFICATIONS = "classifications"


class FileSetting(Enum):
    """Enum for file settings keys used in settings.yaml."""

    FULL_DATA_CSV = "full_data_csv"
    FULL_DATA_PARQUET = "full_data_parquet"
    OSIRIS_DATA = "osiris_data"
    OSIRIS_DATA_W_CONTACTS = "osiris_data_w_contacts"
    PERSON_DATA = "person_data"


class BackupSetting(Enum):
    """Enum for backup settings keys used in settings.yaml."""

    BACKUP_ALL = "backup_all"
    BACKUP_DIRS = "backup_dirs"
    MAX_BACKUPS = "max_backups"
    BACKUP_OVERVIEWS = "backup_overviews"


class SheetSetting(Enum):
    """Enum for sheet-related settings keys used in settings.yaml."""

    RAW_DATA_COL_ORDER = "raw_data_col_order"
    FINAL_DATA_COL_ORDER = "final_data_col_order"
    COMPLETE_DATA_NAME = "complete_data_name"
    COMPLETE_DATA_COLS = "complete_data_cols"
    DATA_ENTRY_NAME = "data_entry_name"
    DATA_ENTRY_COLS = "data_entry_cols"
    NEW_FIELDS = "new_fields"


class Functions(str, Enum):
    """CLI option for picking which functions to run.

    Used in `easy_access_cli.cli()`.
    """

    both = "both"
    read = "read"
    export = "export"


@dataclass
class EasyAccessSettings:
    """Configuration settings for the Easy Access Tool."""

    export: bool = False
    only_changes: bool = True
    refresh_osiris_data: bool = False
    only_retrieve_missing_osiris_data: bool = False
    other_sheet: Path | None = None
    enrich_with_osiris_data: bool = True
    dirs: dict[DirSetting, Directory] = field(default_factory=dict)
    disable_writes: bool = False
    faculty: str | None = None

    @classmethod
    def create_for_runtime(
        cls, main_settings: "Settings", **kwargs: Any
    ) -> "EasyAccessSettings":
        """Creates EasyAccessSettings for a specific runtime execution.

        Derives directory configurations from the main Settings object
        and overrides them with any provided keyword arguments (typically from CLI).

        Args:
            main_settings: The main application Settings object.
            **kwargs: Keyword arguments to override default EasyAccessSettings.

        Returns:
            An instance of EasyAccessSettings configured for runtime.
        """
        dirs: dict[DirSetting, Directory] = main_settings.dirs
        return cls(dirs=dirs, **kwargs)


@dataclass
class DataSettings:
    """Holds settings related to data structure, column definitions, and naming.

    Attributes:
        data_entry_cols: Configuration for columns in the data entry sheet.
        complete_data_cols: Configuration for columns in the complete data sheet.
        complete_data_name: Name for the complete data sheet.
        data_entry_name: Name for the data entry sheet.
        raw_data_col_order: Order of columns in the raw data.
        final_data_col_order: Order of columns in the final processed data.
        new_fields: Definitions for new fields to be added during processing,
                    mapping column names to their settings (e.g., 'values', 'default').
        url_truncation_marker: Marker to indicate a truncated URL.
        url_default_base: Default base URL for constructing full URLs.
    """

    data_entry_cols: list[ColInfo] = field(default_factory=list, init=False)
    complete_data_cols: list[ColInfo] = field(default_factory=list, init=False)
    complete_data_name: str = field(default="Complete Data", init=False)
    data_entry_name: str = field(default="Data Entry", init=False)
    raw_data_col_order: list[str] = field(default_factory=list, init=False)
    final_data_col_order: list[str] = field(default_factory=list, init=False)
    new_fields: dict[str, dict[str, str | list[str]]] = field(
        default_factory=dict, init=False
    )
    url_truncation_marker: str = field(default="...", init=False)
    url_default_base: str = field(
        default="https://utwente.instructure.com/files", init=False
    )


@dataclass
class BackupSettings:
    """Holds settings related to data backup procedures.

    Attributes:
        backup_all: Whether to back up all relevant data.
        backup_dirs: A set of specific directories to include in the backup.
        max_backups: The maximum number of backups to retain.
        backup_overviews: Whether to back up overview files.
        backup_location: The directory where backups will be stored.
    """

    backup_all: bool = True
    backup_dirs: list[Directory] = field(default_factory=list)
    max_backups: int = 3
    backup_overviews: bool = True
    backup_location: Directory | None = field(default=None)


@dataclass(frozen=True)
class SettingsProgramme:
    """Represents a university programme with its details.

    Attributes:
        name: The full name of the programme.
        abbreviation: The abbreviation for the programme.
        programme_type: The type of programme (e.g., 'b' for bachelor, 'm' for master).
        cluster: The cluster the programme belongs to, if any.
        faculty_name: The name of the faculty the programme belongs to.
        faculty_abbreviation: The abbreviation of the faculty.
    """

    name: str | None = None
    abbreviation: str | None = None
    programme_type: Literal["b", "m", "o"] = "o"
    cluster: str | None = None
    faculty_name: str | None = None
    faculty_abbreviation: str | None = None


@dataclass
class SettingsFaculty:
    """Represents a university faculty and its associated programmes.

    Attributes:
        name: The full name of the faculty.
        abbreviation: The abbreviation for the faculty.
        programmes: A list of programmes offered by this faculty.
    """

    name: str
    abbreviation: str = ""
    programmes: list[SettingsProgramme] = field(default_factory=list)


@dataclass
class UniversitySettings:
    """Holds settings related to the university structure, including faculties and programmes.

    Attributes:
        name: The name of the university.
        abbreviation: The abbreviation for the university.
        lms: Learning Management System URLs or identifiers.
        course_catalogue: Course catalogue URLs or identifiers.
        employee_catalogue: Employee directory URLs or identifiers.
        faculties: A list of faculties within the university.
        programmes: A set of all unique programmes across all faculties.
        manual_department_mappings: Manual overrides for mapping programmes/departments to faculties.
    """

    name: str = field(default="", init=False)
    abbreviation: str = field(default="", init=False)
    lms: dict[str, str] = field(default_factory=dict, init=False)
    course_catalogue: dict[str, str] = field(default_factory=dict, init=False)
    employee_catalogue: dict[str, str] = field(default_factory=dict, init=False)
    faculties: list[SettingsFaculty] = field(default_factory=list, init=False)
    programmes: set[SettingsProgramme] = field(default_factory=set, init=False)
    manual_department_mappings: dict[str, str] = field(default_factory=dict, init=False)

    def make_programme_set(self) -> None:
        """Populates the `programmes` set from the list of faculties.

        Iterates through each faculty and its programmes, adding them to the
        `self.programmes` set, enriching them with faculty information.
        """
        if self.faculties:
            for faculty in self.faculties:
                programmes: list[SettingsProgramme] = faculty.programmes
                if not programmes:
                    continue
                for programme in programmes:
                    # Use asdict for dataclasses if available and preferred,
                    # otherwise direct attribute access or __dict__ is common.
                    prog_dict: dict[str, Any] = vars(programme).copy()  # Make a copy
                    prog_dict["faculty_name"] = faculty.name
                    prog_dict["faculty_abbreviation"] = faculty.abbreviation
                    self.programmes.add(SettingsProgramme(**prog_dict))

    @property
    def faculty_abbreviations(self) -> set[str]:
        """A set of all unique faculty abbreviations."""
        return {
            faculty.abbreviation for faculty in self.faculties if faculty.abbreviation
        }

    @property
    def department_mapping(self) -> dict[str, str]:
        """A mapping of programme names/abbreviations to faculty abbreviations.

        Combines programmatically generated mappings with manual overrides from settings.
        Manual overrides take precedence.
        """
        if not self.programmes:
            self.make_programme_set()

        generated_data = {
            f"{programme.abbreviation + ': ' if programme.abbreviation else ''}{programme.name}": programme.faculty_abbreviation
            if programme.faculty_abbreviation
            else ""
            for programme in self.programmes
            if programme.name  # Ensure programme name exists
        }

        merged_data = generated_data.copy()
        merged_data.update(self.manual_department_mappings)
        return merged_data

    @property
    def course_mapping(self) -> dict[str, dict[str, str]]:
        """A mapping of faculty abbreviations to their course cluster mappings.

        Each faculty maps to a dictionary where programme names/abbreviations map to their cluster.
        """
        if not self.programmes:  # Ensure programmes are populated
            self.make_programme_set()

        course_mapping_dict: dict[str, dict[str, str]] = {}
        for faculty in self.faculties:
            if not faculty.abbreviation:  # Skip faculty if no abbreviation
                logger.warning(
                    f"Faculty '{faculty.name}' has no abbreviation. Skipping!"
                )
                continue
            faculty_key: str = faculty.abbreviation
            data: dict[str, str] = {}
            for programme in faculty.programmes:
                if (
                    programme.cluster and programme.name
                ):  # Ensure cluster and name exist
                    programme_key = f"{programme.abbreviation + ': ' if programme.abbreviation else ''}{programme.name}"
                    data[programme_key] = programme.cluster
            if data:
                course_mapping_dict[faculty_key] = data
        return course_mapping_dict


@dataclass
class Settings:
    """Main dataclass holding all application settings, loaded from a YAML file.

    Attributes:
        input_file_path: Path to the settings YAML file (default: "settings.yaml").
        settings_file: File object representing the settings file.
        raw_settings: Raw dictionary loaded from the YAML file.
        dirs: Dictionary mapping directory setting keys (DirSetting) to Directory objects.
        files: Dictionary mapping file setting keys (FileSetting) to File objects.
        fine_amount: Default fine amount for certain calculations.
        data_settings: Nested DataSettings object.
        university_settings: Nested UniversitySettings object.
        backup_settings: Nested BackupSettings object.
        classification_options: List of available classification options.
        dashboard_reload: Boolean indicating if the dashboard should auto-reload.
        db_path: Path to the SQLite database file.
        KEY_TO_PARSER_MAPPING: Internal mapping of setting keys to parser methods.
    """

    input_file_path: str = "settings.yaml"
    settings_file: File = field(init=False)
    raw_settings: dict[str, Any] = field(default_factory=dict, init=False, repr=False)
    dirs: dict[DirSetting, Directory] = field(default_factory=dict, init=False)
    files: dict[FileSetting, File] = field(default_factory=dict, init=False)
    fine_amount: float = field(default=0.3, init=False, repr=False)
    data_settings: DataSettings = field(default_factory=DataSettings, init=False)
    university_settings: UniversitySettings = field(
        default_factory=UniversitySettings, init=False
    )
    backup_settings: BackupSettings = field(default_factory=BackupSettings, init=False)
    classification_options: list[str] = field(default_factory=list)
    dashboard_reload: bool = field(default=True, init=False)
    db_path: Path = field(default=Path("db.sqlite3"), init=False)
    KEY_TO_PARSER_MAPPING: dict[str, Callable[[Any], None]] = field(
        init=False, repr=False
    )

    def _parse_dict_safely(self, data: Any, key: str) -> dict[str, str]:
        """Safely parses a dictionary from settings, ensuring string keys and values."""
        raw_dict = data.get(key, {})
        if isinstance(raw_dict, dict):
            return {str(k): str(v) for k, v in raw_dict.items()}
        logger.warning(
            f"Setting '{key}' is not a valid dictionary. Using empty dictionary."
        )
        return {}

    def parse_university(self, value: dict[str, Any]) -> None:
        """Parses the 'university' section of settings.yaml.

        Args:
            value: The dictionary representing the 'university' settings.
        """
        self.university_settings.name = str(value.get("name", ""))
        self.university_settings.abbreviation = str(value.get("abbreviation", ""))
        self.university_settings.lms = self._parse_dict_safely(value, "lms")
        self.university_settings.course_catalogue = self._parse_dict_safely(
            value, "course_catalogue"
        )
        self.university_settings.employee_catalogue = self._parse_dict_safely(
            value, "employee_catalogue"
        )

        manual_mappings = value.get("manual_department_mappings", {})
        if isinstance(manual_mappings, dict):
            self.university_settings.manual_department_mappings = {
                str(k): str(v) for k, v in manual_mappings.items()
            }
        else:
            logger.warning(
                "'manual_department_mappings' in settings.yaml is not a valid dictionary. Skipping."
            )
            self.university_settings.manual_department_mappings = {}

        faculties_data = value.get("faculties", [])
        parsed_faculties: list[SettingsFaculty] = []
        if isinstance(faculties_data, list):
            for faculty_data_item in faculties_data:
                if isinstance(faculty_data_item, dict):
                    name: str = str(faculty_data_item.get("name", "Unknown Faculty"))
                    abbreviation: str = str(faculty_data_item.get("abbreviation", ""))
                    programmes_data = faculty_data_item.get("programmes", [])
                    programmes: list[SettingsProgramme] = []
                    if isinstance(programmes_data, list):
                        for programme_dict_any in programmes_data:
                            if isinstance(programme_dict_any, dict):
                                # Ensure all keys are strings for SettingsProgramme
                                programme_dict: dict[str, Any] = {
                                    str(k): v for k, v in programme_dict_any.items()
                                }
                                try:
                                    programmes.append(
                                        SettingsProgramme(**programme_dict)
                                    )
                                except TypeError as e:
                                    logger.error(
                                        f"Error parsing programme {programme_dict} for faculty {name}: {e}"
                                    )
                            else:
                                logger.warning(
                                    f"Programme entry for faculty {name} is not a dict: {programme_dict_any}"
                                )
                    else:
                        logger.warning(
                            f"Programmes data for faculty {name} is not a list: {programmes_data}"
                        )
                    parsed_faculties.append(
                        SettingsFaculty(
                            name=name, abbreviation=abbreviation, programmes=programmes
                        )
                    )
                else:
                    logger.warning(f"Faculty entry is not a dict: {faculty_data_item}")
        else:
            logger.warning(f"Faculties data is not a list: {faculties_data}")
        self.university_settings.faculties = parsed_faculties
        self.university_settings.make_programme_set()

    def parse_data_settings(self, data_settings_yaml: dict[str, Any]) -> None:
        """Parses the 'data_settings' section of settings.yaml.

        Args:
            data_settings_yaml: The dictionary representing the 'data_settings'.
        """
        if "url_truncation_marker" in data_settings_yaml:
            val = data_settings_yaml.pop("url_truncation_marker")
            if isinstance(val, str):
                self.data_settings.url_truncation_marker = val
            else:
                logger.warning(
                    f"data_settings.url_truncation_marker is not a string: {val}"
                )

        if "url_default_base" in data_settings_yaml:
            val = data_settings_yaml.pop("url_default_base")
            if isinstance(val, str):
                self.data_settings.url_default_base = val
            else:
                logger.warning(f"data_settings.url_default_base is not a string: {val}")

        for key_str, value_data in data_settings_yaml.items():
            try:
                key_enum = SheetSetting(value=key_str)
            except ValueError:
                logger.error(
                    f"Unrecognized data setting {key_str} (with value: {value_data}). Skipping."
                )
                continue
            match key_enum:
                case SheetSetting.DATA_ENTRY_COLS:
                    if isinstance(value_data, list):
                        parsed_cols: list[ColInfo] = []
                        for col_info_dict_any in value_data:
                            if isinstance(col_info_dict_any, dict):
                                col_info_dict: dict[str, Any] = {
                                    str(k): v for k, v in col_info_dict_any.items()
                                }
                                try:
                                    parsed_cols.append(ColInfo(**col_info_dict))
                                except TypeError as e:
                                    logger.error(
                                        f"Error parsing ColInfo from {col_info_dict}: {e}"
                                    )
                            else:
                                logger.warning(
                                    f"Expected dict for ColInfo, got {type(col_info_dict_any)}: {col_info_dict_any}"
                                )
                        self.data_settings.data_entry_cols = parsed_cols
                    else:
                        logger.warning(
                            f"Expected list for DATA_ENTRY_COLS, got {type(value_data)}"
                        )
                case SheetSetting.COMPLETE_DATA_COLS:
                    # Parsing for 'complete_data_cols' can be implemented here if needed.
                    pass
                case SheetSetting.COMPLETE_DATA_NAME:
                    self.data_settings.complete_data_name = str(value_data)
                case SheetSetting.DATA_ENTRY_NAME:
                    self.data_settings.data_entry_name = str(value_data)
                case SheetSetting.RAW_DATA_COL_ORDER:
                    if isinstance(value_data, list):
                        self.data_settings.raw_data_col_order = [
                            str(item) for item in value_data
                        ]
                    else:
                        logger.warning(
                            f"Expected list for RAW_DATA_COL_ORDER, got {type(value_data)}"
                        )
                case SheetSetting.FINAL_DATA_COL_ORDER:
                    if isinstance(value_data, list):
                        self.data_settings.final_data_col_order = [
                            str(item) for item in value_data
                        ]
                    else:
                        logger.warning(
                            f"Expected list for FINAL_DATA_COL_ORDER, got {type(value_data)}"
                        )
                case SheetSetting.NEW_FIELDS:
                    parsed_new_fields: dict[str, dict[str, str | list[str]]] = {}
                    if isinstance(value_data, dict):
                        for colname, settings_val_any in value_data.items():
                            if isinstance(settings_val_any, dict):
                                settings_val: dict[str, Any] = {
                                    str(k): v for k, v in settings_val_any.items()
                                }
                                new_field_entry: dict[str, str | list[str]] = {}
                                if "values" in settings_val and isinstance(
                                    settings_val["values"], list
                                ):
                                    new_field_entry["values"] = [
                                        str(v) for v in settings_val["values"]
                                    ]
                                if "default" in settings_val and isinstance(
                                    settings_val["default"], str
                                ):
                                    new_field_entry["default"] = str(
                                        settings_val["default"]
                                    )
                                if (
                                    new_field_entry
                                ):  # Only add if it has valid 'values' or 'default'
                                    parsed_new_fields[str(colname)] = new_field_entry
                            else:
                                logger.warning(
                                    f"Expected dict for new_field settings, got {type(settings_val_any)} for {colname}"
                                )
                    else:
                        logger.warning(
                            f"Expected dict for NEW_FIELDS, got {type(value_data)}"
                        )
                    self.data_settings.new_fields = parsed_new_fields
                case _:
                    logger.warning(
                        f"Unrecognized data setting {key_str} (with value: {value_data}). Skipping."
                    )

    def parse_backup(self, backup_settings_yaml: dict[str, Any]) -> None:
        """Parses the 'backup' section of settings.yaml.

        Args:
            backup_settings_yaml: The dictionary representing the 'backup' settings.
        """
        for key_str, value_data in backup_settings_yaml.items():
            try:
                key_enum = BackupSetting(value=key_str)
            except ValueError:
                logger.error(
                    f"Unrecognized backup setting {key_str} (with value: {value_data}). Skipping."
                )
                continue
            match key_enum:
                case BackupSetting.BACKUP_ALL:
                    self.backup_settings.backup_all = bool(value_data)
                case BackupSetting.BACKUP_DIRS:
                    backup_dirs_list: list[Directory] = []
                    if isinstance(value_data, list):
                        for v_item in value_data:
                            if isinstance(v_item, str):
                                try:
                                    dir_setting_val = DirSetting(value=v_item)
                                    if dir_setting_val in self.dirs:
                                        backup_dirs_list.append(
                                            self.dirs[dir_setting_val]
                                        )
                                except ValueError:
                                    logger.warning(
                                        f"Invalid DirSetting value '{v_item}' in backup_dirs. Skipping."
                                    )
                    self.backup_settings.backup_dirs = backup_dirs_list

                case BackupSetting.MAX_BACKUPS:
                    try:
                        self.backup_settings.max_backups = int(value_data)
                    except (ValueError, TypeError):
                        logger.warning(
                            f"Invalid value for MAX_BACKUPS: {value_data}. Using default {self.backup_settings.max_backups}"
                        )
                case BackupSetting.BACKUP_OVERVIEWS:
                    self.backup_settings.backup_overviews = bool(value_data)
                case _:
                    logger.warning(
                        f"Unrecognized backup setting {key_str} (with value: {value_data}). Skipping."
                    )
        self.backup_settings.backup_location = self.dirs.get(DirSetting.FULL_BACKUPS)

    def parse_directories(self, raw_dir_strs: dict[str, str]) -> None:
        """Parses the 'directories' section of settings.yaml.

        Args:
            raw_dir_strs: Dictionary mapping directory keys to path strings.
        """
        tmp_raw_copyright_data_full_path: str | None = (
            None  # Store path string temporarily
        )

        for key_str, path_str in raw_dir_strs.items():
            try:
                key_enum_val = DirSetting(value=key_str)
            except ValueError:
                logger.error(
                    f"Unrecognized directory type '{key_str}' (with path: '{path_str}'). Skipping."
                )
                continue
            try:
                if key_enum_val == DirSetting.RAW_COPYRIGHT_DATA_FULL:
                    # Store the relative path string, handle it after RAW_COPYRIGHT_DATA is parsed
                    tmp_raw_copyright_data_full_path = path_str
                else:
                    self.dirs[key_enum_val] = Directory(path=path_str)
            except Exception as e:
                logger.error(
                    f"Error while creating directory '{key_enum_val.value}' with path '{path_str}': {e}"
                )

        # Handle RAW_COPYRIGHT_DATA_FULL after RAW_COPYRIGHT_DATA is potentially set
        if (
            tmp_raw_copyright_data_full_path
            and DirSetting.RAW_COPYRIGHT_DATA in self.dirs
        ):
            base_path = self.dirs[DirSetting.RAW_COPYRIGHT_DATA].full
            self.dirs[DirSetting.RAW_COPYRIGHT_DATA_FULL] = Directory(
                path=base_path / tmp_raw_copyright_data_full_path
            )
        elif tmp_raw_copyright_data_full_path:
            # If RAW_COPYRIGHT_DATA was not in settings, RAW_COPYRIGHT_DATA_FULL might be an absolute path
            # or relative to CWD. For now, assume it's meant to be used as is if base is missing.
            logger.warning(
                f"'{DirSetting.RAW_COPYRIGHT_DATA.value}' not found, attempting to use '{tmp_raw_copyright_data_full_path}' directly for '{DirSetting.RAW_COPYRIGHT_DATA_FULL.value}'."
            )
            try:
                self.dirs[DirSetting.RAW_COPYRIGHT_DATA_FULL] = Directory(
                    path=tmp_raw_copyright_data_full_path
                )
            except Exception as e:
                logger.error(
                    f"Error while creating directory '{DirSetting.RAW_COPYRIGHT_DATA_FULL.value}' with path '{tmp_raw_copyright_data_full_path}': {e}"
                )

    def parse_files(self, raw_file_strs: dict[str, Any]) -> None:
        """Parses the 'files' section of settings.yaml.

        This includes the main script data folder and specific file paths.

        Args:
            raw_file_strs: Dictionary representing the 'files' settings.
        """
        folder_path_str = raw_file_strs.get("folder")
        script_data_dir_path: Path
        if isinstance(folder_path_str, str):
            script_data_dir = Directory(path=folder_path_str)
            self.dirs[DirSetting.SCRIPT_DATA] = script_data_dir
            script_data_dir_path = script_data_dir.full
        else:
            # Default to current working directory if "folder" is not specified
            cwd_dir = Directory(path=os.getcwd())
            self.dirs[DirSetting.SCRIPT_DATA] = cwd_dir
            script_data_dir_path = cwd_dir.full
            logger.info(
                f"'folder' not specified in 'files' settings, using CWD: {script_data_dir_path}"
            )

        files_section = raw_file_strs.get("files")
        if isinstance(files_section, dict):
            for file_key_str, file_name_str_any in files_section.items():
                if not isinstance(file_name_str_any, str):
                    logger.warning(
                        f"Expected string for file name, got {type(file_name_str_any)} for key '{file_key_str}'. Skipping."
                    )
                    continue
                file_name_str: str = file_name_str_any
                try:
                    file_key_enum = FileSetting(value=file_key_str)
                    self.files[file_key_enum] = File(
                        path=script_data_dir_path / file_name_str
                    )
                except ValueError:
                    logger.error(f"Unrecognized file type '{file_key_str}'. Skipping.")
                except Exception as e:
                    logger.error(f"Error while adding file '{file_key_str}': {e}")

        subfolders_section = raw_file_strs.get("subfolders")
        if isinstance(subfolders_section, dict):
            for subfolder_key_str, subfolder_name_any in subfolders_section.items():
                if not isinstance(subfolder_name_any, str):
                    logger.warning(
                        f"Expected string for subfolder name, got {type(subfolder_name_any)} for key '{subfolder_key_str}'. Skipping."
                    )
                    continue
                subfolder_name_str: str = subfolder_name_any
                try:
                    dir_key_enum = DirSetting(value=subfolder_key_str)
                    self.dirs[dir_key_enum] = Directory(
                        path=script_data_dir_path / subfolder_name_str
                    )
                except ValueError:
                    logger.error(
                        f"Unrecognized subfolder type '{subfolder_key_str}'. Skipping."
                    )
                except Exception as e:
                    logger.error(
                        f"Error while adding subfolder '{subfolder_name_str}': {e}"
                    )
        elif isinstance(subfolders_section, list):
            for subfolder_name_any in subfolders_section:
                if not isinstance(subfolder_name_any, str):
                    logger.warning(
                        f"Expected string in subfolder list, got {type(subfolder_name_any)}. Skipping."
                    )
                    continue
                subfolder_name_str: str = subfolder_name_any
                try:
                    dir_key_enum = DirSetting(value=subfolder_name_str)
                    self.dirs[dir_key_enum] = Directory(
                        path=script_data_dir_path / subfolder_name_str
                    )
                except ValueError:
                    logger.error(
                        f"Unrecognized subfolder type from list: '{subfolder_name_str}'. Skipping."
                    )
                except Exception as e:
                    logger.error(
                        f"Error while adding subfolder '{subfolder_name_str}' from list: {e}"
                    )

    def parse_unsorted(self, rest_values: dict[str, Any]) -> None:
        """Parses settings from the 'unsorted' key in settings.yaml.

        These are set as attributes on the Settings instance.
        Special handling for 'fine_amount', 'classification_options', and 'dashboard_reload'.

        Args:
            rest_values: Dictionary of unsorted settings.
        """
        for key, value in rest_values.items():
            if key == "fine_amount":
                parsed = safe_float(value)
                if parsed is not None:
                    self.fine_amount = parsed
                else:
                    logger.warning(
                        f"Could not parse 'fine_amount': {value} as float. Using default: {self.fine_amount}"
                    )
            elif key == "classification_options":
                if (
                    not self.classification_options
                ):  # Only set if not already set by data_settings standardization
                    if isinstance(value, list):
                        self.classification_options = [
                            str(opt)
                            for opt in value
                            if isinstance(opt, str | int | float)
                        ]
                    else:
                        logger.warning(
                            f"'classification_options' in unsorted settings is not a list: {value}"
                        )
            elif key == "dashboard_reload":
                if isinstance(value, bool):
                    self.dashboard_reload = value
                else:
                    logger.warning(
                        f"Could not parse 'dashboard_reload': {value} as bool. Using default: {self.dashboard_reload}"
                    )
            else:
                if hasattr(self, key):
                    logger.warning(
                        f"Unsorted key '{key}' clashes with an existing Settings attribute. Overwriting."
                    )
                setattr(self, key, value)

    def parse_database_settings(self, db_settings_value: Any) -> None:
        """Parses the database path from settings.

        Args:
            db_settings_value: The value associated with the database path setting.
                               Can be a string or a dictionary with a "path" key.
        """
        if isinstance(db_settings_value, str):
            self.db_path = Path(db_settings_value)
        elif (
            isinstance(db_settings_value, dict)
            and "path" in db_settings_value
            and isinstance(db_settings_value["path"], str)
        ):
            self.db_path = Path(db_settings_value["path"])
        else:
            logger.warning(
                f"Could not parse 'database_path': {db_settings_value}. Using default: {self.db_path}"
            )

    def _standardize_classification_options(self) -> None:
        """Standardizes classification options.

        Sets `self.classification_options` based on dropdown options from
        the 'manual_classification' column in `data_settings.data_entry_cols`.
        This ensures a single source of truth if defined there, otherwise uses
        options from 'unsorted' settings if available.
        """
        manual_classification_col: ColInfo | None = None
        for col_info in self.data_settings.data_entry_cols:
            if col_info.name == "manual_classification":
                manual_classification_col = col_info
                break

        if manual_classification_col and manual_classification_col.dropdown_options:
            options_list = [
                opt.strip()
                for opt in manual_classification_col.dropdown_options.split(",")
                if opt.strip()  # Ensure non-empty options
            ]
            if self.classification_options and set(options_list) != set(
                self.classification_options
            ):
                # logger.debug(
                #    "Overriding 'classification_options' with those from "
                #    "'data_settings.data_entry_cols.manual_classification.dropdown_options'."
                #    f"Selected options: {options_list}"
                # )
                self.classification_options = options_list
        elif (
            not self.classification_options
        ):  # Only warn if no options were set from 'unsorted' either
            logger.warning(
                "Could not find 'manual_classification' column with dropdown options in data_settings, "
                "and no 'classification_options' found in 'unsorted' settings. "
                "Classification options may be incomplete."
            )

    def __post_init__(self) -> None:
        """Initializes settings after dataclass creation.

        Loads settings from the YAML file and parses them.
        """
        self.settings_file = File(path=self.input_file_path)
        self.load()  # Populates self.raw_settings

        self.KEY_TO_PARSER_MAPPING = {
            "university": self.parse_university,
            "data_settings": self.parse_data_settings,
            "directories": self.parse_directories,
            "files": self.parse_files,
            "unsorted": self.parse_unsorted,
            "backup": self.parse_backup,
            "database_path": self.parse_database_settings,
        }

        if self.raw_settings:  # Check if loading was successful
            self.parse_settings()  # Parses based on KEY_TO_PARSER_MAPPING
            self._standardize_classification_options()  # Standardize after all parsing
        else:
            logger.error("Raw settings are empty. Cannot parse settings.")

    def load(self) -> None:
        """Loads settings from the YAML file specified by `self.settings_file`.

        Populates `self.raw_settings`. If loading fails, `self.raw_settings`
        will be an empty dictionary.
        """
        try:
            with open(file=self.settings_file.path, encoding="utf-8") as f:
                loaded_yaml = yaml.load(stream=f, Loader=yaml.FullLoader)
                if isinstance(loaded_yaml, dict):
                    self.raw_settings = loaded_yaml
                else:
                    logger.error(
                        f"Settings file {self.settings_file} did not load as a dictionary. Loaded: {type(loaded_yaml)}"
                    )
                    self.raw_settings = {}
        except FileNotFoundError:
            logger.error(f"Settings file not found: {self.settings_file.path}")
            self.raw_settings = {}
        except yaml.YAMLError as e:
            logger.error(f"Error parsing YAML from {self.settings_file.path}: {e}")
            self.raw_settings = {}
        except Exception as e:
            logger.error(
                f"Unexpected error loading settings from {self.settings_file.path}: {e}"
            )
            self.raw_settings = {}

    def parse_settings(self) -> None:
        """Parses all settings from the loaded `self.raw_settings`.

        Iterates through `self.raw_settings` and calls the appropriate
        parser method based on `self.KEY_TO_PARSER_MAPPING`.
        Backup settings are parsed last to ensure `self.dirs` is initialized.
        """
        if not self.raw_settings:
            logger.warning("Cannot parse settings: raw_settings is empty.")
            return

        # Parse all keys except 'backup' first
        for key, value in self.raw_settings.items():
            if key == "backup":
                continue  # Defer backup parsing

            parser = self.KEY_TO_PARSER_MAPPING.get(key)
            if parser:
                try:
                    parser(value)
                except Exception as e:
                    logger.error(
                        f"Error parsing setting key '{key}' with value '{value}': {e}"
                    )
            else:
                logger.warning(
                    f"Unrecognized setting key: '{key}'. It will be ignored unless handled by 'unsorted'."
                )

        # Parse 'backup' settings last, ensuring dependencies like 'directories' are parsed
        if "backup" in self.raw_settings:
            try:
                self.parse_backup(backup_settings_yaml=self.raw_settings["backup"])
            except Exception as e:
                logger.error(f"Error parsing 'backup' settings: {e}")
                logger.debug(self.raw_settings["backup"])


@dataclass
class SampleSettings:
    """Dataclass holding sample generation settings, loaded from a YAML file.

    Attributes:
        input_file_path: Path to the sample settings YAML file (default: "sample.yaml").
        settings_file: File object representing the sample settings file.
        raw_settings: Raw dictionary loaded from the YAML file.
        input: Parsed settings for input data generation.
        output: Parsed settings for output data generation.
        KEY_TO_PARSER_MAPPING: Internal mapping of setting keys to parser methods.
    """

    input_file_path: str = "sample.yaml"
    settings_file: File = field(init=False)
    raw_settings: dict[str, Any] | None = None  # Can be None if loading fails
    input: dict[str, Any] = field(default_factory=dict, init=False)
    output: dict[str, Any] = field(default_factory=dict, init=False)
    KEY_TO_PARSER_MAPPING: dict[str, Callable[[dict[str, Any]], None]] = field(
        init=False, repr=False
    )

    def __post_init__(self) -> None:
        """Initializes sample settings after dataclass creation."""
        self.settings_file = File(path=self.input_file_path)
        self.load()

        self.KEY_TO_PARSER_MAPPING = {
            "input": self.parse_input,
            "output": self.parse_output,
        }

        if self.raw_settings:
            self.parse_settings()
        else:
            logger.warning(
                f"Sample settings raw_settings is None for {self.input_file_path}. Skipping parsing."
            )

    def load(self) -> None:
        """Loads sample settings from the YAML file.

        Populates `self.raw_settings`. If loading fails, `self.raw_settings` is set to None.
        """
        try:
            with open(file=self.settings_file.path, encoding="utf-8") as f:
                loaded_yaml = yaml.load(stream=f, Loader=yaml.FullLoader)
                if isinstance(loaded_yaml, dict):
                    self.raw_settings = loaded_yaml
                else:
                    logger.error(
                        f"Sample settings file {self.settings_file.path} did not load as a dictionary."
                    )
                    self.raw_settings = None  # Explicitly None on failure
        except FileNotFoundError:
            logger.error(f"Sample settings file not found: {self.settings_file.path}")
            self.raw_settings = None
        except yaml.YAMLError as e:
            logger.error(
                f"Error parsing YAML from sample settings file {self.settings_file.path}: {e}"
            )
            self.raw_settings = None
        except Exception as e:
            logger.error(
                f"Unexpected error loading sample settings from {self.settings_file.path}: {e}"
            )
            self.raw_settings = None

    def parse_settings(self) -> None:
        """Parses all sample settings from `self.raw_settings`."""
        if not self.raw_settings:
            logger.warning(
                "Cannot parse sample settings: raw_settings is empty or None."
            )
            return

        for key, value in self.raw_settings.items():
            parser = self.KEY_TO_PARSER_MAPPING.get(key)
            if parser:
                if isinstance(value, dict):
                    try:
                        parser(value)  # Call the bound method with the value dict
                    except Exception as e:
                        logger.error(f"Error parsing sample setting key '{key}': {e}")
                else:
                    logger.warning(
                        f"Expected dictionary for sample setting key '{key}', got {type(value)}. Skipping."
                    )
            else:
                logger.warning(
                    f"[SampleSettings] Unrecognized key: '{key}'. Setting as attribute (if new)."
                )
                if not hasattr(
                    self, key
                ):  # Avoid overwriting existing attributes like 'input', 'output'
                    setattr(self, key, value)
                else:
                    logger.warning(
                        f"[SampleSettings] Key '{key}' conflicts with existing attribute. Not set from top level."
                    )

    def parse_input(self, data_dict: dict[str, Any]) -> None:
        """Parses the 'input' section of sample.yaml.

        Args:
            data_dict: The dictionary representing the 'input' settings.
        """
        self.input = {"file": "", "filters": []}  # Initialize with defaults

        file_val = data_dict.get("file")
        if isinstance(file_val, str):
            self.input["file"] = file_val
        elif file_val is not None:
            logger.warning(
                f"Sample settings 'input.file' expected a string, got {type(file_val)}."
            )

        filters_val = data_dict.get("filters")
        if isinstance(filters_val, list):
            self.input["filters"] = (
                filters_val  # Assuming filters are correctly structured
            )
        elif filters_val is not None:
            logger.warning(
                f"Sample settings 'input.filters' expected a list, got {type(filters_val)}."
            )

    def parse_output(self, data_dict: dict[str, Any]) -> None:
        """Parses the 'output' section of sample.yaml.

        Args:
            data_dict: The dictionary representing the 'output' settings.
        """
        self.output = {  # Default structure
            "file": "",
            "filters": [],
            "selection": [],
            "columns": [],
            "max_rows": 0,
            "remove_duplicates": True,
        }

        for key, default_value in self.output.items():
            if key in data_dict:
                val = data_dict[key]
                expected_type = type(default_value)
                if isinstance(val, expected_type):
                    self.output[key] = val
                elif key == "max_rows" and isinstance(val, int | float):
                    self.output[key] = int(val)
                elif key == "remove_duplicates" and isinstance(val, bool | int):
                    self.output[key] = bool(val)
                else:
                    logger.warning(
                        f"Sample settings 'output.{key}' expected type {expected_type}, got {type(val)}. Using default."
                    )


def load_osiris_data() -> dict[str, Any] | None:
    """Loads Osiris data from the JSON file specified in settings.

    Returns:
        A dictionary containing Osiris data if successful, otherwise None.
        The dictionary structure is expected to be `dict[str, Any]`.
    """
    osiris_file_path_obj = SETTINGS.files.get(FileSetting.OSIRIS_DATA_W_CONTACTS)
    if not osiris_file_path_obj:
        logger.error(
            f"Osiris data file path not found in settings ('{FileSetting.OSIRIS_DATA_W_CONTACTS.value}'). "
            "OSIRIS data enrichment will not be possible."
        )
        return None

    osiris_file_path = osiris_file_path_obj.path

    try:
        with open(file=osiris_file_path, encoding="utf-8") as f:
            data: Any = json.load(f)
        if not isinstance(data, dict):
            logger.error(
                f"Osiris data file ({osiris_file_path}) does not contain a valid JSON object (expected dict). "
                "OSIRIS data enrichment will not be possible."
            )
            return None
        return {str(k): v for k, v in data.items()}
    except FileNotFoundError:
        logger.error(
            f"Osiris data file ({osiris_file_path}) not found. "
            "OSIRIS data enrichment will not be possible.\n"
            "Please run the CLI again with the refresh_osiris_data flag set to True to retrieve the required data."
        )
        return None
    except json.JSONDecodeError:
        logger.error(
            f"Error decoding JSON from Osiris data file ({osiris_file_path}). "
            "OSIRIS data enrichment will not be possible."
        )
        return None
    except Exception as e:
        logger.error(
            f"An unexpected error occurred while loading Osiris data from {osiris_file_path}: {e}. "
            "OSIRIS data enrichment will not be possible."
        )
        return None


# Global settings

# Set up logging configuration
# install rich traceback as default
install(show_locals=True)

# initialize settings from (default: read from 'settings.yaml')

SETTINGS: Settings = Settings()
# SAMPLESETTINGS: SampleSettings = SampleSettings()

# create global variables from certain settings for easier access [NOTE: remove these and replace with SETTINGS.<attr> in the future for consistency/robustness]
DEPARTMENT_MAPPING: dict[str, str] = SETTINGS.university_settings.department_mapping
COURSE_MAPPING: dict[str, dict[str, str]] = SETTINGS.university_settings.course_mapping
FINE_AMOUNT: float = SETTINGS.fine_amount
OSIRIS_DATA: dict[str, Any] | None = load_osiris_data()
