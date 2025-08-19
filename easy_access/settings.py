import json
import logging  # Already imported, ensuring it's here
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Literal  # Added TypeAlias for older Python if needed

import yaml

from easy_access.utils import Directory, File

# Standard Python logging
std_logger = logging.getLogger(__name__)


"""
This module defines dataclasses and enums for managing application settings,
primarily loaded from a `settings.yaml` file. It provides a structured way to
access configuration for directories, files, data processing parameters,
university-specific information, and backup preferences.

The main class, `Settings`, orchestrates the loading and parsing of these
configurations. Other dataclasses like `EasyAccessSettings`, `DataSettings`,
`UniversitySettings`, etc., represent specific sections of the settings.
Enums like `DirSetting` and `FileSetting` provide controlled vocabulary for keys.
"""

# Type alias for more complex dict structures if needed, though Any is often used.
# JsonValue: TypeAlias = str | int | float | bool | None | list['JsonValue'] | dict[str, 'JsonValue']


@dataclass
class ColInfo:
    """
    Describes metadata for a single column, typically used in constructing data entry sheets.

    Attributes:
        name (str): The original column name as it appears in data sources or internal DataFrames.
        dropdown_options (str): A string representing options for a dropdown list in a sheet
                                (e.g., a comma-separated list for Excel data validation). Empty if no dropdown.
        is_url (bool): If True, the column's content should be formatted as a hyperlink.
        is_new (bool): If True, this column is added during processing and is not in the original data.
        is_editable (bool): If True, this column is intended to be user-editable in output sheets.
        new_name (str): If provided, the column will be renamed to this value in output sheets.
        default_val (str): For new columns (`is_new` is True), this value is used as the default.
        max_width (int): Tracks the maximum character width encountered for this column's data during processing.
                         Used for column width auto-sizing. Initialized to a default minimum.
        count_max_width_over_40 (int): Counts how many cells in this column exceed a width of 40 characters.
                                       Used for heuristics in deciding whether to wrap text.
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
        """bool: True if `dropdown_options` is not empty, indicating a dropdown should be used."""
        return len(self.dropdown_options) > 0


class DirSetting(Enum):
    """Enumerates keys for standard directory paths used throughout the application."""

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
    """Enumerates keys for standard file paths used throughout the application."""

    FULL_DATA_CSV = "full_data_csv"
    FULL_DATA_PARQUET = "full_data_parquet"
    OSIRIS_DATA = "osiris_data"
    OSIRIS_DATA_W_CONTACTS = "osiris_data_w_contacts"
    PERSON_DATA = "person_data"


class BackupSetting(Enum):
    """Enumerates keys for backup-related settings."""

    BACKUP_ALL = "backup_all"
    BACKUP_DIRS = "backup_dirs"
    MAX_BACKUPS = "max_backups"
    BACKUP_OVERVIEWS = "backup_overviews"


class SheetSetting(Enum):
    """Enumerates keys for sheet-specific data settings (column orders, sheet names, etc.)."""

    RAW_DATA_COL_ORDER = "raw_data_col_order"
    FINAL_DATA_COL_ORDER = "final_data_col_order"
    COMPLETE_DATA_NAME = "complete_data_name"
    COMPLETE_DATA_COLS = (
        "complete_data_cols"  # Note: Seems unused in current parsing logic
    )
    DATA_ENTRY_NAME = "data_entry_name"
    DATA_ENTRY_COLS = "data_entry_cols"
    NEW_FIELDS = "new_fields"


class Functions(str, Enum):  # Inherits from str to allow direct use as string values
    """
    Defines options for controlling which main functions of the tool are run,
    typically used with CLI parameters.
    """

    both: str = "both"  # Represents running both read and export/processing functions
    read: str = "read"
    export: str = "export"


@dataclass
class EasyAccessSettings:
    """
    Runtime configuration settings for the Easy Access Tool, often derived from CLI parameters or defaults.

    Attributes:
        export (bool): If True, indicates that an export operation should be performed.
        only_changes (bool): If True, processing might be limited to only new or changed items.
        refresh_osiris_data (bool): If True, triggers a refresh of Osiris-related data.
        only_retrieve_missing_osiris_data (bool): If True (and `refresh_osiris_data` is True),
                                                 limits Osiris refresh to only items missing this data.
        other_sheet (Path | None): Path to an alternative Excel sheet to use as a primary data source,
                                   instead of the default raw copyright data.
        enrich_with_osiris_data (bool): If True, enables enrichment of data with Osiris information.
        dirs (dict[DirSetting, Directory]): Dictionary mapping `DirSetting` enums to `Directory` objects,
                                           providing paths to various working directories.
        disable_writes (bool): If True, all operations that write files to disk should be disabled.
        faculty (str | None): If set, processing will be limited to this specific faculty.
    """

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
    def from_env(cls, **kwargs: Any) -> "EasyAccessSettings":
        """
        Creates an EasyAccessSettings instance, primarily using directory settings
        from the global SETTINGS object and overriding other values with provided kwargs.

        Args:
            **kwargs: Keyword arguments to override default EasyAccessSettings attributes.

        Returns:
            EasyAccessSettings: An instance of EasyAccessSettings.
        """
        dirs_from_global: dict[DirSetting, Directory] = SETTINGS.dirs
        return cls(dirs=dirs_from_global, **kwargs)


@dataclass
class DataSettings:
    """
    Holds settings related to data structure, column names, and sheet configurations.
    These are typically initialized from `settings.yaml`.
    """

    data_entry_cols: list[ColInfo] = field(default_factory=list, init=False)
    complete_data_cols: list[ColInfo] = field(
        default_factory=list, init=False
    )  # Currently not explicitly parsed
    complete_data_name: str = field(default="Complete Data", init=False)
    data_entry_name: str = field(default="Data Entry", init=False)
    raw_data_col_order: list[str] = field(default_factory=list, init=False)
    final_data_col_order: list[str] = field(default_factory=list, init=False)
    new_fields: dict[str, dict[str, Any]] = field(default_factory=dict, init=False)


@dataclass
class BackupSettings:
    """Holds settings related to data backup operations."""

    backup_all: bool = True
    backup_dirs: list[Directory] = field(default_factory=list)
    max_backups: int = 3
    backup_overviews: bool = True
    backup_location: Directory | None = field(default=None)


@dataclass(frozen=True)
class SettingsProgramme:
    """Represents a university programme with its associated metadata, loaded from settings."""

    name: str | None = None
    abbreviation: str | None = None
    programme_type: Literal["b", "m", "o"] = (
        "o"  # b=bachelor, m=master, o=other/unknown
    )
    cluster: str | None = None
    faculty_name: str | None = (
        None  # Populated by UniversitySettings.make_programme_set
    )
    faculty_abbreviation: str | None = (
        None  # Populated by UniversitySettings.make_programme_set
    )


@dataclass
class SettingsFaculty:
    """Represents a university faculty and its list of programmes, loaded from settings."""

    name: str
    abbreviation: str = ""
    programmes: list[SettingsProgramme] = field(default_factory=list)


@dataclass
class UniversitySettings:
    """
    Holds university-specific settings like names, abbreviations, faculty structures,
    and programme details, loaded from `settings.yaml`.
    """

    name: str = field(default="", init=False)
    abbreviation: str = field(default="", init=False)
    lms: dict[str, str] = field(default_factory=dict, init=False)
    course_catalogue: dict[str, str] = field(default_factory=dict, init=False)
    employee_catalogue: dict[str, str] = field(default_factory=dict, init=False)
    faculties: list[SettingsFaculty] = field(default_factory=list, init=False)
    programmes: set[SettingsProgramme] = field(default_factory=set, init=False)

    def make_programme_set(self) -> None:
        """
        Populates the `self.programmes` set from the `self.faculties` list,
        enriching each programme with its parent faculty's name and abbreviation.
        Ensures unique programme entries.
        """
        self.programmes.clear()
        if self.faculties:
            for faculty in self.faculties:
                faculty_programmes: list[SettingsProgramme] = faculty.programmes
                if not faculty_programmes:
                    continue
                for programme in faculty_programmes:
                    # Create a new SettingsProgramme ensuring all fields are present
                    prog_data = {
                        "name": programme.name,
                        "abbreviation": programme.abbreviation,
                        "programme_type": programme.programme_type,
                        "cluster": programme.cluster,
                        "faculty_name": faculty.name,
                        "faculty_abbreviation": faculty.abbreviation,
                    }
                    self.programmes.add(SettingsProgramme(**prog_data))

    @property
    def faculty_abbreviations(self) -> set[str]:
        """Set of all faculty abbreviations."""
        return {
            faculty.abbreviation for faculty in self.faculties if faculty.abbreviation
        }

    @property
    def department_mapping(self) -> dict[str, str]:
        """
        Generates a mapping from a descriptive programme string (used as 'department' in some contexts)
        to its faculty abbreviation.
        """
        if not self.programmes:  # Ensure programmes are populated
            self.make_programme_set()

        mapping_data: dict[str, str] = {
            f"{programme.abbreviation + ': ' if programme.abbreviation else ''}{programme.name}": programme.faculty_abbreviation
            or ""
            for programme in self.programmes
            if programme.name  # Ensure programme.name is not None
        }
        # Manual additions/overrides
        # TODO: Consider moving these manual mappings to settings.yaml if they change often.
        manual_overrides: dict[str, str] = {
            "Master Risicomanagement": "BMS",  # Assuming BMS, adjust if needed
            "Master Public Management": "BMS",  # Assuming BMS, adjust if needed
            "BMS: Behavioural, Management and Social Sciences": "BMS",
            "EEMCS: Electrical Engineering, Mathematics and Computer Science": "EEMCS",
            "ET: Engineering Technology": "ET",
        }
        mapping_data.update(manual_overrides)
        return mapping_data

    @property
    def course_mapping(self) -> dict[str, dict[str, str]]:
        """
        Generates a mapping from faculty abbreviation to a dictionary of its programme names
        (department like) to cluster names.
        """
        if not self.programmes:  # Ensure programmes are populated
            self.make_programme_set()

        course_map: dict[str, dict[str, str]] = {}
        for faculty in self.faculties:
            if not faculty.abbreviation:
                continue  # Skip faculty if no abbreviation

            faculty_programme_map: dict[str, str] = {}
            for programme in faculty.programmes:
                if (
                    programme.cluster and programme.name
                ):  # Ensure cluster and name exist
                    prog_key = f"{programme.abbreviation + ': ' if programme.abbreviation else ''}{programme.name}"
                    faculty_programme_map[prog_key] = programme.cluster

            if (
                faculty_programme_map
            ):  # Only add faculty if it has programmes with clusters
                course_map[faculty.abbreviation] = faculty_programme_map
        return course_map


@dataclass
class Settings:
    """
    Main application settings class, loaded from `settings.yaml`.
    Orchestrates parsing of various settings sections into structured dataclasses.
    """

    input_file_path: str = "settings.yaml"
    settings_file: File = field(init=False, repr=False)
    raw_settings: dict[str, Any] = field(default_factory=dict, init=False, repr=False)
    dirs: dict[DirSetting, Directory] = field(default_factory=dict, init=False)
    files: dict[FileSetting, File] = field(default_factory=dict, init=False)
    fine_amount: float = field(default=0.3, init=False, repr=False)
    data_settings: DataSettings = field(default_factory=DataSettings, init=False)
    university_settings: UniversitySettings = field(
        default_factory=UniversitySettings, init=False
    )
    backup_settings: BackupSettings = field(default_factory=BackupSettings, init=False)
    classification_options: list[str] = field(
        default_factory=list, init=False
    )  # Parsed under 'unsorted'
    KEY_TO_PARSER_MAPPING: dict[str, Any] = field(default_factory=dict, init=False)

    def __post_init__(self) -> None:
        """Loads and parses settings after initialization."""
        self.KEY_TO_PARSER_MAPPING = {
            "university": self.parse_university,
            "data_settings": self.parse_data_settings,
            "directories": self.parse_directories,
            "files": self.parse_files,
            "unsorted": self.parse_unsorted,  # Must be parsed after others that might use its values implicitly
            "backup": self.parse_backup,
        }
        self.settings_file = File(path=self.input_file_path)
        self.load()
        if self.raw_settings:
            self.parse_settings()

    def load(self) -> None:
        """Loads settings from the YAML file specified by `input_file_path`."""
        try:
            with open(file=self.settings_file.path, encoding="utf-8") as f:
                self.raw_settings = yaml.load(stream=f, Loader=yaml.FullLoader)
        except FileNotFoundError:
            std_logger.error(
                f"Settings file not found: {self.settings_file.path}. Default settings will be used."
            )
            self.raw_settings = {}
        except yaml.YAMLError as e:
            std_logger.error(
                f"Error parsing YAML settings from {self.settings_file.path}: {e}"
            )
            self.raw_settings = {}
        except Exception as e:
            std_logger.error(
                f"Unexpected error loading settings from {self.settings_file.path}: {e}"
            )
            self.raw_settings = {}

    def parse_settings(self) -> None:
        """Parses raw settings from the loaded YAML into structured dataclass fields."""
        if not isinstance(self.raw_settings, dict):  # Ensure raw_settings is a dict
            std_logger.error(
                f"Failed to load settings as a dictionary. Raw settings type: {type(self.raw_settings)}"
            )
            return

        for key, value in self.raw_settings.items():
            if key == "backup":  # Defer parsing backup until dirs are parsed
                continue
            parser_method = self.KEY_TO_PARSER_MAPPING.get(key)
            if parser_method:
                try:
                    parser_method(self, value)
                except Exception as e:
                    std_logger.error(
                        f"Error parsing settings section '{key}': {e}. Skipping this section."
                    )
            else:
                std_logger.warning(
                    f"Unrecognized settings key: '{key}'. Value will be set as a direct attribute if not handled by 'unsorted'."
                )
                # setattr(self, key, value) # This is handled by parse_unsorted if key is 'unsorted'

        # Parse backup settings after other sections (especially directories)
        if "backup" in self.raw_settings:
            self.parse_backup(backup_settings=self.raw_settings["backup"])

        # Handle 'unsorted' which might include classification_options
        if "unsorted" in self.raw_settings and isinstance(
            self.raw_settings["unsorted"], dict
        ):
            self.parse_unsorted(self.raw_settings["unsorted"])

    def parse_university(self, value: dict[str, Any]) -> None:
        """Parses the 'university' section of settings."""
        self.university_settings.name = str(value.get("name", ""))
        self.university_settings.abbreviation = str(value.get("abbreviation", ""))
        self.university_settings.lms = value.get("lms", {})
        self.university_settings.course_catalogue = value.get("course_catalogue", {})
        self.university_settings.employee_catalogue = value.get(
            "employee_catalogue", {}
        )

        parsed_faculties: list[SettingsFaculty] = []
        for faculty_data in value.get("faculties", []):
            if not isinstance(faculty_data, dict):
                continue
            name: str = faculty_data.get("name", "")
            abbreviation: str = faculty_data.get("abbreviation", "")
            programmes_data = faculty_data.get("programmes", [])
            parsed_programmes: list[SettingsProgramme] = []
            if isinstance(programmes_data, list):
                for prog_data in programmes_data:
                    if isinstance(prog_data, dict):
                        # Ensure all fields for SettingsProgramme are present or defaulted
                        parsed_programmes.append(
                            SettingsProgramme(
                                name=prog_data.get("name"),
                                abbreviation=prog_data.get("abbreviation"),
                                programme_type=prog_data.get("programme_type", "o"),  # type: ignore
                                cluster=prog_data.get("cluster"),
                            )
                        )
            parsed_faculties.append(
                SettingsFaculty(
                    name=name, abbreviation=abbreviation, programmes=parsed_programmes
                )
            )
        self.university_settings.faculties = parsed_faculties
        self.university_settings.make_programme_set()

    def parse_data_settings(self, data_settings_val: dict[str, Any]) -> None:
        """Parses the 'data_settings' section of settings."""
        for key_str, value in data_settings_val.items():
            try:
                setting_key = SheetSetting(value=key_str)
            except ValueError:
                std_logger.warning(
                    f"Unrecognized data setting key '{key_str}'. Skipping."
                )
                continue

            if setting_key == SheetSetting.DATA_ENTRY_COLS and isinstance(value, list):
                self.data_settings.data_entry_cols = [
                    ColInfo(**col) for col in value if isinstance(col, dict)
                ]
            elif setting_key == SheetSetting.COMPLETE_DATA_COLS and isinstance(
                value, list
            ):  # Not used in current code but parse if present
                self.data_settings.complete_data_cols = [
                    ColInfo(**col) for col in value if isinstance(col, dict)
                ]
            elif setting_key == SheetSetting.COMPLETE_DATA_NAME and isinstance(
                value, str
            ):
                self.data_settings.complete_data_name = value
            elif setting_key == SheetSetting.DATA_ENTRY_NAME and isinstance(value, str):
                self.data_settings.data_entry_name = value
            elif setting_key == SheetSetting.RAW_DATA_COL_ORDER and isinstance(
                value, list
            ):
                self.data_settings.raw_data_col_order = [
                    str(col_name) for col_name in value
                ]
            elif setting_key == SheetSetting.FINAL_DATA_COL_ORDER and isinstance(
                value, list
            ):
                self.data_settings.final_data_col_order = [
                    str(col_name) for col_name in value
                ]
            elif setting_key == SheetSetting.NEW_FIELDS and isinstance(value, dict):
                parsed_new_fields: dict[str, dict[str, Any]] = {}
                for colname, settings_dict in value.items():
                    if isinstance(settings_dict, dict):
                        field_spec: dict[str, Any] = {}
                        if "values" in settings_dict:
                            field_spec["values"] = settings_dict["values"]
                        if "default" in settings_dict:
                            field_spec["default"] = settings_dict["default"]
                        parsed_new_fields[str(colname)] = field_spec
                self.data_settings.new_fields = parsed_new_fields
            else:
                std_logger.warning(
                    f"Skipping data setting '{key_str}' due to unexpected value type or unhandled case."
                )

    def parse_backup(self, backup_settings: dict[str, Any]) -> None:
        """Parses the 'backup' section of settings."""
        for key_str, value in backup_settings.items():
            try:
                setting_key = BackupSetting(value=key_str)
            except ValueError:
                std_logger.warning(
                    f"Unrecognized backup setting key '{key_str}'. Skipping."
                )
                continue

            if setting_key == BackupSetting.BACKUP_ALL:
                self.backup_settings.backup_all = bool(value)
            elif setting_key == BackupSetting.BACKUP_DIRS and isinstance(value, list):
                self.backup_settings.backup_dirs = []  # Initialize/clear
                for dir_key_str in value:
                    try:
                        dir_enum_val = DirSetting(str(dir_key_str))
                        if dir_enum_val in self.dirs:
                            self.backup_settings.backup_dirs.append(
                                self.dirs[dir_enum_val]
                            )
                        else:
                            std_logger.warning(
                                f"Directory key '{dir_key_str}' in backup_dirs not found in parsed self.dirs."
                            )
                    except ValueError:
                        std_logger.warning(
                            f"Invalid DirSetting key '{dir_key_str}' in backup_dirs."
                        )
            elif setting_key == BackupSetting.MAX_BACKUPS:
                self.backup_settings.max_backups = int(value)
            elif setting_key == BackupSetting.BACKUP_OVERVIEWS:
                self.backup_settings.backup_overviews = bool(value)
            else:
                std_logger.warning(f"Unhandled backup setting '{key_str}'.")

        self.backup_settings.backup_location = self.dirs.get(DirSetting.FULL_BACKUPS)

    def parse_directories(self, raw_dir_strs: dict[str, str]) -> None:
        """Parses the 'directories' section of settings."""
        tmp_raw_full_data_path: str | None = None
        for key_str, path_str in raw_dir_strs.items():
            try:
                dir_key = DirSetting(value=key_str)
            except ValueError:
                std_logger.warning(
                    f"Unrecognized directory key '{key_str}' in settings. Skipping."
                )
                continue

            try:
                if dir_key == DirSetting.RAW_COPYRIGHT_DATA_FULL:
                    tmp_raw_full_data_path = path_str  # Store temporarily
                else:
                    self.dirs[dir_key] = Directory(path=path_str)
            except Exception as e:
                std_logger.error(
                    f"Error creating Directory object for key '{dir_key}' with path '{path_str}': {e}"
                )

        # Handle RAW_COPYRIGHT_DATA_FULL specifically after RAW_COPYRIGHT_DATA is parsed
        if tmp_raw_full_data_path and DirSetting.RAW_COPYRIGHT_DATA in self.dirs:
            try:
                self.dirs[DirSetting.RAW_COPYRIGHT_DATA_FULL] = Directory(
                    path=self.dirs[DirSetting.RAW_COPYRIGHT_DATA].full
                    / tmp_raw_full_data_path
                )
            except Exception as e:
                std_logger.error(
                    f"Error creating Directory for RAW_COPYRIGHT_DATA_FULL: {e}"
                )
        elif tmp_raw_full_data_path:
            std_logger.warning(
                "RAW_COPYRIGHT_DATA_FULL path found but RAW_COPYRIGHT_DATA base directory not yet parsed. RAW_COPYRIGHT_DATA_FULL might be incorrect."
            )

    def parse_files(self, raw_file_config: dict[str, Any]) -> None:
        """Parses the 'files' section of settings (which includes 'folder', 'files' dict, 'subfolders' list)."""
        script_data_folder_str = raw_file_config.get("folder")
        if isinstance(script_data_folder_str, str):
            self.dirs[DirSetting.SCRIPT_DATA] = Directory(path=script_data_folder_str)
        else:
            self.dirs[DirSetting.SCRIPT_DATA] = Directory(
                path=Path.cwd()
            )  # Default if no folder specified

        base_path: Path = self.dirs[DirSetting.SCRIPT_DATA].full

        files_dict = raw_file_config.get("files", {})
        if isinstance(files_dict, dict):
            for key_str, file_path_str in files_dict.items():
                try:
                    file_key = FileSetting(value=key_str)
                    self.files[file_key] = File(path=base_path / str(file_path_str))
                except ValueError:
                    std_logger.warning(
                        f"Unrecognized file key '{key_str}' in files settings. Skipping."
                    )
                except Exception as e:
                    std_logger.error(f"Error creating File object for '{key_str}': {e}")

        subfolders_list = raw_file_config.get("subfolders", [])
        if isinstance(subfolders_list, list):
            for subfolder_key_str in subfolders_list:
                try:
                    dir_key = DirSetting(value=str(subfolder_key_str))
                    self.dirs[dir_key] = Directory(
                        path=base_path / str(subfolder_key_str)
                    )
                except ValueError:
                    std_logger.warning(
                        f"Unrecognized DirSetting key '{subfolder_key_str}' in subfolders. Skipping."
                    )
                except Exception as e:
                    std_logger.error(
                        f"Error creating Directory for subfolder '{subfolder_key_str}': {e}"
                    )

    def parse_unsorted(self, unsorted_values: dict[str, Any]) -> None:
        """Sets attributes from the 'unsorted' section directly on the Settings object."""
        for key, value in unsorted_values.items():
            if key == "classification_options" and isinstance(value, list):
                self.classification_options = [str(opt) for opt in value]
            else:
                setattr(self, key, value)
                std_logger.debug(f"Set unsorted attribute '{key}' to Settings object.")


'''
@dataclass
class SampleSettings:
    """Dataclass for handling sample settings, typically for testing or examples."""

    input_file_path: str = "sample.yaml"
    raw_settings: dict[str, Any] | None = None  # Loaded from YAML
    input: dict[str, Any] = field(default_factory=dict, init=False)
    output: dict[str, Any] = field(default_factory=dict, init=False)
    settings_file: File = field(init=False, repr=False)

    def __post_init__(self) -> None:
        """Post-initialization hook to load and parse settings."""
        self.settings_file = File(path=self.input_file_path)
        self.load()
        if self.raw_settings:
            self.parse_settings()

    def load(self) -> None:
        """Load the settings from the specified YAML file."""
        try:
            with open(file=self.settings_file.path, encoding="utf-8") as f:
                self.raw_settings = yaml.load(stream=f, Loader=yaml.FullLoader)
        except FileNotFoundError:
            std_logger.error(
                f"Sample settings file not found: {self.settings_file.path}."
            )
            self.raw_settings = {}
        except yaml.YAMLError as e:
            std_logger.error(
                f"Error parsing YAML from sample settings file {self.settings_file.path}: {e}"
            )
            self.raw_settings = {}
        except Exception as e:
            std_logger.error(
                f"Unexpected error loading sample settings from {self.settings_file.path}: {e}"
            )
            self.raw_settings = {}

    def parse_settings(self) -> None:
        """Parse all settings from the loaded YAML raw_settings."""
        if not isinstance(self.raw_settings, dict):
            std_logger.error(
                f"Failed to load sample settings as a dictionary. Raw settings type: {type(self.raw_settings)}"
            )
            return

        for key, value in self.raw_settings.items():
            parser_method = self.KEY_TO_PARSER_MAPPING.get(key)
            if parser_method:
                try:
                    parser_method(self, value)
                except Exception as e:
                    std_logger.error(
                        f"Error parsing sample settings section '{key}': {e}. Skipping."
                    )
            else:
                std_logger.warning(
                    f"[SampleSettings] Unrecognized key: '{key}'. Value will be set as direct attribute."
                )
                setattr(self, key, value)

    def parse_input(self, data_dict: dict[str, Any]) -> None:
        """
        Parse the settings for the key 'input' in sample.yaml.
        (Currently a placeholder with basic structure).
        """
        self.input = {
            "file": data_dict.get("file", ""),
            "filters": data_dict.get("filters", []),
        }
        if not data_dict.get("file"):  # Log if essential parts are missing
            std_logger.warning('Sample settings "input" section is missing "file" key.')

    def parse_output(self, data_dict: dict[str, Any]) -> None:
        """
        Parse the settings for the key 'output' in sample.yaml.
        (Currently a placeholder with basic structure).
        """
        self.output = {
            "file": data_dict.get("file", ""),
            "filters": data_dict.get("filters", []),
            "selection": data_dict.get("selection", []),
            "columns": data_dict.get("columns", []),
            "max_rows": data_dict.get("max_rows", 0),
            "remove_duplicates": data_dict.get("remove_duplicates", True),
        }
        if not data_dict.get("file"):
            std_logger.warning(
                'Sample settings "output" section is missing "file" key.'
            )
'''


def load_osiris_data() -> dict[str, dict[str, Any]] | None:
    """
    Loads Osiris data from the JSON file specified in settings.

    Returns:
        dict[str, dict[str, Any]] | None: Parsed Osiris data as a dictionary,
                                         or None if the file is not found or unreadable.
    """
    try:
        # Ensure SETTINGS and its 'files' attribute are initialized
        if (
            not hasattr(SETTINGS, "files")
            or not isinstance(SETTINGS.files, dict)
            or FileSetting.OSIRIS_DATA_W_CONTACTS not in SETTINGS.files
        ):
            std_logger.error(
                "SETTINGS.files not properly initialized or OSIRIS_DATA_W_CONTACTS key missing."
            )
            return None

        osiris_file: File = SETTINGS.files[FileSetting.OSIRIS_DATA_W_CONTACTS]
        if not osiris_file.exists:  # Use File object's exists method
            std_logger.error(
                f"{osiris_file.path} not found. OSIRIS data enrichment will not be possible."
            )
            return None

        with open(file=osiris_file.path, encoding="utf-8") as fp:
            return json.load(fp=fp)
    except FileNotFoundError:
        std_logger.error(  # Should be caught by osiris_file.exists above, but as fallback
            "OSIRIS data file not found. OSIRIS data enrichment will not be possible."
        )
    except json.JSONDecodeError as e:
        std_logger.error(f"Error decoding JSON from OSIRIS data file: {e}")
    except Exception as e:
        std_logger.error(f"Error loading OSIRIS data: {e}")
        std_logger.error(
            "OSIRIS data file unreadable or other error. OSIRIS data enrichment will not be possible.\nPlease run the cli again with the refresh_osiris_data flag set to True to retrieve the required data."
        )
    return None


# Standard logging configuration should be done once at application entry point (e.g., run.py or main CLI script)
# Avoid reconfiguring logging here if it's already set up elsewhere.
# If this module can be run standalone for some reason, then a basicConfig might be suitable here,
# but typically library modules should not configure root logger.

# suppress some annoying warnings when reading excel files - This is standard logging, so it's fine.
logging.getLogger(name="fastexcel.types.dtype").setLevel(level=logging.ERROR)

# initialize settings from (default: read from 'settings.yaml')
SETTINGS = Settings()
# SAMPLESETTINGS = SampleSettings()  # This will also attempt to load 'sample.yaml'

# create global variables from certain settings
# These should be accessed via SETTINGS object preferably, to avoid global state issues.
# However, if they are widely used as globals, ensure they are initialized after SETTINGS.
DEPARTMENT_MAPPING: dict[str, str] = SETTINGS.university_settings.department_mapping
COURSE_MAPPING: dict[str, dict[str, str]] = SETTINGS.university_settings.course_mapping
FINE_AMOUNT: float = SETTINGS.fine_amount
OSIRIS_DATA: dict[str, Any] | None = (
    load_osiris_data()
)  # OSIRIS_DATA value type more general
