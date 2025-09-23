import contextlib
import os
import sys
from collections.abc import Callable
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Literal

import webcolors
import yaml
from loguru import logger
from rich.traceback import install

# sys is already imported above
from easy_access.utils import Directory, File, safe_float

"""Manages application settings, loaded from YAML configuration files."

This module defines dataclasses for structuring settings and provides
functionality to load and parse them from 'settings.yaml' and 'sample.yaml'.
It configures logging and establishes global setting constants for use
throughout the application.
"""


def configure_logger() -> None:
    """Configures the Loguru logger for console and file output."""
    log_dir = Directory(path="logs")
    logger.remove()

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
        level="TRACE",
        enqueue=True,
        backtrace=True,
        diagnose=True,
    )

    try:
        # Re-add file sink inside try to catch path-related errors
        logger.add(
            sink=log_dir.full / "app_{time}.log",
            rotation="1 month",
            format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {message}",
            level="TRACE",
            enqueue=True,
            colorize=False,
        )
    except Exception as e:  # pragma: no cover - defensive
        # If adding the file sink fails, ensure at least console output is available
        logger.warning(f"Failed to add file log sink ({log_dir.full}): {e}")


# DEFINED DIRECTORIES


class DirSetting(Enum):
    """Enum for directory settings keys used in settings.yaml."""

    RAW_COPYRIGHT_DATA = "raw_copyright_data"
    RAW_COPYRIGHT_DATA_FULL = "full_data"
    FACULTIES_DIR = "faculties_dir"
    OVERVIEWS_BACKUP = "overviews_backup"
    SCRIPT_DATA = "script_data"
    PDF_DOWNLOADS = "pdf_downloads"


# SETTINGS FOR EXPORTED SHEETS


class SheetSetting(Enum):
    """Enum for sheet-related settings keys used in settings.yaml."""

    RAW_DATA_COL_ORDER = "raw_data_col_order"
    FINAL_DATA_COL_ORDER = "final_data_col_order"
    COMPLETE_DATA_NAME = "complete_data_name"
    COMPLETE_DATA_COLS = "complete_data_cols"
    DATA_ENTRY_NAME = "data_entry_name"
    DATA_ENTRY_COLS = "data_entry_cols"
    NEW_FIELDS = "new_fields"


@dataclass
class StyleInfo:
    """
    Contains style information for a column in a DataEntrySheet.
    Currently used for conditional formatting.

    For colors, use hex color codes without the leading '#', e.g. 'FF0000' for red;
    alternatively, common color names like 'red', 'blue', 'green' are also supported, they will be converted to hex codes.
    """

    from openpyxl.formatting.rule import Rule

    bg_color: str = ""
    text_color: str = ""
    border_color: str = ""
    bold: bool = False
    activate_on: list[str] = field(
        default_factory=list
    )  # values that trigger the style, if this is a conditional style
    _cf_rule: Rule | None = field(default=None, repr=False)

    def __post_init__(self) -> None:
        """Initializes the conditional formatting rule if conditional formatting rules are specified."""
        self._parse_colors()

        if self.activate_on:
            self._create_conditional_formatting_rule()

    def _parse_colors(self) -> None:
        """
        Parses the input color str to ensure hex format without leading '#'.
        """
        colors = {
            "bg_color": self.bg_color,
            "text_color": self.text_color,
            "border_color": self.border_color,
        }
        for attr, color in colors.items():
            if not color:
                continue
            hex_color = color
            with contextlib.suppress(ValueError):
                hex_color = webcolors.name_to_hex(color).lstrip("#")

            with contextlib.suppress(ValueError):
                hex_color = webcolors.normalize_hex(hex_color)

            try:
                if hex_color.startswith("#"):
                    hex_color = hex_color.lstrip("#")
                hex_color = hex_color.lower()
            except Exception:
                ...
            try:
                setattr(self, attr, hex_color)
            except Exception as e:
                logger.warning(
                    f"Could not set color attribute {attr} to hex {hex_color}: {e}"
                )

    def _create_conditional_formatting_rule(self) -> Rule | None:
        try:
            from openpyxl.formatting.rule import CellIsRule
            from openpyxl.styles import Border, Font, PatternFill, Side

            fill = PatternFill(
                start_color=self.bg_color if self.bg_color else "FFFFFF",
                end_color=self.bg_color if self.bg_color else "FFFFFF",
                fill_type="solid" if self.bg_color else None,
            )

            side = Side(
                style="thin", color=self.border_color if self.border_color else "000000"
            )
            border = (
                Border(left=side, right=side, top=side, bottom=side)
                if self.border_color
                else None
            )
            font = Font(
                color=self.text_color if self.text_color else "000000", bold=self.bold
            )

            self._cf_rule = CellIsRule(
                operator="equal",
                formula=self.activate_on,
                fill=fill,
                border=border,
                font=font,
            )

        except Exception as e:
            logger.warning(
                f"Could not add conditional formatting: {e}. Input colors: {self.bg_color}, {self.text_color}, {self.border_color}"
            )

    def modify_activate_on(self, new_activate_on: list[str]) -> None:
        """Modifies the activate_on list and updates the conditional formatting rule."""
        self.activate_on = new_activate_on
        self._create_conditional_formatting_rule()

    @property
    def cf_rule(self) -> Rule | None:
        """Returns the conditional formatting rule, if any."""
        self._create_conditional_formatting_rule()
        return self._cf_rule


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
    style: StyleInfo = field(default_factory=lambda: StyleInfo())

    def __post_init__(self) -> None:
        if "ENUM" in self.dropdown_options:
            enum_name = self.dropdown_options.split("ENUM:")[-1].strip()
            # try to import that enum from db.models
            try:
                from easy_access.db import enums

                enum_class = getattr(enums, enum_name, None)
                if enum_class and issubclass(enum_class, Enum):
                    self.dropdown_options = ",".join([e.value for e in enum_class])
                else:
                    logger.warning(
                        f"Could not find enum class '{enum_name}' in db.models. Leaving dropdown_options as is."
                    )
            except ImportError as e:
                logger.warning(
                    f"Error importing db.models to load enum '{enum_name}': {e}. Leaving dropdown_options as is."
                )

    @property
    def has_dropdown(self) -> bool:
        """True if dropdown_options are specified, False otherwise."""
        return len(self.dropdown_options) > 0


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


# SCRIPT PIPELINE SETTINGS


@dataclass
class EasyAccessSettings:
    """Configuration settings for the Easy Access Tool."""

    export: bool = False
    only_changes: bool = True
    refresh_osiris_data: bool = False
    only_retrieve_missing_osiris_data: bool = False
    other_sheet: Path | None = None
    enrich_with_osiris_data: bool = True  # whether to enrich with osiris data
    dirs: dict[DirSetting, Directory] = field(default_factory=dict)
    disable_writes: bool = False  # skip file writes
    faculty: str | None = None
    no_file_exists: bool = False  # skip file existence checks
    no_pdf_download: bool = False  # skip pdf downloading
    no_pdf_parse: bool = False  # skip pdf parsing
    # Enable the new workflow-based exporter (inbox/in_progress/done)
    export_workflow: bool = True

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


class EnrichmentSetting(Enum):
    """Enum for enrichment settings keys used in settings.yaml."""

    COURSE_TTL_DAYS = "course_ttl_days"
    PERSON_TTL_DAYS = "person_ttl_days"


@dataclass
class EnrichmentSettings:
    """Holds settings related to OSIRIS data enrichment.

    Attributes:
        course_ttl_days: Time-to-live in days for course data freshness.
                         Courses older than this will be refetched.
        person_ttl_days: Time-to-live in days for person data freshness.
                         Persons older than this will be refetched.
        file_exists_ttl_days: Time-to-live in days for file existence checks.
                             Files older than this will be rechecked.
        file_exists_rate_limit_delay: Delay in seconds between file existence API calls.
                                     Helps avoid rate limiting from Canvas API.
    """

    course_ttl_days: int | None = 90
    person_ttl_days: int | None = 90
    file_exists_ttl_days: int | None = 7
    file_exists_rate_limit_delay: float = 0.01


# UNIVERSITY DATA


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
    canvas_api_token: str | None = None

    def add_api_key(self):
        """
        After initing this class, it will try to find an api token in:
        - environment
        - .env / .secret file(s)
        - api_keys.py module
        in that order of precedence.
        """
        if not self.canvas_api_token:
            self.canvas_api_token = os.getenv("CANVAS_API_TOKEN")
        if not self.canvas_api_token:
            self.canvas_api_token = self._load_env_file()
        if not self.canvas_api_token:
            self.canvas_api_token = self._load_api_keys_module()

    def _load_env_file(self) -> str | None:
        """Loads the API token from a .env or .secret file in
        cwd or parent dir (max 2 levels deep)."""
        env_files = [".env", ".secret"]

        folders_to_search = [os.getcwd()]
        for _i in range(2):  #
            folders_to_search.append(os.path.dirname(folders_to_search[-1]))

        for folder in folders_to_search:
            for env_file in env_files:
                env_file_path = os.path.join(folder, env_file)
                if os.path.exists(env_file_path):
                    with open(env_file_path) as f:
                        for line in f:
                            if line.startswith("CANVAS_API_TOKEN="):
                                return line.split("=", 1)[1].strip()
        return None

    def _load_api_keys_module(self) -> str | None:
        """Tries to loads the API token from the api_keys.py module."""
        with contextlib.suppress(ImportError, ModuleNotFoundError):
            from api_keys import CANVAS_API_TOKEN

            found_token = CANVAS_API_TOKEN
            return found_token
        with contextlib.suppress(ImportError, ModuleNotFoundError):
            from easy_access.api_keys import CANVAS_API_TOKEN

            found_token = CANVAS_API_TOKEN
            return found_token
        return None

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


# MAIN SETTINGS OBJECT


@dataclass
class Settings:
    """Main dataclass holding all application settings, loaded from a YAML file.

    Attributes:
        input_file_path: Path to the settings YAML file (default: "settings.yaml").
        settings_file: File object representing the settings file.
        raw_settings: Raw dictionary loaded from the YAML file.
        dirs: Dictionary mapping directory setting keys (DirSetting) to Directory objects.
        fine_amount: Default fine amount for certain calculations.
        data_settings: Nested DataSettings object.
        university_settings: Nested UniversitySettings object.
        enrichment_settings: Nested EnrichmentSettings object.
        classification_options: List of available classification options.
        dashboard_reload: Boolean indicating if the dashboard should auto-reload.
        db_path: Path to the SQLite database file.
        KEY_TO_PARSER_MAPPING: Internal mapping of setting keys to parser methods.
    """

    input_file_path: str = "settings.yaml"
    settings_file: File = field(init=False)
    raw_settings: dict[str, Any] = field(default_factory=dict, init=False, repr=False)
    dirs: dict[DirSetting, Directory] = field(default_factory=dict, init=False)
    data_settings: DataSettings = field(default_factory=DataSettings, init=False)
    university_settings: UniversitySettings = field(
        default_factory=UniversitySettings, init=False
    )
    enrichment_settings: EnrichmentSettings = field(
        default_factory=EnrichmentSettings, init=False
    )
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
        self.university_settings.add_api_key()

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

                                # StyleInfo parsing
                                # use a `style` key with a dict value to specify StyleInfo attributes
                                # ensure to include a `activate_on` key with a list of values to trigger the style, otherwise the style will not be applied

                                if "style" in col_info_dict and isinstance(
                                    col_info_dict["style"], dict
                                ):
                                    style_dict = col_info_dict.pop("style")
                                    try:
                                        col_info_dict["style"] = StyleInfo(
                                            **{str(k): v for k, v in style_dict.items()}
                                        )
                                    except TypeError as e:
                                        logger.error(
                                            f"Error parsing StyleInfo from {style_dict}: {e}"
                                        )
                                        col_info_dict["style"] = StyleInfo()

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

    def parse_enrichment(self, enrichment_settings_yaml: dict[str, Any]) -> None:
        """Parses the 'enrichment' section of settings.yaml.

        Args:
            enrichment_settings_yaml: The dictionary representing the 'enrichment' settings.
        """
        for key_str, value_data in enrichment_settings_yaml.items():
            try:
                key_enum = EnrichmentSetting(value=key_str)
            except ValueError:
                logger.error(
                    f"Unrecognized enrichment setting {key_str} (with value: {value_data}). Skipping."
                )
                continue
            match key_enum:
                case EnrichmentSetting.COURSE_TTL_DAYS:
                    try:
                        self.enrichment_settings.course_ttl_days = int(value_data)
                    except (ValueError, TypeError):
                        logger.warning(
                            f"Invalid value for COURSE_TTL_DAYS: {value_data}. Using default None"
                        )
                case EnrichmentSetting.PERSON_TTL_DAYS:
                    try:
                        self.enrichment_settings.person_ttl_days = int(value_data)
                    except (ValueError, TypeError):
                        logger.warning(
                            f"Invalid value for PERSON_TTL_DAYS: {value_data}. Using default None"
                        )
                case _:
                    logger.warning(
                        f"Unrecognized enrichment setting {key_str} (with value: {value_data}). Skipping."
                    )

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
            "enrichment": self.parse_enrichment,
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

        for key, value in self.raw_settings.items():
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


@dataclass
class OverrideSettings(Settings):
    """Dataclass to override settings, inherits from the main Settings class.

    Attributes:
        input_file_path: Path to the override settings file (YAML)
        The settings file only required the settings to override.
        Not all settings can be overridden.
        Main focus is on the data_settings section, to allow for quick changes to column mappings, order, new fields, and sheet names.
    """

    override_input_file_path: Path | str | None = (
        None  # No default, must be provided for overrides
    )
    backup_data_settings: DataSettings | None = field(default=None, init=False)
    override_settings: dict[str, Any] = field(
        default_factory=dict, init=False, repr=False
    )

    def __post_init__(self) -> None:
        """Initializes override settings after dataclass creation."""
        super().__post_init__()  # Call parent post-init to load and parse settings
        if self.override_input_file_path:
            self.override_default_settings()

    def override_default_settings(self) -> None:
        """
        Figures out if the override settings file is a .yaml or .xlsx file and calls the appropriate method to override settings.
        """

        if not self.override_input_file_path:
            logger.info(
                "No override settings file path provided. Using default settings."
            )
            return

        file_path = (
            Path(self.override_input_file_path)
            if not isinstance(self.override_input_file_path, Path)
            else self.override_input_file_path
        )
        if not file_path.exists():
            logger.error(f"Override settings file does not exist: {file_path}")
            return

        if file_path.suffix.lower() in [".yaml", ".yml"]:
            self._override_from_yaml(file_path)
        else:
            try:
                self._override_from_yaml(file_path)
            except Exception as e:
                logger.error(
                    f"Error overriding settings from file {file_path}: {e}. Using default settings."
                )
                return

        if self.override_settings:
            self.parse_override()

    def _override_from_yaml(self, file_path: Path) -> None:
        try:
            with open(file=file_path, encoding="utf-8") as f:
                loaded_yaml = yaml.load(stream=f, Loader=yaml.FullLoader)
                if isinstance(loaded_yaml, dict):
                    self.override_settings = loaded_yaml
        except Exception as e:
            logger.error(f"Error loading override settings from {file_path}: {e}")
            return

    def parse_override(self) -> None:
        """
        Based on the loaded override settings in self.override_settings, override the relevant settings in self.
        Currently only supports overriding specific fields in data_settings.
        """
        required_cols = {
            "material_id": False,
            "url": False,
            "workflow_status": False,
            "manual_classification": False,
            "v2_manual_classification": False,
            "v2_lengte": False,
            "v2_overnamestatus": False,
        }
        self.backup_data_settings = self.data_settings
        if not self.override_settings:
            logger.info("No override settings to apply.")
            return

        data_settings_override = self.override_settings.get("data_settings")
        if not isinstance(data_settings_override, dict):
            logger.error("No valid 'data_settings' section in override settings.")
            return
        new_col_settings = data_settings_override.get("data_entry_cols")

        # this should be a list of dicts
        # containing the cols the include in the data entry sheet
        # IN order of appearance
        # grab the initialized data entry cols from self.data_settings to grab existing settings
        # the iterate over the new list, creating a new self.data_settings object that copies over the existing settings but overrides order/name/inclusion based on the override
        if not isinstance(new_col_settings, list):
            logger.error(
                "'data_entry_cols' in override settings is not a list. Cannot override."
            )
            return
        existing_cols = {col.name: col for col in self.data_settings.data_entry_cols}
        overridden_cols: list[ColInfo] = []
        for col_dict in new_col_settings:
            if not isinstance(col_dict, dict):
                logger.warning(
                    f"Expected dict or str for column override, got {type(col_dict)}. Skipping."
                )
                continue
            col_name = col_dict.get("name")

            if not isinstance(col_name, str):
                logger.warning(
                    f"Column override missing 'name' or 'name' is not a string: {col_dict}. Skipping."
                )
                continue

            if col_name in existing_cols:
                if col_name in required_cols:
                    required_cols[col_name] = True

                try:
                    overridden_col = existing_cols[col_name]
                    if "style" in col_dict and isinstance(col_dict["style"], dict):
                        overridden_col.style = StyleInfo(**col_dict["style"])
                    overridden_cols.append(overridden_col)
                except TypeError as e:
                    logger.error(
                        f"Error creating ColInfo for overridden column '{col_name}': {e}. Skipping."
                    )
                if col_name == "v2_lengte":
                    logger.debug(f'Input override for "v2_lengte": {col_dict}')
                    logger.debug(f'Parsed colinfo for "v2_lengte": {overridden_col}')
            else:
                logger.warning(
                    f"Column '{col_name}' in override settings not found in existing data entry columns. Skipping."
                )

        if overridden_cols:
            for req_col, found in required_cols.items():
                if not found:
                    overridden_cols.append(existing_cols[req_col])

            self.data_settings.data_entry_cols = overridden_cols
            logger.info(
                f"Overridden data entry columns with {len(overridden_cols)} columns from override settings."
            )


# GLOBALS
# (deprecate? dev only?)

configure_logger()
install(show_locals=True)

SETTINGS: Settings = Settings()  # deprecate this?
