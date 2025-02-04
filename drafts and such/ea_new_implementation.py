
from dataclasses import dataclass, field
from datetime import datetime, date
import pathlib
import os
import shutil
from enum import Enum, StrEnum
import dotenv
from rich.console import Console
from typing import Literal, Type, Annotated, Any
import polars as pl

"""
This script is a new approach to the easy access project, see easy_access_cli.py for the current implementation.

In short, the script should do the following:

- read in raw data from the copyright tool, exported as a excel file once a week, found in the COPYRIGHT_EXPORT_DIR
- process this raw data into a standard format
- match the data with the exising data in the faculty sheets, found in the FACULTIES_DIR (recursively read all files in this dir except those with 'overview' in the name)
    - use the material_id to identify each item (unique identifier)
- enrich it with additional data (e.g. course names, data from osiris, library systems, personell data, Pure, OpenAlex, etc)
- store all of this data into a sqlite db, with separate tables for various entities (e.g. faculties, employees, courses, items, etc)
- produce overview sheets
- produce reports
- split data into multiple files (e.g. per faculty, per period, per filetype, etc), including data entry sheets

"""
dotenv.load_dotenv("settings.env")
cons = Console(emoji=True, markup=True)
print: callable = cons.print

class Directory:
    """
    Class representing a directories with functions for common manipulations & searches.
    Init with an absolute or relative path (relative to the cwd).
    If the given dir does not yet exist, it will be created. Disable this by setting the 'create_dir' parameter to False.

    Two additional parameters can be set:
    ignorestr (str): If set, files with this string in their name will be ignored.
    file_type (list[str]|str): If set, only files with this extension will be returned. Use extensions including the . (e.g. '.csv').
    """
    full: pathlib.Path
    def __init__(self, path: str, create_dir: bool = True, ignorestr: str = 'overview', file_type:list[str]|str = '.'):

        self.input_path_str = path
        self.create_dir = create_dir

        self.ignorestr = ignorestr
        self.file_type = file_type

        # check if the path is absolute
        if pathlib.Path(path).is_absolute():
            self.full = pathlib.Path(path)
        else:
            self.full = pathlib.Path.cwd() / path

        self.post_init()

    def post_init(self) -> None:
        """
        Checks to see if this is actually a dir,
        or create it if create_dir is set to True.
        """
        if not self.full.exists():
            if self.create_dir:
                self.create()
            else:
                raise FileNotFoundError(
                    f"Directory {self.full} does not exist and create_dir is set to False."
                )
        if not self.full.is_dir():
            raise NotADirectoryError(f"Directory {self.full} is not a directory.")

    @property
    def files(self) -> list["File"]:
        """
        Gets all files in the dir as a list of File objects.
        """
        return [
            File(self.full / file) for file in self.full.iterdir() if all([file.is_file(),self.ignorestr not in file.name, file.suffix in self.file_type])
        ]

    @property
    def files_r(self) -> list["File"]:
        """
        Recursively gets all files in the dir (so including files in subdirs) as a list of File objects.
        """
        return [
            File(self.full / file) for file in self.full.rglob("*") if all([file.is_file(),self.ignorestr not in file.name, file.suffix in self.file_type])
        ]

    @property
    def dirs(self) -> list["Directory"]:
        """
        Returns a list of all dirs in this Directory as a list of Directory objects.
        """
        return [Directory(str(d), False) for d in self.full.iterdir() if d.is_dir()]

    @property
    def dirs_r(self) -> list["Directory"]:
        """
        Recursively gets all dirs in the dir (so including dirs in subdirs) as a list of Directory objects.
        """
        return [Directory(str(d), False) for d in self.full.rglob("*") if d.is_dir()]

    @property
    def newest_file(self) -> "File":
        """
        Returns the newest file in the dir as a File object.
        """
        return max(self.files, key=lambda x: x.created) if self.files else None


    @property
    def newest_file_r(self) -> str:
        """
        Recursively gets the newest file in the dir, so including files in subdirs, as a File object.
        """
        return max(self.files_r, key=lambda x: x.created) if self.files_r else None

    @property
    def exists(self) -> bool:
        return self.full.exists()

    @property
    def is_dir(self) -> bool:
        return self.full.is_dir()

    def create(self) -> None:
        try:
            self.full.mkdir(parents=True, exist_ok=False)
        except FileExistsError:
            pass

    def __eq__(self, other) -> bool:
        return self.full == other.full

    def __str__(self):
        return str(self.full)

    def __repr__(self):
        return f"DirPath('{self.input_path_str}') -> {self.full}"

class File:
    """
    Simple class for files + operations
    Parameters:
        path: str or Path
            absolute or relative from the cwd.
            Should always end with filename including extension.
    """

    def __init__(self, path: str | pathlib.Path):
        self._path_init_str = str(path)

        assert isinstance(path, str) or isinstance(path, pathlib.Path)

        if isinstance(path, pathlib.Path):
            self._path = path
            self._name = path.name
            self._extension = path.suffix
            self._dir = Directory(str(self._path.absolute().parent))
        elif isinstance(path, str):
            if "/" in path:
                self._name = path.rsplit("/", 1)[-1]
                self._dir = Directory(path.rsplit("/", 1)[0], create_dir=True)
            else:
                self._name = path
                self._dir = Directory(os.getcwd())

            self._extension = self._name.split(".")[-1]
            self._path = self._dir.full / self._name

    @property
    def exists(self) -> bool:
        return self._path.exists()

    @property
    def is_file(self) -> bool:
        return self._path.is_file()

    @property
    def path(self) -> pathlib.Path:
        return self._path

    @property
    def name(self) -> str:
        return self._name

    @property
    def extension(self) -> str:
        return self._extension

    @property
    def dir(self) -> Directory:
        return self._dir

    @property
    def created(self) -> datetime:
        return datetime.fromtimestamp(self._path.stat().st_birthtime)

    @property
    def modified(self) -> datetime:
        return datetime.fromtimestamp(self._path.stat().st_mtime)

    def copy(self, new_path: str) -> "File":
        shutil.copy(self._path, new_path)
        return File(new_path)

    def move(self, new_path: str) -> "File":
        shutil.move(self._path, new_path)
        return File(new_path)

    def rename(self, new_name: str) -> "File":
        self._path = self._dir.full / new_name
        return File(self._path)

    def delete(self) -> None:
        os.remove(self._path)

    def __eq__(self, other) -> bool:
        return self._path == other.path

    def __str__(self):
        return str(self._path)

    def __repr__(self):
        if self._path_init_str != str(self._path):
            return f"FilePath('{self._path_init_str}') -> {self._path}"
        else:
            return f"FilePath('{self._path}')"


#---------------------
# Enums
#---------------------

def create_period_type(start_year: int, end_year: int) -> Type:
    """
    Creates a type alias for valid period strings within the specified range.
    """
    periods = ["1A", "1B", "2A", "2B", "3", "SEM1", "SEM2", "JAAR"]
    valid_periods = [f"{year}-{period}" for year in range(start_year, end_year + 1) for period in periods]
    return Literal[*valid_periods]

PeriodType = create_period_type(2000, 2025)

def create_period_enum(start_year: int, end_year: int) -> type[Enum]:
    """
    Programmatically creates a Period Enum with all valid periods between start_year and end_year.
    Format: YYYY-XXX
    where YYYY is the academic year (e.g. 2021) and XXX is one of the periods listed here:
    1A, 1B, 2A, 2B, 3, SEM1, SEM2, JAAR.
    Ex: 2020-1A, 2020-SEM1, 2021-JAAR
    """
    periods = ["1A", "1B", "2A", "2B", "3", "SEM1", "SEM2", "JAAR"]
    enum_members = {}
    for year in range(start_year, end_year + 1):
        for period in periods:
            enum_key = f"Y{year}_{period}"
            enum_value = f"{year}-{period}"
            enum_members[enum_key] = enum_value

    return Enum("Period", enum_members)

# Create the actual Period enum
Period = create_period_enum(2000, 2025)

class Faculty(StrEnum):
    """
    All valid faculty names.
    """

    TNW = "TNW"
    EEMCS = "EEMCS"
    ET = "ET"
    ITC = "ITC"
    BMS = "BMS"
    UNKNOWN = "UNKNOWN"

class ProgrammeType(Enum):
    BACHELOR = 'B'
    MASTER = 'M'
    FACULTY = 'Faculty'
    SUPPORT = 'Support'
    OTHER = 'Other'
    UNKNOWN = 'Unknown'


class CopyRightClassification(StrEnum):
    LANGE_OVERNAME = "lange overname"
    MIDDELLANGE_OVERNAME = "middellange overname"
    KORTE_OVERNAME = "korte overname"

    EIGEN_WERK_POWERPOINT = "eigen materiaal - powerpoint"
    EIGEN_WERK_OVERIGE = "eigen materiaal - overig"
    EIGEN_MATERIAAL_TITELINDICATIE = "eigen materiaal - titelindicatie"

    OPEN_ACCESS = "open access"

    IN_ONDERZOEK = "in onderzoek"
    VERWIJDERVERZOEK_VERSTUURD = "verwijderverzoek verstuurd"
    LICENTIE_BESCHIKBAAR = "licentie beschikbaar"

    ONBEKEND = "onbekend"

class CopyRightScope(StrEnum):
    ALTIJD = "altijd"
    JAAR = "jaar"
    MODULE = "module"

class CanvasStatus(StrEnum):
    UNPUBLISHED = "Unpublished"
    DELETED = "Deleted"
    PUBLISHED = "Published"

# ---------------------
# dataclasses
# --------------------

@dataclass
class Programme:

    name: str
    faculty: Faculty = Faculty.UNKNOWN
    abbreviation: str = ""
    programme_type: ProgrammeType = ProgrammeType.UNKNOWN

    @staticmethod
    def from_string(input_str:str) -> "Programme":
        """
        Creates a Programme instance from a string taken from the raw copyright export, column Department.
        """
        if not input_str.strip():
            print("Empty input string.")
            return Programme(name="", faculty=Faculty.UNKNOWN, abbreviation="", programme_type=ProgrammeType.UNKNOWN)

        faculty = Faculty.UNKNOWN
        programme_type = ProgrammeType.UNKNOWN

        # No colon case
        if ": " not in input_str:
            words = input_str.split()
            if words and "Master" in words[0]:
                programme_type = ProgrammeType.MASTER
            else:
                print(f"Could not determine programme type for {input_str}.")
            return Programme(name=input_str, faculty=faculty, abbreviation="", programme_type=programme_type)


        # Else split on colon
        abbrev, name = input_str.split(": ", 1)
        # No dash case
        if "-" not in abbrev:
            if abbrev in Faculty.__members__:
                faculty = Faculty[abbrev]
                programme_type = ProgrammeType.FACULTY
            else:
                print(f'Couldnt determine programme type from: {abbrev} for {input_str}')
            return Programme(name=name, faculty=faculty, abbreviation=abbrev, programme_type=programme_type)

        # Else split on dash
        programme_str, abbrev = abbrev.split("-",1)
        try:
            programme_type = ProgrammeType(programme_str)
        except ValueError as e:
            programme_type = ProgrammeType.OTHER
            print(f'Got other programme type: {programme_str} for {input_str}')

        return Programme(name=name, faculty=faculty, abbreviation=abbrev, programme_type=programme_type)

@dataclass
class Employee:
    """
    Class representing an employee.
    """
    name: str
    email: str = ""
    faculties: list[Faculty|None] = field(default_factory=list)
    programmes: list[Programme|None] = field(default_factory=list)

@dataclass
class OsirisCourse:
    course_code: int
    periods: list[Annotated[str, PeriodType]]
    programmes: list[Programme]
    name: str
    contacts: list[Employee] = field(default_factory=list)
    teachers: list[Employee] = field(default_factory=list)
    examiners: list[Employee] = field(default_factory=list)

@dataclass
class CanvasCourse:
    """
    A Canvas course code, e.g. 2024-CALB-1B
    """
    raw_code: str
    raw_name: str
    period: Annotated[str, PeriodType]
    canvas_code: str
    osiris_courses: list[OsirisCourse] = field(default_factory=list)
    name: str
    url: str

#---------------------
# main item dataclasses
# for raw/processed items, sheets, etc
#---------------------
@dataclass
class RawItem:
    """
    Hold the fields for a single row of raw data, imported from the CSV file that was exported from the Copyright Tool.
    """

    # the columns in the excel file, with their types. If None is present in the typelist, the field is optional.
    # If None is the only type, disregard the field.
    INPUT_COLS: dict[str, list[Any | None]] = {
        "Material id":[int],
        "Period":[Annotated[str, PeriodType]],
        "Department":[Programme],
        "Course code":[str],
        "Course name":[str],
        "url":[str],
        "Filename":[str],
        "Title":[str|None],
        "Owner":[Employee|None],
        "Filetype":[str|None],
        "Classification":[CopyRightClassification],
        "Type":[None],
        "ML Prediction":[CopyRightClassification],
        "Manual classification":[CopyRightClassification|None],
        "Manual identifier":[None],
        "Scope":[CopyRightScope|None],
        "Remarks":[str|None],
        "Auditor":[Employee|None],
        "Last change":[date],
        "Status":[CanvasStatus],
        "Google search file":[None],
        "ISBN":[int|None],
        "DOI":[str|None],
        "In collection":[bool|None],
        "pagecount":[str],
        "wordcount":[str],
        "picturecount":[str],
        "Author":[str],
        "Publisher":[str],
        "Reliability":[str],
        "Pages * Students":[int],
        "#students_registered":[int],
    }

    material_id: int
    title: str
    filetype: str
    course_code: int
    period: Annotated[str, PeriodType]
    ...#more fields here





@dataclass
class Item:
    """
    Data for a single item. Starts from a RawItem, is then processed and enriched to form the items used in the faculty sheets.
    """

    raw: RawItem
    faculty: Faculty
    ...


@dataclass
class RawData:
    """
    Holds raw data from a single Excel file exported from the Copyright Tool.
    """

    rows: list[RawItem] = field(default_factory=list, init=False, repr=False)
    sheetname: str = field(default='Sheet1')
    import_date: date = field(default_factory=date.today()) # when the data was imported
    file: File # the source file.
    df_raw_data: pl.DataFrame = field(init=False, repr=False)

    def __post_init__(self) -> None:
        """
        Initial processing of the raw data.
        Read in the excelsheet and convert it to a list of RawItem objects.
        """
        self.df_raw_data = pl.read_excel(self.file.path)
        self.rows = [RawItem(**row) for row in self.df_raw_data.to_dicts()]


    def __process_raw_data__(self) -> None:
        """
        Read  the raw data and process it into a list of RawItem objects.
        Store as self.rows.
        """

class ProcessedData:
    ...
# ---------------------
# reading / ingesting / processing / enriching data
# ---------------------



# functionaly that is needed here:
# - read in excel with raw data from copyright export, turn into RawData object
# - read in excel files with processed data (e.g. faculty sheets, overview sheets), turn into FacultySheet object
# - process the raw data into the format needed for the faculty sheets
# - enrich the raw data with additional data (e.g. course names, data from osiris, library systems, personell data, Pure, OpenAlex, etc)
# - update existing sheets with new data
# - create overview sheets
# - produce reports
# - split data into multiple files (e.g. per faculty, per period, per filetype, etc)
# - produce dashboard


class DataReader:
    data_entry_dir: Directory
    all_entries_dir: Directory
    copyright_dir: Directory

    raw_copyright_data: RawData
    all_entries_data: ProcessedData
    data_entry_datalist: list[ProcessedData]

    def __init__(self, copyright_dir: Directory | str | pathlib.Path | None = None, data_entry_dir: Directory | str | pathlib.Path | None = None, all_entries_dir: Directory | str | pathlib.Path | None = None):

        self.copyright_dir =  Directory(os.getenv("COPYRIGHT_EXPORT_DIR")) if not copyright_dir else (Directory(copyright_dir) if not isinstance(copyright_dir, Directory) else copyright_dir)
        self.data_entry_dir = Directory(os.getenv("FACULTIES_DIR")) if not data_entry_dir else (Directory(data_entry_dir) if not isinstance(data_entry_dir, Directory) else data_entry_dir)
        self.all_entries_dir = Directory(os.getenv("ALL_ENTRIES_DIR")) if not all_entries_dir else (Directory(all_entries_dir) if not isinstance(all_entries_dir, Directory) else all_entries_dir)

    def read_copyright_export(self) -> None:
        """
        Reads in data from the latest copyright export file in the copyright dir.
        """


    def read_all_entries_data(self) -> None:
        ...
