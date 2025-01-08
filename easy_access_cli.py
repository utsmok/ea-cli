"""
Easy Access Sheet Toolkit
September 2024
Samuel Mok / s.mok@utwente.nl / cip@utwente.nl
homepage: https://github.com/utsmok/easyaccesscli/

Note: only tested on windows systems

see readme.md for more info

Q U I C K    S T A R T
    Run with standard settings:
        > uv run easy_access_cli.py
    View cli instructions:
        > uv run easy_access_cli.py --help

    if you don't have uv installed yet:

I N S T A L L   U V
    UV is an all-in-one python manager.
    Install by opening Powershell (press windows key, type 'powershell', enter) and pasting the following lines:

        > powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

    and press enter to install. For more info, see the uv docs: https://docs.astral.sh/uv/getting-started/installation/
    Once uv is installed, close PowerShell, start it again, and type

        > uv python install

    and the setup is all done! Now you can run the cli help with:

        > uv run easy_access_cli.py --help

This python file contains the following:
    - class EasyAccessToolkit with core functionality to ingest & process data from SURF's copyRight tool for easy access, and export various sheets for end-users
    - typer function cli provides a command line interface
    - helper classes File and Directory for handling... files and directories.
"""
import requests
import json
import os
import pathlib
import shutil
from datetime import datetime, timedelta
from enum import Enum
import locale
import dotenv
import openpyxl
import openpyxl.worksheet
import openpyxl.worksheet.datavalidation
import openpyxl.worksheet.table
import openpyxl.worksheet.worksheet
import polars as pl
import typer
from openpyxl.worksheet.table import Table as ExcelTable
from openpyxl.worksheet.table import TableStyleInfo
from rich.console import Console
from typing_extensions import Annotated
from rich.table import Table
from rich.terminal_theme import SVG_EXPORT_THEME
import copy

from pathlib import WindowsPath



# load settings.env to local environment
dotenv.load_dotenv("settings.env")

# rich Console + overload the print function
cons = Console(emoji=True, markup=True)
print: callable = cons.print

# shorthands for printing with different colors
def info(text: str):
    """
    Prints an information message.
    """
    print(f"[cyan]:information: |> [/cyan] {text}")

def warn(text: str):
    """
    Prints a warning message.
    """
    print(f"[bold red]:warning: |>  {text}[/bold red]")

def cool(text: str):
    """
    Prints a nice message.
    """
    print(f"[yellow] :smiling_face_with_sunglasses: |>  {text} [/yellow]")

# ----------------------------------------------------------------------------------------------------------------------
# Classes for handling files and directories
# ----------------------------------------------------------------------------------------------------------------------
class Directory:
    """
    Simple class for directories + operations.
    Init with an absolute path, or a path relative to the current working directory.
    If the dir does not yet exist, it will be created. Disable this by setting the 'create_dir' parameter to False.
    """
    full: pathlib.Path
    def __init__(self, path: str, create_dir: bool = True):

        self.input_path_str = path
        self.create_dir = create_dir

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
            File(self.full / file) for file in self.full.iterdir() if file.is_file()
        ]

    @property
    def files_r(self) -> list["File"]:
        """
        Recursively gets all files in the dir, so including files in subdirs, as a list of File objects.
        """
        return [
            File(self.full / file) for file in self.full.rglob("*") if file.is_file()
        ]

    @property
    def dirs(self, r: bool = False) -> list["Directory"]:
        """
        Returns a list of all dirs in this Directory as a list of Directory objects.
        If r is set to True, it will return all children dirs recursively.
        """
        if not r:
            return [Directory(str(d), False) for d in self.full.iterdir() if d.is_dir()]
        if r:
            return [Directory(str(d), False) for d in self.full.rglob("*") if d.is_dir()]

    def newest_file(self, file_type:list[str]|str|None = None) -> "File":
        """
        Returns the newest file in the dir as a File object.
        Parameters:
            file_type (str): If set, only files with this extension will be returned.
            input the extension incl dot; or a list of them.
        """
        all_files = self.files
        if file_type:
            if isinstance(file_type, str):
                file_type = [file_type]
            all_files = [file for file in all_files if file.extension in file_type if 'overview' not in file.name]
        if not all_files:
            return None
        return max(all_files, key=lambda x: x.created)


    @property
    def newest_file_r(self) -> str:
        """
        Recursively gets the newest file in the dir, so including files in subdirs, as a File object.
        """
        all_files = self.files_r
        return max(all_files, key=lambda x: x.created)

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
            relative from the current working directory.
            OR
            absolute path to the file.
            Should always end with the filename including extension.
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

class Functions(str, Enum):
    """
    CLI option for picking which functions to run, see cli()
    """

    both = "both"
    read = "read"
    export = "export"
    test = "test"

# ----------------------------------------------------------------------------------------------------------------------
# Main functions
# ----------------------------------------------------------------------------------------------------------------------
def cli(
    do: Annotated[
        Functions,
        typer.Option(
            case_sensitive=False,
            help="Which tool to run: read in new data, export current data, or both.",
            rich_help_panel="Functions",
        ),
    ] = Functions.read,
    changes: Annotated[
        bool,
        typer.Option(
            help="Only add items that have been changed to new faculty sheets.",
            rich_help_panel="Functions",
        ),
    ] = True,
    remove_previous: Annotated[
        bool,
        typer.Option(
            help="First remove the latest excel sheet for each faculty/programme, then run the rest of the tool.",
            rich_help_panel="Functions",
        ),
    ] = False,
    save: Annotated[
        bool,
        typer.Option(
            help="If enabled, will store results in excel files. If disabled will only print to console.",
            rich_help_panel="Functions",
        ),
    ] = True,
    other_sheet: Annotated[
        str | None,
        typer.Option(
            help="(relative) path to a xlsx sheet to read in instead of CopyRight Data.",
            rich_help_panel="Read in data from alternate source",
        ),
    ] = None,
):
    """
    Runs the Easy Access toolkit with the specified settings.\n
    Make sure that these two files are present in the current dir and contain the required info:\n\n
        'settings.env': The directories to use\n
        'department_mapping.json': The mapping between department names and faculty names\n
    \n
    Visit the repo for more instructions & the latest version: https://github.com/utsmok/ea-cli. (<- you can click this in your terminal!)\n
    \n\n
    Example usage\n
    --------------\n
    ea-cli\n
    ea-cli --do export\n
    ea-cli --no-changes\n
    ea-cli --do read --changes\n
    """
    def delete_latest_file(subdir:Directory) -> None:
        # only if it's been created in the last 3 days
        newest_file: File = subdir.newest_file(['.xlsx', '.xls'])
        if not newest_file:
            return None
        if newest_file.created > datetime.now() - timedelta(days=3):
            print(f'Current latest file in subdir:\n {newest_file.name}')
            conf = input('Remove this file? [y/N] ')
            if conf.lower() == 'y':
                newest_file.delete()
                print(f'[red]Deleted[/red] {newest_file.name}.')
            else:
                print(f'[cyan]Keeping[/cyan] {newest_file.name} and moving on.\n\n')
            return conf
    if remove_previous:

        sheet_dir = Directory(os.getenv("FACULTIES_DIR"))
        all_items_dir = Directory(os.getenv("ALL_ITEMS_DIR"))
        dirlist = sheet_dir.dirs
        dirlist.append(all_items_dir)

        for subdir in dirlist:
            print(f'[green]subdir {subdir}[/green]\n------------------')
            conf = delete_latest_file(subdir)
            if not conf:
                continue
            while conf.lower() == 'y':
                conf = delete_latest_file(subdir)
                if not conf:
                    break
            if subdir.dirs:
                print(f'[magenta]sub-subdir {subdir}[/magenta]\n')
                for subsubdir in subdir.dirs:
                    conf = delete_latest_file(subsubdir)
                    if not conf:
                        continue
                    while conf.lower() == 'y':
                        conf = delete_latest_file(subsubdir)
                        if not conf:
                            break


    if do not in [Functions.both, Functions.read, Functions.export, Functions.test]:
        warn(
            "No functions selected! Aborting the script. Next time, enable at least one of 'Function' options; for details run ea-cli --help."
        )
        cool("Thank you for using the Easy Access tool!")
        raise typer.Exit(code=1)


    dirs = {
        "copyright_export": None,
        "copyright_import": None,
        "faculties": None,
        "all_items": None,
    }

    tool = EasyAccessTool(
        functions=do, only_changes=changes, dirs=dirs, other_sheet=other_sheet, save_files=save
    )
    tool.run()

    cool("All done! Thank you for using the Easy Access tool!")

class EasyAccessTool:


    """
    This class contains all the actual functionality of the script.
    For an overview see the comments & docstrings per function, plus readme.md.
    """

    # which functions to run when self.run() is called
    settings: list[callable] = []

    # keep track of relevant files and directories
    files: dict[str, File]
    dirs: dict[str, Directory] = {
        "root": Directory(os.getcwd()),
        "copyright_export": Directory(os.getenv("COPYRIGHT_EXPORT_DIR")),
        "copyright_import": Directory(os.getenv("COPYRIGHT_IMPORT_DIR")),
        "all_items": Directory(os.getenv("ALL_ITEMS_DIR")),
        "faculties": Directory(os.getenv("FACULTIES_DIR")),
    }


    # initialize the various dataframes used to get data from / write to .xlsx files
    raw_copyright_data: pl.DataFrame = pl.DataFrame()  # data directly from copyright tool
    copyright_data: pl.DataFrame = pl.DataFrame()  # data with normalized column names & some cleanup
    faculty_sheet_data: pl.DataFrame = pl.DataFrame()  # data from the faculty sheets
    all_items_sheet_data: pl.DataFrame = (
        pl.DataFrame()
    )  # data from the 'all_items' sheet

    # this mapping is used to get the corresponding faculty from the copyright data column 'departments'
    # it should be present in the file 'department_mapping.json' in the same directory as easy_access.cli.py
    # a department_mapping.json file for the University of Twente is included in the repo
    dept_mapping_path = File("department_mapping.json")
    DEPARTMENT_MAPPING = json.load(open(dept_mapping_path.path, encoding="utf-8"))

    # this mapping is used to map courses to certain groupings
    # to be used in combination with the faculty / department
    course_mapping_path = File("course_mapping.json")
    COURSE_MAPPING = json.load(open(course_mapping_path.path, encoding="utf-8"))

    # list of all found/used faculties
    faculties: list[str]

    # latest copyRight file & when it was created
    latest_file: File
    latest_file_date: str

    # starting excel style number for the data entry tables
    # simple hack to ensure repeatable styles
    style_iter: int = 2

    # the amount to multiply pages_x_students by to calculate the fine
    # this is roughly the average fine per student per page as defined by UvO
    fine_amount: float = 0.3

    # path to another sheet to read in instead of CopyRight data
    other_sheet: File | None = None

    # flag to indicate if there are no new items to add
    no_new_items: bool = False

    # standard column order for the complete data sheets
    column_order = [
        "material_id",
        "period",
        "department",
        "course_code",
        "course_name",
        "url",
        "filename",
        "title",
        "owner",
        "filetype",
        "classification",
        "type",
        "ml_prediction",
        "manual_classification",
        "manual_identifier",
        "scope",
        "remarks",
        "auditor",
        "last_change",
        "status",
        "google_search_file",
        "isbn",
        "doi",
        "in_collection",
        "pagecount",
        "wordcount",
        "picturecount",
        "author",
        "publisher",
        "reliability",
        "pages_x_students",
        "count_students_registered",
        "retrieved_from_copyright_on",
        "workflow_status",
        "faculty",
    ]

    # debug option: completely disables all new file writes
    disable_writes = False

    def __init__(
        self,
        functions: Functions | None = Functions.both,
        dirs: dict[str, str] | None = None,
        only_changes: bool = True,
        other_sheet: str | None = None,
        save_files: bool = True,
    ) -> None:
        """
        Parameters:
            setting:  str | None
                Pick which functions to run when self.run() is called. If no argument is passed, it will run all functions.
                pick from one of the presets below:
                    'none' -> don't run any functions
                    'all' -> run all functions
                    'new_data' -> read in new data, process it, create new faculty and all itemssheets
                    'read_sheets' -> read in faculty sheet data, process, create import sheet
            dirs: dict[str,str] | None
                A dict containing the str path to the directories to use. If not provided, it will use the default dirs from settings.env.
            only_changes: bool = True
                True (default): only add items that have been changed to the created sheets
                False: add all items from the CopyRight export to the created sheets
            other_sheets: list[str] | None
                A list of paths to additional .xlsx files to ingest instead the raw data from CopyRight.
        """


        self.disable_writes = not save_files
        self.enrich_with_osiris_data = True # set to False to disable enriching the data with OSIRIS data

        # determine which functions to run
        # first check if we need to read in copyRight data (default), or other sheets

        if other_sheet:
            self.other_sheet = File(other_sheet)

        # export only new items, or all items found in the copyright export?
        self.only_changes = only_changes

        # set the functions to run
        if functions is None:
            self.settings = []
        elif functions == Functions.both:
            if not self.other_sheet:
                self.settings = [
                    self.read_copyright_export,
                    self.process_copyright_export,  # read in new data
                    self.read_all_items_sheets,
                    self.read_faculty_sheets,  # read in data manually added to sheets
                    self.create_import_sheet,  # from the old data, create a sheet to import into CopyRight
                    self.create_faculty_sheets,
                    self.create_all_items_sheet, # create new sheets with new data
                    self.create_faculty_overview
                ]
            else:
                self.settings = [
                    self.read_other_sheet,  # read in new data
                    self.read_all_items_sheets,
                    self.read_faculty_sheets,  # read in data manually added to sheets
                    self.create_import_sheet,  # from the old data, create a sheet to import into CopyRight
                    self.create_faculty_sheets,
                    self.create_all_items_sheet, # create new sheets with new data
                    self.create_faculty_overview
                ]

        elif functions == Functions.read:
            if not self.other_sheet:
                self.settings = [
                    self.read_copyright_export,
                    self.process_copyright_export,  # read in new data
                    self.create_faculty_sheets,
                    self.create_all_items_sheet, # create new sheets with new data
                    self.create_faculty_overview

                ]
            else:
                self.settings = [
                    self.read_other_sheet,
                    self.process_copyright_export,  # read in new data
                    self.create_faculty_sheets,
                    self.create_all_items_sheet,  # create new sheets with new data
                    self.create_faculty_overview

                ]
        elif functions == Functions.export:
            self.settings = [
                self.read_faculty_sheets,  # read in data manually added to sheets
                self.create_import_sheet,  # from the current manually added data, create a sheet to import into CopyRight
            ]
            if self.other_sheet:
                warn(
                    f"Note: Only exporting data, so the contents of other sheet {self.other_sheet} will have no effect on the output."
                )

        # if dirs is set, add them to the self.dirs dict
        if dirs:
            for key, value in dirs.items():
                if value:
                    self.dirs[key] = Directory(value)

    def run(self) -> None:
        """
        Runs the functions as specified in the settings dict.
        """
        for func in self.settings:
            func()

    def read_other_sheet(self) -> None:
        """
        Reads in the data from another sheet as the datasource, instead of using CopyRight data.
        Sheet should be formatted in the same way as the faculty output sheets.
        It will read in the first sheet in the .xlsx file.
        It will do a quick check on the columns in the sheets to prevent the most basic errors.
        """

        info(f"Reading in data from {self.other_sheet.name}")
        self.copyright_data = pl.read_excel(self.other_sheet.path)
        self.latest_file_date = self.other_sheet.modified.strftime("%Y-%m-%d")
        info(
            f"Read {len(self.copyright_data)} items from {self.other_sheet.name}. Item was lasted changed on {self.latest_file_date}"
        )

        if "workflow_status" not in self.copyright_data.columns:
            self.copyright_data = self.copyright_data.with_columns(
                pl.Series("workflow_status", ["ToDo"] * len(self.copyright_data))
            )
        if "retrieved_from_copyright_on" not in self.copyright_data.columns:
            if "added_to_sheet_on" not in self.copyright_data.columns:
                self.copyright_data = self.copyright_data.with_columns(
                    pl.Series(
                        "retrieved_from_copyright_on",
                        [self.latest_file_date] * len(self.copyright_data),
                    )
                )
            else:
                self.copyright_data = self.copyright_data.rename(
                    {"added_to_sheet_on": "retrieved_from_copyright_on"}
                )

        self.latest_file_date = max(
            self.copyright_data.select(pl.col("retrieved_from_copyright_on"))
            .to_series()
            .to_list()
        )
        self.copyright_data = self.copyright_data.select(self.column_order)

    def read_copyright_export(self) -> None:
        """
        Reads in data from the latest copyright export file in the copyright dir.

        """
        info(
            f"Reading in newest Copyright Data from directory: {self.dirs['copyright_export']}"
        )
        try:
            all_files = self.dirs["copyright_export"].files
            self.latest_file = max(all_files, key=lambda x: x.created)
            self.latest_file_date = self.latest_file.created.strftime("%Y-%m-%d")
            info(
                f"Selected newest copyright export file:\n          {self.latest_file.name}\n          created @ {self.latest_file_date}"
            )
            self.raw_copyright_data = pl.read_excel(self.latest_file.path)

        except FileNotFoundError:
            warn(f"No files found in {self.dirs['copyright_export']}")
            raise typer.Exit(code=1)
        except PermissionError:
            warn(f"Permission denied to read {self.latest_file.name}")
            raise typer.Exit(code=1)
        except ValueError:
            warn("No files found in {self.dirs['copyright_export']}")
            raise typer.Exit(code=1)

    def process_copyright_export(self) -> None:
        """
        Process the raw copyright data:
        rename column headers, add extra columns, format some data, and match to faculty.

        If 'only_changes' it will compare this data to the items present in the faculty sheets,
        and only include new items in the export.
        """
        if self.copyright_data.is_empty():
            self.copyright_data = self.raw_copyright_data.rename(
                lambda col: col.replace(" ", "_")
                .replace("#", "count_")
                .replace("*", "x")
                .lower()
            ).with_columns(
                pl.Series(
                    "retrieved_from_copyright_on",
                    [self.latest_file_date] * len(self.raw_copyright_data),
                ),
                pl.Series("workflow_status", ["ToDo"] * len(self.raw_copyright_data)),
                pl.col("last_change").dt.strftime("%Y-%m-%d"),
                faculty=pl.col("department").replace_strict(
                    self.DEPARTMENT_MAPPING, default="Unmapped"
                ),
            )
        self.faculties = (
            self.copyright_data.select(pl.col("faculty").unique()).to_series().to_list()
        )

        # enrich copyright_data with OSIRIS data if bool is set
        if self.enrich_with_osiris_data:
            self.copyright_data = self.add_osiris_data(self.copyright_data)

        if self.only_changes:
            self.read_faculty_sheets()
            if self.faculty_sheet_data.is_empty():
                info(
                    "No faculty sheets found. Adding all items without checking for changes."
                )
            else:
                '''
                In this part, all items in self.copyright_data that are not present in self.faculty_sheet_data
                will be added to self.faculty_sheet_data.
                This is done by comparing columns material_id and last_change.
                Items are added to faculty_sheet_data if:
                - material_id is not present in self.faculty_sheet_data
                - material_id is found but last_change date is different
                '''
                not_in_faculty = self.copyright_data.join(
                    self.faculty_sheet_data, on="material_id", how="anti"
                )
                matching_id_diff_change = (
                    self.copyright_data.join(
                        self.faculty_sheet_data, on="material_id", how="inner"
                    )
                    .filter(pl.col("last_change") != pl.col("last_change_right"))
                    .select(pl.all().exclude("last_change_right"))
                    .drop_nulls(pl.col("material_id"))
                    .filter(pl.col("status") == "Deleted")
                )

                if not_in_faculty.is_empty():
                    if matching_id_diff_change.is_empty():
                        info("No new items to add!")
                        self.no_new_items = True
                    else:
                        self.copyright_data = matching_id_diff_change
                if not matching_id_diff_change.is_empty():
                    self.copyright_data = pl.concat(
                        [not_in_faculty, matching_id_diff_change]
                    )
                else:
                    self.copyright_data = not_in_faculty

                # read in all items sheets
                self.read_all_items_sheets()
                if not self.all_items_sheet_data.is_empty() and False:
                    '''
                    Here we will do the following:
                    - select rows from self.all_items_sheet_data with a manual classification
                    - find matching row in self.faculty_data
                    - if a match is found:
                        - check if faculty_data has a manual classification
                        - if not, overwrite the row in faculty_data with the row from all_items_sheet_data
                        - if yes, don't do anything
                    - if no match is found:
                        - this should be a new item, so should automatically be added through the normal process above
                    '''
                    all_items_rows_with_classification = self.all_items_sheet_data.filter(
                        pl.col("manual_classification").is_not_null() & (pl.col("manual_classification") != "-") & (pl.col("manual_classification") != "")
                    )
                    copyright_data_rows_without_classifications = self.faculty_sheet_data.filter(
                        pl.col("manual_classification").is_null() | (pl.col("manual_classification") == "") | (pl.col("manual_classification") == "-")
                    )
                    rows_to_be_updated = copyright_data_rows_without_classifications.join(
                        all_items_rows_with_classification,
                        on="material_id",
                        how="inner"
                    )
                    material_ids_with_new_cip_classification = rows_to_be_updated.select(
                        pl.col("material_id")
                    ).to_series().to_list()

                    if not material_ids_with_new_cip_classification:
                        info("No new CIP classifications found.")
                    else:
                        manual_classification_updates = self.all_items_sheet_data.filter(
                            pl.col("material_id").is_in(material_ids_with_new_cip_classification)
                        )

                        manual_classification_updates = manual_classification_updates.with_columns([
                            pl.col(col).replace("-", None) for col in manual_classification_updates.columns
                        ]).drop_nulls('manual_classification')
                        manual_classification_updates.write_excel('cip_updates.xlsx')
                        manual_classification_updates = manual_classification_updates.select(
                            pl.col("material_id"),
                            pl.col("manual_classification"),
                            pl.col("scope"),
                            pl.col("remarks"),
                        )

                        info(f"Updating copyright data with {manual_classification_updates.shape[0]} new CIP classifications.")

                        joined_df = self.faculty_sheet_data.join(manual_classification_updates, on="material_id", how="inner")
                        update_columns = manual_classification_updates.columns[1:]

                        updated_df = joined_df.with_columns([
                            pl.when(pl.col(f"{col}_right").is_not_null())
                            .then(pl.col(f"{col}_right"))
                            .otherwise(pl.col(col))
                            .alias(col)
                            for col in update_columns
                        ])

                        updated_df = updated_df.drop([f"{col}_right" for col in update_columns])

                        # add new rows to self.copyright_data
                        already_present_ids = self.copyright_data.select(pl.col("material_id")).to_series().to_list()
                        new_ids = set(material_ids_with_new_cip_classification) - set(already_present_ids)
                        new_ids = list(new_ids)
                        present_in_both = set(already_present_ids) & set(material_ids_with_new_cip_classification)
                        present_in_both = list(present_in_both)
                        rows_to_add = updated_df.filter(pl.col("material_id").is_in(new_ids))
                        if self.copyright_data.is_empty():
                            self.copyright_data = rows_to_add
                        else:
                            self.copyright_data = pl.concat([self.copyright_data, rows_to_add])
                        print(f'added {(rows_to_add.shape[0])} rows with updated cip classifications to self.copyright_data')
                        if len(present_in_both) > 0:
                            # for each material id in present_in_both,
                            # find the row in self.copyright_data
                            # replace that entire row with the corresponding row from updated_df
                            for material_id in present_in_both:
                                row_to_replace = self.copyright_data.filter(pl.col("material_id") == material_id).to_series().to_list()[0]
                                new_row = updated_df.filter(pl.col("material_id") == material_id).to_series().to_list()[0]
                                self.copyright_data = self.copyright_data.with_columns(
                                    pl.when(pl.col("material_id") == material_id)
                                    .then(new_row)
                                    .otherwise(row_to_replace)
                                )
                                print(f'replaced row with material_id {material_id}')

    def create_faculty_sheets(self) -> None:
        """
        Splits the processed copyright data into one sheet per faculty
        and exports the result to excel sheets.
        """
        info(f'Exporting new items to faculty sheets for date {self.latest_file_date}')
        if self.faculties:
            self.faculties.sort()
        for faculty in self.faculties:
            if faculty in self.COURSE_MAPPING:
                self.create_programme_sheets(faculty)

            faculty_dir = Directory(self.dirs["faculties"].full / faculty)
            if faculty is None or faculty == "":
                faculty = "no_faculty_found"

            filename = f"{faculty}_{self.latest_file_date}.xlsx"
            i = 1
            while os.path.exists(faculty_dir.full / filename):
                filename = f"{faculty}_{self.latest_file_date}_{i}.xlsx"
                i += 1

            faculty_data = self.copyright_data.filter(pl.col("faculty") == faculty)
            gap = ' '*(15-len(faculty))
            if faculty_data.is_empty():
                warn(f'{faculty}:{gap}{faculty_data.shape[0]} (no new items, skipping)')
                continue
            else:
                info(f'{faculty}:{gap}{faculty_data.shape[0]}')
            if not self.disable_writes:
                faculty_data.write_excel(faculty_dir.full / filename)
                self.finalize_sheet(File(str(faculty_dir.full / filename)))

    def create_programme_sheets(self, faculty: str) -> None:
        """
        For a given faculty, split processed copyright data into one sheet per programme.
        Export to faculty_dir / per_programme / programme_name}_{date}.xlsx
        """

        programme_dir = Directory(self.dirs['faculties'].full / faculty / "per_programme")
        course_to_sheet: dict[str,str] = self.COURSE_MAPPING[faculty]
        data: list[dict[str,pl.DataFrame]] = []
        info(f'creating programme sheets for {faculty}')
        for course, group in course_to_sheet.items():
            course_data = self.copyright_data.filter(pl.col("department") == course)
            gap = ' '*(40-len(course))
            if course_data.is_empty():
                warn(f'{course}:{gap}{course_data.shape[0]} (no new items, skipping)')
                continue
            else:
                info(f'retrieved programme sheet data for {course}')
                info(f'{course}:{gap}{course_data.shape[0]}')
                data.append({'sheet':group, 'data': course_data})
        final_data: dict[str,pl.DataFrame] = {}
        for item in data:
            if item.get('sheet') in final_data:
                final_data[item.get('sheet')] = pl.concat([final_data[item.get('sheet')], item['data']])
            else:
                final_data[item.get('sheet')] = item['data']

        if not self.disable_writes:
            for groupname, df in final_data.items():
                filename = programme_dir.full / f'{groupname}_{self.latest_file_date}.xlsx'
                df.write_excel(filename)
                self.finalize_sheet(File(str(filename)))
                info(f'created programme sheet {groupname}_{self.latest_file_date}.xlsx')

    def finalize_sheet(self, file: File) -> None:
        """
        Add the data entry sheet to a fresh faculty excel file.
        This sheet will contain a selection of columns, will be styled,
        and contain dropdowns for data entry.
        """

        wb = openpyxl.load_workbook(filename=str(file.path))
        wb.active.title = "Complete data"

        # Create the Data Entry sheet
        # -----------------------------

        entry_sheet: openpyxl.worksheet.worksheet.Worksheet = wb.create_sheet(
            "Data entry", index=1
        )

        keep_cols = [6, 34, 14, 16, 17, 13, 1, 7, 8, 9, 28, 3, 5, ""]
        col_names = [
            "url",
            "workflow_status",
            "manual_classification",
            "scope",
            "remarks",
            "ml_prediction",
            "material_id",
            "filename",
            "title",
            "owner",
            "author",
            "department",
            "course_name",
        ]
        max_row = 0
        url = False
        max_col_letter = None
        for new_col, old in enumerate(keep_cols, start=1):
            if isinstance(old, str):
                break
            else:
                if not max_col_letter:
                    max_col_letter = "A"
                else:
                    max_col_letter = chr(ord(max_col_letter) + 1)
            for row, cell in enumerate(
                wb.active.iter_rows(min_col=old, max_col=old, values_only=True), start=1
            ):
                if row == 1 and cell[0] == "url":
                    url = True
                elif row == 1:
                    url = False

                if not url:
                    entry_sheet.cell(row=row, column=new_col).value = cell[0]
                else:
                    entry_sheet.cell(row=row, column=new_col).hyperlink = cell[0]
                if row > max_row:
                    max_row = row

        for col, name in enumerate(col_names, start=1):
            entry_sheet.cell(row=1, column=col, value=name)

        # Dropdown items for certain cells
        # -----------------------------------
        dropdowndata = [
            (2, "B", '"ToDo,Done,InProgress"'),  # workflow status
            (
                3,
                "C",
                '"open access, eigen materiaal - powerpoint, eigen materiaal - overig, lange overname, eigen materiaal - titelindicatie"',
            ),  # manual classification
        ]
        for colnum, col_letter, itemlist in dropdowndata:
            dv = openpyxl.worksheet.datavalidation.DataValidation(
                type="list", formula1=itemlist, allow_blank=False
            )
            dv.error = "Please select a valid option from the list"
            dv.errorTitle = "Invalid option"
            dv.prompt = "Please select from the list"
            dv.promptTitle = "List selection"
            entry_sheet.add_data_validation(dv)
            dv.add(f"{col_letter}2:{col_letter}{max_row}")

        # Style as table
        # -----------------
        table = ExcelTable(displayName="DataEntry", ref=f"A1:{max_col_letter}{max_row}")
        tabstyle = TableStyleInfo(
            name=f"TableStyleMedium{self.style_iter}",
            showRowStripes=True,
        )
        self.style_iter = self.style_iter + 1
        table.tableStyleInfo = tabstyle
        entry_sheet.add_table(table)
        wb.save(filename=str(file.path))

    def style_region_as_table(self, sheet: openpyxl.worksheet.worksheet.Worksheet, table_name: str, max_row: int, max_col_letter: str, start_row: int = 1, start_col_letter: str = 'A', style: None | str = None) -> openpyxl.worksheet.worksheet.Worksheet:
        '''
        NOTE: This function is currently unused because it doesn't work properly during testing.
        Functions needs fixing :)
        Add an Excel Table including styling to a given sheet.

        Returns the sheet with the table added.
        Parameters:
            sheet: openpyxl.worksheet.worksheet.Worksheet
                The sheet to add the table to.
            table_name: str
                The name of the table.
            max_row: int
                The maximum row number of the table.
            max_col_letter: str
                The final col of the table
            start_row: int, default 1
                The row number to start the table at.
            start_col_letter: str, default "A"
                the col to start the table at.
            style: str | None, default None
                The number of the table style to use. If None, the next available style will be used (from style_iter).
        '''
        info(f'adding table to area {start_col_letter}{start_row}:{max_col_letter}{max_row} with name {table_name}')
        table = ExcelTable(displayName=table_name, ref=f"{start_col_letter}{start_row}:{max_col_letter}{max_row}")

        if not style:
            style = self.style_iter
            self.style_iter = self.style_iter + 1

        style = f"TableStyleMedium{style}"
        tabstyle = TableStyleInfo(
            name=style,
            showRowStripes=True,
        )

        info(f'adding table style {tabstyle}')
        table.tableStyleInfo = tabstyle
        info(f'Table info: {table}')
        sheet.add_table(table)
        return sheet

    def create_all_items_sheet(self) -> None:
        """
        Add all items in the current Copyright data to a single sheet.
        """
        if not self.no_new_items:
            filename = f"all_items_{self.latest_file_date}.xlsx"
            i = 1
            while os.path.exists(self.dirs["all_items"].full / filename):
                filename = f"all_items_{self.latest_file_date}_{i}.xlsx"
                i += 1
            if not self.disable_writes:
                self.copyright_data.write_excel(self.dirs["all_items"].full / filename)
                info(f"Created sheet: {self.dirs['all_items'].full / filename}")

    def read_faculty_sheets(self) -> None:
        """
        Reads in all data from all sheets in the faculties dir
        and stores it in self.faculty_sheet_data as a single concatted dataframe.
        """
        self.faculty_sheet_data = self.get_all_faculty_data()

    def read_all_items_sheets(self) -> None:
        """
        Reads in all data from all 'all_items' sheets
        and stores it in self.all_items_sheet_data as a single concatted dataframe.
        """

        self.all_items_sheet_data = self.read_complete_data_from_sheets(self.dirs["all_items"].files_r, 'Sheet1')

    def read_complete_data_from_sheets(self, files: list[File], sheetname: str = "Complete data") -> pl.DataFrame:
        """
        Reads the data from sheet 'Complete data' for each file in 'files'.
        """
        file_data = []
        for file in files:
            if file.extension not in [".xls", ".xlsx"]:
                warn(f"{file.path} is not an excel file, skipping.")
                continue
            if 'overview' in file.name:
                info(f'skipping {file.path}')
                continue
            try:
                current_data = pl.read_excel(file.path, sheet_name=sheetname)
            except Exception:
                current_data = pl.read_excel(file.path)

            current_data = self.validate_ea_sheet(current_data, file)
            if current_data.is_empty():
                continue
            else:
                file_data.append(current_data)
        if file_data:
            result: pl.DataFrame = pl.concat(file_data)
        else:
            result = pl.DataFrame()
        return result.unique()

    def validate_ea_sheet(self, df: pl.DataFrame, file: File) -> pl.DataFrame:
        """
        For a given dataframe created from an EA excel sheet,
        check the data for errors.
        If found, try to fix, else print the errors.
        If the sheet is not validated, return an empty dataframe.

        Current implementation is bare:
        - is sheet empty? if yes: print error
        - set all columns to type str

        TODO: Implement this function fully.
        TODO: handle multiple sheets in the same file

        """
        valid = True
        errlist = []
        if df.is_empty():
            valid = False
            errlist.append("Sheet is empty")
        if valid:
            # set all columns to type str
            df = df.with_columns(pl.exclude(pl.Utf8).cast(str))
            # check that the sheet has the correct columns
            ...
        if valid:
            # check the values in the columns
            ...
        if not valid:
            info(f"Errors in sheet {file}:")
            for err in errlist:
                warn(err)
            return pl.DataFrame()
        else:
            return df

    def create_import_sheet(self) -> None:
        """
        combine self.faculty_sheet_data and self.all_items_sheet_data
        clean it up
        change from UT Easy Access format to SURF CopyRight format
        create & export an .xlsx sheet that can be sent to SURF to be imported into CopyRight.
        """
        # TODO
        ...

    def add_osiris_data(self, df: pl.DataFrame) -> pl.DataFrame:
        """
        For a given df with faculty data, retrieve missing/required data from OSIRIS

        first make a list of what to search for
        then scrape the data / use api
        then enrich the df and return it

        base the df structure on self.copyright_data
        """

        enriched_df: pl.DataFrame = copy.copy(df)
        def determine_course_code(code: str, name: str) -> set | None:
            '''
            For a given course code and name, determine the correct course code(s).
            Returns a set of course codes or None if no valid course code could be found.
            '''
            found = False
            tempresults = set()
            first_try = code.split('-')[1].strip()
            if len(first_try) >= 8 and first_try.isdigit():
                tempresults.add(first_try)
                found = True
            else:
                second_try = name.split(';')[1].split('(')[0]
                for c in second_try.split(','):
                    c = c.strip()
                    if c.isdigit() and len(c) >= 8:
                        tempresults.add(c)
                        found = True
            if not found:
                warn(f'No valid course code found for {code} - {name}')
                info(f'code extraction results: {first_try}, name extraction results: {second_try}')
            return tempresults

        def get_data_from_osiris(input_number: int, jaar: int = 2024) -> dict[str, dict[str,str|list|set]]:
            startstring: str = '{"from":0,"size":25,"sort":[{"cursus_lange_naam.raw":{"order":"asc"}},{"cursus":{"order":"asc"}},{"collegejaar":{"order":"desc"}}],"aggs":{"agg_terms_collegejaar":{"filter":{"bool":{"must":[]}},"aggs":{"agg_collegejaar_buckets":{"terms":{"field":"collegejaar","size":2500,"order":{"_term":"desc"}}}}},"agg_terms_blokken_nested.periode_omschrijving":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_blokken_nested.periode_omschrijving":{"terms":{"field":"blokken_nested.periode_omschrijving","size":2500,"order":{"_term":"asc"},"exclude":"Periode: [0-9][0-9]-[0-9][0-9]-[0-9][0-9][0-9][0-9]"}},"nested_aggs":{"nested":{"path":"blokken_nested"},"aggs":{"nested_aggs":{"filter":{"bool":{"must":[]}},"aggs":{"agg_blokken_nested.periode_omschrijving_buckets":{"terms":{"field":"blokken_nested.periode_omschrijving","size":2500,"order":{"_term":"asc"},"exclude":"Periode: [0-9][0-9]-[0-9][0-9]-[0-9][0-9][0-9][0-9]"},"aggs":{"items":{"reverse_nested":{}}}}}}}}}},"agg_terms_faculteit_naam":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_faculteit_naam_buckets":{"terms":{"field":"faculteit_naam","size":2500,"order":{"_term":"asc"}}}}},"agg_terms_coordinerend_onderdeel_oms":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_coordinerend_onderdeel_oms_buckets":{"terms":{"field":"coordinerend_onderdeel_oms","size":2500,"order":{"_term":"asc"}}}}},"agg_terms_categorie_omschrijving":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_categorie_omschrijving_buckets":{"terms":{"field":"categorie_omschrijving","size":2500,"order":{"_term":"asc"}}}}},"agg_terms_voertalen.voertaal_omschrijving":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_voertalen.voertaal_omschrijving_buckets":{"terms":{"field":"voertalen.voertaal_omschrijving","size":2500,"order":{"_term":"asc"}}}}}},"post_filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"query":{"bool":{"must":[{"multi_match":{"query":'
            jaar: int = 2024 #startjaar academisch jaar, 2024 = 2024-2025
            if jaar != 2024:
                if isinstance(jaar, int):
                    startstring.replace('"2024-2025"',f'"{jaar}-{jaar+1}"')
                elif jaar == '':
                    startstring.replace('"2024-2025"','')

            code: str = f'"{input_number}"'
            endstring: str = ',"type":"phrase_prefix","fields":["cursus","cursus_korte_naam","cursus_lange_naam"],"max_expansions":200}}]}}}'
            body: str = startstring + code + endstring
            url: str = "https://utwente.osiris-student.nl/student/osiris/student/cursussen/zoeken"
            headers: dict[str,str] = {
                                "host": "utwente.osiris-student.nl",
                                "connection": "keep-alive",
                                "content-length": "2183",
                                "sec-ch-ua-platform": "\"Windows\"",
                                "authorization": "undefined undefined",
                                "cache-control": "no-cache, no-store, must-revalidate, private",
                                "pragma": "no-cache",
                                "client_type": "web",
                                "release_version": "c0d3b6a1d72bf1610166027c903b46fc10580f30",
                                "manifest": "24.46_B346_c0d3b6a1",
                                "sec-ch-ua-mobile": "?0",
                                "sec-ch-ua": "\"Google Chrome\";v=\"131\", \"Chromium\";v=\"131\", \"Not_A Brand\";v=\"24\"",
                                "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
                                "accept": "application/json, text/plain, */*",
                                "content-type": "application/json",
                                "taal": "NL",
                                "origin": "https//utwente.osiris-student.nl",
                                "sec-fetch-site": "same-origin",
                                "sec-fetch-mode": "cors",
                                "sec-fetch-dest": "empty",
                                "referer": "https//utwente.osiris-student.nl/onderwijscatalogus/extern/cursussen",
                                "accept-encoding": "gzip, deflate, br, zstd",
                                "accept-language": "en-GB,en-US;q=0.9,en;q=0.8"
                            }
            x = requests.post(
                url=url,
                headers=headers,
                data=body
            )

            results = x.json().get('hits',{}).get('hits')
            datadict = {}
            if not results:
                return
            else:
                print(str(len(results))+f' hit(s) for code {input_number} for year {jaar} - {jaar+1}.')
                for h,result in enumerate(results):
                    print(f'------- Result {h} -----------\n')
                    rawdata:dict = result.get('_source')
                    teachers = []
                    # pretty print the raw data
                    print(rawdata.keys())
                    for key, value in rawdata.items():
                        if value == "" or not value or value == [] or value == {}:
                            continue
                        gaplen = 25 - len(key)
                        if gaplen <= 0:
                            gaplen = 1
                            key = key[:21]+"..."
                        gap = " "+"─"*(gaplen-1)
                        if isinstance(value, list):
                            if len(value) == 0:
                                continue
                            if len(value) == 1:
                                print(f"{key}{gap}─ {list(value[0].values())[0]}")
                                items = list(value[0].values())[0]
                            else:
                                gap = f"{key}{gap}┬ "
                                i = 0
                                items = [list(item.values())[0] for item in value]
                                itemset = set(items)
                                items = list(itemset)
                                for item in items:
                                    i += 1
                                    if i - (len(items)) == 0:
                                        gap = " "*(len(key)+gaplen)+"└ "
                                    elif i == 2:
                                        gap = " "*(len(key)+gaplen)+"├ "
                                    if isinstance(item, dict):
                                        print(f"{gap}{list(item.values())[0]}")
                                    else:
                                        print(f"{gap}{item}")
                            if key == 'docenten':
                                teachers = items

                        else:
                            if '\n' not in str(value):
                                print(f"{key}{gap}─ {value}")
                            else:
                                lines = value.split("\n")
                                printer = f"{key}{gap}┬ "
                                i = 0
                                for line in lines:
                                    i = i+1
                                    if i - len(lines) == 0:
                                        printer = " "*(len(key)+gaplen)+"└ "
                                    elif i > 1:
                                        printer = f"{" "*(len(key)+gaplen)}├ "
                                    print(f'{printer}{line}')

                    datadict[rawdata.get('cursus')] = {
                        'cursuscode':rawdata.get('cursus'),
                        'internal_id':rawdata.get('id_cursus'),
                        'year':rawdata.get('collegejaar'),
                        'short_name':rawdata.get('cursus_korte_naam'),
                        'name':rawdata.get('cursus_lange_naam'),
                        'faculty':rawdata.get('faculteit'),
                        'faculty_long':rawdata.get('faculteit_naam'),
                        'programme':rawdata.get('coordinerend_onderdeel_oms'),
                        'ec':rawdata.get('punten'),
                        'language':[x.get('voertaal_omschrijving') for x in rawdata.get('voertalen')],
                        'notes':rawdata.get('opmerking_cursus'),
                        'category':rawdata.get('categorie_omschrijving'),
                        'teachers': teachers,
                        'contacts':set(),
                        'docenten':set(),
                        'examinators':set(),
                        'unknown_role':set()
                    }
                    print('\n')

                headers_course = {'accept':'application/json, text/plain, */*' ,
                'accept-language':'en-US,en;q=0.9,nl-NL;q=0.8,nl;q=0.7' ,
                'authorization':'undefined undefined' ,
                'cache-control':'no-cache, no-store, must-revalidate, private' ,
                'client_type':'web' ,
                'content-type':'application/json' ,
                'dnt':'1' ,
                'manifest':'24.46_B346_c0d3b6a1' ,
                'pragma':'no-cache' ,
                'priority':'u=1, i' ,
                'referer':'https://utwente.osiris-student.nl/onderwijscatalogus/extern/cursussen' ,
                'release_version':'c0d3b6a1d72bf1610166027c903b46fc10580f30' ,
                'sec-ch-ua':'"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"' ,
                'sec-ch-ua-mobile':'?0' ,
                'sec-ch-ua-platform':'"Windows"' ,
                'sec-fetch-dest':'empty' ,
                'sec-fetch-mode':'cors' ,
                'sec-fetch-site':'same-origin' ,
                'taal':'NL' ,
                'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'}
                newdatadict = datadict.copy()
                for course, data in datadict.items():
                    internal_id = data.get('internal_id')
                    url_course = f'https://utwente.osiris-student.nl/student/osiris/owc/cursussen/{internal_id}'
                    course_details = requests.get(
                        url=url_course,
                        headers=headers_course,
                    )

                    if course_details.status_code == 200:
                        course_data = course_details.json()
                        for datapoint in course_data.get('items'):
                            if datapoint.get('rubriek') == 'rubriek-docenten':
                                docentdata = datapoint.get('velden')
                        if docentdata:
                            for docentitem in docentdata:
                                if docentitem.get('waarde'):
                                    for docenttype in docentitem.get('waarde'):
                                        for persoon in docenttype.get('velden'):
                                            if docenttype.get('omschrijving') == 'Contactpersoon':
                                                newdatadict[course]['contacts'].add(persoon.get('docent'))
                                            elif docenttype.get('omschrijving') == 'Docent':
                                                newdatadict[course]['docenten'].add(persoon.get('docent'))
                                            elif docenttype.get('omschrijving') == 'Examinator':
                                                newdatadict[course]['examinators'].add(persoon.get('docent'))
                                            else:
                                                newdatadict[course]['unknown_role'].add(persoon)
                    else:
                        print('Error!')
                        print(course_details.status_code)

                print(newdatadict)
                return newdatadict
        # step 1: determine list of courseids to search for
        # each row in the df should have 1 or multiple course codes attached to it.
        # we are going to search for each of these course codes in OSIRIS.
        # we will need to extract these codes first.

        # heuristic:

        # 1. FROM COLUMN COURSE_CODE
        # - from column 'course_code', get the course code as a string
        # - Should look like YYYY - XXXXXXXXXXX - 1A, where YYYY is the year, XXXXXXXXXXX is the course code, and 1A is the period.
        # - split on '-', select the second part.
        # - course code should be numeric and (probably?) 9 digits long.
        # - period is (probably) one value from: JAAR, 1A, 1B, 2A, 2B, 3A, SEM1, SEM2, SEM3

        # example values that should result in extracted course code + period:
        # 2024-191158500-JAAR --> Course code: 191158500, Period: JAAR
        # 2024-201800005-1A --> Course code: 201800005, Period: 1A
        # 2024-202400157-1A --> Course code: 202400157, Period: 1A
        # 2024-201800236-SEM1 --> Course code: 201800236, Period: SEM1
        #
        # example values that should be processed further:
        # 2024-IDVWI-1A --> Course code: IDVWI, Period: 1A --> ERROR: not a valid course code
        # 2024-ELECMSE-1B --> Course code: ELECMSE, Period: 1B --> ERROR: not a valid course code

        # 2. IF NO COURSE CODE FOUND: EXTRACT FROM COLUMN COURSE_NAME
        # - in cases where the 'course code' is a string with only letters, it is likely this course has multiple course codes attached to it.
        # - in this case, a list of all related course codes should be extracted from the 'course name' column.
        # - retrieve the string to parse from the 'course name' column.
        # - split the string on ';'. Retrieve the second part. Split this on '(', keep only the first part. This should give you the course codes separated by commas.
        # - Each course code should consist solely of digits w/ len >= 8.
        # - if no valid codes are found, mark as 'no code found'.

        # example values that should result in extracted course codes:
        # Circuit Analysis 1 and 2; 202001116,202200163 (2024-JAAR) --> Course codes: [202001116, 202200163]
        # Characterization of Nanostructures 2023; 193700010,201600043 (2024-1A) --> Course codes: [193700010, 201600043]
        #
        # example values that should not result in extracted course codes:
        # Circuit Analysis 1 and 2; CA12,CA34 (2024-JAAR) --> Course codes: [CA12, CA34] --> ERROR: no valid course codes -> return empty list

        # first we extract the cols as lists using to_dict()
        course_data_dict = enriched_df.select(pl.col('course_code'),pl.col('course_name')).to_dict()
        course_code_list = course_data_dict.get('course_code').to_list()
        course_name_list = course_data_dict.get('course_name').to_list()

        # then we build a set of all the course codes we need to look up
        lookup_values = set()
        for code, name in zip(course_code_list, course_name_list):

            result = determine_course_code(code, name)

            lookup_values.update(result)


        if len(lookup_values) == 0:
            info('No course codes found, skipping OSIRIS data enrichment')
            return enriched_df
        else:
            info(f'Found {len(lookup_values)} course codes to look up in OSIRIS')

        # then retrieve data from OSIRIS for each of the values in lookup_values
        course_data_dict = {}
        not_found = set()
        for code in lookup_values:
            if code in course_data_dict:
                continue
            result = get_data_from_osiris(code)
            if result:
                course_data_dict.update(result)
            else:
                for year in [2023, 2022, 2021, 2020, 2019, '']:
                    warn(f'No data found for course code {code}, trying with year {year}')
                    result = get_data_from_osiris(code, 2023)
                    if result:
                        course_data_dict.update(result)
                        continue
                not_found.add(code)


        info('data for '+str(len(list(course_data_dict.keys())))+' courses found in OSIRIS from '+str(len(lookup_values))+' course codes.')
        info(f'Course codes not found: ')
        for code in not_found:
            print('            '+str(code))
        input('press any key to continue')
        # now, for each row in 'enriched_df', detect course codes (just like we did before).
        # look up that course code as a key in course_data_dict to retrieve a dict with details for that course.
        # add the following data to the row (colname -> keyname in course_data_dict[course_code]):
        #   - col 'osiris_id' = internal_id
        #   - col 'osiris_contact' = contacts
        #   - col 'osiris_programme' = programme
        #   - col 'osiris_name' = name
        #   - col 'osiris_short_name' = short_name
        #   - col 'osiris_faculty' = faculty
        #   - col 'osiris_notes' = notes

        # for each set of data, if len > 1, join as strs with ' | ' as separator.

        def add_data_to_row(row: dict, data:set, colname:str) -> dict:
            if len(data) == 0:
                row[colname] = None
            elif len(data) == 1:
                row[colname] = list(data)[0]
            else:
                data = list(data)
                data = [str(x) for x in data]
                row[colname] = ' | '.join(data)
            return row

        new_data = []
        for row in enriched_df.to_dicts():

            codes = determine_course_code(row.get('course_code'), row.get('course_name'))
            if not codes:
                continue
            internal_id = set()
            contact = set()
            programme = set()
            name = set()
            short_name = set()
            faculty = set()
            notes = set()


            for code in list(codes):
                if code in course_data_dict:
                    data = course_data_dict.get(code)
                    internal_id.add(data.get('internal_id'))
                    contact.update(data.get('contacts'))
                    programme.add(data.get('programme'))
                    name.add(data.get('name'))
                    short_name.add(data.get('short_name'))
                    faculty.add(data.get('faculty'))
                    notes.add(data.get('notes'))
                else:
                    warn(f'No data found for course code {code}')

            setcollection = [
                (internal_id, 'osiris_internal_id'),
                (contact, 'osiris_contacts'),
                (programme, 'osiris_programme'),
                (name, 'osiris_name'),
                (short_name, 'osiris_short_name'),
                (faculty, 'osiris_faculty'),
                (notes, 'osiris_notes')
            ]

            for itemset, itemcol in setcollection:
                row = add_data_to_row(row, itemset, itemcol)

            print(row)
            new_data.append(row)

        enriched_df = pl.DataFrame(new_data)
        print(enriched_df)
        input('done with enriching! press any key to continue.')
        return enriched_df

    def get_all_faculty_data(self) -> pl.DataFrame:
        """
        Read in all available faculty sheets
        and merge the 'complete data' and 'data entry' sheets for each one.
        concat all the data into a single dataframe and return it.
        """
        all_faculty_data = pl.DataFrame()
        for faculty in self.faculties:
            info(f'getting data for faculty {faculty}')
            faculty_data = self.get_faculty_data(faculty)
            if faculty_data.is_empty():
                continue
            all_faculty_data = pl.concat([all_faculty_data, faculty_data])

        return all_faculty_data.unique()

    def get_faculty_data(self, faculty: str, del_overview: bool = False) -> pl.DataFrame:
        """
        for a given faculty, read in all available faculty sheets
        and merge the 'complete data' and 'data entry' sheets for each one.
        concat all the data into a single dataframe and return it.

        Parameters:
            faculty: str
                the faculty to get the data for. Will scan through all sheets in path self.dirs['faculties'].full / faculty.
            del_overview: bool
                if True, delete the existing overview sheets for this faculty.
        """
        def join_coalesce_all(df1: pl.DataFrame, df2: pl.DataFrame, on: str, prefer_right=set()) -> pl.DataFrame:
            to_coalesce = set(df1.columns) & set(df2.columns) - set([on])
            coalesced = {c: pl.coalesce(pl.col(c + "_right"), pl.col(c))
                            if c in prefer_right else
                            pl.coalesce(pl.col(c), pl.col(c + "_right"))
                        for c in to_coalesce}
            return (
                df1.join(df2, on=on, how="full", suffix="_right")
                    .with_columns(**coalesced)
                    .drop([c + "_right" for c in to_coalesce])
                    .drop(['material_id_right'])
            )

        if faculty is None or faculty == "":
            return pl.DataFrame()

        faculty_dir = Directory(self.dirs["faculties"].full / faculty)
        faculty_files = faculty_dir.files_r

        all_faculty_data = pl.DataFrame()
        for file in faculty_files:
            if file.extension not in [".xls", ".xlsx"]:
                continue
            if 'overview' in file.name and faculty in file.name:
                if del_overview:
                    file.delete()
                continue
            elif 'overview' in file.name:
                continue
            else:
                full_data = pl.read_excel(file.path, sheet_name="Complete data")
                data_entry = pl.read_excel(file.path, sheet_name="Data entry")

                full_data = self.validate_ea_sheet(full_data, file)
                data_entry = self.validate_ea_sheet(data_entry, file)

                # merge data_entry into full_data on column material_id.
                # data from data_entry will overwrite data from full_data
                # if a col is present in data_entry, but not in full_data, it will be added
                # keep the columns in full_data that are not in data_entry

                merged_data = join_coalesce_all(full_data, data_entry, on="material_id", prefer_right=set(data_entry.columns) - {"material_id"})
                # set all columns to type str for easy concatting

                all_faculty_data = pl.concat([all_faculty_data, merged_data])

        return all_faculty_data

    def create_faculty_overview(self) -> None:
        """
        per faculty:
        Read in all available faculty sheets
        use this data to generate a single sheet with 'complete data' for all items in the faculty,
        PLUS create an overview (a pdf maybe?) with calculated data, e.g.:
            - number of items per classification
            - expected fine
            - ...
        """

        def create_programme_overviews(faculty: str) -> None:
            """
            also create an overview sheet for each programme
            if applicable
            """
            all_faculty_data = self.get_faculty_data(faculty)
            course_to_group: dict[str,str] = self.COURSE_MAPPING[faculty]
            data: list[dict[str,pl.DataFrame]] = []

            for course, group in course_to_group.items():
                programme_data = all_faculty_data.filter(pl.col("department") == course)
                if programme_data.is_empty():
                    continue
                else:
                    programme_data = programme_data.with_columns(
                        pl.col('pages_x_students').cast(pl.Int32).mul(self.fine_amount).alias('possible_fine')
                    )
                    programme_data = programme_data.with_columns(
                        infringement=pl.when(pl.col("manual_classification").is_null() |
                                            (pl.col("manual_classification") == "") |
                                            (pl.col("manual_classification") == "-"))
                                        .then(pl.lit("undetermined"))
                                        .when(pl.col("manual_classification").str.to_lowercase().str.contains("open|eigen|overig|deleted"))
                                        .then(pl.lit("no"))
                                        .when(pl.col("manual_classification").str.to_lowercase().str.contains("lange"))
                                        .then(pl.lit("yes"))
                                        .otherwise(pl.lit("maybe"))
                    )
                    data.append({'group':group, 'data': programme_data})


            final_data: dict[str,pl.DataFrame] = {}
            for item in data:
                info(f'group: {item.get("group")} --> + {item.get("data").shape[0]} items')
                if item.get('group') in final_data:
                    final_data[item.get('group')] = pl.concat([final_data[item.get('group')], item['data']])
                else:
                    final_data[item.get('group')] = item['data']
            # add columns:
            # 'possible_fine': for each row multiply col pages_x_students with 0.30 to get the amount

            # 'infringement': possible values: 'yes', 'no', 'maybe', 'undetermined'.
            # based on the value in 'manual_classification'
            # if 'manual_classification' is empty (None, "", '-', NaN): set to 'undetermined'
            # if the str in 'manual_classification' contains 'open' or 'eigen': set no 'no'
            # if 'lange overname' is in 'manual_classification': set 'yes'
            # else set to 'maybe'


            # calculate the total possible fine by adding up all values in the 'possible_fine' column
            # for all items that do not have 'no' in the 'infringement' column

            for groupname, df in final_data.items():
                if not self.disable_writes:
                    for file in Directory(self.dirs['faculties'].full / faculty / "per_programme").files:
                        if file.extension not in [".xls", ".xlsx"]:
                            continue
                        if 'overview' in file.name and groupname in file.name:
                            file.delete()
                            continue
                print(f'{groupname} has {df.shape[0]} items')
                programme_file = File(self.dirs['faculties'].full / faculty / "per_programme" / f'{groupname}_total_overview_updated_{today}.xlsx')
                if not self.disable_writes:
                    info(f'saving file with {df.shape[0]} rows to {programme_file.path}')
                    programme_data.write_excel(programme_file.path)
                else:
                    info(f'writing is disabled')

        # loop over the faculties
        # for each, read in all data and store
        overview_data: list[dict] = []
        today = datetime.now().strftime("%Y-%m-%d_%H_%M")
        self.faculties.sort()
        for faculty in self.faculties:
            if faculty in self.COURSE_MAPPING:
                create_programme_overviews(faculty)
            fac_data = {'faculty': faculty}
            all_faculty_data = self.get_faculty_data(faculty, del_overview=True)

            # add columns:
            # 'possible_fine': for each row multiply col pages_x_students with 0.30 to get the amount
            if all_faculty_data.is_empty():
                continue

            all_faculty_data = self.add_osiris_data(all_faculty_data)
            all_faculty_data = all_faculty_data.with_columns(
                pl.col('pages_x_students').cast(pl.Int32).mul(self.fine_amount).alias('possible_fine')
            )

            # 'infringement': possible values: 'yes', 'no', 'maybe', 'undetermined'.
            # based on the value in 'manual_classification'
            # if 'manual_classification' is empty (None, "", '-', NaN): set to 'undetermined'
            # if the str in 'manual_classification' contains 'open' or 'eigen': set no 'no'
            # if 'lange overname' is in 'manual_classification': set 'yes'
            # else set to 'maybe'

            all_faculty_data = all_faculty_data.with_columns(
                infringement=pl.when(pl.col("manual_classification").is_null() |
                                    (pl.col("manual_classification") == "") |
                                    (pl.col("manual_classification") == "-"))
                                .then(pl.lit("undetermined"))
                                .when(pl.col("manual_classification").str.to_lowercase().str.contains("open|eigen|overig|deleted"))
                                .then(pl.lit("no"))
                                .when(pl.col("manual_classification").str.to_lowercase().str.contains("lange"))
                                .then(pl.lit("yes"))
                                .otherwise(pl.lit("maybe"))
            )

            # calculate the total possible fine by adding up all values in the 'possible_fine' column
            # for all items that do not have 'no' in the 'infringement' column

            total_possible_fine = all_faculty_data.filter(pl.col("infringement") != "no").select(pl.sum('possible_fine')).to_series().to_list()[0]
            definitive_fine = all_faculty_data.filter(pl.col("infringement") == "yes").select(pl.sum('possible_fine')).to_series().to_list()[0]
            locale.setlocale(locale.LC_ALL, 'nl_NL.utf8')
            fac_data['total_possible_fine'] = str(locale.currency(total_possible_fine, grouping=True, symbol=True))
            fac_data['definitive_fine']= str(locale.currency(definitive_fine, grouping=True, symbol=True))
            fac_data['items_total'] = str(all_faculty_data.shape[0])
            fac_data['possible_infringements'] = str(all_faculty_data.filter(pl.col("infringement") != "no").shape[0])
            fac_data['definitive_infringements'] = str(all_faculty_data.filter(pl.col("infringement") == "yes").shape[0])
            fac_data['definitive_non_infringements'] = str(all_faculty_data.filter(pl.col("infringement") == "no").shape[0])
            fac_data['items_without_man_cl'] = str(all_faculty_data.filter(pl.col("infringement") == "undetermined").shape[0])
            fac_data['items_to_do'] = str(all_faculty_data.filter(pl.col("workflow_status") == "ToDo").shape[0])
            overview_data.append(fac_data)
            fac_file = File(self.dirs['faculties'].full / faculty / f'{faculty}_total_overview_updated_{today}.xlsx')
            info(f'saving file with {all_faculty_data.shape[0]} rows to {fac_file.path}')
            if not self.disable_writes:
                all_faculty_data.write_excel(fac_file.path)
            locale.setlocale(locale.LC_ALL, '')

        # now we have the data for all faculties, and written the excel files to disk.
        # print the overview table to the console, and export it as an html file to the faculties/overviews dir.
        cons = Console(record=True)

        datatable = Table(title=f'Faculty Overview {today}')
        datatable.add_column('Faculty', justify='right', style='yellow bold')
        datatable.add_column('Probable fine', justify='left', style='red bold')
        datatable.add_column('Max fine', justify='left')
        datatable.add_column('Items total', justify='center', style='cyan bold')
        datatable.add_column('Infringements', justify='center')
        datatable.add_column('Non-infringements', justify='center')
        datatable.add_column('To be classified', justify='center', style='magenta bold')
        datatable.add_column('To do', justify='center', style='magenta bold')
        factable = copy.deepcopy(datatable)
        for fac in overview_data:
            # save html overview for each faculty in their dir
            # also add that data to the overview html
            cur_fac_table = copy.deepcopy(factable)
            cur_fac_table.add_row(fac['faculty'],
                            fac['definitive_fine'],
                            fac['total_possible_fine'],
                            fac['items_total'],
                            fac['definitive_infringements']+f" ({int(fac['definitive_infringements'])/int(fac['items_total'])*100:.0f}%)",
                            fac['definitive_non_infringements']+f" ({int(fac['definitive_non_infringements'])/int(fac['items_total'])*100:.0f}%)",
                            fac['items_without_man_cl']+f" ({int(fac['items_without_man_cl'])/int(fac['items_total'])*100:.0f}%)",
                            fac['items_to_do'] + f" ({int(fac['items_to_do'])/int(fac['items_total'])*100:.0f}%)"
            )
            cons.print(cur_fac_table)
            cons.print('''Explanation of columns:

                - [yellow bold]Faculty[/yellow bold]: the abbreviation of the faculty -- all data is per faculty
                - [red bold]Probable fine[/red bold]: the sum of all fines for items that are manually classified as 'lange overname'
                - [bold]Max fine[/bold]: the sum of all fines for all items except those manually classified as 'eigen materiaal' or 'open access'
                - [cyan bold]Items total[/cyan bold]: the total number of items selected by the 'CopyRight tool' (i.e. all pdfs with 40+ pages)
                - [bold]Infringements[/bold]: the number of items that are manually classified as 'lange overname' -- plus as a percentage of total number of items
                - [bold]Non-infringements[/bold]: the number of items manually classified as 'eigen materiaal' or 'open access' -- plus as a percentage of total number of items
                - [magenta bold]To be classified[/magenta bold]: the number of items that are not yet manually classified -- plus as a percentage of total number of items
                ''')
            facdir = Directory(self.dirs['faculties'].full / fac['faculty'])
            # delete any old html files
            if not self.disable_writes:

                for file in facdir.files:
                    if file.name.endswith('.html'):
                        file.delete()

                cons.save_html(facdir.full / f'summary_{today}.html', theme=SVG_EXPORT_THEME)

            datatable.add_row(fac['faculty'],
                            fac['definitive_fine'],
                            fac['total_possible_fine'],
                            fac['items_total'],
                            fac['definitive_infringements']+f" ({int(fac['definitive_infringements'])/int(fac['items_total'])*100:.0f}%)",
                            fac['definitive_non_infringements']+f" ({int(fac['definitive_non_infringements'])/int(fac['items_total'])*100:.0f}%)",
                            fac['items_without_man_cl']+f" ({int(fac['items_without_man_cl'])/int(fac['items_total'])*100:.0f}%)",
                            fac['items_to_do'] + f" ({int(fac['items_to_do'])/int(fac['items_total'])*100:.0f}%)"

                        )

        # now save the complete table to all_items

        cons.print(datatable)
        cons.print('''Explanation of columns:

                - [yellow bold]Faculty[/yellow bold]: Faculty abbreviation
                - [red bold]Probable fine[/red bold]: Total fine for items that have 'lange overname' as manual classification
                - [bold]Max fine[/bold]: Total fine for all items excluding items manually classified as 'eigen materiaal' or 'open access'
                - [cyan bold]Items total[/cyan bold]: Total amount of 'lange overnames' found by the 'CopyRight tool' (all pdfs with 40+ pages)
                - [bold]Infringements[/bold]: Items manually classified as 'lange overname', (% of total)
                - [bold]Non-infringements[/bold]: Items manually classified as 'eigen materiaal' or 'open access', (% of total)
                - [magenta bold]To be classified[/magenta bold]: Items not yet manually classified, (% of total)
                - [magenta bold]To do[/magenta bold]: Items in need of action by faculty, (% of total)
                ''')
        if not self.disable_writes:
            cons.save_html(self.dirs['all_items'].full / f'faculty_overview_{today}.html', theme=SVG_EXPORT_THEME)

if __name__ == "__main__":
    typer.run(cli)
