# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "bs4",
#     "python-dotenv",
#     "httpx",
#     "lxml",
#     "openpyxl",
#     "polars",
#     "rich",
#     "typer",
#     "fastexcel",
#     "xlsxwriter",
# ]
# ///
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

from dataclasses import dataclass, field
import json
import asyncio
import os
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
import logging

from utils import Directory, File, info, cool, warn, print
from enrichment import enrich_sheets, update_osiris_data

cli_app = typer.Typer()
# suppress some annoying warnings when reading excel files
logging.getLogger("fastexcel.types.dtype").setLevel(logging.ERROR)
# load settings.env to local environment
dotenv.load_dotenv("settings.env")



class Functions(str, Enum):
    """
    CLI option for picking which functions to run, see cli()
    """

    both = "both"
    read = "read"
    export = "export"


# ----------------------------------------------------------------------------------------------------------------------
# Main functions
# ----------------------------------------------------------------------------------------------------------------------
@cli_app.command()
def cli(
    do: Annotated[
        Functions,
        typer.Option(
            case_sensitive=False,
            help="Which tool to run: read in new data, export current data, or both.",
            rich_help_panel="Functions",
        ),
    ] = "read",
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
    osiris_update: Annotated[
        bool,
        typer.Option(
            help="If enabled, will retrieve fresh osiris data for all course + people page data.",
            rich_help_panel="Functions",
        ),
    ] = False,
    other_sheet: Annotated[
        str | None,
        typer.Option(
            help="(relative) path to a xlsx sheet to read in instead of CopyRight Data.",
            rich_help_panel="Read in data from alternate source",
        ),
    ] = None,
    retrieve_all: Annotated[
        bool,
        typer.Option(
            help="If enabled, will retrieve all data from folders where users can enter data, and store it as a parquet file.",
            rich_help_panel="Functions",
        ),
    ] = False,
):
    """
    Runs the Easy Access toolkit with the specified settings. Add --help for details.\n
    Make sure that these files are present in the current working dir and contain the required info:\n\n
        'settings.env': The directories to use\n
        'department_mapping.json': The mapping between department names and faculty names\n
    \n
    Optional files:\n\n
        'course_mapping.json': The mapping between course names and programmes -- to create sheets per programme\n
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

    def delete_latest_file(subdir: Directory) -> None:
        """
        This function will be called if remove_previous is set to true.
        """
        # only if it's been created in the last 3 days
        newest_file: File = subdir.newest_file([".xlsx", ".xls"])
        if not newest_file:
            return None
        if newest_file.created > datetime.now() - timedelta(days=3):
            print(f"Current latest file in subdir:\n {newest_file.name}")
            conf = input("Remove this file? [y/N] ")
            if conf.lower() == "y":
                newest_file.delete()
                print(f"[red]Deleted[/red] {newest_file.name}.")
            else:
                print(f"[cyan]Keeping[/cyan] {newest_file.name} and moving on.\n\n")
            return conf

    if remove_previous:
        sheet_dir = Directory(os.getenv("FACULTIES_DIR"))
        all_items_dir = Directory(os.getenv("ALL_ITEMS_DIR"))
        dirlist = sheet_dir.dirs
        dirlist.append(all_items_dir)

        for subdir in dirlist:
            print(f"[green]subdir {subdir}[/green]\n------------------")
            conf = delete_latest_file(subdir)
            if not conf:
                continue
            while conf.lower() == "y":
                conf = delete_latest_file(subdir)
                if not conf:
                    break
            if subdir.dirs:
                print(f"[magenta]sub-subdir {subdir}[/magenta]\n")
                for subsubdir in subdir.dirs:
                    conf = delete_latest_file(subsubdir)
                    if not conf:
                        continue
                    while conf.lower() == "y":
                        conf = delete_latest_file(subsubdir)
                        if not conf:
                            break

    if do not in [Functions.both, Functions.read, Functions.export]:
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
        functions=do,
        only_changes=changes,
        dirs=dirs,
        other_sheet=other_sheet,
        save_files=save,
        refresh_osiris_data=osiris_update,
        retrieve_all=retrieve_all,
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
        "overviews_backup": Directory(os.getenv("OVERVIEWS_BACKUP_DIR")),
    }

    # initialize the various dataframes used to get data from / write to .xlsx files
    raw_copyright_data: pl.DataFrame = (
        pl.DataFrame()
    )  # data directly from copyright tool
    copyright_data: pl.DataFrame = (
        pl.DataFrame()
    )  # data with normalized column names & some cleanup
    faculty_sheet_data: pl.DataFrame = pl.DataFrame()  # data from the faculty sheets
    all_items_sheet_data: pl.DataFrame = (
        pl.DataFrame()
    )  # data from the 'all_items' sheet

    # this mapping is used to get the corresponding faculty from the copyright data column 'departments'
    # it should be present in the file 'department_mapping.json' in the same directory as easy_access.cli.py
    # a department_mapping.json file for the University of Twente is included in the repo
    dept_mapping_path = File("department_mapping.json")
    DEPARTMENT_MAPPING = json.load(open(dept_mapping_path.path, encoding="utf-8"))

    # this mapping is used to map courses to programmes
    # to be used in combination with the faculty / department
    course_mapping_path = File("course_mapping.json")
    COURSE_MAPPING = json.load(open(course_mapping_path.path, encoding="utf-8"))

    # list of all found/used faculties
    faculties: list[str]

    # latest copyright export file & when it was created
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

    # standard basic column order for the complete data sheets
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
        refresh_osiris_data: bool = False,
        retrieve_all: bool = False,
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

        # init parameters
        if other_sheet:
            self.other_sheet = File(other_sheet)
        self.only_changes = only_changes
        self.disable_writes = not save_files
        self.retrieve_all = retrieve_all
        self.refresh_osiris_data = refresh_osiris_data
        self.enrich_with_osiris_data = True # debug option -- should probable always be set to True

        # if dirs is set, add them to the self.dirs dict
        if dirs:
            for key, value in dirs.items():
                if value:
                    self.dirs[key] = Directory(value)

        # set the functions to run based on input param 'functions'
        self.set_functions(functions)

    def set_functions(self, functions: Functions | None) -> None:
        """
        Sets the functions to run based on the input parameter 'functions'.
        stores it in self.settings as a list of functions to run.
        Returns None.
        """
        if functions is None:
            self.settings = []

        elif functions == Functions.both:
            if not self.other_sheet:
                self.settings = [
                    self.read_copyright_export,
                    self.process_copyright_export,  # read in new data
                    self.read_all_items_sheets,
                    self.create_import_sheet,  # from the old data, create a sheet to import into CopyRight
                    self.create_faculty_sheets,
                    self.create_all_items_sheet,  # create new sheets with new data
                    self.create_faculty_overview,
                ]
            else:
                self.settings = [
                    self.read_other_sheet,  # read in new data
                    self.read_all_items_sheets,
                    self.read_faculty_sheets,  # read in data manually added to sheets
                    self.create_import_sheet,  # from the old data, create a sheet to import into CopyRight
                    self.create_faculty_sheets,
                    self.create_all_items_sheet,  # create new sheets with new data
                    self.create_faculty_overview,
                ]

        elif functions == Functions.read:
            if not self.other_sheet:
                self.settings = [
                    self.read_copyright_export,
                    self.process_copyright_export,  # read in new data
                    self.create_faculty_sheets,
                    self.create_all_items_sheet,  # create new sheets with new data
                    self.create_faculty_overview,
                ]
            else:
                self.settings = [
                    self.read_other_sheet,
                    self.process_copyright_export,  # read in new data
                    self.create_faculty_sheets,
                    self.create_all_items_sheet,  # create new sheets with new data
                    self.create_faculty_overview,
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

    def run(self) -> None:
        """
        Runs the functions as specified in the settings dict.
        """
        if self.retrieve_all:
            # will retrieve all data from the directories where users can enter data
            # and store it as a parquet file and csv file in the root dir
            self.retrieve_all_data()
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
            # cast all columns to str
            self.raw_copyright_data = self.raw_copyright_data.with_columns(
                pl.exclude(pl.Utf8).cast(str)
            )

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
                pl.col("last_change")
                .str.replace(r"^-$", "")
                .str.strip_chars()
                .str.strptime(pl.Date, "%Y-%m-%d", strict=False)
                .dt.strftime("%Y-%m-%d"),
                faculty=pl.col("department").replace_strict(
                    self.DEPARTMENT_MAPPING, default="Unmapped"
                ),
            )
        self.faculties = (
            self.copyright_data.select(pl.col("faculty").unique()).to_series().to_list()
        )
        # refresh OSIRIS data if bool is set
        if self.refresh_osiris_data:
            asyncio.run(update_osiris_data(self.copyright_data))

        # enrich copyright_data with OSIRIS data if bool is set
        if self.enrich_with_osiris_data:
            self.copyright_data = enrich_sheets(self.copyright_data)
            # set dtype of all columns to str
            self.copyright_data = self.copyright_data.with_columns(
                pl.exclude(pl.Utf8).cast(str)
            )
        if self.only_changes:
            self.read_faculty_sheets(include_overview=False)
            if self.faculty_sheet_data.is_empty():
                info(
                    "No faculty sheets found. Adding all items without checking for changes."
                )
            else:
                """
                In this part, all items in self.copyright_data that are not present in self.faculty_sheet_data
                will be added to self.faculty_sheet_data.
                This is done by comparing columns material_id and last_change.
                Items are added to faculty_sheet_data if:
                - material_id is not present in self.faculty_sheet_data
                - material_id is found but last_change date is different
                """
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
                    """
                    Here we will do the following:
                    - select rows from self.all_items_sheet_data with a manual classification
                    - find matching row in self.faculty_data
                    - if a match is found:
                        - check if faculty_data has a manual classification
                        - if not, overwrite the row in faculty_data with the row from all_items_sheet_data
                        - if yes, don't do anything
                    - if no match is found:
                        - this should be a new item, so should automatically be added through the normal process above
                    """
                    all_items_rows_with_classification = (
                        self.all_items_sheet_data.filter(
                            pl.col("manual_classification").is_not_null()
                            & (pl.col("manual_classification") != "-")
                            & (pl.col("manual_classification") != "")
                        )
                    )
                    copyright_data_rows_without_classifications = (
                        self.faculty_sheet_data.filter(
                            pl.col("manual_classification").is_null()
                            | (pl.col("manual_classification") == "")
                            | (pl.col("manual_classification") == "-")
                        )
                    )
                    rows_to_be_updated = (
                        copyright_data_rows_without_classifications.join(
                            all_items_rows_with_classification,
                            on="material_id",
                            how="inner",
                        )
                    )
                    material_ids_with_new_cip_classification = (
                        rows_to_be_updated.select(pl.col("material_id"))
                        .to_series()
                        .to_list()
                    )

                    if not material_ids_with_new_cip_classification:
                        info("No new CIP classifications found.")
                    else:
                        manual_classification_updates = (
                            self.all_items_sheet_data.filter(
                                pl.col("material_id").is_in(
                                    material_ids_with_new_cip_classification
                                )
                            )
                        )

                        manual_classification_updates = (
                            manual_classification_updates.with_columns(
                                [
                                    pl.col(col).replace("-", None)
                                    for col in manual_classification_updates.columns
                                ]
                            ).drop_nulls("manual_classification")
                        )
                        manual_classification_updates.write_excel("cip_updates.xlsx")
                        manual_classification_updates = (
                            manual_classification_updates.select(
                                pl.col("material_id"),
                                pl.col("manual_classification"),
                                pl.col("scope"),
                                pl.col("remarks"),
                            )
                        )

                        info(
                            f"Updating copyright data with {manual_classification_updates.shape[0]} new CIP classifications."
                        )

                        joined_df = self.faculty_sheet_data.join(
                            manual_classification_updates, on="material_id", how="inner"
                        )
                        update_columns = manual_classification_updates.columns[1:]

                        updated_df = joined_df.with_columns(
                            [
                                pl.when(pl.col(f"{col}_right").is_not_null())
                                .then(pl.col(f"{col}_right"))
                                .otherwise(pl.col(col))
                                .alias(col)
                                for col in update_columns
                            ]
                        )

                        updated_df = updated_df.drop(
                            [f"{col}_right" for col in update_columns]
                        )

                        # add new rows to self.copyright_data
                        already_present_ids = (
                            self.copyright_data.select(pl.col("material_id"))
                            .to_series()
                            .to_list()
                        )
                        new_ids = set(material_ids_with_new_cip_classification) - set(
                            already_present_ids
                        )
                        new_ids = list(new_ids)
                        present_in_both = set(already_present_ids) & set(
                            material_ids_with_new_cip_classification
                        )
                        present_in_both = list(present_in_both)
                        rows_to_add = updated_df.filter(
                            pl.col("material_id").is_in(new_ids)
                        )
                        if self.copyright_data.is_empty():
                            self.copyright_data = rows_to_add
                        else:
                            self.copyright_data = pl.concat(
                                [self.copyright_data, rows_to_add]
                            )
                        print(
                            f"added {(rows_to_add.shape[0])} rows with updated cip classifications to self.copyright_data"
                        )
                        if len(present_in_both) > 0:
                            # for each material id in present_in_both,
                            # find the row in self.copyright_data
                            # replace that entire row with the corresponding row from updated_df
                            for material_id in present_in_both:
                                row_to_replace = (
                                    self.copyright_data.filter(
                                        pl.col("material_id") == material_id
                                    )
                                    .to_series()
                                    .to_list()[0]
                                )
                                new_row = (
                                    updated_df.filter(
                                        pl.col("material_id") == material_id
                                    )
                                    .to_series()
                                    .to_list()[0]
                                )
                                self.copyright_data = self.copyright_data.with_columns(
                                    pl.when(pl.col("material_id") == material_id)
                                    .then(new_row)
                                    .otherwise(row_to_replace)
                                )
                                print(f"replaced row with material_id {material_id}")

    def create_faculty_sheets(self) -> None:
        """
        Splits the processed copyright data into one sheet per faculty
        and exports the result to excel sheets.
        """
        info(f"Exporting new items to faculty sheets for date {self.latest_file_date}")
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
            gap = " " * (15 - len(faculty))
            if faculty_data.is_empty():
                warn(f"{faculty}:{gap}{faculty_data.shape[0]} (no new items, skipping)")
                continue
            else:
                info(f"{faculty}:{gap}{faculty_data.shape[0]}")
            if not self.disable_writes:
                faculty_data.write_excel(faculty_dir.full / filename)
                self.finalize_sheet(
                    File(str(faculty_dir.full / filename)), faculty_data
                )

    def create_programme_sheets(self, faculty: str) -> None:
        """
        For a given faculty, split processed copyright data into one sheet per programme.
        Export to faculty_dir / per_programme / programme_name}_{date}.xlsx
        """

        programme_dir = Directory(
            self.dirs["faculties"].full / faculty / "per_programme"
        )
        course_to_sheet: dict[str, str] = self.COURSE_MAPPING[faculty]
        data: list[dict[str, pl.DataFrame]] = []
        info(f"creating programme sheets for {faculty}")
        for course, group in course_to_sheet.items():
            course_data = self.copyright_data.filter(pl.col("department") == course)
            gap = " " * (40 - len(course))
            if course_data.is_empty():
                warn(f"{course}:{gap}{course_data.shape[0]} (no new items, skipping)")
                continue
            else:
                info(f"retrieved programme sheet data for {course}")
                info(f"{course}:{gap}{course_data.shape[0]}")
                data.append({"sheet": group, "data": course_data})
        final_data: dict[str, pl.DataFrame] = {}
        for item in data:
            if item.get("sheet") in final_data:
                final_data[item.get("sheet")] = pl.concat(
                    [final_data[item.get("sheet")], item["data"]]
                )
            else:
                final_data[item.get("sheet")] = item["data"]

        if not self.disable_writes:
            for groupname, df in final_data.items():
                filename = (
                    programme_dir.full / f"{groupname}_{self.latest_file_date}.xlsx"
                )
                df.write_excel(filename)
                self.finalize_sheet(File(str(filename)), df)
                info(
                    f"created programme sheet {groupname}_{self.latest_file_date}.xlsx"
                )

    def finalize_sheet(self, file: File, data: pl.DataFrame) -> None:
        """
        New implementation of finalize_sheet
        this function mainly build the second sheet for data entry.
        Input: an excel file with the complete data, and a dataframe with that same data to be processed for the data entry sheet

        Adds the sheet to the workbook and saves it, doesnt return any data.
        """
        from openpyxl.worksheet.filters import (
            FilterColumn,
            Filters,
        )
        from openpyxl.styles import Alignment, NamedStyle

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

        @dataclass
        class DataEntrySheet:
            """
            Use to add a dateentry sheet to an excel file.
            Has functions to add data from dataframe, format as table, add datavalidation, and save
            """

            sheet_name: str
            cols: list[
                ColInfo
            ]  # a list with the cols in order of appearance from left to right
            table_style: TableStyleInfo
            workbook: openpyxl.Workbook
            sheet: openpyxl.worksheet.worksheet.Worksheet = field(init=False)
            file_path: str
            max_row: int = 0
            word_wrap_style: Alignment = NamedStyle(
                name="wordwrap", alignment=Alignment(wrapText=True)
            )

            def __post_init__(self):
                self.sheet = wb.create_sheet(self.sheet_name, index=1)

            def add_data(self, data: pl.DataFrame) -> None:
                self.max_row = data.shape[0]
                colnum = 0

                for col in self.cols:
                    colnum += 1
                    if col.new_name:
                        col_name = col.new_name
                    else:
                        col_name = col.name
                    if col.is_new:
                        # create new coldata
                        col_data = [col.default_val] * self.max_row
                    else:
                        # retrieve coldata from dataframe
                        col_data = data.select(pl.col(col.name)).to_series().to_list()
                        if col.default_val != "":
                            for item_num, item in enumerate(col_data):
                                if item == "" or not item:
                                    col_data[item_num] = col.default_val

                    self.sheet.cell(1, colnum).value = col_name
                    for row, cell_data in enumerate(col_data, start=2):
                        if not cell_data:
                            self.sheet.cell(row, colnum).value = cell_data
                            continue

                        if col.is_url:
                            if "/" not in cell_data:
                                self.sheet.cell(row, colnum).value = cell_data
                            else:
                                self.sheet.cell(row, colnum).value = (
                                    ".../" + cell_data.split("/")[-1]
                                )
                            self.sheet.cell(row, colnum).hyperlink = cell_data
                            if len(self.sheet.cell(row, colnum).value) > col.max_width:
                                col.max_width = len(self.sheet.cell(row, colnum).value)
                            if len(self.sheet.cell(row, colnum).value) > 40:
                                col.count_max_width_over_40 += 1

                        else:
                            self.sheet.cell(row, colnum).value = cell_data
                            if len(cell_data) > col.max_width:
                                col.max_width = len(cell_data)
                            if len(cell_data) > 40:
                                col.count_max_width_over_40 += 1

                for colnum, col in enumerate(self.cols):
                    col_letter = chr(ord("A") + colnum)
                    colnum += 1
                    if col.has_dropdown:
                        dv = openpyxl.worksheet.datavalidation.DataValidation(
                            type="list", formula1=col.dropdown_options, allowBlank=True
                        )
                        dv.error = "Please select a valid option from the list"
                        dv.errorTitle = "Invalid option"
                        dv.prompt = "Please select from the list"
                        dv.promptTitle = "List selection"
                        self.sheet.add_data_validation(dv)
                        if self.max_row == 1:
                            dv.add(f"{col_letter}2")
                        else:
                            dv.add(f"{col_letter}2:{col_letter}{self.max_row + 1}")
                    if col.max_width > 40 and (
                        (col.count_max_width_over_40 > 5)
                        or (col.count_max_width_over_40 > self.max_row - 2)
                    ):
                        # Too much long items: cap width to 40 & enable word wrap for this col
                        for row in range(2, self.max_row + 1):
                            self.sheet.cell(row, colnum).style = self.word_wrap_style
                        self.sheet.column_dimensions[col_letter].bestFit = False
                        self.sheet.column_dimensions[col_letter].width = 40
                    else:
                        # Acceptable width, don't enable word wrap but fit width to contents
                        self.sheet.column_dimensions[col_letter].width = col.max_width

                info(f"Added data to {self.sheet_name} in file {self.file_path}.")
                self.create_table()
                self.save()

            def create_table(self) -> None:
                max_col_letter = chr(ord("A") + len(self.cols) - 1)
                table = ExcelTable(
                    displayName=self.sheet_name.replace(" ", ""),
                    ref=f"A1:{max_col_letter}{self.max_row + 1}",
                )
                table.tableStyleInfo = self.table_style
                self.sheet.add_table(table)
                info(
                    f"Created table with {self.max_row} rows and {len(self.cols)} cols in sheet {self.sheet_name} of file {self.file_path}"
                )

            def save(self) -> None:
                self.workbook.save(filename=self.file_path)
                info(f"Saved .xlsx file with DataEntrySheet to {self.file_path}")

        wb = openpyxl.load_workbook(filename=str(file.path))
        wb.active.title = "Complete data"

        tabstyle = TableStyleInfo(
            name=f"TableStyleMedium{self.style_iter}",
            showRowStripes=True,
        )
        self.style_iter = self.style_iter + 1

        sheet = DataEntrySheet(
            workbook=wb,
            sheet_name="Data entry",
            cols=[
                ColInfo("material_id"),
                ColInfo("url", is_url=True),
                ColInfo(
                    "workflow_status",
                    is_new=True,
                    is_editable=True,
                    dropdown_options='"ToDo,Done,InProgress"',
                    default_val="ToDo",
                ),
                ColInfo(
                    "manual_classification",
                    is_editable=True,
                    default_val="-",
                    dropdown_options='"open access,eigen materiaal - powerpoint,eigen materiaal - overig,lange overname,eigen materiaal - titelindicatie,anders,korte overname,middellange overname,-"',
                ),
                ColInfo("remarks", is_editable=True),
                ColInfo("ml_prediction"),
                ColInfo("filename"),
                ColInfo("title"),
                ColInfo("owner", new_name="uploaded_by"),
                ColInfo("author", new_name="detected_author"),
                ColInfo("contact_name"),
                ColInfo("contact_email"),
                ColInfo("contact_org"),
                ColInfo("osiris_catalogue_url", is_url=True),
                ColInfo("course_name", new_name="course_name_canvas"),
                ColInfo("department", new_name="programme_canvas"),
                ColInfo("osiris_programme", new_name="programme_osiris"),
                ColInfo("osiris_course_codes_found"),
                ColInfo("osiris_course_code_data_selected"),
            ],
            table_style=tabstyle,
            file_path=str(file.path),
        )

        sheet.add_data(data)

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

    def read_all_items_sheets(self) -> None:
        """
        Reads in all data from all 'all_items' sheets
        and stores it in self.all_items_sheet_data as a single concatted dataframe.
        """

        self.all_items_sheet_data = self.read_complete_data_from_sheets(
            self.dirs["all_items"].files_r, "Sheet1"
        )

    def read_complete_data_from_sheets(
        self, files: list[File], sheetname: str = "Complete data"
    ) -> pl.DataFrame:
        """
        Reads the data from sheet 'Complete data' for each file in 'files'.
        """
        file_data = []
        for file in files:
            if file.extension not in [".xls", ".xlsx"]:
                continue
            if "overview" in file.name:
                info(f"skipping {file.path}")
                continue
            try:
                current_data = pl.read_excel(
                    file.path, sheet_name=sheetname, infer_schema_length=None
                )
            except Exception:
                current_data = pl.read_excel(file.path, infer_schema_length=None)

            current_data = self.validate_ea_sheet(current_data, file)
            if current_data.is_empty():
                continue
            else:
                file_data.append(current_data)
        if file_data:
            result: pl.DataFrame = pl.concat(file_data, how="diagonal_relaxed")
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

    def read_faculty_sheets(self, include_overview: bool = True) -> None:
        """
        Reads in all data from all sheets in the faculties dir
        and stores it in self.faculty_sheet_data as a single concatted dataframe.
        """
        self.faculty_sheet_data = self.get_all_faculty_data(
            include_overview=include_overview
        )

    def get_all_faculty_data(self, include_overview: bool = True) -> pl.DataFrame:
        """
        Read in all available faculty sheets
        and merge the 'complete data' and 'data entry' sheets for each one.
        concat all the data into a single dataframe and return it.
        """
        all_faculty_data = pl.DataFrame()
        for faculty in self.faculties:
            info(f"getting data for faculty {faculty}")
            faculty_data = self.get_faculty_data(
                faculty, include_overview=include_overview
            )
            if faculty_data.is_empty():
                continue
            all_faculty_data = pl.concat(
                [all_faculty_data, faculty_data], how="diagonal_relaxed"
            )

        return all_faculty_data.unique()

    def get_faculty_data(
        self, faculty: str, del_overview: bool = False, include_overview: bool = True
    ) -> pl.DataFrame:
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

        def join_coalesce_all(
            df1: pl.DataFrame, df2: pl.DataFrame, on: str, prefer_right=set()
        ) -> pl.DataFrame:
            to_coalesce = set(df1.columns) & set(df2.columns) - set([on])
            coalesced = {
                c: pl.coalesce(pl.col(c + "_right"), pl.col(c))
                if c in prefer_right
                else pl.coalesce(pl.col(c), pl.col(c + "_right"))
                for c in to_coalesce
            }
            return (
                df1.join(df2, on=on, how="full", suffix="_right")
                .with_columns(**coalesced)
                .drop([c + "_right" for c in to_coalesce])
                .drop(["material_id_right"])
            )

        if faculty is None or faculty == "":
            return pl.DataFrame()

        faculty_dir = Directory(self.dirs["faculties"].full / faculty)
        faculty_files = faculty_dir.files_r

        all_faculty_data = pl.DataFrame()

        prefer_overview_cols = False  # set to True to give edits in total_overview file higher priority than edits in each weekly faculty excel

        total_overview = pl.DataFrame()
        overview_file: File = None
        latest_mod_date = None
        if include_overview:
            for file in faculty_files:
                if "total_overview" in file.name and faculty in file.name:
                    try:
                        total_overview_complete = pl.read_excel(
                            file.path, sheet_name="Complete data"
                        )
                        total_overview_data_entry = pl.read_excel(
                            file.path, sheet_name="Data entry"
                        )
                        total_overview = join_coalesce_all(
                            total_overview_complete,
                            total_overview_data_entry,
                            on="material_id",
                            prefer_right=set(total_overview_data_entry.columns)
                            - {"material_id"},
                        )
                    except ValueError:
                        total_overview = pl.read_excel(file.path)

                    overview_file = file
        for file in faculty_files:
            if file.extension not in [".xls", ".xlsx"]:
                continue
            elif "overview" in file.name:
                continue
            else:
                latest_mod_date = (
                    file.modified
                    if latest_mod_date is None
                    else max(latest_mod_date, file.modified)
                )
                full_data = pl.read_excel(file.path, sheet_name="Complete data")
                data_entry = pl.read_excel(file.path, sheet_name="Data entry")

                full_data = self.validate_ea_sheet(full_data, file)
                data_entry = self.validate_ea_sheet(data_entry, file)

                # merge data_entry into full_data on column material_id.
                # data from data_entry will overwrite data from full_data
                # if a col is present in data_entry, but not in full_data, it will be added
                # keep the columns in full_data that are not in data_entry

                merged_data = join_coalesce_all(
                    full_data,
                    data_entry,
                    on="material_id",
                    prefer_right=set(data_entry.columns) - {"material_id"},
                )
                merged_data = merged_data.unique(subset="material_id")
                all_faculty_data = pl.concat(
                    [all_faculty_data, merged_data], how="diagonal_relaxed"
                )
                all_faculty_data = all_faculty_data.unique(subset="material_id")

        if not total_overview.is_empty():
            if (latest_mod_date < overview_file.modified) and (
                not prefer_overview_cols
            ):
                info(
                    f"Latest mod date for overview is newer than latest mod date for any other sheet for {faculty}."
                )
                all_overview_man_class = (
                    total_overview.select(pl.col("manual_classification"))
                    .to_series()
                    .to_list()
                )
                all_overview_man_class = [
                    i for i in all_overview_man_class if i not in [None, "", "-", " "]
                ]
                all_faculty_data_man_class = (
                    all_faculty_data.select(pl.col("manual_classification"))
                    .to_series()
                    .to_list()
                )
                all_faculty_data_man_class = [
                    i
                    for i in all_faculty_data_man_class
                    if i not in [None, "", "-", " "]
                ]
                print(
                    f"{len(all_overview_man_class)} manual classifications in total_overview. {len(all_faculty_data_man_class)} manual classifications in all_faculty_data."
                )
                if len(all_faculty_data_man_class) < len(all_overview_man_class):
                    print(
                        f"all_faculty_data has less manual classifications than total_overview. Will prefer overview columns for {faculty}."
                    )
                    prefer_overview_cols = True
            if prefer_overview_cols:
                preffered_cols = set(total_overview.columns) - {"material_id"}
                all_faculty_data = join_coalesce_all(
                    all_faculty_data,
                    total_overview,
                    on="material_id",
                    prefer_right=preffered_cols,
                )
            else:
                preffered_cols = set(all_faculty_data.columns) - {"material_id"}
                all_faculty_data = join_coalesce_all(
                    total_overview,
                    all_faculty_data,
                    on="material_id",
                    prefer_right=preffered_cols,
                )
            all_faculty_data = all_faculty_data.unique(subset="material_id")

        overview_fac_dir = Directory(self.dirs["overviews_backup"].full / faculty)

        if del_overview:
            if overview_file:
                overview_file.move(overview_fac_dir.full / overview_file.name)
            else:
                for file in faculty_files:
                    if "total_overview" in file.name and faculty in file.name:
                        file.move(overview_fac_dir.full / file.name)
                        break
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
            course_to_group: dict[str, str] = self.COURSE_MAPPING[faculty]
            data: list[dict[str, pl.DataFrame]] = []

            for course, group in course_to_group.items():
                programme_data = all_faculty_data.filter(pl.col("department") == course)
                if programme_data.is_empty():
                    continue
                else:
                    programme_data = programme_data.with_columns(
                        pl.col("pages_x_students")
                        .cast(pl.Int32)
                        .mul(self.fine_amount)
                        .alias("possible_fine")
                    )
                    programme_data = programme_data.with_columns(
                        infringement=pl.when(
                            pl.col("manual_classification").is_null()
                            | (pl.col("manual_classification") == "")
                            | (pl.col("manual_classification") == "-")
                        )
                        .then(pl.lit("undetermined"))
                        .when(
                            pl.col("manual_classification")
                            .str.to_lowercase()
                            .str.contains("open|eigen|overig|deleted")
                        )
                        .then(pl.lit("no"))
                        .when(
                            pl.col("manual_classification")
                            .str.to_lowercase()
                            .str.contains("lange")
                        )
                        .then(pl.lit("yes"))
                        .otherwise(pl.lit("maybe"))
                    )
                    data.append({"group": group, "data": programme_data})

            final_data: dict[str, pl.DataFrame] = {}
            for item in data:
                info(
                    f"group: {item.get('group')} --> + {item.get('data').shape[0]} items"
                )
                if item.get("group") in final_data:
                    final_data[item.get("group")] = pl.concat(
                        [final_data[item.get("group")], item["data"]]
                    )
                else:
                    final_data[item.get("group")] = item["data"]
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
            overview_fac_programme_dir = Directory(
                self.dirs["overviews_backup"].full / faculty / "per_programme"
            )
            for groupname, df in final_data.items():
                if not self.disable_writes:
                    for file in Directory(
                        self.dirs["faculties"].full / faculty / "per_programme"
                    ).files:
                        if file.extension not in [".xls", ".xlsx"]:
                            continue
                        if "overview" in file.name and groupname in file.name:
                            file.move(overview_fac_programme_dir.full / file.name)
                            continue
                print(f"{groupname} has {df.shape[0]} items")
                programme_file = File(
                    self.dirs["faculties"].full
                    / faculty
                    / "per_programme"
                    / f"{groupname}_total_overview_updated_{today}.xlsx"
                )
                if not self.disable_writes:
                    info(
                        f"saving file with {df.shape[0]} rows to {programme_file.path}"
                    )
                    programme_data.write_excel(programme_file.path)
                else:
                    info(f"writing is disabled")

        # loop over the faculties
        # for each, read in all data and store
        overview_data: list[dict] = []
        today = datetime.now().strftime("%Y-%m-%d_%H_%M")
        self.faculties.sort()
        for faculty in self.faculties:
            if faculty in self.COURSE_MAPPING:
                create_programme_overviews(faculty)
            fac_data = {"faculty": faculty}
            all_faculty_data = self.get_faculty_data(faculty, del_overview=True)

            # add columns:
            # 'possible_fine': for each row multiply col pages_x_students with 0.30 to get the amount
            if all_faculty_data.is_empty():
                continue

            all_faculty_data = all_faculty_data.with_columns(
                pl.col("pages_x_students")
                .cast(pl.Int32)
                .mul(self.fine_amount)
                .alias("possible_fine")
            )

            # 'infringement': possible values: 'yes', 'no', 'maybe', 'undetermined'.
            # based on the value in 'manual_classification'
            # if 'manual_classification' is empty (None, "", '-', NaN): set to 'undetermined'
            # if the str in 'manual_classification' contains 'open' or 'eigen': set no 'no'
            # if 'lange overname' is in 'manual_classification': set 'yes'
            # else set to 'maybe'

            all_faculty_data = all_faculty_data.with_columns(
                infringement=pl.when(
                    pl.col("manual_classification").is_null()
                    | (pl.col("manual_classification") == "")
                    | (pl.col("manual_classification") == "-")
                )
                .then(pl.lit("undetermined"))
                .when(
                    pl.col("manual_classification")
                    .str.to_lowercase()
                    .str.contains("open|eigen|overig|deleted")
                )
                .then(pl.lit("no"))
                .when(
                    pl.col("manual_classification")
                    .str.to_lowercase()
                    .str.contains("lange")
                )
                .then(pl.lit("yes"))
                .otherwise(pl.lit("maybe"))
            )

            # calculate the total possible fine by adding up all values in the 'possible_fine' column
            # for all items that do not have 'no' in the 'infringement' column

            total_possible_fine = (
                all_faculty_data.filter(pl.col("infringement") != "no")
                .select(pl.sum("possible_fine"))
                .to_series()
                .to_list()[0]
            )
            definitive_fine = (
                all_faculty_data.filter(pl.col("infringement") == "yes")
                .select(pl.sum("possible_fine"))
                .to_series()
                .to_list()[0]
            )
            locale.setlocale(locale.LC_ALL, "nl_NL.utf8")
            fac_data["total_possible_fine"] = str(
                locale.currency(total_possible_fine, grouping=True, symbol=True)
            )
            fac_data["definitive_fine"] = str(
                locale.currency(definitive_fine, grouping=True, symbol=True)
            )
            fac_data["items_total"] = str(all_faculty_data.shape[0])
            fac_data["possible_infringements"] = str(
                all_faculty_data.filter(pl.col("infringement") != "no").shape[0]
            )
            fac_data["definitive_infringements"] = str(
                all_faculty_data.filter(pl.col("infringement") == "yes").shape[0]
            )
            fac_data["definitive_non_infringements"] = str(
                all_faculty_data.filter(pl.col("infringement") == "no").shape[0]
            )
            fac_data["items_without_man_cl"] = str(
                all_faculty_data.filter(pl.col("infringement") == "undetermined").shape[
                    0
                ]
            )
            fac_data["items_to_do"] = str(
                all_faculty_data.filter(pl.col("workflow_status") == "ToDo").shape[0]
            )
            overview_data.append(fac_data)
            fac_file = File(
                self.dirs["faculties"].full
                / faculty
                / f"{faculty}_total_overview_updated_{today}.xlsx"
            )
            info(
                f"saving file with {all_faculty_data.shape[0]} rows to {fac_file.path}"
            )
            if not self.disable_writes:
                all_faculty_data.write_excel(fac_file.path)
                self.finalize_sheet(fac_file, all_faculty_data)
            locale.setlocale(locale.LC_ALL, "")

        # now we have the data for all faculties, and written the excel files to disk.
        # print the overview table to the console, and export it as an html file to the faculties/overviews dir.
        cons = Console(record=True)

        datatable = Table(title=f"Faculty Overview {today}")
        datatable.add_column("Faculty", justify="right", style="yellow bold")
        datatable.add_column("Probable fine", justify="left", style="red bold")
        datatable.add_column("Max fine", justify="left")
        datatable.add_column("Items total", justify="center", style="cyan bold")
        datatable.add_column("Infringements", justify="center")
        datatable.add_column("Non-infringements", justify="center")
        datatable.add_column("To be classified", justify="center", style="magenta bold")
        datatable.add_column("To do", justify="center", style="magenta bold")
        factable = copy.deepcopy(datatable)
        for fac in overview_data:
            # save html overview for each faculty in their dir
            # also add that data to the overview html
            cur_fac_table = copy.deepcopy(factable)
            cur_fac_table.add_row(
                fac["faculty"],
                fac["definitive_fine"],
                fac["total_possible_fine"],
                fac["items_total"],
                fac["definitive_infringements"]
                + f" ({int(fac['definitive_infringements']) / int(fac['items_total']) * 100:.0f}%)",
                fac["definitive_non_infringements"]
                + f" ({int(fac['definitive_non_infringements']) / int(fac['items_total']) * 100:.0f}%)",
                fac["items_without_man_cl"]
                + f" ({int(fac['items_without_man_cl']) / int(fac['items_total']) * 100:.0f}%)",
                fac["items_to_do"]
                + f" ({int(fac['items_to_do']) / int(fac['items_total']) * 100:.0f}%)",
            )
            cons.print(cur_fac_table)
            cons.print("""Explanation of columns:

                - [yellow bold]Faculty[/yellow bold]: the abbreviation of the faculty -- all data is per faculty
                - [red bold]Probable fine[/red bold]: the sum of all fines for items that are manually classified as 'lange overname'
                - [bold]Max fine[/bold]: the sum of all fines for all items except those manually classified as 'eigen materiaal' or 'open access'
                - [cyan bold]Items total[/cyan bold]: the total number of items selected by the 'CopyRight tool' (i.e. all pdfs with 40+ pages)
                - [bold]Infringements[/bold]: the number of items that are manually classified as 'lange overname' -- plus as a percentage of total number of items
                - [bold]Non-infringements[/bold]: the number of items manually classified as 'eigen materiaal' or 'open access' -- plus as a percentage of total number of items
                - [magenta bold]To be classified[/magenta bold]: the number of items that are not yet manually classified -- plus as a percentage of total number of items
                """)
            facdir = Directory(self.dirs["faculties"].full / fac["faculty"])
            # delete any old html files
            if not self.disable_writes:
                for file in facdir.files:
                    if file.name.endswith(".html"):
                        file.delete()

                cons.save_html(
                    facdir.full / f"summary_{today}.html", theme=SVG_EXPORT_THEME
                )

            datatable.add_row(
                fac["faculty"],
                fac["definitive_fine"],
                fac["total_possible_fine"],
                fac["items_total"],
                fac["definitive_infringements"]
                + f" ({int(fac['definitive_infringements']) / int(fac['items_total']) * 100:.0f}%)",
                fac["definitive_non_infringements"]
                + f" ({int(fac['definitive_non_infringements']) / int(fac['items_total']) * 100:.0f}%)",
                fac["items_without_man_cl"]
                + f" ({int(fac['items_without_man_cl']) / int(fac['items_total']) * 100:.0f}%)",
                fac["items_to_do"]
                + f" ({int(fac['items_to_do']) / int(fac['items_total']) * 100:.0f}%)",
            )

        # now save the complete table to all_items

        cons.print(datatable)
        cons.print("""Explanation of columns:

                - [yellow bold]Faculty[/yellow bold]: Faculty abbreviation
                - [red bold]Probable fine[/red bold]: Total fine for items that have 'lange overname' as manual classification
                - [bold]Max fine[/bold]: Total fine for all items excluding items manually classified as 'eigen materiaal' or 'open access'
                - [cyan bold]Items total[/cyan bold]: Total amount of 'lange overnames' found by the 'CopyRight tool' (all pdfs with 40+ pages)
                - [bold]Infringements[/bold]: Items manually classified as 'lange overname', (% of total)
                - [bold]Non-infringements[/bold]: Items manually classified as 'eigen materiaal' or 'open access', (% of total)
                - [magenta bold]To be classified[/magenta bold]: Items not yet manually classified, (% of total)
                - [magenta bold]To do[/magenta bold]: Items in need of action by faculty, (% of total)
                """)
        if not self.disable_writes:
            cons.save_html(
                self.dirs["all_items"].full / f"faculty_overview_{today}.html",
                theme=SVG_EXPORT_THEME,
            )

    def retrieve_all_data(self) -> pl.DataFrame:
        """
        Goes through all files to retrieve all available data.
        Then, for each material_id, grab only unique rows.
        Keep track of where the data came from.

        Returns a dataframe with all unique rows including provenance.
        """
        found_dfs = dict()
        today = f"{datetime.now().isoformat(sep=' ', timespec='minutes')}"
        cool("Retrieving all data. Please wait, this can take a while.")
        numfiles = 0
        dirs = {
            "faculties": self.dirs.get("faculties"),
            "all_items": self.dirs.get("all_items"),
        }

        for name, dir in dirs.items():
            cur_df = pl.DataFrame()
            for file in dir.files_r:
                if file.extension not in [".xls", ".xlsx", ".csv"]:
                    continue
                if file.extension in [".xls", ".xlsx"]:
                    try:
                        file_content = pl.read_excel(file.path, sheet_id=0)
                        numfiles += 1
                    except Exception as e:
                        print(f"Couldnt read file {file.path}: {e}")
                        continue

                    if not isinstance(file_content, dict):
                        file_content = {"sheet1": file_content}

                    for dataframe in file_content.values():
                        dataframe = dataframe.with_columns(
                            pl.lit(str(file.name)).alias("from_file")
                        )
                        if cur_df.is_empty():
                            cur_df = dataframe
                            continue
                        cur_df = pl.concat([cur_df, dataframe], how="diagonal_relaxed")

            info(f"Retrieved {cur_df.shape[0]} rows from dir {name}")
            cur_df = cur_df.unique()
            info(f"{cur_df.shape[0]} remaining after removing duplicate rows")
            if "material_id" in cur_df.columns:
                cur_df_mat_ids = cur_df.select("material_id").to_series().to_list()
                unique_cur_df_mat_ids = set(cur_df_mat_ids)
                info(
                    f"found {len(cur_df_mat_ids)} rows in dir {name} for {len(unique_cur_df_mat_ids)} unique material ids."
                )
            found_dfs[name] = cur_df

        info(f"Done retrieving data from {numfiles} files. Now merging all.")
        full_df = pl.DataFrame()
        for df in found_dfs.values():
            if full_df.is_empty():
                full_df = df
                continue
            full_df = pl.concat([full_df, df], how="diagonal_relaxed")

        info(f"after concatting all dfs, full_df has {full_df.shape[0]} rows")
        select_cols = ["from_file"]
        if "material_id" in full_df.columns:
            select_cols.append("material_id")
        if "Material id" in full_df.columns:
            select_cols.append("Material id")
        material_id_and_file = full_df.select(select_cols).to_dicts()

        final_dict: dict[int, list] = dict()
        for row in material_id_and_file:
            if row.get("material_id"):
                mat_id = int(row.get("material_id"))
            elif row.get("Material id"):
                mat_id = int(row.get("Material id"))
            if not mat_id:
                continue
            if mat_id in final_dict:
                if row.get("from_file") not in final_dict.get(mat_id):
                    final_dict[mat_id].append(row.get("from_file"))
            else:
                final_dict[mat_id] = list()
                final_dict[mat_id].append(row.get("from_file"))

        update_dict: dict[str, str] = dict()
        for material_id, filenames in final_dict.items():
            if len(filenames) > 1:
                update_dict[str(material_id)] = ", ".join(filenames)
            else:
                update_dict[str(material_id)] = filenames[0]

        full_df = full_df.drop("from_file")

        full_df = full_df.unique(
            subset=[
                "material_id",
                "manual_classification",
                "remarks",
                "workflow_status",
            ]
        )
        full_df = full_df.with_columns(
            [
                pl.col("material_id").replace(update_dict).alias("from_file"),
                pl.lit(today).alias("last_sheet_update"),
            ]
        )

        info(
            f"{full_df.shape[0]} rows remaining after selecting unique rows based on material_id, manual classification, remarks, and workflow_status."
        )
        info(f"Now comparing data with previously stored items.")
        df_merged = pl.DataFrame()
        try:
            stored_df = pl.read_parquet("full_df.parquet")
        except Exception as e:
            warn(f"Error reading full_df.parquet: {e}.")
            df_merged = full_df

        if df_merged.is_empty():
            if "last_sheet_update" not in stored_df.columns:
                info(
                    f"stored_df has no last_sheet_update data. Overwriting stored data with new data"
                )
                df_merged = full_df
            else:
                compare_cols = [
                    "manual_classification",
                    "remarks",
                    "workflow_status",
                    "retrieved_from_copyright_on",
                    "last_change",
                    "status",
                ]
                full_df = full_df.with_columns()  # ...finish this
                # This expression should make sure data from full_df is kept if it contains updated info.
                # Rows that have remained the same or are missing data are not updated; i.e. keep the stored_df row.

                df_merged = (
                    full_df.join(
                        stored_df, on="material_id", how="left", suffix="_stored"
                    )
                    .with_columns(
                        [
                            pl.fold(
                                True,
                                lambda acc, x: acc & x,
                                [
                                    (pl.col(c) == pl.col(f"{c}_stored"))
                                    for c in compare_cols
                                ],
                            ).alias("all_match")
                        ]
                    )
                    .with_columns(
                        [
                            pl.fold(
                                False,
                                lambda acc, x: acc | x,
                                [
                                    (
                                        pl.col(c).is_null()
                                        & pl.col(f"{c}_stored").is_not_null()
                                    )
                                    for c in compare_cols
                                ],
                            ).alias("any_missing_in_full_df")
                        ]
                    )
                    .with_columns(
                        [
                            pl.when(pl.col("any_missing_in_full_df"))
                            .then(pl.col(f"{c}_stored"))
                            .otherwise(pl.col(c))
                            .alias(c)
                            for c in compare_cols
                        ]
                        + [
                            pl.when(pl.col("any_missing_in_full_df"))
                            .then(pl.col("last_sheet_update_stored"))
                            .otherwise(
                                pl.when(pl.col("all_match"))
                                .then(pl.col("last_sheet_update_stored"))
                                .otherwise(pl.col("last_sheet_update"))
                            )
                            .alias("last_sheet_update")
                        ]
                    )
                )
                dropcols = [c for c in df_merged.columns if "_stored" in c]
                df_merged = df_merged.drop(dropcols).drop(
                    ["all_match", "any_missing_in_full_df"]
                )

        df_merged.write_parquet("full_df.parquet")
        df_merged.write_csv("full_data.csv")

if __name__ == "__main__":
    cli_app()
