import asyncio
import os
from collections import defaultdict

import polars as pl

from easy_access.db.base import init
from easy_access.db.ingest import load_base_data, load_raw_copyright_data
from easy_access.db.retrieve import retrieve_copyright_items, retrieve_full_data
from easy_access.db.update import DataSource, update_copyright_items
from easy_access.settings import (
    COURSE_MAPPING,
    SETTINGS,
    DirSetting,
    EasyAccessSettings,
)
from easy_access.sheets.analysis import create_faculty_overviews
from easy_access.sheets.enrichment import update_osiris_data
from easy_access.sheets.sheet import (
    create_export_sheet,
    finalize_sheet,
    read_copyright_export,
    store_complete_data,
)
from easy_access.utils import Directory, File, cool, info, print, warn


class EasyAccessTool:
    """
    This class contains all the actual functionality of the script.
    For an overview see the comments & docstrings per function, plus readme.md.
    """

    faculties: list[str] = []
    # latest copyright export file & when it was created
    latest_file_date: str

    def __init__(self, settings: EasyAccessSettings) -> None:
        """
        Parameter:
            settings: EasyAccessSettings object containing all the settings for the script.
        """

        self.settings = settings
        self.functions: list[callable] = []
        self.dirs = settings.dirs
        self.disable_writes = settings.disable_writes
        self.copyright_data = pl.DataFrame()
        self.mat_ids_on_disk = set()

        self.only_changes = settings.only_changes
        self.refresh_osiris_data = settings.refresh_osiris_data
        self.enrich_with_osiris_data = settings.enrich_with_osiris_data
        self.only_retrieve_missing_osiris_data = (
            settings.only_retrieve_missing_osiris_data
        )
        self.style_iter = 2

        self.set_functions(settings.export)

    def set_functions(self, export: bool) -> None:
        """
        Sets the functions to run based on the input parameter 'export'.
        stores it in self.settings as a list of functions to run.
        """

        self.functions.extend(
            [
                self.process_raw_copyright_data,
                self.create_overviews,
                self.create_faculty_sheets,
                self.create_all_items_sheet,
            ]
        )

        if export:
            self.functions.extend([self.create_export_sheet])

    def _load_data(self):
        """
        Loads all relevant data sources: Qlik export, DB, overview sheets, weekly sheets.
        Populates self.raw_qlik_df, self.initial_db_df, self.weekly_dfs, self.overview_dfs, self.sheet_errors.
        """
        from easy_access.sheets.sheet import (
            read_copyright_export,
            read_overview_sheets,
            read_weekly_sheets,
        )

        self.latest_file_date, self.raw_qlik_df = read_copyright_export()
        self.initial_db_df = retrieve_copyright_items()
        self.weekly_dfs: dict[str, pl.DataFrame] = {}
        self.overview_dfs: dict[str, pl.DataFrame] = {}

        self.sheet_errors = defaultdict(list)
        # Faculties may be set by settings or from Qlik data
        if self.settings.faculty:
            self.faculties = [self.settings.faculty]
        else:
            self.faculties = (
                self.raw_qlik_df.select(pl.col("faculty").unique())
                .to_series()
                .sort()
                .to_list()
            )
        for faculty in self.faculties:
            try:
                self.overview_dfs[faculty] = read_overview_sheets(faculty)
            except Exception as e:
                self.sheet_errors[faculty].append(f"Error reading overview sheets: {e}")
            try:
                self.weekly_dfs[faculty] = read_weekly_sheets(faculty)
            except Exception as e:
                self.sheet_errors[faculty].append(f"Error reading weekly sheets: {e}")
        # Write errors to files
        for faculty, errors in self.sheet_errors.items():
            if errors:
                error_log_path = (
                    self.dirs[DirSetting.FACULTIES_DIR].full
                    / faculty
                    / "sheet_errors.txt"
                )
                with open(error_log_path, "w", encoding="utf-8") as f:
                    for err in errors:
                        f.write(err + "\n")

    def _synchronize_data(self):
        """
        Central synchronization step: applies updates in the correct order and with correct priorities.
        """
        # Step 1: Overview sheets
        for _, overview_df in self.overview_dfs.items():
            if not overview_df.is_empty():
                asyncio.get_event_loop().run_until_complete(
                    update_copyright_items(overview_df, DataSource.OVERVIEW_SHEET)
                )
        # Step 2: Qlik (raw export)
        if not self.raw_qlik_df.is_empty():
            asyncio.get_event_loop().run_until_complete(
                update_copyright_items(self.raw_qlik_df, DataSource.RAW_QLIK_DATA)
            )
        # Step 3: Weekly sheets
        for _, weekly_df in self.weekly_dfs.items():
            if not weekly_df.is_empty():
                asyncio.get_event_loop().run_until_complete(
                    update_copyright_items(weekly_df, DataSource.WEEKLY_SHEET)
                )

    def _generate_output_sheets(self):
        """
        Generates all output sheets (overviews, faculty sheets, all items sheet) from the final DB state.
        """
        self.remove_current_overviews()
        # Use the final DB state for all outputs
        from easy_access.sheets.analysis import create_faculty_overviews

        faculty_dict = {}
        for faculty in self.faculties:
            data = retrieve_full_data(selected_faculties=faculty)
            if not data.is_empty():
                faculty_dict[faculty] = data
        self.style_iter = create_faculty_overviews(
            faculty_dict, self.style_iter, self.disable_writes
        )
        self.create_faculty_sheets()
        self.create_all_items_sheet()

    def run(self) -> None:
        """
        Centralized run: load all data, synchronize, then generate outputs.
        """
        info("Starting EasyAccessTool run (centralized sync mode)")
        self._load_data()
        self._synchronize_data()
        self.copyright_data = retrieve_full_data()
        self._generate_output_sheets()
        if self.settings.export:
            self.create_export_sheet()

    def process_raw_copyright_data(self) -> None:
        """
        Reads in the latest copyright export (using read_copyright_export).

        """
        file = None
        try:
            if not isinstance(self.settings.other_sheet, File):
                file = File(self.settings.other_sheet)
        except Exception:
            pass

        self.latest_file_date, self.copyright_data = read_copyright_export(file)
        if self.copyright_data.is_empty():
            warn(
                "No new Copyright data found to process! No new items will be added. Checking if there are other changes..."
            )
        else:
            # make sure base data is loaded

            fresh_db = asyncio.get_event_loop().run_until_complete(init())
            if fresh_db:
                asyncio.get_event_loop().run_until_complete(load_base_data())
            # load new data into db
            asyncio.get_event_loop().run_until_complete(
                load_raw_copyright_data(self.copyright_data)
            )

        # retrieve full data from db
        self.copyright_data = self.clean_and_validate_df(retrieve_copyright_items())

        # get osiris data for the new items (or refresh all depending on settings)
        if self.refresh_osiris_data:
            asyncio.get_event_loop().run_until_complete(
                update_osiris_data(
                    self.copyright_data, self.only_retrieve_missing_osiris_data
                )
            )

        # set faculty names
        self.faculties = (
            self.copyright_data.select(pl.col("faculty").unique())
            .to_series()
            .sort()
            .to_list()
        )
        if self.settings.faculty:
            # only include data for the given selected faculty
            info(f"Selected single faculty: {self.settings.faculty}")
            self.faculties = [self.settings.faculty]

        # determine which material_ids are already on stored in the faculty sheets
        updated_items, mat_ids = asyncio.get_event_loop().run_until_complete(
            self.update_db_from_faculty_sheets()
        )
        print(
            f"{len(mat_ids)} material_ids found in faculty sheets, {len(self.copyright_data)} items currently in copyright_data."
        )

        if mat_ids:
            self.mat_ids_on_disk = mat_ids
        if updated_items:
            # if items were updated, refresh the data for the final time
            self.copyright_data = retrieve_copyright_items()
            self.copyright_data = self.clean_and_validate_df(self.copyright_data)

        cool(
            f"process copyright export done. {self.copyright_data.shape[0]} rows in self.copyright_data."
        )

    def create_faculty_sheets(self) -> None:
        """
        Splits the processed copyright data into one sheet per faculty
        and exports the result to excel sheets.
        """
        if self.disable_writes:
            warn("Writes are disabled. Skipping programme & faculty sheet creation.")
            return
        info(f"Exporting new items to faculty sheets for date {self.latest_file_date}")
        int_mat_ids = [int(x) for x in self.mat_ids_on_disk if x]

        filtered_data: pl.DataFrame = retrieve_full_data(
            excluded_material_ids=int_mat_ids
        )
        if filtered_data.is_empty() and self.only_changes:
            warn("No new items found to export to faculty sheets.")
            return
        if self.faculties:
            self.faculties.sort()
        for faculty in self.faculties:
            gap = " " * (15 - len(faculty))
            faculty_data: pl.DataFrame = filtered_data.filter(
                pl.col("faculty") == faculty
            )
            if faculty_data.is_empty():
                warn(f"{faculty}:{gap}{faculty_data.shape[0]} (no new items, skipping)")
                continue

            if faculty in COURSE_MAPPING:
                self.create_programme_sheets(faculty, input_data=faculty_data)

            faculty_dir = Directory(self.dirs[DirSetting.FACULTIES_DIR].full / faculty)
            if faculty is None or faculty == "":
                faculty = "no_faculty_found"
            filename = f"{faculty}_{self.latest_file_date}.xlsx"
            i = 1
            while os.path.exists(faculty_dir.full / filename):
                filename = f"{faculty}_{self.latest_file_date}_{i}.xlsx"
                i += 1
            else:
                info(f"{faculty}:{gap}{faculty_data.shape[0]}")
            store_complete_data(faculty_dir.full / filename, faculty_data)
            self.style_iter = finalize_sheet(
                File(str(faculty_dir.full / filename)), faculty_data, self.style_iter
            )

    def create_programme_sheets(
        self, faculty: str, input_data: pl.DataFrame | None = None
    ) -> None:
        """
        For a given faculty, split processed copyright data into one sheet per programme.
        Export to faculty_dir / per_programme / programme_name}_{date}.xlsx
        """

        if self.disable_writes:
            warn("Writes are disabled. Skipping programme & faculty sheet creation.")
            return

        programme_dir = Directory(
            self.dirs[DirSetting.FACULTIES_DIR].full / faculty / "per_programme"
        )
        course_to_sheet: dict[str, str] = COURSE_MAPPING[faculty]
        data: list[dict[str, pl.DataFrame]] = []
        if not isinstance(input_data, pl.DataFrame):
            input_data = self.copyright_data
        if input_data.is_empty():
            warn(f"No data for {faculty} -- skipping programme sheet creation.")
            return
        info(f"creating programme sheets for {faculty}")
        for course, group in course_to_sheet.items():
            course_data = input_data.filter(pl.col("department") == course)
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

        for groupname, df in final_data.items():
            filename = programme_dir.full / f"{groupname}_{self.latest_file_date}.xlsx"

            store_complete_data(filename, df)
            self.style_iter = finalize_sheet(File(str(filename)), df, self.style_iter)
            info(f"created programme sheet {groupname}_{self.latest_file_date}.xlsx")

    def create_all_items_sheet(self) -> None:
        """
        Add all filtered items in the current Copyright data to a single sheet.
        """
        if self.disable_writes:
            warn("Writes are disabled. Skipping create_all_items_sheet.")
            return

        filtered_data = self.copyright_data.filter(
            ~pl.col("material_id").is_in(self.mat_ids_on_disk)
        )
        if filtered_data.is_empty() and self.only_changes:
            warn("No new items found to export to all items sheet.")
            return

        filename = f"all_items_{self.latest_file_date}.xlsx"
        i = 1
        while os.path.exists(self.dirs[DirSetting.ALL_ITEMS_DIR].full / filename):
            filename = f"all_items_{self.latest_file_date}_{i}.xlsx"
            i += 1
        store_complete_data(
            self.dirs[DirSetting.ALL_ITEMS_DIR].full / filename, filtered_data
        )
        info(f"Created sheet: {self.dirs[DirSetting.ALL_ITEMS_DIR].full / filename}")

    def clean_and_validate_df(self, df: pl.DataFrame) -> pl.DataFrame:
        """
        Current implementation is bare:
        - if sheet is not empty:
            - set all columns to type str
            - replace truncated url values

        TODO: Implement this function fully.
        """

        if not df.is_empty():
            # set all columns to type str
            df = df.with_columns(pl.exclude(pl.Utf8).cast(str))

            # replace truncated url values
            if "url" in df.columns:
                df = df.with_columns(
                    pl.col("url").str.replace(
                        r"\.{3}", "https://utwente.instructure.com/files"
                    )
                )
            if "osiris_catalogue_url" in df.columns:
                df = df.with_columns(
                    pl.col("osiris_catalogue_url").str.replace(
                        r"\.{3}", "https://utwente.instructure.com/files"
                    )
                )
            if "is_duplicate" in df.columns:
                df = df.with_columns(
                    pl.col("is_duplicate")
                    .str.replace("0", "FALSE")
                    .str.replace("1", "TRUE")
                )
        return df

    def remove_current_overviews(self) -> None:
        if self.disable_writes:
            warn("Writes are disabled. Skipping remove_current_overviews.")
            return

        for faculty in self.faculties:
            overview_fac_dir = Directory(
                self.dirs[DirSetting.FACULTIES_DIR].full / faculty
            )
            movedir = Directory(self.dirs[DirSetting.OVERVIEWS_BACKUP].full / faculty)
            for file in overview_fac_dir.files_r:
                if "total_overview" in file.name and faculty in file.name:
                    if (
                        SETTINGS.backup_settings.backup_overviews
                        and "llm" not in file.name
                    ):
                        file.move(movedir.full / file.name)
                    else:
                        file.delete()

    def create_overviews(self) -> None:
        """
        1. Creates overview sheets with data per faculty (and per programme if found in COURSE_MAPPING)
        2. Creates tables with summary data and prints them to the console + stores them as .html files
        3. couples data with llm classification data if found and creates llm overview sheets
        No parameters, will pull the data from disk for each faculty in self.faculties.
        """

        faculty_dict: dict[str, pl.DataFrame] = {}
        if not self.faculties:
            self.process_raw_copyright_data()
            if not self.faculties:
                warn("No faculties detected in current data. Cannot produce overviews.")
                return

        # retrieve full data from db -- all items

        self.remove_current_overviews()
        for faculty in self.faculties:
            if not faculty or faculty == "" or faculty == "Unmapped":
                continue
            data = retrieve_full_data(selected_faculties=faculty)
            if data.is_empty():
                continue
            faculty_dict[faculty] = data
        self.style_iter = create_faculty_overviews(
            faculty_dict, self.style_iter, self.disable_writes
        )

    def create_export_sheet(self) -> None:
        for faculty in self.faculties:
            if not faculty or faculty == "" or faculty == "Unmapped":
                continue
            info(f"creating export sheet for {faculty}")
            create_export_sheet(faculty=faculty)
