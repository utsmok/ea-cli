import asyncio
import os

import polars as pl

from easy_access.db.base import init
from easy_access.db.ingest import load_base_data, load_raw_copyright_data
from easy_access.db.retrieve import retrieve_copyright_items, retrieve_full_data
from easy_access.db.update import update_copyright_items
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

"""
    TODO: Make changes to the logic of importing data / syncing up sheets and DB.


    Currently, updating / merging data is done all over the codebase, and there are cases where this results in data loss, errors, or inconsistencies.
    Part of the update logic is in this file (main.py), part in sheets/sheets.py and, part in db/update.py and db/ingest.py.

    We need clear priorities of how we handle merging data from the possible data sources:
        - raw copyright export files (folder raw_copyright_data, excel files)
        - weekly sheets (one sheet per week per faculty, in faculty_sheets dir with a subdir for each faculty)
        - overview sheets (one sheet per faculty, in the same subdir as the weekly sheets)
        - sqlite db (db.sqlite3, not user editable)

    Let's walk through each source and discuss how we handle the various cases.

    # How to handle each data source

    1. Raw copyright export file
        Items **always** initially enter the dataset through a raw copyright export file.
        These are ingested into the sqlite db first after standardization and adding some basic fields (e.g. faculty).

        If an item is already present in the db, we need to compare only a few fields, as most fields are either never updated in the source data, or updates are not relevant for us:
            Always overwrite:
            - status (published/unpublished/deleted)
            - last change
            - students registered, pages * students

            Compare but don't change things (?)
            these fields are entered in the sheets first and read back into the raw data, so our sheets are the source of truth.
            So, dont' update sheets with the raw data, but compare to check for errors/inconsistencies:
            - Manual classification
            - Manual identifier
            - Scope
            - Remarks
            - Auditor

    2. Weekly sheets
        These sheets are made once a week, and never updated/changed by the script. Only shows items that are new in the import of that week.
        Users can change a few fields in the data entry sheet:
        - Manual classification
        - Scope
        - Remarks
        - Workflow status

        Any changes in this field should overwrite the db data for these fields, with one exception: the overview sheet can also be used by users to change this. This needs some more thought / work

    3. Overview sheets
        These sheets are refreshed every week. First, all the data from the overview sheet is read to memory, we create a new weekly sheet, and update the DB using heuristics to determine priorities/changes/conflicts.
        Then the DB should be the source of truth. A new overview sheet is created, which is basically a snapshot of the DB at that moment.

        Users can also change the same fields as in the weekly sheets in this sheet:
        - Manual classification
        - Scope
        - Remarks
        - Workflow status

        In case of conflicts between weekly sheets and overview sheets, we need to determine priorities. This needs some more thought / work.

    4. SQLite DB

        This is used as the source of truth for the current state of the copyright data.
        I believe all details on how to store/update the data are already mentioned above.

    5. User input from webapp

        The (experimental) web frontend for this app can be used to add or change Workflow status, manual classification, and remarks.
        Any changes made using this webapp should directly update the DB, and probably also update the overview sheets -- but we'll need to think about the best way to do this, so for now we'll just ignore this and only update the DB.

    # Order of operations

    Let's walk through the order of operations for each script run when using batch mode (only way to run the script currently):

    A. Read in data from the various source into memory
        1. New items from Qlik: Latest raw copyright data excel file from Qlik or, if given, the 'other_sheet' file
        2. Weekly sheets: Read in all weekly sheets from the faculty_sheets dir (one sheet per week per faculty)
        3. Overview sheets: Read in all overview sheets from the faculty_sheets dir (one sheet per faculty)
        4. DB: Read in all data from the sqlite db

    B. Standardize, normalize, cleanup, verify the ingested data
        - this is done separately for each source, no comparisons yet
        - Required, otherwise we cannot compare the data or might be reading in invalid data, etc
        - NEW: keep track of errors in the weekly + overview sheets per faculty, write these to a separate error sheet per faculty to let the users know what to fix

    C. Use priorities / heuristics to compare the data to determine the current absolute state
        - This is the most important step, and needs to be done carefully to avoid data loss or errors
        - Currently partly done in main.py, partly in db/update.py, and partly in db/ingest.py -- this needs to be cleaned up and made more consistent

        - Let's use this order of operations:

        1. As the overview sheets are remade every week as a direct copy of the DB, start by comparing the DB with the overview sheets row-by-row using material_id as pk.
            - item in overview sheet but not in db: should not happen, mark as error
            - difference in any of the fields below: update the db with the overview sheet data:
                - workflow status
                - manual classification
                - remarks

        Then delete the overview sheets - we've ingested that data and will remake them at the end.
        Next step:

        2. Compare DB with the raw copyright data - pk is material_id
            - item not in db: create new item in db based on the raw copyright data
            - item in db: overwrite the db data with the raw copyright data for the following fields:
                - status (published/unpublished/deleted)
                - last change
                - students registered, pages * students

        Done with raw copyright data. Now on to the weekly sheets:

        3. Compare DB with the weekly sheets - pk is material_id
            - item in sheet but not in db: should not happen, mark as error
            - if there is a difference in any of the fields below, use a detailed heuristic to determine which data to keep, see below.
                - workflow status
                - manual classification
                - remarks

        heuristic for determining which data to keep:
        1. If the current db value is empty, use the value from the weekly sheet
        2. If the current db value is not empty:
            - If available, compare the change date of the db value (using the itemupdate log) with the change date of the weekly sheet. Keep the most recent value.
            - Else, if the workflow status of the weekly sheet is 'higher' than the db value, keep all weekly sheet values; and vice versa. Priority order: Done > InProgress > ToDo
            - Else, if the workflow status is equal but manual classification is different, use priority order for manual classification:
                Open Access, [korte/middel/lange] overname, eigen materiaal [powerpoint/titelindicatie/overig], onbekend, licentie beschikbaar, niet geanalyseerd, in onderzoek, verwijderverzoek verstuurd
            - If all are the same except remarks, merge the strings in both remark fields (if there is overlap in the text, don't add it twice, e.g. "hello this is a remark" + "this is a remark, but better" = "hello this is a remark, but better")



"""


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

    def run(self) -> None:
        """
        Runs the functions as specified in the settings dict.
        """

        for func in self.functions:
            info(f"running {func.__name__}")
            func()

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

    async def update_db_from_faculty_sheets(self) -> tuple[bool, set[str]]:
        """
        Retrieves data from all faculty sheets, and sends items with changes to the database for updating.
        This function is async because it calls the async function update_copyright_items(update_df).

        returns a tuple with:
        bool: True if items to update were selected, False if no items were selected.
        set[str]: a list of all distinct material_ids present in all faculty sheets.

        Bit more details:
        For -all- faculties sheets, retrieve items from the 'data entry' sheet.
        Compare [selected cols] of each item (by matching on material_id) to self.copyright_data.
        perform a comparison to decide which items might need updating. Concat all those to update_df.
        Then send the update_df to the db to update using update_copyright_items(update_df), which will handle the actual db update and detailed comparisons.

        """

        def compare(
            primary: pl.DataFrame, other: pl.DataFrame, select_cols: list[str]
        ) -> pl.DataFrame:
            """
            compare rows based on material_id -- so match up rows from primary to rows in other
            then decide what to do with the primary row:

            if a row is in primary but not in other, KEEP the row
            else compare the cols in select_cols
            if there is no difference (so the cell vals in all cols in select_cols are equal), DROP the row
            else, if any of the vals in other is empty but filled in primary, KEEP the row
            if both have equal amount of missing values, KEEP the row
            all other cases, DROP the row

            Ensure both dataframes have required columns
            """

            cols_in_primary = primary.columns
            cols_in_other = other.columns
            primary_selected = [col for col in select_cols if col in cols_in_primary]
            other_selected = [col for col in select_cols if col in cols_in_other]
            initial_select_cols = select_cols
            # now only select the cols that are in both dataframes
            select_cols = [
                col
                for col in select_cols
                if col in primary_selected and col in other_selected
            ]

            if not select_cols:
                warn(
                    "No columns to compare between dataframes. Skipping comparison; returning primary dataframe."
                )
                return primary
            if len(select_cols) == 1:
                warn(
                    f"Only one column to compare: {select_cols}. Skipping comparison; returning primary dataframe."
                )
                return primary
            if len(select_cols) != len(initial_select_cols):
                warn(
                    f"Not all selected columns are present in both dataframes. Selecting only the common columns: {select_cols}"
                )

            # Get rows in primary but not in other
            not_in_other = primary.join(
                other.select(select_cols), on="material_id", how="anti"
            )

            # Get matching rows to compare
            matching = primary.join(
                other.select(select_cols),
                on="material_id",
                how="inner",
                suffix="_other",
            )

            # Keep rows if:
            # 1. Any values are null in other, but filled in primary
            # 2. Values are different, primary value is non-null and non-empty
            cols_to_compare = [c for c in select_cols if c != "material_id"]

            conditions = []
            for col in cols_to_compare:
                other_col = f"{col}_other"
                # Keep if other is null but primary has value
                conditions.append(
                    (pl.col(other_col).is_null())
                    & (pl.col(col).is_not_null())
                    & (pl.col(col) != "")
                    & (pl.col(col) != "-")
                )
                # Keep if values are different and primary is not null/empty
                conditions.append(
                    (pl.col(col) != pl.col(other_col))
                    & (pl.col(col).is_not_null())
                    & (pl.col(col) != "")
                    & (pl.col(col) != "-")
                )

            different_vals = matching.filter(pl.any_horizontal(conditions)).select(
                cols_in_primary
            )

            return pl.concat([not_in_other, different_vals], how="diagonal_relaxed")

        select_cols = [
            "material_id",
            "workflow_status",
            "remarks",
            "manual_classification",
        ]
        material_ids = set()
        update_df: pl.DataFrame = pl.DataFrame()
        for faculty in self.faculties:
            # get all .xlsx files except llm_classification files
            fac_dir = Directory(self.dirs[DirSetting.FACULTIES_DIR].full / faculty)
            files = fac_dir.files_r
            files = [
                f
                for f in files
                if f.extension == ".xlsx" and "llm_classification" not in f.name
            ]
            if not files:
                warn(f"No files found for faculty {faculty}.")
                continue
            for file in files:
                # load data entry sheet for file and process
                try:
                    data_entry = pl.read_excel(
                        file.path, sheet_name=SETTINGS.data_settings.data_entry_name
                    )
                except Exception as e:
                    warn(f"Error reading {file.path}: {e}")
                    continue
                data_entry = self.clean_and_validate_df(data_entry)
                if data_entry.is_empty():
                    warn(f"No data found in {file.path}.")
                    continue
                material_ids.update(
                    data_entry.select(pl.col("material_id"))
                    .to_series()
                    .unique()
                    .to_list()
                )

                # compare primary df (data_entry) to other (self.copyright_data, update_df)
                # If no rows remaining: continue
                # Else, do the same comparison as above but now compare data_entry to update_df
                # finally concat any remaining rows to update_df and continue to the next file
                if not self.copyright_data.is_empty():
                    data_entry = compare(data_entry, self.copyright_data, select_cols)
                if not data_entry.is_empty():
                    data_entry = compare(data_entry, update_df, select_cols)

                    if not data_entry.is_empty():
                        info(
                            f"retrieved {data_entry.shape[0]} probable updated items from {file.path} ."
                        )
                        update_df = pl.concat(
                            [update_df, data_entry], how="diagonal_relaxed"
                        )

        if not update_df.is_empty():
            info(
                f"Sending {update_df.shape[0]} items from faculty sheets to the database for updating."
            )
            await update_copyright_items(update_df)
            info(f"Returning {len(material_ids)} material_ids from faculty sheets.")

            return (True, material_ids)
        info("No items to update based on faculty sheet contents.")
        return (False, material_ids)

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
