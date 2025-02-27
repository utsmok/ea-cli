import asyncio
import os
import polars as pl
from easy_access.db.base import init
from easy_access.db.ingest import load_base_data, load_raw_copyright_data
from easy_access.db.retrieve import retrieve_copyright_items, retrieve_full_data
from easy_access.utils import Directory, File, info, cool, warn, print
from easy_access.sheets.enrichment import  update_osiris_data
from easy_access.sheets.sheet import (
    finalize_sheet,
    store_complete_data,
    read_copyright_export,
    create_export_sheet
)
from easy_access.sheets.analysis import create_faculty_overviews
from easy_access.settings import (
    SETTINGS,
    DirSetting,
    EasyAccessSettings,
    Functions,
    COURSE_MAPPING
)
from easy_access.db.update import update_copyright_items

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

        self.settings = settings
        self.functions:list[callable] = []
        self.dirs = settings.dirs

        # Initialize data structures
        self.copyright_data = pl.DataFrame()
        self.mat_ids_on_disk = set()

        # Initialize other attributes
        self.only_changes = settings.only_changes
        self.disable_writes = settings.disable_writes
        self.refresh_osiris_data = settings.refresh_osiris_data
        self.enrich_with_osiris_data = settings.enrich_with_osiris_data
        self.only_retrieve_missing_osiris_data = settings.only_retrieve_missing_osiris_data
        self.style_iter = 2

        # Set functions to run
        self.set_functions(settings.functions)

    def set_functions(self, functions: Functions | None) -> None:
        """
        Sets the functions to run based on the input parameter 'functions'.
        stores it in self.settings as a list of functions to run.
        Returns None.
        """
        if functions is None:
            return

        # Common functions
        self.functions.extend([
            self.process_raw_copyright_data,
            self.create_overviews,
            self.create_faculty_sheets,
            self.create_all_items_sheet,
            ])

        # exclusive for 'Both'
        if functions == Functions.both:
            self.functions.extend([
            self.create_export_sheet
        ])

    def run(self) -> None:
        """
        Runs the functions as specified in the settings dict.
        """

        for func in self.functions:
            info(f'running {func.__name__}')
            func()

    def process_raw_copyright_data(self) -> None:
        """
        Reads in the latest copyright export (using read_copyright_export).

        """
        file = None
        try:
            if not isinstance(self.settings.other_sheet, File):
                file = File(self.settings.other_sheet)
        except Exception as e:
            pass

        self.latest_file_date, self.copyright_data = read_copyright_export(file)
        if self.copyright_data.is_empty():
            warn("No new Copyright data found to process! No new items will be added. Checking if there are other changes...")
        else:
            # make sure base data is loaded

            fresh_db = asyncio.get_event_loop().run_until_complete(init())
            if fresh_db:
                asyncio.get_event_loop().run_until_complete(load_base_data())
            # load new data into db
            asyncio.get_event_loop().run_until_complete(load_raw_copyright_data(self.copyright_data))

        # retrieve full data from db
        self.copyright_data = self.clean_and_validate_df(retrieve_copyright_items())

        # get osiris data for the new items (or refresh all depending on settings)
        if self.refresh_osiris_data:
            asyncio.get_event_loop().run_until_complete(update_osiris_data(self.copyright_data, self.only_retrieve_missing_osiris_data))

        # set faculty names
        self.faculties = (
            self.copyright_data.select(pl.col("faculty").unique()).to_series().sort().to_list()
        )

        # determine which material_ids are already on stored in the faculty sheets
        updated_items, mat_ids = asyncio.get_event_loop().run_until_complete(self.update_db_from_faculty_sheets())
        print(f'{len(mat_ids)} material_ids found in faculty sheets, {len(self.copyright_data)} items currently in copyright_data.')

        if mat_ids:
            self.mat_ids_on_disk = mat_ids
        if updated_items:
            # if items were updated, refresh the data for the final time
            self.copyright_data = retrieve_copyright_items()
            self.copyright_data = self.clean_and_validate_df(self.copyright_data)


        cool(f'process copyright export done. {self.copyright_data.shape[0]} rows in self.copyright_data.')

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

        def compare(primary: pl.DataFrame, other:pl.DataFrame, select_cols:list[str]) -> pl.DataFrame:
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
            select_cols = [col for col in select_cols if col in primary_selected and col in other_selected]

            if not select_cols:
                warn('No columns to compare between dataframes. Skipping comparison; returning primary dataframe.')
                return primary
            if len(select_cols) == 1:
                warn(f'Only one column to compare: {select_cols}. Skipping comparison; returning primary dataframe.')
                return primary
            if len(select_cols) != len(initial_select_cols):
                warn(f'Not all selected columns are present in both dataframes. Selecting only the common columns: {select_cols}')

            # Get rows in primary but not in other
            not_in_other = primary.join(
                other.select(select_cols),
                on="material_id",
                how="anti"
            )

            # Get matching rows to compare
            matching = primary.join(
                other.select(select_cols),
                on="material_id",
                how="inner",
                suffix="_other"
            )

            # Keep rows if:
            # 1. Any values are null in other, but filled in primary
            # 2. Values are different, primary value is non-null and non-empty
            cols_to_compare = [c for c in select_cols if c != 'material_id']

            conditions = []
            for col in cols_to_compare:
                other_col = f"{col}_other"
                # Keep if other is null but primary has value
                conditions.append(
                    (pl.col(other_col).is_null()) &
                    (pl.col(col).is_not_null()) &
                    (pl.col(col) != "") &
                    (pl.col(col) != "-")
                )
                # Keep if values are different and primary is not null/empty
                conditions.append(
                    (pl.col(col) != pl.col(other_col)) &
                    (pl.col(col).is_not_null()) &
                    (pl.col(col) != "") &
                    (pl.col(col) != "-")
                )

            different_vals = matching.filter(
                pl.any_horizontal(conditions)
            ).select(cols_in_primary)

            return pl.concat([not_in_other, different_vals], how="diagonal_relaxed")

        select_cols = [
            'material_id',
            'workflow_status',
            'remarks',
            'manual_classification',
        ]
        material_ids = set()
        update_df: pl.DataFrame = pl.DataFrame()
        for faculty in self.faculties:
            # get all .xlsx files except llm_classification files
            # TODO: decide if we want to include overview xlsx files here, or skip them (i.e. can users add data to the overview files, or should they stick to the weekly sheets?)
            fac_dir = Directory(self.dirs[DirSetting.FACULTIES_DIR].full / faculty)
            files = fac_dir.files_r
            files = [f for f in files if f.extension == ".xlsx" and 'llm_classification' not in f.name]
            if not files:
                warn(f'No files found for faculty {faculty}.')
                continue
            for file in files:
                # load data entry sheet for file and process
                try:
                    data_entry = pl.read_excel(file.path, sheet_name=SETTINGS.data_settings.data_entry_name)
                except Exception as e:
                    warn(f'Error reading {file.path}: {e}')
                    continue
                data_entry = self.clean_and_validate_df(data_entry)
                if data_entry.is_empty():
                    warn(f'No data found in {file.path}.')
                    continue
                material_ids.update(data_entry.select(pl.col('material_id')).to_series().unique().to_list())

                # compare primary df (data_entry) to other (self.copyright_data, update_df)
                # If no rows remaining: continue
                # Else, do the same comparison as above but now compare data_entry to update_df
                # finally concat any remaining rows to update_df and continue to the next file
                if not self.copyright_data.is_empty():
                    data_entry = compare(data_entry, self.copyright_data, select_cols)
                if not data_entry.is_empty():
                    data_entry = compare(data_entry, update_df, select_cols)

                    if not data_entry.is_empty():
                        info(f'retrieved {data_entry.shape[0]} probable updated items from {file.path} .')
                        update_df = pl.concat([update_df, data_entry], how="diagonal_relaxed")

        if not update_df.is_empty():
            info(f'Sending {update_df.shape[0]} items from faculty sheets to the database for updating.')
            await update_copyright_items(update_df)
            return (True, material_ids)
        info('No items to update based on faculty sheet contents.')
        return (False, material_ids)

    def create_faculty_sheets(self) -> None:
        """
        Splits the processed copyright data into one sheet per faculty
        and exports the result to excel sheets.
        """
        if self.disable_writes:
            warn(f'Writes are disabled. Skipping programme & faculty sheet creation.')
            return
        info(f"Exporting new items to faculty sheets for date {self.latest_file_date}")
        int_mat_ids = [int(x) for x in self.mat_ids_on_disk if x]

        filtered_data: pl.DataFrame = retrieve_full_data(excluded_material_ids=int_mat_ids)
        if filtered_data.is_empty() and self.only_changes:
            warn("No new items found to export to faculty sheets.")
            return
        if self.faculties:
            self.faculties.sort()
        for faculty in self.faculties:
            gap = " " * (15 - len(faculty))
            faculty_data: pl.DataFrame = filtered_data.filter(pl.col("faculty") == faculty)
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

    def create_programme_sheets(self, faculty: str, input_data:pl.DataFrame | None = None) -> None:
        """
        For a given faculty, split processed copyright data into one sheet per programme.
        Export to faculty_dir / per_programme / programme_name}_{date}.xlsx
        """

        if self.disable_writes:
            warn(f'Writes are disabled. Skipping programme & faculty sheet creation.')
            return

        programme_dir = Directory(
            self.dirs[DirSetting.FACULTIES_DIR].full / faculty / "per_programme"
        )
        course_to_sheet: dict[str, str] = COURSE_MAPPING[faculty]
        data: list[dict[str, pl.DataFrame]] = []
        if not isinstance(input_data, pl.DataFrame):
            input_data = self.copyright_data
        if input_data.is_empty():
            warn(f'No data for {faculty} -- skipping programme sheet creation.')
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
            filename = (
                programme_dir.full / f"{groupname}_{self.latest_file_date}.xlsx"
            )

            store_complete_data(filename, df)
            self.style_iter = finalize_sheet(File(str(filename)), df, self.style_iter)
            info(
                f"created programme sheet {groupname}_{self.latest_file_date}.xlsx"
            )

    def create_all_items_sheet(self) -> None:
        """
        Add all filtered items in the current Copyright data to a single sheet.
        """
        if self.disable_writes:
            warn(f'Writes are disabled. Skipping create_all_items_sheet.')
            return

        filtered_data = self.copyright_data.filter(~pl.col("material_id").is_in(self.mat_ids_on_disk))
        if filtered_data.is_empty() and self.only_changes:
            warn("No new items found to export to all items sheet.")
            return

        filename = f"all_items_{self.latest_file_date}.xlsx"
        i = 1
        while os.path.exists(self.dirs[DirSetting.ALL_ITEMS_DIR].full / filename):
            filename = f"all_items_{self.latest_file_date}_{i}.xlsx"
            i += 1
        store_complete_data(self.dirs[DirSetting.ALL_ITEMS_DIR].full / filename, filtered_data)
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
            if 'url' in df.columns:
                df = df.with_columns(pl.col('url').str.replace(r'\.{3}','https://utwente.instructure.com/files'))
            if 'osiris_catalogue_url' in df.columns:
                df = df.with_columns(pl.col('osiris_catalogue_url').str.replace(r'\.{3}','https://utwente.instructure.com/files'))
            if 'is_duplicate' in df.columns:
                df = df.with_columns(pl.col('is_duplicate').str.replace('0', 'FALSE').str.replace('1', 'TRUE'))
        return df

    def remove_current_overviews(self) -> None:
        if self.disable_writes:
            warn(f'Writes are disabled. Skipping remove_current_overviews.')
            return

        for faculty in self.faculties:
            overview_fac_dir = Directory(self.dirs[DirSetting.FACULTIES_DIR].full / faculty)
            movedir = Directory(self.dirs[DirSetting.OVERVIEWS_BACKUP].full / faculty)
            for file in overview_fac_dir.files_r:
                if "total_overview" in file.name and faculty in file.name:
                    if SETTINGS.backup_settings.backup_overviews and "llm" not in file.name:
                        file.move( movedir.full / file.name)
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
                warn('No faculties detected in current data. Cannot produce overviews.')
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
        self.style_iter = create_faculty_overviews(faculty_dict, self.style_iter, self.disable_writes)

    def create_export_sheet(self) -> None:
        #TODO: change to retrieve data from db first
        warn(f'this function needs updates to work properly, returning for now.')
        return
        material_ids: list[str] = create_export_sheet(data=self.import_sheet_data)
        if material_ids:
            info(f"Updating export status for {len(material_ids)} material ids.")
            warn(f'NOT YET IMPLEMENTED')
            item_data: pl.DataFrame = self.import_sheet_data.filter(pl.col(name='material_id').is_in(other=material_ids))

            # group by faculty
            item_data_per_faculty: dict[str, pl.DataFrame] = {faculty: item_data.filter(pl.col(name='faculty') == faculty) for faculty in item_data.select(pl.col(name='faculty')).unique().to_series().to_list()}
            for faculty, data in item_data_per_faculty.items():
                # for each file in the faculty dir, open it
                # read the data
                # if the material_id is in the data, add col 'exported_on' with the current date (YYYY-MM-DD)
                # save the file
                info(f'Updating export status for {faculty}')
                print(data)
                warn(f'NOT YET IMPLEMENTED')

            # once done, run 'update_export_sheets' to create/update the export sheets for each faculty
            # these sheets show all the exported item for that faculty
            # items in these sheets should be removed from the overview sheets of that faculty
