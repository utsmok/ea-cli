"""
Main orchestrator for the Easy Access tool.

This module defines the `EasyAccessTool` class, which is responsible for
coordinating the various data processing workflows, including:
- Ingesting raw copyright data.
- Synchronizing data between Excel sheets and the SQLite database.
- Generating weekly faculty sheets and overview reports.
- Enriching data with external sources like Osiris.
"""

import asyncio
import contextlib
import datetime
import os
from collections.abc import Callable
from pathlib import Path

import polars as pl
from loguru import logger

from easy_access.db.base import ensure_db_inited
from easy_access.db.ingest import load_base_data, load_raw_copyright_data
from easy_access.db.retrieve import retrieve_copyright_items, retrieve_full_data
from easy_access.db.update import update_copyright_items
from easy_access.settings import (
    DirSetting,
    EasyAccessSettings,
    Settings,
)
from easy_access.sheets.analysis import create_faculty_overviews
from easy_access.sheets.enrichment import update_osiris_data
from easy_access.sheets.sheet import (
    create_export_sheet,
    finalize_sheet,
    read_copyright_export,
    store_complete_data,
)
from easy_access.utils import Directory, File
from easy_access.utilities.file_exists import check_file_exists
from easy_access.api_keys import canvas as api_token

# Existing TODO block remains as it's a design/task list, not a module docstring.
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
    Orchestrates the Easy Access tool's data processing workflows.

    This class handles the main sequence of operations, such as ingesting
    copyright data, updating the database from various sheet sources,
    and generating output sheets for faculties and programs.
    """

    faculties: list[str] = []
    latest_file_date: str  # Date of the latest copyright export file processed
    settings: Settings
    ea_settings: EasyAccessSettings
    functions: list[Callable[[], None]]
    dirs: dict[DirSetting, Directory]
    disable_writes: bool
    copyright_data: pl.DataFrame
    mat_ids_on_disk: set[str]
    only_changes: bool
    refresh_osiris_data: bool
    enrich_with_osiris_data: bool
    only_retrieve_missing_osiris_data: bool
    style_iter: int

    def __init__(self, settings_obj: Settings, ea_settings: EasyAccessSettings) -> None:
        """
        Initializes the EasyAccessTool.

        Args:
            settings_obj: The main Settings object for the application.
            ea_settings: EasyAccessSettings object containing runtime/CLI settings.
        """
        self.settings = settings_obj
        self.ea_settings = ea_settings

        self.functions = []
        self.dirs = self.settings.dirs
        self.disable_writes = self.ea_settings.disable_writes
        self.copyright_data = pl.DataFrame()
        self.mat_ids_on_disk = set()

        self.only_changes = self.ea_settings.only_changes
        self.refresh_osiris_data = self.ea_settings.refresh_osiris_data
        self.enrich_with_osiris_data = (
            self.ea_settings.enrich_with_osiris_data
        )  # This might need to come from main settings or be resolved
        self.only_retrieve_missing_osiris_data = (
            self.ea_settings.only_retrieve_missing_osiris_data
        )
        self.style_iter = 2
        self.latest_file_date = ""  # Initialize to empty string

        self.set_functions(self.ea_settings.export)

    def set_functions(self, export: bool) -> None:
        """
        Sets the sequence of processing functions to be run.

        The function list is stored in `self.functions`.

        Args:
            export: If True, includes the export sheet creation function.
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
        Executes the configured sequence of processing functions.

        Functions are run in the order they appear in `self.functions`.
        """
        for func in self.functions:
            logger.info(f"running {func.__name__}")
            func()

    def process_raw_copyright_data(self) -> None:
        """
        Processes the raw copyright data export.

        Reads the latest copyright export file (or a specified override file),
        loads data into the database, retrieves the full dataset,
        optionally enriches with Osiris data, and identifies material IDs
        already present in faculty sheets.
        """
        input_file_override: File | None = None
        if self.ea_settings.other_sheet is not None:
            # self.ea_settings.other_sheet is guaranteed to be a Path here
            try:
                input_file_override = File(self.ea_settings.other_sheet)
            except Exception as e:
                logger.warning(
                    f"Failed to create File object from other_sheet path '{self.ea_settings.other_sheet}': {e}"
                )
                # Proceed with default behavior (None for input_file_override)

        self.latest_file_date, self.copyright_data = read_copyright_export(
            settings=self.settings, file=input_file_override
        )
        if self.copyright_data.is_empty():
            logger.warning(
                "No new Copyright data found to process! No new items will be added. Checking if there are other changes..."
            )
        else:
            fresh_db: bool | None = asyncio.get_event_loop().run_until_complete(
                ensure_db_inited(settings=self.settings)
            )
            if fresh_db:
                asyncio.get_event_loop().run_until_complete(
                    load_base_data(settings=self.settings)
                )
            asyncio.get_event_loop().run_until_complete(
                load_raw_copyright_data(
                    settings=self.settings, file=self.copyright_data
                )
            )

        self.copyright_data = self.clean_and_validate_df(
            retrieve_copyright_items(settings=self.settings)
        )

        if self.refresh_osiris_data:
            asyncio.get_event_loop().run_until_complete(
                update_osiris_data(
                    settings=self.settings,
                    df=self.copyright_data,
                    only_retrieve_missing=self.only_retrieve_missing_osiris_data,
                )
            )

        self.faculties = (
            self.copyright_data.select(pl.col("faculty").unique())
            .to_series()
            .sort()
            .to_list()
        )
        if self.ea_settings.faculty:
            logger.info(f"Selected single faculty: {self.ea_settings.faculty}")
            self.faculties = [self.ea_settings.faculty]

        updated_items, mat_ids = asyncio.get_event_loop().run_until_complete(
            self.update_db_from_faculty_sheets()
        )
        logger.info(
            f"{len(mat_ids)} material_ids found in faculty sheets, {len(self.copyright_data)} items currently in copyright_data."
        )

        asyncio.get_event_loop().run_until_complete(self.add_file_exists())

        if mat_ids:
            self.mat_ids_on_disk = mat_ids
        if updated_items:
            self.copyright_data = retrieve_copyright_items(settings=self.settings)
            self.copyright_data = self.clean_and_validate_df(self.copyright_data)

        logger.success(
            f"process copyright export done. {self.copyright_data.shape[0]} rows in self.copyright_data."
        )

    async def update_db_from_faculty_sheets(self) -> tuple[bool, set[str]]:
        """
        Updates the database from data found in faculty sheets.

        Retrieves data from all faculty sheets ('data entry' tab), compares items
        with the current `self.copyright_data` and `update_df` (accumulated changes),
        and sends items with detected changes to the database for updating.

        Args:
            None

        Returns:
            A tuple containing:
                - bool: True if items were selected for update, False otherwise.
                - set[str]: A set of all distinct material_ids present in all faculty sheets.
        """

        def compare(
            primary: pl.DataFrame, other: pl.DataFrame, select_cols: list[str]
        ) -> pl.DataFrame:
            """
            Compares two DataFrames based on 'material_id' and selected columns.

            Keeps rows from `primary` if:
            - The row is in `primary` but not in `other`.
            - Values in `select_cols` differ, and `primary` has a non-empty value where `other` is null.
            - Values in `select_cols` differ, and `primary` has a non-null, non-empty value.

            Args:
                primary: The primary DataFrame.
                other: The DataFrame to compare against.
                select_cols: List of column names to use for comparison (must include 'material_id').

            Returns:
                A DataFrame containing rows from `primary` that meet the keep criteria.
            """
            cols_in_primary = primary.columns
            # cols_in_other = other.columns # Not directly used after this
            primary_selected = [col for col in select_cols if col in cols_in_primary]
            other_selected = [
                col for col in select_cols if col in other.columns
            ]  # Corrected to use other.columns
            initial_select_cols = list(select_cols)  # Make a copy

            # now only select the cols that are in both dataframes
            select_cols = [
                col
                for col in select_cols
                if col in primary_selected and col in other_selected
            ]

            if (
                not select_cols or "material_id" not in select_cols
            ):  # Ensure material_id is present
                logger.warning(
                    "Not enough common columns (or 'material_id' missing) to compare. Skipping comparison; returning primary dataframe."
                )
                return primary
            if (
                len(select_cols) == 1 and "material_id" in select_cols
            ):  # Only material_id
                logger.warning(
                    "Only 'material_id' column to compare. Skipping detailed comparison; returning primary dataframe."
                )
                return primary  # Or handle as per logic, this implies no data columns to compare
            if len(select_cols) != len(initial_select_cols):
                logger.warning(
                    f"Not all originally selected columns are present in both dataframes. Using common columns: {select_cols}"
                )

            # Get rows in primary but not in other
            not_in_other: pl.DataFrame = primary.join(
                other.select("material_id"),
                on="material_id",
                how="anti",  # Simpler anti join
            )

            # Get matching rows to compare
            matching: pl.DataFrame = primary.join(
                other.select(select_cols),  # Use the filtered select_cols
                on="material_id",
                how="inner",
                suffix="_other",
            )

            if matching.is_empty():
                return not_in_other

            cols_to_compare = [c for c in select_cols if c != "material_id"]
            if not cols_to_compare:  # No data columns left to compare
                return not_in_other  # Or decide if matching rows with no diff should be dropped

            conditions: list[pl.Expr] = []
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

            if not conditions:  # Should not happen if cols_to_compare is not empty
                different_vals = pl.DataFrame(schema=primary.schema)
            else:
                different_vals = matching.filter(pl.any_horizontal(conditions)).select(
                    cols_in_primary  # Select original columns from primary
                )

            return pl.concat([not_in_other, different_vals], how="diagonal_relaxed")

        # Helper to quietly read Excel files with Polars.
        # Polars/openpyxl sometimes emits many "Could not determine dtype for column"
        # messages during inference. We suppress stdout/stderr during the read and
        # then let our normal cleaning (clean_and_validate_df) cast columns to Utf8.
        def _read_excel_quiet(path: str | Path, sheet_name: str) -> pl.DataFrame:
            try:
                # Redirect noisy output to devnull while reading
                with open(os.devnull, "w") as devnull:
                    with (
                        contextlib.redirect_stdout(devnull),
                        contextlib.redirect_stderr(devnull),
                    ):
                        df = pl.read_excel(path, sheet_name=sheet_name)
                return df
            except Exception:
                # If an error occurs, try once without suppression so we get a helpful
                # traceback/logging in normal operation; if that still fails, re-raise.
                try:
                    return pl.read_excel(path, sheet_name=sheet_name)
                except Exception as e:
                    logger.warning(f"Error reading {path} sheet {sheet_name}: {e}")
                    raise

        select_cols: list[str] = [
            "material_id",
            "workflow_status",
            "remarks",
            "manual_classification",
        ]
        material_ids_found: set[str] = set()
        update_df: pl.DataFrame = pl.DataFrame(
            schema={col: pl.Utf8 for col in select_cols}
        )  # Initialize with schema

        for faculty in self.faculties:
            fac_dir = Directory(self.dirs[DirSetting.FACULTIES_DIR].full / faculty)
            excel_files: list[File] = [
                f
                for f in fac_dir.files_r
                if f.extension == ".xlsx" and "llm_classification" not in f.name
            ]
            if not excel_files:
                logger.warning(f"No Excel files found for faculty {faculty}.")
                continue
            for file_obj in excel_files:
                try:
                    data_entry_df = _read_excel_quiet(
                        file_obj.path, self.settings.data_settings.data_entry_name
                    )
                except Exception as e:
                    logger.warning(f"Error reading {file_obj.path}: {e}")
                    continue
                # Also try to read the 'Complete data' sheet from the same file
                complete_df = None
                try:
                    complete_df = _read_excel_quiet(
                        file_obj.path, self.settings.data_settings.complete_data_name
                    )
                    # normalize column names to match expected format
                    complete_df = self.clean_and_validate_df(complete_df)
                except Exception:
                    complete_df = None

                data_entry_df = self.clean_and_validate_df(data_entry_df)
                if data_entry_df.is_empty():
                    logger.warning(
                        f"No data found in data entry sheet of {file_obj.path}."
                    )
                    continue

                current_file_mat_ids = (
                    data_entry_df.select(pl.col("material_id").cast(pl.Utf8))
                    .to_series()
                    .unique()
                    .to_list()
                )
                material_ids_found.update(current_file_mat_ids)

                # Ensure data_entry_df has all columns from select_cols for comparison
                for col_name in select_cols:
                    if col_name not in data_entry_df.columns:
                        data_entry_df = data_entry_df.with_columns(
                            pl.lit(None).alias(col_name).cast(pl.Utf8)
                        )
                data_entry_df = data_entry_df.select(
                    select_cols
                )  # Ensure correct column order and selection

                changes_vs_db: pl.DataFrame = pl.DataFrame(schema=update_df.schema)
                if not self.copyright_data.is_empty():
                    # Ensure self.copyright_data has all select_cols for comparison
                    temp_copyright_data = self.copyright_data.clone()
                    for col_name in select_cols:
                        if col_name not in temp_copyright_data.columns:
                            temp_copyright_data = temp_copyright_data.with_columns(
                                pl.lit(None).alias(col_name).cast(pl.Utf8)
                            )
                    changes_vs_db = compare(
                        data_entry_df,
                        temp_copyright_data.select(select_cols),
                        select_cols,
                    )

                changes_vs_update_df: pl.DataFrame = pl.DataFrame(
                    schema=update_df.schema
                )
                if not changes_vs_db.is_empty():
                    if not update_df.is_empty():
                        changes_vs_update_df = compare(
                            changes_vs_db, update_df, select_cols
                        )
                    else:
                        changes_vs_update_df = changes_vs_db

                if not changes_vs_update_df.is_empty():
                    logger.info(
                        f"Retrieved {changes_vs_update_df.shape[0]} probable updated items from {file_obj.path}."
                    )
                    # If some of these material_ids are NEW (not in current copyright_data)
                    # and the workbook contains a 'Complete data' sheet, prefer the full
                    # row from that sheet so new CopyrightItem creation has required fields.
                    if complete_df is not None and not self.copyright_data.is_empty():
                        existing_mat_ids = set(
                            str(x)
                            for x in self.copyright_data.select(pl.col("material_id"))
                            .to_series()
                            .to_list()
                        )
                        replacements = []
                        for row in changes_vs_update_df.to_dicts():
                            mid = str(row.get("material_id"))
                            if mid not in existing_mat_ids:
                                try:
                                    full_rows = complete_df.filter(
                                        pl.col("material_id").cast(pl.Utf8) == mid
                                    )
                                    if not full_rows.is_empty():
                                        # use the last matching full row (if duplicates)
                                        replacements.append(
                                            full_rows.tail(1).to_dicts()[0]
                                        )
                                        continue
                                except Exception:
                                    pass
                            replacements.append(row)
                        changes_vs_update_df = pl.DataFrame(replacements)

                    update_df = pl.concat(
                        [update_df, changes_vs_update_df], how="diagonal_relaxed"
                    ).unique(subset=["material_id"], keep="last", maintain_order=True)

        if not update_df.is_empty():
            logger.info(
                f"Sending {update_df.shape[0]} items from faculty sheets to the database for updating."
            )
            await update_copyright_items(
                settings=self.settings,
                data=update_df,
            )
            logger.info(
                f"Returning {len(material_ids_found)} material_ids from faculty sheets."
            )
            return True, material_ids_found

        logger.info("No items to update based on faculty sheet contents.")
        return False, material_ids_found

    def create_faculty_sheets(self) -> None:
        """
        Creates individual Excel sheets for each faculty.

        Splits the processed copyright data by faculty and exports new items
        (not already on disk if `only_changes` is True) to separate Excel files.
        Also triggers programme sheet creation if applicable for the faculty.
        """
        if self.disable_writes:
            logger.warning(
                "Writes are disabled. Skipping programme & faculty sheet creation."
            )
            return

        if not hasattr(self, "latest_file_date") or not self.latest_file_date:
            logger.warning(
                "`latest_file_date` not set. Run `process_raw_copyright_data` first. Skipping faculty sheet creation."
            )
            return

        logger.info(
            f"Exporting new items to faculty sheets for date {self.latest_file_date}"
        )

        int_mat_ids: list[int] = []
        if self.mat_ids_on_disk:
            int_mat_ids = [int(x) for x in self.mat_ids_on_disk if x.isdigit()]

        filtered_data: pl.DataFrame = retrieve_full_data(
            excluded_material_ids=int_mat_ids, settings=self.settings
        )
        if filtered_data.is_empty() and self.only_changes:
            logger.warning("No new items found to export to faculty sheets.")
            return

        sorted_faculties = sorted(self.faculties) if self.faculties else []
        for faculty in sorted_faculties:
            gap = " " * (15 - len(faculty))
            faculty_data: pl.DataFrame = filtered_data.filter(
                pl.col("faculty") == faculty
            )
            if faculty_data.is_empty():
                logger.warning(
                    f"{faculty}:{gap}{faculty_data.shape[0]} (no new items, skipping)"
                )
                continue

            if faculty in self.settings.university_settings.course_mapping:
                self.create_programme_sheets(faculty, input_data=faculty_data)

            faculty_dir = Directory(self.dirs[DirSetting.FACULTIES_DIR].full / faculty)
            current_faculty_name = faculty if faculty else "no_faculty_found"
            filename_base = f"{current_faculty_name}_{self.latest_file_date}"
            output_file_path = self._get_unique_filepath(faculty_dir, filename_base)
            logger.info(
                f"{current_faculty_name}:{gap}{faculty_data.shape[0]} -> {output_file_path.name}"
            )
            store_complete_data(
                settings=self.settings,
                file=output_file_path,
                data=faculty_data,
            )
            style_iter_result: int | None = finalize_sheet(
                settings=self.settings,
                file=File(str(output_file_path)),
                data=faculty_data,
                style_iter=self.style_iter,
            )
            if style_iter_result is not None:
                self.style_iter = style_iter_result

    def create_programme_sheets(
        self, faculty: str, input_data: pl.DataFrame | None = None
    ) -> None:
        """
        Creates individual Excel sheets for each programme within a faculty.

        Splits the provided `input_data` (or `self.copyright_data` if None)
        by programme based on course mappings in settings. Exports data for
        each programme to a separate Excel file.

        Args:
            faculty: The faculty for which to create programme sheets.
            input_data: The DataFrame to process. If None, uses `self.copyright_data`.
        """
        if self.disable_writes:
            logger.warning("Writes are disabled. Skipping programme sheet creation.")
            return

        if not hasattr(self, "latest_file_date") or not self.latest_file_date:
            logger.warning(
                "`latest_file_date` not set. Run `process_raw_copyright_data` first. Skipping programme sheet creation."
            )
            return

        programme_dir = Directory(
            self.dirs[DirSetting.FACULTIES_DIR].full / faculty / "per_programme"
        )
        course_to_sheet_mapping: dict[str, str] | None = (
            self.settings.university_settings.course_mapping.get(faculty)
        )
        if not course_to_sheet_mapping:
            logger.warning(
                f"No course mapping found in settings for faculty {faculty}. Skipping programme sheet creation."
            )
            return

        current_data: pl.DataFrame
        current_data = input_data if input_data is not None else self.copyright_data

        if current_data.is_empty():
            logger.warning(
                f"No data provided for {faculty} -- skipping programme sheet creation."
            )
            return

        logger.info(f"Creating programme sheets for {faculty}")

        # data_for_groups stores tuples of (group_name_str, course_data_df)
        data_for_groups: list[tuple[str, pl.DataFrame]] = []
        for course, group_name_str in course_to_sheet_mapping.items():
            if "department" not in current_data.columns:
                logger.warning(
                    f"'department' column not found in data for faculty {faculty}. Cannot create programme sheets."
                )
                return  # Or continue to next course if appropriate

            course_data = current_data.filter(pl.col("department") == course)
            gap = " " * (40 - len(course))
            if course_data.is_empty():
                logger.warning(
                    f"{course}:{gap}{course_data.shape[0]} (no new items, skipping)"
                )
                continue

            logger.info(f"Retrieved programme sheet data for {course}")
            logger.info(f"{course}:{gap}{course_data.shape[0]}")
            data_for_groups.append((group_name_str, course_data))

        final_grouped_data: dict[str, pl.DataFrame] = {}
        for group_name, df_data in data_for_groups:
            if group_name in final_grouped_data:
                final_grouped_data[group_name] = pl.concat(
                    [final_grouped_data[group_name], df_data]
                )
            else:
                final_grouped_data[group_name] = df_data

        for group_name_str, df_for_group in final_grouped_data.items():
            filename_base = f"{group_name_str}_{self.latest_file_date}"
            output_file_path = self._get_unique_filepath(programme_dir, filename_base)
            store_complete_data(
                settings=self.settings, file=output_file_path, data=df_for_group
            )
            style_iter_result: int | None = finalize_sheet(
                settings=self.settings,
                file=File(str(output_file_path)),
                data=df_for_group,
                style_iter=self.style_iter,
            )
            if style_iter_result is not None:
                self.style_iter = style_iter_result
            logger.info(f"Created programme sheet {output_file_path.name}")

    def create_all_items_sheet(self) -> None:
        """
        Creates a single Excel sheet containing all new copyright items.

        Filters `self.copyright_data` to include only items not already
        present on disk (if `only_changes` is True) and exports them.
        """
        if self.disable_writes:
            logger.warning("Writes are disabled. Skipping create_all_items_sheet.")
            return

        if not hasattr(self, "latest_file_date") or not self.latest_file_date:
            logger.warning(
                "`latest_file_date` not set. Run `process_raw_copyright_data` first. Skipping all_items sheet creation."
            )
            return

        filtered_data: pl.DataFrame
        if self.only_changes:
            valid_mat_ids_on_disk = {
                item for item in self.mat_ids_on_disk if item
            }  # Filter out None or empty strings
            if "material_id" not in self.copyright_data.columns:
                logger.warning(
                    "'material_id' column not in copyright_data. Cannot filter for all_items_sheet."
                )
                filtered_data = self.copyright_data.clone()  # Or handle error
            else:
                filtered_data = self.copyright_data.filter(
                    ~pl.col("material_id")
                    .cast(pl.Utf8)
                    .is_in(list(valid_mat_ids_on_disk))
                )
        else:
            filtered_data = self.copyright_data.clone()

        if filtered_data.is_empty():  # Check after potential filtering
            logger.warning("No new items found to export to all items sheet.")
            return

        filename_base = f"all_items_{self.latest_file_date}"
        output_dir = Directory(
            self.dirs[DirSetting.ALL_ITEMS_DIR].full
        )  # Ensure Directory object
        output_file_path = self._get_unique_filepath(output_dir, filename_base)
        store_complete_data(
            settings=self.settings,
            file=output_file_path,
            data=filtered_data,
        )
        logger.info(f"Created sheet: {output_file_path}")

    def clean_and_validate_df(self, df: pl.DataFrame) -> pl.DataFrame:
        """
        Cleans and validates a DataFrame.

        Current implementation:
        - Casts all non-Utf8 columns to Utf8 (string).
        - Replaces URL truncation markers with a default base URL.
        - Converts 'is_duplicate' column from "0"/"1" to "FALSE"/"TRUE".

        Args:
            df: The input DataFrame.

        Returns:
            The cleaned and validated DataFrame.

        TODO:
            Implement more comprehensive validation and cleaning logic.
        """
        if df.is_empty():
            return df

        df_cleaned = df.clone()

        # set all columns to type str (Utf8 in Polars)
        for col_name in df_cleaned.columns:  # Iterate over column names
            if df_cleaned[col_name].dtype != pl.Utf8:
                df_cleaned = df_cleaned.with_columns(
                    pl.col(col_name).cast(pl.Utf8, strict=False)
                )

        marker = self.settings.data_settings.url_truncation_marker
        base_url = self.settings.data_settings.url_default_base

        if "url" in df_cleaned.columns:
            df_cleaned = df_cleaned.with_columns(
                pl.col("url").str.replace_all(
                    marker, base_url
                )  # Use replace_all for global replacement
            )
        if "osiris_catalogue_url" in df_cleaned.columns:
            df_cleaned = df_cleaned.with_columns(
                pl.col("osiris_catalogue_url").str.replace_all(
                    marker, base_url
                )  # Use replace_all
            )
        if "is_duplicate" in df_cleaned.columns:
            df_cleaned = df_cleaned.with_columns(
                pl.when(pl.col("is_duplicate") == "1")
                .then(pl.lit("TRUE"))
                .when(pl.col("is_duplicate") == "0")
                .then(pl.lit("FALSE"))
                .otherwise(pl.col("is_duplicate"))  # Keep original if not "0" or "1"
                .alias("is_duplicate")
            )
        return df_cleaned

    def remove_current_overviews(self) -> None:
        """
        Removes or backs up existing overview sheets for all faculties.

        Iterates through faculty directories, identifies 'total_overview' files,
        and either moves them to a backup location (if configured) or deletes them.
        """
        if self.disable_writes:
            logger.warning("Writes are disabled. Skipping remove_current_overviews.")
            return

        for faculty in self.faculties:
            if not faculty:  # Skip if faculty name is empty
                continue
            overview_fac_dir = Directory(
                self.dirs[DirSetting.FACULTIES_DIR].full / faculty
            )
            if not overview_fac_dir.exists:
                logger.warning(
                    f"Faculty directory not found: {overview_fac_dir.full}. Skipping overview removal for {faculty}."
                )
                continue

            movedir = Directory(self.dirs[DirSetting.OVERVIEWS_BACKUP].full / faculty)

            files_to_process: list[File] = []
            try:
                files_to_process = (
                    overview_fac_dir.files_r
                )  # Can raise if dir doesn't exist, handled above
            except (
                Exception
            ) as e:  # Catch any other unexpected errors during file listing
                logger.warning(f"Error listing files in {overview_fac_dir.full}: {e}")
                continue

            for file_obj in files_to_process:
                if "total_overview" in file_obj.name and faculty in file_obj.name:
                    if (
                        self.settings.backup_settings.backup_overviews
                        and "llm"
                        not in file_obj.name  # Do not backup llm specific overviews by default
                    ):
                        try:
                            if not movedir.exists:
                                movedir.create()
                            file_obj.move(movedir.full / file_obj.name)
                            logger.info(
                                f"Moved overview: {file_obj.name} to {movedir.full}"
                            )
                        except Exception as e:
                            logger.warning(
                                f"Could not move file {file_obj.name} to backup: {e}"
                            )
                    else:
                        try:
                            file_obj.delete()
                            logger.info(f"Deleted overview: {file_obj.name}")
                        except Exception as e:
                            logger.warning(
                                f"Could not delete file {file_obj.name}: {e}"
                            )

    def create_overviews(self) -> None:
        """
        Generates overview sheets for each faculty.

        Retrieves full data for each faculty, removes/backs up existing
        overviews, and then calls `create_faculty_overviews` to generate
        new overview sheets.
        """
        faculty_dict: dict[str, pl.DataFrame] = {}
        if not self.faculties:
            # Attempt to populate faculties if empty
            logger.warning(
                "Faculties list is empty. Attempting to process raw copyright data to populate it."
            )
            self.process_raw_copyright_data()  # This will set self.faculties
            if not self.faculties:
                logger.warning(
                    "No faculties detected after processing data. Cannot produce overviews."
                )
                return

        self.remove_current_overviews()  # Remove or backup existing overviews first

        for faculty in self.faculties:
            if not faculty or faculty == "Unmapped":  # Skip empty or "Unmapped"
                continue
            data = retrieve_full_data(
                selected_faculties=faculty, settings=self.settings
            )
            if data.is_empty():
                logger.warning(
                    f"No data retrieved for faculty '{faculty}'. Skipping overview creation for this faculty."
                )
                continue
            faculty_dict[faculty] = data

        if not faculty_dict:
            logger.warning(
                "No data collected for any faculty. Skipping faculty overview creation."
            )
            return

        style_iter_result: int | None = create_faculty_overviews(
            settings=self.settings,
            faculty_data=faculty_dict,
            style_iter=self.style_iter,
            disable_writes=self.disable_writes,
        )
        if style_iter_result is not None:
            self.style_iter = style_iter_result

    def create_export_sheet(self) -> None:
        """
        Creates export sheets for each faculty.

        Iterates through configured faculties and calls `create_export_sheet`
        from the `sheets.sheet` module for each one.
        """
        if not self.faculties:
            logger.warning(
                "No faculties configured or detected. Skipping export sheet creation."
            )
            return

        for faculty in self.faculties:
            if not faculty or faculty == "Unmapped":  # Skip empty or "Unmapped"
                continue
            logger.info(f"Creating export sheet for {faculty}")
            create_export_sheet(settings=self.settings, faculty=faculty)

    async def add_file_exists(self, refresh_all: bool = False) -> None:
        """
        Adds the 'file_exists' field to the copyright items in the database.
        """
        all_items = retrieve_copyright_items(
            settings=self.settings, additional_cols=["file_exists", "last_canvas_check"]
        )  # get fresh data from DB
        print(f"number of items total: {all_items.shape[0]}")
        if refresh_all:
            item_selection = pl.DataFrame()
            for_check = all_items
        else:
            item_selection = all_items.filter(pl.col("file_exists") == "")
            print(
                f"Remaining items after filtering out rows with file_exists values: {item_selection.shape[0]} ({-all_items.shape[0] + item_selection.shape[0]})"
            )
            # for each item in 'all_items', if 'status' == 'Deleted',
            # set 'file_exists' to False, and `last_canvas_check` to now
            item_selection = item_selection.with_columns(
                pl.when(pl.col("status") == "Deleted")
                .then(pl.lit(False))
                .otherwise(pl.col("file_exists"))
                .alias("file_exists"),
                pl.when(pl.col("status") == "Deleted")
                .then(pl.lit(datetime.datetime.now()))
                .otherwise(pl.col("last_canvas_check"))
                .alias("last_canvas_check"),
            )
            # pop rows with empty file_exists to for_check
            for_check = item_selection.filter(pl.col("file_exists") == "")
            item_selection = item_selection.filter(pl.col("file_exists") != "")
        if for_check.is_empty():
            logger.info("No items need file existence checking.")
        else:
            with_file_exists = await check_file_exists(api_token, for_check)
            print(with_file_exists.head(10))

            # append with_file_exists to item_selection
            item_selection = item_selection.vstack(with_file_exists)

            print(f"now updating db")
            await update_copyright_items(settings=self.settings, data=item_selection)

    def _get_unique_filepath(
        self, directory: Directory, filename_base: str, extension: str = ".xlsx"
    ) -> Path:
        """
        Generates a unique filepath in the given directory.

        If a file with filename_base + extension exists, it appends _1, _2, etc.
        until a unique name is found.

        Args:
            directory: The directory to place the file in.
            filename_base: The base name for the file (without counter or extension).
            extension: The file extension (defaults to .xlsx).

        Returns:
            A Path object for the unique file.
        """
        if not extension.startswith("."):
            extension = f".{extension}"

        filename = f"{filename_base}{extension}"
        output_file_path = directory.full / filename
        i = 1
        # Ensure directory exists before checking for file existence
        if not directory.exists:
            try:
                directory.create()
                logger.info(f"Created directory: {directory.full}")
            except Exception as e:
                logger.warning(f"Could not create directory {directory.full}: {e}")
                # Fallback: attempt to use current directory or raise error?
                # For now, let it proceed, Path.exists() will handle non-existent parent dirs.

        while output_file_path.exists():
            filename = f"{filename_base}_{i}{extension}"
            output_file_path = directory.full / filename
            i += 1
        return output_file_path
