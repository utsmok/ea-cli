"""
This module defines the main `EasyAccessTool` class, which orchestrates the entire
data processing workflow for the Easy Access Sheet Toolkit. This includes:
- Reading data from various sources (Qlik raw export, weekly Excel sheets, overview Excel sheets).
- Synchronizing this data with a central SQLite database.
- Applying business logic and heuristics for data cleaning, validation, and conflict resolution.
- Generating output Excel sheets for different faculties and programmes.
- Creating summary overviews and export sheets for re-import into other systems.

The tool's behavior is controlled by settings loaded from a `settings.yaml` file via
the `EasyAccessSettings` and global `SETTINGS` objects.
"""

import asyncio
import logging
import traceback
from collections import defaultdict
from collections.abc import Callable  # Callable for self.functions
from datetime import datetime
from pathlib import Path
from typing import Any

import polars as pl

from easy_access.db.base import init as init_tortoise_orm
from easy_access.db.ingest import load_base_data, load_raw_copyright_data
from easy_access.db.models import Classification, WorkflowStatus
from easy_access.db.retrieve import retrieve_copyright_items
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
    # read_copyright_export, # Original global function, now using internal _read_raw_copyright_export
    store_complete_data,
)
from easy_access.utils import Directory, File

# Standard Python logging setup
# Basic config is often done at application entry point (e.g., run.py)
# If this module can be run standalone, this basicConfig is useful.
# Ensure it doesn't conflict if run.py also configures.
# logging.basicConfig(
#     level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
# )
logger = logging.getLogger(__name__)


class EasyAccessTool:
    """
    Orchestrates the Easy Access Sheet Toolkit workflow.

    This class manages settings, reads data from various sources, processes and
    synchronizes it with a database, and generates various output Excel sheets
    including faculty-specific sheets, programme sheets, and overview reports.

    Attributes:
        settings (EasyAccessSettings): Runtime settings for the current tool instance.
        functions (list[Callable[[], None]]): A list of methods to be run sequentially.
        dirs (dict[DirSetting, Directory]): Directory paths loaded from settings.
        disable_writes (bool): If True, file writing operations are skipped.
        copyright_data (pl.DataFrame): The main DataFrame holding current copyright items,
                                       typically loaded and processed from the database.
        faculties (list[str]): List of faculty abbreviations being processed.
        latest_file_date (str | None): Date string (YYYY-MM-DD) of the most recent raw
                                     copyright export file processed.
        mat_ids_on_disk (set[str]): Set of material IDs considered "on disk" or previously
                                    processed, used for filtering new items for sheets.
                                    (Note: its meaning evolved during refactoring).
        style_iter (int): Iterator for cycling through Excel table styles.
    """

    faculties: list[str]
    latest_file_date: str | None  # Can be None if no raw export is found initially
    copyright_data: pl.DataFrame
    mat_ids_on_disk: set[str]  # Material IDs are strings after Polars read usually

    def __init__(self, settings: EasyAccessSettings) -> None:
        """
        Initializes the EasyAccessTool with provided settings.

        Args:
            settings (EasyAccessSettings): Configuration settings for this run.
        """
        logger.info(f"Initializing EasyAccessTool with settings: {settings}")
        self.settings: EasyAccessSettings = settings
        self.functions: list[Callable[[], None]] = []
        self.dirs: dict[DirSetting, Directory] = (
            settings.dirs
        )  # From global SETTINGS via EasyAccessSettings.from_env
        self.disable_writes: bool = settings.disable_writes

        self.copyright_data = pl.DataFrame()  # Initialize as empty DataFrame
        self.mat_ids_on_disk = set()
        self.faculties = []
        self.latest_file_date = None

        # CLI-modifiable settings that are part of EasyAccessSettings
        self.only_changes: bool = settings.only_changes
        self.refresh_osiris_data: bool = settings.refresh_osiris_data
        self.enrich_with_osiris_data: bool = (
            settings.enrich_with_osiris_data
        )  # This is True by default in EasyAccessSettings
        self.only_retrieve_missing_osiris_data: bool = (
            settings.only_retrieve_missing_osiris_data
        )

        self.style_iter: int = 1  # Start styles from 1

        self.set_functions(settings.export)

    def set_functions(self, export_only_mode: bool) -> None:
        """
        Configures the sequence of operations (methods) to be run by the tool.

        If `export_only_mode` is True, only the `create_export_sheet` operation is added.
        Otherwise, the standard processing pipeline is set up.

        Args:
            export_only_mode (bool): If True, sets up only export sheet creation.
        """
        self.functions = []  # Clear any existing
        if export_only_mode:  # This mode is tricky as data might not be loaded.
            logger.info("Tool configured to run in 'export only' mode.")
            # This assumes self.copyright_data will be populated by some means before create_export_sheet runs,
            # or create_export_sheet can handle being called without prior data processing (which current impl cannot).
            # The `run.py` handles this by exiting after calling a global create_export_sheet.
            # If called via EasyAccessTool().run(), data must be loaded first.
            # For now, assuming if export_only_mode is True, process_raw_copyright_data should still run.
            # This 'export' flag in settings seems to control a final step, not an exclusive mode here.
            # The CLI 'export' option in run.py is a separate execution path.
            # Let's assume the 'export' in settings means "include export sheet in the standard run".
            self.functions.extend(
                [
                    self.process_raw_copyright_data,
                    self.create_overviews,
                    self.create_faculty_sheets,
                    self.create_all_items_sheet,
                    self.create_export_sheet,  # Added to the end of the standard pipeline
                ]
            )
        else:
            self.functions.extend(
                [
                    self.process_raw_copyright_data,
                    self.create_overviews,
                    self.create_faculty_sheets,
                    self.create_all_items_sheet,
                ]
            )
        logger.debug(f"Functions to run: {[f.__name__ for f in self.functions]}")

    def run(self) -> None:
        """
        Executes the configured sequence of data processing and sheet generation operations.
        """
        logger.info("EasyAccessTool run started.")
        for func_to_run in self.functions:
            logger.info(f"Running operation: {func_to_run.__name__}...")
            func_to_run()
            logger.info(f"Operation {func_to_run.__name__} completed.")
        logger.info("EasyAccessTool run finished.")

    def process_raw_copyright_data(self) -> None:
        """
        Orchestrates the main data ingestion and synchronization workflow.

        This method follows a specific order of operations:
        1. Initializes the database and loads base data if it's a fresh DB.
        2. Processes overview sheets: Reads them, identifies discrepancies with the DB,
           and updates the DB with user-editable fields from these sheets. Old overview sheets are then removed.
        3. Processes raw Qlik copyright export: Reads the latest export, ingests new items,
           and updates specific fields in the DB for existing items.
        4. Processes weekly sheets: Reads all weekly sheets, identifies items not in the DB (logged as errors),
           and then applies conflict resolution heuristics via `_resolve_weekly_sheet_conflicts_and_update_db`
           to update user-editable fields in the DB.
        5. Logs any errors found during sheet processing to faculty-specific error report files.
        6. Finalizes `self.copyright_data` by cleaning and validating the current DB state.
        7. Optionally refreshes Osiris data if `self.refresh_osiris_data` is True.
        8. Populates `self.faculties` and `self.mat_ids_on_disk` attributes.
        """
        logger.info("Starting centralized data ingestion and synchronization process.")
        all_errors_found: list[dict[str, Any]] = []

        # Initialize DB (Tortoise ORM)
        # `init_tortoise_orm` is async, so it needs to be run in an event loop.
        fresh_db_created: bool = asyncio.get_event_loop().run_until_complete(
            init_tortoise_orm()
        )
        if fresh_db_created:
            logger.info(
                "Fresh database initialized/schemas generated. Loading base data."
            )
            asyncio.get_event_loop().run_until_complete(
                load_base_data()
            )  # load_base_data is async

        # --- 1. Overview Sheets ---
        logger.info("Step 1: Processing overview sheets...")
        overview_data_df: pl.DataFrame = self._read_overview_sheets()
        db_data_df: pl.DataFrame = retrieve_copyright_items()

        if not overview_data_df.is_empty():
            overview_mat_ids = (
                overview_data_df.get_column("material_id").cast(pl.Utf8).unique()
            )
            db_mat_ids = db_data_df.get_column("material_id").cast(pl.Utf8).unique()

            missing_in_db_df = overview_data_df.filter(
                ~pl.col("material_id").is_in(db_mat_ids)
            )
            if not missing_in_db_df.is_empty():
                logger.warning(
                    f"Found {missing_in_db_df.height} items in overview sheets not present in the DB. Logging as errors."
                )
                for row_dict in missing_in_db_df.iter_rows(named=True):
                    all_errors_found.append(
                        {
                            "faculty": row_dict.get("faculty", "Unknown"),
                            "sheet_type": "overview",
                            "file_name": "N/A (aggregated)",
                            "material_id": row_dict["material_id"],
                            "error_type": "Item in overview not in DB",
                            "description": f"Mat ID {row_dict['material_id']} (Faculty: {row_dict.get('faculty', 'Unknown')}) in overview sheet, not in DB.",
                        }
                    )

            # Update DB from overview sheets for specific fields (user-editable, overview takes precedence)
            update_cols_from_overview = [
                "material_id",
                "workflow_status",
                "manual_classification",
                "remarks",
            ]
            overview_update_candidates = overview_data_df.filter(
                pl.col("material_id").is_in(db_mat_ids)
            ).select(
                [
                    col
                    for col in update_cols_from_overview
                    if col in overview_data_df.columns
                ]
            )
            if not overview_update_candidates.is_empty():
                logger.info(
                    f"Updating {overview_update_candidates.height} DB items from overview sheets."
                )
                asyncio.get_event_loop().run_until_complete(
                    update_copyright_items(
                        overview_update_candidates,
                        overwrite=True,
                        update_relations=False,
                    )
                    # Overwrite=True for these fields from overview. update_relations later.
                )
                db_data_df = retrieve_copyright_items()  # Refresh DB data
        else:
            logger.info("No overview sheet data to process.")
        self.remove_current_overviews()  # Backup/delete existing overview sheets

        # --- 2. Raw Copyright Data (Qlik export) ---
        logger.info("Step 2: Processing raw copyright data (Qlik export)...")
        raw_export_df: pl.DataFrame = (
            self._read_raw_copyright_export()
        )  # This sets self.latest_file_date

        if not raw_export_df.is_empty():
            asyncio.get_event_loop().run_until_complete(
                load_raw_copyright_data(raw_export_df)
            )
            logger.info(
                f"Processed {raw_export_df.height} items from raw copyright export into DB."
            )
            db_data_df = retrieve_copyright_items()  # Refresh DB data
        else:
            logger.warning("No raw copyright data found to process.")
            if (
                not self.latest_file_date
            ):  # Ensure latest_file_date is set for subsequent operations
                if (
                    not db_data_df.is_empty()
                    and "retrieved_from_copyright_on" in db_data_df.columns
                ):
                    # Get max date from DB if available
                    latest_db_date = db_data_df.select(
                        pl.col("retrieved_from_copyright_on").max().cast(pl.Utf8)
                    ).item()
                    if latest_db_date:
                        self.latest_file_date = latest_db_date
                if not self.latest_file_date:
                    self.latest_file_date = datetime.now().strftime("%Y-%m-%d")
                logger.info(f"Using date for operations: {self.latest_file_date}")

        # --- 3. Weekly Sheets ---
        logger.info("Step 3: Processing weekly sheets...")
        weekly_data_df: pl.DataFrame = self._read_weekly_sheets()

        if not weekly_data_df.is_empty():
            current_db_mat_ids_str = (
                db_data_df.get_column("material_id").cast(pl.Utf8).unique().to_list()
            )

            missing_in_db_weekly_df = weekly_data_df.filter(
                ~pl.col("material_id").is_in(current_db_mat_ids_str)
            )
            if not missing_in_db_weekly_df.is_empty():
                logger.warning(
                    f"Found {missing_in_db_weekly_df.height} items in weekly sheets not present in DB. Logging as errors."
                )
                for row_dict in missing_in_db_weekly_df.iter_rows(named=True):
                    all_errors_found.append(
                        {
                            "faculty": row_dict.get("faculty", "Unknown"),
                            "sheet_type": "weekly",
                            "file_name": "N/A (aggregated)",
                            "material_id": row_dict["material_id"],
                            "error_type": "Item in weekly sheet not in DB",
                            "description": f"Mat ID {row_dict['material_id']} (Faculty: {row_dict.get('faculty', 'Unknown')}) in weekly sheet, not in DB.",
                        }
                    )

            # Resolve conflicts and update DB from weekly sheets
            db_data_df = asyncio.get_event_loop().run_until_complete(
                self._resolve_weekly_sheet_conflicts_and_update_db(
                    db_data_df, weekly_data_df
                )
            )
        else:
            logger.info("No weekly sheet data to process.")

        # --- Error Reporting ---
        if all_errors_found:
            logger.info(
                f"Consolidating and writing {len(all_errors_found)} errors/inconsistencies found."
            )
            errors_report_df = pl.from_dicts(
                all_errors_found,
                schema={
                    "faculty": pl.Utf8,
                    "sheet_type": pl.Utf8,
                    "file_name": pl.Utf8,
                    "material_id": pl.Utf8,
                    "error_type": pl.Utf8,
                    "description": pl.Utf8,
                },
            )
            if not errors_report_df.is_empty():
                # Group errors by faculty and write them
                for faculty_val, group_df in errors_report_df.group_by("faculty"):
                    faculty_name_for_report = (
                        faculty_val[0]
                        if isinstance(faculty_val, tuple)
                        else str(faculty_val)
                    )
                    self._create_error_report(
                        group_df, faculty_name_for_report or "Unknown_Faculty"
                    )
        else:
            logger.info(
                "No errors or inconsistencies found in sheet data processing requiring reports."
            )

        # --- Finalize Data and Optional Osiris Update ---
        logger.info("Finalizing processed data.")
        self.copyright_data = self.clean_and_validate_df(db_data_df)

        if self.refresh_osiris_data and not self.copyright_data.is_empty():
            logger.info("Refreshing Osiris data as per settings...")
            asyncio.get_event_loop().run_until_complete(
                update_osiris_data(
                    self.copyright_data, self.only_retrieve_missing_osiris_data
                )
            )
            self.copyright_data = self.clean_and_validate_df(
                retrieve_copyright_items()
            )  # Re-fetch and validate
        elif self.copyright_data.is_empty():
            logger.warning(
                "Skipping Osiris data refresh as there is no copyright data after processing."
            )

        # Populate faculty list and material IDs for subsequent steps
        if not self.copyright_data.is_empty():
            self.faculties = (
                self.copyright_data.get_column("faculty")
                .unique()
                .drop_nulls()
                .sort()
                .to_list()
            )
            self.mat_ids_on_disk = set(
                self.copyright_data.get_column("material_id").drop_nulls().to_list()
            )

        if self.settings.faculty:  # Filter by specific faculty if requested
            logger.info(
                f"Filtering final data for selected faculty: {self.settings.faculty}"
            )
            self.faculties = [f for f in self.faculties if f == self.settings.faculty]
            if not self.copyright_data.is_empty():
                self.copyright_data = self.copyright_data.filter(
                    pl.col("faculty") == self.settings.faculty
                )

        logger.info(
            f"Centralized data processing complete. Final master DataFrame has {self.copyright_data.height} rows."
        )
        if not self.faculties:
            logger.warning("No faculties identified in the final dataset.")

    async def _resolve_weekly_sheet_conflicts_and_update_db(
        self, db_data: pl.DataFrame, weekly_data: pl.DataFrame
    ) -> pl.DataFrame:
        """
        Compares data from weekly sheets with current DB data for shared material IDs.
        Applies heuristics (defined in TODO C.3) to decide which source takes
        precedence for `workflow_status`, `manual_classification`, and `remarks`.
        Updates the database for items where changes are determined.

        Args:
            db_data (pl.DataFrame): Current data from the database.
            weekly_data (pl.DataFrame): Aggregated data from all weekly sheets.

        Returns:
            pl.DataFrame: The state of the database data after potential updates.
        """
        if weekly_data.is_empty():
            logger.info("No weekly data provided for conflict resolution.")
            return db_data
        if db_data.is_empty():
            logger.warning("DB data is empty; cannot resolve weekly sheet conflicts.")
            return db_data  # Or weekly_data, depending on desired behavior

        logger.info(
            f"Resolving conflicts for {weekly_data.height} weekly sheet items against {db_data.height} DB items."
        )

        # Define priority orders for conflict resolution (lower index = higher priority)
        workflow_priority_map: dict[str, int] = {
            str(status.value).lower(): idx
            for idx, status in enumerate(
                WorkflowStatus
            )  # Assuming higher value in Enum is higher priority
        }
        # Example: {"done": 0, "inprogress": 1, "todo": 2} if Done is highest
        # The current logic uses `>` so higher number = higher prio. Let's reverse for clarity or adjust logic.
        # Sticking to current logic: higher number = higher priority
        workflow_priority_map = {"todo": 0, "inprogress": 1, "done": 2}

        # Manual classification priority: list defines order, lower index = higher priority
        mc_priority_list_lower: list[str] = [
            str(c.value).lower() for c in Classification
        ]  # Get all from Enum
        # Example: ['open access', 'korte overname', ...]
        manual_classification_priority_map: dict[str, int] = {
            name: idx for idx, name in enumerate(mc_priority_list_lower)
        }

        update_payloads: list[dict[str, Any]] = []

        # Ensure material_id is string for joining, as it's validated to Utf8 in clean_and_validate_df
        weekly_relevant_cols = [
            "material_id",
            "workflow_status",
            "manual_classification",
            "remarks",
        ]
        # Select only relevant columns and ensure they exist
        weekly_subset_cols = [
            col for col in weekly_relevant_cols if col in weekly_data.columns
        ]
        weekly_subset_df = weekly_data.select(weekly_subset_cols)

        db_relevant_cols = [
            "material_id",
            "workflow_status",
            "manual_classification",
            "remarks",
        ]
        db_subset_cols = [col for col in db_relevant_cols if col in db_data.columns]
        db_subset_df = db_data.select(db_subset_cols)

        # Join dataframes on material_id to find common items
        # All material_id columns should be Utf8 due to clean_and_validate_df
        joined_df = weekly_subset_df.join(
            db_subset_df, on="material_id", how="inner", suffix="_db"
        )

        logger.info(
            f"Processing {joined_df.height} items common to weekly sheets and DB for conflict resolution."
        )

        for row_dict in joined_df.iter_rows(named=True):
            mat_id = row_dict["material_id"]
            current_update_payload: dict[str, Any] = {"material_id": mat_id}
            has_changed_in_payload: bool = False

            # --- Workflow Status Resolution ---
            ws_weekly_original = row_dict.get(
                "workflow_status"
            )  # Preserve original case for update
            ws_weekly_lower = str(ws_weekly_original or "").strip().lower()
            ws_db_original = row_dict.get("workflow_status_db")
            ws_db_lower = str(ws_db_original or "").strip().lower()

            if ws_weekly_lower:  # Only consider update if weekly has a value
                if not ws_db_lower:  # DB is empty, weekly is not
                    current_update_payload["workflow_status"] = ws_weekly_original
                    has_changed_in_payload = True
                elif ws_weekly_lower != ws_db_lower:
                    if workflow_priority_map.get(
                        ws_weekly_lower, -1
                    ) > workflow_priority_map.get(ws_db_lower, -1):
                        current_update_payload["workflow_status"] = ws_weekly_original
                        has_changed_in_payload = True

            # --- Manual Classification Resolution ---
            mc_weekly_original = row_dict.get("manual_classification")
            mc_weekly_lower = str(mc_weekly_original or "").strip().lower()
            mc_db_original = row_dict.get("manual_classification_db")
            mc_db_lower = str(mc_db_original or "").strip().lower()

            # Determine effective workflow status for MC logic (what it will be after this update round for WS)
            effective_ws_after_update_lower = (
                str(current_update_payload.get("workflow_status", ws_db_original) or "")
                .strip()
                .lower()
            )

            if mc_weekly_lower:  # Only consider update if weekly has a value
                if not mc_db_lower:  # DB is empty
                    current_update_payload["manual_classification"] = mc_weekly_original
                    has_changed_in_payload = True
                elif mc_weekly_lower != mc_db_lower:
                    # Apply MC priority if WS is same, OR if WS was just updated from weekly sheet
                    ws_is_effectively_same_priority = workflow_priority_map.get(
                        effective_ws_after_update_lower, -1
                    ) == workflow_priority_map.get(ws_weekly_lower, -1)
                    ws_was_changed_by_weekly = (
                        current_update_payload.get("workflow_status")
                        == ws_weekly_original
                    )

                    if ws_is_effectively_same_priority or ws_was_changed_by_weekly:
                        # Higher value in map means lower priority (index in list)
                        if manual_classification_priority_map.get(
                            mc_weekly_lower, float("inf")
                        ) < manual_classification_priority_map.get(
                            mc_db_lower, float("inf")
                        ):
                            current_update_payload["manual_classification"] = (
                                mc_weekly_original
                            )
                            has_changed_in_payload = True

            # --- Remarks Resolution ---
            remarks_weekly_original = str(row_dict.get("remarks") or "").strip()
            remarks_db_original = str(row_dict.get("remarks_db") or "").strip()

            # Effective WS and MC after potential updates for remarks logic
            final_ws_for_remarks_logic = (
                str(current_update_payload.get("workflow_status", ws_db_original) or "")
                .strip()
                .lower()
            )
            final_mc_for_remarks_logic = (
                str(
                    current_update_payload.get("manual_classification", mc_db_original)
                    or ""
                )
                .strip()
                .lower()
            )

            # Only update remarks if other key fields (WS, MC) are considered "aligned" with weekly values
            # AND weekly remark is non-empty and different from DB remark.
            if (
                final_ws_for_remarks_logic == ws_weekly_lower
                and final_mc_for_remarks_logic == mc_weekly_lower
                and remarks_weekly_original
                and remarks_weekly_original != remarks_db_original
            ):
                if not remarks_db_original:  # DB empty, take weekly
                    current_update_payload["remarks"] = remarks_weekly_original
                    has_changed_in_payload = True
                else:  # Both have remarks, and they are different. Merge.
                    # Heuristic: if weekly contains DB remark (case insensitive), take weekly. Else, append.
                    if remarks_db_original.lower() in remarks_weekly_original.lower():
                        current_update_payload["remarks"] = remarks_weekly_original
                    else:  # Append if no clear overlap or weekly is subset.
                        current_update_payload["remarks"] = (
                            f"{remarks_db_original}; {remarks_weekly_original}"
                        )
                    has_changed_in_payload = True
            elif remarks_weekly_original and not remarks_db_original:
                # If DB remarks is empty, and weekly has remarks, and other fields caused a change OR were already same.
                if has_changed_in_payload or (
                    final_ws_for_remarks_logic == ws_weekly_lower
                    and final_mc_for_remarks_logic == mc_weekly_lower
                ):
                    current_update_payload["remarks"] = remarks_weekly_original
                    has_changed_in_payload = (
                        True  # Ensure flag is set if this is the only change
                    )

            if has_changed_in_payload:
                update_payloads.append(current_update_payload)

        if update_payloads:
            logger.info(
                f"Identified {len(update_payloads)} items from weekly sheets for DB update based on heuristics."
            )
            update_df = pl.from_dicts(update_payloads)
            # update_copyright_items handles its own Tortoise init/close and logging.
            # `overwrite` should be False here to allow update_copyright_items to use its own field-level logic if needed,
            # or True if these specific fields determined by heuristics should always take precedence.
            # Given the heuristics are applied here, these are the "winning" values.
            await update_copyright_items(
                update_df, overwrite=True, update_relations=False
            )
            logger.info(
                f"DB updated with {len(update_payloads)} items from weekly sheets."
            )
            return retrieve_copyright_items()  # Return refreshed DB data
        else:
            logger.info(
                "No changes from weekly sheets required DB updates based on heuristics."
            )
            return db_data  # Return original DB data

    def _create_error_report(self, errors_df: pl.DataFrame, faculty: str) -> None:
        """
        Writes errors found during sheet processing to a specific error Excel sheet for the given faculty.

        Args:
            errors_df (pl.DataFrame): DataFrame containing error details.
            faculty (str): The faculty abbreviation for naming the error report.
        """
        if errors_df.is_empty():
            logger.info(f"No errors to report for faculty {faculty}.")
            return

        error_reports_base_dir = self.dirs.get(DirSetting.FACULTIES_DIR)
        if not error_reports_base_dir:
            logger.error(
                "Faculties directory not configured in settings. Cannot save error report."
            )
            return

        error_sheet_dir = Directory(
            error_reports_base_dir.full / faculty / "error_reports"
        )
        error_sheet_dir.mkdir(parents=True, exist_ok=True)

        report_date_str = (
            self.latest_file_date
            if self.latest_file_date
            else datetime.now().strftime("%Y-%m-%d")
        )
        error_filename = (
            f"error_report_{faculty.replace(' ', '_')}_{report_date_str}.xlsx"
        )
        error_filepath = error_sheet_dir.full / error_filename

        try:
            logger.info(
                f"Writing error report for faculty '{faculty}' to {error_filepath} with {errors_df.height} errors."
            )
            worksheet_name = f"Errors_{faculty}".replace(" ", "_").replace(":", "_")[
                :31
            ]  # Sanitize and shorten
            errors_df.write_excel(
                workbook=str(error_filepath), worksheet=worksheet_name
            )
        except Exception as e:
            logger.error(
                f"Failed to write error report for faculty {faculty} to {error_filepath}: {e}"
            )

    def _read_overview_sheets(self) -> pl.DataFrame:
        """
        Reads all 'total_overview_{faculty_name}.xlsx' files from faculty directories.
        It concatenates data from all found overview sheets into a single DataFrame.
        Assumes overview sheets have a structure compatible with data entry sheets.

        Returns:
            pl.DataFrame: A DataFrame containing all data read from overview sheets.
                          Returns an empty DataFrame if no overview sheets are found or readable.
        """
        all_overview_data_list: list[pl.DataFrame] = []
        # Determine which faculties to scan: those set for the tool run, or all available faculty dirs
        faculties_to_scan_list = self.faculties
        if not faculties_to_scan_list:
            faculties_root_dir = self.dirs.get(DirSetting.FACULTIES_DIR)
            if faculties_root_dir and faculties_root_dir.exists:
                try:
                    faculties_to_scan_list = [
                        d.name for d in faculties_root_dir.full.iterdir() if d.is_dir()
                    ]
                except OSError as e:  # Handles permission issues etc.
                    logger.warning(
                        f"Could not list faculty directories in {faculties_root_dir.full}: {e}"
                    )
                    return pl.DataFrame()
            else:
                logger.warning(
                    f"Faculties directory not configured or found: {faculties_root_dir.full if faculties_root_dir else 'N/A'}"
                )
                return pl.DataFrame()

        for faculty_name in faculties_to_scan_list:
            faculty_dir = Directory(
                self.dirs[DirSetting.FACULTIES_DIR].full / faculty_name
            )  # utils.Directory
            if not faculty_dir.exists:
                logger.debug(
                    f"No directory found for faculty {faculty_name} at {faculty_dir.full}"
                )
                continue

            # Pattern for overview files, excluding temp Excel files (e.g., starting with ~$)
            overview_files_list = [
                f
                for f in faculty_dir.files_r  # Recursive search
                if f.name.startswith(f"total_overview_{faculty_name}")
                and f.name.endswith(".xlsx")
                and "llm" not in f.name.lower()
                and not f.name.startswith("~$")
            ]

            if not overview_files_list:
                logger.info(
                    f"No overview sheet found for faculty {faculty_name} in {faculty_dir.full}"
                )
                continue

            # If multiple overview files match (e.g. due to backups not being cleared/moved), use the most recent one.
            # This assumes File object has a 'modified' or 'created' datetime attribute.
            file_to_read_obj: File = max(
                overview_files_list,
                key=lambda f: f.modified if f.exists else datetime.min,
            )
            if len(overview_files_list) > 1:
                logger.warning(
                    f"Multiple overview files found for faculty {faculty_name}, using most recent: {file_to_read_obj.name}."
                )

            try:
                # Try reading the configured data entry sheet name, then fallback to first sheet
                df_sheet = pl.read_excel(
                    file_to_read_obj.path,
                    sheet_name=SETTINGS.data_settings.data_entry_name,
                )
                if (
                    df_sheet.is_empty() and SETTINGS.data_settings.data_entry_name != 0
                ):  # Check if default sheet_id=0 was already tried
                    logger.debug(
                        f"Sheet '{SETTINGS.data_settings.data_entry_name}' in {file_to_read_obj.name} is empty/not found, trying first sheet."
                    )
                    df_sheet = pl.read_excel(file_to_read_obj.path, sheet_index=0)

                if df_sheet.is_empty():
                    logger.warning(
                        f"Overview sheet {file_to_read_obj.name} is empty after trying common sheets."
                    )
                    continue

                df_cleaned = self.clean_and_validate_df(df_sheet)
                # Ensure 'faculty' column is consistent with the folder it came from
                df_with_faculty = df_cleaned.with_columns(
                    pl.lit(faculty_name).alias("faculty")
                )

                all_overview_data_list.append(df_with_faculty)
                logger.info(
                    f"Read {df_with_faculty.height} rows from overview sheet: {file_to_read_obj.path}"
                )
            except Exception as e_read:
                logger.error(
                    f"Error reading overview sheet {file_to_read_obj.path}: {e_read}"
                )

        if not all_overview_data_list:
            logger.info(
                "No data found in any overview sheets across all scanned faculties."
            )
            return pl.DataFrame()

        return pl.concat(all_overview_data_list, how="diagonal_relaxed")

    def _read_raw_copyright_export(self) -> pl.DataFrame:
        """
        Reads data from the latest raw copyright export Excel file or a specified 'other_sheet'.
        It performs initial cleaning (column renaming, basic transformations) and sets
        `self.latest_file_date` based on the chosen file's creation date.

        Returns:
            pl.DataFrame: DataFrame containing data from the raw copyright export.
                          Returns an empty DataFrame if no file is found or on error.
        """
        file_to_read: File | None = None
        source_description: str = ""

        if self.settings.other_sheet:  # If an 'other_sheet' is specified in settings
            try:
                # Ensure it's a File object from utils.py
                other_sheet_path = (
                    self.settings.other_sheet
                )  # This should be a Path object from settings
                if isinstance(other_sheet_path, Path):
                    file_to_read = File(other_sheet_path)  # utils.File
                    if not file_to_read.exists:
                        logger.warning(
                            f"Specified 'other_sheet' not found: {file_to_read.path}. Will try default directory."
                        )
                        file_to_read = None  # Reset to allow fallback
                    else:
                        source_description = (
                            f"specified 'other_sheet': {file_to_read.path}"
                        )
                else:  # Should not happen if settings are parsed correctly
                    logger.warning(
                        f"Misconfigured 'other_sheet' (not a Path): {other_sheet_path}. Trying default."
                    )
            except (
                Exception
            ) as e_other:  # Catch errors related to File object creation or path issues
                logger.error(
                    f"Error accessing 'other_sheet' {self.settings.other_sheet}: {e_other}. Trying default."
                )

        if not file_to_read:  # Fallback to default raw copyright data directory
            raw_data_dir = self.dirs.get(DirSetting.RAW_COPYRIGHT_DATA)
            if not raw_data_dir or not raw_data_dir.exists:
                logger.error(
                    f"Raw copyright data directory not configured or found: {raw_data_dir.full if raw_data_dir else 'N/A'}"
                )
                return pl.DataFrame()

            excel_files = [
                f
                for f in raw_data_dir.files_r
                if f.extension in [".xlsx", ".xls"] and not f.name.startswith("~$")
            ]
            if not excel_files:
                logger.warning(
                    f"No Excel files found in raw copyright data directory: {raw_data_dir.full}"
                )
                return pl.DataFrame()

            try:
                file_to_read = max(
                    excel_files, key=lambda f: f.created
                )  # Get most recent
                source_description = (
                    f"latest file from {raw_data_dir.full}: {file_to_read.name}"
                )
            except Exception as e_max:  # Should not happen if excel_files is not empty
                logger.error(
                    f"Error finding latest raw copyright file in {raw_data_dir.full}: {e_max}"
                )
                return pl.DataFrame()

        if not file_to_read:  # Final check if a file was identified
            logger.error(
                "No raw copyright export file could be identified for processing."
            )
            return pl.DataFrame()

        try:
            logger.info(f"Reading raw copyright data from {source_description}")
            self.latest_file_date = file_to_read.created.strftime(
                "%Y-%m-%d"
            )  # Set instance attribute

            raw_df = pl.read_excel(file_to_read.path)

            # Standardize column names (lowercase, underscores)
            renamed_df = raw_df.rename(
                lambda c: str(c)
                .replace(" ", "_")
                .replace("#", "count_")
                .replace("*", "x")
                .lower()
            )

            # Add metadata columns and perform initial transformations
            transformed_df = renamed_df.with_columns(
                pl.lit(self.latest_file_date).alias("retrieved_from_copyright_on"),
                pl.lit(WorkflowStatus.ToDo.value).alias(
                    "workflow_status"
                ),  # Default workflow
                pl.col("last_change")
                .str.replace(r"^- ভারতবর্ষ$", None)
                .str.strip_chars()
                .str.strptime(pl.Date, "%Y-%m-%d", strict=False)
                .cast(pl.Utf8)
                .alias("last_change"),  # Keep as string
                pl.col("classification").str.to_lowercase().alias("classification"),
                faculty=pl.col("department").replace_strict(
                    SETTINGS.university_settings.department_mapping,
                    default="Unmapped",  # Use mapping from settings
                ),
            )

            # Filter out rows with missing material_id or unwanted filetypes
            initial_count = transformed_df.height
            filtered_df = transformed_df.filter(
                pl.col("material_id").is_not_null()
                & (pl.col("material_id").cast(pl.Utf8) != "")
                & (pl.col("material_id").cast(pl.Utf8) != "-")
            )

            if "filetype" in filtered_df.columns:
                relevant_filetypes = ["pdf", "ppt", "pptx", "doc", "docx", "-"]
                filtered_df = filtered_df.with_columns(
                    pl.col("filetype").cast(pl.Utf8).fill_null("-")
                )
                filtered_df = filtered_df.filter(
                    pl.col("filetype").str.to_lowercase().is_in(relevant_filetypes)
                )

            logger.info(
                f"Read {raw_df.height} rows initially from {source_description}. "
                f"After initial transformations: {initial_count} rows. After filtering: {filtered_df.height} rows."
            )

            return self.clean_and_validate_df(filtered_df)  # Final validation
        except Exception as e_proc:
            logger.error(
                f"Error reading or processing raw copyright export from {source_description}: {e_proc}"
            )
            logger.debug(traceback.format_exc())
            return pl.DataFrame()

    def _read_weekly_sheets(self) -> pl.DataFrame:
        """
        Reads all weekly data sheets from faculty directories.
        Excludes overview sheets and LLM classification files. Concatenates data
        from all found weekly sheets into a single DataFrame.

        Returns:
            pl.DataFrame: A DataFrame containing all data from weekly sheets.
                          Returns an empty DataFrame if no weekly sheets are found or readable.
        """
        all_weekly_data_list: list[pl.DataFrame] = []
        faculties_to_scan_list = self.faculties
        if not faculties_to_scan_list:
            faculties_root_dir = self.dirs.get(DirSetting.FACULTIES_DIR)
            if faculties_root_dir and faculties_root_dir.exists:
                try:
                    faculties_to_scan_list = [
                        d.name for d in faculties_root_dir.full.iterdir() if d.is_dir()
                    ]
                except OSError as e:
                    logger.warning(
                        f"Could not list faculty directories in {faculties_root_dir.full}: {e}"
                    )
                    return pl.DataFrame()
            else:
                logger.warning(
                    f"Faculties directory for weekly sheets not configured or found: {faculties_root_dir.full if faculties_root_dir else 'N/A'}"
                )
                return pl.DataFrame()

        for faculty_name in faculties_to_scan_list:
            faculty_dir = Directory(
                self.dirs[DirSetting.FACULTIES_DIR].full / faculty_name
            )
            if not faculty_dir.exists:
                logger.debug(
                    f"No directory for faculty {faculty_name} at {faculty_dir.full} for weekly sheets."
                )
                continue

            weekly_sheet_files = [
                f
                for f in faculty_dir.files_r
                if f.extension in [".xlsx", ".xls"]
                and not f.name.startswith("total_overview_")
                and "llm_classification" not in f.name.lower()
                and not f.name.startswith("~$")  # Exclude temp Excel files
            ]

            if not weekly_sheet_files:
                logger.info(
                    f"No weekly sheets found for faculty {faculty_name} in {faculty_dir.full}"
                )
                continue

            for file_to_read_obj in weekly_sheet_files:
                try:
                    df_sheet = pl.read_excel(
                        file_to_read_obj.path,
                        sheet_name=SETTINGS.data_settings.data_entry_name,
                    )
                    if (
                        df_sheet.is_empty()
                        and SETTINGS.data_settings.data_entry_name != 0
                    ):
                        df_sheet = pl.read_excel(file_to_read_obj.path, sheet_index=0)

                    if df_sheet.is_empty():
                        logger.warning(
                            f"Weekly sheet {file_to_read_obj.name} is empty."
                        )
                        continue

                    df_cleaned = self.clean_and_validate_df(df_sheet)
                    df_with_faculty = df_cleaned.with_columns(
                        pl.lit(faculty_name).alias("faculty")
                    )

                    all_weekly_data_list.append(df_with_faculty)
                    logger.info(
                        f"Read {df_with_faculty.height} rows from weekly sheet: {file_to_read_obj.path}"
                    )
                except Exception as e_read:
                    logger.error(
                        f"Error reading weekly sheet {file_to_read_obj.path}: {e_read}"
                    )

        if not all_weekly_data_list:
            logger.info(
                "No data found in any weekly sheets across all scanned faculties."
            )
            return pl.DataFrame()

        # Concatenate all weekly data. Duplicates by material_id might exist if item is in multiple weekly sheets.
        # Conflict resolution later will handle this against the DB state.
        combined_weekly_df = pl.concat(all_weekly_data_list, how="diagonal_relaxed")
        logger.info(
            f"Combined {combined_weekly_df.height} rows from all weekly sheets."
        )
        return combined_weekly_df

    def create_faculty_sheets(self) -> None:
        """
        Creates or updates faculty-specific Excel sheets with copyright data.
        Typically includes a "Complete Data" sheet and a "Data Entry" sheet.
        Also orchestrates creation of programme-specific sheets if mappings exist.
        File naming incorporates `self.latest_file_date`.
        """
        if self.disable_writes:
            logger.warning(
                "Writes are disabled by settings; skipping faculty sheet creation."
            )
            return

        if not self.latest_file_date:
            logger.warning(
                "`latest_file_date` not set. Attempting to determine fallback date for faculty sheets."
            )
            # Fallback logic for date, similar to process_raw_copyright_data
            if (
                not self.copyright_data.is_empty()
                and "retrieved_from_copyright_on" in self.copyright_data.columns
            ):
                latest_db_date = self.copyright_data.select(
                    pl.col("retrieved_from_copyright_on").max().cast(pl.Utf8)
                ).item()
                if latest_db_date:
                    self.latest_file_date = latest_db_date
            if not self.latest_file_date:
                self.latest_file_date = datetime.now().strftime("%Y-%m-%d")
            logger.info(f"Using date for faculty sheets: {self.latest_file_date}")

        if self.copyright_data.is_empty():
            logger.warning(
                "No data in `self.copyright_data`. Skipping faculty sheet creation."
            )
            return

        logger.info(
            f"Creating/updating faculty sheets for date: {self.latest_file_date}"
        )

        # `self.faculties` should be populated by `process_raw_copyright_data`
        faculties_to_process = self.faculties
        if not faculties_to_process:
            logger.warning(
                "No faculties set in EasyAccessTool. Deriving from current copyright_data."
            )
            if "faculty" in self.copyright_data.columns:
                faculties_to_process = (
                    self.copyright_data.get_column("faculty")
                    .unique()
                    .drop_nulls()
                    .sort()
                    .to_list()
                )
            if not faculties_to_process:
                logger.error(
                    "No faculties found in data. Cannot create faculty sheets."
                )
                return

        faculties_to_process.sort()  # Ensure consistent order

        for faculty_name in faculties_to_process:
            faculty_df: pl.DataFrame = self.copyright_data.filter(
                pl.col("faculty") == faculty_name
            )

            if faculty_df.is_empty():
                logger.info(
                    f"No items for faculty '{faculty_name}'. Skipping sheet creation."
                )
                continue

            # TODO: Clarify "only_changes" logic for faculty sheets.
            # If self.only_changes is True, this might involve comparing faculty_df
            # with an existing sheet or a snapshot to determine if changes occurred.
            # For now, it proceeds to write if data for faculty exists.

            # Create programme-specific sheets for this faculty
            if faculty_name in COURSE_MAPPING:  # COURSE_MAPPING from settings
                self.create_programme_sheets(faculty_name, input_data=faculty_df)

            faculty_sheet_dir = Directory(
                self.dirs[DirSetting.FACULTIES_DIR].full / faculty_name
            )

            # Sanitize faculty name for filename if needed, though usually abbreviations are clean
            faculty_file_name_part = (
                faculty_name
                if faculty_name and faculty_name.strip() != ""
                else "Unmapped_Faculty"
            )
            base_filename = f"{faculty_file_name_part}_{self.latest_file_date}.xlsx"

            # Handle existing files by versioning
            output_filepath = faculty_sheet_dir.full / base_filename
            version_counter = 1
            while output_filepath.exists():
                versioned_filename = f"{faculty_file_name_part}_{self.latest_file_date}_{version_counter}.xlsx"
                output_filepath = faculty_sheet_dir.full / versioned_filename
                version_counter += 1

            logger.info(
                f"Writing {faculty_df.height} items for faculty '{faculty_name}' to {output_filepath.name}"
            )
            store_complete_data(
                output_filepath, faculty_df
            )  # This creates 'Complete Data' sheet
            self.style_iter = finalize_sheet(
                File(str(output_filepath)), faculty_df, self.style_iter
            )  # Adds 'Data Entry' sheet

    def create_programme_sheets(
        self, faculty_abbr: str, input_data: pl.DataFrame
    ) -> None:
        """
        Creates or updates programme-specific Excel sheets for a given faculty.
        Data is filtered from `input_data` based on 'department' and mapped to
        programme groups using `COURSE_MAPPING` from settings.

        Args:
            faculty_abbr (str): Abbreviation of the faculty.
            input_data (pl.DataFrame): DataFrame containing data for the specified faculty.
                                       Expected to have a 'department' column for filtering.
        """
        if self.disable_writes:
            logger.warning(
                f"Writes disabled; skipping programme sheets for faculty '{faculty_abbr}'."
            )
            return

        if not self.latest_file_date:
            logger.error(
                "`latest_file_date` not set. Cannot create programme sheets with dated names."
            )
            return  # Or use a fallback date, but this indicates a flow issue

        if input_data.is_empty():
            logger.info(
                f"No input data for faculty '{faculty_abbr}' to create programme sheets."
            )
            return

        programme_output_base_dir = (
            self.dirs[DirSetting.FACULTIES_DIR].full / faculty_abbr / "per_programme"
        )
        programme_output_base_dir.mkdir(exist_ok=True, parents=True)

        faculty_course_mapping = COURSE_MAPPING.get(faculty_abbr, {})
        if not faculty_course_mapping:
            logger.info(
                f"No programme (course) mapping found for faculty '{faculty_abbr}'. Skipping programme sheets."
            )
            return

        logger.info(
            f"Creating programme sheets for faculty '{faculty_abbr}' for date {self.latest_file_date}."
        )

        # Group data by 'department' (which often contains course codes or specific programme names)
        # then map these groups to sheet names (programme groups) using faculty_course_mapping.

        # Check if 'department' column exists
        if "department" not in input_data.columns:
            logger.warning(
                f"'department' column missing in data for faculty '{faculty_abbr}'. Cannot create programme sheets."
            )
            return

        grouped_by_dept = input_data.group_by("department")

        final_data_for_programme_sheets: dict[str, pl.DataFrame] = defaultdict(list)  # type: ignore # For pl.concat

        for dept_identifier, group_df in grouped_by_dept:
            # dept_identifier could be a tuple if grouping by multiple columns, ensure it's string
            dept_str_identifier = str(
                dept_identifier[0]
                if isinstance(dept_identifier, tuple)
                else dept_identifier
            )

            target_sheet_group_name = faculty_course_mapping.get(dept_str_identifier)

            if not target_sheet_group_name:
                logger.debug(
                    f"No target sheet group defined for department ID '{dept_str_identifier}' in faculty '{faculty_abbr}'."
                )
                continue

            if group_df.is_empty():  # Should not happen if group_by yields non-empty
                continue

            logger.debug(
                f"Dept ID '{dept_str_identifier}' (Programme Group: {target_sheet_group_name}): {group_df.height} items."
            )
            final_data_for_programme_sheets[target_sheet_group_name].append(group_df)  # type: ignore

        for sheet_group_name, list_of_dfs in final_data_for_programme_sheets.items():
            if not list_of_dfs:
                continue
            df_to_write = pl.concat(list_of_dfs, how="diagonal_relaxed")

            if df_to_write.is_empty():
                continue

            base_filename = f"{sheet_group_name}_{self.latest_file_date}.xlsx"
            version_counter = 1
            output_filepath = programme_output_base_dir / base_filename
            while output_filepath.exists():
                versioned_filename = (
                    f"{sheet_group_name}_{self.latest_file_date}_{version_counter}.xlsx"
                )
                output_filepath = programme_output_base_dir / versioned_filename
                version_counter += 1

            store_complete_data(output_filepath, df_to_write)
            self.style_iter = finalize_sheet(
                File(str(output_filepath)), df_to_write, self.style_iter
            )
            logger.info(
                f"Created programme sheet '{output_filepath.name}' for faculty '{faculty_abbr}' with {df_to_write.height} items."
            )

    def create_all_items_sheet(self) -> None:
        """
        Creates a single Excel sheet containing all items currently in `self.copyright_data`.
        The filename includes `self.latest_file_date`.
        """
        if self.disable_writes:
            logger.warning("Writes disabled; skipping creation of 'all_items' sheet.")
            return

        if self.copyright_data.is_empty():
            logger.warning(
                "No data in `self.copyright_data` to export to 'all_items' sheet."
            )
            return

        export_date_str = (
            self.latest_file_date
            if self.latest_file_date
            else datetime.now().strftime("%Y-%m-%d")
        )

        all_items_output_dir = self.dirs.get(DirSetting.ALL_ITEMS_DIR)
        if not all_items_output_dir:
            logger.error(
                "ALL_ITEMS_DIR not configured in settings. Cannot create all_items sheet."
            )
            return
        all_items_output_dir.mkdir(exist_ok=True, parents=True)

        base_filename = f"all_items_{export_date_str}.xlsx"
        version_counter = 1
        output_filepath = all_items_output_dir.full / base_filename
        while output_filepath.exists():
            versioned_filename = f"all_items_{export_date_str}_{version_counter}.xlsx"
            output_filepath = all_items_output_dir.full / versioned_filename
            version_counter += 1

        store_complete_data(
            output_filepath, self.copyright_data
        )  # store_complete_data logs success
        logger.info(
            f"Created 'all_items' sheet: {output_filepath.name} with {self.copyright_data.height} items."
        )

    def clean_and_validate_df(self, df: pl.DataFrame) -> pl.DataFrame:
        """
        Performs cleaning and basic validation on a DataFrame.

        Steps:
        - Returns empty DataFrame if input is empty.
        - Casts all columns to UTF8 (string) type for consistency.
        - Specifically casts `material_id` to UTF8.
        - Replaces truncated URLs in 'url' and 'osiris_catalogue_url' columns.
        - Standardizes boolean-like values in 'is_duplicate' column to "TRUE"/"FALSE".

        Args:
            df (pl.DataFrame): The input DataFrame to clean.

        Returns:
            pl.DataFrame: The cleaned and validated DataFrame.
        """
        if df.is_empty():
            logger.debug("clean_and_validate_df received an empty DataFrame.")
            return df

        # Cast all to string first for robust subsequent operations
        try:
            df = df.with_columns([pl.all().cast(pl.Utf8, strict=False)])
        except Exception as e:  # Catch potential errors during casting all columns
            logger.warning(
                f"Could not cast all columns to Utf8 in clean_and_validate_df: {e}. Proceeding."
            )

        if "material_id" in df.columns:
            df = df.with_columns(
                pl.col("material_id").cast(pl.Utf8)
            )  # Ensure material_id is string

        url_columns_to_fix = ["url", "osiris_catalogue_url"]
        for col_name in url_columns_to_fix:
            if col_name in df.columns:
                try:
                    # Using str.replace_all for regex pattern if "..." is meant as literal three dots
                    # Polars' str.replace default is literal, so r"\.{3}" might not be needed if it's literal "..."
                    df = df.with_columns(
                        pl.col(col_name).str.replace_all(
                            r"\.\.\.",
                            "https://utwente.instructure.com/files",
                            literal=False,
                        )
                        # Assuming regex was intended by \. If literal "...", then literal=True and pattern "..."
                    )
                except Exception as e_url:
                    logger.warning(
                        f"Error replacing URLs in column '{col_name}': {e_url}"
                    )

        if "is_duplicate" in df.columns:
            try:
                df = df.with_columns(
                    pl.when(
                        pl.col("is_duplicate")
                        .str.to_lowercase()
                        .is_in(["true", "1", "yes"])
                    )  # Added "yes"
                    .then(pl.lit("TRUE", dtype=pl.Utf8))
                    .otherwise(pl.lit("FALSE", dtype=pl.Utf8))
                    .alias("is_duplicate")
                )
            except Exception as e_dup:
                logger.warning(f"Error standardizing 'is_duplicate' column: {e_dup}")

        logger.debug(f"DataFrame cleaned and validated. Shape after: {df.shape}")
        return df

    def remove_current_overviews(self) -> None:
        """
        Removes (or backs up) existing 'total_overview_*.xlsx' files from faculty directories.
        Behavior depends on `self.disable_writes` and `SETTINGS.backup_settings.backup_overviews`.
        """
        if self.disable_writes:
            logger.info(
                "Writes disabled; skipping removal/backup of current overview sheets."
            )
            return

        faculties_to_process = self.faculties
        if not faculties_to_process:
            logger.info(
                "No faculties specified for overview removal. Scanning all faculty directories."
            )
            faculties_root_dir = self.dirs.get(DirSetting.FACULTIES_DIR)
            if faculties_root_dir and faculties_root_dir.exists:
                try:
                    faculties_to_process = [
                        d.name for d in faculties_root_dir.full.iterdir() if d.is_dir()
                    ]
                except OSError as e:
                    logger.warning(
                        f"Could not list faculty directories for overview removal: {e}"
                    )
                    return
            else:
                logger.warning(
                    f"Faculties directory not configured or found for overview removal. Path: {faculties_root_dir.full if faculties_root_dir else 'N/A'}"
                )
                return

        logger.info(
            f"Starting removal/backup of existing overview sheets for faculties: {faculties_to_process}"
        )
        for faculty_name in faculties_to_process:
            overview_faculty_dir = Directory(
                self.dirs[DirSetting.FACULTIES_DIR].full / faculty_name
            )
            if not overview_faculty_dir.exists:
                logger.debug(
                    f"No directory for faculty {faculty_name}, skipping overview removal."
                )
                continue

            backup_target_dir = Directory(
                self.dirs[DirSetting.OVERVIEWS_BACKUP].full / faculty_name
            )
            if SETTINGS.backup_settings.backup_overviews:
                backup_target_dir.mkdir(parents=True, exist_ok=True)

            for file_obj in overview_faculty_dir.files_r:  # Recursive search
                if (
                    file_obj.name.startswith(f"total_overview_{faculty_name}")
                    and file_obj.name.endswith(".xlsx")
                    and "llm" not in file_obj.name.lower()
                    and not file_obj.name.startswith("~$")
                ):
                    if SETTINGS.backup_settings.backup_overviews:
                        try:
                            # Move to backup, handling potential name clashes by versioning
                            target_backup_path = backup_target_dir.full / file_obj.name
                            v_counter = 1
                            while (
                                target_backup_path.exists()
                            ):  # Check for existing file in backup
                                target_backup_path = (
                                    backup_target_dir.full
                                    / f"{file_obj.path.stem}_{v_counter}{file_obj.path.suffix}"
                                )
                                v_counter += 1
                            file_obj.move(target_backup_path)
                            logger.info(
                                f"Moved overview sheet '{file_obj.name}' to backup: {target_backup_path.name}"
                            )
                        except Exception as e_move:
                            logger.error(
                                f"Failed to move overview sheet '{file_obj.name}' to backup: {e_move}"
                            )
                    else:  # No backup, just delete
                        try:
                            file_obj.delete()
                            logger.info(
                                f"Deleted overview sheet '{file_obj.name}' (backup disabled)."
                            )
                        except Exception as e_del:
                            logger.error(
                                f"Failed to delete overview sheet '{file_obj.name}': {e_del}"
                            )

    def create_overviews(self) -> None:
        """
        Generates faculty and programme overview Excel sheets based on `self.copyright_data`.
        It first ensures old overview sheets are removed/backed up.
        """
        if self.copyright_data.is_empty():
            logger.warning(
                "No data in `self.copyright_data`. Cannot produce overviews."
            )
            return

        logger.info(
            "Ensuring old overview sheets are removed before creating new ones."
        )
        self.remove_current_overviews()

        faculty_data_map: dict[str, pl.DataFrame] = {}

        faculties_for_overview_creation = self.faculties
        if (  # noqa: SIM102
            not faculties_for_overview_creation
        ):  # Fallback if self.faculties isn't populated
            if "faculty" in self.copyright_data.columns:
                faculties_for_overview_creation = (
                    self.copyright_data.get_column("faculty")
                    .unique()
                    .drop_nulls()
                    .sort()
                    .to_list()
                )

        if not faculties_for_overview_creation:
            logger.warning(
                "No faculties identified from data. Cannot create overview sheets."
            )
            return

        for faculty_name in faculties_for_overview_creation:
            if not faculty_name or faculty_name == "Unmapped":
                logger.debug(
                    f"Skipping overview creation for faculty name: '{faculty_name}'"
                )
                continue

            faculty_specific_data = self.copyright_data.filter(
                pl.col("faculty") == faculty_name
            )
            if faculty_specific_data.is_empty():
                logger.info(
                    f"No data for faculty '{faculty_name}' in self.copyright_data. Skipping its overview sheet."
                )
                continue
            faculty_data_map[faculty_name] = faculty_specific_data

        if not faculty_data_map:
            logger.warning(
                "No data available for any valid faculty to create overview sheets."
            )
            return

        # Determine file_date_str for naming consistency
        file_date_for_naming = (
            self.latest_file_date
            if self.latest_file_date
            else datetime.now().strftime("%Y-%m-%d")
        )

        self.style_iter = create_faculty_overviews(
            faculty_data_map, self.style_iter, self.disable_writes, file_date_for_naming
        )
        logger.info("Faculty and programme overview sheet creation process finished.")

    def create_export_sheet(self) -> None:
        """
        Creates export Excel sheets for each relevant faculty based on `self.copyright_data`.
        If no specific faculties are set for the tool run, it attempts a global export.
        The data exported is typically filtered for items with workflow_status 'Done'.
        """
        faculties_for_export_processing = self.faculties
        if not faculties_for_export_processing and not self.copyright_data.is_empty():
            faculties_for_export_processing = (
                self.copyright_data.get_column("faculty")
                .unique()
                .drop_nulls()
                .sort()
                .to_list()
            )

        if not faculties_for_export_processing:
            logger.warning("No specific faculties identified for export.")
            if not self.copyright_data.is_empty():
                logger.info("Attempting a global export for all available data.")
                # Call the underlying sheet creation function with all data and a generic faculty_name
                create_export_sheet(
                    data=self.copyright_data, faculty_name="ALL_FACULTIES_EXPORT"
                )
            else:
                logger.warning(
                    "No data available in self.copyright_data for a global export."
                )
            return

        for faculty_name_str in faculties_for_export_processing:
            if (
                not faculty_name_str or faculty_name_str == "Unmapped"
            ):  # Skip invalid/unmapped faculty names
                logger.debug(
                    f"Skipping export sheet creation for invalid/unmapped faculty: '{faculty_name_str}'"
                )
                continue

            logger.info(f"Preparing export sheet for faculty: {faculty_name_str}")
            faculty_data_for_export = self.copyright_data.filter(
                pl.col("faculty") == faculty_name_str
            )

            if faculty_data_for_export.is_empty():
                logger.info(f"No data to export for faculty '{faculty_name_str}'.")
                continue

            # The create_export_sheet function from easy_access.sheets.sheet now handles its own logging
            create_export_sheet(
                data=faculty_data_for_export, faculty_name=faculty_name_str
            )
        logger.info(
            "Export sheet creation process finished for all specified faculties."
        )
