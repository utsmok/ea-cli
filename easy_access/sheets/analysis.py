"""
This module focuses on creating analytical summary sheets, specifically
faculty-level and programme-level overviews from the processed copyright data.
It includes functions to aggregate data, calculate derived fields (like
potential fines or infringement status), and then store these overviews
as styled Excel sheets. It also contains a utility to update the database
with data aggregated during overview generation, though this might be
better integrated into the main data processing flow in the future.
"""

import asyncio
import logging
from collections import defaultdict
from datetime import datetime
from typing import Any # For dict values in faculty_data

import polars as pl

from easy_access.db.ingest import load_base_data
from easy_access.db.update import update_copyright_items, update_copyright_relations
from easy_access.settings import COURSE_MAPPING, FINE_AMOUNT, SETTINGS, DirSetting
from easy_access.sheets.sheet import finalize_sheet, store_complete_data
from easy_access.utils import Directory, File

logger = logging.getLogger(__name__)


def create_programme_overviews(
    all_faculty_data: pl.DataFrame,
    faculty: str,
    style_iter: int,
    file_date_str: str,
) -> int:
    """
    Creates overview Excel sheets for each programme within a given faculty.

    The data for each programme is filtered from `all_faculty_data`. Derived fields
    like 'possible_fine' and 'infringement' status are calculated. Each programme's
    overview is saved as a separate, styled Excel file named using the programme group,
    faculty, and the provided date string.

    Args:
        all_faculty_data (pl.DataFrame): DataFrame containing data for the entire faculty.
        faculty (str): The abbreviation of the faculty whose programmes are being processed.
        style_iter (int): An integer used to cycle through table styles for visual distinction.
        file_date_str (str): Date string (YYYY-MM-DD) to incorporate into output filenames.

    Returns:
        int: The updated `style_iter` value.

    Raises:
        KeyError: If the faculty is not found in `COURSE_MAPPING`.
    """
    if faculty not in COURSE_MAPPING:
        logger.warning(f"Faculty '{faculty}' not found in COURSE_MAPPING. Cannot create programme overviews.")
        return style_iter

    course_to_group: dict[str, str] = COURSE_MAPPING[faculty]
    programme_data_map: dict[str, list[pl.DataFrame]] = defaultdict(list) # Store list of DFs per group to concat later

    for course_key_in_mapping, programme_group_name in course_to_group.items():
        # Filter based on 'department' column which often holds course/programme unique identifiers from Osiris/Canvas
        # The key in COURSE_MAPPING (course_key_in_mapping) should match values in 'department' column.
        programme_specific_data = all_faculty_data.filter(pl.col("department") == course_key_in_mapping)

        if programme_specific_data.is_empty():
            logger.debug(f"No data for course/dept '{course_key_in_mapping}' (group: {programme_group_name}) in faculty '{faculty}'.")
            continue

        # Calculate derived fields
        # Possible Fine Calculation
        if "pages_x_students" in programme_specific_data.columns:
            fine_series = programme_specific_data["pages_x_students"].cast(pl.Int64, strict=False).fill_null(0) * FINE_AMOUNT
            if "possible_fine" in programme_specific_data.columns:
                programme_specific_data = programme_specific_data.with_columns(
                    pl.coalesce(pl.col("possible_fine").cast(pl.Float64, strict=False), fine_series).alias("possible_fine")
                )
            else:
                programme_specific_data = programme_specific_data.with_columns(fine_series.alias("possible_fine"))
        else:
            logger.warning(f"'pages_x_students' column not found for {course_key_in_mapping}, cannot calculate possible_fine.")
            programme_specific_data = programme_specific_data.with_columns(pl.lit(None, dtype=pl.Float64).alias("possible_fine"))

        # Infringement Status Calculation
        if "manual_classification" in programme_specific_data.columns:
            programme_specific_data = programme_specific_data.with_columns(
                infringement=pl.when(pl.col("manual_classification").is_null() | pl.col("manual_classification").str.is_empty() | (pl.col("manual_classification") == "-"))
                .then(pl.lit("undetermined"))
                .when(pl.col("manual_classification").str.to_lowercase().str.contains_any(["open access", "eigen materiaal", "overig", "deleted"])) # Added "eigen materiaal"
                .then(pl.lit("no"))
                .when(pl.col("manual_classification").str.to_lowercase().str.contains("lange"))
                .then(pl.lit("yes"))
                .otherwise(pl.lit("maybe"))
            )
        else:
            logger.warning(f"'manual_classification' column not found for {course_key_in_mapping}, cannot determine infringement status.")
            programme_specific_data = programme_specific_data.with_columns(pl.lit("undetermined").alias("infringement"))

        programme_data_map[programme_group_name].append(programme_specific_data)

    final_programme_dfs: dict[str, pl.DataFrame] = {
        group: pl.concat(dfs, how="diagonal_relaxed") for group, dfs in programme_data_map.items() if dfs
    }

    for group_name, group_df in final_programme_dfs.items():
        logger.info(f"Programme group '{group_name}' for faculty '{faculty}': {group_df.height} items.")

        # Define output directory and handle existing files (backup/move)
        programme_output_dir = SETTINGS.dirs[DirSetting.FACULTIES_DIR].full / faculty / "per_programme"
        programme_output_dir.mkdir(parents=True, exist_ok=True) # Ensure it exists

        backup_dir_for_programme_overviews = SETTINGS.dirs[DirSetting.OVERVIEWS_BACKUP].full / faculty / "per_programme"
        backup_dir_for_programme_overviews.mkdir(parents=True, exist_ok=True)

        # Check for existing overview files for this group and move them to backup
        for existing_file in Directory(programme_output_dir).files: # Uses utils.Directory
            if existing_file.name.startswith(f"{group_name}_total_overview_updated_") and existing_file.extension in [".xlsx", ".xls"]:
                try:
                    existing_file.move(backup_dir_for_programme_overviews / existing_file.name)
                    logger.debug(f"Moved existing programme overview {existing_file.name} to backup.")
                except Exception as e_move:
                    logger.warning(f"Could not move existing programme overview {existing_file.name} to backup: {e_move}")

        output_filename = f"{group_name}_total_overview_updated_{file_date_str}.xlsx"
        programme_file_obj = File(programme_output_dir / output_filename) # utils.File

        logger.info(f"Saving programme overview for '{group_name}' ({faculty}) with {group_df.height} rows to {programme_file_obj.path.name}")
        store_complete_data(programme_file_obj, group_df) # Handles its own logging
        style_iter = finalize_sheet(programme_file_obj, group_df, style_iter) # Handles its own logging
    return style_iter


def create_faculty_overviews(
    faculty_data_map: dict[str, pl.DataFrame], # Renamed from faculty_data for clarity
    style_iter: int,
    disable_writes: bool = False,
    file_date_str: str | None = None
) -> int:
    """
    Creates overview Excel sheets for each faculty and orchestrates programme-level overviews.

    Args:
        faculty_data_map (dict[str, pl.DataFrame]): A dictionary where keys are faculty
            abbreviations (e.g., "BMS") and values are DataFrames containing data for that faculty.
        style_iter (int): An integer used to cycle through table styles.
        disable_writes (bool, optional): If True, file writing operations are skipped. Defaults to False.
        file_date_str (str | None, optional): Date string (YYYY-MM-DD) for naming output files.
                                            If None, current date is used.

    Returns:
        int: The updated `style_iter` value.
    """
    report_date_str = file_date_str if file_date_str else datetime.now().strftime("%Y-%m-%d")
    aggregated_data_for_db_update: list[pl.DataFrame] = []

    if disable_writes:
        logger.info("Write operations disabled; skipping creation of faculty and programme overviews.")

    for faculty_abbr, faculty_df in faculty_data_map.items():
        if disable_writes: # Skip file operations if writes are disabled
            aggregated_data_for_db_update.append(faculty_df) # Still aggregate for potential DB update
            continue

        # Create programme-level overviews if applicable for the faculty
        if faculty_abbr in COURSE_MAPPING: # COURSE_MAPPING from settings
            logger.info(f"Creating programme overviews for faculty: {faculty_abbr}")
            style_iter = create_programme_overviews(
                faculty_df, faculty_abbr, style_iter, report_date_str
            )

        if faculty_df.is_empty():
            logger.info(f"No data for faculty {faculty_abbr}; skipping faculty overview sheet.")
            continue

        # Create faculty-level overview sheet
        faculty_dir = SETTINGS.dirs[DirSetting.FACULTIES_DIR].full / faculty_abbr
        faculty_dir.mkdir(parents=True, exist_ok=True) # Ensure dir exists

        output_filename = f"{faculty_abbr}_total_overview_updated_{report_date_str}.xlsx"
        faculty_overview_file = File(faculty_dir / output_filename) # utils.File

        logger.info(f"Saving faculty overview for {faculty_abbr} with {faculty_df.height} rows to {faculty_overview_file.path.name}.")
        store_complete_data(faculty_overview_file, faculty_df)
        style_iter = finalize_sheet(faculty_overview_file, faculty_df, style_iter)

        aggregated_data_for_db_update.append(faculty_df)

    # The following DB update seems out of place for a function named "create_faculty_overviews".
    # This implies that the data passed in (faculty_data_map) might contain modifications
    # that need to be synced back. This should ideally be part of the main data processing pipeline.
    if aggregated_data_for_db_update and not disable_writes: # Only update DB if data exists and writes not disabled
        logger.info("Aggregated data from overviews will be passed to update_db.")
        # This asyncio call can be problematic if create_faculty_overviews is called from non-async context
        # and no event loop is running.
        try:
            loop = asyncio.get_event_loop()
            if loop.is_running(): # pragma: no cover
                 # This case is complex: if called from within an existing async task,
                 # run_until_complete cannot be used. Await directly or use create_task.
                 # For now, this indicates a potential design issue if called from async.
                 logger.warning("update_db called from within a running event loop. Direct await might be needed.")
                 # loop.create_task(update_db(aggregated_data_for_db_update)) # Fire-and-forget, not ideal
            else: # pragma: no cover
                loop.run_until_complete(update_db(aggregated_data_for_db_update))
        except RuntimeError: # No event loop, try to run it simply (common in scripts)
             asyncio.run(update_db(aggregated_data_for_db_update)) # Python 3.7+
        except Exception as e_async:
            logger.error(f"Error running update_db for overview data: {e_async}")

    return style_iter


async def update_db(datalist: list[pl.DataFrame]) -> None:
    """
    Updates the database with data aggregated from overview sheets.
    This function is intended to be called after overview sheets are generated
    and might contain user modifications that need to be reflected in the DB.

    Args:
        datalist (list[pl.DataFrame]): A list of DataFrames, typically one per faculty,
                                     containing data to update in the database.
    """
    if not datalist:
        logger.info("No data provided to update_db from overviews. Skipping DB update.")
        return

    logger.info("Preparing to update database with aggregated data from overviews.")
    await init_tortoise() # Ensure Tortoise is initialized

    try:
        # It's assumed load_base_data ensures foundational data like Faculties exist.
        # This might be redundant if base data is guaranteed to be stable.
        await load_base_data() # This function handles its own Tortoise connections

        # Concatenate all DataFrames in the list
        # Perform unique operation based on material_id, keeping the first occurrence
        # This is a simple strategy; more complex merging might be needed if data conflicts across DFs.
        combined_df = pl.concat(datalist, how="diagonal_relaxed").unique(subset=["material_id"], keep="first", maintain_order=False)

        if combined_df.is_empty():
            logger.info("Combined data for DB update is empty. No updates will be performed.")
            return

        logger.info(f"Updating/creating {combined_df.height} items in DB from overview data.")
        # update_copyright_items should handle heuristics of what to update.
        # For overview sheets, specific fields are usually user-editable and take precedence.
        # This might require a specific 'source' flag or 'overwrite' mode in update_copyright_items.
        # Assuming update_copyright_items has logic to handle this (e.g. if called with overwrite=True for certain fields)
        await update_copyright_items(combined_df, overwrite=False, update_relations=True) # update_relations=True to be safe

        # update_copyright_relations is called by update_copyright_items if update_relations=True
        # If not, it might be needed here: await update_copyright_relations()
        logger.info("Database update from overview data complete.")
    except Exception as e:
        logger.error(f"Error updating database from overview data: {e}")
        logger.debug(traceback.format_exc())
    finally:
        await Tortoise.close_connections()
