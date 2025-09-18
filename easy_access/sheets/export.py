"""
Export module for creating Excel sheets from processed copyright data.

This module provides functions to export copyright data to various Excel formats:
- Faculty overview sheets
- Programme sheets within faculties
- All items sheet
- Faculty overviews with data entry sheets

All functions follow the new DB-first architecture and use the Settings system.
"""

from datetime import datetime
from pathlib import Path

import polars as pl
from loguru import logger

from easy_access.db.retrieve import retrieve_full_data
from easy_access.settings import DirSetting, Settings
from easy_access.sheets.analysis import create_faculty_overviews
from easy_access.sheets.backup import backup_existing_file
from easy_access.sheets.sheet import (
    finalize_sheet,
    protect_workbook,
    store_complete_data,
)
from easy_access.utils import Directory, File


async def gather_faculty_data(settings: Settings) -> dict[str, pl.DataFrame]:
    """
    Retrieves and organizes copyright data by faculty.

    Returns a dictionary mapping faculty names to their respective DataFrames.
    Only includes faculties with data.
    """
    logger.info("Gathering faculty data for export...")

    # Get all data from DB
    all_data = retrieve_full_data(settings=settings)

    if all_data.is_empty():
        logger.warning("No data found for export.")
        return {}

    if "file_exists" in all_data.columns:
        all_data = all_data.with_columns(pl.col("file_exists").cast(pl.Utf8))
        all_data = all_data.with_columns(
            pl.when(pl.col("file_exists").is_in(["1", "True"]))
            .then(pl.lit("Yes"))
            .otherwise(pl.lit("No"))
            .alias("file_exists")
        )

    # Group by faculty
    faculty_data = {}
    faculties = all_data.select("faculty").unique().to_series().to_list()

    for faculty in sorted(faculties):
        if not faculty or faculty == "Unmapped":
            continue

        faculty_df = all_data.filter(pl.col("faculty") == faculty)
        if not faculty_df.is_empty():
            faculty_data[faculty] = faculty_df
            logger.info(f"Faculty {faculty}: {faculty_df.shape[0]} items")
    logger.info(f"Gathered data for {len(faculty_data)} faculties")
    return faculty_data


async def export_faculty_sheets(
    settings: Settings, faculty_data: dict[str, pl.DataFrame], style_iter: int = 9
) -> int:
    """
    Creates individual Excel sheets for each faculty.

    Args:
        settings: Application settings
        faculty_data: Dictionary of faculty DataFrames
        style_iter: Style iterator for Excel formatting

    Returns:
        Updated style iterator
    """
    logger.info("Exporting faculty sheets...")

    for faculty, data in faculty_data.items():
        if data.is_empty():
            continue

        # Create faculty directory if needed
        faculty_dir = Directory(settings.dirs[DirSetting.FACULTIES_DIR].full / faculty)
        faculty_dir.full.mkdir(parents=True, exist_ok=True)

        # Generate filename with date
        today = datetime.now().strftime("%Y-%m-%d")
        filename_base = f"{faculty}_{today}"
        output_file_path = _get_unique_filepath(faculty_dir.full, filename_base)

        # Determine which items are new (not yet present in any existing regular file for this faculty)
        existing_ids: set[int] = set()
        for file in faculty_dir.files:
            # ignore overview files and non-excel
            if file.extension not in [".xls", ".xlsx"]:
                continue
            if "overview" in file.name or "llm" in file.name:
                continue
            try:
                # read the Complete Data sheet from existing file
                existing_df = pl.read_excel(
                    file.path, sheet_name=settings.data_settings.complete_data_name
                )
                if "material_id" in existing_df.columns:
                    existing_ids.update(
                        existing_df.select("material_id").to_series().to_list()
                    )
            except Exception:
                logger.debug(
                    f"Could not read existing faculty file {file.path}; skipping"
                )

        if existing_ids:
            new_data = data.filter(~pl.col("material_id").is_in(list(existing_ids)))
        else:
            new_data = data

        if new_data.is_empty():
            logger.info(f"No new items for faculty {faculty}; skipping regular export")
            continue

        logger.info(
            f"Creating faculty sheet: {output_file_path.name} ({new_data.shape[0]} new items)"
        )

        # Store complete data (only new items)
        store_complete_data(settings=settings, file=output_file_path, data=new_data)

        # Add data entry sheet and styling
        style_iter = finalize_sheet(
            settings=settings,
            file=File(str(output_file_path)),
            data=new_data,
            style_iter=style_iter,
        )

    return style_iter


async def export_faculty_workflow_files(
    settings: Settings, faculty_data: dict[str, pl.DataFrame], style_iter: int = 9
) -> int:
    """
    Export per-faculty files driven by the `workflow_status` column.

    For each faculty produce three files in the faculty folder:
    - inbox.xlsx (ToDo)
    - in_progress.xlsx (InProgress)
    - done.xlsx (Done) -- protected after write

    Existing files are moved into a timestamped backups folder next to the faculty dir.
    """
    logger.info("Exporting faculty workflow files (inbox/in_progress/done)...")

    for faculty, data in faculty_data.items():
        if data.is_empty():
            continue

        faculty_dir = Directory(settings.dirs[DirSetting.FACULTIES_DIR].full / faculty)
        faculty_dir.full.mkdir(parents=True, exist_ok=True)

        # small backups dir inside faculty dir
        backups_dir_base = (
            settings.dirs[DirSetting.OVERVIEWS_BACKUP].full / "v2_style_backups"
        )
        if not backups_dir_base.exists():
            backups_dir_base.mkdir(parents=True, exist_ok=True)

        backups_dir = backups_dir_base / faculty
        if not backups_dir.exists():
            backups_dir.mkdir(parents=True, exist_ok=True)

        # normalize workflow_status and bucket
        df = data.with_columns(
            pl.col("workflow_status").fill_null("ToDo").cast(pl.Utf8)
        )

        buckets: dict[str, pl.DataFrame] = {
            "inbox": df.filter(pl.col("workflow_status").is_in(["ToDo", "todo"])),
            "in_progress": df.filter(
                pl.col("workflow_status").is_in(
                    ["InProgress", "inprogress", "in_progress"]
                )
            ),
            "done": df.filter(pl.col("workflow_status").is_in(["Done", "done"])),
            "overview": df,  # all items for overview file
        }
        update_stats = {
            "inbox": {"old": 0, "new": buckets["inbox"].shape[0]},
            "in_progress": {"old": 0, "new": buckets["in_progress"].shape[0]},
            "done": {"old": 0, "new": buckets["done"].shape[0]},
            "overview": {"old": 0, "new": buckets["overview"].shape[0]},
        }
        for bucket_name, bucket_df in buckets.items():
            filename = bucket_name + ".xlsx"

            target_path = faculty_dir.full / filename

            # backup existing
            if target_path.exists():
                try:
                    update_stats[bucket_name]["old"] = pl.read_excel(target_path).shape[
                        0
                    ]
                    moved = backup_existing_file(
                        target_path=target_path,
                        backups_dir=backups_dir,
                        manifest={"faculty": faculty, "bucket": bucket_name},
                    )
                    logger.info(f"Backed up existing {target_path.name} -> {moved}")
                except Exception as e:
                    logger.warning(f"Failed to backup existing file {target_path}: {e}")

            # write complete data then add data entry sheet
            try:
                if bucket_df.is_empty():
                    logger.info(
                        f"No items for {faculty} -> {bucket_name}; skipping file creation"
                    )
                    continue

                store_complete_data(settings=settings, file=target_path, data=bucket_df)
                style_iter = finalize_sheet(
                    settings=settings,
                    file=File(str(target_path)),
                    data=bucket_df,
                    style_iter=style_iter,
                )
                logger.info(f"Wrote {len(bucket_df)} rows to {target_path}")
            except Exception as e:
                logger.error(f"Failed writing faculty workflow file {target_path}: {e}")
                raise e
                continue

            # protect done.xlsx and set active sheet to Data Entry
            if bucket_name in ["done", "overview"]:
                try:
                    protect_workbook(
                        target_path,
                        protect_sheets=[
                            settings.data_settings.complete_data_name,
                            settings.data_settings.data_entry_name,
                        ],
                        active_sheet=settings.data_settings.data_entry_name,
                    )
                    logger.info(f"Protected {target_path.name}")
                except Exception as e:
                    logger.warning(f"Failed to protect workbook {target_path}: {e}")

        # Create a simple text file indicating last update time
        # first remove any existing update_info_*.txt files
        for file in faculty_dir.files:
            if file.name.startswith("update_info_") and file.extension == ".txt":
                try:
                    file.path.unlink()
                except Exception:
                    logger.debug(
                        f"Failed to remove old update info file {file.path}; continuing"
                    )
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        with (faculty_dir.full / f"update_info_{timestamp}.txt").open(
            "w", encoding="utf-8"
        ) as fh:
            fh.write(f"\n{'Update information for':{' '}^{40}}\n{faculty:{' '}^{40}}")
            fh.write(
                f"\n{'Last sync with main database:':{' '}^{40}}\n{datetime.now().strftime('%Y-%m-%d -- %H:%M:%S'):{' '}^{40}}"
            )
            # write update stats
            fh.write(f"\n{f'{"-" * 12}-{"-" * 5}-{"-" * 5}-{"-" * 5}':{' '}^{40}}")
            fh.write(
                f"\n{f'{"Sheet":{" "}<12}|{"Old":{" "}^{5}}|{"New":{" "}^{5}}|{"Δ":{" "}^{5}}':{' '}^{40}}"
            )
            fh.write(f"\n{f'{"-" * 12}+{"-" * 5}+{"-" * 5}+{"-" * 5}':{' '}^{40}}")
            for bucket_name, stats in update_stats.items():
                fh.write(
                    f"\n{f'{bucket_name:{" "}<12}|{stats["old"]:{" "}^{5}}|{stats["new"]:{" "}^{5}}|{stats["new"] - stats["old"]:^+5}':{' '}^{40}}"
                )
            fh.write(f"\n{f'{"-" * 12}-{"-" * 5}-{"-" * 5}-{"-" * 5}\n':{' '}^{40}}")
    return style_iter


async def export_programme_sheets(
    settings: Settings, faculty_data: dict[str, pl.DataFrame], style_iter: int = 9
) -> int:
    """
    Creates programme sheets within each faculty directory.

    Args:
        settings: Application settings
        faculty_data: Dictionary of faculty DataFrames
        style_iter: Style iterator for Excel formatting

    Returns:
        Updated style iterator
    """
    logger.warning("Programme sheet export disabled.")
    return style_iter
    # Uncomment below to enable programme sheet export
    logger.info("Exporting programme sheets...")

    for faculty, data in faculty_data.items():
        if data.is_empty():
            continue

        # Check if faculty has course mapping
        course_mapping = settings.university_settings.course_mapping.get(faculty)
        if not course_mapping:
            logger.info(
                f"No course mapping for faculty {faculty}, skipping programme sheets"
            )
            continue

        # Create programme directory
        programme_dir = Directory(
            settings.dirs[DirSetting.FACULTIES_DIR].full / faculty / "per_programme"
        )
        programme_dir.full.mkdir(parents=True, exist_ok=True)

        # Group data by department/course
        if "department" not in data.columns:
            logger.warning(
                f"'department' column not found in data for faculty {faculty}"
            )
            continue

        # Create sheets for each course mapping
        today = datetime.now().strftime("%Y-%m-%d")

        for course, group_name in course_mapping.items():
            course_data = data.filter(pl.col("department") == course)
            if course_data.is_empty():
                continue

            filename_base = f"{group_name}_{today}"
            output_file_path = _get_unique_filepath(programme_dir.full, filename_base)

            logger.info(
                f"Creating programme sheet: {output_file_path.name} ({course_data.shape[0]} items)"
            )

            store_complete_data(
                settings=settings, file=output_file_path, data=course_data
            )

            style_iter = finalize_sheet(
                settings=settings,
                file=File(str(output_file_path)),
                data=course_data,
                style_iter=style_iter,
            )

    return style_iter


async def export_all_items_sheet(settings: Settings, style_iter: int = 9) -> int:
    """
    Creates a single Excel sheet containing all copyright items.

    Args:
        settings: Application settings
        style_iter: Style iterator for Excel formatting

    Returns:
        Updated style iterator
    """
    logger.warning("All items sheet export disabled.")
    return style_iter
    # Uncomment below to enable all items sheet export
    logger.info("Exporting all items sheet...")

    # Get all data
    all_data = retrieve_full_data(settings=settings)

    if all_data.is_empty():
        logger.warning("No data found for all items sheet")
        return style_iter

    # Create all items directory
    all_items_dir = Directory(settings.dirs[DirSetting.ALL_ITEMS_DIR].full)
    all_items_dir.full.mkdir(parents=True, exist_ok=True)

    # Generate filename
    today = datetime.now().strftime("%Y-%m-%d")
    filename_base = f"all_items_{today}"
    output_file_path = _get_unique_filepath(all_items_dir.full, filename_base)

    logger.info(
        f"Creating all items sheet: {output_file_path.name} ({all_data.shape[0]} items)"
    )

    # Store data
    store_complete_data(settings=settings, file=output_file_path, data=all_data)

    # Add data entry sheet
    style_iter = finalize_sheet(
        settings=settings,
        file=File(str(output_file_path)),
        data=all_data,
        style_iter=style_iter,
    )

    return style_iter


async def export_faculty_overviews(
    settings: Settings, faculty_data: dict[str, pl.DataFrame], style_iter: int = 9
) -> int:
    """
    Creates overview sheets for each faculty.

    Args:
        settings: Application settings
        faculty_data: Dictionary of faculty DataFrames
        style_iter: Style iterator for Excel formatting

    Returns:
        Updated style iterator
    """
    logger.info("Exporting faculty overviews...")

    # Use existing create_faculty_overviews function
    updated_style_iter = await create_faculty_overviews(
        settings=settings,
        faculty_data=faculty_data,
        style_iter=style_iter,
        disable_writes=False,
    )

    return updated_style_iter


async def export_reports_async(settings: Settings) -> None:
    """
    Main export orchestrator that creates all types of export sheets.

    This is the pipeline stage that coordinates:
    - Faculty sheets
    - Programme sheets
    - All items sheet
    - Faculty overviews
    """
    logger.info("Starting export reports...")

    # Gather data
    faculty_data = await gather_faculty_data(settings)

    if not faculty_data:
        logger.warning("No faculty data to export")
        return

    style_iter = 9  # Starting style iterator

    # Export faculty sheets
    style_iter = await export_faculty_sheets(settings, faculty_data, style_iter)

    # Export programme sheets
    style_iter = await export_programme_sheets(settings, faculty_data, style_iter)

    # Export all items sheet
    style_iter = await export_all_items_sheet(settings, style_iter)

    # Export faculty overviews
    style_iter = await export_faculty_overviews(settings, faculty_data, style_iter)

    logger.info("Export reports completed")


def _get_unique_filepath(directory: Path, filename_base: str) -> Path:
    """
    Generates a unique filepath by appending suffix if file already exists.

    Args:
        directory: Directory path
        filename_base: Base filename without extension

    Returns:
        Unique filepath with .xlsx extension
    """
    filepath = directory / f"{filename_base}.xlsx"
    counter = 1

    while filepath.exists():
        filepath = directory / f"{filename_base}_{counter}.xlsx"
        counter += 1

    return filepath
