"""
Export module for creating Excel sheets from processed copyright data.

This module provides functions to export copyright data to various Excel formats:
- Faculty overview sheets
- Programme sheets within faculties
- All items sheet
- Faculty overviews with data entry sheets

All functions follow the new DB-first architecture and use the Settings system.
"""

import csv
from datetime import datetime
from io import TextIOWrapper
from pathlib import Path

import polars as pl
from loguru import logger

from easy_access.db.retrieve import retrieve_full_data
from easy_access.settings import DirSetting, OverrideSettings, Settings
from easy_access.sheets.analysis import create_faculty_overviews
from easy_access.sheets.backup import backup_existing_file
from easy_access.sheets.sheet import (
    _read_excel_quiet,
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

    if "canvas_course_id" in all_data.columns:
        base_url = settings.university_settings.lms.url

        all_data = all_data.with_columns(
            pl.when(pl.col("canvas_course_id").is_not_null())
            .then(
                pl.concat_str(
                    [
                        pl.lit(f"{base_url}/courses/"),
                        pl.col("canvas_course_id").cast(pl.Utf8),
                        pl.lit("/files/search?search_term="),
                        pl.col("filename").str.replace_all(" ", "%20"),
                    ],
                    separator="",
                )
            )
            .otherwise(pl.lit(""))
            .alias("course_link")
        )
        # debug: print first 5 unique course links
        unique_links = all_data.select("course_link").unique().to_series().to_list()
        logger.debug(f"Sample course links: {unique_links[:5]}")

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
                existing_df = _read_excel_quiet(
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

        # see if there is a .yaml override file in the faculty directory
        for file in Directory(faculty_dir.full / faculty).files:
            if file.extension in [".yml", ".yaml"] and "settings" in file.name:
                logger.info(f"Applying override settings from {file.name}")
                settings = OverrideSettings(override_input_file_path=file.path)

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

def add_table(fh: TextIOWrapper, update_stats:dict[str, dict[str, int]]) -> None:
    """
    Helper function to write parsed update data into a formatted table in a text file.
    Used by export_faculty_workflow_files.
    """
    fh.write(f"\n{f'{"-" * 12}-{"-" * 5}-{"-" * 5}-{"-" * 5}':{' '}^{40}}")
    fh.write(
        f"\n{f'{"Sheet":{" "}<12}|{"Old":{" "}^{5}}|{"New":{" "}^{5}}|{"Δ":{" "}^{5}}':{' '}^{40}}"
    )
    fh.write(f"\n{f'{"-" * 12}+{"-" * 5}+{"-" * 5}+{"-" * 5}':{' '}^{40}}")
    for bucket_name, stats in update_stats.items():
        fh.write(
            f"\n{f'{bucket_name:{" "}<12}|{stats["old"]:{" "}^{5}}|{stats["new"]:{" "}^{5}}|{stats["new"] - stats["old"]:^+5}':{' '}^{40}}"
        )
    fh.write(f"\n{f'{"-" * 12}-{"-" * 5}-{"-" * 5}-{"-" * 5}':{' '}^{40}}\n")

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
        for file in faculty_dir.files:
            if file.extension in [".yml", ".yaml"] and "settings" in file.name:
                logger.info(f"Applying override settings from {file.name}")
                settings = OverrideSettings(override_input_file_path=file.path)

        # small backups dir inside faculty dir
        backups_dir_base = settings.dirs[DirSetting.OVERVIEWS_BACKUP].full
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
                    update_stats[bucket_name]["old"] = _read_excel_quiet(
                        target_path
                    ).shape[0]
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

            # protect done.xlsx and set active sheet to Data Entry
            if bucket_name in ["done", "overview"]:
                try:
                    protect_workbook(
                        file_path=target_path,
                        settings=settings,
                        protect_sheets=[
                            settings.data_settings.complete_data_name,
                            settings.data_settings.data_entry_name,
                        ],
                        active_sheet=settings.data_settings.data_entry_name,
                    )
                    logger.info(f"Protected {target_path.name}")
                except Exception as e:
                    logger.warning(f"Failed to protect workbook {target_path}: {e}")
        # Store update data
        # 1. data in a simple csv file in the faculty_sheets root folder
        # 2. a simple text file in each faculty folder summarizing the latest update(s)

        # CSV file
        # columns: timestamp, faculty, bucket, old, new, delta
        # append to file if it exists, otherwise create with header
        # only add rows if there was a change (delta != 0)

        summary_file = settings.dirs[DirSetting.FACULTIES_DIR].full / "update_overview.csv"
        mode = "a" if summary_file.exists() else "w"
        diff = False
        with summary_file.open(mode, newline="", encoding="utf-8") as fh:
            writer = csv.writer(fh)
            if mode == "w":
                writer.writerow(["timestamp", "faculty", "bucket", "old", "new", "delta"])
            for bucket_name, stats in update_stats.items():
                if stats["new"] - stats["old"] != 0:
                    diff = True
                    writer.writerow([
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        faculty,
                        bucket_name,
                        stats["old"],
                        stats["new"],
                        stats["new"] - stats["old"],
                    ])

        # Text files
        # we always create a new file with the current timestamp in the name

        # if a file already exists, and there is no diff, we keep the content of the file except:
        #   add line: [sync @ timestamp]: No changes (before the table)
        #   modify the Last sync with main database line to the current timestamp

        # if there are changes, we recreate the entire file with the new stats table
        file_contents = []
        # read existing file content, store as list of lines, delete old file
        for file in faculty_dir.files:
            if file.name.startswith("update_info_") and file.extension == ".txt":
                try:
                    with file.path.open("r", encoding="utf-8") as fh:
                        file_contents = fh.readlines()
                    file.path.unlink()
                except Exception:
                    logger.debug(
                        f"Failed to remove old update info file {file.path}; continuing"
                    )
        # now create a new one with current timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        with (faculty_dir.full / f"update_info_{timestamp}.txt").open(
            "w", encoding="utf-8"
        ) as fh:
            if not diff and file_contents:
                # keep old content, but add a line about no changes
                skip = False
                syncstr = f"[{datetime.now().strftime('%Y-%m-%d')}]"
                skip2 = False
                for line in file_contents:
                    if skip and skip2: # we are in the list of syncdates without changes
                        if line.strip() == syncstr.strip(): # today is already there
                            syncstr = ""  # only add once
                            fh.write(line)
                            continue
                        if '[' in line: # another date line
                            if syncstr: # add today's date line before the next date line
                                fh.write(syncstr+"\n")
                                syncstr = ""  # only add once
                            fh.write(line) # write old date line
                            continue
                        else: # if no more date lines, we are done with this section
                            if syncstr: # add sync line if not yet added
                                fh.write(syncstr+"\n")
                            skip = False
                            skip2 = False
                    if skip2: # header of list of syncs without changes
                        fh.write(f"Syncs without changes:\n")
                        skip = True
                        continue
                    if skip: # this should be the last sync date with changes
                        fh.write(line)
                        skip2 = True
                        skip = False
                        continue
                    if "Last sync with main database" in line:
                        fh.write(line) # write that line and start processing, see above
                        skip = True
                        continue
                    else:
                        fh.write(line)
            else:
                fh.write(f"\n{'Update information for':{' '}^{40}}\n{faculty:{' '}^{40}}")
                fh.write(
                        f"\n{'Last sync with main database:':{' '}^{40}}\n{datetime.now().strftime('%Y-%m-%d -- %H:%M:%S'):{' '}^{40}}"
                    )
                add_table(fh, update_stats)
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
