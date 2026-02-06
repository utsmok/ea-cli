import asyncio
from datetime import datetime

import polars as pl
from loguru import logger

from easy_access.db.ingest import load_base_data
from easy_access.db.update import update_copyright_items

# from easy_access.settings import COURSE_MAPPING, FINE_AMOUNT, SETTINGS, DirSetting # Will be passed
from easy_access.settings import DirSetting, Settings  # Keep for type hinting
from easy_access.sheets.sheet import finalize_sheet, store_complete_data
from easy_access.utils import Directory, File


async def create_faculty_overviews(
    settings: Settings,  # Added settings
    faculty_data: dict[str, pl.DataFrame],
    style_iter: int,
    disable_writes: bool = False,
) -> int:
    """
    Input:
    faculty_data: dict[str, pl.DataFrame]:
        keys are the faculty names (BMS, EEMCS, etc)
        values are the dataframes with the data for each faculty
    """
    # loop over the faculties
    # for each, read in all data and store
    today = datetime.now().strftime("%Y-%m-%d")
    data_to_update = []

    if disable_writes:
        logger.info(
            "write operations disabled, skipping creation of faculty and programme overviews"
        )
    for faculty, all_faculty_data in faculty_data.items():
        if all_faculty_data.is_empty():
            continue

        if not disable_writes:
            fac_file = File(
                path=settings.dirs[DirSetting.FACULTIES_DIR].full
                / faculty
                / f"{faculty}_total_overview_updated_{today}.xlsx"
            )
            # Ensure there is never more than one total_overview sheet per directory.
            # Move any existing overview files for this faculty into the overviews_backup directory.
            overview_dir = Directory(
                settings.dirs[DirSetting.OVERVIEWS_BACKUP].full / faculty
            )
            overview_dir.full.mkdir(parents=True, exist_ok=True)
            faculty_dir = Directory(
                settings.dirs[DirSetting.FACULTIES_DIR].full / faculty
            )
            for file in faculty_dir.files:
                if file.extension in [".xls", ".xlsx"] and "overview" in file.name:
                    try:
                        file.move(overview_dir.full / file.name)
                    except Exception:
                        logger.exception(
                            f"Failed to move existing overview {file.path} to backup"
                        )

            logger.info(
                f"saving file with {all_faculty_data.shape[0]} rows to {fac_file.path}"
            )
            store_complete_data(settings=settings, file=fac_file, data=all_faculty_data)

            style_iter = finalize_sheet(
                settings=settings,
                file=fac_file,
                data=all_faculty_data,
                style_iter=style_iter,
            )
        data_to_update.append(all_faculty_data)

    # [FIX] Removed update_db call from export cycle to ensure export is idempotent.
    # Database updates should happen during ingest or process stages.
    # if not disable_writes:
    #     await update_db(settings=settings, datalist=data_to_update)
    # else:
    #     logger.info("DB update skipped because export was run with disable_writes=True")
    logger.info("Export cycle complete (DB update intentionally skipped in this stage).")

    return style_iter


async def update_db(settings: Settings, datalist: list[pl.DataFrame]):  # Added settings
    logger.info("Moving updates into database.")

    await load_base_data(settings=settings)

    # concat all dfs
    df = pl.concat(datalist, how="diagonal_relaxed")
    df = df.unique()

    await update_copyright_items(settings, df)


def create_faculty_overviews_sync(
    settings: Settings,
    faculty_data: dict[str, pl.DataFrame],
    style_iter: int,
    disable_writes: bool = False,
) -> int:
    """
    Synchronous wrapper for create_faculty_overviews.
    Used by legacy synchronous code that hasn't been migrated to async.
    """
    return asyncio.run(
        create_faculty_overviews(
            settings=settings,
            faculty_data=faculty_data,
            style_iter=style_iter,
            disable_writes=disable_writes,
        )
    )
