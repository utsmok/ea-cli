import asyncio
from collections import defaultdict
from datetime import datetime

import polars as pl
from loguru import logger

from easy_access.db.ingest import load_base_data
from easy_access.db.update import update_copyright_items

# from easy_access.settings import COURSE_MAPPING, FINE_AMOUNT, SETTINGS, DirSetting # Will be passed
from easy_access.settings import DirSetting, Settings  # Keep for type hinting
from easy_access.sheets.sheet import finalize_sheet, store_complete_data
from easy_access.utils import Directory, File


def create_programme_overviews(
    settings: Settings,  # Added settings
    all_faculty_data: pl.DataFrame,
    faculty: str,
    style_iter: int,
):
    """
    create an overview sheet for each programme of the given faculty, using the data in df.
    """
    course_to_group: dict[str, str] = settings.university_settings.course_mapping[
        faculty
    ]
    data: dict[str, pl.DataFrame] = defaultdict(pl.DataFrame)
    today = datetime.now().strftime("%Y-%m-%d")

    for course, group in course_to_group.items():
        programme_data = all_faculty_data.filter(pl.col("department") == course)
        if programme_data.is_empty():
            continue
        else:
            if "possible_fine" in programme_data.columns:
                programme_data = programme_data.with_columns(
                    pl.when(
                        pl.col("possible_fine").is_null()
                        | (pl.col("possible_fine") == "")
                    )
                    .then(
                        pl.col("pages_x_students")
                        .cast(pl.Int32)
                        .mul(settings.fine_amount)
                        .alias("possible_fine")
                    )
                    .otherwise(pl.col("possible_fine"))
                )
            else:
                programme_data = programme_data.with_columns(
                    possible_fine=pl.col("pages_x_students")
                    .cast(pl.Int32)
                    .mul(settings.fine_amount)
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
            data[group] = pl.concat(
                [data[group], programme_data], how="diagonal_relaxed"
            )

    for group, item in data.items():
        logger.info(f"group: {group}: {item.shape[0]} items")

    overview_fac_programme_dir = Directory(
        settings.dirs[DirSetting.OVERVIEWS_BACKUP].full / faculty / "per_programme"
    )
    for groupname, df in data.items():
        for file in Directory(
            settings.dirs[DirSetting.FACULTIES_DIR].full / faculty / "per_programme"
        ).files:
            if file.extension not in [".xls", ".xlsx"]:
                continue
            if "overview" in file.name and groupname in file.name:
                file.move(overview_fac_programme_dir.full / file.name)
                continue
        logger.info(f"{groupname} has {df.shape[0]} items")
        programme_file = File(
            settings.dirs[DirSetting.FACULTIES_DIR].full
            / faculty
            / "per_programme"
            / f"{groupname}_total_overview_updated_{today}.xlsx"
        )
        logger.info(f"saving file with {df.shape[0]} rows to {programme_file.path}")
        store_complete_data(
            settings=settings, file=programme_file, data=df
        )  # Pass settings
        style_iter = finalize_sheet(
            settings=settings, file=programme_file, data=df, style_iter=style_iter
        )  # Pass settings

    return style_iter


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
        if (
            faculty in settings.university_settings.course_mapping
            and not disable_writes
        ):
            style_iter = create_programme_overviews(
                settings=settings,
                all_faculty_data=all_faculty_data,
                faculty=faculty,
                style_iter=style_iter,
            )

        if all_faculty_data.is_empty():
            continue

        if not disable_writes:
            fac_file = File(
                path=settings.dirs[DirSetting.FACULTIES_DIR].full
                / faculty
                / f"{faculty}_total_overview_updated_{today}.xlsx"
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

    await update_db(settings=settings, datalist=data_to_update)
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
