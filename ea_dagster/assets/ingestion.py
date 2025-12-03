"""
Ingestion assets using dlt (Data Load Tool).

This module replaces the manual staging logic in `easy_access.db.ingest`
with dlt-based data loading that writes to SQLite staging tables.
"""

import dlt
import polars as pl
from dagster import AssetExecutionContext, asset

from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import standardize_dataframe


def get_pipeline(dataset_name: str = "staging"):
    """
    Creates a dlt pipeline configured to write to the SQLite database.

    Args:
        dataset_name: The dataset/schema name for tables (default: "staging")

    Returns:
        dlt.Pipeline: Configured pipeline instance
    """
    db_path = SETTINGS.db_path.absolute()
    # dlt uses sqlalchemy destination for SQLite
    return dlt.pipeline(
        pipeline_name="ea_ingest",
        destination=dlt.destinations.sqlalchemy(f"sqlite:///{db_path}"),
        dataset_name=dataset_name,
    )


@asset(group_name="ingestion")
def raw_copyright_data(context: AssetExecutionContext) -> str:
    """
    Reads the latest Excel export from CopyRight, standardizes columns,
    and loads to SQLite via dlt.

    This asset replaces `easy_access.db.ingest.load_raw_copyright_data_to_staging`.

    The dlt pipeline creates a table named `staging__raw_copyright_items` in db.sqlite.

    Returns:
        str: Summary of rows loaded
    """
    # 1. Find the newest file in the raw copyright data directory
    raw_dir = SETTINGS.dirs.get(DirSetting.RAW_COPYRIGHT_DATA)
    if not raw_dir:
        context.log.warning("RAW_COPYRIGHT_DATA directory not configured in settings")
        return "No raw copyright data directory configured"

    file = raw_dir.newest_file(file_type=[".xlsx", ".xls"])
    if not file:
        context.log.info("No raw copyright file found in directory")
        return "No raw copyright file found"

    context.log.info(f"Ingesting file: {file.path}")

    # 2. Read with Polars & Standardize
    # Use existing standardize_dataframe to maintain column renaming logic
    try:
        df = pl.read_excel(file.path)
    except Exception as e:
        context.log.error(f"Failed to read Excel file: {e}")
        return f"Error reading file: {e}"

    df = standardize_dataframe(df)

    # Add retrieval timestamp
    latest_file_date = file.created.strftime("%Y-%m-%d") if file.created else None
    if latest_file_date:
        df = df.with_columns(
            pl.lit(latest_file_date).alias("retrieved_from_copyright_on")
        )

    context.log.info(f"Standardized dataframe with {df.height} rows, {df.width} columns")

    # 3. Load to DB via dlt
    pipeline = get_pipeline()

    # Use write_disposition="replace" to clear old staging data
    # dlt accepts Polars DataFrames via Arrow
    info = pipeline.run(
        df.to_arrow(),
        table_name="raw_copyright_items",
        write_disposition="replace",
    )

    context.log.info(f"dlt load info: {info}")
    return f"Loaded {df.height} rows to staging.raw_copyright_items"


@asset(group_name="ingestion")
def faculty_updates_data(context: AssetExecutionContext) -> str:
    """
    Ingests faculty update sheets using dlt.

    This asset replaces `easy_access.db.ingest.load_faculty_updates_to_staging`.

    The dlt pipeline creates a table named `staging__faculty_updates` in db.sqlite.

    Returns:
        str: Summary of rows loaded
    """
    from easy_access.sheets.sheet import read_faculty_sheets

    # Read faculty sheets using existing logic
    try:
        df = read_faculty_sheets(SETTINGS)
    except Exception as e:
        context.log.error(f"Failed to read faculty sheets: {e}")
        return f"Error reading faculty sheets: {e}"

    if df.is_empty():
        context.log.info("No faculty updates found to ingest")
        return "No faculty updates found"

    # Select only relevant columns for updates
    update_cols = ["material_id", "manual_classification", "remarks", "workflow_status"]
    available_cols = [c for c in update_cols if c in df.columns]

    if not available_cols or "material_id" not in available_cols:
        context.log.warning("Required columns not found in faculty sheets")
        return "Required columns not found in faculty sheets"

    df = df.select(available_cols)
    df = standardize_dataframe(df)

    context.log.info(f"Faculty updates dataframe: {df.height} rows")

    # Load to DB via dlt
    pipeline = get_pipeline()

    info = pipeline.run(
        df.to_arrow(),
        table_name="faculty_updates",
        write_disposition="replace",
    )

    context.log.info(f"dlt load info: {info}")
    return f"Loaded {df.height} rows to staging.faculty_updates"
