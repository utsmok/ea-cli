"""
Ingestion assets using dlt (Data Load Tool).

This module replaces the manual staging logic in `easy_access.db.ingest`
with dlt-based data loading that writes to SQLite staging tables.
"""

import dlt
import polars as pl
from dagster import (
    AssetExecutionContext,
    MetadataValue,
    Output,
    TableColumn,
    TableSchema,
    asset,
)

from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import standardize_dataframe


def get_pipeline(pipeline_name: str, dataset_name: str = "staging") -> dlt.Pipeline:
    """
    Creates a dlt pipeline configured to write to the SQLite database.

    Args:
        pipeline_name: Unique name for this specific extraction job (required).
                       This prevents file lock conflicts in ~/.dlt/pipelines/.
        dataset_name: The database schema/dataset name (default: "staging").
                      Both pipelines should use "staging" so tables end up together.

    Returns:
        dlt.Pipeline: Configured pipeline instance
    """
    db_path = SETTINGS.db_path.absolute()

    # We do NOT cache the pipeline object globally. dlt pipeline objects are
    # lightweight wrappers around a state folder; it is safer to instantiate
    # them fresh per asset execution.
    return dlt.pipeline(
        pipeline_name=pipeline_name,
        destination=dlt.destinations.sqlalchemy(f"sqlite:///{db_path}"),
        dataset_name=dataset_name,
        # 'progress' helps visualize dlt actions in Dagster logs
        progress="log",
    )


@asset(group_name="ingestion")
def raw_copyright_data(context: AssetExecutionContext) -> Output:
    """
    Reads the latest Excel export from CopyRight and loads to SQLite via dlt.
    """
    # 1. Find the newest file
    raw_dir = SETTINGS.dirs.get(DirSetting.RAW_COPYRIGHT_DATA)
    if not raw_dir:
        context.log.warning("RAW_COPYRIGHT_DATA directory not configured")
        return Output(value="No raw copyright data directory configured")

    file = raw_dir.newest_file(file_type=[".xlsx", ".xls"])
    if not file:
        context.log.info("No raw copyright file found")
        return Output(value="No raw copyright file found")

    context.log.info(f"Ingesting file: {file.path}")

    # 2. Read & Standardize
    try:
        df = pl.read_excel(file.path)
    except Exception as e:
        context.log.error(f"Failed to read Excel file: {e}")
        raise

    df = standardize_dataframe(df)
    context.log.info(
        f"Standardized dataframe: {df.height} rows, columns: {list(df.columns)}"
    )

    # Add retrieval timestamp
    latest_file_date = file.created.strftime("%Y-%m-%d") if file.created else None
    if latest_file_date:
        df = df.with_columns(
            pl.lit(latest_file_date).alias("retrieved_from_copyright_on")
        )
        context.log.info(f"Added retrieval timestamp: {latest_file_date}")

    # 3. Load via dlt
    pipeline = get_pipeline(pipeline_name="ingest_raw_copyright")

    info = pipeline.run(
        df.to_arrow(),
        table_name="raw_copyright_items",
        write_disposition="replace",
    )

    context.log.info(f"dlt load info: {info}")

    # Metadata
    columns = [TableColumn(name=col) for col in df.columns]
    schema = TableSchema(columns=columns)
    metadata = {
        "table_name": "staging__raw_copyright_items",
        "table_schema": MetadataValue.table_schema(schema),
    }
    return Output(
        value=f"Loaded {df.height} rows to staging.raw_copyright_items",
        metadata=metadata,
    )


@asset(group_name="ingestion")
def faculty_updates_data(context: AssetExecutionContext) -> Output:
    """
    Ingests faculty update sheets using dlt.
    """
    from easy_access.sheets.sheet import read_faculty_sheets

    # 1. Read faculty sheets
    try:
        df = read_faculty_sheets(SETTINGS)
    except Exception as e:
        context.log.error(f"Failed to read faculty sheets: {e}")
        raise

    if df.is_empty():
        context.log.info("No faculty updates found")
        return Output(value="No faculty updates found")

    context.log.info(f"Read faculty sheets: {df.height} rows")

    # 2. Select & Standardize
    update_cols = ["material_id", "manual_classification", "remarks", "workflow_status"]
    available_cols = [c for c in update_cols if c in df.columns]

    if not available_cols or "material_id" not in available_cols:
        context.log.warning("Required columns not found in faculty sheets")
        return Output(value="Required columns not found")

    df = df.select(available_cols)
    context.log.info(f"Selected columns: {available_cols}")

    df = standardize_dataframe(df)
    context.log.info(
        f"Standardized dataframe: {df.height} rows, columns: {list(df.columns)}"
    )

    # 3. Load via dlt
    # CRITICAL: We pass a UNIQUE pipeline_name for this asset
    pipeline = get_pipeline(pipeline_name="ingest_faculty_updates")

    info = pipeline.run(
        df.to_arrow(),
        table_name="faculty_updates",
        write_disposition="replace",
    )

    context.log.info(f"dlt load info: {info}")

    # Metadata
    columns = [TableColumn(name=col) for col in df.columns]
    schema = TableSchema(columns=columns)
    metadata = {
        "table_name": "staging__faculty_updates",
        "table_schema": MetadataValue.table_schema(schema),
    }
    return Output(
        value=f"Loaded {df.height} rows to staging.faculty_updates", metadata=metadata
    )
