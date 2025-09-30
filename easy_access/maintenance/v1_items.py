"""
the v1 sheets (2024-2025) have lots of data that we we want
this module has functions to:

    Read in the sheets from 2024-2025
    parse and clean the data
    merge it into a single dataframe per faculty

    To be developed:
    - function to ingest that data into the v1_CopyrightItem model
    - map to v2 items and classifications
    - update v2 items with useful data from v1 items

"""

import contextlib
import logging
import os
import warnings
from collections import defaultdict
from pathlib import Path

import polars as pl
from loguru import logger
from rich import print

from easy_access.db.base import close_connections, ensure_db_inited
from easy_access.db.enums import (
    Classification,
    Filetype,
    Period,
    Status,
    WorkflowStatus,
)
from easy_access.db.models import v1_CopyrightItem
from easy_access.pdf.download import download_pdfs_for_items
from easy_access.pdf.parse import parse_pdfs
from easy_access.settings import Settings

# Constants
V1_MODEL_FIELDS = {field for field in v1_CopyrightItem._meta.fields}

# mapping/renaming of specific v1 sheet cols to match V1_MODEL_FIELDS
# key = colname in sheet, value = field name in v1_CopyrightItem model
COLMAPPING = {"uploaded by": "owner", "detected_author": "author"}

V1_FIELD_TYPES = {
    "material_id": pl.Int64,
    "workflow_status": WorkflowStatus,
    "retrieved_from_copyright_on": pl.Datetime,
    "ml_prediction": Classification,
    "period": Period,
    "filetype": Filetype,
    "classification": Classification,
    "status": Status,
    "last_change": pl.Date,
    "pagecount": pl.Int64,
    "wordcount": pl.Int64,
    "picturecount": pl.Int64,
    "reliability": pl.Int64,
    "pages_x_students": pl.Int64,
    "count_students_registered": pl.Int64,
}

#
#        Excel file reading and parsing functions
#


def _read_excel_quiet(file_path: str | Path, **kwargs) -> pl.DataFrame:
    """
    Reads an Excel file quietly, suppressing dtype inference messages.

    We redirect stdout/stderr during the read to avoid noisy messages from the
    underlying libraries. If the quiet read fails, a second attempt without
    suppression is performed to raise a visible error.
    """
    # Temporarily raise log level for noisy libraries and silence warnings
    noisy_loggers = ["polars", "openpyxl", "pyxlsb", "lxml"]
    prev_levels = {}
    for name in noisy_loggers:
        lg = logging.getLogger(name)
        prev_levels[name] = lg.level
        lg.setLevel(logging.ERROR)
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            with (
                open(os.devnull, "w") as devnull,
                contextlib.redirect_stdout(devnull),
                contextlib.redirect_stderr(devnull),
            ):
                return pl.read_excel(file_path, **kwargs)
    finally:
        for name, level in prev_levels.items():
            logging.getLogger(name).setLevel(level)


def import_v1_sheets(
    path: Path,
) -> dict[str, dict[str, pl.DataFrame]]:
    """
    Reads in data from the v1 sheets from the given path, and returns a dict structure with dfs for each of the found sheets.
    """
    all_files = list(path.glob("*.xlsx"))
    logger.info(f"Found {len(all_files)} excel files in {path}")
    sheets = {}
    weekly = 0
    for file in all_files:
        try:
            sheets[file.name] = {
                "complete_data": _read_excel_quiet(file, sheet_name="Complete data"),
                "data_entry": _read_excel_quiet(file, sheet_name="Data entry"),
            }
        except Exception as e:
            print(f"Error reading {file.name}: {e}")
            continue
        if "overview" in file.name:
            print(f"found overview sheet: {file.name}")

        else:
            weekly += 1
    logger.info(f"Found {weekly} weekly sheets in total.")

    return sheets


def clean_and_cast_cols(
    df: pl.DataFrame, conflict_cols: list[tuple[str, str]]
) -> pl.DataFrame:
    """
    Cleans and casts columns in the dataframe to appropriate types.
    Handles conflict columns by prioritizing non-null values from the first column.
    """

    # first we normalize col values for key col manual_classification

    man_class_cols: list[str] = []
    if "manual_classification" in df.columns:
        man_class_cols.append("manual_classification")
    if "manual_classification_entry" in df.columns:
        man_class_cols.append("manual_classification_entry")
    if not man_class_cols:
        print("No manual_classification columns found in df.")
    for col in man_class_cols:
        # all values in these cols should be a valid value in the Classification enum.
        # all lowercased strings.
        # Other values should be translated to the correct value.
        # Possible values:
        # - Classification.[CATEGORY] (so the enum class+key as string, e.g. "Classification.ONBEKEND")
        # - anders and/or Classification.ANDERS: do not exist in this enum, translate to "onbekend"

        ALLOWED_VALUES = {c.value.lower(): c.value for c in Classification}
        TRANSLATIONS = {
            "anders": "onbekend",
            "classification.anders": "onbekend",
            "to be determined": "onbekend",
            "eigen materiaal powerpoint": "eigen materiaal - powerpoint",
            "eigen materiaal overig": "eigen materiaal - overig",
            "deleted": "onbekend",
            "removed?": "onbekend",
            "remove": "onbekend",
            "error": "onbekend",
            "overig": "onbekend",
            "publiek domein": "open access",
            "eigen werk - overig": "eigen materiaal - overig",
            "eigen material - overig": "eigen materiaal - overig",
            "removed": "onbekend",
            "eigen werk powerpoint": "eigen materiaal - powerpoint",
            "eigen werk overig": "eigen materiaal - overig",
            "empty": "onbekend",
            "unknown-removed": "onbekend",
            "software manual says it is licensed to the ut but on the side it says only for one computer system": "onbekend",
            "unknown - removed": "onbekend",
            "public domain": "open access",
            "unknown - deleted": "onbekend",
            "done": "onbekend",
        }

        TRANSLATIONS.update(
            {"classification." + c.name.lower(): c.value for c in Classification}
        )
        TRANSLATIONS.update(ALLOWED_VALUES)

        # first strip excess whitespace and lowercase
        df = df.with_columns(
            pl.col(col).str.strip_chars().str.to_lowercase().alias(col)
        )

        # use the translation dict to map values
        df = df.with_columns(
            pl.col(col).replace(TRANSLATIONS, default="onbekend").alias(col)
        )
        print(f"After translation: {(df[col]).unique().sort().to_list()}")
        # check if any values are not in the allowed values
        values = (df[col]).to_list()
        invalid_values = {
            v for v in values if v is not None and v not in ALLOWED_VALUES
        }
        if invalid_values:
            logger.warning(f"Invalid values in {col}: {invalid_values}")

    for col_entry, col_base in conflict_cols:
        if col_base not in df.columns:
            df = df.rename({col_entry: col_base})
            continue

        # for empty col_base fields, overwrite with col_entry, otherwise keep col_base
        if df[col_base].dtype == pl.Utf8:
            df = df.with_columns(
                pl.when(pl.col(col_base).is_null() | (pl.col(col_base) == pl.lit("")))
                .then(pl.col(col_entry))
                .otherwise(pl.col(col_base))
                .alias(col_base)
            )
        else:
            df = df.with_columns(
                pl.when(pl.col(col_base).is_null())
                .then(pl.col(col_entry))
                .otherwise(pl.col(col_base))
                .alias(col_base)
            )

        # specific fields additions where both have values:
        if col_base == "workflow_status":
            # Done > InProgress > ToDo: on conflict, keep the 'highest' status
            df = df.with_columns(
                pl.when(
                    pl.col(col_base).eq(pl.lit("ToDo"))
                    & (
                        (pl.col(col_entry).eq(pl.lit("InProgress")))
                        | (pl.col(col_entry).eq(pl.lit("Done")))
                    )
                )
                .then(pl.col(col_entry))
                .when(
                    (pl.col(col_base) == pl.lit("InProgress"))
                    & (pl.col(col_entry) == pl.lit("Done"))
                )
                .then(pl.col(col_entry))
                .otherwise(pl.col(col_base))
                .alias(col_base)
            )
        elif col_base == "manual_classification":
            # if _entry is "in onderzoek" but base isn't, use base
            # otherwise keep _entry
            df = df.with_columns(
                pl.when(
                    (pl.col(col_entry) == pl.lit("in onderzoek"))
                    & (pl.col(col_base) != pl.lit("in onderzoek"))
                )
                .then(pl.col(col_base))
                .otherwise(pl.col(col_entry))
                .alias(col_base)
            )
        elif col_base == "remarks":
            # if one has more text (longer), use that one
            df = df.with_columns(
                pl.when(
                    pl.col(col_entry).str.len_chars() > pl.col(col_base).str.len_chars()
                )
                .then(pl.col(col_entry))
                .otherwise(pl.col(col_base))
                .alias(col_base)
            )
        df = df.drop(col_entry)

    # finally cast to correct types as per v1_CopyrightItem model
    # fields not included in V1_FIELD_TYPES will be left as-is (string)

    df = df.with_columns(
        [
            pl.col(col).cast(dtype)  # type: ignore
            for col, dtype in V1_FIELD_TYPES.items()
            if (col in df.columns and dtype not in [pl.Date, pl.Datetime])
        ]
    )

    # handle date/datetime columns separately to catch errors
    for col, dtype in V1_FIELD_TYPES.items():
        if col in df.columns and dtype in [pl.Datetime]:
            if df[col].dtype == pl.Datetime:
                continue
            match col:
                case "last_change":
                    format = "%Y-%m-%d %H:%M:%S"
                case "retrieved_from_copyright_on":
                    format = "%Y-%m-%d %H:%M:%S%"
                case _:
                    format = None
            try:
                df = df.with_columns(
                    pl.col(col).str.strptime(dtype, format=format, strict=False)
                )
            except Exception as e:
                logger.error(f"Error parsing datetime column {col}: {e}")
        if col in df.columns and dtype in [pl.Date]:
            if df[col].dtype == pl.Date:
                continue
            try:
                format = "%Y-%m-%d"
                # first remove time part if present
                df = df.with_columns(
                    pl.col(col).str.split(" ").list.first().str.strip_chars().alias(col)
                )
                df = df.with_columns(
                    pl.col(col).str.to_date(strict=False, format=format)
                )
            except Exception as e:
                logger.error(f"Error parsing date column {col}: {e}")
    return df


def process_v1_sheet(
    filename: str,
    dfs: dict[str, pl.DataFrame],
    existing_item_ids: set[int] | list[int] | None = None,
) -> pl.DataFrame:
    """
    input: two dataframes from a signle v1-style sheet:
        - complete_df: dataframe from the 'Complete data' sheet
        - entry_df: dataframe from the 'Data entry' sheet

    does the following:

    - merge the two dataframes on 'material_id' (inner join)
    - grab the list of attributes from the v1_CopyrightItem model (V1_MODEL_FIELDS)
    - select the matching columns from the dataframe

    - clean/parse the data as needed (e.g. convert date strings to date objects, handle missing values, enums, etc)

    - decide what to do on column data conflicts (e.g. if an item appears in both sheets with different data, which one to keep?)

    returns the cleaned/processed dataframe
    """
    complete_df = dfs["complete_data"]
    entry_df = dfs["data_entry"]
    details = False

    try:
        # do immedate merge? or first clean/parse individual dfs before merging?
        combined_df = complete_df.join(
            entry_df, on="material_id", how="inner", suffix="_entry"
        )
    except Exception as e:
        logger.error(f"[{filename}] Error merging dataframes: {e}")
        return pl.DataFrame()

    if existing_item_ids is not None:
        combined_df = combined_df.with_columns(pl.col("material_id").cast(pl.Int64))
        combined_df = combined_df.filter(
            ~pl.col("material_id").is_in(existing_item_ids)
        )
        if combined_df.is_empty():
            return combined_df
    # Col selection
    cols_to_keep = []
    for col in combined_df.columns:
        col_base = col
        if "_entry" in col:
            col_base = col.replace("_entry", "")
        if col in V1_MODEL_FIELDS or col_base in V1_MODEL_FIELDS:
            cols_to_keep.append(col)

    combined_df = combined_df.select(cols_to_keep)

    # String cleaning
    combined_df = combined_df.with_columns(
        pl.selectors.string()
        .str.strip_chars()
        .str.normalize("NFC")
        .str.replace_all(r"\s+", " ")
    )

    # Replace all "-" with None
    combined_df = combined_df.with_columns(
        [
            pl.when(pl.col(col) != "-").then(col)
            for col in combined_df.columns
            if combined_df[col].dtype == pl.Utf8
        ]
    )

    # now clean/parse the cols that are in both sheets and resolve conflicts
    conflict_cols = [
        (col, col.replace("_entry", ""))
        for col in combined_df.columns
        if col.endswith("_entry")
    ]

    combined_df = clean_and_cast_cols(combined_df, conflict_cols)

    if details:
        logger.info(f"Details for {filename}:")
        print_details(combined_df)
    return combined_df


def process_v1_sheets(
    extracted_sheets: dict[str, dict[str, pl.DataFrame]],
    existing_item_ids: set[int] | list[int] | None = None,
) -> dict[str, pl.DataFrame]:
    """
    input: dict of sheets as returned by import_v1_sheets()

    processes each set of sheets using process_v1_sheet(): merge, clean, parse, etc
    returns a dict with the processed dataframes, keyed by filename.
    """

    results = {
        filename: process_v1_sheet(filename, dfs, existing_item_ids)
        for filename, dfs in extracted_sheets.items()
    }

    # filter out keys that have empty dataframes
    return {k: v for k, v in results.items() if not v.is_empty()}


def merge_all_sheets(
    processed_sheets: dict[str, pl.DataFrame],
) -> pl.DataFrame:
    """
    merges all processed dataframes into a single dataframe for further analysis/inspection
    use outer join w/ material_id as key to keep all items
    for conflicting columns, suffix conflicting column with _entry

    once merged, then do another clean_and_cast_cols() to resolve conflicts -- if I'm correct, this should be the same logic as used in process_v1_sheet()
    """

    main_df = pl.DataFrame()
    for filename, df in processed_sheets.items():
        if main_df.is_empty():
            main_df = df
        else:
            try:
                main_df = main_df.join(
                    df, on="material_id", how="outer", suffix="_entry"
                )
            except Exception as e:
                logger.error(f"Error merging {filename}: {e}")
                continue
            try:
                conflict_cols = [
                    (col, col.replace("_entry", ""))
                    for col in main_df.columns
                    if col.endswith("_entry")
                ]
                main_df = clean_and_cast_cols(main_df, conflict_cols)
            except Exception as e:
                logger.error(f"Error cleaning columns in {filename}: {e}")

    print_details(main_df)
    return main_df


#
#   async functions; e.g. to interact with the database
#


async def ingest_v1_items(df: pl.DataFrame) -> None:
    """
    ingests the given dataframe into the v1_CopyrightItem model in the database
    make sure the db is initialized before calling this function
    """

    data = df.to_dicts()

    for row in data:
        await v1_CopyrightItem.update_or_create(**row)


async def ingest_v1_data(settings: Settings, base_dir: Path) -> None:
    """
    given a base directory with subdirectories for each faculty containing v1 sheets,
    ingest all data into the database.
    """
    await ensure_db_inited(settings)
    existing_v1_item_ids: list[int] = await v1_CopyrightItem.all().values_list(
        "material_id", flat=True
    )  # type: ignore

    individual_results = {}
    processed_results = {}
    final_results: dict[str, pl.DataFrame] = {}
    for subdir in base_dir.iterdir():
        if subdir.is_dir():
            logger.info(f"Processing faculty directory: {subdir.name}")

            result = import_v1_sheets(subdir)
            individual_results[subdir.name] = result
            final = process_v1_sheets(result, existing_v1_item_ids)
            if not final:
                logger.info(f"No new items to process for {subdir.name}. Skipping.")
                continue
            processed_results[subdir.name] = final
            main = merge_all_sheets(final)
            final_results[subdir.name] = main

    for _faculty, df in final_results.items():
        await ingest_v1_items(df)
    await close_connections()


async def add_v1_hashes(settings: Settings) -> None:
    """
    for each v1_CopyrightItem that has no filehash, download the pdf if possible and calculate the hash
    then save the hash to the item
    """

    await ensure_db_inited(settings)
    v1_items = await v1_CopyrightItem.filter(filehash=None).all()
    print(f"Found {len(v1_items)} v1 items without a filehash.")
    items_as_dicts = [
        {"material_id": item.material_id, "url": item.url, "filename": item.filename}
        for item in v1_items
        if item.url
    ]
    await download_pdfs_for_items(settings, items_as_dicts)
    await parse_pdfs(
        filter_ids=[item.material_id for item in v1_items], parse_text=False
    )
    await close_connections()


#
#           Utility/debugging functions
#


def print_details(combined_df: pl.DataFrame):
    """utility function to print details about each column in the dataframe"""
    for col in combined_df.columns:
        unique = combined_df[col].unique().sort().drop_nulls().drop_nans().to_list()
        nulls = combined_df[col].is_null().sum()
        print(
            f"--------- {col} | {combined_df[col].dtype if col in combined_df.columns else combined_df[col].dtype} ---------"
        )
        print(f"       {combined_df.height} items            ")
        print(
            f"    {len(unique)} unique ({(len(unique) * 100 / combined_df.height):.2f}%) |   {nulls} nulls ({(nulls * 100 / combined_df.height):.2f}%)    "
        )
        print("Samples:")
        print(unique[:5])


def inspect_fields(complete_df: pl.DataFrame, entry_df: pl.DataFrame) -> pl.DataFrame:
    """
    DEBUG/INITIAL INSPECTION
    function that checks which v1 model fields are in the df, which are missing, and which are untagged (not mapped to v1 model fields)
    prints the results and returns the combined df for further inspection if needed
    """
    try:
        combined_df = complete_df.join(
            entry_df, on="material_id", how="inner", suffix="_entry"
        )
    except Exception as e:
        print(f"Error merging dataframes: {e}")
        return pl.DataFrame()

    v1_fields_in_df: dict[str, (dict[str, list[str]] | list[str])] = {
        "found": defaultdict(list),
        "not_found": list(),
        "untagged": list(),
    }

    dfcols = combined_df.columns
    dfcols_entry = [
        col.replace("_entry", "") for col in dfcols if col.endswith("_entry")
    ]
    dfcols_tag = {col: False for col in dfcols}
    for col in dfcols:
        if col in COLMAPPING:
            v1_fields_in_df["found"][COLMAPPING[col]].append(col)
            dfcols_tag[col] = True
        else:
            if col.endswith("_entry"):
                col = col.replace("_entry", "")
                if col in COLMAPPING:
                    v1_fields_in_df["found"][COLMAPPING[col]].append(col + "_entry")
                    dfcols_tag[col + "_entry"] = True

    for col in V1_MODEL_FIELDS:
        found = False
        if col in dfcols:
            v1_fields_in_df["found"][col].append(col)
            found = True
            dfcols_tag[col] = True
        if col in dfcols_entry:
            v1_fields_in_df["found"][col].append(col + "_entry")
            found = True
            dfcols_tag[col + "_entry"] = True
        if not found:
            v1_fields_in_df["not_found"].append(col)
            dfcols_tag[col] = True

    v1_fields_in_df["untagged"] = [
        col for col, tagged in dfcols_tag.items() if not tagged
    ]

    print(v1_fields_in_df)

    return combined_df
