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

from collections import defaultdict
from pathlib import Path

import polars as pl
from rich import print

from easy_access.db.enums import (
    Classification,
    Filetype,
    Period,
    Status,
    WorkflowStatus,
)
from easy_access.db.models import v1_CopyrightItem

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


def import_v1_sheets(
    path: Path,
) -> dict[str, dict[str, pl.DataFrame]]:
    """
    Reads in data from the v1 sheets from the given path, and returns a dict structure with dfs for each of the found sheets.
    """
    all_files = list(path.glob("*.xlsx"))
    print(f"Found {len(all_files)} excel files in {path}")
    sheets = {}
    weekly = 0
    for file in all_files:
        try:
            sheets[file.name] = {
                "complete_data": pl.read_excel(file, sheet_name="Complete data"),
                "data_entry": pl.read_excel(file, sheet_name="Data entry"),
            }
        except Exception as e:
            print(f"Error reading {file.name}: {e}")
            continue
        if "overview" in file.name:
            print(f"found overview sheet: {file.name}")

        else:
            weekly += 1
    print(f"Found {weekly} weekly sheets in total.")

    return sheets


def clean_and_cast_cols(
    df: pl.DataFrame, conflict_cols: list[tuple[str, str]]
) -> pl.DataFrame:
    """
    Cleans and casts columns in the dataframe to appropriate types.
    Handles conflict columns by prioritizing non-null values from the first column.
    """
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
            pl.col(col).cast(dtype)
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
                print(f"Error parsing datetime column {col}: {e}")
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
                print(f"Error parsing date column {col}: {e}")
    return df


def process_v1_sheet(filename: str, dfs: dict[str, pl.DataFrame]) -> pl.DataFrame:
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
        print(f"[{filename}] Error merging dataframes: {e}")
        return pl.DataFrame()

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
        print(f"Details for {filename}:")
        print_details(combined_df)
    return combined_df


def process_v1_sheets(
    extracted_sheets: dict[str, dict[str, pl.DataFrame]],
) -> dict[str, pl.DataFrame]:
    """
    input: dict of sheets as returned by import_v1_sheets()

    processes each set of sheets using process_v1_sheet(): merge, clean, parse, etc
    returns a dict with the processed dataframes, keyed by filename.
    """

    return {
        filename: process_v1_sheet(filename, dfs)
        for filename, dfs in extracted_sheets.items()
    }


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
                print(f"Error merging {filename}: {e}")
                continue
            try:
                conflict_cols = [
                    (col, col.replace("_entry", ""))
                    for col in main_df.columns
                    if col.endswith("_entry")
                ]
                main_df = clean_and_cast_cols(main_df, conflict_cols)
            except Exception as e:
                print(f"Error cleaning columns in {filename}: {e}")

    print_details(main_df)
    return main_df


#
#   Database ingestion functions
#


async def ingest_v1_items(df: pl.DataFrame) -> None:
    """
    ingests the given dataframe into the v1_CopyrightItem model in the database
    """
    items: list[v1_CopyrightItem] = []
    data = df.to_dicts()

    for row in data:
        await v1_CopyrightItem.create(**row)


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
