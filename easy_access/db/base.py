"""
This module provides base functionalities for database operations, including
database initialization for Tortoise ORM, schema generation, and utility
functions for standardizing data and creating ORM model instances from dictionaries.

It defines the database path and includes functions to set this path dynamically.
The primary ORM used is Tortoise ORM, with `init()` being the main entry point
for its setup. An SQLAlchemy engine (`init_engine`) is also included, potentially
for use with Polars' direct database reading capabilities.
"""

import logging
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any # For dict values

import polars as pl
from sqlalchemy import Engine, create_engine # Engine for type hint
from tortoise import Tortoise

from easy_access.db.models import CopyrightItem, Faculty
from easy_access.utils import File

db_path: Path = Path("db.sqlite3") # Default database path
logger = logging.getLogger(__name__)

# Module-level constant for CopyrightItem keys used in from_dict conversion
# Ensures consistency and avoids redefining this large set in the function.
# These should match the fields in the CopyrightItem model or be expected keys in input dicts.
COPYRIGHT_ITEM_EXPECTED_KEYS: set[str] = {
    "material_id", "period", "department", "course_code", "course_name", "url",
    "filename", "title", "owner", "filetype", "classification", "ml_prediction",
    "manual_classification", "manual_identifier", "scope", "remarks", "auditor",
    "last_change", "status", "isbn", "doi", "in_collection", "pagecount",
    "wordcount", "picturecount", "author", "publisher", "reliability",
    "pages_x_students", "count_students_registered", "retrieved_from_copyright_on",
    "workflow_status", "possible_fine", "infringement", "faculty",
    # Potentially add keys from canvas if they are used before renaming, e.g. programme_canvas
    "programme_canvas", "course_name_canvas",
}


def init_engine(path_str: str | None = None) -> Engine:
    """
    Initializes and returns an SQLAlchemy engine instance.
    This engine can be used by Polars for direct database reads.

    Args:
        path_str (str | None, optional): The database file name or full path.
                                      Defaults to "db.sqlite3".

    Returns:
        Engine: The initialized SQLAlchemy engine.
    """
    if not path_str:
        path_str = "db.sqlite3"
    return create_engine(f"sqlite:///{path_str}")


def set_db_path(path: str | Path | File) -> None:
    """
    Sets the global `db_path` for the database and re-initializes the
    SQLAlchemy engine (if it's being used globally by other modules like retrieve.py).

    Args:
        path (str | Path | File): The new path to the database file.
                                  Can be a string, pathlib.Path, or a utils.File object.

    Raises:
        ValueError: If the provided path is not a str, Path, or File object.
    """
    global db_path, engine # Allow modification of global engine if retrieve.py relies on it

    resolved_path: Path
    if isinstance(path, File):
        resolved_path = path.path
    elif isinstance(path, str):
        resolved_path = Path(path)
    elif isinstance(path, Path):
        resolved_path = path
    else:
        raise ValueError("path must be a Path, str, or File object")

    db_path = resolved_path.resolve() # Ensure it's an absolute path
    logger.info(f"Global database path set to: {db_path}")

    # Re-initialize the global engine in retrieve.py if it's used there.
    # This is a bit of a side effect. Ideally, modules requiring an engine
    # would get it explicitly.
    try:
        # Attempt to update retrieve.py's engine if it exists.
        # This is fragile and not recommended.
        import easy_access.db.retrieve
        easy_access.db.retrieve.engine = init_engine(str(db_path))
        logger.debug(f"SQLAlchemy engine in db.retrieve re-initialized with path: {db_path}")
    except ImportError:
        logger.debug("db.retrieve module not imported, or its engine not updated via set_db_path.")
    except Exception as e:
        logger.warning(f"Could not re-initialize engine in db.retrieve: {e}")



async def init(force_create_tables: bool = False) -> bool:
    """
    Initializes the Tortoise ORM connection and generates database schemas.
    Schemas are generated if the database file doesn't exist or if the
    `models.py` file has been modified more recently than the database file,
    or if `force_create_tables` is True.

    Args:
        force_create_tables (bool, optional): If True, forces schema generation
                                              even if heuristics suggest it's not needed.
                                              Defaults to False.
    Returns:
        bool: True if tables were (re)generated, False otherwise.
    """
    create_tables_flag: bool = force_create_tables
    if not db_path.exists():
        logger.info(f"Database file {db_path} not found. Schemas will be created.")
        create_tables_flag = True

    if not create_tables_flag: # Only check model file if not already forcing creation
        models_py_path = Path(__file__).parent / "models.py" # More robust path to models.py
        if models_py_path.exists() and db_path.exists(): # db_path must exist for stat
            try:
                models_py_mod_time = models_py_path.stat().st_mtime
                db_mod_time = db_path.stat().st_mtime
                if models_py_mod_time > db_mod_time:
                    logger.info("models.py is newer than the database file. Schemas will be regenerated.")
                    create_tables_flag = True
            except FileNotFoundError: # Should not happen due to .exists() checks
                 logger.warning("Could not stat models.py or db_path for schema generation check.")


    await Tortoise.init(
        db_url=f"sqlite://{str(db_path)}",
        modules={"models": ["easy_access.db.models"]}
    )

    if create_tables_flag:
        logger.info("Generating database schemas.")
        await Tortoise.generate_schemas(safe=True) # safe=True avoids dropping columns not in models
        return True
    else:
        # Ensure schemas are present even if not creating anew (e.g. if DB exists but tables are missing)
        # Tortoise.generate_schemas(safe=True) is idempotent for existing tables matching models.
        await Tortoise.generate_schemas(safe=True)
        logger.debug("Database schemas verified (safe generation).")
        return False


async def create() -> None:
    """
    Explicitly generates database schemas using Tortoise ORM.
    This is typically called by `init()` if tables need to be created.
    """
    logger.info("Explicitly called create_db_and_tables(). Generating schemas.")
    await Tortoise.generate_schemas(safe=True)


def standardize_dataframe(df: pl.DataFrame) -> pl.DataFrame:
    """
    Standardizes a DataFrame containing copyright item data.
    Operations include:
    - Renaming columns to a standard format (lowercase, underscores).
    - Casting all columns to string type initially for robust processing.
    - Replacing placeholder '-' values with None (though Polars might handle this as nulls).
    - Filtering out rows with missing `material_id`.
    - Optionally filtering by `filetype` if the column exists.
    - Dropping specified irrelevant columns ('type', 'google_search_file').

    Args:
        df (pl.DataFrame): The input DataFrame.

    Returns:
        pl.DataFrame: The standardized DataFrame.
    """
    if df.is_empty():
        return df

    # Standardize column names
    renamed_df = df.rename(
        lambda col_name: str(col_name).replace(" ", "_").replace("#", "count_").replace("*", "x").lower()
    )

    # Cast all to string first for safety, then handle specific type conversions later if needed
    # This also helps with '-' replacement if they are not already nulls.
    string_casted_df = renamed_df.with_columns(pl.all().cast(pl.Utf8, strict=False))

    # Replace '-' with null, then filter null material_ids
    # Polars usually reads empty strings or specific markers like '-' as nulls or empty strings depending on source.
    # Explicit replacement can be done if needed, but often pl.col().is_not_null() is sufficient.
    # For now, assuming material_id is critical and must exist.
    # Example for replacing '-':
    # final_df = string_casted_df.with_columns(
    #     [pl.when(pl.col(c) == "-").then(None).otherwise(pl.col(c)).name.keep() for c in string_casted_df.columns]
    # )
    final_df = string_casted_df.filter(pl.col("material_id").is_not_null() & (pl.col("material_id") != ""))


    if "filetype" in final_df.columns:
        final_df = final_df.filter(
            (pl.col("filetype").str.to_lowercase().is_in(["pdf", "ppt", "doc", "pptx", "docx", "-"])) | # Added pptx, docx
            (pl.col("filetype").is_null()) |
            (pl.col("filetype") == "") # Consider empty string as valid or to be mapped to unknown
        )

    # Drop columns if they exist
    cols_to_drop = ["type", "google_search_file"]
    for col_to_drop in cols_to_drop:
        if col_to_drop in final_df.columns:
            final_df = final_df.drop(col_to_drop)

    return final_df


async def copyright_item_from_dict(item_data: dict[str, Any]) -> CopyrightItem | None:
    """
    Converts a dictionary of copyright item data into a `CopyrightItem` ORM object.
    Performs type conversions and links to the Faculty model.

    Args:
        item_data (dict[str, Any]): A dictionary where keys are field names and
                                   values are the corresponding data for a copyright item.

    Returns:
        CopyrightItem | None: An initialized `CopyrightItem` model instance, or None if
                              conversion fails or essential data (like faculty) is missing.
    """
    faculty_abbr: str = item_data.get("faculty", "UNM") # Default to "UNM"
    if faculty_abbr == "Unmapped" or not faculty_abbr: # Handle "Unmapped" or empty
        faculty_abbr = "UNM"

    faculty: Faculty | None = None
    try:
        faculty = await Faculty.get_or_none(abbreviation=faculty_abbr)
        if not faculty:
            logger.warning(f"Faculty with abbreviation '{faculty_abbr}' not found. Trying 'UNM'.")
            faculty = await Faculty.get_or_none(abbreviation="UNM") # Fallback
            if not faculty: # Still not found, this is an issue for FK constraint
                 logger.error(f"Fallback faculty 'UNM' also not found. Cannot create CopyrightItem for material_id {item_data.get('material_id')}")
                 return None
    except Exception as e:
        logger.error(f"Error retrieving faculty '{faculty_abbr}' for material_id {item_data.get('material_id', 'Unknown')}: {e}")
        return None # Cannot proceed without a valid faculty if it's a required relation

    # Prepare data for CopyrightItem model instantiation
    model_data: dict[str, Any] = {}

    for key, value in item_data.items():
        if key not in COPYRIGHT_ITEM_EXPECTED_KEYS: # Filter to expected keys
            continue

        if key == "material_id":
            model_data[key] = int(value) if value is not None else None
        elif key == "last_change" or key == "retrieved_from_copyright_on":
            if isinstance(value, str):
                try:
                    # Attempt to parse various date/datetime formats that might appear
                    if " " in value: # Likely datetime
                        model_data[key] = datetime.strptime(value.split(" ")[0], "%Y-%m-%d")
                    else: # Likely date
                        model_data[key] = datetime.strptime(value, "%Y-%m-%d")
                except ValueError:
                    model_data[key] = None # Set to None if parsing fails
                    logger.debug(f"Could not parse date string '{value}' for field '{key}'.")
            elif isinstance(value, datetime):
                 model_data[key] = value
            else:
                model_data[key] = None
        elif key in ["pagecount", "wordcount", "picturecount", "reliability", "pages_x_students", "count_students_registered"]:
            model_data[key] = int(value) if value is not None and str(value).isdigit() else 0
        elif key == "faculty":
            model_data[key] = faculty # Assign the fetched Faculty object
        elif key == "filetype":
            model_data[key] = value if value else Filetype.UNKNOWN.value # Default if empty
        elif key == "classification":
             model_data[key] = value if value else Classification.LANGE_OVERNAME.value
        elif key == "status":
            model_data[key] = value if value else Status.PUBLISHED.value
        elif key == "workflow_status":
            model_data[key] = value if value else WorkflowStatus.ToDo.value
        elif key == "infringement":
            model_data[key] = value if value else Infringement.UNDETERMINED.value
        else:
            model_data[key] = value if value is not None else None # Handle other fields, ensure None for empty

    # Handle cases where department/course_name might come from alternative keys
    if not model_data.get("department") and item_data.get("programme_canvas"):
        model_data["department"] = item_data.get("programme_canvas")
    if not model_data.get("course_name") and item_data.get("course_name_canvas"):
        model_data["course_name"] = item_data.get("course_name_canvas")

    # Ensure all required fields for the model are present, even if None
    # Tortoise fields handle null=True, but good to be explicit for non-nullable if not defaulted in model
    if model_data.get("material_id") is None:
        logger.warning(f"material_id is None for item: {item_data}. Skipping creation.")
        return None


    try:
        # Filter final_dict to only include keys that are actual fields in CopyrightItem model
        # This is safer than relying on COPYRIGHT_ITEM_EXPECTED_KEYS directly if it's not perfectly synced
        # However, Tortoise ORM handles extra keys in **kwargs by ignoring them.
        final_item = CopyrightItem(**model_data)
        return final_item
    except Exception as e:
        logger.error(f"Error creating CopyrightItem instance for material_id {model_data.get('material_id', 'Unknown')}: {e}")
        logger.debug(f"Data used for CopyrightItem creation: {model_data}")
        logger.debug(traceback.format_exc())
        return None

[end of easy_access/db/base.py]
