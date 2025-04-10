"""
Base & util functions for db-related operations
"""

import traceback
from datetime import datetime
from pathlib import Path

import polars as pl
from tortoise import Tortoise

from easy_access.db.models import CopyrightItem, Faculty
from easy_access.db.retrieve import init_engine
from easy_access.utils import File, warn

db_path: Path = Path("db.sqlite3")


def set_db_path(path: str | Path | File) -> None:
    global db_path
    if isinstance(path, File):
        path = path.path
    if isinstance(path, str):
        path = Path(path)
    if not isinstance(path, Path):
        raise ValueError("path must be a Path, str, or File object")
    db_path = path
    init_engine(str(db_path))


async def init() -> None:
    create_tables = False
    if not db_path.exists():
        create_tables = True

    # check if models.py file has been modified since last db modification
    models_py_path = Path("easy_access/db/models.py")
    if models_py_path.exists() and db_path.exists():
        models_py_mod_time = models_py_path.stat().st_mtime
        db_mod_time = db_path.stat().st_mtime
        if models_py_mod_time > db_mod_time:
            create_tables = True
    await Tortoise.init(
        db_url=f"sqlite://{str(db_path)}", modules={"models": ["easy_access.db.models"]}
    )
    await Tortoise.generate_schemas(safe=True)

    if create_tables:
        await Tortoise.generate_schemas(safe=True)
        return True


async def create() -> None:
    await Tortoise.generate_schemas(safe=True)


def standardize_dataframe(df: pl.DataFrame) -> pl.DataFrame:
    """
    rename cols to standard format
    cast all cols to str
    replace '-' with None
    filter missing material_ids
    filter to select only relevant itemtypes
    drop useless cols
    """
    df = (
        df.with_columns(pl.exclude(pl.String).cast(str))
        .rename(
            lambda col: col.replace(" ", "_")
            .replace("#", "count_")
            .replace("*", "x")
            .lower()
        )
        .with_columns(
            pl.when(pl.col(pl.String) != "-").then(pl.col(pl.String)).name.keep()
        )
        .filter(pl.col("material_id").is_not_null())
    )

    if "filetype" in df.columns:
        df = df.filter(
            (pl.col("filetype").is_in(["pdf", "ppt", "doc", "-"]))
            | (pl.col("filetype").is_null())
        )
    if "type" in df.columns:
        df = df.drop("type")
    if "google_search_file" in df.columns:
        df = df.drop("google_search_file")
    return df


async def copyright_item_from_dict(item: dict[str, str]) -> CopyrightItem:
    """
    Turns a dict with data for a CopyrightItem into a CopyrightItem object
    """
    copyright_item_keys = {
        "material_id",
        "period",
        "department",
        "course_code",
        "course_name",
        "url",
        "filename",
        "title",
        "owner",
        "filetype",
        "classification",
        "ml_prediction",
        "manual_classification",
        "manual_identifier",
        "scope",
        "remarks",
        "auditor",
        "last_change",
        "status",
        "isbn",
        "doi",
        "in_collection",
        "pagecount",
        "wordcount",
        "picturecount",
        "author",
        "publisher",
        "reliability",
        "pages_x_students",
        "count_students_registered",
        "retrieved_from_copyright_on",
        "workflow_status",
        "possible_fine",
        "infringement",
        "faculty",
    }
    try:
        if item.get("faculty") == "Unmapped" or not item.get("faculty"):
            abbr = "UNM"
        else:
            abbr = item.get("faculty")
            if not abbr:
                abbr = "UNM"
        faculty = await Faculty.get(abbreviation=abbr)
    except Exception as e:
        warn(f"Error getting faculty with {item.get('faculty')}: {e}")
    try:
        if not faculty:
            faculty = await Faculty.get(abbreviation="UNM")

        item["faculty"] = faculty
        item["material_id"] = int(item.get("material_id"))
        item["last_change"] = (
            datetime.strptime(item.get("last_change", ""), "%Y-%m-%d")
            if isinstance(item.get("last_change"), str)
            else None
        )

        item["retrieved_from_copyright_on"] = (
            (
                datetime.strptime(
                    item["retrieved_from_copyright_on"].split(" ")[0], "%Y-%m-%d"
                )
            )
            if item.get("retrieved_from_copyright_on")
            else None
        )

        item["pagecount"] = int(item.get("pagecount")) if item.get("pagecount") else 0
        item["wordcount"] = int(item.get("wordcount")) if item.get("wordcount") else 0
        item["picturecount"] = (
            int(item.get("picturecount")) if item.get("picturecount") else 0
        )
        item["reliability"] = (
            int(item.get("reliability")) if item.get("reliability") else 0
        )
        item["pages_x_students"] = (
            int(item.get("pages_x_students")) if item.get("pages_x_students") else 0
        )
        item["count_students_registered"] = (
            int(item.get("count_students_registered"))
            if item.get("count_students_registered")
            else 0
        )
        item["filetype"] = (
            item.get("filetype", "unknown") if item.get("filetype") else "unknown"
        )

        if not item.get("department"):
            item["department"] = item.get("programme_canvas")
        if not item.get("course_name"):
            item["course_name"] = item.get("course_name_canvas")
        final_dict = {}
        for key in item:
            if key in copyright_item_keys:
                final_dict[key] = item[key]

        final_item = CopyrightItem(**final_dict)
        return final_item

    except Exception as e:
        warn(
            f"Error while trying to create CopyrightItem with mat_id {item['material_id']}:{e}"
        )
        print(traceback.format_exc())
        print(item)

        return None
