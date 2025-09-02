"""
Base & util functions for db-related operations
"""

import traceback
from datetime import datetime
from pathlib import Path

from loguru import logger
from sqlalchemy import Engine, create_engine
from tortoise import Model, Tortoise

from easy_access.db.models import CopyrightItem, Faculty, Classification, Status
from easy_access.settings import Settings

# Module-level flag to memoize initialization
_DB_INITIALIZED: bool = False


async def ensure_db_inited(settings: Settings | None = None) -> bool | None:
    """Ensure Tortoise ORM is initialized once per process.

    This is a thin wrapper around :func:`init` that memoizes the initialized
    state so callers don't have to remember to call ``await init(...)``.
    If the DB is already initialized this becomes a no-op.

    Args:
        settings: Optional Settings instance forwarded to ``init`` when
            initialization is required.
    """
    global _DB_INITIALIZED
    if _DB_INITIALIZED:
        return None
    # If callers pass None, let the underlying init raise if it needs settings
    if settings is None:
        res = await init(settings)  # type: ignore[arg-type]
    else:
        res = await init(settings=settings)
    _DB_INITIALIZED = True
    return res


def init_engine(settings: Settings) -> Engine:
    db_file_path = settings.db_path
    return create_engine(f"sqlite:///{str(db_file_path)}")


async def init(settings: Settings) -> bool | None:
    db_file_path = settings.db_path
    create_tables = False
    if not db_file_path.exists():
        create_tables = True

    # check if models.py file has been modified since last db modification
    models_py_path = Path("easy_access/db/models.py")
    if models_py_path.exists() and db_file_path.exists():
        models_py_mod_time = models_py_path.stat().st_mtime
        db_mod_time = db_file_path.stat().st_mtime
        if models_py_mod_time > db_mod_time:
            create_tables = True
    await Tortoise.init(
        db_url=f"sqlite://{str(db_file_path)}",
        modules={"models": ["easy_access.db.models"]},
    )
    await Tortoise.generate_schemas(safe=True)
    if create_tables:
        await Tortoise.generate_schemas(safe=True)
        return True


async def create() -> None:
    await Tortoise.generate_schemas(safe=True)




async def copyright_item_from_dict(
    item: dict[str, str | Model | int | datetime | None],
) -> CopyrightItem:
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
        "file_exists",
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
        logger.warning(f"Error getting faculty with {item.get('faculty')}: {e}")
    try:
        if not faculty:
            faculty = await Faculty.get(abbreviation="UNM")

        item["faculty"] = faculty
        # enforce sensible defaults for fields required by the model to avoid
        # immediate constructor errors when values are missing in staged data.
        if not item.get("classification"):
            # fallback to the project's default classification
            item["classification"] = Classification.LANGE_OVERNAME.value
        if not item.get("status"):
            item["status"] = Status.PUBLISHED.value

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

        if not item.get("file_exists"):
            item["file_exists"] = None

        final_dict = {}
        # Only include keys that are in the allowed set and have non-None values.
        # Passing explicit None for non-nullable fields causes model construction errors.
        for key, val in item.items():
            if key in copyright_item_keys and val is not None:
                final_dict[key] = val

        final_item = CopyrightItem(**final_dict)
        return final_item

    except Exception as e:
        logger.warning(
            f"Error while trying to create CopyrightItem with mat_id {item.get('material_id')}:{e}"
        )
        logger.error(traceback.format_exc())
        logger.debug(item)

        return None
