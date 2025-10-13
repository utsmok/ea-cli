"""
Base & util functions for db-related operations
"""

import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

from loguru import logger
from sqlalchemy import Engine, create_engine

from easy_access.db.compat import get_or_create
from easy_access.db.enums import Classification, Status
from easy_access.db.models import (
    CopyrightItem,
    Faculty,
    Organization,
)
from easy_access.db.session import get_engine, get_session, init_db, shutdown_db
from easy_access.settings import Settings

# Module-level flag to memoize initialization
_DB_INITIALIZED: bool = False


async def ensure_db_inited(settings: Settings | None = None) -> bool | None:
    """Ensure SQLAlchemy async engine is initialized once per process.

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
    """Create a synchronous SQLAlchemy engine for polars/pandas reads.

    This is used by retrieve.py for read-heavy operations with polars.
    For ORM operations, use the async session from session.py instead.
    """
    db_file_path = settings.db_path
    return create_engine(f"sqlite:///{str(db_file_path)}")


async def init(settings: Settings) -> bool | None:
    """Initialize the async database connection and create tables if needed.

    Args:
        settings: Application settings containing DB configuration

    Returns:
        True if tables were created, None otherwise
    """

    from easy_access.db.models_base import Base

    db_file_path = settings.db_path
    create_tables = False

    if not db_file_path.exists():
        create_tables = True

    # Check if models have been modified since last DB modification
    models_py_path = Path("easy_access/db/sa_models.py")
    if models_py_path.exists() and db_file_path.exists():
        models_py_mod_time = models_py_path.stat().st_mtime
        db_mod_time = db_file_path.stat().st_mtime
        if models_py_mod_time > db_mod_time:
            create_tables = True

    # Initialize the async engine
    init_db(settings)

    if create_tables:
        # Create all tables using the engine directly
        engine = get_engine()
        if engine:
            async with engine.begin() as conn:
                await conn.run_sync(Base.metadata.create_all)

        # Also create the default faculties
        await init_faculties(settings)
        return True

    return None


async def init_faculties(settings: Settings) -> None:
    """Initialize default faculties in the database.

    Args:
        settings: Application settings containing university and faculty config
    """
    # Create university entry in 'Organization' table
    main_uni, success = await get_or_create(
        Organization,
        name=settings.university_settings.name,
        defaults={
            "abbreviation": settings.university_settings.abbreviation,
            "full_abbreviation": settings.university_settings.abbreviation,
            "hierarchy_level": 0,
        },
    )

    # Retrieve faculties from settings
    faculties = settings.university_settings.faculties

    for faculty in faculties:
        faculty_obj, success = await get_or_create(
            Faculty,
            name=faculty.name,
            defaults={
                "abbreviation": faculty.abbreviation,
                "full_abbreviation": faculty.abbreviation,
                "hierarchy_level": 1,
            },
        )

        # Set parent organization if not already set
        if not faculty_obj.parent_organization:
            faculty_obj.parent_organization = main_uni
            async for session in get_session():
                async with session.begin():
                    session.add(faculty_obj)
                    await session.flush()

    # Also create an "Unmapped" faculty to use as a fallback
    unmapped, success = await get_or_create(
        Faculty,
        name="Unmapped",
        defaults={
            "abbreviation": "UNM",
            "full_abbreviation": "UNM",
            "hierarchy_level": 1,
        },
    )

    if not unmapped.parent_organization:
        unmapped.parent_organization = main_uni
        async for session in get_session():
            async with session.begin():
                session.add(unmapped)
                await session.flush()


async def create() -> None:
    """Create all database tables (deprecated - use init instead)."""
    from easy_access.db.models_base import Base

    async for session in get_session():
        async with session.begin():
            await session.run_sync(Base.metadata.create_all)


async def close_connections() -> None:
    """Close all database connections."""
    await shutdown_db()


async def copyright_item_from_dict(
    item: dict[str, Any],
) -> CopyrightItem | None:
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
        faculty = None  # Initialize faculty to None on error

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

        item["material_id"] = (
            int(item.get("material_id"))
            if isinstance(item.get("material_id"), str | int)
            else 0
        )  # type: ignore[arg-type]

        # Validate that material_id is present and valid
        if not item.get("material_id") or item["material_id"] == 0:
            logger.warning(f"Invalid or missing material_id: {item.get('material_id')}")
            return None
        last_change_val = item.get("last_change")
        item["last_change"] = (
            datetime.strptime(last_change_val, "%Y-%m-%d")
            if isinstance(last_change_val, str) and last_change_val
            else None
        )

        retrieved_val = item.get("retrieved_from_copyright_on")
        item["retrieved_from_copyright_on"] = (
            datetime.strptime(retrieved_val.split(" ")[0], "%Y-%m-%d")
            if isinstance(retrieved_val, str) and retrieved_val
            else None
        )

        item["pagecount"] = (
            int(item.get("pagecount"))
            if isinstance(item.get("pagecount"), str | int)
            else 0
        )  # type: ignore[arg-type]
        item["wordcount"] = (
            int(item.get("wordcount"))
            if isinstance(item.get("wordcount"), str | int)
            else 0
        )  # type: ignore[arg-type]
        item["picturecount"] = (
            int(item.get("picturecount"))
            if isinstance(item.get("picturecount"), str | int)
            else 0  # type: ignore[arg-type]
        )
        item["reliability"] = (
            int(item.get("reliability"))
            if isinstance(item.get("reliability"), str | int)
            else 0  # type: ignore[arg-type]
        )
        item["pages_x_students"] = (
            int(item.get("pages_x_students"))
            if isinstance(item.get("pages_x_students"), str | int)
            else 0  # type: ignore[arg-type]
        )
        item["count_students_registered"] = (
            int(item.get("count_students_registered"))
            if isinstance(item.get("count_students_registered"), str | int)
            else 0  # type: ignore[arg-type]
        )
        item["filetype"] = (
            item.get("filetype", "unknown") if item.get("filetype") else "unknown"
        )

        if not item.get("department"):
            item["department"] = item.get("programme_canvas")
        if not item.get("course_name"):
            item["course_name"] = item.get("course_name_canvas")

        # Normalize file_exists before checking if it's falsy
        from easy_access.db.update import _normalize_file_exists

        original_file_exists = item.get("file_exists")

        # Only set to None if the original value was None or empty string
        # After normalization, we want to preserve True/False values
        if original_file_exists is None or original_file_exists == "":
            item["file_exists"] = None
        else:
            normalized_file_exists = _normalize_file_exists(original_file_exists)
            item["file_exists"] = normalized_file_exists

        final_dict = {}
        # Only include keys that are in the allowed set and have non-None values.
        # Passing explicit None for non-nullable fields causes model construction errors.
        # However, file_exists should always be included even if None (null=True in model)
        for key, val in item.items():
            if key in copyright_item_keys and (key == "file_exists" or val is not None):
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
