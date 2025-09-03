"""
functions to update existing data in the database
"""

import contextlib
import traceback
from datetime import UTC, date, datetime
from enum import Enum, StrEnum
from itertools import batched
from typing import Any

import polars as pl
from loguru import logger
from tortoise import Tortoise
from tortoise.transactions import in_transaction

from easy_access.db.base import (
    copyright_item_from_dict,
    ensure_db_inited,
)
from easy_access.db.models import (
    PDF,
    CopyrightItem,
    Course,
    Infringement,
    ItemUpdate,
    StagedCopyrightItem,
    StagedFacultyUpdate,
    StagedProcessingFailure,
    Status,
    WorkflowStatus,
)
from easy_access.merge_rules import (
    build_merge_rules_from_settings,
    get_mergeable_fields,
)
from easy_access.settings import Settings
from easy_access.utils import (
    determine_course_code,
    safe_compare_greater,
    safe_enum,
    safe_float,
    safe_int,
    standardize_dataframe,
)


# Custom exceptions for better error handling
class MergeError(Exception):
    """Base exception for merge-related errors."""

    pass


class MergeConflictError(MergeError):
    """Raised when there are conflicts during field merging."""

    pass


class TypeCastError(MergeError):
    """Raised when type casting fails during field comparison."""

    pass


class DatabaseOperationError(MergeError):
    """Raised when database operations fail."""

    pass


class ValidationError(MergeError):
    """Raised when data validation fails."""

    pass


# Constants for comparison logic
DEFAULT_RANK = 20
MIN_CHANGES_THRESHOLD = 3


class FieldComparisonStrategy:
    """Base class for field comparison strategies."""

    def should_update(
        self, new_value: Any, old_value: Any, ordering: Any
    ) -> tuple[bool, str]:
        """
        Determine if a field should be updated.

        Args:
            new_value: New value for the field
            old_value: Current value in the database
            ordering: Ordering rules for the field

        Returns:
            Tuple of (should_update, reason)
        """
        raise NotImplementedError


class RankedFieldStrategy(FieldComparisonStrategy):
    """Strategy for ranked fields (higher priority = lower index)."""

    def should_update(
        self, new_value: Any, old_value: Any, ordering: list
    ) -> tuple[bool, str]:
        if not isinstance(ordering, list):
            return False, ""

        new_rank = DEFAULT_RANK
        old_rank = DEFAULT_RANK

        if new_value in ordering:
            new_rank = ordering.index(new_value)
        if old_value in ordering:
            old_rank = ordering.index(old_value)

        if new_rank < old_rank:
            return True, "new rank < old rank"

        return False, ""


class StringFieldStrategy(FieldComparisonStrategy):
    """Strategy for string fields (longer strings take precedence)."""

    def should_update(
        self, new_value: Any, old_value: Any, ordering: Any
    ) -> tuple[bool, str]:
        if not (isinstance(new_value, str) and isinstance(old_value, str)):
            return False, ""

        new_value = new_value.strip()
        old_value = old_value.strip()

        if len(new_value) > len(old_value):
            return True, "new len > old len"

        return False, ""


class NumericFieldStrategy(FieldComparisonStrategy):
    """Strategy for numeric/date fields using safe comparison."""

    def should_update(
        self, new_value: Any, old_value: Any, ordering: Any
    ) -> tuple[bool, str]:
        try:
            if safe_compare_greater(new_value, old_value):
                return True, "new > old"
        except Exception:
            logger.debug(f"Could not compare values: {new_value} vs {old_value}")

        return False, ""


class DateFieldStrategy(FieldComparisonStrategy):
    """Strategy for date/datetime fields (newer dates take precedence)."""

    def should_update(
        self, new_value: Any, old_value: Any, ordering: Any
    ) -> tuple[bool, str]:
        if not (
            isinstance(new_value, date | datetime)
            and isinstance(old_value, date | datetime)
        ):
            return False, ""

        if new_value > old_value:
            return True, "new date > old date"

        return False, ""


class EnumFieldStrategy(FieldComparisonStrategy):
    """Strategy for enum fields (uses ranking if provided, otherwise no update)."""

    def should_update(
        self, new_value: Any, old_value: Any, ordering: Any
    ) -> tuple[bool, str]:
        # If ordering is provided, use ranked comparison
        if isinstance(ordering, list) and ordering:
            new_rank = DEFAULT_RANK
            old_rank = DEFAULT_RANK

            if new_value in ordering:
                new_rank = ordering.index(new_value)
            if old_value in ordering:
                old_rank = ordering.index(old_value)

            if new_rank < old_rank:
                return True, "new enum rank < old enum rank"

        return False, ""


class FileExistsStrategy(FieldComparisonStrategy):
    """Strategy for file_exists field (always update when received)."""

    def should_update(
        self, new_value: Any, old_value: Any, ordering: Any
    ) -> tuple[bool, str]:
        return True, "file_exists value received, always update"


def get_comparison_strategy(
    field: str, db_item: CopyrightItem | None = None
) -> FieldComparisonStrategy:
    """
    Get the appropriate comparison strategy for a field.

    Args:
        field: Field name
        db_item: Database item to check field types (optional)

    Returns:
        FieldComparisonStrategy instance
    """
    # Special case for file_exists
    if field == "file_exists":
        return FileExistsStrategy()

    # If we have a db_item, check the field type to determine strategy
    if db_item is not None:
        try:
            old_value = getattr(db_item, field)
            if isinstance(old_value, date | datetime):
                return DateFieldStrategy()
            elif isinstance(old_value, Enum):
                return EnumFieldStrategy()
            elif isinstance(old_value, str):
                return StringFieldStrategy()
        except AttributeError:
            pass

    # Default strategy for numeric fields
    return NumericFieldStrategy()


async def preprocess_input_data(
    data: pl.DataFrame | list[dict],
) -> tuple[list[dict], list[dict]]:
    """
    Preprocess input data by standardizing and separating new items from existing items.

    Args:
        data: Input data as DataFrame or list of dicts

    Returns:
        Tuple of (new_items, update_items)
    """
    new_items = []
    update_items = []

    if isinstance(data, pl.DataFrame):
        data = standardize_dataframe(data)
        existing_mat_ids = await CopyrightItem.all().values("material_id")
        existing_mat_ids = {safe_int(m["material_id"]) for m in existing_mat_ids}
        existing_mat_ids = {m for m in existing_mat_ids if m is not None}

        # Candidate new items (may be partial if coming from faculty sheets)
        candidate_new_items = (
            data.with_columns(pl.col("material_id").cast(int))
            .filter(~pl.col("material_id").is_in(existing_mat_ids))
            .to_dicts()
        )

        # Only create new CopyrightItem objects when the incoming row contains
        # the required, non-nullable fields present in the model.
        required_for_creation = [
            "period",
            "department",
            "course_code",
            "course_name",
        ]

        skipped_mat_ids: list[Any] = []

        for itm in candidate_new_items:
            ok = True
            for rc in required_for_creation:
                v = itm.get(rc)
                if v is None or (isinstance(v, str) and v.strip() == ""):
                    ok = False
                    break
            if ok:
                new_items.append(itm)
            else:
                mid_val = itm.get("material_id")
                if mid_val is None:
                    skipped_mat_ids.append("<no-id>")
                else:
                    safe_mid = safe_int(mid_val)
                    if safe_mid is not None:
                        skipped_mat_ids.append(safe_mid)
                    else:
                        skipped_mat_ids.append(str(mid_val))

        if skipped_mat_ids:
            logger.warning(
                f"Skipping {len(skipped_mat_ids)} new items missing required fields (not creating in DB): {skipped_mat_ids[:20]}"
            )

        update_items = (
            data.with_columns(pl.col("material_id").cast(int))
            .filter(pl.col("material_id").is_in(existing_mat_ids))
            .to_dicts()
        )

    elif isinstance(data, list):
        update_items = data

    logger.info(f"# of new items: {len(new_items)}")
    logger.info(f"# of items to update: {len(update_items)}")

    return new_items, update_items


async def process_new_items(new_items: list[dict]) -> list[CopyrightItem]:
    """
    Process and create new copyright items in the database.

    Args:
        new_items: List of dictionaries representing new items to create

    Returns:
        List of created CopyrightItem objects

    Raises:
        DatabaseOperationError: When bulk creation fails
    """
    new_objects = []
    if new_items:
        new_objects = [await copyright_item_from_dict(item) for item in new_items]
        new_objects = [item for item in new_objects if item]
        try:
            await CopyrightItem.bulk_create(objects=new_objects)
        except Exception as e:
            logger.warning(
                f"Bulk creation failed: {e}. Attempting one-by-one creation."
            )
            failed_items = []
            for item in new_objects:
                try:
                    await item.save()
                except Exception as save_error:
                    logger.error(
                        f"Failed to save item {item.material_id}: {save_error}"
                    )
                    failed_items.append(item.material_id)

            if failed_items:
                raise DatabaseOperationError(
                    f"Failed to create items with material_ids: {failed_items}"
                ) from e
        logger.success(f"Created {len(new_objects)} new copyright items in db.")

    return new_objects


async def process_existing_items(
    update_items: list[dict],
    added_fields: dict,
    changeable_fields: dict,
    overwrite: bool = False,
) -> tuple[list[CopyrightItem], dict]:
    """
    Process updates to existing copyright items.

    Args:
        update_items: List of dictionaries representing items to update
        added_fields: Dictionary of field priorities for added fields
        changeable_fields: Dictionary of field priorities for changeable fields
        overwrite: Whether to overwrite existing values instead of using comparison logic

    Returns:
        Tuple of (changelist, updates) where changelist contains modified items
        and updates contains change details

    Raises:
        DatabaseOperationError: When processing individual items fails
    """
    logger.info(f"Updating {len(update_items)} existing items.")
    changelist = []
    updates = {}

    for new_item in update_items:
        try:
            if overwrite:
                changes, db_item = await _process_item_overwrite(
                    new_item, added_fields, changeable_fields
                )
            else:
                changes, db_item = await _process_item_normal(
                    new_item, added_fields, changeable_fields
                )

            if len(list(changes.keys())) >= MIN_CHANGES_THRESHOLD:
                changes["modified_at"] = datetime.now()
                updates[new_item.get("material_id")] = changes
                changelist.append(db_item)
        except DatabaseOperationError:
            # Re-raise database operation errors
            raise
        except Exception as e:
            # Log unexpected errors but continue processing other items
            logger.error(
                f"Unexpected error processing item {new_item.get('material_id')}: {e}"
            )
            continue

    return changelist, updates


async def _process_item_overwrite(
    new_item: dict, added_fields: dict, changeable_fields: dict
) -> tuple[dict, CopyrightItem | None]:
    """
    Process a single item in overwrite mode.

    Args:
        new_item: Dictionary with new field values
        added_fields: Dictionary of field priorities for added fields
        changeable_fields: Dictionary of field priorities for changeable fields

    Returns:
        Tuple of (changes dict, db_item)

    Raises:
        DatabaseOperationError: When database operations fail
    """
    try:
        db_item = await CopyrightItem.get(material_id=new_item.get("material_id"))

        changes = {
            "material_id": new_item.get("material_id"),
            "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

        logger.debug(f"now in overwrite function for {new_item.get('material_id')}")
        for k in changeable_fields | added_fields:
            logger.debug(f"checking field {k}")
            logger.debug(f"new_item.get(k): {new_item.get(k)}")
            logger.debug(f"getattr(db_item, k): {getattr(db_item, k)}")
            if new_item.get(k) is None:
                continue
            if new_item.get(k) != getattr(db_item, k):
                if str(new_item.get(k)) == getattr(db_item, k):
                    # if the new value is the same as the old value, skip it
                    continue
                changes = record_field_change(
                    changes,
                    k,
                    new_item.get(k),
                    getattr(db_item, k),
                    "[overwrite] new value != old value",
                    db_item,
                )
                logger.debug(f"changes: {changes}")
        # if any changes were made we'll have 3 or more keys in the changes dict
        # if not, no need to update the db
        logger.debug("final changes:")
        logger.debug(changes)
        if len(changes) < MIN_CHANGES_THRESHOLD:
            logger.debug(f"No changes for item {new_item.get('material_id')}.")
    except Exception as e:
        logger.warning(f"Could not update item {new_item.get('material_id')}: {e}")
        logger.warning(traceback.format_exc())
        raise DatabaseOperationError(
            f"Failed to process item {new_item.get('material_id')} in overwrite mode: {e}"
        ) from e

    return changes, db_item


async def _process_item_normal(
    new_item: dict, added_fields: dict, changeable_fields: dict
) -> tuple[dict, CopyrightItem | None]:
    """
    Process a single item in normal mode using comparison logic.

    Args:
        new_item: Dictionary with new field values
        added_fields: Dictionary of field priorities for added fields
        changeable_fields: Dictionary of field priorities for changeable fields

    Returns:
        Tuple of (changes dict, db_item)

    Raises:
        DatabaseOperationError: When database operations fail
    """
    try:
        db_item = await CopyrightItem.get(material_id=new_item.get("material_id"))
        changes = {}
        changes, db_item = compare_and_update_fields(
            new_item, db_item, added_fields, changes
        )
        changes, db_item = compare_and_update_fields(
            new_item, db_item, changeable_fields, changes
        )
    except Exception as e:
        logger.warning(f"Could not update item {new_item.get('material_id')}: {e}")
        raise DatabaseOperationError(
            f"Failed to process item {new_item.get('material_id')}: {e}"
        ) from e

    return changes, db_item


async def execute_bulk_database_operations(
    changelist: list[CopyrightItem],
    updates: dict,
    cur_user: str | dict | None,
    new_objects: list[CopyrightItem],
    update_relations: bool,
    settings: Settings,
) -> None:
    """
    Execute bulk database operations including updates and changelog creation.

    Args:
        changelist: List of modified CopyrightItem objects
        updates: Dictionary of change details keyed by material_id
        cur_user: Current user information
        new_objects: List of newly created CopyrightItem objects
        update_relations: Whether to update relations after processing
        settings: Application settings
    """
    if changelist:
        # get all values from 'updates'
        # then get list of all distinct keys from all those dicts
        # then drop keys 'material_id' and 'update_time'
        # then add all those keys to the fields to update

        all_keys = {key for item in updates.values() for key in item}
        all_keys.discard("material_id")
        all_keys.discard("update_time")
        changed_fields = list(all_keys)
        changed_fields.append("modified_at")

        # process in batches of max 50:
        for change_batch in batched(changelist, 50):
            logger.info(
                f"Updating fields {changed_fields} for {len(change_batch)} items that were changed."
            )
            await CopyrightItem.bulk_update(change_batch, fields=changed_fields)

        if cur_user:
            user_email = (
                cur_user.get("email") if isinstance(cur_user, dict) else cur_user
            )
            [
                changes.update({"modified_by": user_email})
                for changes in updates.values()
            ]
            logger.info(f"items modified by {user_email}")

        for update_batch in batched(
            [(mat_id, changes) for mat_id, changes in updates.items()], 50
        ):
            logger.info(f"Updating {len(update_batch)} changelog items in db.")

            mat_ids = [mat_id for mat_id, _ in update_batch]
            await ItemUpdate.bulk_create(
                [
                    ItemUpdate(change_details=changes, material_id=mat_id)
                    for mat_id, changes in update_batch
                ]
            )

            # now add each ItemUpdate as a m2m relation to the corresponding CopyrightItem
            for mat_id in mat_ids:
                item = await CopyrightItem.get(material_id=mat_id)
                update = (
                    await ItemUpdate.filter(material_id=mat_id)
                    .order_by("-created_at")
                    .first()
                )
                await item.changes.add(update)

    if (changelist or new_objects) and update_relations:
        logger.success("Updating relations for all CopyrightItems.")


def record_field_change(
    changes: dict,
    field: str,
    new_value: Any,
    old_value: Any,
    reason: str,
    db_item: Any = None,
) -> dict:
    """
    Record a field change in the changes dictionary and update the database item.

    Args:
        changes: Dictionary to record changes in
        field: Field name being changed
        new_value: New value for the field
        old_value: Old value for the field
        reason: Reason for the change
        db_item: Database item to update (optional, for backward compatibility)

    Returns:
        Updated changes dictionary
    """
    if field == "file_exists" and db_item:
        db_item.last_canvas_check = datetime.now()
        changes["last_canvas_check"] = {
            "old": str(old_value),
            "new": str(new_value),
        }
    else:
        logger.debug(
            f"[{reason}] [{field}] {old_value} ({type(old_value)}) --> {new_value} ({type(new_value)})"
        )
    changes[field] = {"old": str(old_value), "new": str(new_value)}

    if db_item:
        setattr(db_item, field, new_value)
    return changes


def compare_and_update_fields(
    new_item: dict, db_item: CopyrightItem, fielddict: dict, changes: dict
) -> tuple[dict, CopyrightItem]:
    """
    Compare fields between new item and database item, updating the database item
    and recording changes according to merge rules.

    Args:
        new_item: Dictionary with new field values
        db_item: Existing database item
        fielddict: Dictionary of field names to ordering rules
        changes: Dictionary to record changes in

    Returns:
        Tuple of (changes dict, updated db_item)
    """
    if not changes:
        changes = {
            "material_id": new_item.get("material_id"),
            "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

    for field, ordering in fielddict.items():
        new_value = new_item.get(field)
        old_value = getattr(db_item, field)

        # Early return: skip if new value is None
        if new_value is None:
            continue

        # Special handling for file_exists
        if field == "file_exists":
            new_value = _normalize_file_exists(new_value)
            if not isinstance(new_value, bool):
                continue

            changes = record_field_change(
                changes,
                field,
                new_value,
                old_value,
                "file_exists value received, always update",
                db_item,
            )
            continue

        # Type casting for comparison
        try:
            cast_success = _cast_values_for_comparison(
                field, new_value, old_value, db_item
            )
            if not cast_success:
                continue
        except TypeCastError as e:
            logger.warning(f"Type casting failed for field '{field}': {e}")
            continue

        # Skip if values are the same after casting
        if new_value == old_value:
            continue

        # Use strategy pattern for field-specific comparison
        strategy = get_comparison_strategy(field, db_item)
        should_update, reason = strategy.should_update(new_value, old_value, ordering)

        # Handle null-to-value case
        if new_value is not None and old_value is None:
            should_update = True
            reason = "no old value"

        if should_update:
            changes = record_field_change(
                changes, field, new_value, old_value, reason, db_item
            )

    return changes, db_item


def _cast_values_for_comparison(
    field: str, new_value: Any, old_value: Any, db_item: CopyrightItem
) -> bool:
    """
    Cast values for comparison and update them in place.

    Args:
        field: Field name
        new_value: New value (will be modified)
        old_value: Old value (will be modified)
        db_item: Database item for context

    Returns:
        True if casting succeeded, False otherwise

    Raises:
        TypeCastError: When type casting fails
    """
    try:
        if isinstance(old_value, datetime):
            new_value = _cast_datetime_value(new_value)
            old_value = old_value.replace(tzinfo=UTC) if old_value else old_value
        elif isinstance(old_value, Enum):
            old_value = _cast_enum_value(old_value, type(old_value))
        elif isinstance(old_value, float):
            new_value = _cast_numeric_value(new_value, float)
            old_value = round(old_value, 2)
        elif isinstance(old_value, int):
            new_value = _cast_numeric_value(new_value, int)
        return True
    except Exception as e:
        logger.debug(
            f"Error {e} while typecasting data for field comparison of {field}"
        )
        raise TypeCastError(f"Failed to cast values for field '{field}': {e}") from e


def _cast_datetime_value(value: Any) -> datetime | None:
    """Cast a value to datetime with multiple format fallbacks."""
    if not value:
        return None

    try:
        return datetime.strptime(value, "%Y-%m-%d %H:%M:%S%z").replace(tzinfo=UTC)
    except Exception:
        try:
            return datetime.strptime(value, "%Y-%m-%d %H:%M:%S").replace(tzinfo=UTC)
        except Exception:
            with contextlib.suppress(Exception):
                return datetime.strptime(value, "%Y-%m-%d").replace(tzinfo=UTC)
    return None


def _cast_numeric_value(value: Any, target_type: type) -> int | float | None:
    """Cast a value to int or float with safe parsing."""
    if target_type is int:
        parsed = safe_int(value)
        return parsed if parsed is not None else None
    elif target_type is float:
        parsed = safe_float(value)
        return round(parsed, 2) if parsed is not None else None
    return None


def _normalize_file_exists(value: Any) -> bool | None:
    """Normalize file_exists values to boolean."""
    match value:
        case True | 1 | "1" | "true" | "True":
            return True
        case False | 0 | "0" | "false" | "False":
            return False
        case None | "":
            return None
        case _:
            return None


def _cast_enum_value(value: Any, enum_class: type) -> Any:
    """Cast a value to enum, returning the enum value."""
    if isinstance(value, enum_class):
        return value.value
    return value


class DataSource(StrEnum):
    """Enum for the source of the data.

    This is used to determine how to handle the data when updating the database.
    """

    RAW_QLIK_DATA = "raw_qlik_data"
    WEEKLY_SHEET = "weekly_sheet"
    OVERVIEW_SHEET = "overview_sheet"
    EA_SCRIPT = "ea_script"
    WEB_DASHBOARD = "dashboard"


async def update_copyright_items(
    settings: Settings,
    data: pl.DataFrame | list[dict],
    update_relations: bool = True,
    overwrite: bool = False,
    user_info: dict | None = None,
) -> None:
    """
    Update the db with copyrightitems from the dataframe (or pre-filtered list of dicts from a df).
    Adds new if they don't exist, or updates if they do.
    See `compare_items` and the dicts added_fields, changeable_fields, core_fields for details on how the comparison is done.

    Once done, and if any updates were made, will call `update_copyright_relations` to update the m2m relations.

    Options:
    - `update_relations`: if True, will call `update_copyright_relations` to update the m2m relations after the update.
    - `overwrite`: if True, will overwrite the existing items in the db with the new ones instead of using the comparison logic.
    """

    if user_info is None:
        user_info = dict()

    cur_user = (
        user_info.get("email")
        if user_info.get("email")
        else {"email": "cip-admin@utwente.nl"}
    )

    await ensure_db_inited(settings)
    # Build merge rules dynamically from settings
    added_fields, changeable_fields = build_merge_rules_from_settings(settings)

    logger.info(f"Received {len(data)} raw copyright items as input for an update.")

    # Preprocess input data
    new_items, update_items = await preprocess_input_data(data)

    # Process new items
    new_objects = await process_new_items(new_items)

    # Process existing items
    changelist, updates = await process_existing_items(
        update_items, added_fields, changeable_fields, overwrite
    )

    # Execute bulk database operations
    await execute_bulk_database_operations(
        changelist, updates, cur_user, new_objects, update_relations, settings
    )

    logger.success("Done updating CopyrightItems!")
    await Tortoise.close_connections()


async def process_staged_raw_data(settings: Settings) -> None:
    """
    Processes the staged raw data and updates the main CopyrightItem table.
    """
    await ensure_db_inited(settings)
    staged_items = await StagedCopyrightItem.all()
    if not staged_items:
        logger.info("No staged raw data to process.")
        return

    logger.info(f"Processing {len(staged_items)} staged raw items...")

    # Helper: list of known staged fields (keeps mapping explicit and safe)
    staged_fields = [
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
        "manual_classification",
        "manual_identifier",
        "scope",
        "remarks",
        "ml_prediction",
        "isbn",
        "doi",
        "in_collection",
        "pagecount",
        "wordcount",
        "picturecount",
        "author",
        "publisher",
        "auditor",
        "last_change",
        "status",
        "reliability",
        "pages_x_students",
        "count_students_registered",
        "retrieved_from_copyright_on",
        "workflow_status",
        "faculty",
        "file_exists",
    ]

    processed_ids: list[int] = []

    # Process in batches to limit transaction size
    for batch_idx, batch in enumerate(batched(staged_items, 50)):
        batch_processed_ids: list[int] = []
        complex_item_dicts: list[dict] = []
        logger.info(f"Processing batch {batch_idx} with {len(batch)} items")
        async with in_transaction():
            for _item_idx, staged_item in enumerate(batch):
                try:
                    # Build a safe dict from known fields
                    item_dict: dict = {}
                    for f in staged_fields:
                        # Use getattr to avoid ORM internals
                        item_dict[f] = getattr(staged_item, f, None)

                    # Ensure material_id is present and castable
                    mid = item_dict.get("material_id")
                    faculty_val = item_dict.get("faculty")
                    if mid is None:
                        logger.warning(
                            f"[STAGED][SKIP] material_id=None, faculty={faculty_val}, stage=raw_data: Missing required material_id"
                        )
                        continue

                    logger.debug(
                        f"[STAGED][PROCESS] material_id={mid}, faculty={faculty_val}, stage=raw_data: Starting processing"
                    )

                    existing_item = await CopyrightItem.get_or_none(material_id=mid)

                    if not existing_item:
                        # Create new item using canonical normalizer
                        logger.debug(
                            f"[STAGED][CREATE] material_id={mid}, faculty={faculty_val}, stage=raw_data: Creating new item"
                        )
                        new_item = await copyright_item_from_dict(item_dict)
                        if new_item:
                            await new_item.save()
                            smid = safe_int(mid)
                            if smid is not None:
                                batch_processed_ids.append(smid)
                            logger.info(
                                f"[STAGED][SUCCESS] material_id={mid}, faculty={faculty_val}, stage=raw_data: Created new item"
                            )
                        else:
                            logger.warning(
                                f"[STAGED][FAIL] material_id={mid}, faculty={faculty_val}, stage=raw_data: Failed to create item from dict"
                            )
                    else:
                        # Check if this item has complex fields that need merging
                        mergeable_fields = get_mergeable_fields()
                        has_complex_fields = False
                        for field in mergeable_fields:
                            if item_dict.get(field) is not None:
                                has_complex_fields = True
                                break

                        if has_complex_fields:
                            # Delegate to canonical merge path
                            logger.debug(
                                f"[STAGED][MERGE] material_id={mid}, faculty={faculty_val}, stage=raw_data: Delegating to complex merge"
                            )
                            complex_item_dicts.append(item_dict)
                            smid = safe_int(mid)
                            if smid is not None:
                                batch_processed_ids.append(smid)
                        else:
                            # Conservative updates for trivial fields
                            update_fields = []
                            # status: try to coerce to Status enum safely
                            status_val = item_dict.get("status")
                            if status_val:
                                new_status = safe_enum(Status, status_val)
                                if new_status and existing_item.status != new_status:
                                    existing_item.status = new_status
                                    update_fields.append("status")

                            # last_change: accept datetime/date or parse common string formats
                            lc_val = item_dict.get("last_change")
                            if lc_val:
                                parsed_date = None
                                if isinstance(lc_val, datetime):
                                    parsed_date = lc_val.date()
                                elif isinstance(lc_val, date):
                                    parsed_date = lc_val
                                elif isinstance(lc_val, str):
                                    try:
                                        # try isoformat first
                                        parsed_dt = datetime.fromisoformat(lc_val)
                                        parsed_date = parsed_dt.date()
                                    except Exception:
                                        for fmt in (
                                            "%Y-%m-%d %H:%M:%S%z",
                                            "%Y-%m-%d %H:%M:%S",
                                            "%Y-%m-%d",
                                        ):
                                            try:
                                                parsed_dt = datetime.strptime(
                                                    lc_val, fmt
                                                )
                                                parsed_date = parsed_dt.date()
                                                break
                                            except Exception:
                                                continue
                                if (
                                    parsed_date
                                    and existing_item.last_change != parsed_date
                                ):
                                    existing_item.last_change = parsed_date
                                    update_fields.append("last_change")

                            if update_fields:
                                logger.debug(
                                    f"[STAGED][UPDATE] material_id={mid}, faculty={faculty_val}, stage=raw_data: Updating fields {update_fields}"
                                )
                                await existing_item.save(update_fields=update_fields)
                                smid = safe_int(mid)
                                if smid is not None:
                                    batch_processed_ids.append(smid)
                                logger.info(
                                    f"[STAGED][SUCCESS] material_id={mid}, faculty={faculty_val}, stage=raw_data: Updated existing item"
                                )

                except Exception as e:
                    err_msg = str(e)
                    mid_val = getattr(staged_item, "material_id", None)
                    faculty_val = getattr(staged_item, "faculty", None)
                    logger.error(
                        f"[STAGED][ERROR] material_id={mid_val}, faculty={faculty_val}, stage=raw_data: {err_msg}"
                    )
                    logger.debug(
                        f"[STAGED][TRACE] material_id={mid_val}, faculty={faculty_val}, stage=raw_data: {traceback.format_exc()}"
                    )
                    # Record failure in the DB for later inspection/retry
                    try:
                        payload = {
                            f: getattr(staged_item, f, None) for f in staged_fields
                        }
                        await StagedProcessingFailure.create(
                            material_id=safe_int(
                                getattr(staged_item, "material_id", None)
                            ),
                            staged_payload=payload,
                            error_message=err_msg[:1900],
                        )
                        logger.info(
                            f"[STAGED][RECORDED] material_id={mid_val}, faculty={faculty_val}, stage=raw_data: Failure recorded in StagedProcessingFailure"
                        )
                    except Exception as record_error:
                        logger.error(
                            f"[STAGED][RECORD_FAIL] material_id={mid_val}, faculty={faculty_val}, stage=raw_data: Could not record failure: {record_error}"
                        )
                    # Do not re-raise; keep other rows processing. Failed staged rows remain for manual inspection.

        # Process complex merges outside transaction since update_copyright_items does its own operations
        if complex_item_dicts:
            logger.info(
                f"Processing {len(complex_item_dicts)} complex merges for batch {batch_idx}"
            )
            try:
                await update_copyright_items(
                    settings=settings, data=complex_item_dicts, update_relations=False
                )
                logger.info(
                    f"Successfully processed {len(complex_item_dicts)} complex merges"
                )
            except Exception as e:
                logger.exception(f"Error processing complex merges: {e}")
                # Don't fail the whole batch, just log the error

        # After successful transaction, remove successfully processed staged rows
        if batch_processed_ids:
            try:
                await StagedCopyrightItem.filter(
                    material_id__in=batch_processed_ids
                ).delete()
                logger.info(
                    f"Cleared {len(batch_processed_ids)} processed staged rows."
                )
                processed_ids.extend(batch_processed_ids)
            except Exception as e:
                logger.exception(
                    f"Error deleting staged rows {batch_processed_ids}: {e}"
                )

    logger.info(
        f"Finished processing staged raw data. Successfully processed {len(processed_ids)} rows."
    )


async def process_staged_faculty_updates(settings: Settings) -> None:
    """
    Processes the staged faculty updates and updates the main CopyrightItem table.
    """
    await ensure_db_inited(settings)
    staged_updates = await StagedFacultyUpdate.all()
    if not staged_updates:
        logger.info("No staged faculty updates to process.")
        return

    logger.info(f"Processing {len(staged_updates)} staged faculty updates...")

    # Process staged faculty updates in batches inside transactions; delete only processed rows
    processed_updates: list[int] = []
    for batch in batched(staged_updates, 100):
        batch_processed: list[int] = []
        async with in_transaction():
            for update in batch:
                try:
                    mid = update.material_id
                    logger.debug(
                        f"[STAGED][PROCESS] material_id={mid}, stage=faculty_update: Starting processing"
                    )

                    item = await CopyrightItem.get_or_none(material_id=mid)
                    if not item:
                        logger.warning(
                            f"[STAGED][SKIP] material_id={mid}, stage=faculty_update: Item not found in database"
                        )
                        continue

                    update_fields = []
                    if (
                        update.manual_classification
                        and item.manual_classification != update.manual_classification
                    ):
                        item.manual_classification = update.manual_classification
                        update_fields.append("manual_classification")
                        logger.debug(
                            f"[STAGED][UPDATE] material_id={mid}, stage=faculty_update: Updating manual_classification"
                        )

                    if update.remarks and item.remarks != update.remarks:
                        item.remarks = update.remarks
                        update_fields.append("remarks")
                        logger.debug(
                            f"[STAGED][UPDATE] material_id={mid}, stage=faculty_update: Updating remarks"
                        )

                    if (
                        update.workflow_status
                        and item.workflow_status != update.workflow_status
                    ):
                        wf_st = safe_enum(WorkflowStatus, update.workflow_status)
                        if wf_st:
                            item.workflow_status = wf_st
                            update_fields.append("workflow_status")
                            logger.debug(
                                f"[STAGED][UPDATE] material_id={mid}, stage=faculty_update: Updating workflow_status"
                            )

                    if update_fields:
                        await item.save(update_fields=update_fields)
                        smid = safe_int(mid)
                        if smid is not None:
                            batch_processed.append(smid)
                        logger.info(
                            f"[STAGED][SUCCESS] material_id={mid}, stage=faculty_update: Updated fields {update_fields}"
                        )
                    else:
                        logger.debug(
                            f"[STAGED][SKIP] material_id={mid}, stage=faculty_update: No fields to update"
                        )

                except Exception as e:
                    mid_val = getattr(update, "material_id", None)
                    logger.error(
                        f"[STAGED][ERROR] material_id={mid_val}, stage=faculty_update: {str(e)}"
                    )
                    logger.debug(
                        f"[STAGED][TRACE] material_id={mid_val}, stage=faculty_update: {traceback.format_exc()}"
                    )

        if batch_processed:
            try:
                await StagedFacultyUpdate.filter(
                    material_id__in=batch_processed
                ).delete()
                logger.info(
                    f"Cleared {len(batch_processed)} processed staged faculty updates."
                )
                processed_updates.extend(batch_processed)
            except Exception:
                logger.exception(
                    f"Error deleting processed staged faculty updates: {batch_processed}"
                )

    logger.info(
        f"Finished processing staged faculty updates. Successfully processed {len(processed_updates)} rows."
    )


async def calculate_derived_fields(settings: Settings) -> None:
    """
    Calculates and updates derived fields like 'possible_fine' and 'infringement'
    for all copyright items.
    """
    from easy_access.db.retrieve import retrieve_copyright_items

    logger.info("Calculating derived fields for all items...")
    df = retrieve_copyright_items(
        settings=settings,
        additional_cols=[
            "possible_fine",
            "infringement",
            "manual_classification",
            "pages_x_students",
        ],
    )

    if df.is_empty():
        logger.info("No items found to calculate derived fields for.")
        return

    # Calculate possible_fine
    df = df.with_columns(
        pl.when(pl.col("possible_fine").is_null())
        .then(
            pl.col("pages_x_students")
            .cast(pl.Int64, strict=False)
            .fill_null(0)
            .mul(settings.fine_amount)
        )
        .otherwise(pl.col("possible_fine"))
        .alias("possible_fine")
    )

    # Calculate infringement
    df = df.with_columns(
        pl.when(
            pl.col("manual_classification").is_null()
            | (pl.col("manual_classification") == "")
            | (pl.col("manual_classification") == "-")
        )
        .then(pl.lit(Infringement.UNDETERMINED.value))
        .when(
            pl.col("manual_classification")
            .str.to_lowercase()
            .str.contains("open|eigen|overig|deleted")
        )
        .then(pl.lit(Infringement.NO.value))
        .when(pl.col("manual_classification").str.to_lowercase().str.contains("lange"))
        .then(pl.lit(Infringement.YES.value))
        .otherwise(pl.lit(Infringement.MAYBE.value))
        .alias("infringement")
    )

    update_df = df.select(["material_id", "possible_fine", "infringement"])

    await update_copyright_items(settings=settings, data=update_df, overwrite=True)
    logger.info("Finished calculating derived fields.")
