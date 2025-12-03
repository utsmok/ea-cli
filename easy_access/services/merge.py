"""
Merge logic for copyright item updates.

This module contains the comparison and update logic for merging
copyright item fields. It is decoupled from database operations
and operates on Python objects directly.

NOTE: This module contains pure logic with NO database commits.
Database state changes (setattr on items) are performed by the caller.
"""

import contextlib
from datetime import UTC, datetime
from enum import Enum
from typing import Any

from loguru import logger

from easy_access.db.models import WorkflowStatus
from easy_access.services.strategies import get_comparison_strategy
from easy_access.utils import safe_float, safe_int


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


def record_field_change(
    changes: dict,
    field: str,
    new_value: Any,
    old_value: Any,
    reason: str,
    db_item: Any = None,
) -> dict:
    """
    Record a field change in the changes dictionary and optionally update the item.

    Args:
        changes: Dictionary to record changes in
        field: Field name being changed
        new_value: New value for the field
        old_value: Old value for the field
        reason: Reason for the change
        db_item: Object to update (optional). If provided, setattr is called.

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


def _cast_values_for_comparison(
    field: str, new_value: Any, old_value: Any, db_item: Any
) -> tuple[bool, Any, Any]:
    """
    Cast values for comparison and return the modified values.

    Args:
        field: Field name
        new_value: New value
        old_value: Old value
        db_item: Database item for context (used to determine types)

    Returns:
        Tuple of (success, new_value, old_value)

    Raises:
        TypeCastError: When type casting fails
    """
    try:
        if isinstance(old_value, datetime):
            new_value = _cast_datetime_value(new_value)
            old_value = old_value.replace(tzinfo=UTC) if old_value else old_value
        elif isinstance(old_value, Enum):
            new_value = _cast_enum_value(new_value, type(old_value))
            old_value = old_value.value
        elif isinstance(old_value, float):
            new_value = _cast_numeric_value(new_value, float)
            old_value = round(old_value, 2)
        elif isinstance(old_value, int):
            new_value = _cast_numeric_value(new_value, int)
        return True, new_value, old_value
    except Exception as e:
        logger.debug(
            f"Error {e} while typecasting data for field comparison of {field}"
        )
        raise TypeCastError(f"Failed to cast values for field '{field}': {e}") from e


def compare_and_update_fields(
    new_item: dict, db_item: Any, fielddict: dict, changes: dict
) -> tuple[dict, Any]:
    """
    Compare fields between new item and database item, updating the item
    and recording changes according to merge rules.

    This function performs the core merge logic:
    1. For each field in fielddict, compare new vs old values
    2. Apply type casting as needed
    3. Use appropriate strategy to determine if update is needed
    4. Record changes and update the item instance

    Args:
        new_item: Dictionary with new field values
        db_item: Existing database item (Tortoise model instance)
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
            cast_success, new_value, old_value = _cast_values_for_comparison(
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

        # Hard safety net for workflow_status: never allow downgrade regardless of ordering.
        if field == "workflow_status" and isinstance(ordering, list):
            try:
                # Canonical rank map (lower index = higher priority)
                canonical = [
                    WorkflowStatus.Done.value,
                    WorkflowStatus.InProgress.value,
                    WorkflowStatus.ToDo.value,
                ]
                if new_value in canonical and old_value in canonical:
                    new_rank = canonical.index(new_value)
                    old_rank = canonical.index(old_value)
                    # Only update if new has higher priority (smaller index) or old is None
                    if new_rank < old_rank:
                        should_update = True
                        reason = "workflow_status upgrade (canonical ordering)"
                    elif new_rank >= old_rank and old_value is not None:
                        should_update = False
                        reason = "workflow_status downgrade prevented"
            except Exception:
                pass

        # Handle null-to-value case
        if new_value is not None and old_value is None:
            should_update = True
            reason = "no old value"

        if should_update:
            changes = record_field_change(
                changes, field, new_value, old_value, reason, db_item
            )

    return changes, db_item


def calculate_changes(
    new_data: dict, 
    current_item: Any, 
    added_fields: dict,
    changeable_fields: dict,
) -> tuple[dict, Any]:
    """
    Calculate changes needed for a copyright item.

    This is a higher-level function that applies comparison logic
    for both added and changeable fields.

    Args:
        new_data: Dictionary with new field values
        current_item: Current database item
        added_fields: Dictionary of field priorities for added fields
        changeable_fields: Dictionary of field priorities for changeable fields

    Returns:
        Tuple of (changes dict, updated item)
    """
    changes = {}
    changes, current_item = compare_and_update_fields(
        new_data, current_item, added_fields, changes
    )
    changes, current_item = compare_and_update_fields(
        new_data, current_item, changeable_fields, changes
    )
    return changes, current_item
