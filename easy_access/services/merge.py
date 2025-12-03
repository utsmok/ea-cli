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
    """
    Parse a value into a timezone-aware datetime using common date/time formats.

    Parameters:
        value (Any): The input to parse; expected to be a string in one of these formats:
            "YYYY-MM-DD HH:MM:SS±ZZZZ", "YYYY-MM-DD HH:MM:SS", or "YYYY-MM-DD".

    Returns:
        datetime | None: A `datetime` with UTC tzinfo if parsing succeeds, otherwise `None`.
    """
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
    """
    Convert a value to an integer or float suitable for comparison.

    Parameters:
        value (Any): The input to be parsed as a numeric value.
        target_type (type): Desired numeric type; expected to be `int` or `float`.

    Returns:
        int | float | None: An `int` when `target_type` is `int` and parsing succeeds; a `float` rounded to two decimal places when `target_type` is `float` and parsing succeeds; `None` if parsing fails or `target_type` is unsupported.
    """
    if target_type is int:
        parsed = safe_int(value)
        return parsed if parsed is not None else None
    elif target_type is float:
        parsed = safe_float(value)
        return round(parsed, 2) if parsed is not None else None
    return None


def _normalize_file_exists(value: Any) -> bool | None:
    """
    Normalize various truthy/falsey representations for the `file_exists` field to `True`, `False`, or `None`.

    Parameters:
        value (Any): Input value that may be a boolean, integer, string, None, or empty string.

    Returns:
        bool | None: `True` for common truthy values (`True`, `1`, `"1"`, `"true"`, `"True"`),
        `False` for common falsey values (`False`, `0`, `"0"`, `"false"`, `"False"`), and `None` for `None`, empty string, or unrecognized values.
    """
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
    """
    Retrieve the underlying value when given an enum instance, otherwise leave the input unchanged.

    Parameters:
        value (Any): The value to inspect; may be an instance of `enum_class`.
        enum_class (type): The Enum class to check against.

    Returns:
        Any: The enum member's `.value` if `value` is an instance of `enum_class`, otherwise `value` unchanged.
    """
    if isinstance(value, enum_class):
        return value.value
    return value


def _cast_values_for_comparison(
    field: str, new_value: Any, old_value: Any, db_item: Any
) -> tuple[bool, Any, Any]:
    """
    Prepare and normalize new and old values for field comparison according to the old value's type.

    Parameters:
        field (str): Field name (used for context in error messages).
        new_value (Any): Candidate new value to be cast for comparison.
        old_value (Any): Existing value whose type determines casting rules.
        db_item (Any): Optional context object (unused by most casts but available for type resolution).

    Returns:
        tuple: (success, new_value, old_value) where `success` is `True` on successful casting, `new_value` is the cast/normalized new value, and `old_value` is the normalized old value suitable for comparison.

    Raises:
        TypeCastError: If converting `new_value` to the type implied by `old_value` fails.
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
    Compare fields from a new item against an existing item, apply allowed updates, and record any changes.

    Per-field comparison uses type-aware casting and a pluggable comparison strategy to decide whether to update. The function mutates the provided `db_item` when updates are applied and records each change in `changes`. If `changes` is empty, it is initialized with `material_id` from `new_item` and a generated `update_time`. The function also performs special handling for `file_exists` normalization and enforces a canonical upgrade-only rule for `workflow_status`.

    Parameters:
        new_item (dict): Source data with proposed field values.
        db_item (Any): Existing item instance to compare against and update when changes are accepted.
        fielddict (dict): Mapping of field names to ordering/rules used by comparison strategies.
        changes (dict): Accumulator for recorded changes; returned and updated in-place.

    Returns:
        tuple: (`changes` dict with recorded field changes, updated `db_item` instance)
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
    Compute and apply field-level updates for added and changeable fields, returning recorded changes and the updated item.

    Parameters:
        new_data (dict): Incoming data with candidate field values.
        current_item (Any): Existing item object to compare against and optionally update.
        added_fields (dict): Mapping of fields and their ordering rules to treat as newly added.
        changeable_fields (dict): Mapping of fields and their ordering rules eligible for updates.

    Returns:
        tuple[dict, Any]: A tuple containing the changes dictionary (field -> {old, new, reason}) and the potentially updated item.
    """
    changes = {}
    changes, current_item = compare_and_update_fields(
        new_data, current_item, added_fields, changes
    )
    changes, current_item = compare_and_update_fields(
        new_data, current_item, changeable_fields, changes
    )
    return changes, current_item
