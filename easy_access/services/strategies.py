"""
Field comparison strategies for merge operations.

These strategies implement the Strategy pattern for comparing fields 
during copyright item updates. Each strategy defines its own logic 
for determining whether a field should be updated.

NOTE: This module contains pure logic with NO database dependencies.
"""

from datetime import date, datetime
from enum import Enum
from typing import Any

from loguru import logger

from easy_access.utils import safe_compare_greater

# Constants for comparison logic
DEFAULT_RANK = 20


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
    field: str, db_item: Any | None = None
) -> FieldComparisonStrategy:
    """
    Get the appropriate comparison strategy for a field.

    Args:
        field: Field name
        db_item: Database item to check field types (optional).
                 Can be a Tortoise model instance or any object with attributes.

    Returns:
        FieldComparisonStrategy instance
    """
    # Special cases for file_exists
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
