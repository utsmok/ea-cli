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
        Decide whether a field value should be replaced during a merge.
        
        Parameters:
            new_value: The incoming candidate value.
            old_value: The existing value to compare against.
            ordering: Optional ordering or ranking information used by some strategies (type and meaning depend on strategy).
        
        Returns:
            tuple(bool, str): `True` if the field should be updated, `False` otherwise; second element is a short reason for the decision or an empty string.
        """
        raise NotImplementedError


class RankedFieldStrategy(FieldComparisonStrategy):
    """Strategy for ranked fields (higher priority = lower index)."""

    def should_update(
        self, new_value: Any, old_value: Any, ordering: list
    ) -> tuple[bool, str]:
        """
        Determine whether a field should be updated based on a ranked ordering where lower index indicates higher priority.
        
        Parameters:
            new_value (Any): Candidate value to consider for update.
            old_value (Any): Existing value to compare against.
            ordering (list): List representing value priority (index 0 = highest priority). If not a list, no update is performed.
        
        Returns:
            tuple[bool, str]: `True` and a short reason if `new_value` has higher priority than `old_value` according to `ordering`; `False` and an empty string otherwise. When a value is not present in `ordering`, a default rank is used.
        """
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
        """
        Determine whether a string field should be updated by preferring longer trimmed strings.
        
        Parameters:
            new_value (Any): Candidate value; update is considered only if this is a string.
            old_value (Any): Existing value; update is considered only if this is a string.
            ordering (Any): Unused by this strategy.
        
        Returns:
            tuple[bool, str]: `True` and reason "new len > old len" if the trimmed `new_value` is longer than the trimmed `old_value`, `False` and an empty string otherwise.
        """
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
        """
        Decides whether the new value should replace the old value by checking if the new value is greater.
        
        Uses safe_compare_greater to compare numeric or date-like values; comparison errors are logged and treated as not updateable.
        
        Parameters:
            new_value: The candidate value to consider for update.
            old_value: The existing value to compare against.
            ordering: Ignored by this strategy (present for API compatibility).
        
        Returns:
            A tuple where the first element is `True` and the second is "new > old" if `new_value` is greater than `old_value`, `False` and an empty string otherwise.
        """
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
        """
        Decides whether a date/datetime field should be updated based on which value is later.
        
        Parameters:
            new_value (date | datetime): The incoming candidate date/time value.
            old_value (date | datetime): The existing stored date/time value.
            ordering (Any): Ignored for date comparisons.
        
        Returns:
            tuple[bool, str]: `True` and the reason "new date > old date" if `new_value` is later than `old_value`, `False` and an empty string otherwise. If either value is not a `date` or `datetime`, returns `False` and an empty string.
        """
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
        """
        Determine whether an enum-like new_value should replace old_value based on a provided ranking.
        
        Parameters:
            new_value: The candidate value to consider for update.
            old_value: The current value to compare against.
            ordering (list | Any): A list defining preferred values in priority order (lower index = higher priority).
                If `ordering` is not a non-empty list, ranking is not applied.
                Values not present in `ordering` are treated as having the default rank.
        
        Returns:
            tuple[bool, str]: `True` and a short reason if `new_value` has higher priority (lower rank) than `old_value`, `False` and an empty string otherwise.
        """
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
        """
        Always require an update when a file existence indication is provided.
        
        Returns:
            tuple[bool, str]: First element is `True` indicating the field should be updated; second element is a human-readable reason string explaining the update decision.
        """
        return True, "file_exists value received, always update"


def get_comparison_strategy(
    field: str, db_item: Any | None = None
) -> FieldComparisonStrategy:
    """
    Selects a FieldComparisonStrategy appropriate for the given field and optional database item.
    
    Parameters:
        field (str): Name of the field to choose a comparison strategy for.
        db_item (Any | None): Optional object whose current attribute value will be inspected to infer the most suitable strategy; if the attribute is missing or db_item is None, inference is skipped.
    
    Returns:
        FieldComparisonStrategy: An instance suitable for comparing/merging values for the specified field (e.g., a file-existence strategy for "file_exists", or a strategy inferred from the current attribute type when db_item is provided).
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