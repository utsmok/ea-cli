"""
Services layer for Easy Access CLI.

This module contains pure business logic that is decoupled from database operations.
Services handle comparison strategies, merge logic, and field transformations.
"""

from easy_access.services.merge import (
    calculate_changes,
    compare_and_update_fields,
    record_field_change,
)
from easy_access.services.strategies import (
    DateFieldStrategy,
    EnumFieldStrategy,
    FieldComparisonStrategy,
    FileExistsStrategy,
    NumericFieldStrategy,
    RankedFieldStrategy,
    StringFieldStrategy,
    get_comparison_strategy,
)

__all__ = [
    # Strategies
    "FieldComparisonStrategy",
    "RankedFieldStrategy",
    "StringFieldStrategy",
    "NumericFieldStrategy",
    "DateFieldStrategy",
    "EnumFieldStrategy",
    "FileExistsStrategy",
    "get_comparison_strategy",
    # Merge logic
    "calculate_changes",
    "compare_and_update_fields",
    "record_field_change",
]
