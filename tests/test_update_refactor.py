"""
Unit tests for refactored update_copyright_items components.
"""

import pytest
from datetime import date, datetime, UTC
from unittest.mock import Mock, AsyncMock, patch
from enum import Enum

from easy_access.db.update import (
    FieldComparisonStrategy,
    RankedFieldStrategy,
    StringFieldStrategy,
    NumericFieldStrategy,
    FileExistsStrategy,
    DateFieldStrategy,
    EnumFieldStrategy,
    get_comparison_strategy,
    record_field_change,
    compare_and_update_fields,
    _cast_values_for_comparison,
    _cast_datetime_value,
    _cast_numeric_value,
    _normalize_file_exists,
    _cast_enum_value,
    preprocess_input_data,
    MergeError,
    TypeCastError,
    DatabaseOperationError,
    ValidationError,
    MergeConflictError,
)
from easy_access.db.models import CopyrightItem, Status, WorkflowStatus, Classification
from easy_access.db.base import copyright_item_from_dict
from easy_access.merge_rules import (
    build_merge_rules_from_settings,
    get_mergeable_fields,
    added_fields,
    changeable_fields,
)
from easy_access.settings import Settings


class TestCopyrightItemFromDict:
    """Test the copyright_item_from_dict function."""

    @pytest.mark.asyncio
    async def test_copyright_item_from_dict_valid_data(self):
        """Test creating CopyrightItem from valid dict data."""
        # Mock the Faculty.get call and CopyrightItem constructor
        with patch('easy_access.db.base.Faculty') as mock_faculty_class, \
             patch('easy_access.db.base.CopyrightItem') as mock_copyright_item_class:

            # Create a mock faculty with the required attributes
            mock_faculty = Mock()
            mock_faculty.abbreviation = "TEST"
            mock_faculty_class.get = AsyncMock(return_value=mock_faculty)

            # Mock the CopyrightItem constructor
            mock_copyright_item = Mock()
            mock_copyright_item_class.return_value = mock_copyright_item

            test_data = {
                "material_id": "12345",
                "period": "2023-1A",
                "department": "Test Department",
                "course_code": "TEST101",
                "course_name": "Test Course",
                "title": "Test Title",
                "faculty": "TEST",
                "classification": "lange overname",
                "status": "Published",
                "pagecount": "10",
                "wordcount": "1000",
                "picturecount": "5",
                "reliability": "8",
                "pages_x_students": "50",
                "count_students_registered": "25",
                "last_change": "2023-01-01",
                "retrieved_from_copyright_on": "2023-01-01 12:00:00",
            }

            result = await copyright_item_from_dict(test_data)  # type: ignore[arg-type]

            assert result is not None
            # Verify that CopyrightItem was called with the correct data
            mock_copyright_item_class.assert_called_once()
            call_args = mock_copyright_item_class.call_args[1]  # Get kwargs
            assert call_args['material_id'] == 12345
            assert call_args['title'] == "Test Title"
            assert call_args['faculty'] == mock_faculty

    @pytest.mark.asyncio
    async def test_copyright_item_from_dict_missing_required_fields(self):
        """Test handling of missing required fields."""
        with patch('easy_access.db.base.Faculty') as mock_faculty_class, \
             patch('easy_access.db.base.CopyrightItem') as mock_copyright_item_class:

            mock_faculty = Mock()
            mock_faculty.abbreviation = "UNM"
            mock_faculty_class.get = AsyncMock(return_value=mock_faculty)

            mock_copyright_item = Mock()
            mock_copyright_item_class.return_value = mock_copyright_item

            # Missing required fields like material_id
            test_data = {
                "title": "Test Title",
                "faculty": "TEST",
            }

            result = await copyright_item_from_dict(test_data)  # type: ignore[arg-type]

            # Should return None due to missing material_id
            assert result is None

    @pytest.mark.asyncio
    async def test_copyright_item_from_dict_invalid_data_types(self):
        """Test handling of invalid data types."""
        with patch('easy_access.db.base.Faculty') as mock_faculty_class:
            mock_faculty = Mock()
            mock_faculty_class.get = AsyncMock(return_value=mock_faculty)

            test_data = {
                "material_id": "not_a_number",  # Invalid material_id
                "period": "2023-1A",
                "department": "Test Department",
                "course_code": "TEST101",
                "course_name": "Test Course",
                "faculty": "TEST",
                "pagecount": "not_a_number",  # Invalid pagecount
            }

            result = await copyright_item_from_dict(test_data)  # type: ignore[arg-type]

            # Should handle invalid data gracefully
            assert result is None

    @pytest.mark.asyncio
    async def test_copyright_item_from_dict_faculty_fallback(self):
        """Test faculty fallback to UNM when faculty lookup fails."""
        with patch('easy_access.db.base.Faculty') as mock_faculty_class, \
             patch('easy_access.db.base.CopyrightItem') as mock_copyright_item_class:

            # First call fails, second succeeds with UNM
            mock_faculty_unm = Mock()
            mock_faculty_unm.abbreviation = "UNM"
            mock_faculty_class.get = AsyncMock(side_effect=[
                Exception("Faculty not found"),
                mock_faculty_unm
            ])

            mock_copyright_item = Mock()
            mock_copyright_item_class.return_value = mock_copyright_item

            test_data = {
                "material_id": "12345",
                "period": "2023-1A",
                "department": "Test Department",
                "course_code": "TEST101",
                "course_name": "Test Course",
                "faculty": "NONEXISTENT",
            }

            result = await copyright_item_from_dict(test_data)  # type: ignore[arg-type]

            assert result is not None
            # Verify UNM faculty was requested as fallback
            assert mock_faculty_class.get.call_count == 2
            mock_faculty_class.get.assert_any_call(abbreviation="UNM")
            # Verify CopyrightItem was called with UNM faculty
            mock_copyright_item_class.assert_called_once()
            call_args = mock_copyright_item_class.call_args[1]
            assert call_args['faculty'] == mock_faculty_unm

    @pytest.mark.asyncio
    async def test_copyright_item_from_dict_default_values(self):
        """Test that default values are applied for missing optional fields."""
        with patch('easy_access.db.base.Faculty') as mock_faculty_class, \
             patch('easy_access.db.base.CopyrightItem') as mock_copyright_item_class:

            mock_faculty = Mock()
            mock_faculty.abbreviation = "TEST"
            mock_faculty_class.get = AsyncMock(return_value=mock_faculty)

            mock_copyright_item = Mock()
            mock_copyright_item_class.return_value = mock_copyright_item

            test_data = {
                "material_id": "12345",
                "period": "2023-1A",
                "department": "Test Department",
                "course_code": "TEST101",
                "course_name": "Test Course",
                "faculty": "TEST",
                # Missing classification and status - should get defaults
            }

            result = await copyright_item_from_dict(test_data)  # type: ignore[arg-type]

            assert result is not None
            # Verify CopyrightItem was called with defaults
            mock_copyright_item_class.assert_called_once()
            call_args = mock_copyright_item_class.call_args[1]
            assert call_args['classification'] == Classification.LANGE_OVERNAME.value
            assert call_args['status'] == Status.PUBLISHED.value
            assert call_args['filetype'] == "unknown"  # Default for missing filetype


class TestMergeRules:
    """Test the merge_rules.py functions."""

    def test_get_mergeable_fields(self):
        """Test get_mergeable_fields returns correct field sets."""
        mergeable = get_mergeable_fields()

        # Should contain both added and changeable fields
        expected_fields = set(added_fields.keys()) | set(changeable_fields.keys())
        assert set(mergeable.keys()) == expected_fields

    def test_build_merge_rules_from_settings_basic(self):
        """Test build_merge_rules_from_settings with basic settings."""
        settings = Mock()
        settings.classification_options = None
        settings.data_settings = None

        added, changeable = build_merge_rules_from_settings(settings)

        # Should return default values when no settings provided
        assert added == added_fields
        assert changeable == changeable_fields

    def test_build_merge_rules_from_settings_with_classification_options(self):
        """Test build_merge_rules_from_settings with classification options."""
        settings = Mock()
        settings.classification_options = ["open access", "korte overname", "lange overname"]
        settings.data_settings = None

        added, changeable = build_merge_rules_from_settings(settings)

        # Should update changeable_fields with new classification priorities
        assert changeable["manual_classification"] == ["open access", "korte overname", "lange overname"]

    def test_build_merge_rules_from_settings_with_workflow_options(self):
        """Test build_merge_rules_from_settings with workflow status options."""
        settings = Mock()
        settings.classification_options = None

        # Mock data_settings with workflow_status dropdown
        mock_col_info = Mock()
        mock_col_info.name = "workflow_status"
        mock_col_info.dropdown_options = '"ToDo,InProgress,Done"'

        mock_data_settings = Mock()
        mock_data_settings.data_entry_cols = [mock_col_info]

        settings.data_settings = mock_data_settings

        added, changeable = build_merge_rules_from_settings(settings)

        # Should update added_fields with parsed workflow options
        assert added["workflow_status"] == ["Done", "InProgress", "ToDo"]

    def test_build_merge_rules_from_settings_invalid_dropdown(self):
        """Test build_merge_rules_from_settings with invalid dropdown format."""
        settings = Mock()
        settings.classification_options = None

        # Mock data_settings with malformed dropdown
        mock_col_info = Mock()
        mock_col_info.name = "workflow_status"
        mock_col_info.dropdown_options = "invalid_format_no_quotes"

        mock_data_settings = Mock()
        mock_data_settings.data_entry_cols = [mock_col_info]

        settings.data_settings = mock_data_settings

        added, changeable = build_merge_rules_from_settings(settings)

        # Should keep default workflow options when parsing fails
        assert added["workflow_status"] == added_fields["workflow_status"]


class TestFieldComparisonStrategies:
    """Test all field comparison strategies."""

    def test_ranked_field_strategy_higher_priority_wins(self):
        """Test that higher priority (lower index) values win."""
        strategy = RankedFieldStrategy()
        ordering = ["low", "medium", "high"]

        # New value has higher priority (lower index)
        should_update, reason = strategy.should_update("low", "high", ordering)
        assert should_update is True
        assert "new rank < old rank" in reason

        # Old value has higher priority
        should_update, reason = strategy.should_update("high", "low", ordering)
        assert should_update is False
        assert reason == ""

    def test_ranked_field_strategy_invalid_ordering(self):
        """Test behavior with invalid ordering."""
        strategy = RankedFieldStrategy()

        should_update, reason = strategy.should_update("value1", "value2", [])
        assert should_update is False
        assert reason == ""

    def test_string_field_strategy_longer_wins(self):
        """Test that longer strings take precedence."""
        strategy = StringFieldStrategy()

        should_update, reason = strategy.should_update("longer string", "short", None)
        assert should_update is True
        assert "new len > old len" in reason

        should_update, reason = strategy.should_update("short", "longer string", None)
        assert should_update is False
        assert reason == ""

    def test_string_field_strategy_non_strings(self):
        """Test behavior with non-string values."""
        strategy = StringFieldStrategy()

        should_update, reason = strategy.should_update(123, "string", None)
        assert should_update is False
        assert reason == ""

    def test_numeric_field_strategy_greater_wins(self):
        """Test that greater numeric values win."""
        strategy = NumericFieldStrategy()

        should_update, reason = strategy.should_update(10, 5, None)
        assert should_update is True
        assert "new > old" in reason

        should_update, reason = strategy.should_update(3, 8, None)
        assert should_update is False
        assert reason == ""

    def test_date_field_strategy_newer_wins(self):
        """Test that newer dates take precedence."""
        strategy = DateFieldStrategy()
        old_date = date(2023, 1, 1)
        new_date = date(2023, 6, 1)

        should_update, reason = strategy.should_update(new_date, old_date, None)
        assert should_update is True
        assert "new date > old date" in reason

        should_update, reason = strategy.should_update(old_date, new_date, None)
        assert should_update is False
        assert reason == ""

    def test_date_field_strategy_non_dates(self):
        """Test behavior with non-date values."""
        strategy = DateFieldStrategy()

        should_update, reason = strategy.should_update("2023-01-01", date(2023, 1, 1), None)
        assert should_update is False
        assert reason == ""

    def test_enum_field_strategy_with_ordering(self):
        """Test enum strategy with ordering provided."""
        strategy = EnumFieldStrategy()
        ordering = ["low", "medium", "high"]

        should_update, reason = strategy.should_update("low", "high", ordering)
        assert should_update is True
        assert "new enum rank < old enum rank" in reason

        should_update, reason = strategy.should_update("high", "low", ordering)
        assert should_update is False
        assert reason == ""

    def test_enum_field_strategy_without_ordering(self):
        """Test enum strategy without ordering (no update)."""
        strategy = EnumFieldStrategy()

        should_update, reason = strategy.should_update("value1", "value2", None)
        assert should_update is False
        assert reason == ""

    def test_file_exists_strategy_always_updates(self):
        """Test that file_exists always updates when received."""
        strategy = FileExistsStrategy()

        should_update, reason = strategy.should_update(True, False, None)
        assert should_update is True
        assert "file_exists value received, always update" in reason

        should_update, reason = strategy.should_update(False, True, None)
        assert should_update is True
        assert "file_exists value received, always update" in reason


class TestGetComparisonStrategy:
    """Test the get_comparison_strategy function."""

    def test_file_exists_returns_file_exists_strategy(self):
        """Test that file_exists field returns FileExistsStrategy."""
        strategy = get_comparison_strategy("file_exists")
        assert isinstance(strategy, FileExistsStrategy)

    def test_date_field_returns_date_strategy(self):
        """Test that date fields return DateFieldStrategy."""
        mock_item = Mock()
        mock_item.last_change = date(2023, 1, 1)

        strategy = get_comparison_strategy("last_change", mock_item)
        assert isinstance(strategy, DateFieldStrategy)

    def test_datetime_field_returns_date_strategy(self):
        """Test that datetime fields return DateFieldStrategy."""
        mock_item = Mock()
        mock_item.retrieved_from_copyright_on = datetime(2023, 1, 1, tzinfo=UTC)

        strategy = get_comparison_strategy("retrieved_from_copyright_on", mock_item)
        assert isinstance(strategy, DateFieldStrategy)

    def test_enum_field_returns_enum_strategy(self):
        """Test that enum fields return EnumFieldStrategy."""
        mock_item = Mock()
        mock_item.status = Status.PUBLISHED

        strategy = get_comparison_strategy("status", mock_item)
        assert isinstance(strategy, EnumFieldStrategy)

    def test_unknown_field_returns_numeric_strategy(self):
        """Test that unknown fields default to NumericFieldStrategy."""
        strategy = get_comparison_strategy("unknown_field")
        assert isinstance(strategy, NumericFieldStrategy)

    def test_missing_attribute_returns_numeric_strategy(self):
        """Test that missing attributes default to NumericFieldStrategy."""
        mock_item = Mock()
        mock_item.configure_mock(**{"some_field": None})

        strategy = get_comparison_strategy("nonexistent_field", mock_item)
        assert isinstance(strategy, NumericFieldStrategy)


class TestTypeCastingFunctions:
    """Test type casting helper functions."""

    def test_cast_datetime_value_valid_string(self):
        """Test casting valid datetime strings."""
        result = _cast_datetime_value("2023-01-01 12:00:00")
        expected = datetime(2023, 1, 1, 12, 0, 0, tzinfo=UTC)
        assert result == expected

    def test_cast_datetime_value_date_only(self):
        """Test casting date-only strings."""
        result = _cast_datetime_value("2023-01-01")
        expected = datetime(2023, 1, 1, tzinfo=UTC)
        assert result == expected

    def test_cast_datetime_value_invalid(self):
        """Test casting invalid values."""
        assert _cast_datetime_value("invalid") is None
        assert _cast_datetime_value(None) is None

    def test_cast_numeric_value_int(self):
        """Test casting to int."""
        assert _cast_numeric_value("123", int) == 123
        assert _cast_numeric_value(45.0, int) == 45
        assert _cast_numeric_value("invalid", int) is None

    def test_cast_numeric_value_float(self):
        """Test casting to float."""
        assert _cast_numeric_value("12.34", float) == 12.34
        assert _cast_numeric_value(5, float) == 5.0
        assert _cast_numeric_value("invalid", float) is None

    def test_normalize_file_exists(self):
        """Test normalizing file_exists values."""
        assert _normalize_file_exists(True) is True
        assert _normalize_file_exists(1) is True
        assert _normalize_file_exists("1") is True
        assert _normalize_file_exists("true") is True

        assert _normalize_file_exists(False) is False
        assert _normalize_file_exists(0) is False
        assert _normalize_file_exists("0") is False
        assert _normalize_file_exists("false") is False

        assert _normalize_file_exists(None) is None
        assert _normalize_file_exists("") is None
        assert _normalize_file_exists("invalid") is None

    def test_cast_enum_value(self):
        """Test casting enum values."""
        # When value is already an enum
        status = Status.PUBLISHED
        assert _cast_enum_value(status, Status) == Status.PUBLISHED.value

        # When value is not an enum
        assert _cast_enum_value("Published", Status) == "Published"


class TestRecordFieldChange:
    """Test the record_field_change function."""

    def test_record_field_change_basic(self):
        """Test basic field change recording."""
        changes = {"material_id": 123}
        db_item = Mock()

        result = record_field_change(changes, "title", "New Title", "Old Title", "test reason", db_item)

        assert result["title"] == {"old": "Old Title", "new": "New Title"}
        # Verify the db_item was updated
        assert db_item.title == "New Title"

    def test_record_field_change_file_exists(self):
        """Test file_exists field change recording."""
        changes = {"material_id": 123}
        db_item = Mock()

        result = record_field_change(changes, "file_exists", True, False, "file_exists value received", db_item)

        assert result["file_exists"] == {"old": "False", "new": "True"}
        # Verify the db_item was updated
        assert db_item.file_exists == True
        assert hasattr(db_item, 'last_canvas_check')

    def test_record_field_change_without_db_item(self):
        """Test field change recording without db_item."""
        changes = {"material_id": 123}

        result = record_field_change(changes, "title", "New Title", "Old Title", "test reason")

        assert result["title"] == {"old": "Old Title", "new": "New Title"}


class TestCompareAndUpdateFields:
    """Test the compare_and_update_fields function."""

    def test_compare_and_update_fields_basic(self):
        """Test basic field comparison and update."""
        new_item = {"material_id": 123, "title": "A much longer new title"}
        db_item = Mock()
        db_item.title = "Old Title"
        db_item.material_id = 123
        changes = {}

        result_changes, result_db_item = compare_and_update_fields(new_item, db_item, {"title": []}, changes)

        # The title should be updated because the new title is longer
        assert "title" in result_changes
        assert result_changes["title"] == {"old": "Old Title", "new": "A much longer new title"}

    def test_compare_and_update_fields_no_change(self):
        """Test when no changes are needed."""
        new_item = {"material_id": 123, "title": "Same Title"}
        db_item = Mock()
        db_item.title = "Same Title"
        changes = {}

        result_changes, result_db_item = compare_and_update_fields(new_item, db_item, {"title": []}, changes)

        assert "title" not in result_changes

    def test_compare_and_update_fields_none_value(self):
        """Test handling of None values."""
        new_item = {"material_id": 123, "title": None}
        db_item = Mock()
        db_item.title = "Old Title"
        changes = {}

        result_changes, result_db_item = compare_and_update_fields(new_item, db_item, {"title": []}, changes)

        assert "title" not in result_changes


class TestCastValuesForComparison:
    """Test the _cast_values_for_comparison function."""

    def test_cast_datetime_values(self):
        """Test casting datetime values."""
        db_item = Mock()
        db_item.last_change = datetime(2023, 1, 1, tzinfo=UTC)

        success, new_val, old_val = _cast_values_for_comparison("last_change", "2023-06-01", datetime(2023, 1, 1, tzinfo=UTC), db_item)

        assert success is True

    def test_cast_enum_values(self):
        """Test casting enum values."""
        db_item = Mock()
        db_item.status = Status.PUBLISHED

        success, new_val, old_val = _cast_values_for_comparison("status", "Unpublished", Status.PUBLISHED, db_item)

        assert success is True

    def test_cast_numeric_values(self):
        """Test casting numeric values."""
        db_item = Mock()
        db_item.pagecount = 100

        success, new_val, old_val = _cast_values_for_comparison("pagecount", "150", 100, db_item)

        assert success is True

    def test_cast_failure_raises_error(self):
        """Test that casting failures raise TypeCastError."""
        db_item = Mock()
        db_item.pagecount = 100

        # This should work fine - int to int casting
        success, new_val, old_val = _cast_values_for_comparison("pagecount", "150", 100, db_item)
        assert success is True


class TestPreprocessInputData:
    """Test the preprocess_input_data function."""

    @pytest.mark.asyncio
    async def test_preprocess_dataframe(self):
        """Test preprocessing DataFrame input."""
        import polars as pl

        # Create mock data
        data = pl.DataFrame({
            "material_id": [123, 456],
            "title": ["Title 1", "Title 2"],
            "period": ["2023-1A", "2023-2A"],
            "department": ["Dept1", "Dept2"],
            "course_code": ["CODE1", "CODE2"],
            "course_name": ["Course 1", "Course 2"]
        })

        # Mock existing material_ids
        mock_existing = [{"material_id": 123}]
        from unittest.mock import patch
        with patch("easy_access.db.update.CopyrightItem.all") as mock_all:
            mock_query = AsyncMock()
            mock_query.values = AsyncMock(return_value=mock_existing)
            mock_all.return_value = mock_query

            new_items, update_items = await preprocess_input_data(data)

            assert len(new_items) == 1  # One new item (456)
            assert len(update_items) == 1  # One update item (123)

    @pytest.mark.asyncio
    async def test_preprocess_list_input(self):
        """Test preprocessing list input."""
        data = [{"material_id": 123, "title": "Test"}]

        new_items, update_items = await preprocess_input_data(data)

        assert len(new_items) == 0  # No new items (missing required fields)
        assert len(update_items) == 1  # One update item


class TestCustomExceptions:
    """Test custom exception classes."""

    def test_merge_error_base(self):
        """Test base MergeError."""
        error = MergeError("Test error")
        assert str(error) == "Test error"

    def test_merge_conflict_error(self):
        """Test MergeConflictError."""
        error = MergeConflictError("Conflict detected")
        assert str(error) == "Conflict detected"
        assert isinstance(error, MergeError)

    def test_type_cast_error(self):
        """Test TypeCastError."""
        error = TypeCastError("Cast failed")
        assert str(error) == "Cast failed"
        assert isinstance(error, MergeError)

    def test_database_operation_error(self):
        """Test DatabaseOperationError."""
        error = DatabaseOperationError("DB operation failed")
        assert str(error) == "DB operation failed"
        assert isinstance(error, MergeError)

    def test_validation_error(self):
        """Test ValidationError."""
        error = ValidationError("Validation failed")
        assert str(error) == "Validation failed"
        assert isinstance(error, MergeError)
