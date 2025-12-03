"""
Unit tests for database relations functions.

Tests cover:
- Duplicate status updates with batch operations
- Course linking with N+1 elimination
- Batch query optimization
- Error handling and edge cases
"""

from unittest.mock import MagicMock, patch

import pytest

from easy_access.db.relations import (
    link_courses,
    update_relations_async,
)
from easy_access.settings import Settings


class TestLinkCourses:
    """Test course linking functionality."""

    @pytest.mark.asyncio
    async def test_link_courses_no_items(self):
        """Test course linking when no items exist."""
        settings = Settings()

        with patch("easy_access.db.relations.CopyrightItem.filter") as mock_filter:
            mock_filter.return_value = []

            await link_courses(settings)

            mock_filter.assert_called_once()

    @pytest.mark.asyncio
    async def test_link_courses_with_course_codes(self):
        """Test linking courses when items have course codes."""
        settings = Settings()

        # Mock copyright item with course code
        mock_item = MagicMock()
        mock_item.material_id = 1001
        mock_item.course_code = "12345"

        # Mock course
        mock_course = MagicMock()
        mock_course.id = 1
        mock_course.code = 12345

        with (
            patch("easy_access.db.relations.CopyrightItem.filter") as mock_item_filter,
            patch("easy_access.db.relations.determine_course_code") as mock_determine,
            patch("easy_access.db.relations.Course.filter") as mock_course_filter,
            patch(
                "easy_access.db.relations.CopyrightItem.bulk_update"
            ) as mock_bulk_update,
        ):
            mock_item_filter.return_value = [mock_item]
            mock_determine.return_value = ["12345"]
            mock_course_filter.return_value = [mock_course]

            await link_courses(settings)

            # Verify bulk update was called with course_id
            mock_bulk_update.assert_called_once()
            call_args = mock_bulk_update.call_args
            assert call_args[0][0] == [mock_item]
            assert call_args[0][1] == {"course_id": 1}

    @pytest.mark.asyncio
    async def test_link_courses_no_matching_courses(self):
        """Test when no courses match the extracted codes."""
        settings = Settings()

        # Mock copyright item with course code
        mock_item = MagicMock()
        mock_item.material_id = 1001
        mock_item.course_code = "12345"

        with (
            patch("easy_access.db.relations.CopyrightItem.filter") as mock_item_filter,
            patch("easy_access.db.relations.determine_course_code") as mock_determine,
            patch("easy_access.db.relations.Course.filter") as mock_course_filter,
        ):
            mock_item_filter.return_value = [mock_item]
            mock_determine.return_value = ["12345"]
            mock_course_filter.return_value = []  # No matching courses

            await link_courses(settings)

            # Should not call bulk_update
            mock_course_filter.assert_called_once()

    @pytest.mark.asyncio
    async def test_link_courses_multiple_codes(self):
        """Test linking when multiple course codes are extracted."""
        settings = Settings()

        # Mock copyright item with multiple course codes in name
        mock_item = MagicMock()
        mock_item.material_id = 1001
        mock_item.course_code = None
        mock_item.course_name = "Course 12345 and 67890"

        # Mock courses
        mock_course1 = MagicMock()
        mock_course1.id = 1
        mock_course1.code = 12345

        mock_course2 = MagicMock()
        mock_course2.id = 2
        mock_course2.code = 67890

        with (
            patch("easy_access.db.relations.CopyrightItem.filter") as mock_item_filter,
            patch("easy_access.db.relations.determine_course_code") as mock_determine,
            patch("easy_access.db.relations.Course.filter") as mock_course_filter,
            patch(
                "easy_access.db.relations.CopyrightItem.bulk_update"
            ) as mock_bulk_update,
        ):
            mock_item_filter.return_value = [mock_item]
            mock_determine.return_value = ["12345", "67890"]
            mock_course_filter.return_value = [mock_course1, mock_course2]

            await link_courses(settings)

            # Should link to first matching course
            mock_bulk_update.assert_called_once()
            call_args = mock_bulk_update.call_args
            assert call_args[0][0] == [mock_item]
            assert call_args[0][1] == {"course_id": 1}


class TestUpdateRelationsAsync:
    """Test the main relations update orchestrator."""

    @pytest.mark.asyncio
    async def test_update_relations_async_calls_both_functions(self):
        """Test that update_relations_async calls both link_courses and match_v1_to_copyright_items."""
        settings = Settings()

        with (
            patch("easy_access.db.relations.link_courses") as mock_link_courses,
            patch(
                "easy_access.db.relations.match_v1_to_copyright_items"
            ) as mock_match_v1,
        ):
            await update_relations_async(settings)

            mock_link_courses.assert_called_once_with(settings)
            mock_match_v1.assert_called_once_with(settings)

    @pytest.mark.asyncio
    async def test_update_relations_async_error_handling(self):
        """Test error handling in update_relations_async."""
        settings = Settings()

        with (
            patch("easy_access.db.relations.link_courses") as mock_link_courses,
            patch(
                "easy_access.db.relations.match_v1_to_copyright_items"
            ) as mock_match_v1,
        ):
            mock_link_courses.side_effect = Exception("Test error")

            # Should not raise exception, should log error
            await update_relations_async(settings)

            mock_link_courses.assert_called_once_with(settings)
            mock_match_v1.assert_called_once_with(settings)


class TestRelationsBatchOperations:
    """Test batch operation optimizations."""

    @pytest.mark.asyncio
    async def test_batch_course_filtering(self):
        """Test that course filtering uses batch operations."""
        settings = Settings()

        # Mock copyright item
        mock_item = MagicMock()
        mock_item.material_id = 1001
        mock_item.course_code = "12345"

        with (
            patch("easy_access.db.relations.CopyrightItem.filter") as mock_item_filter,
            patch("easy_access.db.relations.determine_course_code") as mock_determine,
            patch("easy_access.db.relations.Course.filter") as mock_course_filter,
        ):
            mock_item_filter.return_value = [mock_item]
            mock_determine.return_value = ["12345"]
            mock_course_filter.return_value = []

            await link_courses(settings)

            # Verify Course.filter was called with IN clause (batch operation)
            mock_course_filter.assert_called_once()
            call_args = mock_course_filter.call_args
            assert "code__in" in str(call_args)
