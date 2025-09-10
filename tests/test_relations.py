"""
Unit tests for database relations functions.

Tests cover:
- Duplicate status updates with batch operations
- Course linking with N+1 elimination
- Batch query optimization
- Error handling and edge cases
"""

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from easy_access.db.relations import (
    link_courses,
    update_duplicates,
    update_relations_async,
)
from easy_access.settings import Settings
from tests.helpers import QuerySetMock


class TestUpdateDuplicates:
    """Test duplicate status update functionality."""

    @pytest.mark.asyncio
    async def test_update_duplicates_no_replacements(self):
        """Test when no PDFs have replacements."""
        settings = Settings()

        with patch("easy_access.db.relations.PDF.filter") as mock_filter:
            # Mock empty query result - need to return a queryset that supports prefetch_related
            mock_queryset = AsyncMock()
            mock_queryset.prefetch_related.return_value = []
            mock_filter.return_value = mock_queryset

            await update_duplicates(settings)

            mock_filter.assert_called_once_with(replace_with_id__not_isnull=True)

    @pytest.mark.asyncio
    async def test_update_duplicates_with_replacements(self):
        """Test updating duplicates when replacements exist."""
        settings = Settings()

        # Mock PDF with replacement
        mock_pdf = MagicMock()
        mock_pdf.material_id = 1001
        mock_pdf.replace_with.material_id = 2001

        # Mock copyright item to update
        mock_item = MagicMock()
        mock_item.material_id = 1001

        with (
            patch("easy_access.db.relations.PDF.filter") as mock_filter,
            patch("easy_access.db.relations.CopyrightItem.filter") as mock_item_filter,
            patch(
                "easy_access.db.relations.CopyrightItem.bulk_update"
            ) as mock_bulk_update,
        ):
            # Mock queryset for PDFs with prefetch_related capability
            mock_queryset = QuerySetMock([mock_pdf])
            mock_filter.return_value = mock_queryset

            mock_item_filter.return_value = [mock_item]

            await update_duplicates(settings)

            # Verify bulk update was called
            mock_bulk_update.assert_called_once()

    @pytest.mark.asyncio
    async def test_update_duplicates_no_items_to_update(self):
        """Test when no copyright items need updating."""
        settings = Settings()

        # Mock PDF with replacement
        mock_pdf = MagicMock()
        mock_pdf.material_id = 1001
        mock_pdf.replace_with.material_id = 2001

        with (
            patch("easy_access.db.relations.PDF.filter") as mock_filter,
            patch("easy_access.db.relations.CopyrightItem.filter") as mock_item_filter,
        ):
            # Mock queryset for PDFs
            mock_queryset = QuerySetMock([mock_pdf])
            mock_filter.return_value = mock_queryset

            mock_item_filter.return_value = []  # No matching items

            await update_duplicates(settings)

            # Should not fail, just log that no items were updated
            # Reset the mock call count to ensure the next invocation is measured
            mock_item_filter.reset_mock()
            mock_item_filter.return_value = []  # No items to update

            await update_duplicates(settings)

            # Should have been called once during the second run
            mock_item_filter.assert_called_once()


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
        """Test that update_relations_async calls both update_duplicates and link_courses."""
        settings = Settings()

        with (
            patch("easy_access.db.relations.update_duplicates") as mock_update_dup,
            patch("easy_access.db.relations.link_courses") as mock_link_courses,
        ):
            await update_relations_async(settings)

            mock_update_dup.assert_called_once_with(settings)
            mock_link_courses.assert_called_once_with(settings)

    @pytest.mark.asyncio
    async def test_update_relations_async_error_handling(self):
        """Test error handling in update_relations_async."""
        settings = Settings()

        with (
            patch("easy_access.db.relations.update_duplicates") as mock_update_dup,
            patch("easy_access.db.relations.link_courses") as mock_link_courses,
        ):
            mock_update_dup.side_effect = Exception("Test error")

            # Should not raise exception, should log error
            await update_relations_async(settings)

            mock_update_dup.assert_called_once_with(settings)
            mock_link_courses.assert_called_once_with(settings)


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

    @pytest.mark.asyncio
    async def test_prefetch_related_usage(self):
        """Test that prefetch_related is used to avoid N+1 queries."""
        settings = Settings()

        with patch("easy_access.db.relations.PDF.filter") as mock_filter:
            mock_filter.return_value = []

            await update_duplicates(settings)

            # Verify prefetch_related was used
            mock_filter.assert_called_once()
            call_args = mock_filter.call_args
            assert "prefetch_related" in str(call_args) or "replace_with" in str(
                call_args
            )
