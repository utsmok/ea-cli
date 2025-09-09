"""
Unit tests for file existence verification functions.

Tests cover:
- Item selection logic based on TTL poli        with patch('easy_access.maintenance.file_existence.select_items_needing_file_check') as mock_select:
            mock_select.return_value = []

            result = await select_items_needing_file_check(settings, force=True)

            assert result == []
- Single file existence checking with Canvas API
- Batch database updates
- Main refresh function with concurrent processing
- Error handling and edge cases
"""

from datetime import datetime
from unittest.mock import AsyncMock, MagicMock, patch

import httpx
import pytest

from easy_access.maintenance.file_existence import (
    check_single_file_existence,
    refresh_file_existence_async,
    select_items_needing_file_check,
    update_file_existence_batch,
)
from easy_access.settings import Settings
from tests.helpers import QuerySetMock


class TestSelectItemsNeedingFileCheck:
    """Test functions for selecting items that need file existence verification."""

    @pytest.mark.asyncio
    async def test_select_items_no_conditions(self):
        """Test selecting items when no TTL conditions are applied."""
        settings = Settings()

        mock_item1 = MagicMock()
        mock_item1.material_id = "123"
        mock_item1.url = "https://example.com/files/123/download"

        mock_item2 = MagicMock()
        mock_item2.material_id = "456"
        mock_item2.url = "https://example.com/files/456/download"

        with patch(
            "easy_access.maintenance.file_existence.CopyrightItem.raw",
            new_callable=AsyncMock,
        ) as mock_raw:
            # Mock the raw query to return the items
            mock_raw.return_value = [mock_item1, mock_item2]

            result = await select_items_needing_file_check(
                settings, ttl_days=None, force=True
            )

            assert len(result) == 2
            assert result[0]["material_id"] == "123"
            assert result[0]["url"] == "https://example.com/files/123/download"
            assert result[1]["material_id"] == "456"
            assert result[1]["url"] == "https://example.com/files/456/download"

            assert len(result) == 2
            assert result[0]["material_id"] == "123"
            assert result[0]["url"] == "https://example.com/files/123/download"
            assert result[1]["material_id"] == "456"
            assert result[1]["url"] == "https://example.com/files/456/download"

    @pytest.mark.asyncio
    async def test_select_items_with_ttl(self):
        """Test selecting items with TTL-based conditions."""
        settings = Settings()

        mock_item = MagicMock()
        mock_item.material_id = "123"
        mock_item.url = "https://example.com/files/123/download"

        with patch(
            "easy_access.maintenance.file_existence.CopyrightItem.raw",
            new_callable=AsyncMock,
        ) as mock_raw:
            mock_raw.return_value = [mock_item]

            result = await select_items_needing_file_check(
                settings, ttl_days=30, force=False
            )

            assert len(result) == 1
            assert result[0]["material_id"] == "123"

    @pytest.mark.asyncio
    async def test_select_items_force_check(self):
        """Test selecting items with force flag."""
        settings = Settings()

        with patch(
            "easy_access.maintenance.file_existence.CopyrightItem.raw",
            new_callable=AsyncMock,
        ) as mock_raw:
            mock_raw.return_value = []

            result = await select_items_needing_file_check(
                settings, ttl_days=30, force=True
            )

            assert result == []

    @pytest.mark.asyncio
    async def test_select_items_batch_size_limit(self):
        """Test that limits the number of returned items."""
        settings = Settings()

        mock_items = []
        for i in range(1500):
            mock_item = MagicMock()
            mock_item.material_id = str(i)
            mock_item.url = f"https://example.com/files/{i}/download"
            mock_items.append(mock_item)

        with patch(
            "easy_access.maintenance.file_existence.CopyrightItem.raw",
            new_callable=AsyncMock,
        ) as mock_raw:
            # Mock should return only up to batch_size items
            mock_raw.return_value = mock_items[:1000]

            result = await select_items_needing_file_check(
                settings, limit=1000, force=True
            )

            assert len(result) == 1000

    @pytest.mark.asyncio
    async def test_select_items_invalid_data(self):
        """Test handling of items with missing material_id or url."""
        settings = Settings()

        mock_item1 = MagicMock()
        mock_item1.material_id = None
        mock_item1.url = "https://example.com/files/123/download"

        mock_item2 = MagicMock()
        mock_item2.material_id = "456"
        mock_item2.url = None

        with patch(
            "easy_access.maintenance.file_existence.CopyrightItem.raw",
            new_callable=AsyncMock,
        ) as mock_raw:
            mock_raw.return_value = [mock_item1, mock_item2]

            result = await select_items_needing_file_check(settings, force=True)

            assert len(result) == 0


class TestCheckSingleFileExistence:
    """Test functions for checking file existence of individual items."""

    @pytest.mark.asyncio
    async def test_check_file_exists_success(self):
        """Test successful file existence check."""
        item_data = {
            "material_id": "123",
            "url": "https://utwente.instructure.com/courses/123/files/456/download",
        }

        mock_response = MagicMock()
        mock_response.status_code = 200

        mock_session = AsyncMock()
        mock_session.get.return_value = mock_response

        result = await check_single_file_existence(item_data, mock_session)

        assert result["material_id"] == "123"
        assert result["file_exists"] is True
        assert isinstance(result["last_canvas_check"], datetime)

    @pytest.mark.asyncio
    async def test_check_file_not_exists(self):
        """Test file existence check when file doesn't exist."""
        item_data = {
            "material_id": "123",
            "url": "https://utwente.instructure.com/courses/123/files/456/download",
        }

        mock_response = MagicMock()
        mock_response.status_code = 404

        mock_session = AsyncMock()
        mock_session.get.return_value = mock_response

        result = await check_single_file_existence(item_data, mock_session)

        assert result["material_id"] == "123"
        assert result["file_exists"] is False
        assert isinstance(result["last_canvas_check"], datetime)

    @pytest.mark.asyncio
    async def test_check_file_invalid_url_format(self):
        """Test handling of invalid URL format."""
        item_data = {"material_id": "123", "url": "https://invalid-url.com"}

        mock_session = AsyncMock()

        result = await check_single_file_existence(item_data, mock_session)

        assert result["material_id"] == "123"
        assert result["file_exists"] is False
        assert isinstance(result["last_canvas_check"], datetime)

    @pytest.mark.asyncio
    async def test_check_file_http_error(self):
        """Test handling of HTTP errors during file check."""
        item_data = {
            "material_id": "123",
            "url": "https://utwente.instructure.com/courses/123/files/456/download",
        }

        mock_session = AsyncMock()
        mock_session.get.side_effect = httpx.RequestError("Connection failed")

        result = await check_single_file_existence(item_data, mock_session)

        assert result["material_id"] == "123"
        assert result["file_exists"] is False
        assert isinstance(result["last_canvas_check"], datetime)

    @pytest.mark.asyncio
    async def test_check_file_url_with_query_params(self):
        """Test file existence check with URL containing query parameters."""
        item_data = {
            "material_id": "123",
            "url": "https://utwente.instructure.com/courses/123/files/456/download?download_frd=1",
        }

        mock_response = MagicMock()
        mock_response.status_code = 200

        mock_session = AsyncMock()
        mock_session.get.return_value = mock_response

        result = await check_single_file_existence(item_data, mock_session)

        assert result["material_id"] == "123"
        assert result["file_exists"] is True

        # Verify the correct file_id was extracted (before query params)
        call_args = mock_session.get.call_args
        assert "456" in call_args[0][0]


class TestUpdateFileExistenceBatch:
    """Test functions for batch updating file existence status."""

    @pytest.mark.asyncio
    async def test_update_batch_empty_results(self):
        """Test batch update with empty results list."""
        await update_file_existence_batch([])

        # Should not raise any errors

    @pytest.mark.asyncio
    async def test_update_batch_single_item(self):
        """Test batch update with single item."""
        results = [
            {
                "material_id": "123",
                "file_exists": True,
                "last_canvas_check": datetime.now(),
            }
        ]

        with patch(
            "easy_access.maintenance.file_existence.CopyrightItem.filter"
        ) as mock_filter:
            mock_update = AsyncMock()
            mock_filter.return_value.update = mock_update

            await update_file_existence_batch(results)

            mock_filter.assert_called_once_with(material_id="123")
            mock_update.assert_called_once()

    @pytest.mark.asyncio
    async def test_update_batch_multiple_items(self):
        """Test batch update with multiple items."""
        results = [
            {
                "material_id": "123",
                "file_exists": True,
                "last_canvas_check": datetime.now(),
            },
            {
                "material_id": "456",
                "file_exists": False,
                "last_canvas_check": datetime.now(),
            },
        ]

        with patch(
            "easy_access.maintenance.file_existence.CopyrightItem.filter"
        ) as mock_filter:
            mock_update = AsyncMock()
            mock_filter.return_value.update = mock_update

            await update_file_existence_batch(results)

            assert mock_filter.call_count == 2
            assert mock_update.call_count == 2


class TestRefreshFileExistenceAsync:
    """Test functions for the main file existence refresh process."""

    @pytest.mark.asyncio
    async def test_refresh_no_api_token(self):
        """Test refresh when no API token is available."""
        settings = Settings()

        with (
            patch("easy_access.maintenance.file_existence.ensure_db_inited"),
            patch("easy_access.maintenance.file_existence.close_connections"),
            patch("easy_access.maintenance.file_existence.getattr") as mock_getattr,
        ):
            # Mock getattr to return None for canvas_api_token
            mock_getattr.return_value = None

            result = await refresh_file_existence_async(settings)

            assert result["error"] == "No API token"
            assert result["checked"] == 0

    @pytest.mark.asyncio
    async def test_refresh_no_items_to_check(self):
        """Test refresh when no items need checking."""
        settings = Settings()

        with (
            patch("easy_access.maintenance.file_existence.ensure_db_inited"),
            patch("easy_access.maintenance.file_existence.close_connections"),
            patch(
                "easy_access.maintenance.file_existence.select_items_needing_file_check"
            ) as mock_select,
            patch("easy_access.maintenance.file_existence.getattr") as mock_getattr,
        ):
            mock_getattr.return_value = "test_token"
            mock_select.return_value = []

            result = await refresh_file_existence_async(settings)

            assert result["checked"] == 0
            assert result["updated"] == 0

    @pytest.mark.asyncio
    async def test_refresh_successful_processing(self):
        """Test successful file existence refresh process."""
        settings = Settings()

        items_to_check = [
            {"material_id": "123", "url": "https://example.com/files/123/download"},
            {"material_id": "456", "url": "https://example.com/files/456/download"},
        ]

        with (
            patch("easy_access.maintenance.file_existence.ensure_db_inited"),
            patch("easy_access.maintenance.file_existence.close_connections"),
            patch(
                "easy_access.maintenance.file_existence.select_items_needing_file_check"
            ) as mock_select,
            patch(
                "easy_access.maintenance.file_existence.update_file_existence_batch"
            ) as mock_update,
            patch("httpx.AsyncClient") as mock_client_class,
            patch("easy_access.maintenance.file_existence.getattr") as mock_getattr,
        ):
            mock_getattr.return_value = "test_token"
            mock_select.return_value = items_to_check

            # Mock HTTP client
            mock_client = AsyncMock()
            mock_client_class.return_value.__aenter__.return_value = mock_client
            mock_client_class.return_value.__aexit__.return_value = None

            # Mock responses for file checks
            mock_response1 = MagicMock()
            mock_response1.status_code = 200
            mock_response2 = MagicMock()
            mock_response2.status_code = 404

            mock_client.get.side_effect = [mock_response1, mock_response2]

            result = await refresh_file_existence_async(settings, max_concurrent=2)

            assert result["checked"] == 2
            assert result["exists"] == 1
            assert result["not_exists"] == 1
            assert "duration_seconds" in result

            mock_update.assert_called_once()

    @pytest.mark.asyncio
    async def test_refresh_with_concurrency_limit(self):
        """Test refresh with concurrency limiting."""
        settings = Settings()

        items_to_check = [
            {"material_id": str(i), "url": f"https://example.com/files/{i}/download"}
            for i in range(10)
        ]

        with (
            patch("easy_access.maintenance.file_existence.ensure_db_inited"),
            patch("easy_access.maintenance.file_existence.close_connections"),
            patch(
                "easy_access.maintenance.file_existence.select_items_needing_file_check"
            ) as mock_select,
            patch(
                "easy_access.maintenance.file_existence.update_file_existence_batch"
            ),
            patch("httpx.AsyncClient") as mock_client_class,
            patch("easy_access.maintenance.file_existence.getattr") as mock_getattr,
        ):
            mock_getattr.return_value = "test_token"
            mock_select.return_value = items_to_check

            mock_client = AsyncMock()
            mock_client_class.return_value.__aenter__.return_value = mock_client
            mock_client_class.return_value.__aexit__.return_value = None

            # Mock responses
            mock_responses = [MagicMock() for _ in range(10)]
            for resp in mock_responses:
                resp.status_code = 200
            mock_client.get.side_effect = mock_responses

            result = await refresh_file_existence_async(settings, max_concurrent=3)

            assert result["checked"] == 10
            assert result["exists"] == 10
            assert result["not_exists"] == 0

    @pytest.mark.asyncio
    async def test_refresh_with_force_flag(self):
        """Test refresh with force flag enabled."""
        settings = Settings()

        with (
            patch("easy_access.maintenance.file_existence.ensure_db_inited"),
            patch("easy_access.maintenance.file_existence.close_connections"),
            patch(
                "easy_access.maintenance.file_existence.select_items_needing_file_check"
            ) as mock_select,
            patch("easy_access.maintenance.file_existence.getattr") as mock_getattr,
        ):
            mock_getattr.return_value = "test_token"
            mock_select.return_value = []

            await refresh_file_existence_async(settings, force=True)

            # Verify force flag was passed to select function
            mock_select.assert_called_once()
            call_args = mock_select.call_args
            assert call_args[0][3] is True  # force is the 4th positional argument

    @pytest.mark.asyncio
    async def test_refresh_with_custom_batch_size(self):
        """Test refresh with custom batch size."""
        settings = Settings()

        with (
            patch("easy_access.maintenance.file_existence.ensure_db_inited"),
            patch("easy_access.maintenance.file_existence.close_connections"),
            patch(
                "easy_access.maintenance.file_existence.select_items_needing_file_check"
            ) as mock_select,
            patch("easy_access.maintenance.file_existence.getattr") as mock_getattr,
        ):
            mock_getattr.return_value = "test_token"
            mock_select.return_value = []

            await refresh_file_existence_async(settings, batch_size=500)

            # Verify batch size was passed to select function
            mock_select.assert_called_once()
            call_args = mock_select.call_args
            assert call_args[0][2] == 500  # batch_size is the 3rd positional argument


class TestFileExistenceIntegration:
    """Integration tests for file existence verification."""

    @pytest.mark.asyncio
    async def test_end_to_end_file_existence_check(self):
        """Test complete file existence verification workflow."""
        settings = Settings()

        # Mock database items
        mock_item = MagicMock()
        mock_item.material_id = "123"
        mock_item.url = "https://utwente.instructure.com/courses/123/files/456/download"

        with (
            patch("easy_access.maintenance.file_existence.ensure_db_inited"),
            patch("easy_access.maintenance.file_existence.close_connections"),
            patch(
                "easy_access.maintenance.file_existence.CopyrightItem.raw",
                new_callable=AsyncMock,
            ) as mock_raw,
            patch(
                "easy_access.maintenance.file_existence.CopyrightItem.filter"
            ) as mock_filter,
            patch("httpx.AsyncClient") as mock_client_class,
            patch("easy_access.maintenance.file_existence.getattr") as mock_getattr,
        ):
            mock_getattr.side_effect = (
                lambda obj, attr, default=None: "test_token"
                if attr == "canvas_api_token"
                else getattr(obj, attr, default)
                if hasattr(obj, attr)
                else default
            )

            mock_raw.return_value = [mock_item]

            # Mock HTTP client and response
            mock_client = AsyncMock()
            mock_client_class.return_value.__aenter__.return_value = mock_client
            mock_client_class.return_value.__aexit__.return_value = None

            mock_response = MagicMock()
            mock_response.status_code = 200
            mock_client.get.return_value = mock_response

            # Mock database update
            mock_filter.return_value = QuerySetMock()
            mock_filter.return_value.update = AsyncMock()

            # Execute the workflow
            result = await refresh_file_existence_async(settings)

            # Verify results
            assert result["checked"] == 1
            assert result["exists"] == 1
            assert result["not_exists"] == 0

            # Verify database was updated (update is on the filter() return value)
            filter_ret = mock_filter.return_value
            assert hasattr(filter_ret, "update")
            filter_ret.update.assert_awaited()

    @pytest.mark.asyncio
    async def test_error_handling_in_concurrent_processing(self):
        """Test error handling during concurrent file checking."""
        settings = Settings()

        items_to_check = [
            {"material_id": "123", "url": "https://example.com/files/123/download"},
            {"material_id": "456", "url": "https://example.com/files/456/download"},
        ]

        with (
            patch("easy_access.maintenance.file_existence.ensure_db_inited"),
            patch("easy_access.maintenance.file_existence.close_connections"),
            patch(
                "easy_access.maintenance.file_existence.select_items_needing_file_check"
            ) as mock_select,
            patch(
                "easy_access.maintenance.file_existence.update_file_existence_batch"
            ),
            patch("httpx.AsyncClient") as mock_client_class,
            patch("easy_access.maintenance.file_existence.getattr") as mock_getattr,
        ):
            mock_getattr.return_value = "test_token"
            mock_select.return_value = items_to_check

            mock_client = AsyncMock()
            mock_client_class.return_value.__aenter__.return_value = mock_client
            mock_client_class.return_value.__aexit__.return_value = None

            # Mock one success and one failure
            mock_response1 = MagicMock()
            mock_response1.status_code = 200
            mock_response2 = MagicMock()
            mock_response2.status_code = 500

            mock_client.get.side_effect = [mock_response1, mock_response2]

            result = await refresh_file_existence_async(settings)

            # Should still process both items despite one error
            assert result["checked"] == 2
            assert result["exists"] == 1
            assert result["not_exists"] == 1
