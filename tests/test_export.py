"""
Unit tests for export functions in easy_access.sheets.export module.

Tests cover:
- Faculty data gathering
- Faculty sheet export
- Programme sheet export
- All items sheet export
- Faculty overview export
- Main export orchestrator
- File uniqueness handling
"""

import asyncio
import tempfile
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

import polars as pl
import pytest

from easy_access.sheets.export import (
    _get_unique_filepath,
    export_all_items_sheet,
    export_faculty_overviews,
    export_faculty_sheets,
    export_programme_sheets,
    export_reports_async,
    gather_faculty_data,
)
from easy_access.settings import DirSetting, Settings


class TestGatherFacultyData:
    """Test faculty data gathering functionality."""

    @pytest.mark.asyncio
    async def test_gather_faculty_data_success(self):
        """Test successful gathering of faculty data."""
        settings = Settings()
        settings = Settings()
        # Mock data
        mock_data = pl.DataFrame({
            "faculty": ["Faculty A", "Faculty A", "Faculty B", "Faculty B"],
            "title": ["Item 1", "Item 2", "Item 3", "Item 4"],
            "department": ["Dept 1", "Dept 1", "Dept 2", "Dept 2"]
        })

        with patch("easy_access.sheets.export.retrieve_full_data", return_value=mock_data):
            result = await gather_faculty_data(settings)

            assert len(result) == 2
            assert "Faculty A" in result
            assert "Faculty B" in result
            assert result["Faculty A"].shape[0] == 2
            assert result["Faculty B"].shape[0] == 2

    @pytest.mark.asyncio
    async def test_gather_faculty_data_empty(self):
        """Test gathering faculty data when no data exists."""
        settings = Settings()
        mock_data = pl.DataFrame()

        with patch("easy_access.sheets.export.retrieve_full_data", return_value=mock_data):
            result = await gather_faculty_data(settings)

            assert result == {}

    @pytest.mark.asyncio
    async def test_gather_faculty_data_filters_unmapped(self):
        """Test that unmapped faculty is filtered out."""
        settings = Settings()
        mock_data = pl.DataFrame({
            "faculty": ["Faculty A", "Unmapped", "Faculty B"],
            "title": ["Item 1", "Item 2", "Item 4"]
        })

        with patch("easy_access.sheets.export.retrieve_full_data", return_value=mock_data):
            result = await gather_faculty_data(settings)

            assert len(result) == 2
            assert "Faculty A" in result
            assert "Faculty B" in result
            assert "Unmapped" not in result


class TestExportFacultySheets:
    """Test faculty sheet export functionality."""

    @pytest.mark.asyncio
    async def test_export_faculty_sheets_success(self):
        """Test successful export of faculty sheets."""
        settings = Settings()
        faculty_data = {
            "Faculty A": pl.DataFrame({
                "title": ["Item 1", "Item 2"],
                "faculty": ["Faculty A", "Faculty A"]
            })
        }

        with patch("easy_access.sheets.export.store_complete_data") as mock_store, \
             patch("easy_access.sheets.export.finalize_sheet") as mock_finalize, \
             patch("easy_access.sheets.export._get_unique_filepath") as mock_get_path, \
             patch("pathlib.Path.mkdir") as mock_mkdir:

            mock_get_path.return_value = Path("/tmp/test.xlsx")
            mock_finalize.return_value = 10

            result = await export_faculty_sheets(settings, faculty_data, 9)

            assert result == 10
            mock_store.assert_called_once()
            mock_finalize.assert_called_once()

    @pytest.mark.asyncio
    async def test_export_faculty_sheets_empty_data(self):
        """Test export with empty faculty data."""
        settings = Settings()
        faculty_data = {
            "Faculty A": pl.DataFrame()
        }

        with patch("easy_access.sheets.export.store_complete_data") as mock_store:
            result = await export_faculty_sheets(settings, faculty_data, 9)

            assert result == 9
            mock_store.assert_not_called()


class TestExportProgrammeSheets:
    """Test programme sheet export functionality."""

    @pytest.mark.asyncio
    async def test_export_programme_sheets_success(self):
        """Test successful export of programme sheets."""
        settings = Settings()
        faculty_data = {
            "Faculty A": pl.DataFrame({
                "title": ["Item 1", "Item 2"],
                "faculty": ["Faculty A", "Faculty A"],
                "department": ["Course 1", "Course 1"]
            })
        }

        # Mock course mapping
        mock_course_mapping = MagicMock()
        mock_course_mapping.get.return_value = {"Course 1": "Programme 1"}
        with patch.object(settings, 'university_settings', MagicMock(course_mapping=mock_course_mapping)), \
             patch("easy_access.sheets.export.store_complete_data") as mock_store, \
             patch("easy_access.sheets.export.finalize_sheet") as mock_finalize, \
             patch("easy_access.sheets.export._get_unique_filepath") as mock_get_path, \
             patch("pathlib.Path.mkdir") as mock_mkdir:

            mock_get_path.return_value = Path("/tmp/test.xlsx")
            mock_finalize.return_value = 10

            result = await export_programme_sheets(settings, faculty_data, 9)

            assert result == 10
            mock_store.assert_called_once()
            mock_finalize.assert_called_once()

    @pytest.mark.asyncio
    async def test_export_programme_sheets_no_mapping(self):
        """Test export when no course mapping exists."""
        settings = Settings()
        faculty_data = {
            "Faculty A": pl.DataFrame({
                "title": ["Item 1"],
                "faculty": ["Faculty A"],
                "department": ["Course 1"]
            })
        }

        # No course mapping
        mock_course_mapping = MagicMock()
        mock_course_mapping.get.return_value = None
        with patch.object(settings, 'university_settings', MagicMock(course_mapping=mock_course_mapping)), \
             patch("easy_access.sheets.export.store_complete_data") as mock_store:
            result = await export_programme_sheets(settings, faculty_data, 9)

            assert result == 9
            mock_store.assert_not_called()

    @pytest.mark.asyncio
    async def test_export_programme_sheets_no_department_column(self):
        """Test export when department column is missing."""
        settings = Settings()
        faculty_data = {
            "Faculty A": pl.DataFrame({
                "title": ["Item 1"],
                "faculty": ["Faculty A"]
            })
        }

        mock_course_mapping = MagicMock()
        mock_course_mapping.get.return_value = {"Course 1": "Programme 1"}
        with patch.object(settings, 'university_settings', MagicMock(course_mapping=mock_course_mapping)), \
             patch("easy_access.sheets.export.store_complete_data") as mock_store:
            result = await export_programme_sheets(settings, faculty_data, 9)

            assert result == 9
            mock_store.assert_not_called()


class TestExportAllItemsSheet:
    """Test all items sheet export functionality."""

    @pytest.mark.asyncio
    async def test_export_all_items_sheet_success(self):
        """Test successful export of all items sheet."""
        settings = Settings()
        mock_data = pl.DataFrame({
            "title": ["Item 1", "Item 2"],
            "faculty": ["Faculty A", "Faculty B"]
        })

        with patch("easy_access.sheets.export.retrieve_full_data", return_value=mock_data), \
             patch("easy_access.sheets.export.store_complete_data") as mock_store, \
             patch("easy_access.sheets.export.finalize_sheet") as mock_finalize, \
             patch("easy_access.sheets.export._get_unique_filepath") as mock_get_path, \
             patch("pathlib.Path.mkdir") as mock_mkdir:

            mock_get_path.return_value = Path("/tmp/test.xlsx")
            mock_finalize.return_value = 10

            result = await export_all_items_sheet(settings, 9)

            assert result == 10
            mock_store.assert_called_once()
            mock_finalize.assert_called_once()

    @pytest.mark.asyncio
    async def test_export_all_items_sheet_empty_data(self):
        """Test export when no data exists."""
        settings = Settings()
        mock_data = pl.DataFrame()

        with patch("easy_access.sheets.export.retrieve_full_data", return_value=mock_data), \
             patch("easy_access.sheets.export.store_complete_data") as mock_store:

            result = await export_all_items_sheet(settings, 9)

            assert result == 9
            mock_store.assert_not_called()


class TestExportFacultyOverviews:
    """Test faculty overview export functionality."""

    @pytest.mark.asyncio
    async def test_export_faculty_overviews_success(self):
        """Test successful export of faculty overviews."""
        settings = Settings()
        faculty_data = {
            "Faculty A": pl.DataFrame({
                "title": ["Item 1"],
                "faculty": ["Faculty A"]
            })
        }

        with patch("easy_access.sheets.export.create_faculty_overviews") as mock_create:
            mock_create.return_value = 10

            result = await export_faculty_overviews(settings, faculty_data, 9)

            assert result == 10
            mock_create.assert_called_once_with(
                settings=settings,
                faculty_data=faculty_data,
                style_iter=9,
                disable_writes=False
            )


class TestExportReportsAsync:
    """Test main export orchestrator functionality."""

    @pytest.mark.asyncio
    async def test_export_reports_async_success(self):
        """Test successful execution of main export orchestrator."""
        settings = Settings()
        faculty_data = {
            "Faculty A": pl.DataFrame({
                "title": ["Item 1"],
                "faculty": ["Faculty A"]
            })
        }

        with patch("easy_access.sheets.export.gather_faculty_data", return_value=faculty_data), \
             patch("easy_access.sheets.export.export_faculty_sheets", return_value=10) as mock_faculty, \
             patch("easy_access.sheets.export.export_programme_sheets", return_value=11) as mock_programme, \
             patch("easy_access.sheets.export.export_all_items_sheet", return_value=12) as mock_all_items, \
             patch("easy_access.sheets.export.export_faculty_overviews", return_value=13) as mock_overviews:

            await export_reports_async(settings)

            mock_faculty.assert_called_once()
            mock_programme.assert_called_once()
            mock_all_items.assert_called_once()
            mock_overviews.assert_called_once()

    @pytest.mark.asyncio
    async def test_export_reports_async_no_data(self):
        """Test export orchestrator when no faculty data exists."""
        settings = Settings()
        with patch("easy_access.sheets.export.gather_faculty_data", return_value={}), \
             patch("easy_access.sheets.export.export_faculty_sheets") as mock_faculty:

            await export_reports_async(settings)

            mock_faculty.assert_not_called()


class TestGetUniqueFilepath:
    """Test file uniqueness functionality."""

    def test_get_unique_filepath_no_conflict(self, tmp_path):
        """Test filepath generation when no conflict exists."""
        filename_base = "test_file"
        expected_path = tmp_path / "test_file.xlsx"

        result = _get_unique_filepath(tmp_path, filename_base)

        assert result == expected_path

    def test_get_unique_filepath_with_conflict(self, tmp_path):
        """Test filepath generation when file already exists."""
        # Create existing file
        existing_file = tmp_path / "test_file.xlsx"
        existing_file.touch()

        filename_base = "test_file"
        expected_path = tmp_path / "test_file_1.xlsx"

        result = _get_unique_filepath(tmp_path, filename_base)

        assert result == expected_path

    def test_get_unique_filepath_multiple_conflicts(self, tmp_path):
        """Test filepath generation with multiple existing files."""
        # Create multiple existing files
        (tmp_path / "test_file.xlsx").touch()
        (tmp_path / "test_file_1.xlsx").touch()
        (tmp_path / "test_file_2.xlsx").touch()

        filename_base = "test_file"
        expected_path = tmp_path / "test_file_3.xlsx"

        result = _get_unique_filepath(tmp_path, filename_base)

        assert result == expected_path


class TestExportIntegration:
    """Integration tests for export functionality."""

    @pytest.mark.asyncio
    async def test_export_pipeline_integration(self):
        """Test the complete export pipeline integration."""
        settings = Settings()
        # Mock comprehensive data
        mock_data = pl.DataFrame({
            "faculty": ["Faculty A", "Faculty A", "Faculty B"],
            "title": ["Item 1", "Item 2", "Item 3"],
            "department": ["Course 1", "Course 1", "Course 2"]
        })

        # Mock course mapping
        mock_course_mapping = MagicMock()
        mock_course_mapping.get.side_effect = lambda faculty: {
            "Faculty A": {"Course 1": "Programme 1"},
            "Faculty B": {"Course 2": "Programme 2"}
        }.get(faculty)
        with patch.object(settings, 'university_settings', MagicMock(course_mapping=mock_course_mapping)), \
             patch("easy_access.sheets.export.retrieve_full_data", return_value=mock_data), \
             patch("easy_access.sheets.export.store_complete_data") as mock_store, \
             patch("easy_access.sheets.export.finalize_sheet") as mock_finalize, \
             patch("easy_access.sheets.export.create_faculty_overviews") as mock_overviews, \
             patch("easy_access.sheets.export._get_unique_filepath") as mock_get_path, \
             patch("pathlib.Path.mkdir") as mock_mkdir:

            mock_get_path.return_value = Path("/tmp/test.xlsx")
            mock_finalize.return_value = 10
            mock_overviews.return_value = 11

            await export_reports_async(settings)

            # Verify multiple calls for different sheet types
            assert mock_store.call_count >= 3  # faculty + programme + all items
            assert mock_finalize.call_count >= 3
            mock_overviews.assert_called_once()
