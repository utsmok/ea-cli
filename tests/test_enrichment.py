"""
Unit tests for enrichment functions.

Tests cover:
- Course/person gathering and selection logic
- HTML parsing and data extraction
- Stale data detection
- Mock HTTP requests for external API testing
"""

import pytest
import httpx
from unittest.mock import AsyncMock, MagicMock, patch
from datetime import datetime, timedelta
import asyncio

from easy_access.enrichment.osiris import (
    gather_target_course_codes,
    select_missing_or_stale_courses,
    gather_target_person_names,
    select_missing_or_stale_persons,
    fetch_course_data,
    fetch_person_data,
    persist_courses,
    persist_persons,
    enrich_async,
)
from easy_access.db.models import CopyrightItem, Course, Person
from easy_access.settings import Settings


class TestEnrichmentGathering:
    """Test functions for gathering target data for enrichment."""

    @pytest.mark.asyncio
    async def test_gather_target_course_codes_empty_db(self, setup_test_db):
        """Test gathering course codes when database is empty."""
        settings = Settings()

        with patch('easy_access.enrichment.osiris.CopyrightItem.all') as mock_all:
            # Create a proper async mock that returns an empty list
            mock_queryset = AsyncMock()
            mock_queryset.distinct.return_value = []
            mock_all.return_value = mock_queryset

            result = await gather_target_course_codes(settings)
            assert result == set()
            mock_all.assert_called_once()

    @pytest.mark.asyncio
    async def test_gather_target_course_codes_with_data(self, setup_test_db):
        """Test gathering course codes with actual data."""
        settings = Settings()

        # Mock copyright items with course codes
        mock_item1 = MagicMock()
        mock_item1.course_code = "123456"
        mock_item1.course_name = "Test Course"

        mock_item2 = MagicMock()
        mock_item2.course_code = "789012"
        mock_item2.course_name = "Another Course"

        with patch('easy_access.enrichment.osiris.CopyrightItem.all') as mock_all, \
             patch('easy_access.enrichment.osiris.determine_course_code') as mock_determine:
            # Setup the queryset mock
            mock_queryset = AsyncMock()
            mock_queryset.distinct.return_value = [mock_item1, mock_item2]
            mock_all.return_value = mock_queryset

            # Mock determine_course_code to return course codes
            mock_determine.side_effect = [["123456"], ["789012"]]

            result = await gather_target_course_codes(settings)

            assert result == {123456, 789012}
            mock_all.assert_called_once()
            assert mock_determine.call_count == 2

    @pytest.mark.asyncio
    async def test_gather_target_person_names(self, setup_test_db):
        """Test gathering person names from copyright items."""
        settings = Settings()

        # Mock copyright items with person data
        mock_item1 = MagicMock()
        mock_item1.author = "John Doe"
        mock_item1.auditor = None

        mock_item2 = MagicMock()
        mock_item2.author = None
        mock_item2.auditor = "Jane Smith"

        with patch('easy_access.enrichment.osiris.CopyrightItem.all') as mock_all:
            mock_queryset = AsyncMock()
            mock_queryset.distinct.return_value = [mock_item1, mock_item2]
            mock_all.return_value = mock_queryset

            result = await gather_target_person_names(settings)
            expected = {"John Doe", "Jane Smith"}
            assert result == expected


class TestEnrichmentSelection:
    """Test functions for selecting stale or missing data."""

    @pytest.mark.asyncio
    async def test_select_missing_or_stale_courses_no_ttl(self, setup_test_db):
        """Test selecting courses when no TTL is set."""
        settings = Settings()
        course_codes = {12345, 67890}

            with patch('easy_access.enrichment.osiris.Course.filter', new=AsyncMock(return_value=[])):
                # Mock that no courses exist, so all are missing
                result = await select_missing_or_stale_courses(settings, course_codes, None)
                assert result == course_codes

    @pytest.mark.asyncio
    async def test_select_missing_or_stale_courses_with_ttl(self, setup_test_db):
        """Test selecting courses based on TTL."""
        settings = Settings()
        course_codes = {12345, 67890}
        ttl_days = 30

        # Mock existing course that's fresh
        mock_fresh_course = MagicMock()
        mock_fresh_course.cursuscode = 12345
        mock_fresh_course.modified_at = datetime.now() - timedelta(days=10)

        # Mock existing course that's stale
        mock_stale_course = MagicMock()
        mock_stale_course.cursuscode = 67890
        mock_stale_course.modified_at = datetime.now() - timedelta(days=40)

            with patch('easy_access.enrichment.osiris.Course.filter', new=AsyncMock(return_value=[mock_fresh_course, mock_stale_course])):
                result = await select_missing_or_stale_courses(settings, course_codes, ttl_days)
                # Only the stale course should be returned, fresh course is excluded
                assert result == {67890}

    @pytest.mark.asyncio
    async def test_select_missing_or_stale_persons(self, setup_test_db):
        """Test selecting persons based on TTL."""
        settings = Settings()
        person_names = {"John Doe", "Jane Smith"}
        ttl_days = 30

        # Mock existing person that's stale
        mock_stale_person = MagicMock()
        mock_stale_person.input_name = "John Doe"
        mock_stale_person.modified_at = datetime.now() - timedelta(days=40)

            with patch('easy_access.enrichment.osiris.Person.filter', new=AsyncMock(return_value=[mock_stale_person])):
                result = await select_missing_or_stale_persons(settings, person_names, ttl_days)
                # John Doe is stale, Jane Smith is missing
                assert result == {"John Doe", "Jane Smith"}


class TestEnrichmentFetching:
    """Test functions for fetching data from external APIs."""

    @pytest.mark.asyncio
    async def test_fetch_course_data_success(self, setup_test_db):
        """Test successful course data fetching."""
        course_code = 12345
        mock_client = AsyncMock()

        # Mock successful HTTP response from OSIRIS API
        mock_response = AsyncMock()
        mock_response.status_code = 200
        # make .json synchronous to match usage in code (response.json().get(...))
        mock_response.json = MagicMock(return_value={
            "hits": {
                "hits": [
                    {
                        "_source": {
                            "id_cursus": "INT123",
                            "collegejaar": "2024-2025",
                            "cursus_korte_naam": "TEST",
                            "cursus_lange_naam": "Test Course",
                            "faculteit": "EEMCS",
                            "faculteit_naam": "Electrical Engineering, Mathematics and Computer Science",
                            "coordinerend_onderdeel_oms": "Computer Science",
                            "punten": 5.0,
                            "voertalen": [{"voertaal_omschrijving": "English"}],
                            "opmerking_cursus": "Test course notes",
                            "categorie_omschrijving": "Bachelor",
                            "docenten": ["Dr. John Doe", "Prof. Jane Smith"]
                        }
                    }
                ]
            }
        }

        # Mock the detailed course info call (empty for simplicity)
        mock_details_response = AsyncMock()
        mock_details_response.status_code = 200
        mock_details_response.json.return_value = {"items": []}

        mock_client.post.return_value = mock_response
        mock_client.get.return_value = mock_details_response

        result = await fetch_course_data(course_code, mock_client)

        assert result['cursuscode'] == 12345
        assert result['name'] == 'Test Course'
        assert result['short_name'] == 'TEST'
        assert result['faculty'] == 'EEMCS'
        assert isinstance(result['teachers'], set)
        mock_client.post.assert_called_once()

    @pytest.mark.asyncio
    async def test_fetch_course_data_not_found(self, setup_test_db):
        """Test course data fetching when course not found."""
        course_code = 99999
        mock_client = AsyncMock()

        # Mock empty response (no results)
        mock_response = AsyncMock()
        mock_response.status_code = 200
        mock_response.json = MagicMock(return_value={
            "hits": {
                "hits": []
            }
        }
        mock_client.post.return_value = mock_response

        result = await fetch_course_data(course_code, mock_client)

        assert result == {}
        mock_client.post.assert_called_once()

    @pytest.mark.asyncio
    async def test_fetch_person_data_success(self, setup_test_db):
        """Test successful person data fetching."""
        person_name = "John Doe"
        mock_client = AsyncMock()

        # Mock search results page
        search_response = AsyncMock()
        search_response.status_code = 200
        search_response.text = '<html><body><a data-link="john.doe">John Doe</a></body></html>'

        # Mock person detail page
        detail_response = AsyncMock()
        detail_response.status_code = 200
        detail_response.text = '''
        <html>
        <body>
        <h1 class="pageheader__title">John Doe</h1>
        <a href="mailto:john.doe@utwente.nl">john.doe@utwente.nl</a>
        <div class="widget-linklist--smallicons">
            <div class="widget-linklist__text">Computer Science (EEMCS)</div>
        </div>
        <div id="tabpanel-education">
            <a href="https://utwente.osiris-student.nl/course/12345">CS101 - Introduction to CS</a>
            <a href="https://www.utwente.nl/programme/cs">Computer Science Master</a>
        </div>
        </body>
        </html>
        '''

        mock_client.get.side_effect = [search_response, detail_response]

        with patch('easy_access.enrichment.osiris.Levenshtein.ratio', return_value=0.95):
            result = await fetch_person_data(person_name, mock_client)

        assert result['input_name'] == 'John Doe'
        assert result['main_name'] == 'John Doe'
        assert result['email'] == 'john.doe@utwente.nl'
        assert result['faculty'] == 'EEMCS'
        assert len(result['courses']) == 1
        assert result['courses'][0]['course_code'] == 'CS101'
        assert mock_client.get.call_count == 2

    @pytest.mark.asyncio
    async def test_fetch_person_data_not_found(self, setup_test_db):
        """Test person data fetching when person not found."""
        person_name = "Unknown Person"
        mock_client = AsyncMock()

        # Mock empty search results
        search_response = AsyncMock()
        search_response.status_code = 200
        search_response.text = '<html><body>No results found</body></html>'
        mock_client.get.return_value = search_response

        result = await fetch_person_data(person_name, mock_client)

        assert result == {}
        mock_client.get.assert_called_once()


class TestEnrichmentPersistence:
    """Test functions for persisting enrichment data."""

    @pytest.mark.asyncio
    async def test_persist_courses_new(self, setup_test_db):
        """Test persisting new courses."""
        courses_data = {
            12345: {
                'cursuscode': 12345,
                'name': 'Test Course',
                'short_name': 'TEST',
                'faculty': 'EEMCS'
            }
        }

        with patch('easy_access.enrichment.osiris.Course.filter') as mock_filter, \
             patch('easy_access.enrichment.osiris.Course.bulk_create') as mock_create:

            # No existing courses
            mock_filter.return_value = AsyncMock()
            mock_filter.return_value.__aiter__.return_value = []

            await persist_courses(courses_data)

            mock_create.assert_called_once()
            created_courses = mock_create.call_args[0][0]  # First positional arg
            assert len(created_courses) == 1

    @pytest.mark.asyncio
    async def test_persist_courses_update(self, setup_test_db):
        """Test updating existing courses."""
        courses_data = {
            12345: {
                'cursuscode': 12345,
                'name': 'Updated Course Name',
                'short_name': 'TEST'
            }
        }

        # Mock existing course
        mock_existing_course = MagicMock()
        mock_existing_course.cursuscode = 12345

        with patch('easy_access.enrichment.osiris.Course.filter') as mock_filter, \
             patch('easy_access.enrichment.osiris.Course.bulk_update') as mock_update:

            mock_filter.return_value = AsyncMock()
            mock_filter.return_value.__aiter__.return_value = [mock_existing_course]

            await persist_courses(courses_data)

            # Should call update for existing courses
            mock_update.assert_called_once()

    @pytest.mark.asyncio
    async def test_persist_persons_new(self, setup_test_db):
        """Test persisting new persons."""
        persons_data = {
            "John Doe": {
                'input_name': 'John Doe',
                'main_name': 'John Doe',
                'email': 'john@utwente.nl',
                'faculty': 'EEMCS'
            }
        }

        with patch('easy_access.enrichment.osiris.Person.filter') as mock_filter, \
             patch('easy_access.enrichment.osiris.Person.bulk_create') as mock_create:

            # No existing persons
            mock_filter.return_value = AsyncMock()
            mock_filter.return_value.__aiter__.return_value = []

            await persist_persons(persons_data)

            mock_create.assert_called_once()
            created_persons = mock_create.call_args[0][0]  # First positional arg
            assert len(created_persons) == 1

    @pytest.mark.asyncio
    async def test_persist_persons_update(self, setup_test_db):
        """Test updating existing persons."""
        persons_data = {
            "John Doe": {
                'input_name': 'John Doe',
                'main_name': 'Dr. John Doe',
                'email': 'john@utwente.nl'
            }
        }

        # Mock existing person
        mock_existing_person = MagicMock()
        mock_existing_person.input_name = "John Doe"

        with patch('easy_access.enrichment.osiris.Person.filter') as mock_filter, \
             patch('easy_access.enrichment.osiris.Person.bulk_update') as mock_update:

            mock_filter.return_value = AsyncMock()
            mock_filter.return_value.__aiter__.return_value = [mock_existing_person]

            await persist_persons(persons_data)

            # Should call update for existing persons
            mock_update.assert_called_once()


class TestEnrichmentIntegration:
    """Integration tests for enrichment workflow."""

    @pytest.mark.asyncio
    async def test_enrich_async_workflow(self, setup_test_db):
        """Test the complete enrichment workflow."""
        settings = Settings()

        with patch('easy_access.enrichment.osiris.ensure_db_inited') as mock_db, \
             patch('easy_access.enrichment.osiris.close_connections') as mock_close, \
             patch('easy_access.enrichment.osiris.gather_target_course_codes') as mock_gather_codes, \
             patch('easy_access.enrichment.osiris.select_missing_or_stale_courses') as mock_select_courses, \
             patch('easy_access.enrichment.osiris.fetch_and_parse_courses') as mock_fetch_courses, \
             patch('easy_access.enrichment.osiris.persist_courses') as mock_persist_courses, \
             patch('easy_access.enrichment.osiris.gather_target_person_names') as mock_gather_persons, \
             patch('easy_access.enrichment.osiris.select_missing_or_stale_persons') as mock_select_persons, \
             patch('easy_access.enrichment.osiris.fetch_and_parse_persons') as mock_fetch_persons, \
             patch('easy_access.enrichment.osiris.persist_persons') as mock_persist_persons:

            # Setup mocks
            mock_gather_codes.return_value = {12345}
            mock_select_courses.return_value = {12345}
            mock_fetch_courses.return_value = {12345: {'name': 'Test Course', 'teachers': []}}
            mock_gather_persons.return_value = {"John Doe"}
            mock_select_persons.return_value = {"John Doe"}
            mock_fetch_persons.return_value = {"John Doe": {'name': 'John Doe'}}

            await enrich_async(settings)

            # Verify all functions were called
            mock_gather_codes.assert_called_once_with(settings)
            mock_select_courses.assert_called_once()
            mock_fetch_courses.assert_called_once()
            mock_persist_courses.assert_called_once()
            mock_gather_persons.assert_called_once_with(settings)
            mock_select_persons.assert_called_once()
            mock_fetch_persons.assert_called_once()
            mock_persist_persons.assert_called_once()
