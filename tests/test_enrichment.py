"""Minimal, well-formed tests for the enrichment module.

These tests avoid complex QuerySet chaining by using a small
QuerySet-like mock object that is awaitable and exposes an
async `update` method. HTTP responses are mocked so `.json()`
is synchronous (MagicMock) matching how the production code uses it.
"""

import pytest
from unittest.mock import AsyncMock, MagicMock, patch

from easy_access.enrichment import osiris
from easy_access.enrichment.osiris import fetch_course_data, fetch_person_data, persist_courses, persist_persons


class QuerySetMock:
    """A tiny QuerySet-like mock that's awaitable and chainable.

    - awaiting the object returns the provided result list
    - `.update` is an AsyncMock so `await qs.update(...)` works
    - `prefetch_related` and `distinct` return self for chaining
    """

    def __init__(self, results=None):
        self._results = results or []
        self.update = AsyncMock()

    def __await__(self):
        async def _coro():
            return self._results

        return _coro().__await__()

    def prefetch_related(self, *args, **kwargs):
        return self

    def distinct(self, *args, **kwargs):
        return self

    def __iter__(self):
        return iter(self._results)


@pytest.mark.asyncio
async def test_persist_courses_new(setup_test_db):
    courses_data = {12345: {'cursuscode': 12345, 'name': 'Test Course'}}

    with patch('easy_access.enrichment.osiris.Course.filter', return_value=QuerySetMock([])), \
         patch('easy_access.enrichment.osiris.Course.create', new=AsyncMock()) as mock_create:
        await persist_courses(courses_data)
        assert mock_create.await_count == 1 or mock_create.called


@pytest.mark.asyncio
async def test_persist_courses_update(setup_test_db):
    courses_data = {12345: {'cursuscode': 12345, 'name': 'Updated'}}

    mock_existing = MagicMock()
    mock_existing.cursuscode = 12345

    qs = QuerySetMock([mock_existing])

    with patch('easy_access.enrichment.osiris.Course.filter', return_value=qs):
        await persist_courses(courses_data)
        # ensure update was attempted on the queryset
        qs.update.assert_awaited()


@pytest.mark.asyncio
async def test_persist_persons_new(setup_test_db):
    persons_data = {'John Doe': {'input_name': 'John Doe', 'main_name': 'John Doe'}}

    with patch('easy_access.enrichment.osiris.Person.filter', return_value=QuerySetMock([])), \
         patch('easy_access.enrichment.osiris.Person.create', new=AsyncMock()) as mock_create:
        await persist_persons(persons_data)
        assert mock_create.await_count == 1 or mock_create.called


@pytest.mark.asyncio
async def test_persist_persons_update(setup_test_db):
    persons_data = {'John Doe': {'input_name': 'John Doe', 'main_name': 'Dr. John Doe'}}

    mock_existing = MagicMock()
    mock_existing.input_name = 'John Doe'

    qs = QuerySetMock([mock_existing])

    with patch('easy_access.enrichment.osiris.Person.filter', return_value=qs):
        await persist_persons(persons_data)
        qs.update.assert_awaited()


@pytest.mark.asyncio
async def test_fetch_course_data_success(setup_test_db):
    course_code = 12345
    mock_client = AsyncMock()

    # Build a fake response whose .json() is synchronous
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json = MagicMock(return_value={
        'hits': {
            'hits': [
                {'_source': {'cursus_lange_naam': 'Test Course', 'cursus_korte_naam': 'TEST', 'faculteit': 'EEMCS', 'docenten': ['Dr. John Doe']}}
            ]
        }
    })

    mock_details = MagicMock()
    mock_details.status_code = 200
    mock_details.json = MagicMock(return_value={'items': []})

    mock_client.post.return_value = mock_response
    mock_client.get.return_value = mock_details

    result = await fetch_course_data(course_code, mock_client)
    assert result.get('cursuscode') == 12345
    assert result.get('name') == 'Test Course'
    assert result.get('short_name') == 'TEST'


@pytest.mark.asyncio
async def test_fetch_person_data_success(setup_test_db):
    person_name = 'John Doe'
    mock_client = AsyncMock()

    search_resp = MagicMock()
    search_resp.status_code = 200
    search_resp.text = '<a data-link="john.doe">John Doe</a>'

    detail_resp = MagicMock()
    detail_resp.status_code = 200
    detail_resp.text = '<h1 class="pageheader__title">John Doe</h1><a href="mailto:john.doe@utwente.nl">john.doe@utwente.nl</a>'

    mock_client.get.side_effect = [search_resp, detail_resp]

    with patch('easy_access.enrichment.osiris.Levenshtein.ratio', return_value=0.95):
        result = await fetch_person_data(person_name, mock_client)

    assert result.get('input_name') == 'John Doe'
    assert result.get('email') == 'john.doe@utwente.nl'

