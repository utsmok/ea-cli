import pytest
from unittest.mock import MagicMock, patch
from unittest.mock import AsyncMock
from datetime import datetime, timedelta

from easy_access.enrichment import osiris


class QuerySetMock:
    def __init__(self, results=None):
        self._results = results or []

    def __await__(self):
        async def _coro():
            return self._results

        return _coro().__await__()

    def __iter__(self):
        return iter(self._results)


@pytest.mark.asyncio
async def test_select_missing_or_stale_courses_ttl_none():
    course_codes = {1, 2, 3}

    # existing: only course 1
    mock_course = MagicMock()
    mock_course.cursuscode = 1

    # tracked missing contains 3
    tracked_missing = [MagicMock(cursuscode=3)]

    with patch('easy_access.enrichment.osiris.Course.filter', return_value=QuerySetMock([mock_course])), \
         patch('easy_access.enrichment.osiris.MissingCourse.filter', return_value=QuerySetMock(tracked_missing)):
        to_fetch = await osiris.select_missing_or_stale_courses(None, course_codes, ttl_days=None)
        # missing codes are 2 and 3
        assert to_fetch == {2, 3}


@pytest.mark.asyncio
async def test_select_missing_or_stale_courses_with_ttl():
    course_codes = {1, 2, 3, 4}
    ttl = 7
    now = datetime.now().astimezone()

    # course 1 is stale (modified_at is None => considered stale), course 2 is fresh
    stale_course = MagicMock()
    stale_course.cursuscode = 1
    stale_course.modified_at = None

    fresh_course = MagicMock()
    fresh_course.cursuscode = 2
    # provide an object with astimezone() to match production code expectations
    class _FakeDT:
        def __init__(self, dt):
            self._dt = dt

        def astimezone(self, tz):
            # return a datetime aligned with the requested tz
            return datetime.now(tz) - timedelta(days=1)

    fresh_course.modified_at = _FakeDT(datetime.now())

    # tracked missing contains 3 and is stale (modified_at None)
    tracked_missing = [MagicMock(cursuscode=3, modified_at=None)]

    with patch('easy_access.enrichment.osiris.Course.filter', return_value=QuerySetMock([stale_course, fresh_course])), \
         patch('easy_access.enrichment.osiris.MissingCourse.filter', return_value=QuerySetMock(tracked_missing)):
        to_fetch = await osiris.select_missing_or_stale_courses(None, course_codes, ttl_days=ttl)
        # expect stale existing (1), retry tracked missing (3), and new missing (4)
        assert to_fetch == {1, 3, 4}


@pytest.mark.asyncio
async def test_select_missing_or_stale_persons_ttl_none():
    names = {'Alice', 'Bob', 'Carol'}

    # existing only contains Alice
    existing_person = MagicMock()
    existing_person.input_name = 'Alice'

    with patch('easy_access.enrichment.osiris.Person.filter', return_value=QuerySetMock([existing_person])):
        to_fetch = await osiris.select_missing_or_stale_persons(None, names, ttl_days=None)
        assert to_fetch == {'Bob', 'Carol'}


@pytest.mark.asyncio
async def test_select_missing_or_stale_persons_with_ttl():
    names = {'Alice', 'Eve', 'Frank'}
    ttl = 3
    now = datetime.now().astimezone()

    # Alice exists but is stale
    alice = MagicMock()
    alice.input_name = 'Alice'
    alice.main_name = 'Alice'
    alice.modified_at = now - timedelta(days=ttl + 1)

    # Eve exists but unresolved (main_name is None) and stale
    eve = MagicMock()
    eve.input_name = 'Eve'
    eve.main_name = None
    eve.modified_at = now - timedelta(days=ttl + 2)

    with patch('easy_access.enrichment.osiris.Person.filter', return_value=QuerySetMock([alice, eve])):
        to_fetch = await osiris.select_missing_or_stale_persons(None, names, ttl_days=ttl)
        # expect Alice (stale), Eve (unresolved+stale) and Frank (missing)
        assert to_fetch == {'Alice', 'Eve', 'Frank'}


@pytest.mark.asyncio
async def test_fetch_course_data_error_returns_empty():
    mock_client = MagicMock()
    mock_client.post.side_effect = Exception('network')

    result = await osiris.fetch_course_data(99999, mock_client)
    assert result == {}


@pytest.mark.asyncio
async def test_fetch_person_data_error_returns_empty():
    mock_client = MagicMock()
    mock_client.get.side_effect = Exception('network')

    result = await osiris.fetch_person_data('Nonexistent Person', mock_client)
    assert result == {}


@pytest.mark.asyncio
async def test_fetch_and_parse_courses_records_missing_and_returns_found():
    # Prepare a fake fetch function that returns data for code 2 and empty for 1
    async def fake_fetch(code, client):
        if code == 1:
            return {}
        return {'cursuscode': code, 'internal_id': 10 + code, 'year': 2024}

    # Patch fetch_course_data and MissingCourse helpers
    with patch('easy_access.enrichment.osiris.fetch_course_data', new=AsyncMock(side_effect=fake_fetch)), \
         patch('easy_access.enrichment.osiris.MissingCourse.get_or_none', new=AsyncMock(return_value=None)) as mock_get_or_none, \
         patch('easy_access.enrichment.osiris.MissingCourse.create', new=AsyncMock()) as mock_create, \
         patch('easy_access.enrichment.osiris.MissingCourse.filter', return_value=AsyncMock()) as mock_filter:
        # pass a simple settings object (module doesn't use it here)
        results = await osiris.fetch_and_parse_courses(object(), {1, 2}, max_concurrent=2)

        # Expect only code 2 present in results
        assert 2 in results and 1 not in results
        # MissingCourse.create should have been called for code 1
        assert mock_create.await_count >= 1 or mock_create.called
