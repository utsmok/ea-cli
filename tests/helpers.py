"""Small test helpers shared across tests.

Provides a QuerySet-like mock useful for awaited ``Model.filter(...)``
or ``Model.all()`` calls in tests. The object is awaitable and exposes an
async ``update`` so production code that does `await qs.update(...)` or
`await Model.filter(...).prefetch_related(... )` will work with this mock.
"""

from unittest.mock import AsyncMock


class QuerySetMock:
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
