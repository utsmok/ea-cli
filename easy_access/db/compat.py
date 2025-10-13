"""Small compatibility layer exposing a tiny subset of the old Tortoise API.

This file provides helper adapters so we can incrementally replace Tortoise
call sites with SQLAlchemy implementations. Keep this file intentionally
small — add functions as needed when migrating modules.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from .session import get_session


async def bulk_create(
    table, rows: Iterable[dict[str, Any]], batch_size: int = 500
) -> None:
    """Insert rows in batches using SQLAlchemy Core/ORM table insert.

    `table` may be a SQLAlchemy ORM class or a Table object; keep usage
    simple in early migration phases.
    """
    async for session in get_session():
        async with session.begin():
            # Using SQLAlchemy Core insert for performance and simplicity
            await session.execute(table.__table__.insert(), list(rows))


async def get_by_id(model, id_value) -> Any | None:
    async for session in get_session():
        return await session.get(model, id_value)


async def filter_many(query_callable) -> list[Any]:
    """Run a query callable that accepts an AsyncSession and returns a statement.

    Example:
        async def q(sess):
            return select(MyModel).where(MyModel.foo == 'bar')

        results = await filter_many(q)
    """
    async for session in get_session():
        stmt = await query_callable(session)
    result = await session.scalars(stmt)
    return list(result.all())

    # If the session factory yielded nothing (should not happen), return an
    # empty list to satisfy callers.
    return []
