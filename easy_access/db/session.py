from __future__ import annotations

from collections.abc import AsyncIterator

from sqlalchemy.ext.asyncio import (
    AsyncEngine,
    AsyncSession,
    async_sessionmaker,
    create_async_engine,
)

_engine: AsyncEngine | None = None
_SessionFactory: async_sessionmaker[AsyncSession] | None = None


def init_db(db_url: str, *, echo: bool = False) -> None:
    """Initialize the async engine and session factory for the application.

    Call this once at application startup with the repository's DATABASE_URL.
    """
    global _engine, _SessionFactory
    if _engine is not None:
        return

    _engine = create_async_engine(db_url, echo=echo)
    _SessionFactory = async_sessionmaker(_engine, expire_on_commit=False)


def get_engine() -> AsyncEngine | None:
    return _engine


def get_session_factory() -> async_sessionmaker[AsyncSession]:
    if _SessionFactory is None:
        raise RuntimeError("Database not initialized. Call init_db() first.")
    return _SessionFactory


async def shutdown_db() -> None:
    """Dispose the async engine and free connections."""
    global _engine
    if _engine is not None:
        await _engine.dispose()
        _engine = None


async def get_session() -> AsyncIterator[AsyncSession]:
    """Async context manager factory for working with sessions.

    Usage:
        async with get_session() as session:
            await session.execute(...)
    """
    factory = get_session_factory()
    async with factory() as session:
        yield session
