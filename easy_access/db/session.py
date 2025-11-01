from __future__ import annotations

from collections.abc import AsyncIterator
from typing import TYPE_CHECKING

from sqlalchemy.ext.asyncio import (
    AsyncEngine,
    AsyncSession,
    async_sessionmaker,
    create_async_engine,
)

if TYPE_CHECKING:
    from easy_access.settings import Settings

_engine: AsyncEngine | None = None
_SessionFactory: async_sessionmaker[AsyncSession] | None = None


def get_database_url(settings: Settings | None = None) -> str:
    """Get the PostgreSQL database URL from Settings, environment, or defaults.

    Args:
        settings: Optional Settings object (currently unused but kept for API consistency)

    Returns:
        Database URL string for PostgreSQL with asyncpg driver
    """
    from os import environ
    from pathlib import Path

    # Check environment variable first
    db_url = environ.get("DATABASE_URL")

    # If not in environment, try to load from .env file
    if not db_url:
        env_file = Path(".env")
        if env_file.exists():
            for line in env_file.read_text().splitlines():
                line = line.strip()
                if line.startswith("DATABASE_URL="):
                    db_url = line.split("=", 1)[1].strip()
                    break

    # Fall back to default matching docker-compose.postgres.yml
    if not db_url:
        db_url = "postgresql+asyncpg://easyaccess:easyaccess@localhost:5432/easyaccess"

    return db_url


def init_db(settings_or_url: Settings | str, *, echo: bool = False) -> None:
    """Initialize the async engine and session factory for the application.

    Call this once at application startup with either a Settings object or DATABASE_URL.

    Args:
        settings_or_url: Either a Settings object or a database URL string
        echo: Whether to echo SQL statements
    """
    global _engine, _SessionFactory
    if _engine is not None:
        return

    # Extract DB URL from Settings if provided
    if isinstance(settings_or_url, str):
        db_url = settings_or_url
    else:
        # It's a Settings object - get URL via helper function
        db_url = get_database_url(settings_or_url)

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
    global _engine, _SessionFactory
    if _engine is not None:
        # Instead of disposing completely, try to reset connection pool
        try:
            # Close all connections in the pool more gently
            await _engine.dispose(close=True)
            # Reset globals to force re-initialization
            _engine = None
            _SessionFactory = None
        except Exception as e:
            # If gentle dispose fails, force reset globals anyway
            _engine = None
            _SessionFactory = None
            # Log the error but don't raise - we want cleanup to succeed
            import logging

            logging.warning(f"Error during database shutdown: {e}")


async def get_session() -> AsyncIterator[AsyncSession]:
    """Async context manager factory for working with sessions.

    Usage:
        async for session in get_session():
            await session.execute(...)
    """
    factory = get_session_factory()
    async with factory() as session:
        yield session
