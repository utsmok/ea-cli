import asyncio
import os
import sys
from logging.config import fileConfig

from alembic import context
from sqlalchemy import engine_from_config, pool
from sqlalchemy.ext.asyncio import async_engine_from_config

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# Import your project's Base metadata here
try:
    from easy_access.db.models_base import metadata as target_metadata
except Exception:  # pragma: no cover - keep imports tolerant during scaffolding
    target_metadata = None
else:
    # Import SA models to ensure tables are registered on Base.metadata
    import contextlib

    with contextlib.suppress(Exception):
        import easy_access.db.sa_models  # noqa: F401 - registers models on import

config = context.config

# Interpret the config file for Python logging.
if config.config_file_name is not None:
    import contextlib

    with contextlib.suppress(Exception):
        fileConfig(config.config_file_name)


def run_migrations_online() -> None:
    cfg_section = config.get_section(config.config_ini_section) or {}
    # allow a sqlite fallback for local autogenerate without Postgres/docker
    url = cfg_section.get("sqlalchemy.url") or ""

    if url.startswith("sqlite"):
        # use sync engine for sqlite autogenerate
        sync_engine = engine_from_config(
            cfg_section, prefix="sqlalchemy.", poolclass=pool.NullPool
        )
        with sync_engine.connect() as connection:
            context.configure(
                connection=connection,
                target_metadata=target_metadata,
                compare_type=True,
            )
            with context.begin_transaction():
                context.run_migrations()
        return

    connectable = async_engine_from_config(
        cfg_section, prefix="sqlalchemy.", poolclass=pool.NullPool
    )

    async def do_run() -> None:
        async with connectable.connect() as connection:
            await connection.run_sync(run_migrations)
        await connectable.dispose()

    asyncio.run(do_run())


def run_migrations(connection) -> None:
    context.configure(
        connection=connection, target_metadata=target_metadata, compare_type=True
    )

    with context.begin_transaction():
        context.run_migrations()


if context.is_offline_mode():
    raise RuntimeError(
        "Offline mode not supported in this template; run migrations online"
    )
else:
    run_migrations_online()
