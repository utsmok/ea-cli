import sys
import os
import asyncio
import pytest
import pytest_asyncio
from tortoise import Tortoise

# Ensure repository root is on PYTHONPATH for tests
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)


@pytest_asyncio.fixture(scope="function")
async def setup_test_db():
    """Setup and teardown test database for each test function."""
    # Initialize Tortoise for tests with in-memory SQLite
    await Tortoise.init(
        db_url="sqlite://:memory:",
        modules={
            "models": ["easy_access.db.models"]
        }
    )

    # Generate the schema
    await Tortoise.generate_schemas(safe=True)

    yield

    # Clean up connections and ensure the event loop is not left with
    # pending tasks or async generators which can cause pytest to hang.
    try:
        await Tortoise.close_connections()
    finally:
        # Cancel any still-running tasks (except the current one)
        try:
            loop = asyncio.get_running_loop()
            pending = [t for t in asyncio.all_tasks(loop) if t is not asyncio.current_task()]
            if pending:
                for t in pending:
                    t.cancel()
                await asyncio.gather(*pending, return_exceptions=True)

            # Shutdown async generators (Python 3.7+)
            if hasattr(loop, 'shutdown_asyncgens'):
                await loop.shutdown_asyncgens()
        except RuntimeError:
            # Event loop already closed or not running; ignore
            pass
