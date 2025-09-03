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

    # Clean up connections after each test
    await Tortoise.close_connections()
