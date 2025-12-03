"""
Dagster resources for the Easy Access pipeline.

This module provides:
- TortoiseDBResource: Manages Tortoise ORM connection lifecycle for Dagster assets
- SettingsResource: Provides access to application settings
"""

from contextlib import asynccontextmanager
from typing import Any

from dagster import ConfigurableResource, InitResourceContext
from loguru import logger


class TortoiseDBResource(ConfigurableResource):
    """
    Manages the Tortoise ORM connection lifecycle for Dagster assets.

    This resource ensures database connections are properly initialized before
    asset execution and closed afterwards to release SQLite locks for dlt
    or other processes.

    Usage in assets:
        @asset
        async def my_asset(db: TortoiseDBResource):
            async with db.yield_for_execution(context):
                # Your async database operations here
                pass
    """

    @asynccontextmanager
    async def yield_for_execution(self, context: InitResourceContext | Any = None):
        """
        Context manager that initializes Tortoise ORM and ensures cleanup.

        Args:
            context: Dagster resource context (optional)

        Yields:
            self: The resource instance for use in the asset
        """
        from easy_access.db.base import close_connections, init
        from easy_access.settings import SETTINGS

        try:
            # Initialize Tortoise ORM using existing settings
            await init(SETTINGS)
            logger.debug("Tortoise ORM initialized for Dagster asset execution")
            yield self
        finally:
            # Strictly close connections to release SQLite locks
            await close_connections()
            logger.debug("Tortoise ORM connections closed after asset execution")


class SettingsResource(ConfigurableResource):
    """
    Provides access to application settings for Dagster assets.

    This resource wraps the Settings object to make it available
    in asset functions.
    """

    def get_settings(self):
        """
        Returns the application Settings instance.

        Returns:
            Settings: The global SETTINGS object
        """
        from easy_access.settings import SETTINGS

        return SETTINGS
