"""
Repository layer for Easy Access CLI.

This module contains all database access operations, centralized for
easier maintenance and testing. Repositories handle CRUD operations,
bulk operations, and raw SQL queries.
"""

from easy_access.repositories.item_repo import CopyrightItemRepository
from easy_access.repositories.osiris_repo import OsirisRepository
from easy_access.repositories.staging_repo import StagingRepository

__all__ = [
    "CopyrightItemRepository",
    "StagingRepository",
    "OsirisRepository",
]
