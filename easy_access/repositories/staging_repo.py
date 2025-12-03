"""
Repository for staging table database operations.

This module centralizes all Tortoise ORM calls for staging tables,
including StagedCopyrightItem and StagedFacultyUpdate.
"""

import polars as pl
from loguru import logger

from easy_access.db.base import ensure_db_inited
from easy_access.db.models import (
    StagedCopyrightItem,
    StagedFacultyUpdate,
    StagedProcessingFailure,
)
from easy_access.settings import Settings
from easy_access.utils import safe_int, standardize_dataframe


class StagingRepository:
    """Repository for staging table operations."""

    def __init__(self, settings: Settings):
        """Initialize the repository with settings."""
        self.settings = settings

    async def ensure_initialized(self) -> None:
        """Ensure the database is initialized."""
        await ensure_db_inited(self.settings)

    # ----- StagedCopyrightItem Operations -----

    async def get_all_staged_items(self) -> list[StagedCopyrightItem]:
        """
        Get all staged copyright items.

        Returns:
            List of StagedCopyrightItem instances
        """
        return await StagedCopyrightItem.all()

    async def get_staged_items_batch(
        self, offset: int = 0, limit: int = 500
    ) -> list[StagedCopyrightItem]:
        """
        Get a batch of staged copyright items.

        Args:
            offset: Starting offset
            limit: Maximum number of items to return

        Returns:
            List of StagedCopyrightItem instances
        """
        return await StagedCopyrightItem.all().offset(offset).limit(limit)

    async def clear_staged_items(self, material_ids: list[int] | None = None) -> int:
        """
        Clear staged copyright items.

        Args:
            material_ids: Optional list of material IDs to clear.
                         If None, clears all staged items.

        Returns:
            Number of deleted items
        """
        if material_ids:
            deleted = await StagedCopyrightItem.filter(
                material_id__in=material_ids
            ).delete()
        else:
            deleted = await StagedCopyrightItem.all().delete()
        logger.info(f"Cleared {deleted} staged copyright items.")
        return deleted

    async def ingest_raw_data(self, data: pl.DataFrame) -> None:
        """
        Load raw copyright data into the staging table.

        Args:
            data: Polars DataFrame with raw copyright data
        """
        await self.ensure_initialized()
        items = standardize_dataframe(data).to_dicts()
        staged_items = [StagedCopyrightItem(**item) for item in items]

        # Define fields to update on conflict (all fields except primary key)
        update_fields = [
            "period",
            "department",
            "course_code",
            "course_name",
            "url",
            "filename",
            "title",
            "owner",
            "filetype",
            "classification",
            "manual_classification",
            "manual_identifier",
            "scope",
            "remarks",
            "ml_prediction",
            "isbn",
            "doi",
            "in_collection",
            "pagecount",
            "wordcount",
            "picturecount",
            "author",
            "publisher",
            "auditor",
            "last_change",
            "status",
            "reliability",
            "pages_x_students",
            "count_students_registered",
            "retrieved_from_copyright_on",
            "workflow_status",
            "faculty",
            "file_exists",
        ]

        await StagedCopyrightItem.bulk_create(
            staged_items,
            on_conflict=["material_id"],
            update_fields=update_fields,
        )
        logger.info(f"Ingested {len(staged_items)} items into staging table.")

    # ----- StagedFacultyUpdate Operations -----

    async def get_all_staged_faculty_updates(self) -> list[StagedFacultyUpdate]:
        """
        Get all staged faculty updates.

        Returns:
            List of StagedFacultyUpdate instances
        """
        return await StagedFacultyUpdate.all()

    async def clear_staged_faculty_updates(
        self, material_ids: list[int] | None = None
    ) -> int:
        """
        Clear staged faculty updates.

        Args:
            material_ids: Optional list of material IDs to clear.
                         If None, clears all staged updates.

        Returns:
            Number of deleted items
        """
        if material_ids:
            deleted = await StagedFacultyUpdate.filter(
                material_id__in=material_ids
            ).delete()
        else:
            deleted = await StagedFacultyUpdate.all().delete()
        logger.info(f"Cleared {deleted} staged faculty updates.")
        return deleted

    async def ingest_faculty_updates(self, data: pl.DataFrame) -> None:
        """
        Load faculty updates into the staging table.

        Args:
            data: Polars DataFrame with faculty update data
        """
        await self.ensure_initialized()
        data = data.select(
            ["material_id", "manual_classification", "remarks", "workflow_status"]
        )
        items = standardize_dataframe(data).to_dicts()
        staged_items = [StagedFacultyUpdate(**item) for item in items]

        await StagedFacultyUpdate.bulk_create(
            staged_items,
            on_conflict=["material_id"],
            update_fields=["manual_classification", "remarks", "workflow_status"],
        )
        logger.info(f"Ingested {len(staged_items)} faculty updates into staging table.")

    # ----- Processing Failure Operations -----

    async def record_processing_failure(
        self,
        material_id: int | None,
        staged_payload: dict,
        error_message: str,
    ) -> None:
        """
        Record a processing failure for later inspection/retry.

        Args:
            material_id: Material ID of the failed item
            staged_payload: The original staged data
            error_message: The error message
        """
        await StagedProcessingFailure.create(
            material_id=safe_int(material_id),
            staged_payload=staged_payload,
            error_message=error_message[:1900],  # Truncate to fit field size
        )

    async def get_processing_failures(
        self, material_ids: list[int] | None = None
    ) -> list[StagedProcessingFailure]:
        """
        Get processing failures.

        Args:
            material_ids: Optional list of material IDs to filter by

        Returns:
            List of StagedProcessingFailure instances
        """
        if material_ids:
            return await StagedProcessingFailure.filter(
                material_id__in=material_ids
            ).all()
        return await StagedProcessingFailure.all()

    # ----- Utility Methods -----

    @staticmethod
    def get_staged_item_fields() -> list[str]:
        """
        Get the list of fields on StagedCopyrightItem.

        Returns:
            List of field names
        """
        return [
            "material_id",
            "period",
            "department",
            "course_code",
            "course_name",
            "url",
            "filename",
            "title",
            "owner",
            "filetype",
            "classification",
            "manual_classification",
            "manual_identifier",
            "scope",
            "remarks",
            "ml_prediction",
            "isbn",
            "doi",
            "in_collection",
            "pagecount",
            "wordcount",
            "picturecount",
            "author",
            "publisher",
            "auditor",
            "last_change",
            "status",
            "reliability",
            "pages_x_students",
            "count_students_registered",
            "retrieved_from_copyright_on",
            "workflow_status",
            "faculty",
            "file_exists",
        ]

    def staged_item_to_dict(self, staged_item: StagedCopyrightItem) -> dict:
        """
        Convert a StagedCopyrightItem to a dictionary.

        Args:
            staged_item: The staged item to convert

        Returns:
            Dictionary with item data
        """
        item_dict = {}
        for field in self.get_staged_item_fields():
            item_dict[field] = getattr(staged_item, field, None)
        return item_dict
