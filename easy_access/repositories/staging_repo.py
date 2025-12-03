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
        """
        Create a StagingRepository configured with application settings.
        
        Parameters:
            settings (Settings): Application settings that configure database access and repository behavior; stored for use by repository methods.
        """
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
        Return a page of staged copyright items.
        
        Parameters:
            offset (int): Number of records to skip before the returned batch.
            limit (int): Maximum number of records to return.
        
        Returns:
            list[StagedCopyrightItem]: The staged copyright items for the requested page.
        """
        return await StagedCopyrightItem.all().offset(offset).limit(limit)

    async def clear_staged_items(self, material_ids: list[int] | None = None) -> int:
        """
        Clear staged copyright items.
        
        Parameters:
            material_ids (list[int] | None): Optional list of material IDs to delete; if None, deletes all staged items.
        
        Returns:
            int: Number of deleted staged items.
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
        Insert or update copyright records from the given DataFrame into the staging table.
        
        Parameters:
            data (pl.DataFrame): Polars DataFrame containing copyright records with columns matching the staging model fields. Existing records with the same `material_id` will be updated (all staging fields except the primary key).
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
        Retrieve all staged faculty update records.
        
        Returns:
            list[StagedFacultyUpdate]: All StagedFacultyUpdate instances from the staging table.
        """
        return await StagedFacultyUpdate.all()

    async def clear_staged_faculty_updates(
        self, material_ids: list[int] | None = None
    ) -> int:
        """
        Clear staged faculty updates from the staging table.
        
        Parameters:
            material_ids (list[int] | None): Optional list of material IDs to delete; if None, deletes all staged faculty updates.
        
        Returns:
            int: Number of deleted staged faculty updates.
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
        Ingest faculty update records into the staging table, upserting by material_id.
        
        Only the columns `material_id`, `manual_classification`, `remarks`, and `workflow_status` are used from the provided Polars DataFrame; rows are inserted or updated on conflict by `material_id`, with `manual_classification`, `remarks`, and `workflow_status` overwritten on conflict.
        
        Parameters:
            data (pl.DataFrame): Polars DataFrame containing faculty update records. Only the columns listed above are considered.
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
        Record a processing failure for later inspection or retry.
        
        Stores a StagedProcessingFailure using the given payload, converting `material_id` to an integer when possible and truncating `error_message` to 1900 characters to fit the database field.
        
        Parameters:
            material_id (int | None): Identifier of the failed material; will be converted safely to an int or stored as None.
            staged_payload (dict): The original staged record payload being processed.
            error_message (str): Error text describing the failure; only the first 1900 characters are persisted.
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
        Retrieve processing failures, optionally filtered by material IDs.
        
        Parameters:
            material_ids (list[int] | None): Optional list of material IDs to restrict the returned failures to those materials. If None, all processing failures are returned.
        
        Returns:
            list[StagedProcessingFailure]: List of matching StagedProcessingFailure instances.
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
        Provide the canonical list of field names for StagedCopyrightItem.
        
        Returns:
            fields (list[str]): Field names in the canonical order used for staging records.
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
        Convert a StagedCopyrightItem into a dictionary mapping its staged-field names to their values.
        
        Parameters:
            staged_item (StagedCopyrightItem): The staged item to convert.
        
        Returns:
            dict: A mapping from each staged item field name to its value (`None` if the attribute is missing).
        """
        item_dict = {}
        for field in self.get_staged_item_fields():
            item_dict[field] = getattr(staged_item, field, None)
        return item_dict