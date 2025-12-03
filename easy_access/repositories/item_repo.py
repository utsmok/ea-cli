"""
Repository for CopyrightItem database operations.

This module centralizes all Tortoise ORM calls for CopyrightItem,
including CRUD operations, bulk operations, and data retrieval.
"""

from itertools import batched
from typing import Any

import polars as pl
from loguru import logger
from sqlalchemy import Engine

from easy_access.db.base import (
    copyright_item_from_dict,
    ensure_db_inited,
    init_engine,
)
from easy_access.db.models import CopyrightItem, ItemUpdate
from easy_access.settings import Settings
from easy_access.utils import safe_int, standardize_dataframe


class CopyrightItemRepository:
    """Repository for CopyrightItem database operations."""

    def __init__(self, settings: Settings):
        """
        Create a repository bound to the given application settings for lazy engine initialization.
        
        Parameters:
            settings (Settings): Application configuration used to initialize the database engine and control repository behavior.
        """
        self.settings = settings
        self._engine: Engine | None = None

    @property
    def engine(self) -> Engine:
        """Lazily initialize and return the SQLAlchemy engine."""
        if self._engine is None:
            self._engine = init_engine(settings=self.settings)
        return self._engine

    async def ensure_initialized(self) -> None:
        """
        Ensure the underlying database is initialized so the repository can operate.
        
        Prepares database schemas and connections according to the repository settings.
        """
        await ensure_db_inited(self.settings)

    # ----- Single Item Operations -----

    async def get_by_id(self, material_id: int) -> CopyrightItem | None:
        """
        Get a CopyrightItem by its material_id.

        Args:
            material_id: The primary key of the item

        Returns:
            CopyrightItem instance or None if not found
        """
        return await CopyrightItem.get_or_none(material_id=material_id)

    async def get(self, material_id: int) -> CopyrightItem:
        """
        Retrieve a CopyrightItem by its material_id.
        
        Returns:
            CopyrightItem: The matching CopyrightItem.
        
        Raises:
            DoesNotExist: If no item with the given material_id exists.
        """
        return await CopyrightItem.get(material_id=material_id)

    async def save(
        self, item: CopyrightItem, update_fields: list[str] | None = None
    ) -> None:
        """
        Persist a CopyrightItem, optionally restricting which fields are updated.
        
        Parameters:
            item (CopyrightItem): The item to persist.
            update_fields (list[str] | None): Specific field names to update on the existing record; if None, all fields are saved.
        """
        if update_fields:
            await item.save(update_fields=update_fields)
        else:
            await item.save()

    # ----- Query Operations -----

    async def get_all(self, filter_dict: dict | None = None) -> list[CopyrightItem]:
        """
        Retrieve all CopyrightItem records, optionally applying simple equality filters.
        
        Parameters:
            filter_dict (dict | None): Optional mapping of field names to values used as equality filters (passed as kwargs to the query).
        
        Returns:
            list[CopyrightItem]: List of matching CopyrightItem instances.
        """
        if filter_dict:
            return await CopyrightItem.filter(**filter_dict).all()
        return await CopyrightItem.all()

    async def get_existing_material_ids(self) -> set[int]:
        """
        Return the set of material IDs present in the database, converted to integers where possible.
        
        Returns:
            set[int]: Material IDs as integers; any values that cannot be converted or are None are excluded.
        """
        existing = await CopyrightItem.all().values("material_id")
        existing_ids = {safe_int(m["material_id"]) for m in existing}
        return {m for m in existing_ids if m is not None}

    async def exists(self, material_id: int) -> bool:
        """
        Check if a CopyrightItem exists.

        Args:
            material_id: The material_id to check

        Returns:
            True if exists, False otherwise
        """
        return await CopyrightItem.exists(material_id=material_id)

    # ----- Bulk Operations -----

    async def bulk_create(self, items: list[CopyrightItem]) -> None:
        """
        Create multiple CopyrightItem records; if the bulk insert fails, attempt to save each item individually.
        
        Parameters:
            items (list[CopyrightItem]): CopyrightItem instances to create.
        
        Raises:
            Exception: If creation fails for some items after the per-item fallback; exception message includes the list of failed `material_id` values.
        """
        if not items:
            return

        try:
            await CopyrightItem.bulk_create(objects=items)
            logger.success(f"Created {len(items)} new copyright items in db.")
        except Exception as e:
            logger.warning(
                f"Bulk creation failed: {e}. Attempting one-by-one creation."
            )
            failed_items = []
            for item in items:
                try:
                    await item.save()
                except Exception as save_error:
                    logger.error(
                        f"Failed to save item {item.material_id}: {save_error}"
                    )
                    failed_items.append(item.material_id)

            if failed_items:
                raise Exception(
                    f"Failed to create items with material_ids: {failed_items}"
                ) from e

    async def bulk_update(
        self, items: list[CopyrightItem], fields: list[str], batch_size: int = 500
    ) -> None:
        """
        Update multiple CopyrightItem records in batches.
        
        Parameters:
            items (list[CopyrightItem]): CopyrightItem instances to update.
            fields (list[str]): Attribute names to update on each item.
            batch_size (int): Maximum number of items to process per batch.
        """
        if not items:
            return

        for batch in batched(items, batch_size):
            batch_list = list(batch)
            logger.info(f"Updating fields {fields} for {len(batch_list)} items.")
            await CopyrightItem.bulk_update(batch_list, fields=fields)

    # ----- Change Tracking -----

    async def create_changelog_entries(
        self,
        updates: dict[int, dict],
        user_email: str | None = None,
        batch_size: int = 500,
    ) -> None:
        """
        Create changelog entries and associate them with the corresponding CopyrightItem records.
        
        Parameters:
            updates (dict[int, dict]): Mapping from material_id to a dict of change details; each dict becomes the `change_details` for a new ItemUpdate.
            user_email (str | None): If provided, added to each change dict under the key `"modified_by"` to record who made the change.
            batch_size (int): Number of updates to process per batch when creating ItemUpdate records.
        """
        if not updates:
            return

        # Add user email to all changes
        if user_email:
            for changes in updates.values():
                changes["modified_by"] = user_email
            logger.info(f"items modified by {user_email}")

        for update_batch in batched(list(updates.items()), batch_size):
            update_batch_list = list(update_batch)
            logger.info(f"Creating {len(update_batch_list)} changelog entries in db.")

            mat_ids = [mat_id for mat_id, _ in update_batch_list]
            await ItemUpdate.bulk_create(
                [
                    ItemUpdate(change_details=changes, material_id=mat_id)
                    for mat_id, changes in update_batch_list
                ]
            )

            # Add each ItemUpdate as a m2m relation to the corresponding CopyrightItem
            for mat_id in mat_ids:
                item = await CopyrightItem.get(material_id=mat_id)
                update = (
                    await ItemUpdate.filter(material_id=mat_id)
                    .order_by("-created_at")
                    .first()
                )
                await item.changes.add(update)

    # ----- DataFrame Operations -----

    def get_dataframe(self, additional_cols: list[str] | None = None) -> pl.DataFrame:
        """
        Return a Polars DataFrame containing copyright items.
        
        Parameters:
            additional_cols (list[str] | None): Optional extra columns to include. Column names not in the repository's allowed whitelist are ignored and a warning is logged. The alias "faculty_id AS faculty" is permitted.
        
        Returns:
            pl.DataFrame: DataFrame with the requested columns (or only `material_id` if no requested columns are valid).
        
        Raises:
            Exception: If retrieving the data fails.
        """
        from time import time

        full_start = time()
        col_order: set[str] = set(self.settings.data_settings.raw_data_col_order)

        # Remove cols that aren't directly selectable
        for col in ["google_search_file", "type", "replacement_id", "is_duplicate"]:
            col_order.discard(col)

        if "faculty" in col_order:
            col_order.remove("faculty")
            select_cols = {*col_order, "faculty_id AS faculty"}
        else:
            select_cols = col_order

        if additional_cols:
            select_cols.update(additional_cols)

        # Whitelist of allowed column names to prevent SQL injection
        # These are the known columns in the copyright_data table
        allowed_columns = {
            "material_id", "period", "department", "course_code", "course_name",
            "url", "filename", "title", "owner", "filetype", "classification",
            "ml_prediction", "manual_classification", "manual_identifier",
            "v2_manual_classification", "v2_overnamestatus", "v2_lengte",
            "scope", "remarks", "auditor", "last_change", "status",
            "isbn", "doi", "in_collection", "pagecount", "wordcount", "picturecount",
            "author", "publisher", "reliability", "pages_x_students",
            "count_students_registered", "filehash", "last_scan_date_university",
            "last_scan_date_course", "retrieved_from_copyright_on", "workflow_status",
            "possible_fine", "infringement", "file_exists", "last_canvas_check",
            "canvas_course_id", "faculty_id", "is_duplicate", "created_at", "modified_at",
            "faculty_id AS faculty",  # Allow this specific alias
        }
        
        # Validate and filter columns
        validated_cols = {col for col in select_cols if col in allowed_columns}
        if len(validated_cols) != len(select_cols):
            invalid_cols = select_cols - allowed_columns
            logger.warning(f"Removed invalid column names: {invalid_cols}")

        if not validated_cols:
            validated_cols = {"material_id"}  # Default to just material_id

        try:
            query = "SELECT " + ", ".join(validated_cols) + " FROM copyright_data cd"
            query_start = time()
            df = pl.read_database(
                query=query, connection=self.engine.connect(), infer_schema_length=None
            )
            end = time()
            logger.info(f"CopyrightItemRepository.get_dataframe returned {len(df)} rows")
            logger.info(f"query took {end - query_start} seconds")
            logger.info(f"full function took {end - full_start} seconds")
        except Exception as e:
            logger.error(f"Error retrieving copyright items: {e}")
            raise e

        return df

    # ----- Factory Methods -----

    async def create_from_dict(self, item_dict: dict) -> CopyrightItem | None:
        """
        Create a CopyrightItem instance from a dictionary.

        Args:
            item_dict: Dictionary with item data

        Returns:
            CopyrightItem instance or None if creation fails
        """
        return await copyright_item_from_dict(item_dict)

    async def create_from_dicts(self, items: list[dict]) -> list[CopyrightItem]:
        """
        Create CopyrightItem objects from a list of dictionaries.
        
        Parameters:
        	items (list[dict]): Dictionaries representing item data to convert into CopyrightItem instances.
        
        Returns:
        	created_items (list[CopyrightItem]): List of created CopyrightItem instances; dictionaries that could not be converted are omitted.
        """
        new_objects = []
        for item_dict in items:
            obj = await copyright_item_from_dict(item_dict)
            if obj:
                new_objects.append(obj)
        return new_objects

    # ----- Data Preprocessing -----

    async def preprocess_input_data(
        self, data: pl.DataFrame | list[dict]
    ) -> tuple[list[dict], list[dict]]:
        """
        Split input into new items to create and existing items to update.
        
        When given a Polars DataFrame, the function standardizes the frame, determines which rows
        correspond to material_ids already present in the database, and separates rows into:
        - new_items: rows with material_id not in the database and that contain the required
          creation fields (`period`, `department`, `course_code`, `course_name`).
        - update_items: rows whose material_id exists in the database.
        
        When given a list of dictionaries, the list is treated as update_items and returned unchanged.
        
        Parameters:
            data (pl.DataFrame | list[dict]): Input data as a Polars DataFrame or a list of record dicts.
        
        Returns:
            tuple[list[dict], list[dict]]: A tuple (new_items, update_items) where each element is a list of
            dictionaries representing records to create and records to update, respectively.
        """
        new_items = []
        update_items = []

        if isinstance(data, pl.DataFrame):
            data = standardize_dataframe(data)
            existing_mat_ids = await self.get_existing_material_ids()

            # Candidate new items (may be partial if coming from faculty sheets)
            candidate_new_items = (
                data.with_columns(pl.col("material_id").cast(int))
                .filter(~pl.col("material_id").is_in(existing_mat_ids))
                .to_dicts()
            )

            # Only create new CopyrightItem objects when the incoming row contains
            # the required, non-nullable fields present in the model.
            required_for_creation = [
                "period",
                "department",
                "course_code",
                "course_name",
            ]

            skipped_mat_ids: list[Any] = []

            for itm in candidate_new_items:
                ok = True
                for rc in required_for_creation:
                    v = itm.get(rc)
                    if v is None or (isinstance(v, str) and v.strip() == ""):
                        ok = False
                        break
                if ok:
                    new_items.append(itm)
                else:
                    mid_val = itm.get("material_id")
                    if mid_val is None:
                        skipped_mat_ids.append("<no-id>")
                    else:
                        safe_mid = safe_int(mid_val)
                        if safe_mid is not None:
                            skipped_mat_ids.append(safe_mid)
                        else:
                            skipped_mat_ids.append(str(mid_val))

            if skipped_mat_ids:
                logger.warning(
                    f"Skipping {len(skipped_mat_ids)} new items missing required fields (not creating in DB): {skipped_mat_ids[:20]}"
                )

            update_items = (
                data.with_columns(pl.col("material_id").cast(int))
                .filter(pl.col("material_id").is_in(existing_mat_ids))
                .to_dicts()
            )

        elif isinstance(data, list):
            update_items = data

        logger.info(f"# of new items: {len(new_items)}")
        logger.info(f"# of items to update: {len(update_items)}")

        return new_items, update_items