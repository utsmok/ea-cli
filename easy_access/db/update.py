"""
functions to update existing data in the database

This module serves as the orchestrator for update operations.
It coordinates between services (business logic) and repositories (data access).

NOTE: Strategy classes and merge logic have been moved to easy_access/services/.
      They are re-exported here for backward compatibility.
"""

import re
import traceback
from datetime import date, datetime
from enum import StrEnum
from itertools import batched
from typing import Any

import polars as pl
from loguru import logger
from tortoise import Tortoise
from tortoise.expressions import Q
from tortoise.transactions import in_transaction

from easy_access.db.base import (
    copyright_item_from_dict,
    ensure_db_inited,
)
from easy_access.db.enums import (
    CLASSIFICATION_MAPPING_V1_TO_V2,
    Classification,
    ClassificationV2,
)
from easy_access.db.models import (
    CopyrightItem,
    Course,
    Faculty,
    Infringement,
    ItemUpdate,
    Organization,
    Person,
    StagedCopyrightItem,
    StagedFacultyUpdate,
    StagedProcessingFailure,
    Status,
    WorkflowStatus,
)
from easy_access.merge_rules import (
    build_merge_rules_from_settings,
    get_mergeable_fields,
)

# Import repositories for use in orchestration
from easy_access.repositories.item_repo import CopyrightItemRepository  # noqa: F401
from easy_access.repositories.staging_repo import StagingRepository  # noqa: F401
from easy_access.services.merge import (  # noqa: F401
    MergeConflictError,
    MergeError,
    TypeCastError,
    _cast_datetime_value,
    _cast_enum_value,
    _cast_numeric_value,
    _cast_values_for_comparison,
    _normalize_file_exists,
    compare_and_update_fields,
    record_field_change,
)

# Import from services for backward compatibility
# These are the canonical implementations - use services directly for new code
from easy_access.services.strategies import (  # noqa: F401
    DEFAULT_RANK,
    DateFieldStrategy,
    EnumFieldStrategy,
    FieldComparisonStrategy,
    FileExistsStrategy,
    NumericFieldStrategy,
    RankedFieldStrategy,
    StringFieldStrategy,
    get_comparison_strategy,
)
from easy_access.settings import Settings
from easy_access.utils import (
    safe_enum,
    safe_int,
    standardize_dataframe,
)


# Custom exceptions (re-exported from services, plus local ones)
class DatabaseOperationError(MergeError):
    """Raised when database operations fail."""

    pass


class ValidationError(MergeError):
    """Raised when data validation fails."""

    pass


# Constants
MIN_CHANGES_THRESHOLD = 3


# NOTE: Strategy classes have been moved to easy_access/services/strategies.py
# They are imported above for backward compatibility.

# NOTE: The following functions are imported from easy_access/services/merge.py:
# - record_field_change
# - compare_and_update_fields
# - _cast_datetime_value
# - _cast_numeric_value
# - _normalize_file_exists
# - _cast_enum_value
# - _cast_values_for_comparison


async def preprocess_input_data(
    data: pl.DataFrame | list[dict],
) -> tuple[list[dict], list[dict]]:
    """
    Standardize input and split it into items to create and items to update.

    When given a Polars DataFrame, the frame is standardized, existing material_ids are looked up,
    rows whose material_id does not exist are validated for required creation fields and returned
    as `new_items`, and rows whose material_id exists are returned as `update_items`.
    When given a list of dicts, the list is treated as `update_items`. Skipped candidate creations
    (with missing required fields) are logged and not returned.

    Parameters:
        data (pl.DataFrame | list[dict]): Input records as a Polars DataFrame or a list of dictionaries.

    Returns:
        tuple[list[dict], list[dict]]: A tuple (new_items, update_items) where `new_items` are dicts
        suitable for creating new records and `update_items` are dicts for updating existing records.
    """
    new_items = []
    update_items = []

    if isinstance(data, pl.DataFrame):
        data = standardize_dataframe(data)
        existing_mat_ids = await CopyrightItem.all().values("material_id")
        existing_mat_ids = {safe_int(m["material_id"]) for m in existing_mat_ids}
        existing_mat_ids = {m for m in existing_mat_ids if m is not None}

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


async def process_new_items(new_items: list[dict]) -> list[CopyrightItem]:
    """
    Process and create new copyright items in the database.

    Args:
        new_items: List of dictionaries representing new items to create

    Returns:
        List of created CopyrightItem objects

    Raises:
        DatabaseOperationError: When bulk creation fails
    """
    new_objects = []
    if new_items:
        new_objects = [await copyright_item_from_dict(item) for item in new_items]
        new_objects = [item for item in new_objects if item]
        try:
            await CopyrightItem.bulk_create(objects=new_objects)
        except Exception as e:
            logger.warning(
                f"Bulk creation failed: {e}. Attempting one-by-one creation."
            )
            failed_items = []
            for item in new_objects:
                try:
                    await item.save()
                except Exception as save_error:
                    logger.error(
                        f"Failed to save item {item.material_id}: {save_error}"
                    )
                    failed_items.append(item.material_id)

            if failed_items:
                raise DatabaseOperationError(
                    f"Failed to create items with material_ids: {failed_items}"
                ) from e
        logger.success(f"Created {len(new_objects)} new copyright items in db.")

    return new_objects


async def process_existing_items(
    update_items: list[dict],
    added_fields: dict,
    changeable_fields: dict,
    overwrite: bool = False,
) -> tuple[list[CopyrightItem], dict]:
    """
    Process updates to existing copyright items.

    Args:
        update_items: List of dictionaries representing items to update
        added_fields: Dictionary of field priorities for added fields
        changeable_fields: Dictionary of field priorities for changeable fields
        overwrite: Whether to overwrite existing values instead of using comparison logic

    Returns:
        Tuple of (changelist, updates) where changelist contains modified items
        and updates contains change details

    Raises:
        DatabaseOperationError: When processing individual items fails
    """
    logger.info(f"Updating {len(update_items)} existing items.")
    changelist = []
    updates = {}

    for new_item in update_items:
        try:
            if overwrite:
                changes, db_item = await _process_item_overwrite(
                    new_item, added_fields, changeable_fields
                )
            else:
                changes, db_item = await _process_item_normal(
                    new_item, added_fields, changeable_fields
                )

            if len(list(changes.keys())) >= MIN_CHANGES_THRESHOLD:
                changes["modified_at"] = datetime.now()
                updates[new_item.get("material_id")] = changes
                changelist.append(db_item)
        except DatabaseOperationError:
            # Re-raise database operation errors
            raise
        except Exception as e:
            # Log unexpected errors but continue processing other items
            logger.error(
                f"Unexpected error processing item {new_item.get('material_id')}: {e}"
            )
            continue

    return changelist, updates


async def _process_item_overwrite(
    new_item: dict, added_fields: dict, changeable_fields: dict
) -> tuple[dict, CopyrightItem | None]:
    """
    Process a single item in overwrite mode.

    Args:
        new_item: Dictionary with new field values
        added_fields: Dictionary of field priorities for added fields
        changeable_fields: Dictionary of field priorities for changeable fields

    Returns:
        Tuple of (changes dict, db_item)

    Raises:
        DatabaseOperationError: When database operations fail
    """
    try:
        db_item = await CopyrightItem.get(material_id=new_item.get("material_id"))

        changes = {
            "material_id": new_item.get("material_id"),
            "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

        logger.debug(f"now in overwrite function for {new_item.get('material_id')}")
        for k in changeable_fields | added_fields:
            logger.debug(f"checking field {k}")
            logger.debug(f"new_item.get(k): {new_item.get(k)}")
            logger.debug(f"getattr(db_item, k): {getattr(db_item, k)}")
            if new_item.get(k) is None:
                continue
            if new_item.get(k) != getattr(db_item, k):
                if str(new_item.get(k)) == getattr(db_item, k):
                    # if the new value is the same as the old value, skip it
                    continue
                changes = record_field_change(
                    changes,
                    k,
                    new_item.get(k),
                    getattr(db_item, k),
                    "[overwrite] new value != old value",
                    db_item,
                )
                logger.debug(f"changes: {changes}")
        # if any changes were made we'll have 3 or more keys in the changes dict
        # if not, no need to update the db
        logger.debug("final changes:")
        logger.debug(changes)
        if len(changes) < MIN_CHANGES_THRESHOLD:
            logger.debug(f"No changes for item {new_item.get('material_id')}.")
    except Exception as e:
        logger.warning(f"Could not update item {new_item.get('material_id')}: {e}")
        logger.warning(traceback.format_exc())
        raise DatabaseOperationError(
            f"Failed to process item {new_item.get('material_id')} in overwrite mode: {e}"
        ) from e

    return changes, db_item


async def _process_item_normal(
    new_item: dict, added_fields: dict, changeable_fields: dict
) -> tuple[dict, CopyrightItem | None]:
    """
    Process a single item in normal mode using comparison logic.

    Args:
        new_item: Dictionary with new field values
        added_fields: Dictionary of field priorities for added fields
        changeable_fields: Dictionary of field priorities for changeable fields

    Returns:
        Tuple of (changes dict, db_item)

    Raises:
        DatabaseOperationError: When database operations fail
    """
    try:
        db_item = await CopyrightItem.get(material_id=new_item.get("material_id"))
        changes = {}
        changes, db_item = compare_and_update_fields(
            new_item, db_item, added_fields, changes
        )
        changes, db_item = compare_and_update_fields(
            new_item, db_item, changeable_fields, changes
        )
    except Exception as e:
        logger.warning(f"Could not update item {new_item.get('material_id')}: {e}")
        raise DatabaseOperationError(
            f"Failed to process item {new_item.get('material_id')}: {e}"
        ) from e

    return changes, db_item


async def execute_bulk_database_operations(
    changelist: list[CopyrightItem],
    updates: dict,
    cur_user: str | dict | None,
    new_objects: list[CopyrightItem],
    update_relations: bool,
    settings: Settings,
) -> None:
    """
    Execute queued database writes for changed and newly created items and record item-level change logs.

    Parameters:
        changelist (list[CopyrightItem]): CopyrightItem objects with pending field changes.
        updates (dict): Mapping from `material_id` to a dict of changed field values (must include `material_id` and `update_time` keys; other keys are treated as changed fields).
        cur_user (str | dict | None): Current user identifier — either an email string or a dict containing an `"email"` key; used to populate `modified_by` on change records when present.
        new_objects (list[CopyrightItem]): Newly created CopyrightItem objects that may require relation updates.
        update_relations (bool): If true, schedule relation updates for affected items after applying changes.
        settings (Settings): Application settings / configuration context used by the operation.

    Returns:
        None
    """
    if changelist:
        # get all values from 'updates'
        # then get list of all distinct keys from all those dicts
        # then drop keys 'material_id' and 'update_time'
        # then add all those keys to the fields to update

        all_keys = {key for item in updates.values() for key in item}
        all_keys.discard("material_id")
        all_keys.discard("update_time")
        changed_fields = list(all_keys)
        changed_fields.append("modified_at")

        # process in batches of max 50:
        for change_batch in batched(changelist, 500):
            logger.info(
                f"Updating fields {changed_fields} for {len(change_batch)} items that were changed."
            )
            await CopyrightItem.bulk_update(change_batch, fields=changed_fields)

        if cur_user:
            user_email = (
                cur_user.get("email") if isinstance(cur_user, dict) else cur_user
            )
            [
                changes.update({"modified_by": user_email})
                for changes in updates.values()
            ]
            logger.info(f"items modified by {user_email}")

        for update_batch in batched(
            [(mat_id, changes) for mat_id, changes in updates.items()], 500
        ):
            logger.info(f"Updating {len(update_batch)} changelog items in db.")

            mat_ids = [mat_id for mat_id, _ in update_batch]
            await ItemUpdate.bulk_create(
                [
                    ItemUpdate(change_details=changes, material_id=mat_id)
                    for mat_id, changes in update_batch
                ]
            )

            # now add each ItemUpdate as a m2m relation to the corresponding CopyrightItem
            for mat_id in mat_ids:
                item = await CopyrightItem.get(material_id=mat_id)
                update = (
                    await ItemUpdate.filter(material_id=mat_id)
                    .order_by("-created_at")
                    .first()
                )
                await item.changes.add(update)

    if (changelist or new_objects) and update_relations:
        logger.success("Updating relations for all CopyrightItems.")


# NOTE: The following functions have been moved to easy_access/services/merge.py:
# - record_field_change (imported above)
# - compare_and_update_fields (imported above)
# - _cast_values_for_comparison (imported above)
# - _cast_datetime_value (imported above)
# - _cast_numeric_value (imported above)
# - _normalize_file_exists (imported above)
# - _cast_enum_value (imported above)


class DataSource(StrEnum):
    """Enum for the source of the data.

    This is used to determine how to handle the data when updating the database.
    """

    RAW_QLIK_DATA = "raw_qlik_data"
    WEEKLY_SHEET = "weekly_sheet"
    OVERVIEW_SHEET = "overview_sheet"
    EA_SCRIPT = "ea_script"
    WEB_DASHBOARD = "dashboard"


async def update_copyright_items(
    settings: Settings,
    data: pl.DataFrame | list[dict],
    update_relations: bool = True,
    overwrite: bool = False,
    user_info: dict | None = None,
) -> None:
    """
    Update the db with copyrightitems from the dataframe (or pre-filtered list of dicts from a df).
    Adds new if they don't exist, or updates if they do.
    See `compare_items` and the dicts added_fields, changeable_fields, core_fields for details on how the comparison is done.

    Once done, and if any updates were made, will call `update_copyright_relations` to update the m2m relations.

    Options:
    - `update_relations`: if True, will call `update_copyright_relations` to update the m2m relations after the update.
    - `overwrite`: if True, will overwrite the existing items in the db with the new ones instead of using the comparison logic.
    """

    if user_info is None:
        user_info = dict()

    cur_user = (
        user_info.get("email")
        if user_info.get("email")
        else {"email": "cip-admin@utwente.nl"}
    )

    await ensure_db_inited(settings)
    # Build merge rules dynamically from settings
    added_fields, changeable_fields = build_merge_rules_from_settings(settings)

    logger.info(f"Received {len(data)} raw copyright items as input for an update.")

    # Preprocess input data
    new_items, update_items = await preprocess_input_data(data)

    # Process new items
    new_objects = await process_new_items(new_items)

    # Process existing items
    changelist, updates = await process_existing_items(
        update_items, added_fields, changeable_fields, overwrite
    )

    # Execute bulk database operations
    await execute_bulk_database_operations(
        changelist, updates, cur_user, new_objects, update_relations, settings
    )

    logger.success("Done updating CopyrightItems!")
    await Tortoise.close_connections()


async def process_staged_raw_data(settings: Settings) -> None:
    """
    Processes the staged raw data and updates the main CopyrightItem table.
    """
    await ensure_db_inited(settings)
    staged_items = await StagedCopyrightItem.all()
    if not staged_items:
        logger.info("No staged raw data to process.")
        return

    logger.info(f"Processing {len(staged_items)} staged raw items...")

    # Helper: list of known staged fields (keeps mapping explicit and safe)
    staged_fields = [
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

    processed_ids: list[int] = []

    for batch_idx, batch in enumerate(batched(staged_items, 500)):
        batch_processed_ids: list[int] = []
        complex_item_dicts: list[dict] = []
        logger.info(f"Processing batch {batch_idx} with {len(batch)} items")
        async with in_transaction():
            for _item_idx, staged_item in enumerate(batch):
                try:
                    # Build a safe dict from known fields
                    item_dict: dict = {}
                    for f in staged_fields:
                        # Use getattr to avoid ORM internals
                        item_dict[f] = getattr(staged_item, f, None)

                    # Ensure material_id is present and castable
                    mid = item_dict.get("material_id")
                    item_dict.get("faculty")
                    if mid is None:
                        # logger.warning(
                        #    f"[STAGED][SKIP] material_id=None, faculty={faculty_val}, stage=raw_data: Missing required material_id"
                        # )
                        continue

                    # logger.debug(
                    #    f"[STAGED][PROCESS] material_id={mid}, faculty={faculty_val}, stage=raw_data: Starting processing"
                    # )

                    existing_item = await CopyrightItem.get_or_none(material_id=mid)

                    if not existing_item:
                        # Create new item using canonical normalizer
                        # logger.debug(
                        #    f"[STAGED][CREATE] material_id={mid}, faculty={faculty_val}, stage=raw_data: Creating new item"
                        # )
                        new_item = await copyright_item_from_dict(item_dict)
                        if new_item:
                            await new_item.save()
                            smid = safe_int(mid)
                            if smid is not None:
                                batch_processed_ids.append(smid)
                            # logger.info(
                            #    f"[STAGED][SUCCESS] material_id={mid}, faculty={faculty_val}, stage=raw_data: Created new item"
                            # )
                        else:
                            # logger.warning(
                            #    f"[STAGED][FAIL] material_id={mid}, faculty={faculty_val}, stage=raw_data: Failed to create item from dict"
                            # )
                            ...
                    else:
                        # Check if this item has complex fields that need merging
                        mergeable_fields = get_mergeable_fields()
                        has_complex_fields = False
                        for field in mergeable_fields:
                            if item_dict.get(field) is not None:
                                has_complex_fields = True
                                break

                        if has_complex_fields:
                            # Delegate to canonical merge path
                            # logger.debug(
                            #    f"[STAGED][MERGE] material_id={mid}, faculty={faculty_val}, stage=raw_data: Delegating to complex merge"
                            # )
                            complex_item_dicts.append(item_dict)
                            smid = safe_int(mid)
                            if smid is not None:
                                batch_processed_ids.append(smid)
                        else:
                            # Conservative updates for trivial fields
                            update_fields = []
                            # status: try to coerce to Status enum safely
                            status_val = item_dict.get("status")
                            if status_val:
                                new_status = safe_enum(Status, status_val)
                                if new_status and existing_item.status != new_status:
                                    existing_item.status = new_status
                                    update_fields.append("status")

                            # last_change: accept datetime/date or parse common string formats
                            lc_val = item_dict.get("last_change")
                            if lc_val:
                                parsed_date = None
                                if isinstance(lc_val, datetime):
                                    parsed_date = lc_val.date()
                                elif isinstance(lc_val, date):
                                    parsed_date = lc_val
                                elif isinstance(lc_val, str):
                                    try:
                                        # try isoformat first
                                        parsed_dt = datetime.fromisoformat(lc_val)
                                        parsed_date = parsed_dt.date()
                                    except Exception:
                                        for fmt in (
                                            "%Y-%m-%d %H:%M:%S%z",
                                            "%Y-%m-%d %H:%M:%S",
                                            "%Y-%m-%d",
                                        ):
                                            try:
                                                parsed_dt = datetime.strptime(
                                                    lc_val, fmt
                                                )
                                                parsed_date = parsed_dt.date()
                                                break
                                            except Exception:
                                                continue
                                if (
                                    parsed_date
                                    and existing_item.last_change != parsed_date
                                ):
                                    existing_item.last_change = parsed_date
                                    update_fields.append("last_change")

                            if update_fields:
                                # logger.debug(
                                #    f"[STAGED][UPDATE] material_id={mid}, faculty={faculty_val}, stage=raw_data: Updating fields {update_fields}"
                                # )
                                await existing_item.save(update_fields=update_fields)
                                smid = safe_int(mid)
                                if smid is not None:
                                    batch_processed_ids.append(smid)
                                # logger.info(
                                #    f"[STAGED][SUCCESS] material_id={mid}, faculty={faculty_val}, stage=raw_data: Updated existing item"
                                # )

                except Exception as e:
                    err_msg = str(e)
                    getattr(staged_item, "material_id", None)
                    getattr(staged_item, "faculty", None)
                    # logger.error(
                    #    f"[STAGED][ERROR] material_id={mid_val}, faculty={faculty_val}, stage=raw_data: {err_msg}"
                    # )
                    # logger.debug(
                    #    f"[STAGED][TRACE] material_id={mid_val}, faculty={faculty_val}, stage=raw_data: {traceback.format_exc()}"
                    # )
                    # Record failure in the DB for later inspection/retry
                    try:
                        payload = {
                            f: getattr(staged_item, f, None) for f in staged_fields
                        }
                        await StagedProcessingFailure.create(
                            material_id=safe_int(
                                getattr(staged_item, "material_id", None)
                            ),
                            staged_payload=payload,
                            error_message=err_msg[:1900],
                        )
                        # logger.info(
                        #    f"[STAGED][RECORDED] material_id={mid_val}, faculty={faculty_val}, stage=raw_data: Failure recorded in StagedProcessingFailure"
                        # )
                    except Exception:
                        # logger.error(
                        #    f"[STAGED][RECORD_FAIL] material_id={mid_val}, faculty={faculty_val}, stage=raw_data: Could not record failure: {record_error}"
                        # )
                        ...
                    # Do not re-raise; keep other rows processing. Failed staged rows remain for manual inspection.

        # Process complex merges outside transaction since update_copyright_items does its own operations
        if complex_item_dicts:
            logger.info(
                f"Processing {len(complex_item_dicts)} complex merges for batch {batch_idx}"
            )
            try:
                await update_copyright_items(
                    settings=settings, data=complex_item_dicts, update_relations=False
                )
                logger.info(
                    f"Successfully processed {len(complex_item_dicts)} complex merges"
                )
            except Exception as e:
                logger.exception(f"Error processing complex merges: {e}")
                # Don't fail the whole batch, just log the error

        # After successful transaction, remove successfully processed staged rows
        if batch_processed_ids:
            try:
                await StagedCopyrightItem.filter(
                    material_id__in=batch_processed_ids
                ).delete()
                logger.info(
                    f"Cleared {len(batch_processed_ids)} processed staged rows."
                )
                processed_ids.extend(batch_processed_ids)
            except Exception as e:
                logger.exception(
                    f"Error deleting staged rows {batch_processed_ids}: {e}"
                )

    logger.info(
        f"Finished processing staged raw data. Successfully processed {len(processed_ids)} rows."
    )


async def process_staged_faculty_updates(settings: Settings) -> None:
    """
    Processes the staged faculty updates and updates the main CopyrightItem table.
    """

    def _normalize_wf(val: str) -> str | None:
        if not val:
            return None
        s = str(val).strip()
        # direct match to enum values
        enum_vals = {e.value for e in WorkflowStatus}
        if s in enum_vals:
            return s
        # member names
        if s in WorkflowStatus.__members__:
            return WorkflowStatus[s].value
        lower = s.lower()
        mapping = {
            "todo": WorkflowStatus.ToDo.value,
            "to do": WorkflowStatus.ToDo.value,
            "inbox": WorkflowStatus.ToDo.value,
            "in_progress": WorkflowStatus.InProgress.value,
            "in progress": WorkflowStatus.InProgress.value,
            "inprogress": WorkflowStatus.InProgress.value,
            "in-progress": WorkflowStatus.InProgress.value,
            "done": WorkflowStatus.Done.value,
        }
        return mapping.get(lower)

    await ensure_db_inited(settings)
    staged_updates = await StagedFacultyUpdate.all()
    if not staged_updates:
        logger.info("No staged faculty updates to process.")
        return

    logger.info(f"Processing {len(staged_updates)} staged faculty updates...")

    # Process staged faculty updates in batches inside transactions; delete only processed rows
    processed_updates: list[int] = []
    for batch in batched(staged_updates, 100):
        batch_processed: list[int] = []
        async with in_transaction():
            for update in batch:
                try:
                    mid = update.material_id
                    # logger.debug(
                    #    f"[STAGED][PROCESS] material_id={mid}, stage=faculty_update: Starting processing"
                    # )

                    item = await CopyrightItem.get_or_none(material_id=mid)
                    if not item:
                        # logger.warning(
                        #    f"[STAGED][SKIP] material_id={mid}, stage=faculty_update: Item not found in database"
                        # )
                        continue

                    update_fields = []
                    if (
                        update.manual_classification
                        and item.manual_classification != update.manual_classification
                    ):
                        item.manual_classification = update.manual_classification
                        update_fields.append("manual_classification")
                        # logger.debug(
                        #    f"[STAGED][UPDATE] material_id={mid}, stage=faculty_update: Updating manual_classification"
                        # )

                    if update.remarks and item.remarks != update.remarks:
                        item.remarks = update.remarks
                        update_fields.append("remarks")
                        # logger.debug(
                        #    f"[STAGED][UPDATE] material_id={mid}, stage=faculty_update: Updating remarks"
                        # )

                    if update.workflow_status:
                        # Accept explicit workflow status choices made by faculty users.
                        # Normalize common variants so things like "inbox", "in_progress",
                        # "inprogress", "todo" (case-insensitive) are mapped to the
                        # canonical enum values. If a recognizable value is found, always
                        # apply it (this represents an explicit user choice).

                        normalized = _normalize_wf(update.workflow_status)
                        if normalized:
                            wf_st = safe_enum(WorkflowStatus, normalized)
                            if wf_st and item.workflow_status != wf_st:
                                item.workflow_status = wf_st
                                update_fields.append("workflow_status")
                            # logger.debug(
                            #    f"[STAGED][UPDATE] material_id={mid}, stage=faculty_update: Updating workflow_status"
                            # )

                    if update_fields:
                        await item.save(update_fields=update_fields)
                        smid = safe_int(mid)
                        if smid is not None:
                            batch_processed.append(smid)
                        # logger.info(
                        #    f"[STAGED][SUCCESS] material_id={mid}, stage=faculty_update: Updated fields {update_fields}"
                        # )
                    else:
                        # logger.debug(
                        #    f"[STAGED][SKIP] material_id={mid}, stage=faculty_update: No fields to update"
                        # )
                        ...

                except Exception:
                    getattr(update, "material_id", None)
                    # logger.error(
                    #    f"[STAGED][ERROR] material_id={mid_val}, stage=faculty_update: {str(e)}"
                    # )
                    # logger.debug(
                    #    f"[STAGED][TRACE] material_id={mid_val}, stage=faculty_update: {traceback.format_exc()}"
                    # )

        if batch_processed:
            try:
                await StagedFacultyUpdate.filter(
                    material_id__in=batch_processed
                ).delete()
                # logger.info(
                #    f"Cleared {len(batch_processed)} processed staged faculty updates."
                # )
                processed_updates.extend(batch_processed)
            except Exception:
                logger.exception(
                    f"Error deleting processed staged faculty updates: {batch_processed}"
                )

    logger.info(
        f"Finished processing staged faculty updates. Successfully processed {len(processed_updates)} rows."
    )


async def calculate_derived_fields(settings: Settings) -> None:
    """
    Calculates and updates derived fields like 'possible_fine' and 'infringement'
    for all copyright items.
    """
    from easy_access.db.retrieve import retrieve_copyright_items

    df = retrieve_copyright_items(
        settings=settings,
        additional_cols=[
            "possible_fine",
            "infringement",
            "manual_classification",
            "pages_x_students",
        ],
    )

    if df.is_empty():
        logger.info("No items found to calculate derived fields for.")
        return

    # Calculate possible_fine
    df = df.with_columns(
        pl.when(pl.col("possible_fine").is_null())
        .then(
            pl.col("pages_x_students")
            .cast(pl.Int64, strict=False)
            .fill_null(0)
            .mul(settings.fine_amount)
        )
        .otherwise(pl.col("possible_fine"))
        .alias("possible_fine")
    )

    # Calculate infringement
    df = df.with_columns(
        pl.when(
            pl.col("manual_classification").is_null()
            | (pl.col("manual_classification") == "")
            | (pl.col("manual_classification") == "-")
        )
        .then(pl.lit(Infringement.UNDETERMINED.value))
        .when(
            pl.col("manual_classification")
            .str.to_lowercase()
            .str.contains("open|eigen|overig|deleted")
        )
        .then(pl.lit(Infringement.NO.value))
        .when(pl.col("manual_classification").str.to_lowercase().str.contains("lange"))
        .then(pl.lit(Infringement.YES.value))
        .otherwise(pl.lit(Infringement.MAYBE.value))
        .alias("infringement")
    )

    update_df = df.select(["material_id", "possible_fine", "infringement"])

    await update_copyright_items(settings=settings, data=update_df, overwrite=True)
    logger.info("Finished calculating derived fields.")


async def persist_courses(
    courses_data: dict[int, dict],
    *,
    CourseModel=Course,
    PersonModel=Person,
    FacultyModel=Faculty,
) -> None:
    """Persist (upsert) course records and teacher relations.

    Accepts dependency-injected models so tests patching objects in the
    enrichment.osiris module still work when that wrapper forwards its
    patched classes here.
    """
    if not courses_data:
        logger.info("No course data to persist")
        return

    # Split create/update
    existing = await CourseModel.filter(cursuscode__in=list(courses_data.keys()))
    existing_codes = {c.cursuscode for c in existing}

    to_create: list[dict] = []
    to_update: list[dict] = []

    # Preprocess each course dict
    allowed_course_fields = {
        "cursuscode",
        "internal_id",
        "year",
        "name",
        "short_name",
        "ec",
        "programme",
        "notes",
        "category",
        "faculty_id",
    }

    for code, data in courses_data.items():
        if not isinstance(data, dict):
            logger.warning(f"Skipping invalid course data for {code}: not a dict")
            continue
        # Shallow copy so we can mutate safely
        cd = dict(data)

        # Handle faculty FK (stored by abbreviation). Use faculty_id convention.
        faculty_abbr = cd.pop("faculty", None)
        if faculty_abbr:
            faculty_obj = await FacultyModel.get_or_none(abbreviation=faculty_abbr)
            if faculty_obj:
                cd["faculty_id"] = faculty_obj.abbreviation
            else:
                logger.debug(
                    f"Faculty '{faculty_abbr}' not found for course {code}; leaving FK null"
                )

        # Drop unsupported keys (e.g. faculty_long, language, etc.)
        cd = {
            k: v
            for k, v in cd.items()
            if k in allowed_course_fields or k.startswith("_")
        }

        if "ec" in cd:
            if "," in str(cd["ec"]):
                cd["ec"] = cd["ec"].replace(",", ".")
            try:
                cd["ec"] = float(cd["ec"])
            except (ValueError, TypeError):
                cd["ec"] = None
        if code in existing_codes:
            to_update.append(cd | {"cursuscode": code})
        else:
            # Required minimal fields guard
            missing_req = [
                k for k in ["cursuscode", "internal_id", "year", "name"] if k not in cd
            ]
            if missing_req:
                logger.warning(
                    f"Skipping create for course {code}: missing {missing_req}"
                )
                continue
            cd["cursuscode"] = code
            to_create.append(cd)

    # Create
    for cd in to_create:
        try:
            await CourseModel.create(**cd)
        except Exception as exc:  # pragma: no cover (defensive)
            logger.error(f"Error creating course {cd.get('cursuscode')}: {exc}")
            continue

    # Update existing (exclude PK)
    for ud in to_update:
        code = ud.pop("cursuscode")
        try:
            await CourseModel.filter(cursuscode=code).update(**ud)
        except Exception as exc:  # pragma: no cover
            logger.error(f"Error updating course {code}: {exc}")

    logger.info(f"Successfully persisted {len(courses_data)} courses")


async def _apply_course_teacher_relations(course_obj, rel_payload: dict, PersonModel):
    """Handle teacher/person many-to-many assignments for a course.

    We unify all available role sets into a single collection for now; role-specific
    data could be added by creating CourseEmployee entries with a role value.
    """
    if not course_obj or not rel_payload:
        return
    # Aggregate teacher-like sets
    teacher_sets = []
    for key in [
        "teachers",
        "contacts",
        "docenten",
        "examinators",
        "unknown_role",
        "tutors",
    ]:
        val = rel_payload.get(key)
        if isinstance(val, set | list | tuple):
            teacher_sets.append(set(val))
    if not teacher_sets:
        return
    all_teachers = set.union(*teacher_sets)
    for name in sorted(all_teachers):
        if not name or not str(name).strip():
            continue
        person_obj, _created = await PersonModel.get_or_create(
            input_name=str(name).strip(), defaults={"main_name": None}
        )
        try:
            await course_obj.teachers.add(person_obj)
        except Exception as exc:  # pragma: no cover
            logger.debug(
                f"Could not add teacher '{name}' to course {course_obj.cursuscode}: {exc}"
            )


async def persist_persons(
    persons_data: dict[str, dict],
    *,
    PersonModel=Person,
    FacultyModel=Faculty,
    OrganizationModel=Organization,
) -> None:
    """Persist (upsert) person records and their organization relations.

    Drops keys that don't map to Person columns; handles FK + M2M after base create/update.
    """
    if not persons_data:
        logger.info("No person data to persist")
        return

    existing = await PersonModel.filter(input_name__in=list(persons_data.keys()))
    existing_names = {p.input_name for p in existing}

    to_create: list[dict] = []
    to_update: list[dict] = []

    # Allowed direct columns (excluding M2M + unserialized fields)
    direct_fields = {
        "input_name",
        "main_name",
        "match_confidence",
        "first_name",
        "email",
        "people_page_url",
    }

    for input_name, pdata in persons_data.items():
        if not isinstance(pdata, dict):
            logger.warning(f"Skipping invalid person data for {input_name}: not a dict")
            continue
        pd = dict(pdata)
        faculty_abbr = pd.pop("faculty", None)
        if faculty_abbr:
            faculty_obj = await FacultyModel.get_or_none(abbreviation=faculty_abbr)
            if faculty_obj:
                pd["faculty_id"] = faculty_obj.abbreviation
            else:
                logger.debug(
                    f"Faculty '{faculty_abbr}' not found for person {input_name}"
                )
        # Stash org info
        orgs_payload = pd.pop("orgs", [])
        pd["_orgs_payload"] = orgs_payload
        # Drop unsupported keys
        cleaned = {
            k: v
            for k, v in pd.items()
            if k in direct_fields or k.endswith("_id") or k.startswith("_")
        }
        cleaned["input_name"] = input_name  # ensure primary identifier present
        if input_name in existing_names:
            to_update.append(cleaned)
        else:
            to_create.append(cleaned)

    # Create
    for cd in to_create:
        orgs_payload = cd.pop("_orgs_payload", [])
        try:
            person_obj = await PersonModel.create(
                **{k: v for k, v in cd.items() if not k.startswith("_")}
            )
        except Exception as exc:  # pragma: no cover
            logger.error(f"Error creating person {cd.get('input_name')}: {exc}")
            continue
        await _apply_person_org_relations(person_obj, orgs_payload, OrganizationModel)

    # Update
    for ud in to_update:
        orgs_payload = ud.pop("_orgs_payload", [])
        input_name = ud.pop("input_name")
        try:
            await PersonModel.filter(input_name=input_name).update(
                **{k: v for k, v in ud.items() if not k.startswith("_")}
            )
            person_obj = await PersonModel.get_or_none(input_name=input_name)
            if person_obj:
                await _apply_person_org_relations(
                    person_obj, orgs_payload, OrganizationModel
                )
        except Exception as exc:  # pragma: no cover
            logger.error(f"Error updating person {input_name}: {exc}")

    logger.info(f"Successfully persisted {len(persons_data)} persons")


async def _apply_person_org_relations(
    person_obj: Person,
    orgs_payload: dict[str, str],
    OrganizationModel: type[Organization],
):
    if not person_obj or not orgs_payload:
        return
    for org in orgs_payload:
        if not isinstance(org, dict):
            continue
        raw_abbr = org.get("abbr") or org.get("abbreviation") or org.get("name")
        name = org.get("name") or raw_abbr
        if not raw_abbr:
            continue
        full_abbr = raw_abbr  # provided chain (e.g. ET-CEM-MD)
        base_abbr = full_abbr.split("-")[-1] if full_abbr else full_abbr
        hierarchy_level = full_abbr.count("-") + 1 if full_abbr else 1

        # Prefer lookup by full_abbreviation (unique); fallback to base abbreviation
        org_obj = await OrganizationModel.get_or_none(full_abbreviation=full_abbr)
        if not org_obj:
            try:
                org_obj = await OrganizationModel.get_or_none(abbreviation=base_abbr)
            except Exception:
                # probably multiple with the same 'abbreviation'
                # instead filter on abbreviation and hierarchy_level
                org_obj_filtered = OrganizationModel.filter(
                    full_abbreviation=full_abbr, name=name
                )
                num_found = await org_obj_filtered.count()
                # if exactly one match, use it
                if not num_found:
                    org_obj = None
                elif num_found == 1:
                    org_obj = await org_obj_filtered.first()
                else:
                    logger.error(
                        f"Found multiple organizations matching abbreviation='{base_abbr}', hierarchy_level={hierarchy_level}, name='{name}'; cannot disambiguate, skipping"
                    )
                    org_obj = None
        if not org_obj:
            try:
                org_obj = await OrganizationModel.create(
                    parent_organization=None,
                    hierarchy_level=hierarchy_level,
                    name=name,
                    abbreviation=base_abbr,
                    full_abbreviation=full_abbr,
                )
            except Exception as exc:  # pragma: no cover
                # Retry fetch in case of race creating same full_abbreviation
                existing_retry = await OrganizationModel.get_or_none(
                    full_abbreviation=full_abbr
                )
                if existing_retry:
                    org_obj = existing_retry
                else:
                    logger.debug(
                        f"Could not create organization '{full_abbr}' for person {person_obj.input_name}: {exc}"
                    )
                    continue
        try:
            await person_obj.orgs.add(org_obj)
        except Exception as exc:  # pragma: no cover
            logger.debug(
                f"Could not add org '{full_abbr}' to person {person_obj.input_name}: {exc}"
            )


async def update_workflow_status_from_db(settings: Settings) -> None:
    """
    ensures that workflow status matches the current state of the item,
    e.g. if a manual_classification is present that requires no actions, set workflow_status to Done,
    if a file does not exist, also set it to Done, etc.
    """
    await ensure_db_inited(settings)
    # grab all items with non-Done workflow status
    items = await CopyrightItem.filter(~Q(workflow_status=WorkflowStatus.Done)).all()

    DONE_MANUAL_CLASSIFICATIONS = [
        Classification.OPEN_ACCESS.value,
        Classification.EIGEN_MATERIAAL_POWERPOINT.value,
        Classification.EIGEN_MATERIAAL_OVERIG.value,
        Classification.EIGEN_MATERIAAL_TITELINDICATIE.value,
        Classification.EIGEN_MATERIAAL.value,
    ]

    if not items:
        logger.info("No items to update workflow status for.")
        return
    logger.info(f"Updating workflow status for {len(items)} items...")
    updated_count = 0
    for item in items:
        if item.file_exists is False:
            item.workflow_status = WorkflowStatus.Done
            await item.save(update_fields=["workflow_status"])
            updated_count += 1
            continue
        if (
            item.manual_classification
            and item.manual_classification.lower() in DONE_MANUAL_CLASSIFICATIONS
        ):
            logger.info(
                f"Updating item {item.material_id} to Done based on manual_classification '{item.manual_classification}'"
            )
            item.workflow_status = WorkflowStatus.Done
            await item.save(update_fields=["workflow_status"])
            updated_count += 1
            continue


async def map_v1_to_v2_classifications(settings: Settings) -> None:
    """
    Map v1 manual classification values to v2 classification fields for items that lack a v2 classification.

    Selects CopyrightItem rows whose `v2_manual_classification` is null or `ONBEKEND`, normalizes their existing v1 `manual_classification` values, looks up the corresponding v2 mapping via CLASSIFICATION_MAPPING_V1_TO_V2, and updates the item's `v2_manual_classification`, `v2_lengte`, and `v2_overnamestatus` when a mapped value differs from the current v2 fields. Operates directly on the database using Tortoise ORM within a transaction, records per-item mapping details for optional export, and logs summary counts of mapped, modified, failed, and unlinked items.

    Parameters:
        settings (Settings): Application settings used to ensure the database is initialized and to drive ORM interactions.
    """
    # select all items:
    # - without a v2 classification (null or empty)
    await ensure_db_inited(settings)
    selected_items = (
        await CopyrightItem.filter(
            Q(v2_manual_classification__isnull=True)
            | Q(v2_manual_classification=ClassificationV2.ONBEKEND)
        )
        .all()
        .prefetch_related("v1_items", "faculty")
    )
    logger.info(f"Mapping v1 to v2 classifications for {len(selected_items)} items...")
    if not selected_items:
        logger.info("No items to map v1 to v2 classifications for.")
        return
    # add mapping logic here, see add_v2_classification for reference
    # relevant fields on CopyrightItem:
    # - manual_classification (v1)
    # - v2_manual_classification
    # - v2_lengte
    # - v2_overnamestatus

    # perform mapping using the same lookup as the old DataFrame-based helper
    details = []
    mapped_count = 0
    failed_count = 0
    modified_count = 0
    unlinked_upd = 0
    async with in_transaction():
        for item in selected_items:
            detaildict = {}
            try:
                input_val = item.manual_classification
                if not isinstance(input_val, str):
                    try:
                        input_val = input_val.value
                    except Exception:
                        input_val = str(input_val)

                current = input_val or "onbekend"
                if not current or current == "-":
                    current = "onbekend"

                # normalize common variations to improve enum lookup
                if isinstance(current, str):
                    current = current.strip().lower()
                # try to coerce to the v1 Classification enum; fall back to ONBEKEND
                # use a match case statement for this.
                # 1. exact match to enum values
                # 2. match ignoring case
                # 3. match ignoring underscores, hyphens, spaces, and case
                # 4. default to ONBEKEND
                match current:
                    case val if val in {e.value for e in Classification}:
                        key = Classification(val)
                    case val if val.lower() in {
                        e.value.lower() for e in Classification
                    }:
                        key = Classification(
                            next(
                                e.value
                                for e in Classification
                                if e.value.lower() == val.lower()
                            )
                        )
                    case val if re.sub(r"[\s_-]", "", val.lower()) in {
                        re.sub(r"[\s_-]", "", e.value.lower()) for e in Classification
                    }:
                        key = Classification(
                            next(
                                e.value
                                for e in Classification
                                if re.sub(r"[\s_-]", "", e.value.lower())
                                == re.sub(r"[\s_-]", "", val.lower())
                            )
                        )
                    case _:
                        key = Classification.ONBEKEND

                mapped = CLASSIFICATION_MAPPING_V1_TO_V2.get(key)
                if not mapped:
                    mapped = CLASSIFICATION_MAPPING_V1_TO_V2[Classification.ONBEKEND]

                current_v2_classification = item.v2_manual_classification.value

                if item.v2_manual_classification != mapped.classification:
                    item.v2_manual_classification = mapped.classification
                    item.v2_lengte = mapped.length
                    item.v2_overnamestatus = mapped.overname_status
                    faculty = item.faculty
                    abbreviation = faculty.abbreviation
                    v1_items = await item.v1_items.all()
                    v1_id = None
                    if v1_items:
                        v1_id = v1_items[0].material_id
                    if not v1_id:
                        unlinked_upd += 1
                    detaildict = {
                        "material_id": item.material_id,
                        "faculty": abbreviation,
                        "v1_material_id": v1_id,
                        "found_v1_classification": input_val,
                        "v2_classification_before_update": current_v2_classification,
                        "used_v1_classification": key.value,
                        "mapped_v2_classification": mapped.classification.value,
                        "mapped_v2_length": mapped.length.value,
                        "mapped_v2_overname_status": mapped.overname_status.value,
                    }
                    details.append(detaildict)
                    logger.debug(
                        f"material_id {item.material_id}: [v1] {input_val} -> {current} -> {key.value} mapped to [v2] {mapped.classification}"
                    )
                    await item.save(
                        update_fields=[
                            "v2_manual_classification",
                            "v2_lengte",
                            "v2_overnamestatus",
                        ]
                    )
                    modified_count += 1
                mapped_count += 1
            except Exception as exc:
                logger.error(
                    f"Failed to map v1->v2 classification for material_id={getattr(item, 'material_id', None)}: {exc}"
                )
                failed_count += 1

    logger.info(
        f"Finished mapping v1->v2 classifications: mapped={mapped_count}, failed={failed_count}, modified={modified_count}, unlinked updates={unlinked_upd} (out of {len(selected_items)})"
    )
    if len(details) > 0:
        logger.debug("Stored parsed/mapped details to csv for inspection")
        try:
            pl.DataFrame(details).write_csv(
                "v1_to_v2_classification_mapping_details.csv"
            )
        except Exception:
            details = [{a: str(b) for a, b in d.items()} for d in details]
            try:
                pl.DataFrame(details).write_csv(
                    "v1_to_v2_classification_mapping_details.csv"
                )
            except Exception as exc2:
                logger.error(f"Could not write mapping details to csv: {exc2}")
                if len(details) < 20:
                    logger.debug(f"Mapping details: {details}")
                else:
                    logger.debug(f"Mapping details: {details[:20]} ... (truncated)")
