"""
This module contains functions to update existing data in the database.
It includes logic for linking related entities (like LLM classifications or courses
to copyright items) and the core function `update_copyright_items` which handles
complex update logic based on heuristics and data source.
"""

import logging
import traceback
from datetime import UTC, datetime
from enum import (  # Enum was already imported, StrEnum is more specific if used
    Enum,
    StrEnum,
)
from typing import Any

import polars as pl
from tortoise import Tortoise

from easy_access.db.base import (  # Renamed init for clarity
    copyright_item_from_dict,
    standardize_dataframe,
)
from easy_access.db.base import init as init_tortoise
from easy_access.db.models import (
    PDF,  # Not directly used here but good for context if functions were expanded
    Classification,  # Used in changeable_fields
    CopyrightItem,
    Course,
    Infringement,  # Used in added_fields
    ItemUpdate,
    LLMClassification,
    Status,
    WorkflowStatus,  # Used in added_fields
)
from easy_access.utils import (
    determine_course_code,  # Assuming this is still the location
)

logger = logging.getLogger(__name__)


async def link_llm_classifications_to_copyright_items() -> None:
    """
    Links LLMClassification records to their corresponding CopyrightItem records.
    It iterates through LLM classifications that are not yet linked to an item
    and attempts to establish the link based on `used_material_id`.
    """
    await init_tortoise()
    try:
        classifications_to_link = await LLMClassification.filter(
            item__isnull=True
        ).all()
        linked_count = 0
        for classification in classifications_to_link:
            try:
                material_id = classification.used_material_id
                item = await CopyrightItem.get_or_none(material_id=material_id)
                if item:
                    item.llm_classification = classification
                    await item.save(
                        update_fields=["llm_classification_id", "modified_at"]
                    )  # Be specific
                    linked_count += 1
            except Exception as e:
                logger.warning(
                    f"Error linking LLM classification for material_id {classification.used_material_id} "
                    f"to CopyrightItem: {e}"
                )
        logger.info(
            f"Successfully linked {linked_count} of {len(classifications_to_link)} LLM classifications."
        )
    except Exception as e:
        logger.error(
            f"General error in link_llm_classifications_to_copyright_items: {e}"
        )
        logger.debug(traceback.format_exc())
    finally:
        await Tortoise.close_connections()


async def link_courses_to_copyright_items() -> None:
    """
    Links CopyrightItem records to Course records based on course codes.
    It uses `determine_course_code` to extract potential course codes from
    copyright item data and then attempts to link them to existing courses.
    """
    await init_tortoise()
    try:
        items_w_prefetch = (
            await CopyrightItem.all()
        )  # Consider prefetching faculty if used by determine_course_code indirectly
        logger.info(
            f"Retrieved {len(items_w_prefetch)} copyright items for course linking."
        )

        links_added: int = 0
        course_codes_processed: int = (
            0  # Tracks how many codes were attempted, not just found
        )

        for item in items_w_prefetch:
            # Assuming determine_course_code is robust enough.
            # It might be more efficient to batch course_code determinations if it involves DB/heavy ops.
            course_codes_str: set[str] = determine_course_code(
                item.course_code, item.course_name
            )

            if not course_codes_str:
                logger.debug(
                    f"No course codes determined for item {item.material_id} "
                    f"(course_code: '{item.course_code}', course_name: '{item.course_name}')."
                )
                continue

            for code_str in course_codes_str:
                course_codes_processed += 1
                if (
                    not code_str or not code_str.isdigit()
                ):  # Skip if empty or not purely numeric
                    logger.debug(
                        f"Invalid course code format '{code_str}' for item {item.material_id}."
                    )
                    continue
                try:
                    cursuscode = int(code_str)
                    course = await Course.get_or_none(cursuscode=cursuscode)
                    if course:
                        # Add relationship if not already present. Tortoise handles duplicates in add.
                        await item.courses.add(course)
                        links_added += 1
                        logger.debug(
                            f"Linked item {item.material_id} to course {cursuscode}."
                        )
                    else:
                        logger.debug(
                            f"Course with code {cursuscode} not found for item {item.material_id}."
                        )
                except ValueError:  # For int conversion
                    logger.warning(
                        f"Course code '{code_str}' is not a valid integer for item {item.material_id}."
                    )
                except Exception as e:
                    logger.warning(
                        f"Error linking course {code_str} to item {item.material_id}: {e}"
                    )
        logger.info(
            f"Attempted linking for {course_codes_processed} course codes. Successfully added/confirmed {links_added} links."
        )
    except Exception as e:
        logger.error(f"General error in link_courses_to_copyright_items: {e}")
        logger.debug(traceback.format_exc())
    finally:
        await Tortoise.close_connections()


async def update_duplicate_status() -> None:
    """
    Updates the `is_duplicate` and `replacement_id` fields on CopyrightItem records.
    It iterates through items, checks linked PDF records for `replace_with` pointers,
    and updates the copyright item accordingly.
    """
    await init_tortoise()
    try:
        items = await CopyrightItem.all()
        duplicates_updated: int = 0
        for item in items:
            original_is_duplicate = item.is_duplicate
            original_replacement_id = item.replacement_id

            item.is_duplicate = False  # Default to False
            item.replacement_id = None

            pdf_record = await PDF.get_or_none(
                material_id=item.material_id
            )  # Assumes one-to-one or primary PDF
            if (
                pdf_record and pdf_record.replace_with_id
            ):  # Check replace_with_id which is the FK field
                # The related PDF this one should be replaced with
                replacing_pdf = await PDF.get_or_none(
                    material_id=pdf_record.replace_with_id
                )
                if replacing_pdf:
                    item.is_duplicate = True
                    item.replacement_id = (
                        replacing_pdf.material_id
                    )  # Store the material_id of the PDF it's replaced by

            if (
                item.is_duplicate != original_is_duplicate
                or item.replacement_id != original_replacement_id
            ):
                await item.save(
                    update_fields=["is_duplicate", "replacement_id", "modified_at"]
                )
                duplicates_updated += 1
        logger.info(f"Updated duplicate status for {duplicates_updated} items.")
    except Exception as e:
        logger.error(f"General error in update_duplicate_status: {e}")
        logger.debug(traceback.format_exc())
    finally:
        await Tortoise.close_connections()


async def update_copyright_relations() -> None:
    """
    Orchestrates updates to various relationships for CopyrightItems.
    This includes duplicate status, LLM classification links, and course links.
    """
    logger.info("Starting update of copyright item relations.")
    await init_tortoise()  # Ensure DB is ready
    try:
        await update_duplicate_status()
        await link_llm_classifications_to_copyright_items()  # Assumes this function now closes its own connection or is fine with one here
        await link_courses_to_copyright_items()  # Same assumption
        logger.info("Successfully updated copyright item relations.")
    except Exception as e:
        logger.error(f"Error during update_copyright_relations: {e}")
        logger.debug(traceback.format_exc())
    finally:
        await (
            Tortoise.close_connections()
        )  # Ensure connection is closed at the end of the orchestration


class DataSource(StrEnum):
    """
    Enum for the source of the data being processed.
    This helps determine the update strategy for copyright items in the database.
    """

    RAW_QLIK_DATA = "raw_qlik_data"  # Data directly from Qlik export, typically authoritative for some fields.
    WEEKLY_SHEET = "weekly_sheet"  # Data from user input in weekly sheets.
    OVERVIEW_SHEET = "overview_sheet"  # Data from user input in overview sheets.
    EA_SCRIPT = (
        "ea_script"  # Data generated or modified by the Easy Access script itself.
    )
    WEB_DASHBOARD = "dashboard"  # Data from user input via the web dashboard.


async def update_copyright_items(
    data: pl.DataFrame | list[dict[str, Any]],  # Data can be DataFrame or list of dicts
    # source: DataSource = DataSource.EA_SCRIPT, # TODO: Re-introduce when refactoring based on source
    overwrite: bool = False,  # If true, overwrites fields regardless of comparison logic (but still selective)
    user_info: dict[str, Any] | None = None,  # For logging who made the change
    update_relations: bool = True,  # Whether to update related models after item updates
) -> None:
    """
    Updates CopyrightItem records in the database based on provided data.
    New items are created if they don't exist. Existing items are updated
    based on a comparison heuristic or overwritten if specified.
    Logs changes to the ItemUpdate table.

    Args:
        data: Data to update from, either a Polars DataFrame or a list of dictionaries.
        overwrite: If True, new values for specified fields will overwrite existing ones
                   without complex comparison. Defaults to False.
        user_info: Optional dictionary containing user information (e.g., email) for audit logging.
        update_relations: If True (default), calls `update_copyright_relations` after processing items.

    Note:
        The `source` parameter from the original TODO is currently not implemented. The logic
        primarily follows the `overwrite` flag and field-specific heuristics. A future refactor
        could make update strategies more dependent on the `DataSource`.
    """

    current_user_email: str = "script@ea.tool"  # Default user
    if user_info and isinstance(user_info.get("email"), str):
        current_user_email = user_info["email"]

    def _log_change(
        item_changes: dict[str, Any],
        field: str,
        new_value: Any,
        old_value: Any,
        reason: str,
    ) -> dict[str, Any]:
        """Helper to log a field change and update the database item object."""
        logger.debug(
            f"Item {item_changes.get('material_id')}: [{reason}] Field '{field}': '{old_value}' -> '{new_value}'"
        )
        item_changes[field] = {"old": str(old_value), "new": str(new_value)}
        # setattr(db_item, field, new_value) # This should be done on the actual db_item Tortoise object
        return item_changes

    def _compare_and_apply_field_changes(
        new_item_data: dict[str, Any],
        db_item_obj: CopyrightItem,
        fields_to_compare: dict[str, list[Any] | None],
        item_change_log: dict[str, Any],
    ) -> tuple[
        dict[str, Any], bool
    ]:  # Returns updated change log and a flag if db_item_obj was modified
        """Compares fields based on defined heuristics and applies changes to db_item_obj."""
        item_modified_flag = False
        for field, priority_order in fields_to_compare.items():
            new_value = new_item_data.get(field)
            old_value = getattr(db_item_obj, field, None)

            # Type normalization/casting before comparison (example for datetime)
            if isinstance(old_value, datetime) and isinstance(new_value, str):
                try:
                    new_value = datetime.strptime(
                        new_value.split(" ")[0], "%Y-%m-%d"
                    ).replace(tzinfo=UTC)
                except ValueError:
                    new_value = None  # Or handle error
                if old_value:
                    old_value = old_value.replace(
                        tzinfo=UTC
                    )  # Ensure timezone consistency

            if isinstance(old_value, Enum):
                old_value = old_value.value
            if isinstance(new_value, Enum):
                new_value = new_value.value  # Ensure new_value from dict is also compared as value if it's an Enum

            if new_value is None:
                continue  # Skip if new value is None (don't overwrite with None unless intended)
            if new_value == old_value:
                continue

            if (
                old_value is None or old_value == ""
            ):  # DB value is empty/None, new value is not
                item_change_log = _log_change(
                    item_change_log,
                    field,
                    new_value,
                    old_value,
                    "DB value empty, using new value",
                )
                setattr(db_item_obj, field, new_value)
                item_modified_flag = True
            elif priority_order:  # List-based priority
                new_rank = (
                    priority_order.index(new_value)
                    if new_value in priority_order
                    else float("inf")
                )
                old_rank = (
                    priority_order.index(old_value)
                    if old_value in priority_order
                    else float("inf")
                )
                if new_rank < old_rank:  # Lower index means higher priority
                    item_change_log = _log_change(
                        item_change_log,
                        field,
                        new_value,
                        old_value,
                        "New value has higher priority rank",
                    )
                    setattr(db_item_obj, field, new_value)
                    item_modified_flag = True
            elif isinstance(new_value, str) and isinstance(
                old_value, str
            ):  # String length heuristic (e.g. for remarks)
                if len(new_value.strip()) > len(old_value.strip()):
                    item_change_log = _log_change(
                        item_change_log,
                        field,
                        new_value,
                        old_value,
                        "New string value longer",
                    )
                    setattr(db_item_obj, field, new_value)
                    item_modified_flag = True
            # Add other type-specific comparisons if needed (e.g., dates, numbers if None means prefer higher/later)
            # Default for other types if no priority: new value overwrites if different (already handled by initial check)
            elif (
                new_value != old_value
            ):  # Generic overwrite if different and no specific rule matched
                item_change_log = _log_change(
                    item_change_log,
                    field,
                    new_value,
                    old_value,
                    "Generic update, new value different",
                )
                setattr(db_item_obj, field, new_value)
                item_modified_flag = True

        return item_change_log, item_modified_flag

    await init_tortoise()

    # Define field categories for update logic
    # Added by script, typically not user-editable directly in sheets (value is None for simple overwrite if different)
    added_fields: dict[str, list[Any] | None] = {
        "workflow_status": [
            ws.value for ws in WorkflowStatus
        ],  # Highest priority first
        "retrieved_from_copyright_on": None,  # Take latest if applicable, or just update
        "possible_fine": None,  # Overwrite if different
        "infringement": [i.value for i in Infringement],  # Highest priority first
    }
    # User-changeable fields, with priority lists or None for simple overwrite logic
    changeable_fields: dict[str, list[Any] | None] = {
        "manual_classification": [
            c.value for c in Classification
        ],  # Highest priority first
        "manual_identifier": None,
        "remarks": None,
        "scope": None,
        # Fields also updated by raw Qlik data, Qlik is source of truth for these
        "status": [s.value for s in Status],
        "last_change": None,
        "pages_x_students": None,
        "count_students_registered": None,
    }

    logger.info(
        f"Received {len(data)} items for update process. Overwrite: {overwrite}"
    )

    items_to_update_from_input: list[dict[str, Any]]
    new_item_dicts: list[dict[str, Any]] = []

    if isinstance(data, pl.DataFrame):
        # Standardize incoming DataFrame (column names, basic cleaning)
        # standardize_dataframe now returns strings, specific type conversions happen in copyright_item_from_dict or here
        standardized_df = standardize_dataframe(
            data.clone()
        )  # Clone to avoid modifying original

        all_input_ids_str = (
            standardized_df.select(pl.col("material_id").cast(pl.Utf8))
            .to_series()
            .to_list()
        )
        existing_db_items = await CopyrightItem.filter(
            material_id__in=all_input_ids_str
        ).all()
        existing_db_ids = {str(item.material_id) for item in existing_db_items}

        items_to_update_from_input = standardized_df.filter(
            pl.col("material_id").is_in(existing_db_ids)
        ).to_dicts()
        new_item_dicts = standardized_df.filter(
            ~pl.col("material_id").is_in(existing_db_ids)
        ).to_dicts()
    elif isinstance(data, list):  # Assuming list of dicts
        # This path requires more careful handling to separate new vs existing if not pre-filtered
        all_input_ids = [
            str(item.get("material_id")) for item in data if item.get("material_id")
        ]
        existing_db_items = await CopyrightItem.filter(
            material_id__in=all_input_ids
        ).all()
        existing_db_ids = {str(item.material_id) for item in existing_db_items}

        items_to_update_from_input = [
            item for item in data if str(item.get("material_id")) in existing_db_ids
        ]
        new_item_dicts = [
            item for item in data if str(item.get("material_id")) not in existing_db_ids
        ]
    else:
        logger.error(f"Unsupported data type for update_copyright_items: {type(data)}")
        return

    logger.info(f"# of new items to create: {len(new_item_dicts)}")
    logger.info(
        f"# of existing items to check for updates: {len(items_to_update_from_input)}"
    )

    created_orm_items: list[CopyrightItem] = []
    if new_item_dicts:
        for item_dict in new_item_dicts:
            orm_item = await copyright_item_from_dict(
                item_dict
            )  # This handles type conversion
            if orm_item:
                created_orm_items.append(orm_item)
        if created_orm_items:
            try:
                await CopyrightItem.bulk_create(created_orm_items)
                logger.info(
                    f"Bulk created {len(created_orm_items)} new copyright items."
                )
            except Exception as e:
                logger.warning(
                    f"Error during bulk_create of new items: {e}. Attempting individual save."
                )
                for orm_item in created_orm_items:
                    try:
                        await orm_item.save()
                    except Exception as ie:
                        logger.error(
                            f"Error saving new item {orm_item.material_id}: {ie}"
                        )

    db_items_to_bulk_update: list[CopyrightItem] = []
    item_update_logs: list[ItemUpdate] = []

    # Map existing DB items by material_id for quick lookup
    existing_db_items_map = {str(item.material_id): item for item in existing_db_items}

    for new_item_data_dict in items_to_update_from_input:
        mat_id_str = str(new_item_data_dict.get("material_id"))
        db_item_instance = existing_db_items_map.get(mat_id_str)

        if (
            not db_item_instance
        ):  # Should not happen due to pre-filtering, but as a safeguard
            logger.warning(
                f"Item with material_id {mat_id_str} not found in DB map for update. Skipping."
            )
            continue

        item_change_log_details: dict[str, Any] = {
            "material_id": db_item_instance.material_id,  # Use actual int ID from DB object
            "update_time": datetime.now(UTC).strftime("%Y-%m-%d %H:%M:%S"),
            "modified_by": current_user_email,
        }
        item_was_modified_in_db = False

        if overwrite:
            logger.debug(f"Overwrite mode for item {mat_id_str}.")
            for field_key in list(changeable_fields.keys()) + list(added_fields.keys()):
                if field_key in new_item_data_dict:
                    new_value_from_input = new_item_data_dict[field_key]
                    old_db_value = getattr(db_item_instance, field_key, None)

                    # Basic type normalization for comparison, can be expanded
                    if isinstance(old_db_value, Enum):
                        old_db_value = old_db_value.value
                    if isinstance(new_value_from_input, Enum):
                        new_value_from_input = new_value_from_input.value
                    if isinstance(old_db_value, datetime) and isinstance(
                        new_value_from_input, str
                    ):
                        try:
                            new_value_from_input = datetime.strptime(
                                new_value_from_input.split(" ")[0], "%Y-%m-%d"
                            )
                        except ValueError:
                            pass  # Keep as string if not parsable here, model field will handle

                    if str(new_value_from_input) != str(
                        old_db_value
                    ):  # Simple string comparison after basic normalization
                        item_change_log_details = _log_change(
                            item_change_log_details,
                            field_key,
                            new_value_from_input,
                            old_db_value,
                            "[overwrite]",
                        )
                        setattr(
                            db_item_instance, field_key, new_value_from_input
                        )  # Apply change to ORM object
                        item_was_modified_in_db = True
        else:  # Heuristic comparison
            item_change_log_details, field_changed1 = _compare_and_apply_field_changes(
                new_item_data_dict,
                db_item_instance,
                added_fields,
                item_change_log_details,
            )
            item_change_log_details, field_changed2 = _compare_and_apply_field_changes(
                new_item_data_dict,
                db_item_instance,
                changeable_fields,
                item_change_log_details,
            )
            if field_changed1 or field_changed2:
                item_was_modified_in_db = True

        if item_was_modified_in_db:
            db_item_instance.modified_at = datetime.now(UTC)
            db_items_to_bulk_update.append(db_item_instance)
            item_update_logs.append(
                ItemUpdate(
                    change_details=item_change_log_details,
                    material_id=db_item_instance.material_id,
                )
            )  # type: ignore

    if db_items_to_bulk_update:
        # Determine all unique fields that were changed across all items for bulk_update
        all_changed_field_keys: set[str] = set()
        for log_entry_details in item_update_logs:
            for key in (
                log_entry_details.change_details.keys()
            ):  # Iterate over keys in change_details dict
                if key not in ["material_id", "update_time", "modified_by"]:
                    all_changed_field_keys.add(key)

        update_fields_list = list(all_changed_field_keys)
        if (
            "modified_at" not in update_fields_list
        ):  # Always include modified_at if any changes occurred
            update_fields_list.append("modified_at")

        if not update_fields_list or (
            len(update_fields_list) == 1
            and update_fields_list[0] == "modified_at"
            and not all_changed_field_keys
        ):
            logger.info(
                "No substantive fields to update in bulk apart from modified_at, or list is empty."
            )
        else:
            logger.info(
                f"Bulk updating {len(db_items_to_bulk_update)} items for fields: {update_fields_list}."
            )
            await CopyrightItem.bulk_update(
                db_items_to_bulk_update, fields=update_fields_list
            )

        if item_update_logs:
            await ItemUpdate.bulk_create(item_update_logs)
            logger.info(f"Created {len(item_update_logs)} ItemUpdate log entries.")
            # Link ItemUpdates to CopyrightItems (M2M)
            for (
                db_item_updated
            ) in db_items_to_bulk_update:  # Iterate over actual updated ORM objects
                # Fetch the most recent ItemUpdate for this material_id (the one just created)
                latest_log_entry = (
                    await ItemUpdate.filter(material_id=db_item_updated.material_id)
                    .order_by("-created_at")
                    .first()
                )
                if latest_log_entry:
                    await db_item_updated.changes.add(latest_log_entry)

    if (created_orm_items or db_items_to_bulk_update) and update_relations:
        logger.info("Updating relations for CopyrightItems after updates/creations.")
        await (
            update_copyright_relations()
        )  # This function should also handle its own Tortoise connection

    logger.info("Finished update_copyright_items process.")
    await Tortoise.close_connections()  # Close connection at the very end
