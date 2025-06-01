"""
functions to update existing data in the database
"""

from dataclasses import dataclass
from datetime import date, datetime
from enum import StrEnum
from itertools import batched
from typing import Any

import polars as pl
from loguru import logger
from rich import print
from tortoise import Tortoise

from easy_access.db.base import copyright_item_from_dict, init, standardize_dataframe
from easy_access.db.models import (
    PDF,
    Classification,
    CopyrightItem,
    Course,
    ItemUpdate,
    LLMClassification,
    Status,
    WorkflowStatus,
    str_to_classification,
)
from easy_access.utils import cool, determine_course_code, info, warn


async def link_llm_classifications_to_copyright_items() -> None:
    """
    Link llm classifications to copyright items
    """
    # get all LLM classifications without a corresponding CopyrightItem

    classifications_to_link = await LLMClassification().filter(item__isnull=True)

    for classification in classifications_to_link:
        try:
            material_id = classification.used_material_id
            item = await CopyrightItem.get_or_none(material_id=material_id)
            if item:
                item.llm_classification = classification
                await item.save()
        except Exception as e:
            warn(
                f"Error while trying to get llm classification for item {item.material_id}: {e}"
            )
            continue

    info(
        f"Done linking {len(classifications_to_link)} LLM classifications to CopyrightItems."
    )


async def link_courses_to_copyright_items() -> None:
    items_w_prefetch = await CopyrightItem.all()
    info(f"got {len(items_w_prefetch)} items from db")

    # for each of the items, extract the course code (using determine_course_code)
    # then match with existing course item in db
    # if missing, add to list to retrieve later

    links_added = 0
    course_codes_found = 0
    for item in items_w_prefetch:
        course_codes = determine_course_code(item.course_code, item.course_name)
        if not course_codes or len(course_codes) == 0:
            warn(
                f"Could not determine course code for item {item.material_id} with input course code {item.course_code} and course name {item.course_name}."
            )
        course_codes = list(course_codes)

        for course_code in course_codes:
            if not course_code:
                continue
            try:
                cursuscode = int(course_code)
                course_codes_found += 1

                course = await Course.get_or_none(cursuscode=cursuscode)
                if course:
                    await item.courses.add(course)
                    links_added += 1
            except Exception as e:
                warn(
                    f"Error while trying to get course {course_code} for item {item.material_id}: {e}"
                )
    cool(f"Added {links_added} links to {course_codes_found} found coursecodes.")


async def update_duplicate_status() -> None:
    """
    For each item, retrieve the PDF
    if the PDF has value in 'replace_with', set the item's is_duplicate status to True
    grab the material_id from the replace_with field and store it in the 'replacement_id' field
    """

    items = await CopyrightItem.all()
    duplicates = 0
    for item in items:
        item.is_duplicate = False
        item.replacement_id = None
        mat_id = item.material_id
        pdf = await PDF.get_or_none(material_id=mat_id)
        if pdf:
            replaced_item = await pdf.replace_with
            if replaced_item:
                item.is_duplicate = True
                item.replacement_id = replaced_item.material_id
                duplicates += 1
        await item.save(update_fields=["is_duplicate", "replacement_id"])
    cool(f"Updated {duplicates} duplicate statuses.")


async def update_copyright_relations() -> None:
    """
    Go through the copyright items in the db
    use the values of the item fields to find links to other tables.
    """
    await init()
    # await update_duplicate_status()

    # await link_llm_classifications_to_copyright_items()
    await link_courses_to_copyright_items()
    await Tortoise.close_connections()


class DataSource(StrEnum):
    """Enum for the source of the data.

    This is used to determine how to handle the data when updating the database.
    """

    RAW_QLIK_DATA = "raw_qlik_data"
    WEEKLY_SHEET = "weekly_sheet"
    OVERVIEW_SHEET = "overview_sheet"
    EA_SCRIPT = "ea_script"
    WEB_DASHBOARD = "dashboard"


@dataclass
class ErrorItem:
    """Class to represent an item with an error."""

    material_id: int
    error: str
    full_data: dict
    source: DataSource

    #  two erroritems are the same if they have the same material_id, source, and error
    def __eq__(self, other):
        if not isinstance(other, ErrorItem):
            return False
        return (
            self.material_id == other.material_id
            and self.source == other.source
            and self.error == other.error
        )


async def update_copyright_items(
    data: pl.DataFrame | list[dict],
    source: DataSource,  # Add this parameter
    update_relations: bool = True,
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
        user_info.get("email") if user_info.get("email") else "cip-admin@utwente.nl"
    )
    await init()

    def change(
        changes: dict, field: str, new_value: Any, old_value: Any, reason: str
    ) -> dict:
        logger.debug(f"[{reason}] [{field}] {old_value} --> {new_value}")
        changes[field] = {"old": str(old_value), "new": str(new_value)}
        setattr(db_item, field, new_value)
        return changes

    # Standardize input data
    if isinstance(data, pl.DataFrame):
        data = standardize_dataframe(data)
        items_to_process = data.to_dicts()
    elif isinstance(data, list):
        items_to_process = data
    else:
        warn(
            "Invalid data type passed to update_copyright_items. Expected DataFrame or list."
        )
        return

    if not items_to_process:
        info("No items provided for update.")
        return

    info(
        f"Received {len(items_to_process)} items from source '{source.value}' for potential update."
    )

    errors: list[ErrorItem] = []  # stores items with errors
    # Separate new vs existing items based on DB check
    existing_mat_ids = await CopyrightItem.all().values_list("material_id", flat=True)
    existing_mat_ids = set(existing_mat_ids)

    new_items_dicts = [
        item
        for item in items_to_process
        if int(item.get("material_id", 0)) not in existing_mat_ids
    ]
    update_items_dicts = [
        item
        for item in items_to_process
        if int(item.get("material_id", 0)) in existing_mat_ids
    ]

    # New items can only be created if source == RAW_QLIK_DATA
    new_objects = []
    if new_items_dicts:
        if source == DataSource.RAW_QLIK_DATA:
            info(f"Processing {len(new_items_dicts)} new items from Qlik data.")
            new_objects = [
                await copyright_item_from_dict(item) for item in new_items_dicts
            ]
            new_objects = [item for item in new_objects if item]
            if new_objects:
                try:
                    await CopyrightItem.bulk_create(objects=new_objects)
                    cool(f"Created {len(new_objects)} new copyright items in db.")
                except Exception as e:
                    warn(
                        f"error while trying to bulk save items. Error: {e}. Trying one-by-one."
                    )
                    for item in new_objects:
                        try:
                            await item.save()
                        except Exception as e:
                            logger.error(e)
                            warn(
                                f"error {e} while trying to save item {item} with material_id {item.material_id}. Skipping for now."
                            )
        else:
            warn(
                f"Item {len(new_items_dicts)} items found in source '{source.value}' that do not exist in the database. Marking as errors."
            )
            for item in new_items_dicts:
                errors.append(
                    ErrorItem(
                        material_id=item.get("material_id"),
                        error="Item in sheet but not found in DB",
                        full_data=item,
                        source=source,
                    )
                )

    # Helper update logic for each source
    def _apply_qlik_update(
        new_item: dict[str, str], db_item: CopyrightItem, changes: dict[str, str]
    ):
        # Overwrite these fields always
        for field in [
            "status",
            "last_change",
            "count_students_registered",
            "pages_x_students",
        ]:
            new_value = new_item.get(field)
            old_value = getattr(db_item, field)
            if field == "status":
                # Convert to enum for comparison
                new_value = Status(new_value) if new_value else None
                old_value = db_item.status
            if field == "last_change":
                # Convert to date for comparison
                new_value = date.fromisoformat(new_value)
                old_value = db_item.last_change
            if field == "count_students_registered" or field == "pages_x_students":
                # Convert to int for comparison
                new_value = int(new_value) if new_value else None
                old_value = (
                    db_item.count_students_registered
                    if field == "count_students_registered"
                    else db_item.pages_x_students
                )
            if new_value is not None and new_value != old_value:
                changes = change(changes, field, new_value, old_value, "qlik overwrite")
        # Compare but do not update, just log inconsistencies
        for field in [
            "manual_classification",
            "manual_identifier",
            "scope",
            "remarks",
            "auditor",
        ]:
            new_value = new_item.get(field)
            old_value = getattr(db_item, field)
            if new_value and new_value != old_value:
                warn(
                    f"Inconsistency found for {field} in material_id {new_item.get('material_id')}: DB='{old_value}' vs Qlik='{new_value}' (not updating DB)"
                )
                errors.append(
                    ErrorItem(
                        material_id=new_item.get("material_id"),
                        error=f"Inconsistency found for {field}: DB='{old_value}' vs Qlik='{new_value}'",
                        full_data=new_item,
                        source=source,
                    )
                )
        return changes, db_item

    def _apply_overview_update(
        new_item: dict[str, str], db_item: CopyrightItem, changes: dict[str, str]
    ):
        for field in ["workflow_status", "manual_classification", "remarks"]:
            new_value = new_item.get(field)
            old_value = getattr(db_item, field)
            if field == "workflow_status":
                new_value = WorkflowStatus(new_value) if new_value else None
                old_value = db_item.workflow_status
            if new_value is not None and new_value != old_value:
                changes = change(
                    changes, field, new_value, old_value, "overview overwrite"
                )
        return changes, db_item

    def _apply_web_update(
        new_item: dict[str, str], db_item: CopyrightItem, changes: dict[str, str]
    ):
        for field in ["workflow_status", "manual_classification", "remarks"]:
            new_value = new_item.get(field)
            if field == "workflow_status":
                # Convert to enum for comparison
                new_value = WorkflowStatus(new_value) if new_value else None
            if field == "manual_classification":
                # Convert to enum for comparison
                new_value = str_to_classification(new_value) if new_value else None
            old_value = getattr(db_item, field)
            if new_value is not None and new_value != old_value:
                changes = change(changes, field, new_value, old_value, "web overwrite")
        return changes, db_item

    def _apply_weekly_update(
        new_item: dict[str, str], db_item: CopyrightItem, changes: dict[str, str]
    ):
        # Heuristic: recency, status priority, classification priority, remarks merge
        # 1. If old_value is empty, use new_value
        # 2. If both exist, try to compare change dates (not implemented here, placeholder)
        # 3. Else, compare workflow_status by priority
        # 4. Else, compare manual_classification by priority
        # 5. Else, merge remarks
        # Priority lists
        workflow_priority = [
            WorkflowStatus.Done,
            WorkflowStatus.InProgress,
            WorkflowStatus.ToDo,
        ]
        classification_priority = [
            Classification.OPEN_ACCESS,
            Classification.KORTE_OVERNAME,
            Classification.MIDDELLANGE_OVERNAME,
            Classification.LANGE_OVERNAME,
            Classification.EIGEN_MATERIAAL_POWERPOINT,
            Classification.EIGEN_MATERIAAL_TITELINDICATIE,
            Classification.EIGEN_MATERIAAL_OVERIG,
            Classification.EIGEN_MATERIAAL,
            Classification.ANDERS,
            Classification.ONBEKEND,
            Classification.LICENTIE_BESCHIKBAAR,
            Classification.NIET_GEANALYSEERD,
            Classification.IN_ONDERZOEK,
            Classification.VERWIJDERVERZOEK_VERSTUURD,
        ]
        # Workflow status
        ws_new: str = new_item.get("workflow_status")
        # Convert to enum for comparison
        ws_new = WorkflowStatus(ws_new) if ws_new else None
        ws_old: WorkflowStatus | None = db_item.workflow_status
        if not ws_old and ws_new:
            changes = change(
                changes, "workflow_status", ws_new, ws_old, "weekly: old empty"
            )
        elif ws_new and ws_old and ws_new != ws_old:
            try:
                new_rank = (
                    workflow_priority.index(ws_new)
                    if ws_new in workflow_priority
                    else 99
                )
                old_rank = (
                    workflow_priority.index(ws_old)
                    if ws_old in workflow_priority
                    else 99
                )
                if new_rank < old_rank:
                    changes = change(
                        changes,
                        "workflow_status",
                        ws_new,
                        ws_old,
                        "weekly: higher priority",
                    )
            except Exception:
                pass
        # Manual classification
        mc_new = new_item.get("manual_classification")
        # Convert to enum for comparison
        mc_new = str_to_classification(mc_new) if mc_new else None
        mc_old = str_to_classification(classification_str=db_item.manual_classification)
        if not mc_old and mc_new:
            changes = change(
                changes, "manual_classification", mc_new, mc_old, "weekly: old empty"
            )
        elif mc_new and mc_old and mc_new != mc_old:
            try:
                new_rank = (
                    classification_priority.index(mc_new)
                    if mc_new in classification_priority
                    else 99
                )
                old_rank = (
                    classification_priority.index(mc_old)
                    if mc_old in classification_priority
                    else 99
                )
                if new_rank < old_rank:
                    changes = change(
                        changes,
                        "manual_classification",
                        mc_new,
                        mc_old,
                        "weekly: higher priority",
                    )
            except Exception:
                pass
        # Remarks: merge if different
        remarks_new = new_item.get("remarks")
        remarks_old = db_item.remarks
        if remarks_new and remarks_old and remarks_new != remarks_old:
            merged = remarks_old
            if remarks_new not in remarks_old:
                merged = remarks_old + ", " + remarks_new
            changes = change(
                changes, "remarks", merged, remarks_old, "weekly: merge remarks"
            )
        elif remarks_new and not remarks_old:
            changes = change(
                changes, "remarks", remarks_new, remarks_old, "weekly: old empty"
            )
        return changes, db_item

    # Main update loop for existing items
    info(f"Updating {len(update_items_dicts)} existing items.")
    changelist: list[CopyrightItem] = []
    updates: dict[str, dict[str, str]] = {}
    for new_item in update_items_dicts:
        try:
            db_item = await CopyrightItem.get(material_id=new_item.get("material_id"))
            changes = {
                "material_id": new_item.get("material_id"),
                "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            if source == DataSource.RAW_QLIK_DATA:
                changes, db_item = _apply_qlik_update(new_item, db_item, changes)
            elif source == DataSource.OVERVIEW_SHEET:
                changes, db_item = _apply_overview_update(new_item, db_item, changes)
            elif source == DataSource.WEEKLY_SHEET:
                changes, db_item = _apply_weekly_update(new_item, db_item, changes)
            elif source == DataSource.WEB_DASHBOARD:
                changes, db_item = _apply_web_update(new_item, db_item, changes)
            elif source == DataSource.EA_SCRIPT:
                # For now, treat as overview update (can be customized)
                changes, db_item = _apply_overview_update(new_item, db_item, changes)
            # If any changes were made (more than just material_id, update_time)
            if len(changes) > 2:
                changelist.append(db_item)
                updates[new_item.get("material_id")] = changes
        except Exception as e:
            warn(f"Error updating item {new_item.get('material_id')}: {e}")
            errors.append(
                ErrorItem(
                    material_id=new_item.get("material_id"),
                    error=f"{e}",
                    full_data=new_item,
                    source=source,
                )
            )

    if len(changelist) > 0:
        print(f"changelist: {len(changelist)} items")

        # get all values from 'updates'
        # then get list of all distinct keys from all those dicts
        # then drop keys 'material_id' and 'update_time'
        # then add all those keys to the fields to update
        all_keys = {key for item in updates.values() for key in item}
        all_keys.discard("material_id")
        all_keys.discard("update_time")
        changed_fields = list(all_keys)
        changed_fields.append("modified_at")
        info(f"Updating {len(changelist)} items in db for fields {changed_fields}.")
        for batch in batched(changelist, 100):
            await CopyrightItem.bulk_update(batch, fields=changed_fields)
        if user_info:
            [changes.update({"modified_by": cur_user}) for changes in updates.values()]
            print(f"items modified by {cur_user}")

        info(f"Updating {len(updates)} changelog items in db.")
        for batch in batched(updates.values(), 100):
            await ItemUpdate.bulk_create(
                [
                    ItemUpdate(
                        change_details=change, material_id=change.get("material_id")
                    )
                    for change in batch
                ]
            )

        # now add each ItemUpdate as a m2m relation to the corresponding CopyrightItem
        for mat_id in updates:
            item = await CopyrightItem.get(material_id=mat_id)
            update = (
                await ItemUpdate.filter(material_id=mat_id)
                .order_by("-created_at")
                .first()
            )
            await item.changes.add(update)

    if new_objects:
        cool("Updating relations for CopyrightItems.")
        await update_copyright_relations()

    cool("Done updating CopyrightItems!")
    await Tortoise.close_connections()
