"""
functions to update existing data in the database
"""

from datetime import datetime, timezone
from enum import Enum
from typing import Any

import polars as pl
from loguru import logger
from tortoise import Tortoise

from easy_access.db.base import copyright_item_from_dict, init, standardize_dataframe
from easy_access.db.models import (
    PDF,
    Classification,
    CopyrightItem,
    Course,
    Infringement,
    ItemUpdate,
    LLMClassification,
    Status,
    WorkflowStatus,
)
from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import cool, determine_course_code, info, warn


async def link_llm_classifications_to_copyright_items() -> None:
    """
    Link llm classifications to copyright items
    """
    # get all LLM classifications without a corresponding CopyrightItem

    classifications_to_link = await LLMClassification().filter(item__isnull=True)

    missing_classifications = []
    # load dedupe info from .replace files
    replace_files = {
        f.name.split("_")[0]: f.name.split("_")[1]
        for f in SETTINGS.dirs[DirSetting.CLASSIFICATIONS].files
        if f.name.endswith(".replace")
    }

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
    await update_duplicate_status()

    await link_llm_classifications_to_copyright_items()
    await link_courses_to_copyright_items()
    await Tortoise.close_connections()


async def update_copyright_items(data: pl.DataFrame | list[dict]) -> None:
    """
    Update the db with copyrightitems from the dataframe (or pre-filtered list of dicts from a df).
    Adds new if they don't exist, or updates if they do.
    See `compare_items` and the dicts added_fields, changeable_fields, core_fields for details on how the comparison is done.

    Once done, and if any updates were made, will call `update_copyright_relations` to update the m2m relations.
    """

    def compare_fields(
        new_item: dict, db_item: CopyrightItem, fielddict: dict, changes: dict
    ) -> tuple[dict, CopyrightItem]:
        def change(
            changes: dict, field: str, new_value: Any, old_value: Any, reason: str
        ) -> dict:
            logger.debug(f"[{reason}] [{field}] {old_value} --> {new_value}")
            changes[field] = {"old": str(old_value), "new": str(new_value)}
            setattr(db_item, field, new_value)
            return changes

        if not changes:
            changes = {
                "material_id": new_item.get("material_id"),
                "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

        for field, ordering in fielddict.items():
            new_value = new_item.get(field)
            old_value = getattr(db_item, field)

            try:
                if isinstance(old_value, datetime):
                    new_value = None
                    try:
                        new_value = (
                            datetime.strptime(new_value, "%Y-%m-%d %H:%M:%S%z").replace(
                                tzinfo=timezone.utc
                            )
                            if new_value
                            else None
                        )
                    except Exception:
                        try:
                            new_value = (
                                datetime.strptime(
                                    new_value, "%Y-%m-%d %H:%M:%S"
                                ).replace(tzinfo=timezone.utc)
                                if new_value
                                else None
                            )
                        except Exception:
                            try:
                                new_value = (
                                    datetime.strptime(new_value, "%Y-%m-%d").replace(
                                        tzinfo=timezone.utc
                                    )
                                    if new_value
                                    else None
                                )
                            except Exception:
                                ...

                    old_value = old_value.replace(tzinfo=timezone.utc)
                if isinstance(old_value, Enum):
                    old_value = old_value.value
                if isinstance(old_value, float):
                    new_value = round(float(new_value), 2) if new_value else None
                    old_value = round(old_value, 2)
                if isinstance(old_value, int):
                    new_value = int(new_value) if new_value else None
            except Exception as e:
                logger.debug(
                    f"error {e} while typecasting data for field comparison of {field}"
                )
                continue
            if new_value:
                if old_value == new_value:
                    continue
                if not old_value:
                    changes = change(
                        changes, field, new_value, old_value, "no old value"
                    )
                elif isinstance(ordering, list):
                    new_rank = 20
                    old_rank = 20
                    if new_value in ordering:
                        new_rank = ordering.index(new_value)
                    if old_value in ordering:
                        old_rank = ordering.index(old_value)
                    if new_rank < old_rank:
                        changes = change(
                            changes, field, new_value, old_value, "new rank < old rank"
                        )
                else:
                    if isinstance(new_value, str) and isinstance(old_value, str):
                        new_value = new_value.strip()
                        old_value = old_value.strip()
                        if len(new_value) > len(old_value):
                            changes = change(
                                changes,
                                field,
                                new_value,
                                old_value,
                                "new len > old len",
                            )
                    else:
                        if type(new_value) is type(old_value):
                            if new_value > old_value:
                                changes = change(
                                    changes, field, new_value, old_value, "new > old"
                                )
                        else:
                            logger.debug(
                                f"[incomparable types] [{field}] {type(new_value)=}, {type(old_value)=}"
                            )
            else:
                # new value is None
                # do nothing?
                pass

        return changes, db_item

    await init()
    # Fields added by script.
    # dict with field name as key,
    # value are the ordered possible values:in case of conflict, take the earliest value
    # for values that are None, sort and take the highest/latest/... or implement some other logic
    added_fields = {
        "workflow_status": [
            WorkflowStatus.Done.value,
            WorkflowStatus.InProgress.value,
            WorkflowStatus.ToDo.value,
        ],
        "retrieved_from_copyright_on": None,
        "possible_fine": None,
        "infringement": [
            Infringement.YES.value,
            Infringement.NO.value,
            Infringement.UNDETERMINED.value,
        ],
    }

    # fields changable by checkers. See above for details
    changeable_fields = {
        "manual_classification": [
            Classification.OPEN_ACCESS.value,
            Classification.KORTE_OVERNAME.value,
            Classification.MIDDELLANGE_OVERNAME.value,
            Classification.LANGE_OVERNAME.value,
            Classification.EIGEN_MATERIAAL_POWERPOINT.value,
            Classification.EIGEN_MATERIAAL_TITELINDICATIE.value,
            Classification.EIGEN_MATERIAAL_OVERIG.value,
            Classification.EIGEN_MATERIAAL.value,
            Classification.ONBEKEND.value,
            Classification.LICENTIE_BESCHIKBAAR.value,
            Classification.NIET_GEANALYSEERD.value,
            Classification.IN_ONDERZOEK.value,
            Classification.VERWIJDERVERZOEK_VERSTUURD.value,
        ],
        "manual_identifier": None,
        "remarks": None,
        "scope": None,
    }

    core_fields = {
        "title": None,
        "classification": None,
        "ml_prediction": None,
        "auditor": None,
        "last_change": None,
        "status": [
            Status.DELETED,
            Status.PUBLISHED,
            Status.UNPUBLISHED,
        ],
        "isbn": None,
        "doi": None,
        "in_collection": None,
        "pagecount": None,
        "wordcount": None,
        "picturecount": None,
        "author": None,
        "publisher": None,
        "reliability": None,
        "pages_x_students": None,
        "count_students_registered": None,
    }
    # standardize df
    # loop over items
    # if item is not in db: add it
    # else compare values in specific fields to determine if we need to update
    info(f"Received {len(data)} raw copyright items as input for an update.")
    new_items = []
    if isinstance(data, pl.DataFrame):
        data = standardize_dataframe(data)
        existing_mat_ids = await CopyrightItem.all().values("material_id")
        existing_mat_ids = {int(m["material_id"]) for m in existing_mat_ids}
        new_items = (
            data.with_columns(pl.col("material_id").cast(int))
            .filter(~pl.col("material_id").is_in(existing_mat_ids))
            .to_dicts()
        )
        update_items = (
            data.with_columns(pl.col("material_id").cast(int))
            .filter(pl.col("material_id").is_in(existing_mat_ids))
            .to_dicts()
        )
    else:
        if isinstance(data, list):
            update_items = data

    info(f"# of new items: {len(new_items)}")
    new_objects = []
    if new_items:
        new_objects = [await copyright_item_from_dict(item) for item in new_items]
        new_objects = [item for item in new_objects if item]
        try:
            await CopyrightItem.bulk_create(objects=new_objects)
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
        cool(f"Created {len(new_objects)} new copyright items in db.")

    info(f"Updating {len(update_items)} existing items.")
    changelist = []
    updates = {}

    for new_item in update_items:
        try:
            db_item = await CopyrightItem.get(material_id=new_item.get("material_id"))
            changes = {}
            changes, db_item = compare_fields(new_item, db_item, added_fields, changes)
            changes, db_item = compare_fields(
                new_item, db_item, changeable_fields, changes
            )

            if new_item.get("last_change"):
                # if new_item has a newer last_change value, we need to update the core CopyRight fields
                item_last_changed = None
                db_last_changed = None
                try:
                    item_last_changed = datetime.strptime(
                        new_item.get("last_change"), "%Y-%m-%d"
                    ).replace(tzinfo=timezone.utc)
                    db_last_changed = (
                        db_item.last_change.replace(tzinfo=timezone.utc)
                        if db_item.last_change
                        else None
                    )
                except Exception:
                    pass

                if isinstance(item_last_changed, datetime) and isinstance(
                    db_last_changed, datetime
                ):
                    if item_last_changed > db_last_changed:
                        # replace the core fields in the db item with new values
                        for field in core_fields:
                            if new_item.get("field"):
                                if new_item.get("field") != getattr(db_item, field):
                                    logger.debug(
                                        f"[core field] Changing field {field} for item {db_item.material_id} from {getattr(db_item, field)} to {new_item.get(field)}"
                                    )
                                    changes[field] = {
                                        "old": getattr(db_item, field),
                                        "new": str(new_item.get(field)),
                                    }
                                    setattr(db_item, field, new_item.get(field))
        except Exception as e:
            warn(f"Could not update item {new_item.get('material_id')}: {e}")
        finally:
            if len(list(changes.keys())) >= 3:
                changes["modified_at"] = datetime.now()
                updates[new_item.get("material_id")] = changes
                changelist.append(db_item)

    if changelist:
        # get all values from 'updates'
        # then get list of all distinct keys from all those dicts
        # then drop keys 'material_id' and 'update_time'
        # then add all those keys to the fields to update
        all_keys = {key for item in updates.values() for key in item.keys()}
        all_keys.discard("material_id")
        all_keys.discard("update_time")
        changed_fields = list(all_keys)
        changed_fields.append("modified_at")
        info(f"Updating {len(changelist)} items in db for fields {changed_fields}.")
        info(f"Updating {len(updates)} changelog items in db.")
        await CopyrightItem.bulk_update(changelist, fields=changed_fields)
        await ItemUpdate.bulk_create(
            [
                ItemUpdate(change_details=changes, material_id=mat_id)
                for mat_id, changes in updates.items()
            ]
        )

        # now add each ItemUpdate as a m2m relation to the corresponding CopyrightItem
        for mat_id in updates.keys():
            item = await CopyrightItem.get(material_id=mat_id)
            update = (
                await ItemUpdate.filter(material_id=mat_id)
                .order_by("-created_at")
                .first()
            )
            await item.changes.add(update)

    if changelist or new_objects:
        cool("Updating relations for all CopyrightItems.")
        await update_copyright_relations()

    cool("Done updating CopyrightItems!")
    await Tortoise.close_connections()
