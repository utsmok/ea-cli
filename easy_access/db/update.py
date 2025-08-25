"""
functions to update existing data in the database
"""

import contextlib
import traceback
from datetime import UTC, datetime
from enum import Enum, StrEnum
from typing import Any

import polars as pl
from loguru import logger
from tortoise import Tortoise

from easy_access.db.base import (
    copyright_item_from_dict,
    ensure_db_inited,
    standardize_dataframe,
)
from easy_access.db.models import (
    PDF,
    Classification,
    CopyrightItem,
    Course,
    Infringement,
    ItemUpdate,
    LLMClassification,
    WorkflowStatus,
)
from easy_access.settings import Settings
from easy_access.utils import determine_course_code


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
            logger.warning(
                f"Error while trying to get llm classification for item {item.material_id}: {e}"
            )
            continue

    logger.info(
        f"Done linking {len(classifications_to_link)} LLM classifications to CopyrightItems."
    )


async def link_courses_to_copyright_items() -> None:
    items_w_prefetch = await CopyrightItem.all()
    logger.info(f"got {len(items_w_prefetch)} items from db")

    # for each of the items, extract the course code (using determine_course_code)
    # then match with existing course item in db
    # if missing, add to list to retrieve later

    links_added = 0
    course_codes_found = 0
    for item in items_w_prefetch:
        course_codes = determine_course_code(item.course_code, item.course_name)
        if not course_codes or len(course_codes) == 0:
            logger.warning(
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
                logger.warning(
                    f"Error while trying to get course {course_code} for item {item.material_id}: {e}"
                )
    logger.success(
        f"Added {links_added} links to {course_codes_found} found coursecodes."
    )


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
    logger.success(f"Updated {duplicates} duplicate statuses.")


async def update_copyright_relations(settings: Settings) -> None:
    """
    Go through the copyright items in the db
    use the values of the item fields to find links to other tables.
    """
    await ensure_db_inited(settings)
    await update_duplicate_status()

    await link_llm_classifications_to_copyright_items()
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


async def update_copyright_items(
    settings: Settings,
    data: pl.DataFrame | list[dict],
    update_relations: bool = True,
    overwrite: bool = False,
    user_info: dict | None = None,
) -> None:
    """
    TODO: Decide if we want to refactor this?

    IDEA:

    Split functionality based on type of input. Reuse logic where possible.
    Type of input mainly determines what to overwrite/update/compare.
    See also the notes in main.py for more details on the refactoring.

    Types of input:
    -> Raw data import from Qlik
    -> User input from weekly sheet
    -> User input from overview sheet
    -> User input from web dashboard
    -> Directly from script / overwrite


    All input should be dataframes or lists of dicts (which are turned into dataframes then?),
    standardized/normalized to the same format (handle that in this function? or expect input to be standardized?)

    Does not return anything, directly updates the db.

    New params:
    -> data: pl.DataFrame | list[dict] (or only accept pl.DataFrame?)
        The input data used to update the items. Pref. a standardized dataframe.
    -> source: DataSource (StrEnum), default: DataSource.EA_SCRIPT
        The input source of data, used to determine how to handle updating. Use the enum DataSource.
    --> update_relations: bool, default: True
        If True, will call `update_copyright_relations` to update the m2m relations after the update.
    --> overwrite: bool, default: False
        If True, will overwrite the existing items in the db with the new ones instead of using the comparison logic.
        This is automatically set to True if the source is DataSource.EA_SCRIPT or DataSource.WEB_DASHBOARD.
    --> user_info: dict | None, default: None
        The user info used to determine who made the changes. If None, will use the default user info.
        Currently only uses the 'email' field, stored in 'modified_by' in ItemUpdate.changes.

    """
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

    def change(
        changes: dict, field: str, new_value: Any, old_value: Any, reason: str
    ) -> dict:
        logger.debug(f"[{reason}] [{field}] {old_value} --> {new_value}")
        changes[field] = {"old": str(old_value), "new": str(new_value)}
        setattr(db_item, field, new_value)
        return changes

    def compare_fields(
        new_item: dict, db_item: CopyrightItem, fielddict: dict, changes: dict
    ) -> tuple[dict, CopyrightItem]:
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
                                tzinfo=UTC
                            )
                            if new_value
                            else None
                        )
                    except Exception:
                        try:
                            new_value = (
                                datetime.strptime(
                                    new_value, "%Y-%m-%d %H:%M:%S"
                                ).replace(tzinfo=UTC)
                                if new_value
                                else None
                            )
                        except Exception:
                            with contextlib.suppress(Exception):
                                new_value = (
                                    datetime.strptime(new_value, "%Y-%m-%d").replace(
                                        tzinfo=UTC
                                    )
                                    if new_value
                                    else None
                                )
                    old_value = old_value.replace(tzinfo=UTC)
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

    await ensure_db_inited(settings)
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

    # standardize df
    # loop over items
    # if item is not in db: add it
    # else compare values in specific fields to determine if we need to update
    logger.info(f"Received {len(data)} raw copyright items as input for an update.")

    new_items = []
    if isinstance(data, pl.DataFrame):
        data = standardize_dataframe(data)
        existing_mat_ids = await CopyrightItem.all().values("material_id")
        existing_mat_ids = {int(m["material_id"]) for m in existing_mat_ids}
        # Candidate new items (may be partial if coming from faculty sheets)
        candidate_new_items = (
            data.with_columns(pl.col("material_id").cast(int))
            .filter(~pl.col("material_id").is_in(existing_mat_ids))
            .to_dicts()
        )
        # Only create new CopyrightItem objects when the incoming row contains
        # the required, non-nullable fields present in the model. Faculty-sheet
        # updates intentionally only include a few fields (material_id, workflow_status,
        # remarks, manual_classification); those should not trigger creation of a
        # full CopyrightItem. Filter out partial rows to avoid ValueError on save.
        required_for_creation = [
            "period",
            "department",
            "course_code",
            "course_name",
        ]
        new_items = []
        skipped_mat_ids: list[int] = []
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
                try:
                    skipped_mat_ids.append(int(itm.get("material_id")))
                except Exception:
                    skipped_mat_ids.append(itm.get("material_id"))
        if skipped_mat_ids:
            logger.warning(
                f"Skipping {len(skipped_mat_ids)} new items missing required fields (not creating in DB): {skipped_mat_ids[:20]}"
            )
        update_items = (
            data.with_columns(pl.col("material_id").cast(int))
            .filter(pl.col("material_id").is_in(existing_mat_ids))
            .to_dicts()
        )
    else:
        if isinstance(data, list):
            update_items = data

    logger.info(f"# of new items: {len(new_items)}")
    logger.info(f"# of items to update: {len(update_items)}")
    new_objects = []
    if new_items:
        new_objects = [await copyright_item_from_dict(item) for item in new_items]
        new_objects = [item for item in new_objects if item]
        try:
            await CopyrightItem.bulk_create(objects=new_objects)
        except Exception as e:
            logger.warning(
                f"error while trying to bulk save items. Error: {e}. Trying one-by-one."
            )
            for item in new_objects:
                try:
                    await item.save()
                except Exception as e:
                    logger.error(e)
                    logger.warning(
                        f"error {e} while trying to save item {item} with material_id {item.material_id}. Skipping for now."
                    )
        logger.success(f"Created {len(new_objects)} new copyright items in db.")

    logger.info(f"Updating {len(update_items)} existing items.")
    changelist = []
    updates = {}
    for new_item in update_items:
        if overwrite:
            try:
                # if overwrite is True, just create a new item and skip the rest
                db_item = await CopyrightItem.get(
                    material_id=new_item.get("material_id")
                )

                changes = {
                    "material_id": new_item.get("material_id"),
                    "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }

                logger.debug(
                    f"now in overwrite function for {new_item.get('material_id')}"
                )
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
                        changes = change(
                            changes,
                            k,
                            new_item.get(k),
                            getattr(db_item, k),
                            "[overwrite] new value != old value",
                        )
                        logger.debug(f"changes: {changes}")
                # if any changes were made we'll have 3 or more keys in the changes dict
                # if not, no need to update the db
                logger.debug("final changes:")
                logger.debug(changes)
                if len(changes) >= 3:
                    changes["modified_at"] = datetime.now()
                    updates[new_item.get("material_id")] = changes
                    changelist.append(db_item)
                else:
                    logger.debug(f"No changes for item {new_item.get('material_id')}.")
            except Exception as e:
                logger.warning(
                    f"Could not update item {new_item.get('material_id')}: {e}"
                )
                logger.warning(traceback.format_exc())
        else:
            try:
                db_item = await CopyrightItem.get(
                    material_id=new_item.get("material_id")
                )
                changes = {}
                changes, db_item = compare_fields(
                    new_item, db_item, added_fields, changes
                )
                changes, db_item = compare_fields(
                    new_item, db_item, changeable_fields, changes
                )

            except Exception as e:
                logger.warning(
                    f"Could not update item {new_item.get('material_id')}: {e}"
                )
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
        all_keys = {key for item in updates.values() for key in item}
        all_keys.discard("material_id")
        all_keys.discard("update_time")
        changed_fields = list(all_keys)
        changed_fields.append("modified_at")
        logger.info(
            f"Updating {len(changelist)} items in db for fields {changed_fields}."
        )
        logger.info(f"Updating {len(updates)} changelog items in db.")
        await CopyrightItem.bulk_update(changelist, fields=changed_fields)
        if user_info:
            [changes.update({"modified_by": cur_user}) for changes in updates.values()]
            logger.info(f"items modified by {cur_user}")

        await ItemUpdate.bulk_create(
            [
                ItemUpdate(change_details=changes, material_id=mat_id)
                for mat_id, changes in updates.items()
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

    if (changelist or new_objects) and update_relations:
        logger.success("Updating relations for all CopyrightItems.")
        await update_copyright_relations(settings=settings)

    logger.success("Done updating CopyrightItems!")
    await Tortoise.close_connections()
