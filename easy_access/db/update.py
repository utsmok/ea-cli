"""
functions to update existing data in the database
"""

import contextlib
import traceback
from datetime import UTC, datetime, date
from enum import Enum, StrEnum
from itertools import batched
from typing import Any

import polars as pl
from loguru import logger
from tortoise import Tortoise
from tortoise.transactions import in_transaction

from easy_access.db.base import (
    copyright_item_from_dict,
    ensure_db_inited,
)
from easy_access.db.models import (
    PDF,
    Classification,
    CopyrightItem,
    Course,
    Infringement,
    ItemUpdate,
    StagedCopyrightItem,
    StagedFacultyUpdate,
    WorkflowStatus,
    Status,
)
from easy_access.settings import Settings
from easy_access.utils import determine_course_code, standardize_dataframe


# --- small parsing helpers to centralize validation and reduce type casting bugs ---
def safe_int(x: Any) -> int | None:
    if x is None:
        return None
    try:
        return int(x)
    except Exception:
        try:
            # sometimes float-like strings
            return int(float(x))
        except Exception:
            return None


def safe_float(x: Any) -> float | None:
    if x is None:
        return None
    try:
        return float(x)
    except Exception:
        return None


def safe_date(x: Any) -> date | None:
    if x is None:
        return None
    if isinstance(x, date) and not isinstance(x, datetime):
        return x
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, str):
        try:
            return datetime.fromisoformat(x).date()
        except Exception:
            for fmt in ("%Y-%m-%d %H:%M:%S%z", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
                try:
                    return datetime.strptime(x, fmt).date()
                except Exception:
                    continue
    return None


def safe_enum(enum_cls, value: Any):
    if value is None:
        return None
    try:
        # Try direct construction
        return enum_cls(value)
    except Exception:
        try:
            # Try name lookup
            if isinstance(value, str) and value in enum_cls.__members__:
                return enum_cls[value]
        except Exception:
            pass
    return None

# --- end helpers ---


def safe_compare_greater(a: Any, b: Any) -> bool:
    """Return True if a > b using safe normalization for numbers and dates, else False."""
    if a is None or b is None:
        return False
    # numeric comparison
    if isinstance(a, (int, float, str)) and isinstance(b, (int, float, str)):
        try:
            return float(a) > float(b)
        except Exception:
            return False
    # date/datetime comparison
    if isinstance(a, (date, datetime)) and isinstance(b, (date, datetime)):
        try:
            na = datetime.combine(a, datetime.min.time()) if isinstance(a, date) and not isinstance(a, datetime) else a
            nb = datetime.combine(b, datetime.min.time()) if isinstance(b, date) and not isinstance(b, datetime) else b
            return na > nb
        except Exception:
            return False
    return False


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
                cursuscode = safe_int(course_code)
                if cursuscode is None:
                    continue
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
        # leave replacement_id untouched unless we find a replacement
        mat_id = item.material_id
        pdf = await PDF.get_or_none(material_id=mat_id)
        if pdf:
            # use the foreign-key id to safely resolve the replacement without awaiting a possibly-None relation
            rid = getattr(pdf, "replace_with_id", None)
            replaced_item = None
            if rid:
                try:
                    replaced_item = await PDF.get_or_none(material_id=rid)
                except Exception:
                    replaced_item = None
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
        if field == "file_exists":
            db_item.last_canvas_check = datetime.now()
            changes["last_canvas_check"] = {
                "old": str(old_value),
                "new": str(new_value),
            }
        else:
            logger.debug(
                f"[{reason}] [{field}] {old_value}  {type(old_value)}) --> {new_value} ({type(new_value)})"
            )
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
                if field == "file_exists":
                    match new_value:
                        case True | 1 | "1" | "true" | "True":
                            new_value = True
                        case False | 0 | "0" | "false" | "False":
                            new_value = False
                        case None | "":
                            new_value = None
                        case _:
                            logger.warning(
                                f"[skip][file_exists] unexpected value: {new_value}"
                            )
                            continue

                    if not isinstance(new_value, bool):
                        continue

                    changes = change(
                        changes,
                        field,
                        new_value,
                        old_value,
                        "file_exists value received, always update",
                    )

                    continue

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
                    try:
                        if isinstance(new_value, (int, float, str)):
                            new_value = round(float(new_value), 2)
                        else:
                            new_value = None
                    except Exception:
                        new_value = None
                    old_value = round(old_value, 2)
                if isinstance(old_value, int):
                    try:
                        if new_value is None:
                            new_value = None
                        elif isinstance(new_value, (int, float, str)):
                            new_value = int(new_value)
                        else:
                            new_value = None
                    except Exception:
                        new_value = None

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
                        # use safe comparison helper to compare magnitude for numeric/date values
                        try:
                            if safe_compare_greater(new_value, old_value):
                                changes = change(
                                    changes, field, new_value, old_value, "new > old"
                                )
                        except Exception:
                            logger.debug(
                                f"Could not compare values for field {field}: {new_value} vs {old_value}"
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
        "file_exists": [False, 0, True, 1],
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
    update_items = []

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
        # the required, non-nullable fields present in the model.
        required_for_creation = [
            "period",
            "department",
            "course_code",
            "course_name",
        ]

        new_items = []
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

        # process in batches of max 50:

        for change_batch in batched(changelist, 50):
            logger.info(
                f"Updating fields {changed_fields} for {len(change_batch)} items that were changed."
            )

            await CopyrightItem.bulk_update(change_batch, fields=changed_fields)

        if user_info:
            [changes.update({"modified_by": cur_user}) for changes in updates.values()]
            logger.info(f"items modified by {cur_user}")

        for update_batch in batched(
            [(mat_id, changes) for mat_id, changes in updates.items()], 50
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
        await update_copyright_relations(settings=settings)

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
    ]

    processed_ids: list[int] = []

    # Process in batches to limit transaction size
    for batch in batched(staged_items, 50):
        batch_processed_ids: list[int] = []
        async with in_transaction() as conn:
            for staged_item in batch:
                try:
                    # Build a safe dict from known fields
                    item_dict: dict = {}
                    for f in staged_fields:
                        # Use getattr to avoid ORM internals
                        item_dict[f] = getattr(staged_item, f, None)

                    # Ensure material_id is present and castable
                    mid = item_dict.get("material_id")
                    if mid is None:
                        logger.warning(
                            f"Skipping staged row without material_id: {staged_item}"
                        )
                        continue

                    existing_item = await CopyrightItem.get_or_none(material_id=mid)

                    if not existing_item:
                        # Create new item using canonical normalizer
                        new_item = await copyright_item_from_dict(item_dict)
                        if new_item:
                            await new_item.save()
                            smid = safe_int(mid)
                            if smid is not None:
                                batch_processed_ids.append(smid)
                    else:
                        # Conservative updates for trivial fields
                        update_fields = []
                        # status: try to coerce to Status enum safely
                        status_val = item_dict.get("status")
                        if status_val:
                            try:
                                new_status = None
                                try:
                                    new_status = Status(status_val)
                                except Exception:
                                    # maybe name lookup
                                    if status_val in Status.__members__:
                                        new_status = Status[status_val]
                                if new_status and existing_item.status != new_status:
                                    existing_item.status = new_status
                                    update_fields.append("status")
                            except Exception:
                                logger.debug(f"Could not parse status value '{status_val}' for material {mid}")

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
                                    for fmt in ("%Y-%m-%d %H:%M:%S%z", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
                                        try:
                                            parsed_dt = datetime.strptime(lc_val, fmt)
                                            parsed_date = parsed_dt.date()
                                            break
                                        except Exception:
                                            continue
                            if parsed_date and existing_item.last_change != parsed_date:
                                existing_item.last_change = parsed_date
                                update_fields.append("last_change")

                        if update_fields:
                            await existing_item.save(update_fields=update_fields)
                            smid = safe_int(mid)
                            if smid is not None:
                                batch_processed_ids.append(smid)

                except Exception as e:
                    logger.exception(
                        f"Error processing staged item {getattr(staged_item,'material_id',None)}: {e}"
                    )
                    # Do not re-raise; keep other rows processing. Failed staged rows remain for manual inspection.

        # After successful transaction, remove successfully processed staged rows
        if batch_processed_ids:
            try:
                await StagedCopyrightItem.filter(material_id__in=batch_processed_ids).delete()
                logger.info(f"Cleared {len(batch_processed_ids)} processed staged rows.")
                processed_ids.extend(batch_processed_ids)
            except Exception as e:
                logger.exception(f"Error deleting staged rows {batch_processed_ids}: {e}")

    logger.info(f"Finished processing staged raw data. Successfully processed {len(processed_ids)} rows.")

async def process_staged_faculty_updates(settings: Settings) -> None:
    """
    Processes the staged faculty updates and updates the main CopyrightItem table.
    """
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
                    item = await CopyrightItem.get_or_none(material_id=update.material_id)
                    if not item:
                        continue
                    update_fields = []
                    if update.manual_classification and item.manual_classification != update.manual_classification:
                        item.manual_classification = update.manual_classification
                        update_fields.append("manual_classification")
                    if update.remarks and item.remarks != update.remarks:
                        item.remarks = update.remarks
                        update_fields.append("remarks")
                    if update.workflow_status and item.workflow_status != update.workflow_status:
                        # attempt to coerce to WorkflowStatus enum
                        wf_st = None
                        try:
                            if update.workflow_status in WorkflowStatus.__members__:
                                wf_st = WorkflowStatus[update.workflow_status]
                            else:
                                wf_st = WorkflowStatus(update.workflow_status)
                        except Exception:
                            wf_st = None

                        if wf_st:
                            item.workflow_status = wf_st
                            update_fields.append("workflow_status")

                    if update_fields:
                        await item.save(update_fields=update_fields)
                        smid = safe_int(update.material_id)
                        if smid is not None:
                            batch_processed.append(smid)
                except Exception:
                    logger.exception(f"Error applying staged faculty update {getattr(update,'material_id',None)}")

        if batch_processed:
            try:
                await StagedFacultyUpdate.filter(material_id__in=batch_processed).delete()
                logger.info(f"Cleared {len(batch_processed)} processed staged faculty updates.")
                processed_updates.extend(batch_processed)
            except Exception:
                logger.exception(f"Error deleting processed staged faculty updates: {batch_processed}")

    logger.info(f"Finished processing staged faculty updates. Successfully processed {len(processed_updates)} rows.")


async def calculate_derived_fields(settings: Settings) -> None:
    """
    Calculates and updates derived fields like 'possible_fine' and 'infringement'
    for all copyright items.
    """
    from easy_access.db.retrieve import retrieve_copyright_items

    logger.info("Calculating derived fields for all items...")
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
