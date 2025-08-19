"""
This module contains functions for ingesting new data into the database.
It handles loading foundational data like organizational structures, courses, and persons,
as well as transactional data such as raw copyright exports, LLM classifications,
and PDF metadata. Most functions are asynchronous due to reliance on Tortoise ORM.
"""

import json
import logging
import traceback
from collections import Counter
from typing import Any  # For dict values

import polars as pl
from tortoise import Tortoise

from easy_access.db.base import (
    copyright_item_from_dict,
    standardize_dataframe,
)
from easy_access.db.base import (
    init as init_tortoise,  # Renamed to avoid conflict if there was a local 'init'
)
from easy_access.db.models import (
    PDF,
    CopyrightItem,
    Course,
    CourseEmployee,
    Faculty,
    LLMClassification,
    MissingCourse,  # Should be handled if courses are missing
    Organization,
    Person,
    Programme,
    WorkflowStatus,
)
from easy_access.db.update import (  # These are async, ensure they handle their own connections or are awaited properly
    update_copyright_items,  # Used for updating existing items found in raw load
    update_copyright_relations,
)
from easy_access.settings import (
    DEPARTMENT_MAPPING,  # Used in load_raw_copyright_data's helper
    SETTINGS,
    DirSetting,
    FileSetting,
    SettingsFaculty,
)
from easy_access.utils import Directory, File

logger = logging.getLogger(__name__)


async def load_osiris_data() -> None:
    """
    Loads course data from the `osiris_data.json` file into the `Course` table.
    Skips courses that already exist in the database based on `cursuscode`.
    Links courses to faculties if faculty information is available.
    """
    await init_tortoise()
    osiris_data_file: File = SETTINGS.files[FileSetting.OSIRIS_DATA]
    if not osiris_data_file.exists:
        logger.warning(
            f"{osiris_data_file.path} not found; Osiris course data will not be loaded to DB."
        )
        return

    try:
        with open(osiris_data_file.path, encoding="utf-8") as f:
            osiris_data_dict: dict[str, dict[str, Any]] = json.load(f)
    except json.JSONDecodeError as e:
        logger.error(f"Error decoding JSON from {osiris_data_file.path}: {e}")
        return
    except Exception as e:
        logger.error(f"Error reading {osiris_data_file.path}: {e}")
        return

    course_dicts_to_create: list[dict[str, Any]] = []
    try:
        existing_course_codes_query = await Course.all().values_list(
            "cursuscode", flat=True
        )
        existing_course_codes: set[int] = set(existing_course_codes_query)  # type: ignore # Tortoise returns list of tuples or values

        for course_data in osiris_data_dict.values():
            cursuscode_str = course_data.get("cursuscode")
            if not cursuscode_str or not str(cursuscode_str).isdigit():
                logger.debug(
                    f"Skipping Osiris course data due to missing or invalid cursuscode: {course_data.get('name')}"
                )
                continue

            cursuscode = int(cursuscode_str)
            if cursuscode in existing_course_codes:
                continue

            ec_str = str(course_data.get("ec", "0")).replace(",", ".")

            course_dict: dict[str, Any] = {
                "cursuscode": cursuscode,
                "internal_id": int(course_data["internal_id"])
                if course_data.get("internal_id")
                and str(course_data.get("internal_id")).isdigit()
                else None,
                "name": course_data.get("name"),
                "short_name": course_data.get("short_name"),
                "ec": int(round(float(ec_str)))
                if ec_str and ec_str.replace(".", "", 1).isdigit()
                else None,
                "programme": course_data.get("programme"),
                "notes": course_data.get("notes"),
                "category": course_data.get("category"),
            }

            year_str = course_data.get("year")
            if year_str and "-" in year_str:
                course_dict["year"] = int(year_str.split("-")[0])
            elif year_str and year_str.isdigit():
                course_dict["year"] = int(year_str)

            faculty_abbr = course_data.get("faculty")
            if faculty_abbr:
                faculty = await Faculty.get_or_none(abbreviation=faculty_abbr)
                if faculty:
                    course_dict["faculty"] = faculty
                else:
                    logger.debug(
                        f"Faculty '{faculty_abbr}' not found for course '{course_dict['name']}'."
                    )

            course_dicts_to_create.append(course_dict)
            existing_course_codes.add(
                cursuscode
            )  # Add to set to prevent re-processing if duplicated in source JSON

        if course_dicts_to_create:
            await Course.bulk_create(
                objects=[Course(**c) for c in course_dicts_to_create]
            )
            logger.info(
                f"Created {len(course_dicts_to_create)} new courses from Osiris data."
            )
        else:
            logger.info("No new courses found in Osiris data to load.")

    except Exception as e:  # Catch broader errors during processing
        logger.error(f"An error occurred during Osiris data processing: {e}")
        logger.debug(traceback.format_exc())
    finally:
        await Tortoise.close_connections()


async def load_org_data_from_settings() -> None:
    """
    Loads organizational data (University, Faculties, Programmes) from the global SETTINGS
    into the database. Creates records if they don't already exist.
    """
    await init_tortoise()
    try:
        faculties_from_settings: list[SettingsFaculty] = (
            SETTINGS.university_settings.faculties
        )

        university, _ = await Organization.get_or_create(
            abbreviation="UT",  # Assuming "UT" is the primary key or unique identifier
            defaults={
                "name": SETTINGS.university_settings.name or "University of Twente",
                "full_abbreviation": SETTINGS.university_settings.abbreviation or "UT",
                "parent_organization": None,
                "hierarchy_level": 0,
            },
        )

        for faculty_setting in faculties_from_settings:
            faculty_obj, _ = await Faculty.get_or_create(
                abbreviation=faculty_setting.abbreviation,
                defaults={
                    "name": faculty_setting.name,
                    "full_abbreviation": faculty_setting.abbreviation,  # Assuming full_abbr is same as abbr for Faculty
                    "parent_organization": university,
                    "hierarchy_level": 1,  # Faculties are level 1
                },
            )

            programmes_to_create: list[dict[str, Any]] = []
            existing_programmes_q = (
                await Programme.filter(faculty=faculty_obj)
                .all()
                .values_list("name", "abbreviation")
            )
            existing_programmes_set: set[tuple[str, str | None]] = {
                (name, abbr) for name, abbr in existing_programmes_q
            }  # type: ignore

            for programme_setting in faculty_setting.programmes:
                if not programme_setting.name:
                    continue  # Skip if no name

                prog_abbr = programme_setting.abbreviation
                if not prog_abbr:  # Attempt to generate abbreviation if missing
                    abbr_parts: list[str] = []
                    if programme_setting.programme_type == "b":
                        abbr_parts.append("B")
                    elif programme_setting.programme_type == "m":
                        abbr_parts.append("M")
                    else:
                        abbr_parts.append("O")  # Other/Unknown

                    name_parts = (
                        programme_setting.name.lower()
                        .replace("bachelor", "")
                        .replace("master", "")
                        .strip()
                        .split(" ")
                    )
                    abbr_parts.extend(
                        [part[0] for part in name_parts if part][:2]
                    )  # First letter of first two words
                    prog_abbr = "-".join(abbr_parts).upper()

                if (programme_setting.name, prog_abbr) not in existing_programmes_set:
                    programmes_to_create.append(
                        {
                            "name": programme_setting.name,
                            "abbreviation": prog_abbr,
                            "programme_type": programme_setting.programme_type,
                            "faculty": faculty_obj,
                            "cluster": programme_setting.cluster,
                        }
                    )
                    existing_programmes_set.add((programme_setting.name, prog_abbr))

            if programmes_to_create:
                await Programme.bulk_create(
                    objects=[Programme(**p) for p in programmes_to_create]
                )
                logger.info(
                    f"Created {len(programmes_to_create)} programmes for faculty {faculty_setting.abbreviation}."
                )
        logger.info("Organization data from settings loaded/verified.")
    except Exception as e:
        logger.error(f"Error loading organization data from settings: {e}")
        logger.debug(traceback.format_exc())
    finally:
        await Tortoise.close_connections()


async def load_person_data() -> None:
    """
    Loads person data from `person_data.json` file.
    Creates `Person` and `Organization` records, linking them as specified.
    Skips persons already existing in the database based on `input_name`.
    """
    await init_tortoise()
    person_data_file: File = SETTINGS.files[FileSetting.PERSON_DATA]
    if not person_data_file.exists:
        logger.warning(
            f"{person_data_file.path} not found; person data will not be loaded."
        )
        return

    try:
        with open(person_data_file.path, encoding="utf-8") as f:
            persons_json_data: list[dict[str, Any]] = json.load(f)
    except json.JSONDecodeError as e:
        logger.error(f"Error decoding JSON from {person_data_file.path}: {e}")
        return
    except Exception as e:
        logger.error(f"Error reading {person_data_file.path}: {e}")
        return

    try:
        existing_person_input_names_q = await Person.all().values_list(
            "input_name", flat=True
        )
        existing_person_input_names: set[str] = set(existing_person_input_names_q)  # type: ignore

        university_org = await Organization.get_or_none(
            abbreviation="UT", hierarchy_level=0
        )
        if not university_org:
            logger.error(
                "University 'UT' root organization not found. Cannot properly set parent organizations."
            )
            # Decide if to proceed or halt. For now, will proceed but parent links might be broken.

        for person_data_dict in persons_json_data:
            input_name = str(person_data_dict.get("input_name", "")).strip()
            if not input_name or input_name in existing_person_input_names:
                continue

            faculty_abbr = person_data_dict.get("faculty")
            faculty_obj: Faculty | None = None
            if faculty_abbr:
                faculty_obj = await Faculty.get_or_none(abbreviation=str(faculty_abbr))

            person_model_data: dict[str, Any] = {
                "input_name": input_name,
                "main_name": str(person_data_dict.get("main_name", "")).strip() or None,
                "match_confidence": float(person_data_dict["match_confidence"])
                if person_data_dict.get("match_confidence") is not None
                else None,
                "first_name": str(
                    person_data_dict.get("other_names", [None])[0] or ""
                ).strip()
                or None,  # Assuming first item is first name
                "email": str(person_data_dict.get("email", "")).strip() or None,
                "faculty": faculty_obj,
                "people_page_url": str(
                    person_data_dict.get("people_page_url", "")
                ).strip()
                or None,
            }

            created_person = await Person.create(**person_model_data)
            existing_person_input_names.add(input_name)  # Add to set after creation

            org_objects_for_person: list[Organization] = []
            org_data_list = person_data_dict.get("orgs", [])
            if isinstance(org_data_list, list):
                org_data_list.sort(
                    key=lambda x: x.get("abbr", "").count("-")
                )  # Process parents first

                for org_dict_item in org_data_list:
                    if not isinstance(org_dict_item, dict):
                        continue
                    try:
                        org_name = str(org_dict_item.get("name", "")).strip()
                        org_full_abbr = str(org_dict_item.get("abbr", "")).strip()
                        if not org_name or not org_full_abbr:
                            continue

                        org_sole_abbr = org_full_abbr.split("-")[-1].strip()
                        hierarchy_level = (
                            org_full_abbr.count("-") + 1
                        )  # Level 1 for "FAC", 2 for "FAC-DEPT"

                        org_model_data: dict[str, Any] = {
                            "name": org_name,
                            "abbreviation": org_sole_abbr,
                            "full_abbreviation": org_full_abbr,
                            "hierarchy_level": hierarchy_level,
                        }

                        # Determine parent organization
                        parent_org_obj: Organization | None = None
                        if (
                            hierarchy_level == 1 and university_org
                        ):  # Faculty, parent is University
                            parent_org_obj = university_org
                        elif hierarchy_level > 1:
                            parent_full_abbr = org_full_abbr.rsplit("-", 1)[0].strip()
                            parent_org_obj = await Organization.get_or_none(
                                full_abbreviation=parent_full_abbr
                            )

                        org_model_data["parent_organization"] = parent_org_obj

                        # Use get_or_create for organizations to avoid duplicates
                        # Unique constraint on full_abbreviation is important here
                        org_obj, _ = await Organization.get_or_create(
                            full_abbreviation=org_full_abbr, defaults=org_model_data
                        )
                        org_objects_for_person.append(org_obj)
                    except Exception as e_org:
                        logger.warning(
                            f"Error processing organization {org_dict_item} for person {input_name}: {e_org}"
                        )

            if org_objects_for_person:
                await created_person.orgs.add(*org_objects_for_person)
        logger.info("Person data loading complete.")
    except Exception as e:
        logger.error(f"An error occurred during person data loading: {e}")
        logger.debug(traceback.format_exc())
    finally:
        await Tortoise.close_connections()


async def load_linked_persons_for_courses() -> Counter[str]:
    """
    Links `Person` records to `Course` records via the `CourseEmployee` through-table,
    based on roles like 'teacher', 'contact', etc., specified in `osiris_data.json`.

    Returns:
        Counter[str]: A counter of roles and how many links were created for each.
    """
    await init_tortoise()
    counter: Counter[str] = Counter()
    osiris_data_file: File = SETTINGS.files[FileSetting.OSIRIS_DATA]

    if not osiris_data_file.exists:
        logger.warning(
            f"{osiris_data_file.path} not found; cannot link persons to courses."
        )
        return counter

    try:
        with open(osiris_data_file.path, encoding="utf-8") as f:
            osiris_data_dict: dict[str, dict[str, Any]] = json.load(f)
    except Exception as e:
        logger.error(f"Error reading {osiris_data_file.path} for linking persons: {e}")
        return counter

    try:
        for (
            course_data_item
        ) in osiris_data_dict.values():  # Renamed course_data to course_data_item
            cursuscode_str = course_data_item.get("cursuscode")
            if not cursuscode_str or not str(cursuscode_str).isdigit():
                continue

            course_obj = await Course.get_or_none(cursuscode=int(cursuscode_str))
            if not course_obj:
                logger.debug(f"Course {cursuscode_str} not found, cannot link persons.")
                continue

            roles_to_process: list[
                tuple[str, str]
            ] = [  # (role_name_in_json, role_name_in_db)
                ("teachers", "teacher"),
                ("contacts", "contact"),
                ("tutors", "tutor"),
                (
                    "docenten",
                    "docent",
                ),  # Assuming 'docenten' is another list of contacts/teachers
                ("examinators", "examinator"),
                ("unknown_role", "unknown_role"),
            ]

            for json_role_key, db_role_name in roles_to_process:
                person_names_list = course_data_item.get(json_role_key, [])
                if isinstance(person_names_list, list):
                    for person_name in person_names_list:
                        if not isinstance(person_name, str):
                            continue  # Skip if name is not string
                        person_obj = await Person.get_or_none(
                            input_name=person_name.strip()
                        )
                        if person_obj:
                            _, created = await CourseEmployee.get_or_create(
                                course=course_obj, person=person_obj, role=db_role_name
                            )
                            if created:
                                counter[db_role_name] += 1
                        else:
                            logger.debug(
                                f"Person '{person_name}' not found for course {cursuscode_str}, role '{db_role_name}'."
                            )
        logger.info(f"Finished linking persons to courses. Link counts: {counter}")
    except Exception as e:
        logger.error(f"An error occurred linking persons to courses: {e}")
        logger.debug(traceback.format_exc())
    finally:
        await Tortoise.close_connections()
    return counter


async def load_base_data() -> None:
    """
    Orchestrates the loading of all foundational data into the database.
    This includes organizational structure, courses, persons, and their relationships.
    Should be run when setting up a new database or refreshing this core data.
    """
    logger.info(
        "Starting to load all base data (organizations, courses, persons, links)..."
    )
    await init_tortoise()  # Ensure DB is ready
    # `create()` is likely redundant if `init_tortoise` calls `generate_schemas`
    # await create() # `create` was an alias for generate_schemas, init_tortoise handles this.

    try:
        # Log counts before loading
        counts_before: dict[str, int] = {
            "Faculty": await Faculty.all().count(),
            "Programme": await Programme.all().count(),
            "Course": await Course.all().count(),
            "Person": await Person.all().count(),
            "Organization": await Organization.all().count(),
            "MissingCourse": await MissingCourse.all().count(),
            "CourseEmployee": await CourseEmployee.all().count(),
        }
        logger.info(f"DB counts before loading base data: {counts_before}")

        await load_org_data_from_settings()
        await load_osiris_data()
        await load_person_data()
        link_results = await load_linked_persons_for_courses()

        counts_after: dict[str, int] = {
            "Faculty": await Faculty.all().count(),
            "Programme": await Programme.all().count(),
            "Course": await Course.all().count(),
            "Person": await Person.all().count(),
            "Organization": await Organization.all().count(),
            "MissingCourse": await MissingCourse.all().count(),
            "CourseEmployee": await CourseEmployee.all().count(),
        }
        logger.info(f"DB counts after loading base data: {counts_after}")
        logger.info(f"Overview of new Course-Employee links added: {link_results}")
        logger.info("Base data loading process completed.")
    except Exception as e:
        logger.error(f"Error during comprehensive base data loading: {e}")
        logger.debug(traceback.format_exc())
    finally:
        await Tortoise.close_connections()


async def load_raw_copyright_data(
    source_data: File | pl.DataFrame,
) -> None:  # Removed None option, require data
    """
    Loads raw copyright items into the database from a given source (File or DataFrame).
    New items are created, and existing items are identified for update.
    After processing, LLM classifications and copyright relations are updated.

    Args:
        source_data (File | pl.DataFrame): The source of raw copyright data.
                                         If a File object, it's read as an Excel sheet.
                                         If a DataFrame, it's used directly.
    """

    # Nested helper function to read and do initial processing of an Excel file
    def _read_and_prep_excel(file_obj: File) -> pl.DataFrame:
        """Reads copyright data from an Excel file and performs initial preparation."""
        logger.info(f"Reading raw copyright data from file: {file_obj.name}")
        try:
            latest_file_date_str = file_obj.created.strftime(
                "%Y-%m-%d"
            )  # Assuming File obj has 'created'
            raw_df = pl.read_excel(file_obj.path)

            # Standardize column names (lowercase, underscores)
            renamed_df = raw_df.rename(
                lambda c: str(c)
                .replace(" ", "_")
                .replace("#", "count_")
                .replace("*", "x")
                .lower()
            )

            # Add metadata columns and perform initial transformations
            # Ensure DEPARTMENT_MAPPING is available
            transformed_df = renamed_df.with_columns(
                pl.lit(latest_file_date_str).alias("retrieved_from_copyright_on"),
                pl.lit(WorkflowStatus.ToDo.value).alias(
                    "workflow_status"
                ),  # Default workflow
                pl.col("last_change")
                .str.replace(
                    r"^- ভারতবর্ষ$", None
                )  # Example of specific cleaning if needed
                .str.strip_chars()
                .str.strptime(pl.Date, "%Y-%m-%d", strict=False)
                .dt.strftime("%Y-%m-%d"),
                pl.col("classification").str.to_lowercase(),
                faculty=pl.col("department").replace_strict(
                    DEPARTMENT_MAPPING, default="Unmapped"
                ),
            )
            logger.info(f"Retrieved {len(transformed_df)} items from {file_obj.name}.")
            return transformed_df
        except Exception as e_read:  # Catch errors during file read/prep
            logger.error(
                f"Error reading or prepping Excel file {file_obj.name}: {e_read}"
            )
            return pl.DataFrame()  # Return empty DataFrame on error

    await init_tortoise()
    process_error: Exception | None = None

    try:
        logger.info(
            f"Items in DB before loading raw copyright data: {await CopyrightItem.all().count()}"
        )

        df_to_process: pl.DataFrame
        if isinstance(source_data, File):
            df_to_process = _read_and_prep_excel(source_data)
        elif isinstance(source_data, pl.DataFrame):
            logger.info(
                f"Loading {len(source_data)} raw copyright items from provided DataFrame."
            )
            df_to_process = source_data.clone()  # Use a copy
        else:
            logger.error(
                f"Unsupported data type for load_raw_copyright_data: {type(source_data)}"
            )
            return  # Exit if data type is wrong

        if df_to_process.is_empty():
            logger.warning("No data to process after reading/preparation stage.")
            return

        # Standardize further (filters, drops etc.)
        # This standardize_dataframe is from db.base
        standardized_items_df = standardize_dataframe(df_to_process)
        if standardized_items_df.is_empty():
            logger.warning("No data remaining after standardization.")
            return

        item_dicts_for_processing = standardized_items_df.to_dicts()

        existing_mat_ids_query = await CopyrightItem.all().values_list(
            "material_id", flat=True
        )
        existing_mat_ids: set[int] = set(existing_mat_ids_query)  # type: ignore
        logger.info(
            f"Read {len(item_dicts_for_processing)} standardized items. {len(existing_mat_ids)} items already in DB."
        )

        new_item_orm_list: list[CopyrightItem] = []
        items_for_update_list: list[
            dict[str, Any]
        ] = []  # List of dicts for update_copyright_items

        for item_dict_idx, item_data_dict in enumerate(item_dicts_for_processing):
            material_id_val = item_data_dict.get("material_id")
            if material_id_val is None or not str(material_id_val).isdigit():
                logger.warning(
                    f"Skipping item due to missing/invalid material_id: {item_data_dict.get('filename', 'N/A')}"
                )
                continue

            material_id = int(material_id_val)

            if material_id in existing_mat_ids:
                items_for_update_list.append(item_data_dict)
            else:
                created_orm_item = await copyright_item_from_dict(
                    item_data_dict
                )  # This converts types
                if created_orm_item:
                    new_item_orm_list.append(created_orm_item)
                else:
                    logger.warning(
                        f"Failed to convert item dict to ORM object: {material_id}"
                    )

            if (item_dict_idx + 1) % 100 == 0:
                logger.debug(
                    f"Processed {item_dict_idx + 1}/{len(item_dicts_for_processing)} raw items for DB ingestion."
                )

        if new_item_orm_list:
            await CopyrightItem.bulk_create(new_item_orm_list)
            logger.info(f"Bulk created {len(new_item_orm_list)} new copyright items.")

        item_count_after_new = await CopyrightItem.all().count()
        logger.info(
            f"Total items in DB after creating new ones: {item_count_after_new}"
        )

        if items_for_update_list:
            logger.info(
                f"Found {len(items_for_update_list)} existing items from raw data to check for updates."
            )
            # update_copyright_items handles its own Tortoise init/close and logging for updates.
            # Pass overwrite=True because raw Qlik data is usually authoritative for its specific fields.
            await update_copyright_items(
                items_for_update_list, overwrite=True, update_relations=False
            )
            # update_relations=False here because we do it once at the end.

        # Update relations after all new items and updates are processed
        if new_item_orm_list or items_for_update_list:
            logger.info(
                "Updating LLM classifications and copyright relations after raw data load."
            )
            await load_llm_classifications()  # This should handle its own init/close
            await (
                update_copyright_relations()
            )  # This should also handle its own init/close

    except Exception as e_main:
        logger.error(f"Major error during raw copyright data loading: {e_main}")
        logger.debug(traceback.format_exc())
        process_error = e_main  # Store error to re-raise after cleanup
    finally:
        await Tortoise.close_connections()
        if process_error and not (
            new_item_orm_list or items_for_update_list
        ):  # Only re-raise if nothing was processed
            raise process_error


async def load_llm_classifications() -> None:
    """
    Loads LLM classification data from JSON files into the `LLMClassification` table.
    Skips classifications that already exist based on `used_material_id`.
    Deletes JSON files that are empty or lack essential data after attempting to load.
    """
    await init_tortoise()
    try:
        existing_class_mat_ids_q = await LLMClassification.all().values_list(
            "used_material_id", flat=True
        )
        existing_class_mat_ids: set[int] = set(existing_class_mat_ids_q)  # type: ignore
        logger.info(
            f"Found {len(existing_class_mat_ids)} existing LLM classifications in DB."
        )

        classifications_dir: Directory = SETTINGS.dirs[DirSetting.CLASSIFICATIONS]
        all_json_files: list[File] = [
            f for f in classifications_dir.files if f.name.endswith(".json")
        ]

        # Clean up old files (if any) - this logic might be better placed in a separate utility
        old_files_to_delete = [f for f in all_json_files if "_old" in f.name]
        if old_files_to_delete:
            logger.info(
                f"Found {len(old_files_to_delete)} old LLM JSON files. Removing..."
            )
            for old_f in old_files_to_delete:
                old_f.delete()

        # Re-list files after potential deletion
        current_json_files: list[File] = [
            f for f in classifications_dir.files if f.name.endswith(".json")
        ]
        logger.info(f"Found {len(current_json_files)} LLM JSON files to process.")

        new_classifications_data: list[dict[str, Any]] = []
        files_to_delete_after_processing: list[File] = []

        for json_file in current_json_files:
            try:
                material_id_str = json_file.name.replace(".json", "").strip()
                if not material_id_str.isdigit():
                    logger.warning(
                        f"Skipping JSON file with non-numeric name: {json_file.name}"
                    )
                    continue

                material_id = int(material_id_str)
                if material_id in existing_class_mat_ids:
                    continue  # Skip already processed

                with open(json_file.path, encoding="utf-8", errors="ignore") as f:
                    data_dict = json.load(f)

                data_dict["used_material_id"] = (
                    material_id  # Ensure this key is present
                )
                if not data_dict.get("allowed_usage"):  # Check for essential data
                    files_to_delete_after_processing.append(json_file)
                else:
                    new_classifications_data.append(data_dict)
            except json.JSONDecodeError:
                logger.warning(
                    f"Invalid JSON in file {json_file.name}. Marking for deletion."
                )
                files_to_delete_after_processing.append(json_file)
            except Exception as e_file:
                logger.warning(
                    f"Error processing LLM JSON file {json_file.name}: {e_file}"
                )
                files_to_delete_after_processing.append(
                    json_file
                )  # Mark for deletion on other errors too

        logger.info(
            f"Retrieved {len(new_classifications_data)} new LLM classifications from JSON files."
        )
        for f_del in files_to_delete_after_processing:
            if (
                f_del in new_classifications_data
            ):  # Ensure we don't try to use it if it's marked for deletion
                new_classifications_data.remove(
                    f_del
                )  # This comparison is wrong, need to remove by mat_id or filename
            logger.info(f"Deleting invalid/empty LLM JSON file: {f_del.name}")
            f_del.delete()

        if new_classifications_data:
            orm_objects_to_create: list[LLMClassification] = []
            for classification_dict in new_classifications_data:
                # Convert dict to LLMClassification ORM object, handling potential errors
                # Assuming direct mapping for now; add validation/conversion if needed
                try:
                    # Ensure all list fields are indeed lists, not strings from JSON.
                    # Tortoise JSONField should handle stringified lists if necessary, but better to be clean.
                    list_fields = [
                        "author_names",
                        "doi",
                        "isbn",
                        "source_url",
                        "license",
                        "topic",
                    ]
                    for field_name in list_fields:
                        if field_name in classification_dict and isinstance(
                            classification_dict[field_name], str
                        ):
                            try:
                                potential_list = json.loads(
                                    classification_dict[field_name]
                                )
                                if isinstance(potential_list, list):
                                    classification_dict[field_name] = potential_list
                                else:  # Not a list, keep as string or wrap in list
                                    classification_dict[field_name] = [
                                        str(classification_dict[field_name])
                                    ]
                            except json.JSONDecodeError:  # Not a JSON list string
                                classification_dict[field_name] = (
                                    [str(classification_dict[field_name])]
                                    if classification_dict[field_name]
                                    else []
                                )
                        elif field_name in classification_dict and not isinstance(
                            classification_dict[field_name], list
                        ):
                            classification_dict[field_name] = (
                                [str(classification_dict[field_name])]
                                if classification_dict[field_name]
                                else []
                            )

                    orm_objects_to_create.append(
                        LLMClassification(**classification_dict)
                    )
                except Exception as e_orm:
                    logger.warning(
                        f"Error creating LLMClassification ORM object from dict {classification_dict.get('used_material_id')}: {e_orm}"
                    )

            if orm_objects_to_create:
                await LLMClassification.bulk_create(orm_objects_to_create)
                logger.info(
                    f"Created {len(orm_objects_to_create)} new LLM classifications in DB."
                )
        else:
            logger.info("No valid new LLM classifications found to add to DB.")

        # This linking should happen after all CopyrightItems are potentially in the DB
        # It's called from load_raw_copyright_data and update_copyright_relations already.
        # Calling it here again might be redundant if part of a larger flow.
        # await link_llm_classifications_to_copyright_items()
    except Exception as e:
        logger.error(f"General error in load_llm_classifications: {e}")
        logger.debug(traceback.format_exc())
    finally:
        await Tortoise.close_connections()


async def load_pdfs() -> None:
    """
    Loads metadata for PDF files found in the designated PDF downloads directory
    into the `PDF` table. It links PDFs to existing `CopyrightItem` records
    based on material ID extracted from filenames. Skips PDFs if no related
    CopyrightItem is found or if PDF metadata already exists.
    """
    await init_tortoise()
    try:
        pdfs_dir: Directory = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]
        if not pdfs_dir.exists:
            logger.warning(
                f"PDF downloads directory {pdfs_dir.full} not found. Cannot load PDF metadata."
            )
            return

        all_pdf_files_on_disk: list[File] = [
            f
            for f in pdfs_dir.files_r
            if f.name.endswith(".pdf") and not f.name.startswith("~$")
        ]

        pdf_files_by_mat_id: dict[int, File] = {}
        for pdf_file in all_pdf_files_on_disk:
            # Extract material_id from filename (e.g., "12345_document.pdf" or "12345.pdf")
            name_part = (
                pdf_file.name.split("_")[0]
                if "_" in pdf_file.name
                else pdf_file.name.split(".")[0]
            )
            if name_part.isdigit():
                pdf_files_by_mat_id[int(name_part)] = pdf_file
            else:
                logger.debug(
                    f"Could not extract material_id from PDF filename: {pdf_file.name}"
                )

        if not pdf_files_by_mat_id:
            logger.info(
                "No PDF files found with extractable material_ids in their names."
            )
            return

        existing_pdf_db_mat_ids_q = await PDF.all().values_list(
            "material_id", flat=True
        )
        existing_pdf_db_mat_ids: set[int] = set(existing_pdf_db_mat_ids_q)  # type: ignore

        pdf_metadata_to_create: list[dict[str, Any]] = []
        for mat_id, pdf_file_obj in pdf_files_by_mat_id.items():
            if mat_id in existing_pdf_db_mat_ids:
                continue  # Skip if PDF metadata already in DB for this material_id

            related_item = await CopyrightItem.get_or_none(material_id=mat_id)
            if not related_item:
                logger.debug(
                    f"No CopyrightItem found for material_id {mat_id} (PDF: {pdf_file_obj.name}). Skipping PDF metadata."
                )
                continue

            pdf_metadata_to_create.append(
                {
                    "material_id": mat_id,
                    "current_file_name": pdf_file_obj.name,
                    "original_file_name": related_item.filename,  # From CopyrightItem
                    "original_page_count": related_item.pagecount,  # From CopyrightItem
                    # Other PDF metadata fields (author, title etc.) would be populated by a PDF parsing step later.
                }
            )

        if pdf_metadata_to_create:
            await PDF.bulk_create(
                objects=[PDF(**p_dict) for p_dict in pdf_metadata_to_create]
            )
            logger.info(
                f"Created {len(pdf_metadata_to_create)} new PDF metadata records in DB."
            )
        else:
            logger.info(
                "No new PDF metadata records to create (either exist or no related CopyrightItem)."
            )

    except Exception as e:
        logger.error(f"An error occurred during PDF metadata loading: {e}")
        logger.debug(traceback.format_exc())
    finally:
        await Tortoise.close_connections()
