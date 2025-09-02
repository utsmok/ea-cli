"""
functions to ingest new data into the database
"""

import contextlib
import json
import os
from collections import Counter
from pathlib import Path

import polars as pl
from loguru import logger
from tortoise import Tortoise

from easy_access.db.base import (
    copyright_item_from_dict,
    create,
    ensure_db_inited,
    standardize_dataframe,
)
from easy_access.db.models import (
    PDF,
    CopyrightItem,
    Course,
    CourseEmployee,
    Faculty,

    MissingCourse,
    Organization,
    Person,
    Programme,
)
from easy_access.db.update import (
    update_copyright_items,
    update_copyright_relations,
)
from easy_access.settings import (  # Keep DirSetting, FileSetting, SettingsFaculty for type hints
    DirSetting,
    FileSetting,
    Settings,  # Add Settings for type hint
    SettingsFaculty,
    # DEPARTMENT_MAPPING, # Will be accessed via settings
    # SETTINGS, # Will be passed as an argument
)
from easy_access.utils import File


async def load_osiris_data(settings: Settings) -> None:
    """
    Creates Courses from the osiris_data.json file.
    Staff data is added later once people data has been loaded.
    """
    await ensure_db_inited(settings)
    if not settings.files[FileSetting.OSIRIS_DATA].exists:
        logger.warning("No osiris_data.json file found; data not loaded to DB.")
        return

    try:
        with open(settings.files[FileSetting.OSIRIS_DATA].path, encoding="utf-8") as f:
            osiris_data: dict[str, dict[str, str | list[str]]] = json.load(f)
    except Exception as e:
        logger.warning(f"Error reading osiris_data.json: {e}")
        return

    course_dicts = []
    existing_course_codes = await Course().all().values("cursuscode")
    existing_course_codes = {int(c["cursuscode"]) for c in existing_course_codes}
    for course_data in osiris_data.values():
        course_code_raw = course_data.get("cursuscode", 0)
        if not isinstance(course_code_raw, int) and not isinstance(
            course_code_raw, str
        ):
            logger.warning(f"Invalid course code: {course_code_raw}. Skipping course.")
            continue
        else:
            course_code = int(course_code_raw)
        if course_code in existing_course_codes or not course_data.get("cursuscode"):
            continue
        course_dict: dict[str, str | list[str] | None] = {
            "cursuscode": course_code,
            "internal_id": int(course_data.get("internal_id")),
            "name": course_data.get("name", None),
            "short_name": course_data.get("short_name", None),
            "ec": int(round(float(course_data.get("ec", 0).replace(",", ".")), 0)),
            "programme": course_data.get("programme", None),
            "notes": course_data.get("notes", None),
            "category": course_data.get("category", None),
        }

        existing_course_codes.add(int(course_data.get("cursuscode")))

        year: str = course_data.get("year")
        if "-" in year:
            year = int(year.split("-")[0])
            course_dict["year"] = year

        faculty = await Faculty.get_or_none(abbreviation=course_data.get("faculty"))
        if faculty:
            course_dict["faculty"] = faculty

        course_dicts.append(course_dict)

    await Course.bulk_create(objects=[Course(**c) for c in course_dicts])


async def load_org_data_from_settings(settings: Settings) -> None:
    """
    From settings.university_settings.faculties, create Faculty and Programme objects.
    """
    faculties: list[SettingsFaculty] = settings.university_settings.faculties
    # first retrieve or create the university org
    university, _ = await Organization.get_or_create(
        defaults={
            "name": "University of Twente",
            "abbreviation": "UT",
            "full_abbreviation": "UT",
            "parent_organization": None,
            "hierarchy_level": 0,
        },
        abbreviation="UT",
    )
    await ensure_db_inited(settings)
    for faculty in faculties:
        faculty_obj, _ = await Faculty.get_or_create(
            defaults={
                "name": faculty.name,
                "abbreviation": faculty.abbreviation,
                "full_abbreviation": faculty.abbreviation,
                "parent_organization": university,
                "hierarchy_level": 1,
            },
            abbreviation=faculty.abbreviation,
        )
        progamme_list = []
        existing_programme_names = await Programme().all().values("name")
        existing_programme_names = {p["name"] for p in existing_programme_names}
        for programme in faculty.programmes:
            if programme.name in existing_programme_names:
                continue
            if programme.abbreviation:
                abbr = programme.abbreviation
            else:
                abbr = ""
                if "master" in programme.name.lower():
                    abbr = "M-"
                elif "bachelor" in programme.name.lower():
                    abbr = "B-"
                else:
                    abbr = "O-"

                removed_prefix = (
                    programme.name.lower()
                    .replace("bachelor", "")
                    .replace("master", "")
                    .strip()
                )
                if " " in removed_prefix:
                    parts = removed_prefix.split(" ")
                    if len(parts) == 2:
                        abbr += parts[0][0] + parts[1][0]
                    elif len(parts) >= 3:
                        abbr += parts[0][0] + parts[1][0] + parts[2][0]
                else:
                    abbr += removed_prefix[:3]

            programme_dict = {
                "name": programme.name,
                "abbreviation": abbr,
                "programme_type": programme.programme_type
                if programme.programme_type
                else None,
                "faculty": faculty_obj,
                "cluster": programme.cluster if programme.cluster else None,
            }
            progamme_list.append(programme_dict)

        await Programme.bulk_create(objects=[Programme(**p) for p in progamme_list])


async def load_person_data(settings: Settings) -> None:
    """
    Load data from the person_data.json file into the db.
    Adds Persons and Orgs, creates MissingCourses where necessary.
    """

    await ensure_db_inited(settings)
    if not settings.files[FileSetting.PERSON_DATA].exists:
        logger.warning("No person_data.json file found; data not loaded to DB.")
        return
    try:
        with open(settings.files[FileSetting.PERSON_DATA].path, encoding="utf-8") as f:
            person_data: list[dict[str, str | float | list[str]]] = json.load(f)
    except Exception as e:
        logger.warning(f"Error reading person_data.json: {e}")
        return

    existing_person_names = await Person().all().values("input_name")
    existing_person_names = {p["input_name"] for p in existing_person_names}
    for person in person_data:
        if person.get("input_name") in existing_person_names:
            continue
        person_dict = {
            "input_name": person.get("input_name").strip(),
            "main_name": person.get("main_name", None),
            "match_confidence": person.get("match_confidence", None),
            "first_name": person.get("other_names", [None])[0],
            "email": person.get("email", None),
            "faculty": await Faculty.get_or_none(
                abbreviation=person.get("faculty", None)
            ),
            "people_page_url": person.get("people_page_url", None),
        }
        for k, v in person_dict.items():
            if isinstance(v, str):
                person_dict[k] = v.strip()
        orgs: list[dict[str, str]] = person.get("orgs", [])
        orgs.sort(key=lambda x: x.get("abbr").count("-"))
        orglist = []
        for org in orgs:
            try:
                hierarchy_level = org.get("abbr").count("-") + 1
                orgname = org.get("name").strip()
                org_sole_abbr = org.get("abbr").split("-")[-1].strip()
                org_full_abbr = org.get("abbr").strip()
                org_dict = {
                    "name": orgname,
                    "abbreviation": org_sole_abbr,
                    "full_abbreviation": org_full_abbr,
                    "hierarchy_level": hierarchy_level,
                }
                org_obj, _ = await Organization.get_or_create(
                    defaults=org_dict, full_abbreviation=org_full_abbr
                )
                if hierarchy_level == 1:
                    org_obj.parent_organization, _ = await Organization.get_or_create(
                        defaults={
                            "name": "University of Twente",
                            "abbreviation": "UT",
                            "full_abbreviation": "UT",
                            "parent_organization": None,
                            "hierarchy_level": 0,
                        },
                        abbreviation="UT",
                    )
                elif hierarchy_level >= 2:
                    parent_org_abbr = org_full_abbr.split("-")[-2].strip()
                    parent_org_full_abbreviation = org_full_abbr.rsplit("-", 1)[
                        0
                    ].strip()
                    org_obj.parent_organization = await Organization.get(
                        abbreviation=parent_org_abbr,
                        full_abbreviation=parent_org_full_abbreviation,
                        hierarchy_level=hierarchy_level - 1,
                    )
                await org_obj.save()
                orglist.append(org_obj)
            except Exception as e:
                logger.warning(f"Error adding org with data {org_dict} to person: {e}")
                continue

        person = await Person.create(**person_dict)
        if orglist:
            await person.orgs.add(*orglist)


async def load_linked_persons_for_courses(settings: Settings) -> Counter:
    """
    Link the Persons to the Courses they are involved in.
    """
    await ensure_db_inited(settings)
    # load osiris data
    if not settings.files[FileSetting.OSIRIS_DATA].exists:
        logger.warning("No osiris_data.json file found; data not loaded to DB.")
        return

    try:
        with open(settings.files[FileSetting.OSIRIS_DATA].path, encoding="utf-8") as f:
            osiris_data: dict[str, dict[str, str | list[str]]] = json.load(f)
    except Exception as e:
        logger.warning(f"Error reading osiris_data.json: {e}")
        return

    counter = Counter()

    for course_data in osiris_data.values():
        course_obj = await Course.get_or_none(cursuscode=course_data.get("cursuscode"))
        if not course_obj:
            continue

        teachers = course_data.get("teachers", [])
        for teacher_name in teachers:
            teacher_obj = await Person.get_or_none(input_name=teacher_name)
            if teacher_obj:
                await CourseEmployee.get_or_create(
                    course=course_obj, person=teacher_obj, role="teacher"
                )
                counter["teachers"] += 1

        contacts = course_data.get("contacts", [])
        for contact_name in contacts:
            contact_obj = await Person.get_or_none(input_name=contact_name)
            if contact_obj:
                await CourseEmployee.get_or_create(
                    course=course_obj, person=contact_obj, role="contact"
                )
                counter["contacts"] += 1

        tutors = course_data.get("tutors", [])
        for tutor_name in tutors:
            tutor_obj = await Person.get_or_none(input_name=tutor_name)
            if tutor_obj:
                await CourseEmployee.get_or_create(
                    course=course_obj, person=tutor_obj, role="tutor"
                )
                counter["tutors"] += 1

        docenten = course_data.get("contacts", [])
        for docent_name in docenten:
            docent_obj = await Person.get_or_none(input_name=docent_name)
            if docent_obj:
                await CourseEmployee.get_or_create(
                    course=course_obj, person=docent_obj, role="docent"
                )
                counter["docenten"] += 1

        examinators = course_data.get("examinators", [])
        for examinators_name in examinators:
            examinator_obj = await Person.get_or_none(input_name=examinators_name)
            if examinator_obj:
                await CourseEmployee.get_or_create(
                    course=course_obj, person=examinator_obj, role="examinator"
                )
                counter["examinators"] += 1

        unknown_roles = course_data.get("unknown_role", [])
        for unknown_role_name in unknown_roles:
            unknown_role_obj = await Person.get_or_none(input_name=unknown_role_name)
            if unknown_role_obj:
                await CourseEmployee.get_or_create(
                    course=course_obj, person=unknown_role_obj, role="unknown_role"
                )
                counter["unknown_roles"] += 1

    return counter


async def load_base_data(settings: Settings) -> None:
    """
    Load the base data into the db: orgs, courses, persons.
    """
    await ensure_db_inited(settings)
    await create()
    try:
        faculty_count = await Faculty.all().count()
        programme_count = await Programme.all().count()
        logger.info(
            f"# of Faculties present in DB before load_org_data: {faculty_count}"
        )
        logger.info(
            f"# of Programmes present in DB before load_org_data: {programme_count}"
        )

        await load_org_data_from_settings(settings=settings)

        faculty_count = await Faculty.all().count()
        programme_count = await Programme.all().count()
        logger.success(
            f"# of Faculties present in DB after load_org_data: {faculty_count}"
        )
        logger.success(
            f"# of Programmes present in DB after load_org_data: {programme_count}"
        )

        course_count = await Course.all().count()
        logger.info(
            f"# of Courses present in DB before load_osiris_data: {course_count}"
        )

        await load_osiris_data(settings=settings)

        course_count = await Course.all().count()
        logger.success(
            f"# of Courses present in DB after load_osiris_data: {course_count}"
        )

        person_count = await Person.all().count()
        org_count = await Organization.all().count()
        missing_orgs = await MissingCourse.all().count()
        logger.info(
            f"# of Persons present in DB before load_person_data: {person_count}"
        )
        logger.info(
            f"# of Organizations present in DB before load_person_data: {org_count}"
        )
        logger.info(
            f"# of MissingCourses present in DB before load_person_data: {missing_orgs}"
        )

        await load_person_data(settings=settings)

        org_count = await Organization.all().count()
        person_count = await Person.all().count()
        missing_orgs = await MissingCourse.all().count()
        logger.success(
            f"# of Persons present in DB after load_person_data: {person_count}"
        )
        logger.success(
            f"# of Organizations present in DB after load_person_data: {org_count}"
        )
        logger.success(
            f"# of MissingCourses present in DB after load_person_data: {missing_orgs}"
        )

        results = await load_linked_persons_for_courses(settings=settings)
        logger.success(f"Overview of new relations added to Courses: {results}")
    except Exception as e:
        logger.warning(f"Error loading base data: {e}")

    await Tortoise.close_connections()


async def load_raw_copyright_data(
    settings: Settings, file: File | pl.DataFrame | None = None
) -> None:
    def _read_excel_quiet(file_path: str | Path, **kwargs) -> pl.DataFrame:
        """Read excel quietly suppressing noisy dtype-inference output."""
        try:
            with open(os.devnull, "w") as devnull:
                with (
                    contextlib.redirect_stdout(devnull),
                    contextlib.redirect_stderr(devnull),
                ):
                    return pl.read_excel(file_path, **kwargs)
        except Exception:
            return pl.read_excel(file_path, **kwargs)

    def read_copyright_export(
        settings_param: Settings, file_param: File | None = None
    ) -> pl.DataFrame:
        """
        Reads in data from the latest copyright export file in the copyright dir;
        or if a file is given, reads in that file.
        Input should be a direct export from the CopyRight tool without any changes.
        """
        try:
            if not file_param:
                logger.info(
                    f"Reading in newest Copyright Data from directory: {settings_param.dirs[DirSetting.RAW_COPYRIGHT_DATA]}"
                )
                file_param = max(
                    settings_param.dirs[DirSetting.RAW_COPYRIGHT_DATA].files,
                    key=lambda x: x.created,
                )

            logger.info(f"Reading in data from:\n            {file_param.name}\n")
            latest_file_date = file_param.created.strftime("%Y-%m-%d")
            raw_copyright_data = _read_excel_quiet(file_param.path, sheet_name=None)
            copyright_data = (
                raw_copyright_data.with_columns(pl.exclude(pl.Utf8).cast(str))
                .rename(
                    lambda col: col.replace(" ", "_")
                    .replace("#", "count_")
                    .replace("*", "x")
                    .lower()
                )
                .with_columns(
                    pl.Series(
                        "retrieved_from_copyright_on",
                        [latest_file_date] * len(raw_copyright_data),
                    ),
                    pl.Series("workflow_status", ["ToDo"] * len(raw_copyright_data)),
                    pl.col("last_change")
                    .str.replace(r"^-$", "")
                    .str.strip_chars()
                    .str.strptime(pl.Date, "%Y-%m-%d", strict=False)
                    .dt.strftime("%Y-%m-%d"),
                    pl.col("classification").str.to_lowercase(),
                    faculty=pl.col("department").replace_strict(
                        settings_param.university_settings.department_mapping,
                        default="Unmapped",
                    ),
                )
            )

            # now drop rows we definitely do not want.
            # - drop row if material_id is null, None, blank, or '-'
            # - keep rows with filetype pdf, ppt, doc, or blank ('-'/None/null/""), drop rest
            logger.info(
                f"Retrieved {len(copyright_data)} items from {file_param.name}."
            )

            copyright_data = copyright_data.filter(pl.col("material_id").is_not_null())
            copyright_data = copyright_data.filter(
                (pl.col("filetype").is_in(["pdf", "ppt", "doc", "-"]))
                | (pl.col("filetype").is_null())
            )

            logger.info(
                f"{len(copyright_data)} items remaining from {file_param.name} after filtering out missing material_ids and specific filetypes."
            )
            return copyright_data
        except FileNotFoundError as e:
            logger.warning(
                f"No files found in {settings_param.dirs[DirSetting.RAW_COPYRIGHT_DATA]}"
            )
            raise e
        except PermissionError as e:
            logger.warning(f"Permission denied to read {file_param.name}")
            raise e
        except ValueError as e:
            logger.warning(f"No file found: {file_param=}.")
            raise e

    """
    Loads in new items from a copyright export file.
    Either give a specific file to read in, or use the default (latest regular raw copyright export in raw_copyright_data dir) .
    """
    await ensure_db_inited(settings)
    error = None
    logger.info(
        f"# of items in db before loading raw items: {await CopyrightItem.all().count()}"
    )
    if file is not None:
        if isinstance(file, pl.DataFrame):
            logger.info(
                f"loading {len(file)} raw copyright items into db from dataframe."
            )
            df = file
        else:
            logger.info(f"loading raw items from {file}")
    try:
        if not isinstance(file, pl.DataFrame):
            if file is None:  # file here is the parameter of load_raw_copyright_data
                logger.info("loading raw items from most recent raw copyright export")
            df = read_copyright_export(settings_param=settings, file_param=file)

        items = standardize_dataframe(df).to_dicts()

        item_list = []

        existing_mat_ids = await CopyrightItem.all().values("material_id")
        existing_mat_ids = {int(m["material_id"]) for m in existing_mat_ids}

        logger.info(
            f"Read in {len(items)} raw copyright items. {len(existing_mat_ids)} items already in db."
        )
        num_total = len(items)
        update_list = []
        for item in items:
            if int(item.get("material_id")) in existing_mat_ids:
                update_list.append(item)
                continue
            created_item: CopyrightItem = await copyright_item_from_dict(item)
            if not created_item:
                logger.info(item)
                inp = input(
                    "Error parsing raw item. enter x to stop, anything else to continue"
                )
                if inp.lower() == "x":
                    break
                continue
            item_list.append(created_item)
            if num_total % 100 == 0:
                logger.info(f"Parsed {len(item_list)}/{num_total} items.")
    except Exception as e:
        logger.warning(f"Error loading raw items: {e}")
        error = e
    finally:
        if item_list:
            await CopyrightItem.bulk_create(objects=item_list)
            await load_llm_classifications(settings=settings)
            await update_copyright_relations(settings=settings)
        logger.success(
            f"# of items in db after loading raw items: {await CopyrightItem.all().count()}"
        )
        if error:
            await Tortoise.close_connections()
            raise error

    if update_list:
        logger.info(f"comparing {len(update_list)} items with items in db for updates.")
        await update_copyright_items(settings, update_list)

    await Tortoise.close_connections()


async def load_pdfs(settings: Settings) -> None:
    """
    Load pdfs from the pdfs dir into the db.
    """

    await ensure_db_inited(settings)
    pdf_file_list: list[File] = settings.dirs[DirSetting.PDF_DOWNLOADS].files
    try:
        pdf_files: dict[str, File] = {
            file.name.split("_")[0]: file
            for file in pdf_file_list
            if file.name.endswith(".pdf") and "_" in file.name
        }
        more_files: dict[str, File] = {
            file.name.split(".")[0]: file
            for file in pdf_file_list
            if file.name.endswith(".pdf") and "_" not in file.name
        }
        pdf_files.update(more_files)
    except Exception as e:
        logger.warning(f"Error loading pdf files: {e}")
        return

    if not pdf_files:
        logger.warning("No PDF files found; data not loaded to DB.")
        return
    existing_pdfs_mat_ids = await PDF.all().values("material_id")

    pdf_files = {
        k: v
        for k, v in pdf_files.items()
        if int(k) not in {int(p["material_id"]) for p in existing_pdfs_mat_ids}
    }
    pdf_dicts = []
    for mat_id, pdf_file in pdf_files.items():
        related_item = await CopyrightItem.get_or_none(material_id=int(mat_id))
        if not related_item:
            logger.warning(
                f"No related item found for pdf with material_id {mat_id}. Skipping."
            )
            continue
        pdf_dict = {
            "material_id": int(mat_id),
            "current_file_name": pdf_file.name,
            "original_file_name": related_item.filename,
            "original_page_count": related_item.pagecount,
            # file exists on disk, so mark download as attempted and succeeded
            "download_attempted": True,
            "download_succeeded": True,
        }
        pdf_dicts.append(pdf_dict)

    logger.info(f"Creating {len(pdf_dicts)} new PDF objects in DB.")
    await PDF.bulk_create(objects=[PDF(**p) for p in pdf_dicts])

    await Tortoise.close_connections()
