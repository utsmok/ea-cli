"""
functions to ingest new data into the database
"""

import polars as pl
from loguru import logger

from easy_access.db.base import create, ensure_db_inited
from easy_access.db.compat import (
    all_values,
    bulk_create,
    count,
    get_or_create,
    get_or_none,
)
from easy_access.db.models import (
    PDF,
    CopyrightItem,
    Faculty,
    Organization,
    Programme,
    StagedCopyrightItem,
    StagedFacultyUpdate,
)
from easy_access.db.session import shutdown_db
from easy_access.settings import DirSetting, Settings, SettingsFaculty
from easy_access.utils import File, standardize_dataframe


async def load_org_data_from_settings(settings: Settings) -> None:
    """
    From settings.university_settings.faculties, create Faculty and Programme objects.
    """
    faculties: list[SettingsFaculty] = settings.university_settings.faculties
    # first retrieve or create the university org
    university, _ = await get_or_create(
        Organization,
        abbreviation="UT",
        defaults={
            "name": "University of Twente",
            "abbreviation": "UT",
            "full_abbreviation": "UT",
            "parent_organization": None,
            "hierarchy_level": 0,
        },
    )
    await ensure_db_inited(settings)
    for faculty in faculties:
        faculty_obj, _ = await get_or_create(
            Faculty,
            abbreviation=faculty.abbreviation,
            defaults={
                "name": faculty.name,
                "abbreviation": faculty.abbreviation,
                "full_abbreviation": faculty.abbreviation,
                "parent_organization": university,
                "hierarchy_level": 1,
            },
        )
        progamme_list = []
        existing_programme_names = await all_values(Programme, "name")
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

        await bulk_create(Programme, rows=progamme_list)


async def load_base_data(settings: Settings) -> None:
    """
    Load the base data into the db: orgs, courses, persons.
    """
    await ensure_db_inited(settings)
    await create()
    try:
        await load_org_data_from_settings(settings=settings)

        faculty_count = await count(Faculty)
        programme_count = await count(Programme)
        logger.success(
            f"# of Faculties present in DB after load_org_data: {faculty_count}"
        )
        logger.success(
            f"# of Programmes present in DB after load_org_data: {programme_count}"
        )

    except Exception as e:
        logger.warning(f"Error loading base data: {e}")

    await shutdown_db()


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
    existing_pdfs_mat_ids = await all_values(PDF, "material_id")

    pdf_files = {
        k: v
        for k, v in pdf_files.items()
        if int(k) not in {int(p["material_id"]) for p in existing_pdfs_mat_ids}
    }
    pdf_dicts = []
    for mat_id, pdf_file in pdf_files.items():
        related_item = await get_or_none(CopyrightItem, material_id=int(mat_id))
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
    await bulk_create(PDF, rows=pdf_dicts)

    await shutdown_db()


async def load_raw_copyright_data_to_staging(
    settings: Settings, data: pl.DataFrame
) -> None:
    """
    Loads raw copyright data into the staging table.
    """
    await ensure_db_inited(settings)
    items = standardize_dataframe(data).to_dicts()

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

    await bulk_create(
        StagedCopyrightItem,
        rows=items,
        on_conflict=["material_id"],
        update_fields=update_fields,
    )


async def load_faculty_updates_to_staging(
    settings: Settings, data: pl.DataFrame
) -> None:
    """
    Loads faculty updates into the staging table.
    """
    await ensure_db_inited(settings)
    data = data.select(
        ["material_id", "manual_classification", "remarks", "workflow_status"]
    )
    items = standardize_dataframe(data).to_dicts()
    await bulk_create(
        StagedFacultyUpdate,
        rows=items,
        on_conflict=["material_id"],
        update_fields=["manual_classification", "remarks", "workflow_status"],
    )
