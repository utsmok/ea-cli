"""
Functions to retrieve data from the database.
"""

import json
import traceback
from collections.abc import Iterable
from time import time
from typing import Any, LiteralString

import polars as pl

# from easy_access.settings import SETTINGS # Will be passed as an argument
from loguru import logger
from sqlalchemy import Engine, text
from tortoise import Tortoise
from tortoise.expressions import Q

from easy_access.db.base import ensure_db_inited, init_engine
from easy_access.db.models import PDF, CopyrightItem, ItemUpdate
from easy_access.settings import Settings  # Import Settings for type hint

engine: Engine | None = None


def retrieve_copyright_items(settings: Settings) -> pl.DataFrame:  # Added settings
    """
    Retrieve all copyright items currently in db
    returns a flat dataframe with the core fields
    """
    full_start = time()
    global engine
    if not engine:
        engine = init_engine(settings=settings)  # Pass settings

    col_order: set[str] = set(
        settings.data_settings.raw_data_col_order
    )  # Use passed settings

    if "google_search_file" in col_order:
        col_order.remove("google_search_file")
    if "type" in col_order:
        col_order.remove("type")
    if "faculty" in col_order:
        col_order.remove("faculty")
        select_cols = {*col_order, "faculty_id AS faculty"}
    else:
        select_cols = col_order

    query: str = "SELECT " + ", ".join(select_cols) + " FROM copyright_data cd"
    query_start = time()
    df: pl.DataFrame = pl.read_database(
        query=query, connection=engine.connect(), infer_schema_length=None
    )
    end = time()
    logger.info(f"db.retrieve.retrieve_copyright_items returned {len(df)} rows")
    logger.info(f"query took {end - query_start} seconds")
    logger.info(f"full function took {end - full_start} seconds")
    return df


def retrieve_duplicate_copyright_items(
    settings: Settings,
) -> pl.DataFrame:  # Added settings
    """
    for copyright_items with duplicates, find the replacing material id
    returns a dataframe with 'material_id', 'is_duplicate', and 'replacement_id' columns
    """
    global engine
    if not engine:
        engine = init_engine(settings=settings)  # Pass settings

    query: str = """
        SELECT material_id, is_duplicate, replacement_id
        FROM copyright_data cd
        WHERE is_duplicate IS TRUE
    """
    df: pl.DataFrame = pl.read_database(
        query=query, connection=engine.connect(), infer_schema_length=None
    )
    return df


def get_valid_faculties(settings: Settings) -> set[str]:  # Added settings
    """Retrieves the set of valid faculty abbreviations from the database."""
    global engine
    if not engine:
        engine = init_engine(settings=settings)  # Pass settings
    with engine.connect() as conn:
        result = conn.execute(text("SELECT abbreviation FROM faculty"))
        return {row[0] for row in result.fetchall()}


def retrieve_full_data(
    selected_material_ids: Iterable[int] | None = None,
    selected_faculties: Iterable[str] | str | None = None,
    excluded_material_ids: Iterable[int] | None = None,
    settings: Settings
    | None = None,  # Added settings, optional for now if not always available
) -> pl.DataFrame:
    """
    Retrieves copyright items, enriched with related data, with optional filtering.

    Args:
        selected_material_ids: Optional iterable of material IDs to select (ONLY these, drop rest).
        selected_faculties: Optional string or list of strings (faculty abbreviations). ONLY return items with these faculties.
        excluded_material_ids: Optional iterable of material IDs to exclude. Excluded_material_ids take precedence over selected_material_ids.
        settings: The application settings.

    Returns:
        A Polars DataFrame.
    """
    global engine
    if not settings:
        # This case should ideally be handled by ensuring settings are always passed.
        # For now, let's assume a global SETTINGS might be fallback, or raise error.
        # from easy_access.settings import SETTINGS as global_settings # Avoid if possible
        # settings = global_settings
        raise ValueError("Settings must be provided to retrieve_full_data")

    if not engine:
        engine = init_engine(settings=settings)  # Pass settings
    valid_faculties = get_valid_faculties(settings=settings)  # Pass settings
    material_join_clause = ""
    faculty_where_clause = ""
    material_exclusion_clause = ""
    with engine.connect() as conn:
        if selected_material_ids is not None:
            if not selected_material_ids:
                logger.warning(
                    "retrieve_full_data received an empty collection of selected_material_ids.  Returning empty DataFrame."
                )
                return pl.DataFrame()
            conn.execute(text("DROP TABLE IF EXISTS temp_material_ids;"))
            conn.execute(
                text("CREATE TEMP TABLE temp_material_ids (material_id INTEGER);")
            )
            conn.execute(
                text(
                    "INSERT INTO temp_material_ids (material_id) VALUES (:material_id)"
                ),
                [{"material_id": mat_id} for mat_id in selected_material_ids],
            )
            material_join_clause = (
                "INNER JOIN temp_material_ids tmid ON cd.material_id = tmid.material_id"
            )

        if excluded_material_ids is not None:
            excluded_material_ids_list = list(excluded_material_ids)
            if excluded_material_ids_list:  # Only add clause if the list is not empty
                excluded_ids_string = ", ".join(map(str, excluded_material_ids_list))
                material_exclusion_clause = (
                    f"AND cd.material_id NOT IN ({excluded_ids_string})"
                )

        if selected_faculties:
            if isinstance(selected_faculties, str):
                selected_faculties = [selected_faculties]  # Make it a list

            # Validate faculties
            invalid_faculties = set(selected_faculties) - valid_faculties
            if invalid_faculties:
                raise ValueError(f"Invalid faculty abbreviations: {invalid_faculties}")

            faculties_string = "', '".join(selected_faculties)  # Escape for SQL
            faculty_where_clause = f"AND cd.faculty_id IN ('{faculties_string}')"

        query: str = f"""
            WITH CourseDataAggregated AS (
                SELECT
                    cdcd.copyright_data_id,
                    (SELECT GROUP_CONCAT(cursuscode, ' | ') FROM (SELECT DISTINCT CAST(cd.cursuscode AS TEXT) as cursuscode FROM course_data cd WHERE cd.cursuscode = cdcd.course_id)) AS cursuscodes,
                    (SELECT GROUP_CONCAT(programme, ' | ') FROM (SELECT DISTINCT cd.programme FROM course_data cd WHERE cd.cursuscode = cdcd.course_id)) AS programmes,
                    (SELECT GROUP_CONCAT(name, ' | ') FROM (SELECT DISTINCT cd.name FROM course_data cd WHERE cd.cursuscode = cdcd.course_id)) AS course_names
                FROM copyright_data_course_data cdcd
            )
            SELECT
                cd.*,
                llm.allowed_usage as llm_allowed_usage,
                llm.allowed_usage_reasoning as llm_allowed_usage_reason,
                llm.copyright_status as llm_copyright,
                llm.copyright_classification_reason as llm_copyright_reason,
                llm.item_type as llm_item_type,
                llm.remarks as llm_remarks,
                llm.item_title as llm_title,
                llm.copyright_holder as llm_copyright_holder,
                llm.publisher_name as llm_publisher,
                llm.isbn as llm_isbn,
                llm.doi as llm_doi,
                llm.source_url as llm_source_url,
                llm.license as llm_license,
                llm.author_names as llm_authors,
                cda.cursuscodes as cursuscodes,
                cda.programmes as programmes,
                cda.course_names as course_names,
                (
                    SELECT GROUP_CONCAT(course_contacts_names, ' | ')
                    FROM (
                        SELECT DISTINCT pd.main_name as course_contacts_names
                        FROM course_employee ce
                        JOIN person_data pd ON ce.person_id = pd.id
                        WHERE ce.role = 'contact' AND ce.course_id = cdcd.course_id
                    )
                ) AS course_contacts_names,
                (
                    SELECT GROUP_CONCAT(course_contacts_emails, ' | ')
                    FROM (
                        SELECT DISTINCT pd.email as course_contacts_emails
                        FROM course_employee ce
                        JOIN person_data pd ON ce.person_id = pd.id
                        WHERE ce.role = 'contact' AND ce.course_id = cdcd.course_id
                    )
                ) AS course_contacts_emails,
                (
                    SELECT GROUP_CONCAT(faculty_abbreviations, ' | ')
                    FROM (
                        SELECT DISTINCT f.abbreviation as faculty_abbreviations
                        FROM course_employee ce
                        JOIN person_data pd ON ce.person_id = pd.id
                        LEFT JOIN faculty f ON pd.faculty_id = f.abbreviation
                        WHERE ce.role = 'contact' AND ce.course_id = cdcd.course_id
                    )
                ) AS course_contacts_faculties,
                (
                    SELECT GROUP_CONCAT(org_full_abbreviations, ' | ')
                    FROM (
                        SELECT DISTINCT org.full_abbreviation as org_full_abbreviations
                        FROM course_employee ce
                        JOIN person_data pd ON ce.person_id = pd.id
                        LEFT JOIN person_data_organization_data pdod ON pd.id = pdod.person_data_id
                        LEFT JOIN organization_data org ON pdod.organization_id = org.id
                        WHERE ce.role = 'contact' AND ce.course_id = cdcd.course_id
                    )
                ) AS course_contacts_organizations
            FROM copyright_data cd
            {material_join_clause}
            LEFT JOIN llm_classification_data llm ON cd.llm_classification_id = llm.id
            LEFT JOIN CourseDataAggregated cda ON cd.material_id = cda.copyright_data_id
            LEFT JOIN copyright_data_course_data cdcd ON cd.material_id = cdcd.copyright_data_id
            WHERE 1=1  -- Placeholder for easier AND clause addition
            {faculty_where_clause}
            {material_exclusion_clause}

            """
        df = pl.read_database(query=query, connection=conn, infer_schema_length=None)

    df = df.drop(
        [
            "llm_classification_id",
            "created_at",
            "modified_at",
            "possible_fine",
            "infringement",
        ]
    ).rename(mapping={"faculty_id": "faculty"})

    if df.is_empty():
        return df
    if df["material_id"].is_null().all():
        return pl.DataFrame()

    # drop nulcols from llm cols
    llm_cols = ["llm_isbn", "llm_doi", "llm_source_url", "llm_license", "llm_authors"]
    droplist = []
    # iterate over llm cols
    # if col is null (dtype=null or all values are null), drop it
    # if all cells in col are "[]", drop it
    for col in llm_cols:
        if (
            (df[col].is_null().all())
            or (df[col].dtype == pl.Null)
            or (df[col].eq("[]").all())
        ):
            df = df.drop(col)
            droplist.append(col)

    select_cols = [col for col in llm_cols if col not in droplist]

    if select_cols:
        for colname in select_cols:
            try:
                df = df.with_columns(
                    pl.col(name=colname)
                    .str.json_decode(infer_schema_length=None)
                    .list.join(separator=" | ")
                )
            except Exception as e:
                logger.warning(
                    f"Error {e} while processing {colname} with type {df[colname].dtype} in retrieve_full_data"
                )
                logger.debug(f"Column: {colname}")
                logger.debug(f"Length: {len(df[colname])}")
                logger.debug(f"Column values: {df[colname]}")
                logger.debug(f"Dtype: {df[colname].dtype}")

    # add droplist cols back in with empty values
    if droplist:
        for (
            _col
        ) in droplist:  # Renamed col to _col as it's not used in the loop body directly
            df = df.with_columns(pl.lit(None).alias(_col))  # Use _col in alias

    return df


def get_llm_classification_schema() -> dict[str, type]:
    """Defines the expected schema for llm_classification_data."""
    return {
        "allowed_usage_llm": pl.Categorical,
        "allowed_usage_reasoning_llm": pl.Utf8,
        "copyright_status_llm": pl.Categorical,
        "copyright_classification_reason_llm": pl.Utf8,
        "item_type_llm": pl.Categorical,
        "item_type_classification_reason_llm": pl.Utf8,
        "publisher_name_llm": pl.Utf8,
        "copyright_holder_llm": pl.Utf8,
        "item_title_llm": pl.Utf8,
        "pdf_page_count_llm": pl.Int64,
        "remarks_llm": pl.Utf8,
        "author_names_llm": pl.List(pl.Utf8),
        "doi_llm": pl.List(pl.Utf8),
        "isbn_llm": pl.List(pl.Utf8),
        "source_url_llm": pl.List(pl.Utf8),
        "license_llm": pl.List(pl.Utf8),
        "topic_llm": pl.List(pl.Utf8),
        "material_id": pl.Int64,
    }  # type: ignore


def retrieve_llm_classifications(
    selected_material_ids: Iterable[int] | None = None,
    settings: Settings | None = None,  # Added settings
) -> pl.DataFrame:
    global engine
    if not settings:
        raise ValueError("Settings must be provided to retrieve_llm_classifications")
    if not engine:
        engine = init_engine(settings=settings)  # Pass settings
    material_join_clause: str = ""
    with engine.connect() as conn:
        # print all table names
        if selected_material_ids is not None:
            if not selected_material_ids:
                logger.warning(
                    "retrieve_llm_classifications received an empty collection of selected_material_ids.  Returning empty DataFrame."
                )
                return pl.DataFrame(schema=get_llm_classification_schema())
            conn.execute(text("DROP TABLE IF EXISTS temp_material_ids;"))
            conn.execute(
                text("CREATE TEMP TABLE temp_material_ids (material_id INTEGER);")
            )
            conn.execute(
                text(
                    "INSERT INTO temp_material_ids (material_id) VALUES (:material_id)"
                ),
                [{"material_id": mat_id} for mat_id in selected_material_ids],
            )
            material_join_clause = "INNER JOIN temp_material_ids tmid ON llm.used_material_id = tmid.material_id"

        query: LiteralString = f"""
        SELECT
            llm.*
        FROM llm_classification_data llm
        {material_join_clause}
        """

        df = pl.read_database(query=query, connection=conn)
        df = df.drop(["created_at", "modified_at", "id"]).rename(
            mapping={
                "allowed_usage": "allowed_usage_llm",
                "allowed_usage_reasoning": "allowed_usage_reasoning_llm",
                "copyright_status": "copyright_status_llm",
                "copyright_classification_reason": "copyright_classification_reason_llm",
                "item_type": "item_type_llm",
                "item_type_classification_reason": "item_type_classification_reason_llm",
                "remarks": "remarks_llm",
                "author_names": "author_names_llm",
                "item_title": "item_title_llm",
                "publisher_name": "publisher_name_llm",
                "copyright_holder": "copyright_holder_llm",
                "doi": "doi_llm",
                "isbn": "isbn_llm",
                "source_url": "source_url_llm",
                "license": "license_llm",
                "topic": "topic_llm",
                "pdf_page_count": "pdf_page_count_llm",
                "used_material_id": "material_id",
            }
        )
        if df.is_empty():
            return pl.DataFrame()
        if df["material_id"].is_null().all():
            return pl.DataFrame()

        return df.with_columns(
            [
                pl.col(name=colname)
                .str.json_decode(infer_schema_length=None)
                .fill_null(pl.lit([]))
                .list.eval(pl.element().cast(pl.Utf8))
                .list.join(separator=" | ")
                for colname in [
                    "isbn_llm",
                    "doi_llm",
                    "source_url_llm",
                    "license_llm",
                    "author_names_llm",
                    "topic_llm",
                ]
            ]
        )


def retrieve_osiris_data(
    material_ids: list[int] | int, settings: Settings | None = None
) -> list[dict[str, Any]]:  # Added settings
    """
    Retrieves copyright data and richly nested related data (faculty, courses,
    persons, organizations) for the given material IDs using SQL JSON functions.

    Args:
        material_ids: A list of material IDs to retrieve data for.
        settings: The application settings.

    Returns:
        A list of nested dictionaries, where each dictionary represents one
        copyright item and its related data. Returns an empty list if
        material_ids is empty or no data is found.
    """
    global engine
    if not settings:
        raise ValueError("Settings must be provided to retrieve_osiris_data")
    if not engine:
        engine = init_engine(settings=settings)  # Pass settings
    if not material_ids:
        logger.warning("No material IDs provided. Returning empty list.")
        return []
    if not isinstance(material_ids, Iterable):
        material_ids = [material_ids]  # Convert to list if not already
    if len(material_ids) == 1:
        mat_id_query = f"WHERE cd.material_id = {material_ids[0]}"
    else:
        mat_id_query = f"WHERE cd.material_id IN ({', '.join(map(str, material_ids))})"

    query: str = f"""
WITH RECURSIVE OrgHierarchyUp (id, name, abbreviation, parent_organization_id, path_ids, path_abbrs) AS (
      -- Base case: Start with all organizations
      SELECT id, name, abbreviation, parent_organization_id,
             CAST(id AS TEXT), CAST(abbreviation AS TEXT)
      FROM organization_data
      -- Removed WHERE clause to handle orgs with NULL parent_organization_id correctly in base case

      UNION ALL

      -- Recursive step: Go up one level (Join child's parent_id to parent's id)
      SELECT
        child.id, child.name, child.abbreviation, parent.parent_organization_id,
        parent.id || '/' || child.path_ids,
        parent.abbreviation || '/' || child.path_abbrs -- Build path bottom-up
      FROM organization_data parent -- This should be the parent
      JOIN OrgHierarchyUp child ON child.parent_organization_id = parent.id -- Join condition connects child UP to parent
      -- No WHERE clause needed here, recursion stops naturally when parent.id has no match (or parent_organization_id is NULL in the parent)
),
-- Corrected FullOrgPaths CTE using standard SQL ROW_NUMBER()
FullOrgPaths AS (
  SELECT id, path_ids, path_abbrs
  FROM (
      SELECT
        id,
        path_ids,
        path_abbrs,
        ROW_NUMBER() OVER (PARTITION BY id ORDER BY LENGTH(path_ids) DESC) as rn
      FROM
        OrgHierarchyUp
  ) AS RankedPaths
  WHERE rn = 1
),
-- Pre-aggregate organizations linked to persons, including hierarchy info
PersonOrgs AS (
  SELECT
    pdod.person_data_id,
    JSON_GROUP_ARRAY(
      JSON_OBJECT(
        'id', org.id,
        'name', org.name,
        'abbreviation', org.abbreviation,
        'full_abbreviation', org.full_abbreviation,
        'hierarchy_level', org.hierarchy_level,
        'parent_organization_id', org.parent_organization_id,
        'full_parent_abbreviations', fop.path_abbrs -- Include the full path of abbreviations
      )
    ) AS organizations_json
  FROM person_data_organization_data pdod
  JOIN organization_data org ON pdod.organization_id = org.id
  LEFT JOIN FullOrgPaths fop ON org.id = fop.id -- Join the full path CTE
  GROUP BY pdod.person_data_id
),
-- Pre-aggregate persons linked to courses
CoursePersons AS (
  SELECT
    ce.course_id,
    JSON_GROUP_ARRAY(
      JSON_OBJECT(
        'id', p.id,
        'main_name', p.main_name,
        'email', p.email,
        'first_name', p.first_name,
        'people_page_url', p.people_page_url,
        'faculty_id', p.faculty_id,
        'role', ce.role,
        'organizations', JSON(po.organizations_json) -- Embed organizations
      ) ORDER BY p.main_name
    ) AS persons_json
  FROM course_employee ce
  JOIN person_data p ON ce.person_id = p.id
  LEFT JOIN PersonOrgs po ON p.id = po.person_data_id
  GROUP BY ce.course_id
),
-- Pre-aggregate courses linked to copyright items
CopyrightCourses AS (
  SELECT
    cdcd.copyright_data_id,
    JSON_GROUP_ARRAY(
       JSON_OBJECT(
        'cursuscode', crs.cursuscode,
        'internal_id', crs.internal_id,
        'name', crs.name,
        'short_name', crs.short_name,
        'year', crs.year,
        'programme', crs.programme,
        'ec', crs.ec,
        'faculty_id', crs.faculty_id,
        'persons', JSON(cp.persons_json)
       ) ORDER BY crs.name
    ) AS courses_json
  FROM copyright_data_course_data cdcd
  JOIN course_data crs ON cdcd.course_id = crs.cursuscode
  LEFT JOIN CoursePersons cp ON crs.cursuscode = cp.course_id
  GROUP BY cdcd.copyright_data_id
)
-- Final Select statement
SELECT
  cd.*, -- Select all from copyright_data
  (
    SELECT JSON_OBJECT(
             'abbreviation', f.abbreviation,
             'name', f.name,
             'full_abbreviation', f.full_abbreviation
           )
    FROM faculty f
    WHERE f.abbreviation = cd.faculty_id
  ) AS faculty_data,
  JSON(cc.courses_json) AS courses
FROM copyright_data cd
LEFT JOIN CopyrightCourses cc ON cd.material_id = cc.copyright_data_id
{mat_id_query}

    """
    results = []
    try:
        with engine.connect() as conn:
            db_result = conn.execute(text(query)).fetchall()
            # ... (rest of the JSON parsing logic remains the same) ...
            for row_mapping in db_result:
                row_mapping = row_mapping._mapping
                item_dict = dict(row_mapping)  # Convert Row to dict
                # Parse top-level JSON
                for key in ["faculty_data", "courses"]:
                    json_string = item_dict.get(key)
                    if isinstance(json_string, str):
                        try:
                            item_dict[key] = json.loads(json_string)
                        except json.JSONDecodeError:
                            logger.warning(
                                f"Warning: Could not decode JSON for key '{key}' in material_id {item_dict.get('material_id')}. Value: {json_string}"
                            )
                            item_dict[key] = None
                    elif json_string is None:
                        # Handle case where subquery returned NULL (e.g., no courses)
                        item_dict[key] = [] if key == "courses" else None
                if item_dict.get("courses"):
                    for course in item_dict["courses"]:
                        if (
                            course
                            and "persons" in course
                            and isinstance(course["persons"], str)
                        ):
                            try:
                                course["persons"] = json.loads(course["persons"])
                                if course.get("persons"):
                                    for person in course["persons"]:
                                        if (
                                            person
                                            and "organizations" in person
                                            and isinstance(person["organizations"], str)
                                        ):
                                            try:
                                                person["organizations"] = json.loads(
                                                    person["organizations"]
                                                )
                                            except json.JSONDecodeError:
                                                person["organizations"] = []
                                        elif person and "organizations" not in person:
                                            person["organizations"] = []
                            except json.JSONDecodeError:
                                course["persons"] = []
                        elif course and "persons" not in course:
                            course["persons"] = []
                results.append(item_dict)

    except Exception as e:
        logger.error(f"Database query failed: {e}")
        logger.error(traceback.format_exc())
    finally:
        pass

    return results


async def retrieve_item_history(
    material_ids: list[int], settings: Settings | None = None
) -> list[ItemUpdate]:  # Added settings
    """
    Retrieves the history of changes for the given material IDs.

    Args:
        material_ids: A list of material IDs to retrieve history for.
        settings: The application settings.

    Returns:
        A list of dictionaries, where each dictionary represents one history entry.
    """
    if not settings:
        raise ValueError("Settings must be provided to retrieve_item_history")
    await ensure_db_inited(settings)

    if not material_ids:
        logger.warning("No material IDs provided. Returning empty list.")
        return []

    if not isinstance(material_ids, Iterable):
        material_ids = [material_ids]  # Convert to list if not already

    # get all ItemUpdate instances with material_id in material_ids

    items = await ItemUpdate().filter(material_id__in=material_ids).all()
    await Tortoise.close_connections()
    return items


async def retrieve_unmarked_deleted_items(settings: Settings) -> list[CopyrightItem]:
    """
    for each PDF in the db that has been marked with 'download_succeeded'==False (NOTE: -not- is None!),
    retrieve the corresponding copyright item using `material_id` (pk for both).

    Return all copyrightitems that DO NOT have the status `deleted` but failed to download.
    These are (probably) actually deleted, but aren't marked as such.
    """
    global engine
    if not settings:
        raise ValueError("Settings must be provided to retrieve_failed_downloads")
    # Ensure Tortoise ORM is initialized before using ORM models
    await ensure_db_inited(settings)

    if not engine:
        engine = init_engine(settings=settings)  # Pass settings

    # Retrieve material_ids as a flat list of ints so it can be used in __in filters
    deleted_pdfs = await PDF.filter(download_succeeded=False).values()
    logger.debug(deleted_pdfs[0:5])
    logger.info(
        f'Retrieved {len(deleted_pdfs)} PDFs with "download_succeeded" set to False'
    )
    logger.debug(deleted_pdfs)
    copyright_items = await CopyrightItem.filter(
        Q(material_id__in=deleted_pdfs)
    ).values()
    logger.debug([item.get("status") for item in copyright_items])
    logger.info(
        f"Retrieved {len(copyright_items)} copyright items associated with non-downloadable PDFs"
    )
    await Tortoise.close_connections()
    return copyright_items
