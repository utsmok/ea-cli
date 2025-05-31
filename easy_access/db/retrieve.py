"""
This module provides functions to retrieve data from the database, primarily
transforming raw table data into Polars DataFrames for use in the application.
It includes functions for fetching copyright items, their full enriched details,
LLM classifications, Osiris data, and item history.

It uses a global SQLAlchemy engine for Polars' direct database reads for
synchronous functions, and Tortoise ORM for asynchronous operations.
"""

import json
import logging
import traceback
from collections.abc import Iterable
from time import time
from typing import Any, LiteralString # Removed Literal as it's >=3.9, using LiteralString for SQL

import polars as pl
from sqlalchemy import Engine, text # Engine for type hint
from tortoise import Tortoise # For async retrieve_item_history

from easy_access.db.base import init as init_tortoise, init_engine # init_tortoise for async
from easy_access.db.models import ItemUpdate # For retrieve_item_history return type
from easy_access.settings import SETTINGS

engine: Engine | None = None # Global SQLAlchemy engine for polars sync functions
logger = logging.getLogger(__name__)


def retrieve_copyright_items() -> pl.DataFrame:
    """
    Retrieves all core copyright items from the database.

    This function fetches data from the `copyright_data` table, performing
    some minor transformations like aliasing `faculty_id` to `faculty`.
    It uses Polars for efficient database reading via an SQLAlchemy engine.

    Returns:
        pl.DataFrame: A Polars DataFrame containing the copyright items.
                      Returns an empty DataFrame if no items are found.
    """
    full_start_time = time()
    global engine
    if not engine: # Initialize engine if not already done
        engine = init_engine()

    # Define the set of columns to select, based on settings.
    # Exclude specific columns like 'google_search_file' and 'type'.
    # Alias 'faculty_id' to 'faculty' for consistency.
    col_order: set[str] = set(SETTINGS.data_settings.raw_data_col_order)

    # Modify col_order in place or create new set for select_cols
    select_cols_str_list: list[str] = []
    if "google_search_file" in col_order: col_order.remove("google_search_file")
    if "type" in col_order: col_order.remove("type")

    final_select_cols: set[str] = set()
    for col in col_order:
        if col == "faculty": # Assuming 'faculty' in raw_data_col_order means 'faculty_id' from DB
            final_select_cols.add("cd.faculty_id AS faculty")
        else:
            final_select_cols.add(f"cd.{col}") # Qualify with alias

    # Fallback if 'faculty' was not in raw_data_col_order but is expected
    if "cd.faculty_id AS faculty" not in final_select_cols and "faculty" not in final_select_cols:
         # Check if 'faculty_id' is in the model to decide if we should add it
         # For now, let's assume if 'faculty' is desired, 'faculty_id AS faculty' is the way
         pass # Or add it explicitly if it's always needed: select_cols.add("cd.faculty_id AS faculty")


    query: LiteralString = f"SELECT {', '.join(final_select_cols)} FROM copyright_data cd"

    df: pl.DataFrame = pl.DataFrame() # Initialize empty DataFrame
    try:
        query_start_time = time()
        # Ensure engine is not None before connecting
        if engine:
            with engine.connect() as connection:
                df = pl.read_database(query=query, connection=connection, infer_schema_length=None)
            query_duration = time() - query_start_time
            logger.debug(f"Query for retrieve_copyright_items took {query_duration:.4f} seconds.")
        else:
            logger.error("SQLAlchemy engine not initialized for retrieve_copyright_items.")

    except Exception as e:
        logger.error(f"Error retrieving copyright items: {e}")
        logger.error(traceback.format_exc())

    full_duration = time() - full_start_time
    logger.info(f"retrieve_copyright_items returned {len(df)} rows in {full_duration:.4f} seconds.")
    return df


def retrieve_duplicate_copyright_items() -> pl.DataFrame:
    """
    Retrieves information about copyright items marked as duplicates.

    Fetches `material_id`, `is_duplicate`, and `replacement_id` for items
    where `is_duplicate` is True.

    Returns:
        pl.DataFrame: A DataFrame with columns for duplicate item identification.
                      Returns an empty DataFrame if no duplicates are found or an error occurs.
    """
    global engine
    if not engine:
        engine = init_engine()

    query: LiteralString = """
        SELECT cd.material_id, cd.is_duplicate, cd.replacement_id
        FROM copyright_data cd
        WHERE cd.is_duplicate = TRUE
    """
    df: pl.DataFrame = pl.DataFrame()
    try:
        if engine:
            with engine.connect() as connection:
                df = pl.read_database(query=query, connection=connection, infer_schema_length=None)
            logger.info(f"Retrieved {len(df)} duplicate copyright items.")
        else:
            logger.error("SQLAlchemy engine not initialized for retrieve_duplicate_copyright_items.")
    except Exception as e:
        logger.error(f"Error retrieving duplicate copyright items: {e}")
        logger.error(traceback.format_exc())
    return df


def get_valid_faculties() -> set[str]:
    """
    Retrieves the set of valid faculty abbreviations from the 'faculty' table.

    Returns:
        set[str]: A set of unique faculty abbreviations. Returns an empty set on error.
    """
    global engine
    if not engine:
        engine = init_engine()

    valid_faculties: set[str] = set()
    try:
        if engine:
            with engine.connect() as conn:
                result = conn.execute(text("SELECT f.abbreviation FROM faculty f")) # Assuming table name is 'faculty'
                valid_faculties = {row[0] for row in result.fetchall() if row[0] is not None}
            logger.debug(f"Retrieved {len(valid_faculties)} valid faculties: {valid_faculties}")
        else:
            logger.error("SQLAlchemy engine not initialized for get_valid_faculties.")
    except Exception as e:
        logger.error(f"Error retrieving valid faculties: {e}")
        logger.error(traceback.format_exc())
    return valid_faculties


def retrieve_full_data(
    selected_material_ids: Iterable[int] | None = None,
    selected_faculties: Iterable[str] | str | None = None,
    excluded_material_ids: Iterable[int] | None = None,
) -> pl.DataFrame:
    """
    Retrieves copyright items enriched with related data (LLM classifications, courses, contacts),
    with optional filtering by material IDs and/or faculties.

    Args:
        selected_material_ids: Optional. Select ONLY these material IDs.
        selected_faculties: Optional. Select items ONLY from these faculty abbreviations.
        excluded_material_ids: Optional. Exclude these material IDs. Takes precedence.

    Returns:
        pl.DataFrame: An enriched DataFrame. Returns an empty DataFrame on error or no data.
    """
    global engine
    if not engine:
        engine = init_engine()

    df: pl.DataFrame = pl.DataFrame()
    if not engine:
        logger.error("SQLAlchemy engine not initialized for retrieve_full_data.")
        return df

    valid_faculties = get_valid_faculties()
    material_join_clause: str = ""
    faculty_where_clause: str = ""
    material_exclusion_clause: str = ""

    try:
        with engine.connect() as conn: # Ensure connection is managed
            if selected_material_ids:
                selected_ids_list = list(selected_material_ids)
                if not selected_ids_list:
                    logger.warning("retrieve_full_data received empty selected_material_ids. Returning empty DataFrame.")
                    return pl.DataFrame()
                # Use temporary table for large lists of IDs for better performance
                conn.execute(text("DROP TABLE IF EXISTS temp_select_material_ids;")) # Use temp_select_material_ids
                conn.execute(text("CREATE TEMP TABLE temp_select_material_ids (material_id INTEGER PRIMARY KEY);"))
                conn.execute(
                    text("INSERT INTO temp_select_material_ids (material_id) VALUES (:material_id)"),
                    [{"material_id": mat_id} for mat_id in selected_ids_list],
                )
                conn.commit() # Commit DDL and inserts for temp table
                material_join_clause = "INNER JOIN temp_select_material_ids tsmid ON cd.material_id = tsmid.material_id"

            if excluded_material_ids:
                excluded_ids_list = [str(id_val) for id_val in excluded_material_ids if id_val is not None]
                if excluded_ids_list:
                    material_exclusion_clause = f"AND cd.material_id NOT IN ({', '.join(excluded_ids_list)})"

            if selected_faculties:
                faculties_to_filter = [selected_faculties] if isinstance(selected_faculties, str) else list(selected_faculties)
                invalid_faculties = set(faculties_to_filter) - valid_faculties
                if invalid_faculties:
                    logger.error(f"Invalid faculty abbreviations provided: {invalid_faculties}")
                    # Decide behavior: raise error, or filter by valid ones only? For now, log and continue with valid.
                    # raise ValueError(f"Invalid faculty abbreviations: {invalid_faculties}")
                    faculties_to_filter = [f for f in faculties_to_filter if f in valid_faculties]

                if faculties_to_filter:
                    faculties_string = "', '".join(faculties_to_filter)
                    faculty_where_clause = f"AND cd.faculty_id IN ('{faculties_string}')"
                else: # All provided faculties were invalid, or list was empty
                    logger.warning("No valid faculties selected for filtering in retrieve_full_data, may return empty if faculty filter was intended.")
                    # To ensure it returns nothing if faculties were specified but all invalid:
                    # faculty_where_clause = "AND 1=0"


            # Note: SQL query uses LEFT JOINs which can be slow on large tables without proper indexing.
            # Consider optimizing JOIN conditions and ensuring foreign keys are indexed.
            # The GROUP_CONCAT subqueries can also be performance intensive.
            query: LiteralString = f"""
                WITH CourseDataAggregated AS (
                    SELECT
                        cdcd.copyright_data_id,
                        GROUP_CONCAT(DISTINCT CAST(crs.cursuscode AS TEXT), ' | ') AS cursuscodes,
                        GROUP_CONCAT(DISTINCT crs.programme, ' | ') AS programmes,
                        GROUP_CONCAT(DISTINCT crs.name, ' | ') AS course_names
                    FROM copyright_data_course_data cdcd
                    JOIN course_data crs ON cdcd.course_id = crs.cursuscode
                    GROUP BY cdcd.copyright_data_id
                )
                SELECT
                    cd.*,
                    llm.allowed_usage as llm_allowed_usage, llm.allowed_usage_reasoning as llm_allowed_usage_reason,
                    llm.copyright_status as llm_copyright, llm.copyright_classification_reason as llm_copyright_reason,
                    llm.item_type as llm_item_type, llm.remarks as llm_remarks, llm.item_title as llm_title,
                    llm.copyright_holder as llm_copyright_holder, llm.publisher_name as llm_publisher,
                    llm.isbn as llm_isbn, llm.doi as llm_doi, llm.source_url as llm_source_url,
                    llm.license as llm_license, llm.author_names as llm_authors,
                    cda.cursuscodes, cda.programmes, cda.course_names,
                    (SELECT GROUP_CONCAT(DISTINCT pd.main_name, ' | ') FROM course_employee ce JOIN person_data pd ON ce.person_id = pd.id WHERE ce.role = 'contact' AND ce.course_id IN (SELECT cdcd_sub.course_id FROM copyright_data_course_data cdcd_sub WHERE cdcd_sub.copyright_data_id = cd.material_id)) AS course_contacts_names,
                    (SELECT GROUP_CONCAT(DISTINCT pd.email, ' | ') FROM course_employee ce JOIN person_data pd ON ce.person_id = pd.id WHERE ce.role = 'contact' AND ce.course_id IN (SELECT cdcd_sub.course_id FROM copyright_data_course_data cdcd_sub WHERE cdcd_sub.copyright_data_id = cd.material_id)) AS course_contacts_emails,
                    (SELECT GROUP_CONCAT(DISTINCT f.abbreviation, ' | ') FROM course_employee ce JOIN person_data pd ON ce.person_id = pd.id LEFT JOIN faculty f ON pd.faculty_id = f.abbreviation WHERE ce.role = 'contact' AND ce.course_id IN (SELECT cdcd_sub.course_id FROM copyright_data_course_data cdcd_sub WHERE cdcd_sub.copyright_data_id = cd.material_id)) AS course_contacts_faculties,
                    (SELECT GROUP_CONCAT(DISTINCT org.full_abbreviation, ' | ') FROM course_employee ce JOIN person_data pd ON ce.person_id = pd.id LEFT JOIN person_data_organization_data pdod ON pd.id = pdod.person_data_id LEFT JOIN organization_data org ON pdod.organization_id = org.id WHERE ce.role = 'contact' AND ce.course_id IN (SELECT cdcd_sub.course_id FROM copyright_data_course_data cdcd_sub WHERE cdcd_sub.copyright_data_id = cd.material_id)) AS course_contacts_organizations
                FROM copyright_data cd
                {material_join_clause}
                LEFT JOIN llm_classification_data llm ON cd.llm_classification_id = llm.id
                LEFT JOIN CourseDataAggregated cda ON cd.material_id = cda.copyright_data_id
                WHERE 1=1
                {faculty_where_clause}
                {material_exclusion_clause}
                GROUP BY cd.material_id -- Ensure one row per copyright item if multiple courses/contacts exist
                ORDER BY cd.material_id;
            """
            df = pl.read_database(query=query, connection=conn, infer_schema_length=None)
            if selected_material_ids: # Clean up temp table
                conn.execute(text("DROP TABLE IF EXISTS temp_select_material_ids;"))
                conn.commit()


        cols_to_drop = ["llm_classification_id", "created_at", "modified_at", "possible_fine", "infringement"]
        existing_cols_to_drop = [col for col in cols_to_drop if col in df.columns]
        if existing_cols_to_drop:
            df = df.drop(existing_cols_to_drop)

        if "faculty_id" in df.columns: # Ensure faculty_id is renamed if present
             df = df.rename(mapping={"faculty_id": "faculty"})


        if df.is_empty():
            logger.info("retrieve_full_data returned no rows.")
            return df

        # This check might be too strict if a valid df can have all null material_ids temporarily
        # if df["material_id"].is_null().all():
        #     logger.warning("All material_ids are null in retrieve_full_data result.")
        #     return pl.DataFrame() # Return empty if all material_ids are null

        # Process JSON-like list columns (e.g. from LLM data)
        llm_list_cols = ["llm_isbn", "llm_doi", "llm_source_url", "llm_license", "llm_authors", "llm_topic"]
        for colname in llm_list_cols:
            if colname in df.columns:
                try:
                    # Attempt to decode if it's a string that looks like a list, then join.
                    # If it's already a list (e.g. from some DBs), just join.
                    # This part is tricky as pl.read_database might already parse some JSON.
                    df = df.with_columns(
                        pl.col(colname).map_elements(
                            lambda x: " | ".join(json.loads(x)) if isinstance(x, str) and x.startswith("[") else
                                      (" | ".join(x) if isinstance(x, list) else x),
                            return_dtype=pl.Utf8
                        ).alias(colname)
                    )
                except Exception as e: # Broad exception for varied parsing issues
                    logger.warning(f"Could not process/join list-like column '{colname}': {e}. Column type: {df[colname].dtype}")
        logger.info(f"retrieve_full_data processed {len(df)} rows.")

    except Exception as e:
        logger.error(f"Error in retrieve_full_data: {e}")
        logger.error(traceback.format_exc())
        return pl.DataFrame() # Return empty DataFrame on error

    return df


def get_llm_classification_schema() -> dict[str, pl.PolarsDataType]: # Changed type to PolarsDataType
    """
    Defines the expected Polars schema for LLM classification data.
    Used for returning an empty DataFrame with correct types if no data is found.

    Returns:
        dict[str, pl.PolarsDataType]: A dictionary mapping column names to Polars data types.
    """
    return {
        "allowed_usage_llm": pl.Categorical, "allowed_usage_reasoning_llm": pl.Utf8,
        "copyright_status_llm": pl.Categorical, "copyright_classification_reason_llm": pl.Utf8,
        "item_type_llm": pl.Categorical, "item_type_classification_reason_llm": pl.Utf8,
        "publisher_name_llm": pl.Utf8, "copyright_holder_llm": pl.Utf8,
        "item_title_llm": pl.Utf8, "pdf_page_count_llm": pl.Int64,
        "remarks_llm": pl.Utf8, "author_names_llm": pl.List(pl.Utf8),
        "doi_llm": pl.List(pl.Utf8), "isbn_llm": pl.List(pl.Utf8),
        "source_url_llm": pl.List(pl.Utf8), "license_llm": pl.List(pl.Utf8),
        "topic_llm": pl.List(pl.Utf8), "material_id": pl.Int64, # Changed from Int64 to Utf8 if material_id is string
    }


def retrieve_llm_classifications(
    selected_material_ids: Iterable[int] | None = None,
) -> pl.DataFrame:
    """
    Retrieves LLM classification data, optionally filtered by material IDs.

    Args:
        selected_material_ids: Optional. Select classifications ONLY for these material IDs.

    Returns:
        pl.DataFrame: A DataFrame of LLM classification data.
                      Returns an empty DataFrame with schema if no data or error.
    """
    global engine
    if not engine:
        engine = init_engine()

    df_schema = get_llm_classification_schema()
    df: pl.DataFrame = pl.DataFrame(schema=df_schema) # Initialize with schema

    if not engine:
        logger.error("SQLAlchemy engine not initialized for retrieve_llm_classifications.")
        return df

    material_join_clause: str = ""
    try:
        with engine.connect() as conn:
            if selected_material_ids:
                selected_ids_list = list(selected_material_ids)
                if not selected_ids_list:
                    logger.warning("retrieve_llm_classifications received empty selected_material_ids.")
                    return df
                conn.execute(text("DROP TABLE IF EXISTS temp_llm_material_ids;")) # Use specific temp table name
                conn.execute(text("CREATE TEMP TABLE temp_llm_material_ids (material_id INTEGER PRIMARY KEY);"))
                conn.execute(
                    text("INSERT INTO temp_llm_material_ids (material_id) VALUES (:material_id)"),
                    [{"material_id": mat_id} for mat_id in selected_ids_list],
                )
                conn.commit()
                material_join_clause = "INNER JOIN temp_llm_material_ids tmid ON llm.used_material_id = tmid.material_id"

            query: LiteralString = f"""
                SELECT llm.* FROM llm_classification_data llm {material_join_clause}
            """
            raw_df = pl.read_database(query=query, connection=conn, infer_schema_length=None)

            if selected_material_ids: # Clean up temp table
                conn.execute(text("DROP TABLE IF EXISTS temp_llm_material_ids;"))
                conn.commit()

            if raw_df.is_empty():
                logger.info("No LLM classifications found for the given criteria.")
                return df # Return empty df with schema

            # Rename columns and process JSON list columns
            rename_map = {
                "allowed_usage": "allowed_usage_llm", "allowed_usage_reasoning": "allowed_usage_reasoning_llm",
                "copyright_status": "copyright_status_llm", "copyright_classification_reason": "copyright_classification_reason_llm",
                "item_type": "item_type_llm", "item_type_classification_reason": "item_type_classification_reason_llm",
                "remarks": "remarks_llm", "author_names": "author_names_llm", "item_title": "item_title_llm",
                "publisher_name": "publisher_name_llm", "copyright_holder": "copyright_holder_llm",
                "doi": "doi_llm", "isbn": "isbn_llm", "source_url": "source_url_llm",
                "license": "license_llm", "topic": "topic_llm", "pdf_page_count": "pdf_page_count_llm",
                "used_material_id": "material_id", # This is key for joining
            }
            # Select only columns that exist in raw_df before renaming
            cols_to_rename = {k: v for k, v in rename_map.items() if k in raw_df.columns}
            df_renamed = raw_df.rename(cols_to_rename)

            # Drop Tortoise internal columns if they exist
            internal_cols_to_drop = ["id", "created_at", "modified_at"]
            existing_internal_cols = [col for col in internal_cols_to_drop if col in df_renamed.columns]
            if existing_internal_cols:
                df_processed = df_renamed.drop(existing_internal_cols)
            else:
                df_processed = df_renamed

            # Ensure material_id is of the correct type (matching schema)
            if "material_id" in df_processed.columns:
                 df_processed = df_processed.with_columns(pl.col("material_id").cast(df_schema["material_id"]))


            json_list_cols_renamed = ["author_names_llm", "doi_llm", "isbn_llm", "source_url_llm", "license_llm", "topic_llm"]
            for colname in json_list_cols_renamed:
                if colname in df_processed.columns:
                    df_processed = df_processed.with_columns(
                        pl.col(colname).map_elements(
                            lambda x: " | ".join(json.loads(x)) if isinstance(x, str) and x.startswith("[") else
                                      (" | ".join(x) if isinstance(x, list) else x),
                            return_dtype=pl.Utf8
                        ).alias(colname)
                    )
            logger.info(f"Retrieved and processed {len(df_processed)} LLM classifications.")
            return df_processed # Return processed data

    except Exception as e:
        logger.error(f"Error retrieving LLM classifications: {e}")
        logger.error(traceback.format_exc())
        return df # Return empty df with schema on error


def retrieve_osiris_data(material_ids: list[int]) -> list[dict[str, Any]]:
    """
    Retrieves copyright data and richly nested related data (faculty, courses,
    persons, organizations) for the given material IDs using SQL JSON functions.
    This function is complex and relies on specific SQL JSON capabilities of the database.

    Args:
        material_ids: A list of material IDs to retrieve data for.

    Returns:
        list[dict[str, Any]]: A list of nested dictionaries, where each dictionary
                              represents one copyright item and its related data.
                              Returns an empty list if material_ids is empty or no data is found, or on error.
    """
    global engine
    if not engine:
        engine = init_engine()

    results: list[dict[str, Any]] = []
    if not engine:
        logger.error("SQLAlchemy engine not initialized for retrieve_osiris_data.")
        return results

    if not material_ids:
        logger.warning("No material IDs provided to retrieve_osiris_data. Returning empty list.")
        return results

    # Ensure material_ids are integers for the query
    try:
        processed_material_ids = [int(mid) for mid in material_ids]
    except (ValueError, TypeError) as e:
        logger.error(f"Invalid material_ids provided: {material_ids}. Error: {e}")
        return results

    mat_id_query_filter: str
    if len(processed_material_ids) == 1:
        mat_id_query_filter = f"WHERE cd.material_id = {processed_material_ids[0]}"
    else:
        ids_str = ", ".join(map(str, processed_material_ids))
        mat_id_query_filter = f"WHERE cd.material_id IN ({ids_str})"

    # This SQL query is highly database-specific (SQLite JSON functions) and complex.
    # TODO: Consider breaking down into smaller queries or views if performance is an issue or for clarity.
    # Ensure all table and column names match the actual schema.
    query: LiteralString = f"""
        WITH RECURSIVE OrgHierarchyUp (id, name, abbreviation, parent_organization_id, path_ids, path_abbrs) AS (
            SELECT id, name, abbreviation, parent_organization_id, CAST(id AS TEXT), CAST(abbreviation AS TEXT) FROM organization_data
            UNION ALL
            SELECT child.id, child.name, child.abbreviation, parent.parent_organization_id,
                   parent.id || '/' || child.path_ids, parent.abbreviation || '/' || child.path_abbrs
            FROM organization_data parent JOIN OrgHierarchyUp child ON child.parent_organization_id = parent.id
        ),
        FullOrgPaths AS (
            SELECT id, path_ids, path_abbrs FROM (
                SELECT id, path_ids, path_abbrs, ROW_NUMBER() OVER (PARTITION BY id ORDER BY LENGTH(path_ids) DESC) as rn
                FROM OrgHierarchyUp
            ) AS RankedPaths WHERE rn = 1
        ),
        PersonOrgs AS (
            SELECT pdod.person_data_id, JSON_GROUP_ARRAY(JSON_OBJECT(
                'id', org.id, 'name', org.name, 'abbreviation', org.abbreviation,
                'full_abbreviation', org.full_abbreviation, 'hierarchy_level', org.hierarchy_level,
                'parent_organization_id', org.parent_organization_id, 'full_parent_abbreviations', fop.path_abbrs
            )) AS organizations_json
            FROM person_data_organization_data pdod
            JOIN organization_data org ON pdod.organization_id = org.id
            LEFT JOIN FullOrgPaths fop ON org.id = fop.id GROUP BY pdod.person_data_id
        ),
        CoursePersons AS (
            SELECT ce.course_id, JSON_GROUP_ARRAY(JSON_OBJECT(
                'id', p.id, 'main_name', p.main_name, 'email', p.email, 'first_name', p.first_name,
                'people_page_url', p.people_page_url, 'faculty_id', p.faculty_id, 'role', ce.role,
                'organizations', JSON(po.organizations_json)
            ) ORDER BY p.main_name) AS persons_json
            FROM course_employee ce JOIN person_data p ON ce.person_id = p.id
            LEFT JOIN PersonOrgs po ON p.id = po.person_data_id GROUP BY ce.course_id
        ),
        CopyrightCourses AS (
            SELECT cdcd.copyright_data_id, JSON_GROUP_ARRAY(JSON_OBJECT(
                'cursuscode', crs.cursuscode, 'internal_id', crs.internal_id, 'name', crs.name,
                'short_name', crs.short_name, 'year', crs.year, 'programme', crs.programme,
                'ec', crs.ec, 'faculty_id', crs.faculty_id, 'persons', JSON(cp.persons_json)
            ) ORDER BY crs.name) AS courses_json
            FROM copyright_data_course_data cdcd JOIN course_data crs ON cdcd.course_id = crs.cursuscode
            LEFT JOIN CoursePersons cp ON crs.cursuscode = cp.course_id GROUP BY cdcd.copyright_data_id
        )
        SELECT cd.*,
            (SELECT JSON_OBJECT('abbreviation', f.abbreviation, 'name', f.name, 'full_abbreviation', f.full_abbreviation)
             FROM faculty f WHERE f.abbreviation = cd.faculty_id) AS faculty_data,
            JSON(cc.courses_json) AS courses
        FROM copyright_data cd
        LEFT JOIN CopyrightCourses cc ON cd.material_id = cc.copyright_data_id
        {mat_id_query_filter};
    """
    try:
        with engine.connect() as conn:
            db_result = conn.execute(text(query)).fetchall()
            for row_proxy in db_result: # Iterate over RowProxy objects
                item_dict = dict(row_proxy._mapping) # Convert RowProxy to dict

                # Safely parse top-level JSON fields
                for key in ["faculty_data", "courses"]:
                    json_string = item_dict.get(key)
                    if isinstance(json_string, str):
                        try: item_dict[key] = json.loads(json_string)
                        except json.JSONDecodeError:
                            logger.warning(f"Could not decode JSON for key '{key}' in material_id {item_dict.get('material_id')}. Value: '{json_string[:100]}...'")
                            item_dict[key] = [] if key == "courses" else None # Default to empty list for courses, None for faculty
                    elif json_string is None and key == "courses":
                        item_dict[key] = [] # Ensure courses is an empty list if null

                # Safely parse nested JSON fields (persons, organizations)
                if isinstance(item_dict.get("courses"), list):
                    for course in item_dict["courses"]:
                        if isinstance(course, dict) and isinstance(course.get("persons"), str):
                            try: course["persons"] = json.loads(course["persons"])
                            except json.JSONDecodeError: course["persons"] = []

                        if isinstance(course.get("persons"), list):
                            for person in course["persons"]:
                                if isinstance(person, dict) and isinstance(person.get("organizations"), str):
                                    try: person["organizations"] = json.loads(person["organizations"])
                                    except json.JSONDecodeError: person["organizations"] = []
                                elif isinstance(person, dict) and "organizations" not in person:
                                     person["organizations"] = [] # Ensure key exists
                results.append(item_dict)
        logger.info(f"Retrieved and parsed Osiris data for {len(results)} material IDs.")
    except Exception as e:
        logger.error(f"Database query or JSON parsing failed in retrieve_osiris_data: {e}")
        logger.error(traceback.format_exc())

    return results


async def retrieve_item_history(material_ids: list[int]) -> list[ItemUpdate]:
    """
    Retrieves the history of changes (ItemUpdate records) for the given material IDs.

    Args:
        material_ids: A list of material IDs.

    Returns:
        list[ItemUpdate]: A list of ItemUpdate ORM objects. Returns an empty list if
                          no IDs are provided or no history is found.
    """
    await init_tortoise() # Ensure Tortoise is initialized

    if not material_ids:
        logger.warning("No material IDs provided to retrieve_item_history. Returning empty list.")
        return []

    # Ensure material_ids are integers, as model expects int.
    try:
        processed_material_ids = [int(mid) for mid in material_ids]
    except (ValueError, TypeError) as e:
        logger.error(f"Invalid material_ids provided for history: {material_ids}. Error: {e}")
        return []

    items: list[ItemUpdate] = []
    try:
        items = await ItemUpdate.filter(material_id__in=processed_material_ids).all().order_by('-created_at')
        logger.info(f"Retrieved {len(items)} history entries for {len(processed_material_ids)} material IDs.")
    except Exception as e:
        logger.error(f"Error retrieving item history: {e}")
        logger.error(traceback.format_exc())
    # Tortoise connections are typically managed globally or per request in web apps.
    # For a script, ensure it's closed if opened by init_tortoise, but init_tortoise itself doesn't open/close.
    # await Tortoise.close_connections() # This might be too aggressive if Tortoise is used elsewhere.
    return items
