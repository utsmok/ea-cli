"""
Functions to retrieve data from the database.
"""
from collections.abc import Iterable
from typing import Literal, LiteralString
from sqlalchemy import Engine, create_engine, text
import polars as pl
from easy_access.settings import SETTINGS
from easy_access.utils import warn

engine: Engine | None = None


def retrieve_copyright_items() -> pl.DataFrame:
    """
    Retrieve all copyright items currently in db
    returns a flat dataframe with the core fields
    """
    if not engine:
        init_engine()

    col_order: set[str] = set(SETTINGS.data_settings.raw_data_col_order)

    if 'google_search_file' in col_order:
        col_order.remove('google_search_file')
    if 'type' in col_order:
        col_order.remove('type')
    if 'faculty' in col_order:
        col_order.remove('faculty')
        select_cols = {*col_order, 'faculty_id AS faculty'}
    else:
        select_cols = col_order

    query: str = "SELECT " + ", ".join(select_cols) + " FROM copyright_data cd"

    df: pl.DataFrame = pl.read_database(query=query, connection=engine.connect(), infer_schema_length=None)
    return df

def retrieve_duplicate_copyright_items() -> pl.DataFrame:
    """
    for copyright_items with duplicates, find the replacing material id
    returns a dataframe with 'material_id', 'is_duplicate', and 'replacement_id' columns
    """
    if not engine:
        init_engine()

    query: str = """
        SELECT material_id, is_duplicate, replacement_id
        FROM copyright_data cd
        WHERE is_duplicate IS TRUE
    """
    df: pl.DataFrame = pl.read_database(query=query, connection=engine.connect(), infer_schema_length=None)
    return df
def init_engine() -> None:
    global engine
    engine = create_engine("sqlite:///db.sqlite3")

def get_valid_faculties() -> set[str]:
    """Retrieves the set of valid faculty abbreviations from the database."""
    if not engine:
        init_engine()
    with engine.connect() as conn:
        result = conn.execute(text("SELECT abbreviation FROM faculty"))
        return {row[0] for row in result.fetchall()}


def retrieve_full_data(
    selected_material_ids: Iterable[int] | None = None,
    selected_faculties: Iterable[str] | str | None = None,
    excluded_material_ids: Iterable[int] | None = None
    ) -> pl.DataFrame:
    """
    Retrieves copyright items, enriched with related data, with optional filtering.

    Args:
        selected_material_ids: Optional iterable of material IDs to select (ONLY these, drop rest).
        selected_faculties: Optional string or list of strings (faculty abbreviations). ONLY return items with these faculties.
        excluded_material_ids: Optional iterable of material IDs to exclude. Excluded_material_ids take precedence over selected_material_ids.

    Returns:
        A Polars DataFrame.
    """
    if not engine:
        init_engine()
    valid_faculties = get_valid_faculties()
    material_join_clause = ""
    faculty_where_clause = ""
    material_exclusion_clause = ""
    with engine.connect() as conn:
        # print all table names
        if selected_material_ids is not None:
            if not selected_material_ids:
                warn(text='retrieve_full_data received an empty collection of selected_material_ids.  Returning empty DataFrame.')
                return pl.DataFrame()
            conn.execute(text("DROP TABLE IF EXISTS temp_material_ids;"))
            conn.execute(text("CREATE TEMP TABLE temp_material_ids (material_id INTEGER);"))
            conn.execute(
                text("INSERT INTO temp_material_ids (material_id) VALUES (:material_id)"),
                [{"material_id": mat_id} for mat_id in selected_material_ids]
            )
            material_join_clause = "INNER JOIN temp_material_ids tmid ON cd.material_id = tmid.material_id"

        if excluded_material_ids is not None:
            excluded_material_ids_list = list(excluded_material_ids)
            if excluded_material_ids_list:  # Only add clause if the list is not empty
                excluded_ids_string = ', '.join(map(str, excluded_material_ids_list))
                material_exclusion_clause = f"AND cd.material_id NOT IN ({excluded_ids_string})"

        if selected_faculties:
            if isinstance(selected_faculties, str):
                selected_faculties = [selected_faculties]  # Make it a list

            # Validate faculties
            invalid_faculties = set(selected_faculties) - valid_faculties
            if invalid_faculties:
                raise ValueError(f"Invalid faculty abbreviations: {invalid_faculties}")

            faculties_string = "', '".join(selected_faculties)  # Escape for SQL
            faculty_where_clause = f"AND cd.faculty_id IN ('{faculties_string}')"


        query: LiteralString = f"""
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
                'created_at',
                'modified_at',
                'possible_fine',
                'infringement'
            ]
        ).rename(
            mapping={
                "faculty_id":"faculty"
            }
        )

    if df.is_empty():
        return df
    if df['material_id'].is_null().all():
        return None

    # drop nulcols from llm cols
    llm_cols = [
                    "llm_isbn",
                    "llm_doi",
                    "llm_source_url",
                    "llm_license",
                    "llm_authors"
                ]
    droplist = []
    # iterate over llm cols
    # if col is null (dtype=null or all values are null), drop it
    for col in llm_cols:
        if df[col].is_null().all():
            df = df.drop(col)
            droplist.append(col)
    select_cols = [col for col in llm_cols if col not in droplist]
    if select_cols:
        df = df.with_columns(
                [

                    pl.col(name=colname).
                    str.json_decode(infer_schema_length=None).
                    list.join(separator=' | ')
                    for colname in select_cols
                ]
            )

    # add droplist cols back in with empty values
    if droplist:
        for col in droplist:
            df = df.with_columns(col = pl.lit(None))

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
        "author_names_llm": pl.List(pl.Utf8),  # Specify list of strings
        "doi_llm": pl.List(pl.Utf8),
        "isbn_llm": pl.List(pl.Utf8),
        "source_url_llm": pl.List(pl.Utf8),
        "license_llm": pl.List(pl.Utf8),
        "topic_llm": pl.List(pl.Utf8),
        "material_id": pl.Int64,
    }

def retrieve_llm_classifications(
    selected_material_ids: Iterable[int] | None = None,
) -> pl.DataFrame:

    if not engine:
        init_engine()
    material_join_clause: Literal[''] = ""
    with engine.connect() as conn:
        # print all table names
        if selected_material_ids is not None:
            if not selected_material_ids:
                warn(text='retrieve_llm_classifications received an empty collection of selected_material_ids.  Returning empty DataFrame.')
                return pl.DataFrame(schema=get_llm_classification_schema())
            conn.execute(text("DROP TABLE IF EXISTS temp_material_ids;"))
            conn.execute(text("CREATE TEMP TABLE temp_material_ids (material_id INTEGER);"))
            conn.execute(
                text("INSERT INTO temp_material_ids (material_id) VALUES (:material_id)"),
                [{"material_id": mat_id} for mat_id in selected_material_ids]
            )
            material_join_clause = "INNER JOIN temp_material_ids tmid ON llm.used_material_id = tmid.material_id"

        query: LiteralString = f"""
        SELECT
            llm.*
        FROM llm_classification_data llm
        {material_join_clause}
        """

        df = pl.read_database(query=query, connection=conn)
        df = df.drop(
            [
                'created_at',
                'modified_at',
                'id'
            ]
        ).rename(
            mapping={
                "allowed_usage":"allowed_usage_llm",
                "allowed_usage_reasoning":"allowed_usage_reasoning_llm",
                "copyright_status":"copyright_status_llm",
                "copyright_classification_reason":"copyright_classification_reason_llm",
                "item_type":"item_type_llm",
                "item_type_classification_reason":"item_type_classification_reason_llm",
                "remarks":"remarks_llm",
                "author_names":"author_names_llm",
                "item_title":"item_title_llm",
                "publisher_name":"publisher_name_llm",
                "copyright_holder":"copyright_holder_llm",
                "doi":"doi_llm",
                "isbn":"isbn_llm",
                "source_url":"source_url_llm",
                "license":"license_llm",
                "topic":"topic_llm",
                "pdf_page_count":"pdf_page_count_llm",
                'used_material_id':'material_id',
            }
        )
        if df.is_empty():
            return None
        if df['material_id'].is_null().all():
            return None

        return df.with_columns([
            pl.col(name=colname)
            .str.json_decode(infer_schema_length=None)
            .fill_null(pl.lit([]))
            .list.eval(pl.element().cast(pl.Utf8))
            .list.join(separator=' | ')
            for colname in [
                "isbn_llm",
                "doi_llm",
                "source_url_llm",
                "license_llm",
                "author_names_llm",
                "topic_llm"
            ]
        ])
