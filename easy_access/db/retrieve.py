"""
Functions to retrieve data from the database.
"""
from sqlalchemy import Engine, create_engine
import polars as pl
from easy_access.settings import SETTINGS

engine: Engine | None = None

def init_engine() -> None:
    global engine
    engine = create_engine("sqlite:///db.sqlite3")

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
def retrieve_full_data() -> pl.DataFrame:
    """
    Retrieve all copyright items from db enriched with data from other tables
    returns a flat dataframe ready for .xlsx export

    """
    if not engine:
        init_engine()

    query="""WITH CourseDataAggregated AS (
        SELECT
            cdcd.copyright_data_id,
            (SELECT GROUP_CONCAT(cursuscode, ' | ') FROM (SELECT DISTINCT CAST(cd.cursuscode AS TEXT) as cursuscode FROM course_data cd WHERE cd.cursuscode = cdcd.course_id)) AS cursuscodes,
            (SELECT GROUP_CONCAT(programme, ' | ') FROM (SELECT DISTINCT cd.programme FROM course_data cd WHERE cd.cursuscode = cdcd.course_id)) AS programmes,
            (SELECT GROUP_CONCAT(name, ' | ') FROM (SELECT DISTINCT cd.name FROM course_data cd WHERE cd.cursuscode = cdcd.course_id)) AS course_names
        FROM copyright_data_course_data cdcd
    ), PersonDataAggregated AS (
        SELECT
        ce.course_id,
        (SELECT GROUP_CONCAT(email, ' | ') FROM (SELECT DISTINCT pd.email FROM person_data pd WHERE pd.id = ce.person_id)) as course_contacts_emails,
        (SELECT GROUP_CONCAT(main_name, ' | ') FROM (SELECT DISTINCT pd.main_name FROM person_data pd WHERE pd.id = ce.person_id)) as course_contacts_names,
        (SELECT GROUP_CONCAT(abbreviation, ' | ') FROM (SELECT DISTINCT f.abbreviation FROM faculty f LEFT JOIN person_data pd ON pd.faculty_id = f.abbreviation WHERE pd.id = ce.person_id)) as course_contacts_faculties,
        (SELECT GROUP_CONCAT(full_abbreviation, ' | ') FROM (
            SELECT DISTINCT org.full_abbreviation
            FROM organization_data org
            LEFT JOIN person_data_organization_data pdod on pdod.organization_id = org.id
            LEFT JOIN person_data pd ON pd.id = pdod.person_data_id
            WHERE pd.id = ce.person_id
        )) as course_contacts_organizations
        FROM course_employee ce
        WHERE ce.role = 'contact'
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
        cda.cursuscodes,
        cda.programmes,
        cda.course_names,
        pda.course_contacts_names,
        pda.course_contacts_emails,
        pda.course_contacts_faculties,
        pda.course_contacts_organizations
    FROM copyright_data cd
    LEFT JOIN llm_classification_data llm ON cd.llm_classification_id = llm.id
    LEFT JOIN CourseDataAggregated cda ON cd.material_id = cda.copyright_data_id
    LEFT JOIN copyright_data_course_data cdcd ON cd.material_id = cdcd.copyright_data_id
    LEFT JOIN PersonDataAggregated pda ON cdcd.course_id = pda.course_id;
    """
    df: pl.DataFrame = pl.read_database(query=query, connection=engine.connect(), infer_schema_length=None)
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
        ).with_columns(
            [
                pl.col(name=colname).
                str.json_decode(infer_schema_length=None).
                list.join(separator=' | ')
                for colname in [
                    "llm_isbn",
                    "llm_doi",
                    "llm_source_url",
                    "llm_license",
                    "llm_authors"
                ]
            ]
        )
    return df
