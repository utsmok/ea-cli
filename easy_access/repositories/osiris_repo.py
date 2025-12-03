"""
Repository for OSIRIS-related database operations.

This module centralizes complex SQL queries for retrieving enriched
copyright data with nested course, person, and organization information.
"""

import json
import traceback
from typing import Any

from loguru import logger
from sqlalchemy import Engine, text

from easy_access.db.base import ensure_db_inited, init_engine
from easy_access.settings import Settings


class OsirisRepository:
    """Repository for OSIRIS data operations."""

    def __init__(self, settings: Settings):
        """Initialize the repository with settings."""
        self.settings = settings
        self._engine: Engine | None = None

    @property
    def engine(self) -> Engine:
        """Lazily initialize and return the SQLAlchemy engine."""
        if self._engine is None:
            self._engine = init_engine(settings=self.settings)
        return self._engine

    async def ensure_initialized(self) -> None:
        """Ensure the database is initialized."""
        await ensure_db_inited(self.settings)

    def fetch_enriched_data(
        self, material_ids: list[int] | int
    ) -> list[dict[str, Any]]:
        """
        Retrieves copyright data and richly nested related data (faculty, courses,
        persons, organizations) for the given material IDs using SQL JSON functions.

        Args:
            material_ids: A list of material IDs to retrieve data for.

        Returns:
            A list of nested dictionaries, where each dictionary represents one
            copyright item and its related data. Returns an empty list if
            material_ids is empty or no data is found.
        """
        if not material_ids:
            logger.warning("No material IDs provided. Returning empty list.")
            return []

        # Normalize to list
        if not isinstance(material_ids, list):
            material_ids = [material_ids]

        # Validate that all material_ids are integers to prevent SQL injection
        validated_ids = []
        for mid in material_ids:
            if isinstance(mid, int):
                validated_ids.append(mid)
            elif isinstance(mid, str) and mid.isdigit():
                validated_ids.append(int(mid))
            else:
                logger.warning(f"Skipping invalid material_id: {mid}")
        
        if not validated_ids:
            logger.warning("No valid material IDs after validation. Returning empty list.")
            return []

        # Build the WHERE clause for material_ids (using validated integer IDs only)
        if len(validated_ids) == 1:
            mat_id_query = f"WHERE cd.material_id = {validated_ids[0]}"
        else:
            mat_id_query = f"WHERE cd.material_id IN ({', '.join(map(str, validated_ids))})"

        query = self._build_enriched_data_query(mat_id_query)
        return self._execute_enriched_query(query)

    def _build_enriched_data_query(self, mat_id_query: str) -> str:
        """
        Build the complex SQL query for enriched data retrieval.

        Args:
            mat_id_query: WHERE clause for filtering by material IDs

        Returns:
            Complete SQL query string
        """
        return f"""
WITH RECURSIVE OrgHierarchyUp (id, name, abbreviation, parent_organization_id, path_ids, path_abbrs) AS (
      -- Base case: Start with all organizations
      SELECT id, name, abbreviation, parent_organization_id,
             CAST(id AS TEXT), CAST(abbreviation AS TEXT)
      FROM organization_data

      UNION ALL

      -- Recursive step: Go up one level (Join child's parent_id to parent's id)
      SELECT
        child.id, child.name, child.abbreviation, parent.parent_organization_id,
        parent.id || '/' || child.path_ids,
        parent.abbreviation || '/' || child.path_abbrs
      FROM organization_data parent
      JOIN OrgHierarchyUp child ON child.parent_organization_id = parent.id
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
        'full_parent_abbreviations', fop.path_abbrs
      )
    ) AS organizations_json
  FROM person_data_organization_data pdod
  JOIN organization_data org ON pdod.organization_id = org.id
  LEFT JOIN FullOrgPaths fop ON org.id = fop.id
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
        'organizations', JSON(po.organizations_json)
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
  cd.*,
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

    def _execute_enriched_query(self, query: str) -> list[dict[str, Any]]:
        """
        Execute the enriched data query and parse results.

        Args:
            query: SQL query to execute

        Returns:
            List of parsed dictionaries
        """
        results = []
        try:
            with self.engine.connect() as conn:
                db_result = conn.execute(text(query)).fetchall()

                for row_mapping in db_result:
                    row_mapping = row_mapping._mapping
                    item_dict = dict(row_mapping)

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
                            item_dict[key] = [] if key == "courses" else None

                    # Parse nested JSON in courses
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
                                                and isinstance(
                                                    person["organizations"], str
                                                )
                                            ):
                                                try:
                                                    person["organizations"] = (
                                                        json.loads(
                                                            person["organizations"]
                                                        )
                                                    )
                                                except json.JSONDecodeError:
                                                    person["organizations"] = []
                                            elif (
                                                person
                                                and "organizations" not in person
                                            ):
                                                person["organizations"] = []
                                except json.JSONDecodeError:
                                    course["persons"] = []
                            elif course and "persons" not in course:
                                course["persons"] = []

                    results.append(item_dict)

        except Exception as e:
            logger.error(f"Database query failed: {e}")
            logger.error(traceback.format_exc())

        return results
