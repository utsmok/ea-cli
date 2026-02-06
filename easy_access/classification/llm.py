# module that uses llms for classifications

# we'll first try with local ollama + dspy as a test

import sqlite3
from dataclasses import dataclass, field
from enum import Enum
from time import time

import dspy
import orjson
from rich import print

start_time = time()
lm = dspy.LM(
    "ollama_chat/qwen3:4b-instruct", api_base="http://localhost:11434", api_key=""
)
dspy.configure(lm=lm)


# RETRIEVE / BUILD TRAINING DATA

db_fields = [
    "created_at",
    "modified_at",
    "id",
    "filename",
    "url",
    "file_size",
    "retrieved_on",
    "current_file_name",
    "author",
    "title",
    "subject",
    "keywords",
    "producer",
    "creation_date",
    "mod_date",
    "creator",
    "summary",
    "description",
    "filehash",
    "extraction_attempted",
    "extraction_successful",
    "num_pages",
    "num_words",
    "num_images",
    "extracted_entities",
    "copyright_item_id",
    "canvas_metadata_id",
    "extracted_text_id",
    "pdf_text_data_id",
    "extracted_text",
    "created_at",
    "modified_at",
    "id",
    "uuid",
    "folder_id",
    "display_name",
    "filename",
    "upload_status",
    "content_type",
    "mime_class",
    "category",
    "download_url",
    "size",
    "thumbnail_url",
    "canvas_created_at",
    "canvas_updated_at",
    "locked",
    "hidden",
    "lock_at",
    "unlock_at",
    "visibility_level",
    "user_id",
    "user_anonymous_id",
    "user_display_name",
    "user_avatar_image_url",
    "user_html_url",
    "user_pronouns",
    "created_at",
    "modified_at",
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
    "ml_prediction",
    "manual_classification",
    "manual_identifier",
    "scope",
    "remarks",
    "auditor",
    "last_change",
    "status",
    "isbn",
    "doi",
    "in_collection",
    "pagecount",
    "wordcount",
    "picturecount",
    "author",
    "publisher",
    "reliability",
    "pages_x_students",
    "count_students_registered",
    "filehash",
    "last_scan_date_university",
    "last_scan_date_course",
    "retrieved_from_copyright_on",
    "workflow_status",
    "possible_fine",
    "infringement",
    "file_exists",
    "last_canvas_check",
    "is_duplicate",
    "faculty_id",
    "v2_lengte",
    "v2_overnamestatus",
    "v2_manual_classification",
    "courses_json",
]

selected_fields = [
    "material_id",
    "v2_lengte",
    "v2_overnamestatus",
    "v2_manual_classification",
    "courses_json",
    "pagecount",
    "wordcount",
    "picturecount",
    "remarks",
    "manual_classification",
    "owner",
    "filename",
    "title",
    "size",
    "extracted_text",
    "author",
    "title",
]
sql_query = """
WITH course_teachers AS (
  SELECT
    ce.course_id,
    JSON_GROUP_ARRAY(
      JSON_OBJECT(
        'teacher_id', p.id,
        'teacher_name', p.main_name,
		'teacher_email', p.email

      )
    ) AS teachers_json
  FROM
    course_employee AS ce
  JOIN
    person_data AS p ON ce.person_id = p.id
  GROUP BY
    ce.course_id
),
copyright_courses AS (
  SELECT
    cdcd.copyright_data_id,
    JSON_GROUP_ARRAY(
      JSON_OBJECT(
        'course_code', cd.cursuscode,
        'course_name', cd.name,
		'course_programme', cd.programme,
        'teachers', JSON(ct.teachers_json)
      )
    ) AS courses_json
  FROM
    (SELECT DISTINCT copyright_data_id, course_id FROM copyright_data_course_data) AS cdcd
  JOIN
    course_data AS cd ON cdcd.course_id = cd.cursuscode
  LEFT JOIN
    course_teachers AS ct ON cd.cursuscode = ct.course_id
  GROUP BY
    cdcd.copyright_data_id
)
SELECT
  pdf_data.*,
  pdf_text_data.id as pdf_text_data_id, -- Aliased to prevent potential 'id' column conflicts
  SUBSTR(pdf_text_data.extracted_text, 1, 10000) AS extracted_text,
  pdf_canvas_metadata.*,
  copyright_data.*,
  cc.courses_json
FROM
  pdf_data
INNER JOIN
  pdf_text_data ON pdf_data.extracted_text_id = pdf_text_data.id
INNER JOIN
  pdf_canvas_metadata ON pdf_data.canvas_metadata_id = pdf_canvas_metadata.id
INNER JOIN
  copyright_data ON pdf_data.copyright_item_id = copyright_data.material_id
LEFT JOIN
  copyright_courses AS cc ON copyright_data.material_id = cc.copyright_data_id
WHERE
  LENGTH(pdf_text_data.extracted_text) > 50  AND (copyright_data.manual_classification) IS NOT NULL
"""

db = sqlite3.connect("e:\\ea-cli\\db.sqlite3")
cur = db.cursor()
cur.execute(sql_query)
rows = cur.fetchall()

print(f"fetched {len(rows)} rows from the database")
print(f" fields are: {[description[0] for description in cur.description]}")

# for each row, select only the fields in `selected_fields`
final_data = []
for row in rows:
    row = dict(
        zip([description[0] for description in cur.description], row, strict=False)
    )
    filtered_row = {key: row[key] for key in selected_fields if key in row}
    # parse courses_json if it exists
    if "courses_json" in filtered_row and filtered_row["courses_json"]:
        try:
            filtered_row["courses_json"] = orjson.loads(filtered_row["courses_json"])  # type: ignore
            # dedupe teachers in each course
            for course in filtered_row["courses_json"]:
                teachers_by_id = {
                    t.get("teacher_id"): t
                    for t in course.get("teachers", [])
                    if t.get("teacher_id")
                }
                course["teachers"] = list(teachers_by_id.values())

        except orjson.JSONDecodeError:
            filtered_row["courses_json"] = []

    final_data.append(filtered_row)

print(f"filtered to {len(final_data)} rows with selected fields")
print(f" selected fields are: {[key for key in final_data[0]]}")
print(" example row:")
max_key_str_len = max(len(str(key)) for key in final_data[0])

for key, value in final_data[0].items():
    # print results, align the colon using f-string formatting

    if key == "extracted_text":
        print(f" {key:<{max_key_str_len}}: {len(value)} characters")
    elif key == "courses_json":
        print(f" {key:<{max_key_str_len}}:")
        print(value)
    else:
        print(f" {key:<{max_key_str_len}}:    {value}")


class DocumentClassification(Enum):
    # example, use actual classifications from easy access model
    OPEN_ACCESS = "open access"  # specific license like CC-BY or GPL
    PUBLIC_DOMAIN = "public domain"  # not owned by anyone, free to use, e.g. law text, very old books
    OWN_MATERIAL = (
        "own material"  # made for or by an employee of the University of Twente
    )
    COPYRIGHTED_MATERIAL = (
        "copyrighted material"  # not free to use, owned by a publisher for instance
    )


@dataclass
class DocumentMetadata:
    # change this to match the data retreived from the database
    title: str = ""
    filename: str = ""
    authors: list[str] = field(default_factory=list)
    uploader: str = ""  # name of the person who uploaded the document
    teachers: list[str] = field(
        default_factory=list
    )  # Teachers associated with the item


class Classify(dspy.Signature):
    """Determine if hosting this pdf on our environment constitutes an copyright infringement or not, based on the text inside the document and the related metadata."""

    document_text: str = dspy.InputField(desc="The text extracted from the pdf.")
    confidence: float = dspy.OutputField(
        desc="Confidence score between 0 and 1 for the classification."
    )
    classification: DocumentClassification = dspy.OutputField(
        desc="The final classification"
    )


classify = dspy.Predict(Classify)
# classification: dspy.Prediction = classify("input")

# print(classification)
print(f"took {time() - start_time:.2f} seconds")
