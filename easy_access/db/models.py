"""
ORM models for the database.
"""

from datetime import datetime
from enum import Enum
from pathlib import Path

from tortoise import fields
from tortoise.models import Model

from easy_access.db.enums import (
    Classification,
    ClassificationV2,
    EntityTypes,
    Filetype,
    Infringement,
    Lengte,
    OvernameStatus,
    Period,
    Status,
    WorkflowStatus,
)
from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import File


class TimestampMixin:
    created_at = fields.DatetimeField(null=True, auto_now_add=True)
    modified_at = fields.DatetimeField(null=True, auto_now=True)


Department = Enum(
    "Department",
    {
        department: department
        for department in SETTINGS.university_settings.department_mapping
    },
)


class v1_CopyrightItem(Model, TimestampMixin):
    """
    Copyright item as imported from the v1 sheets (2024-2025).
    """

    material_id = fields.IntField(primary_key=True)
    matching_copyright_item = fields.ForeignKeyField(
        "models.CopyrightItem", related_name="v1_items", null=True
    )
    workflow_status = fields.CharEnumField(
        enum_type=WorkflowStatus, max_length=255, default=WorkflowStatus.ToDo
    )
    retrieved_from_copyright_on = fields.DatetimeField(null=True)
    url = fields.CharField(max_length=2048, null=True)
    manual_classification = fields.CharField(max_length=2048, null=True, db_index=True)
    remarks = fields.CharField(max_length=10000, null=True)
    scope = fields.CharField(max_length=255, null=True)
    faculty = fields.CharField(max_length=255, null=True)
    ml_prediction = fields.CharEnumField(
        enum_type=Classification, max_length=255, null=True
    )
    filename = fields.CharField(max_length=2048, null=True)
    title = fields.CharField(max_length=2048, null=True)
    filehash = fields.CharField(max_length=255, null=True)
    owner = fields.CharField(max_length=2048, null=True)
    period = fields.CharEnumField(enum_type=Period, max_length=255, null=True)
    department = fields.CharField(
        max_length=2048, db_index=True, null=True
    )  # turn this into a relation w/ programmes later
    course_code = fields.CharField(max_length=255, db_index=True, null=True)
    course_name = fields.CharField(max_length=2048, db_index=True, null=True)
    filetype = fields.CharEnumField(
        enum_type=Filetype, max_length=255, default=Filetype.UNKNOWN
    )
    classification = fields.CharEnumField(
        enum_type=Classification, max_length=255, default=Classification.LANGE_OVERNAME
    )
    manual_identifier = fields.CharField(max_length=2048, null=True)
    auditor = fields.CharField(max_length=10000, null=True)
    last_change = fields.DateField(null=True)
    status = fields.CharEnumField(
        enum_type=Status, max_length=255, default=Status.PUBLISHED, db_index=True
    )
    isbn = fields.CharField(max_length=255, null=True)
    doi = fields.CharField(max_length=255, null=True)
    in_collection = fields.BooleanField(null=True)
    pagecount = fields.IntField(default=0)
    wordcount = fields.IntField(default=0)
    picturecount = fields.IntField(default=0)
    author = fields.CharField(max_length=2048, null=True)
    publisher = fields.CharField(max_length=2048, null=True)
    reliability = fields.IntField(default=0)
    pages_x_students = fields.IntField(default=0)
    count_students_registered = fields.IntField(default=0)
    cursuscodes = fields.CharField(
        max_length=2048, null=True
    )  # probably single course code(?)

    pdf: fields.ReverseRelation["PDF"]


class CopyrightItem(Model, TimestampMixin):
    """
    Core item in the dataset. Contains all the data for one item on Canvas.
    Contains all the data as exported from the copyrighttool.
    """

    material_id = fields.IntField(primary_key=True)
    period = fields.CharEnumField(enum_type=Period, max_length=255)
    department = fields.CharField(
        max_length=2048, db_index=True
    )  # turn this into a relation w/ programmes later
    course_code = fields.CharField(max_length=255, db_index=True)
    course_name = fields.CharField(max_length=2048, db_index=True)
    url = fields.CharField(max_length=255, unique=True, null=True)
    filename = fields.CharField(max_length=2048, db_index=True, null=True)
    title = fields.CharField(max_length=2048, null=True)
    owner = fields.CharField(max_length=2048, null=True)
    filetype = fields.CharEnumField(
        enum_type=Filetype, max_length=255, default=Filetype.UNKNOWN
    )
    classification = fields.CharEnumField(
        enum_type=Classification, max_length=255, default=Classification.LANGE_OVERNAME
    )
    ml_prediction = fields.CharEnumField(
        enum_type=Classification, max_length=255, db_index=True, null=True
    )
    manual_classification = fields.CharField(max_length=2048, null=True, db_index=True)
    manual_identifier = fields.CharField(max_length=2048, null=True)
    v2_manual_classification = fields.CharEnumField(
        enum_type=ClassificationV2,
        max_length=255,
        db_index=True,
        null=True,
        default=ClassificationV2.ONBEKEND,
    )
    v2_overnamestatus = fields.CharEnumField(
        enum_type=OvernameStatus,
        max_length=255,
        db_index=True,
        null=True,
        default=OvernameStatus.ONBEKEND,
    )
    v2_lengte = fields.CharEnumField(
        enum_type=Lengte,
        max_length=255,
        db_index=True,
        default=Lengte.ONBEKEND,
        null=True,
    )
    scope = fields.CharField(max_length=255, null=True)
    remarks = fields.CharField(max_length=10000, null=True, db_index=True)
    auditor = fields.CharField(max_length=10000, null=True)
    last_change = fields.DateField(null=True)
    status = fields.CharEnumField(
        enum_type=Status, max_length=255, default=Status.PUBLISHED, db_index=True
    )
    isbn = fields.CharField(max_length=255, null=True)
    doi = fields.CharField(max_length=255, null=True)
    in_collection = fields.BooleanField(null=True)
    pagecount = fields.IntField()
    wordcount = fields.IntField()
    picturecount = fields.IntField()
    author = fields.CharField(max_length=2048, null=True)
    publisher = fields.CharField(max_length=2048, null=True)
    reliability = fields.IntField()
    pages_x_students = fields.IntField()
    count_students_registered = fields.IntField()
    filehash = fields.CharField(max_length=255, null=True)
    last_scan_date_university = fields.DateField(null=True)
    last_scan_date_course = fields.DateField(null=True)

    # Workflow data, added by the tool -- not present in the raw data!
    retrieved_from_copyright_on = fields.DatetimeField(null=True, db_index=True)
    workflow_status = fields.CharEnumField(
        enum_type=WorkflowStatus,
        max_length=255,
        default=WorkflowStatus.ToDo,
        db_index=True,
    )
    possible_fine = fields.FloatField(null=True)
    infringement = fields.CharEnumField(
        enum_type=Infringement, max_length=255, default=Infringement.UNDETERMINED
    )
    file_exists = fields.BooleanField(
        null=True, default=None
    )  # whether the file exists on Canvas. Null = unchecked.
    last_canvas_check = fields.DatetimeField(
        null=True
    )  # when was the file existence last checked on Canvas
    canvas_course_id = fields.IntField(
        null=True, db_index=True
    )  # the course ID on canvas, to link to courses more easily

    # relations

    courses = fields.ManyToManyField("models.Course", related_name="course_items")
    faculty = fields.ForeignKeyField(
        "models.Faculty", related_name="faculty_items", to_field="abbreviation"
    )
    changes = fields.ManyToManyField("models.ItemUpdate", related_name="item")

    is_duplicate = fields.BooleanField(
        db_index=True, null=True
    )  # if this item is a duplicate of another item -- determined by comparing PDFs

    pdf: fields.ReverseRelation["PDF"]
    v1_items: fields.ReverseRelation["v1_CopyrightItem"]
    class Meta:
        table = "copyright_data"

    def actual_status(self) -> Status:
        if not self.url or self.url.strip() == "":
            return Status.DELETED
        elif self.file_exists in [True, 1, "1"]:
            return self.status
        else:
            return Status.DELETED

    def misaligned_status(self) -> bool:
        return self.status != self.actual_status()

    def status_details(self) -> dict[str, str | bool | datetime | int | Status]:
        return {
            "material_id": self.material_id,
            "filename": self.filename,
            "url": self.url,
            "status": self.status,
            "actual_status": self.actual_status(),
            "file_exists": self.file_exists,
            "last_canvas_check": self.last_canvas_check,
        }

    # course_links can only be retrieved, not set here

    @property
    def course_link(self) -> str:
        if not self.canvas_course_id or not self.filename:
            return ""
        base_url = SETTINGS.university_settings.lms.url
        return f"{base_url}/courses/{self.canvas_course_id}/files?search_term={self.filename.replace(' ', '%20')}"

    def __str__(self):
        return str(self.filename) + " (" + str(self.material_id) + ")"


class ItemUpdate(Model, TimestampMixin):
    """
    Stores changes made to a copyright item.
    """

    id = fields.IntField(primary_key=True)
    change_details = fields.JSONField()
    material_id = fields.IntField()

    class Meta:
        table = "item_updates"


class MissingCourse(Model, TimestampMixin):
    """
    Store cursuscodes that do not yet have a 'Course' entry in the db, to be retrieved and added later.
    """

    cursuscode = fields.IntField(primary_key=True)

    class Meta:
        table = "missing_courses"

    def __str__(self):
        return str(self.cursuscode)


class Course(Model, TimestampMixin):
    """
    Data for a course; mainly from OSIRIS.
    """

    cursuscode = fields.IntField(primary_key=True)
    internal_id = fields.IntField(unique=True)
    year = fields.IntField()  # use academic year; ie 2024-2025 --> 2024
    name = fields.CharField(max_length=2048)
    short_name = fields.CharField(max_length=255, null=True)
    faculty = fields.ForeignKeyField(
        "models.Faculty",
        related_name="faculty_courses",
        null=True,
        to_field="abbreviation",
    )
    ec = fields.FloatField(null=True)
    programme = fields.CharField(
        max_length=2048, null=True
    )  # try to turn into a relation to Programme later.
    notes = fields.CharField(
        max_length=10000, null=True
    )  # made null=True for optional notes
    category = fields.CharField(
        max_length=2048, null=True
    )  # made null=True for optional category
    teachers = fields.ManyToManyField(
        "models.Person",
        through="course_employee",
        forward_key="person_id",
        backward_key="course_cursuscode",
        related_name="courses",
    )

    class Meta:
        table = "course_data"

    def __str__(self):
        return self.name + " (" + str(self.cursuscode) + ")"


class CourseEmployee(Model, TimestampMixin):
    """
    Many-to-many relation between Course and Person, to store the various types of teachers of a course.
    """

    id = fields.IntField(primary_key=True)
    course = fields.ForeignKeyField("models.Course", related_name="course_employee")
    person = fields.ForeignKeyField("models.Person", related_name="course_employee")
    role = fields.CharField(
        max_length=2048, null=True
    )  # made null=True for optional role

    class Meta:
        table = "course_employee"


class Person(Model, TimestampMixin):
    """
    Person data; data from people pages.
    Initialized with just an 'input_name'. If no match is found, the rest will remain None.
    """

    id = fields.IntField(primary_key=True)
    input_name = fields.CharField(max_length=2048, db_index=True, unique=True)
    main_name = fields.CharField(max_length=2048, null=True)
    match_confidence = fields.FloatField(null=True)
    first_name = fields.CharField(max_length=2048, null=True)
    email = fields.CharField(max_length=2048, null=True)
    faculty = fields.ForeignKeyField(
        "models.Faculty",
        related_name="faculty_employees",
        null=True,
        to_field="abbreviation",
    )
    people_page_url = fields.CharField(max_length=2048, null=True)
    orgs = fields.ManyToManyField("models.Organization", related_name="org_employees")

    class Meta:
        table = "person_data"

    def __str__(self):
        return (
            self.main_name + f" ({self.faculty})" if self.main_name else self.input_name
        )


class Organization(Model, TimestampMixin):
    """
    'base' class for organizations, can be used for faculties, departments, etc.
    """

    id = fields.IntField(primary_key=True)
    parent_organization = fields.ForeignKeyField("models.Organization", null=True)
    hierarchy_level = fields.IntField()  # how much levels of parent orgs are above this one. E.g. 0 for the university, 1 for faculty, 2 for departments, 3 for groups.
    name = fields.CharField(max_length=2048, db_index=True)
    abbreviation = fields.CharField(
        max_length=255, db_index=True
    )  # the standalone abbreviation of this org, e.g. HMI
    full_abbreviation = fields.CharField(
        max_length=2048, db_index=True, unique=True
    )  # including the parent orgs abbreviations, e.g. EEMCS-CS-HMI

    class Meta:
        table = "organization_data"
        unique_together = ("name", "abbreviation")

    def __str__(self):
        return self.name + " (" + self.abbreviation + ")"


class Faculty(Organization):
    """
    Faculty data
    Same as Organization, just using a distinct name as it's used a lot.
    """

    abbreviation = fields.CharField(
        max_length=255, db_index=True, unique=True
    )  # the standalone abbreviation of this org, e.g. HMI


class Programme(Model, TimestampMixin):
    """
    Programme data (field 'Department' in raw copyright data)
    """

    faculty = fields.ForeignKeyField(
        "models.Faculty",
        related_name="faculty_programmes",
        to_field="abbreviation",
        null=True,
    )
    cluster = fields.CharField(max_length=2048, null=True)
    name = fields.CharField(max_length=2048, db_index=True)
    abbreviation = fields.CharField(max_length=255, db_index=True)
    programme_type = fields.CharField(
        max_length=255, null=True
    )  # bachelor, master, etc.

    class Meta:
        table = "programme_data"
        unique_together = ("name", "abbreviation")

    def __str__(self):
        return self.name + " (" + self.abbreviation + ")"


class PDF(Model, TimestampMixin):
    """
    data for a PDF file
    Each file should be directly related to a single CopyrightItem OR
    a single v1_CopyrightItem (if the PDF was downloaded based on the v1 sheet).
    """

    id = fields.IntField(primary_key=True)

    # one-to-one relation with parent item
    # exactly one of these should be non-null
    copyright_item = fields.OneToOneField(
        "models.CopyrightItem", related_name="pdf", null=True
    )
    v1_copyright_item = fields.OneToOneField(
        "models.v1_CopyrightItem", related_name="pdf", null=True
    )

    # one-to-one relation with Canvas metadata
    # should always exist, as it is used to retrieve the PDF in the first place
    canvas_metadata = fields.OneToOneField(
        "models.PDFCanvasMetadata", related_name="pdf"
    )

    # core fields: filename, url, material_id, size, retrieval date, current filename (as stored on disk)

    filename = fields.CharField(max_length=2048, null=True)  # original filename
    url = fields.CharField(max_length=2048, null=True)  # original url
    file_size = fields.IntField(null=True)  # in bytes
    retrieved_on = fields.DatetimeField(null=True, default=datetime.utcnow)
    current_file_name = fields.CharField(max_length=2048)

    # created fields: metadata (see which fields to include later), filehash (from CRC and also self-computed)

    author = fields.CharField(max_length=2048, null=True)
    title = fields.CharField(max_length=2048, null=True)
    subject = fields.CharField(max_length=2048, null=True)
    keywords = fields.JSONField(null=True)  # list of keywords
    producer = fields.CharField(max_length=2048, null=True)
    creation_date = fields.DatetimeField(null=True)
    mod_date = fields.DatetimeField(null=True)
    creator = fields.CharField(max_length=2048, null=True)
    summary = fields.CharField(max_length=10000, null=True)
    description = fields.CharField(max_length=10000, null=True)

    filehash = fields.CharField(max_length=255, null=True, db_index=True)

    # info about parsing (no attempt/success/failure, pages, length, ...)
    extraction_attempted = fields.BooleanField(default=False)
    extraction_successful = fields.BooleanField(default=False)

    num_pages = fields.IntField(null=True)
    num_words = fields.IntField(null=True)
    num_images = fields.IntField(null=True)

    # 1-to-1 relation to extracted text
    extracted_text = fields.OneToOneField(
        "models.PDFText", related_name="pdf", null=True
    )

    keywords = fields.JSONField(
        null=True
    )  # list of keywords extracted from text w/ confidence as a dict with key = keyword, value = confidence

    extracted_entities = fields.ManyToManyField(
        "models.Entity", through="pdf_entity", related_name="parent_pdf"
    )

    class Meta:
        table = "pdf_data"

    @property
    def path(self) -> Path:
        return SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full / self.current_file_name

    def as_file(self) -> File:
        return File(self.path)


class PDFCanvasMetadata(Model, TimestampMixin):
    """
    Metadata about a PDF file as retrieved from Canvas.
    """

    id = fields.IntField(primary_key=True)
    uuid = fields.CharField(max_length=255)
    folder_id = fields.IntField(null=True)

    display_name = fields.CharField(max_length=2048)
    filename = fields.CharField(max_length=2048)

    upload_status = fields.CharField(max_length=255)

    content_type = fields.CharField(max_length=255)
    mime_class = fields.CharField(max_length=255)
    category = fields.CharField(max_length=255)

    download_url = fields.CharField(max_length=2048)
    size = fields.IntField()  # in bytes
    thumbnail_url = fields.CharField(max_length=2048, null=True)

    canvas_created_at = fields.DatetimeField()
    canvas_updated_at = fields.DatetimeField()
    canvas_modified_at = fields.DatetimeField(null=True)

    locked = fields.BooleanField()
    hidden = fields.BooleanField()
    lock_at = fields.DatetimeField(null=True)
    unlock_at = fields.DatetimeField(null=True)
    visibility_level = fields.CharField(max_length=255)

    pdf: fields.ReverseRelation["PDF"]

    # data for user who uploaded the file
    user_id = fields.IntField(null=True)
    user_anonymous_id = fields.CharField(max_length=255, null=True)
    user_display_name = fields.CharField(max_length=2048, null=True)
    user_avatar_image_url = fields.CharField(max_length=2048, null=True)
    user_html_url = fields.CharField(max_length=2048, null=True)
    user_pronouns = fields.CharField(max_length=255, null=True)

    class Meta:
        table = "pdf_canvas_metadata"


class PDFText(Model, TimestampMixin):
    """
    Extracted text from a PDF file.
    """

    id = fields.IntField(primary_key=True)
    extracted_text = fields.TextField(null=True)
    num_pages = fields.IntField(null=True)
    text_quality = fields.FloatField(default=0)  # between 0 and 1
    is_ocr = fields.BooleanField(default=False)

    class Meta:
        table = "pdf_text_data"


class PDFEntity(Model):
    """
    Many-to-one relation to store extracted entities from a PDF file.
    """

    pdf = fields.ForeignKeyField("models.PDF", related_name="pdf_entities")
    entity = fields.ForeignKeyField("models.Entity", related_name="pdf_entities")


class Entity(Model, TimestampMixin):
    """
    Extracted entity from a PDF file.
    """

    id = fields.IntField(primary_key=True)
    label = fields.CharField(max_length=255)
    raw_text = fields.CharField(max_length=2048)
    canonical_form = fields.CharField(
        max_length=2048, null=True
    )  # if recognized as a known entity, the canonical form (e.g. full name)
    recognized = fields.BooleanField(
        default=False
    )  # whether the entity was recognized as a known entity (person/publisher/org/....)
    recognition_type = fields.CharEnumField(enum_type=EntityTypes)
    confidence = fields.FloatField(
        null=True
    )  # between 0 and 1. If 1, it was a precise match; e.g. by regex or lookup

    class Meta:
        table = "pdf_entity_data"


class StagedCopyrightItem(Model, TimestampMixin):
    """
    Staging table for raw data ingested from copyright export files.
    Fields are kept as simple as possible to accommodate raw data.

    The raw fields we expect:
        [
        "Material id",
        "Period",
        "Department",
        "Course code",
        "Course name",
        "url",
        "Filename",
        "Title",
        "Owner",
        "Filetype",
        "Classification",
        "Type",
        "ML Prediction",
        "Manual classification",
        "Manual identifier",
        "Scope",
        "Remarks",
        "Auditor",
        "Last change",
        "Status",
        "Google search file",
        "ISBN",
        "DOI",
        "In collection",
        "pagecount",
        "wordcount",
        "picturecount",
        "Author",
        "Publisher",
        "Reliability",
        "Pages * Students",
        "#students_registered"
    ]

    Before ingestion, these should be lowercased, spaces replaced with underscores, * replaced with x, and # replaced with count_.
    e.g. by calling standardize_dataframe in db.base

    """

    material_id = fields.IntField(primary_key=True)
    period = fields.CharField(max_length=255, null=True)
    department = fields.CharField(max_length=2048, null=True)
    course_code = fields.CharField(max_length=255, null=True)
    course_name = fields.CharField(max_length=2048, null=True)
    url = fields.CharField(max_length=255, null=True)
    filename = fields.CharField(max_length=2048, null=True)
    title = fields.CharField(max_length=2048, null=True)
    owner = fields.CharField(max_length=2048, null=True)
    filetype = fields.CharField(max_length=255, null=True)
    classification = fields.CharField(max_length=255, null=True)
    manual_classification = fields.CharField(max_length=2048, null=True)
    manual_identifier = fields.CharField(max_length=2048, null=True)
    scope = fields.CharField(max_length=255, null=True)
    remarks = fields.CharField(max_length=10000, null=True)
    ml_prediction = fields.CharField(max_length=255, null=True)
    isbn = fields.CharField(max_length=255, null=True)
    doi = fields.CharField(max_length=255, null=True)
    in_collection = fields.CharField(max_length=255, null=True)
    pagecount = fields.CharField(max_length=255, null=True)
    wordcount = fields.CharField(max_length=255, null=True)
    picturecount = fields.CharField(max_length=255, null=True)
    author = fields.CharField(max_length=255, null=True)
    publisher = fields.CharField(max_length=255, null=True)
    auditor = fields.CharField(max_length=10000, null=True)
    last_change = fields.DateField(null=True)
    status = fields.CharField(max_length=255, null=True)
    reliability = fields.CharField(max_length=255, null=True)
    pages_x_students = fields.CharField(max_length=255, null=True)
    count_students_registered = fields.CharField(max_length=255, null=True)
    retrieved_from_copyright_on = fields.DatetimeField(null=True)
    workflow_status = fields.CharField(max_length=255, null=True)
    faculty = fields.CharField(max_length=255, null=True)
    file_exists = fields.CharField(max_length=255, null=True)

    class Meta:
        table = "staged_copyright_item"


class StagedFacultyUpdate(Model, TimestampMixin):
    """
    Staging table for updates from faculty sheets.
    """

    material_id = fields.IntField(primary_key=True)
    manual_classification = fields.CharField(max_length=2048, null=True)
    v2_manual_classification = fields.CharField(max_length=255, null=True)
    v2_overnamestatus = fields.CharField(max_length=255, null=True)
    v2_lengte = fields.CharField(max_length=255, null=True)
    remarks = fields.CharField(max_length=10000, null=True)
    workflow_status = fields.CharField(max_length=255, null=True)

    class Meta:
        table = "staged_faculty_update"


class StagedProcessingFailure(Model, TimestampMixin):
    """
    Stores failures encountered while processing staged rows.
    Each row references the staged material_id (if available), the raw payload
    (as JSON), and an error message to aid debugging/retry.
    """

    id = fields.IntField(primary_key=True)
    material_id = fields.IntField(null=True, db_index=True)
    staged_payload = fields.JSONField(null=True)
    error_message = fields.CharField(max_length=2000, null=True)

    class Meta:
        table = "staged_processing_failures"
