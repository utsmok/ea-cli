"""
ORM models for the database.
"""

from datetime import datetime
from enum import Enum
from pathlib import Path

from tortoise import fields
from tortoise.models import Model

from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import File


class TimestampMixin:
    created_at = fields.DatetimeField(null=True, auto_now_add=True)
    modified_at = fields.DatetimeField(null=True, auto_now=True)


class Classification(Enum):
    OPEN_ACCESS = "open access"
    KORTE_OVERNAME = "korte overname"
    MIDDELLANGE_OVERNAME = "middellange overname"
    LANGE_OVERNAME = "lange overname"

    EIGEN_MATERIAAL_POWERPOINT = "eigen materiaal - powerpoint"
    EIGEN_MATERIAAL_TITELINDICATIE = "eigen materiaal - titelindicatie"
    EIGEN_MATERIAAL_OVERIG = "eigen materiaal - overig"
    EIGEN_MATERIAAL = "eigen materiaal"

    ONBEKEND = "onbekend"
    NIET_GEANALYSEERD = "niet geanalyseerd"
    IN_ONDERZOEK = "in onderzoek"
    VERWIJDERVERZOEK_VERSTUURD = "verwijderverzoek verstuurd"
    LICENTIE_BESCHIKBAAR = "licentie beschikbaar"


class Filetype(Enum):
    PDF = "pdf"
    PPT = "ppt"
    DOC = "doc"
    XLSX = "xlsx"
    MP4 = "mp4"
    JPG = "jpg"
    PNG = "png"
    UNKNOWN = "unknown"
    FILE = "file"


class Status(Enum):
    PUBLISHED = "Published"
    UNPUBLISHED = "Unpublished"
    DELETED = "Deleted"


class WorkflowStatus(Enum):
    ToDo = "ToDo"
    Done = "Done"
    InProgress = "InProgress"


class Infringement(Enum):
    YES = "yes"
    NO = "no"
    MAYBE = "maybe"
    UNDETERMINED = "undetermined"


"""
Programatically generate enums for years between 2020 and 2030 for valid periods using one of these formats:
YYYY-[12]{1}[AB]{1} (eg. 2022-1A or 2022-2B)
YYYY-3 (eg. 2022-3)
YYYY-SEM[12]{1} (eg. 2022-SEM1 or 2022-SEM2)
YYYY-JAAR (eg. 2022-JAAR)
"""
Period = Enum(
    "Period",
    {
        f"{year}_{period}": f"{year}-{period}"
        for year in range(2020, 2031)
        for period in ["1A", "1B", "2A", "2B", "3", "SEM1", "SEM2", "JAAR"]
    },
)
Department = Enum(
    "Department",
    {
        department: department
        for department in SETTINGS.university_settings.department_mapping
    },
)


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

    # relations

    courses = fields.ManyToManyField("models.Course", related_name="course_items")
    faculty = fields.ForeignKeyField(
        "models.Faculty", related_name="faculty_items", to_field="abbreviation"
    )
    changes = fields.ManyToManyField("models.ItemUpdate", related_name="item")

    is_duplicate = fields.BooleanField(
        db_index=True, null=True
    )  # if this item is a duplicate of another item -- determined by comparing PDFs
    replacement_id = fields.IntField(
        null=True, db_index=True
    )  # if this item is a duplicate, this field has the material_id of the original item

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
    ec = fields.IntField(null=True)
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
        through="models.CourseEmployee",
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
    first_name = fields.CharField(
        max_length=2048, null=True
    )  #'other_names' should be a list with len 1 containing only the first name
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
    Each file should be directly related to a CopyrightItem
    """

    material_id = fields.IntField(primary_key=True)
    current_file_name = fields.CharField(
        max_length=2048
    )  # this is also used to get the path to the file, see self.path()
    replace_with = fields.ForeignKeyField(
        "models.PDF", related_name="replacement_for", null=True
    )  # if this file is a duplicate, this field points to the original file
    replacement_for: fields.ReverseRelation["PDF"]
    extracted_text = fields.TextField(null=True)
    extracted_text_max_pages = fields.IntField(
        null=True
    )  # how many pages were processed to get the extracted text
    extracted_text_max_length = fields.IntField(
        null=True
    )  # how many characters were kept from the extracted text

    original_file_name = fields.CharField(max_length=2048, null=True)
    original_page_count = fields.IntField(null=True)
    author = fields.CharField(max_length=2048, null=True)
    file_modification_date = fields.DatetimeField(null=True)
    file_creation_date = fields.DatetimeField(null=True)
    producer = fields.CharField(max_length=2048, null=True)
    creator = fields.CharField(max_length=2048, null=True)
    subject = fields.CharField(max_length=2048, null=True)
    title = fields.CharField(max_length=2048, null=True)

    parsing_failed = fields.BooleanField(
        null=True, default=False
    )  # if the file could not be parsed, set to True

    # Download-related fields
    # -> Has an download been attempted?
    # defaults to false for new items. Use to create a list of items to download.
    # Set to True once an initial attempt has been made, then never change again.
    # -> Did the download succeed?
    # The function that downloads the file sets this to True if the files is downloaded and is >0kb. Else it sets it to False.
    # Defaults to None (if no download attempts have been made yet).

    download_attempted = fields.BooleanField(null=True, default=False)
    download_succeeded = fields.BooleanField(null=True, default=None)

    class Meta:
        table = "pdf_data"

    @property
    def path(self) -> Path:
        return SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full / self.current_file_name

    @property
    def age(self) -> int:
        """
        Depending on which date info is available, determine the age of the file in seconds.
        Start with 'file_modification_date', then try 'file_creation_date'.
        If neither are available, use 'modified_at', which should always be present.
        """
        now = datetime.now().timestamp()
        if self.file_modification_date:
            return int(now - self.file_modification_date.timestamp())
        elif self.file_creation_date:
            return int(now - self.file_creation_date.timestamp())
        else:
            return int(now - self.modified_at.timestamp())

    def as_file(self) -> File:
        return File(self.path)

    def __str__(self):
        return self.current_file_name + " (" + str(self.material_id) + ")"


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
