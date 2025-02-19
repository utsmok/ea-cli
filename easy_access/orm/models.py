"""
defines classes for the ORM to store data in a database
"""
from tortoise.models import Model
from tortoise import fields
from enum import Enum
from easy_access.settings import SETTINGS
from easy_access.classification.classifier_api import CopyrightStatus, ItemType, AllowedUsageByUT

class TimestampMixin():
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

    IN_ONDERZOEK = "in onderzoek"
    VERWIJDERVERZOEK_VERSTUURD = "verwijderverzoek verstuurd"

class Filetype(Enum):
    PDF = "pdf"
    PPT = "ppt"
    DOCX = "docx"
    XLSX = "xlsx"
    MP4 = "mp4"
    JPG = "jpg"
    PNG = "png"
    UNKNOWN = "unknown"

class Status(Enum):
    PUBLISHED = "published"
    UNPUBLISHED = "unpublished"
    DELETED = "deleted"

class Faculty(Enum):
    BMS = "BMS"
    EEMCS = "EEMCS"
    ET = "ET"
    ITC = "ITC"
    TNW = "TNW"
    UNMAPPED = "Unmapped"

class WorkflowStatus(Enum):
    ToDo = "ToDo"
    Done = "Done"
    InProgress = "InProgress"

class Infringement(Enum):
    NO = "no"
    YES = "yes"
    UNDETERMINED = "undetermined"
"""
Programatically generate enums for years between 2020 and 2030 for valid periods using one of these formats:
YYYY-[12]{1}[AB]{1} (eg. 2022-1A or 2022-2B)
YYYY-3 (eg. 2022-3)
YYYY-SEM[12]{1} (eg. 2022-SEM1 or 2022-SEM2)
YYYY-JAAR (eg. 2022-JAAR)
"""
Period = Enum('Period', {f"{year}_{period}": f"{year}-{period}" for year in range(2020, 2031) for period in ["1A", "1B", "2A", "2B", "3", "SEM1", "SEM2", "JAAR"]})

# get departments from settings
Department = Enum('Department', {department: department for department in SETTINGS.university_settings.department_mapping.keys()})



class CopyrightItem(Model, TimestampMixin):
    """
    Core item in the dataset. Contains all the data for one item on Canvas.
    Contains all the data as exported from the copyrighttool.
    """

    material_id = fields.IntField(primary_key=True)
    period = fields.CharEnumField(enum_type=Period, max_length=255)
    department = fields.CharEnumField(enum_type=Department, max_length=2048, db_index=True)
    course_code = fields.CharField(max_length=255, db_index=True)
    course_name = fields.CharField(max_length=2048, db_index=True)
    url = fields.CharField(max_length=255)
    filename = fields.CharField(max_length=2048, db_index=True)
    title = fields.CharField(max_length=2048, null=True)
    owner = fields.CharField(max_length=2048, null=True)
    filetype = fields.CharEnumField(enum_type=Filetype, max_length=255, default=Filetype.PDF)
    classification = fields.CharEnumField(enum_type=Classification, max_length=255, default=Classification.LANGE_OVERNAME)
    ml_prediction = fields.CharEnumField(enum_type=Classification, max_length=255, db_index=True)
    manual_classification = fields.CharField(max_length=2048, null=True, db_index=True)
    manual_identifier = fields.CharField(max_length=2048, null=True)
    scope = fields.CharField(max_length=255, null=True)
    remarks = fields.CharField(max_length=10000, null=True)
    auditor = fields.CharField(max_length=10000, null=True)
    last_change = fields.DateField()
    status = fields.CharEnumField(enum_type=Status, max_length=255, default=Status.PUBLISHED, db_index=True)
    isbn = fields.CharField(max_length=255, null=True)
    doi = fields.CharField(max_length=255, null=True)
    in_collection = fields.BooleanField()
    pagecount = fields.IntField()
    wordcount = fields.IntField()
    picturecount = fields.IntField()
    author = fields.CharField(max_length=2048)
    publisher = fields.CharField(max_length=2048)
    reliability = fields.IntField()
    pages_students = fields.IntField()
    students_registered = fields.IntField()

    # Workflow data, added by the tool -- not present in the raw data!
    retrieved_from_copyright_on = fields.DatetimeField(null=True, db_index=True)
    workflow_status = fields.CharEnumField(enum_type=WorkflowStatus, max_length=255, default=WorkflowStatus.ToDo, db_index=True)
    possible_fine = fields.FloatField(null=True)
    infringement = fields.CharEnumField(enum_type=Infringement, max_length=255, default=Infringement.UNDETERMINED)

    # relations

    courses = fields.ManyToManyField('Course', related_name='course_items')
    faculty = fields.ForeignKeyField('Faculty', related_name='faculty_items')
    llm_classification = fields.OneToOneField('LLMClassification', related_name='item', null=True)

    class Meta:
        table = "copyright_data"


class Course(Model, TimestampMixin):
    """
    Data for a course; mainly from OSIRIS.
    """
    cursuscode = fields.IntField(primary_key=True)
    internal_id = fields.IntField()
    year = fields.IntField()  # use academic year; ie 2024-2025 --> 2024
    name = fields.CharField(max_length=2048)
    short_name = fields.CharField(max_length=255)
    faculty = fields.ForeignKeyField('Faculty', related_name='faculty_courses', null=True)
    ec = fields.IntField()
    programme = fields.CharField(max_length=2048) # use enumfield?? possibility of missing programmes in enum...
    notes = fields.CharField(max_length=10000)
    category = fields.CharField(max_length=2048)
    teachers = fields.ManyToManyRelation('Person', related_name='course_teacher')
    contacts = fields.ManyToManyField('Person', related_name='course_contact')
    docenten = fields.ManyToManyField('Person', related_name='course_docent')
    examinators = fields.ManyToManyField('Person', related_name='course_examinator')
    unknown_role = fields.ManyToManyField('Person', related_name='course_unknown_role')
    tutors = fields.ManyToManyField('Person', related_name='course_tutor')

    class Meta:
        table = "course_data"

class Person(Model, TimestampMixin):
    """
    Person data; data from people pages.
    Initialized with just an 'input_name'. If no match is found, the rest will remain None.
    """
    id = fields.IntField(primary_key=True)
    input_name = fields.CharField(max_length=2048, db_index=True)
    main_name = fields.CharField(max_length=2048, null=True)
    match_confidence = fields.FloatField(null=True)
    first_name = fields.CharField(max_length=2048, null=True) #'other_names' should be a list with len 1 containing only the first name
    email = fields.CharField(max_length=2048, null=True)
    faculty = fields.ForeignKeyField('Faculty', related_name='faculty_employees', null=True)
    people_page_url = fields.CharField(max_length=2048, null=True)

    class Meta:
        table = "person_data"

class LLMClassification(Model, TimestampMixin):
    """
    Additional classification data generated by an LLM to enrich items.
    """
    id = fields.IntField(primary_key=True)
    item = fields.OneToOneField('CopyrightItem')
    allowed_usage = fields.CharEnumField(enum_type=AllowedUsageByUT, max_length=255, default=AllowedUsageByUT.UNDETERMINED)
    allowed_usage_reasoning = fields.CharField(max_length=10000)
    copyright_status = fields.CharEnumField(enum_type=CopyrightStatus, max_length=255, default=CopyrightStatus.OTHER)
    copyright_classification_reason = fields.CharField(max_length=10000)
    item_type = fields.CharEnumField(enum_type=ItemType, max_length=255, default=ItemType.UNKNOWN)
    item_type_classification_reason = fields.CharField(max_length=10000)
    pdf_name = fields.CharField(max_length=2048)
    publisher_name = fields.CharField(max_length=2048)
    copyright_holder = fields.CharField(max_length=2048)
    item_title = fields.CharField(max_length=2048)
    pdf_page_count = fields.IntField()
    remarks = fields.CharField(max_length=10000)

    author_names = fields.JSONField()
    doi = fields.JSONField()
    isbn = fields.JSONField()
    source_url = fields.JSONField()
    license = fields.JSONField()
    topic = fields.JSONField()

    class Meta:
        table = "llm_classification_data"

class Organization(Model, TimestampMixin):
    """
    'base' class for organizations, can be used for faculties, departments, etc.
    """
    parent_organization = fields.ForeignKeyField('Organization', related_name='child_organizations')
    hierarchy_level = fields.IntField() # how much levels of parent orgs are above this one. E.g. 0 for the university, 1 for faculty, 2 for departments, 3 for groups.
    name = fields.CharField(max_length=2048, db_index=True)
    abbreviation = fields.CharField(max_length=255, db_index=True) # the standalone abbreviation of this org, e.g. HMI
    full_abbreviation = fields.CharField(max_length=2048, db_index=True) # including the parent orgs abbreviations, e.g. EEMCS-CS-HMI

    class Meta:
        table = "organization_data"

class Programme(Model, TimestampMixin):
    """
    Programme data (field 'Department' in raw copyright data)
    """
    faculty = fields.ForeignKeyField('Faculty', related_name='faculty_programmes')
    cluster = fields.CharField(max_length=2048)
    name = fields.CharField(max_length=2048, db_index=True)
    abbreviation = fields.CharField(max_length=255, db_index=True)
    programme_type = fields.CharField(max_length=255) # bachelor, master, etc.

    class Meta:
        table = "programme_data"
