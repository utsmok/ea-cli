"""
This module defines the Tortoise ORM models for the application's database.
These models represent various entities such as Copyright Items, Courses, Persons,
Faculties, LLM Classifications, and their relationships. Each class corresponds
to a database table.
"""

from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING  # For type checking an optional import like File

from tortoise import fields
from tortoise.models import Model

# Conditional import for type checking to avoid circular dependency if File uses models
if TYPE_CHECKING:
    from easy_access.utils import File  # Used in PDF model

from easy_access.classification.classifier_models import (
    AllowedUsageByUT,
    CopyrightStatus,
    ItemType,
)
from easy_access.settings import SETTINGS, DirSetting


class TimestampMixin:
    """
    A mixin class that adds `created_at` and `modified_at` timestamp fields
    to any model that inherits from it.
    `created_at` is set on creation, `modified_at` is updated on modification.
    """

    created_at: fields.DatetimeField = fields.DatetimeField(
        null=True, auto_now_add=True
    )
    modified_at: fields.DatetimeField = fields.DatetimeField(null=True, auto_now=True)


class Classification(Enum):
    """Manual classification options for copyright items."""

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
    """Enumeration of possible file types for copyright items."""

    PDF = "pdf"
    PPT = "ppt"
    DOC = "doc"
    XLSX = "xlsx"
    MP4 = "mp4"
    JPG = "jpg"
    PNG = "png"
    UNKNOWN = "unknown"
    FILE = "file"  # Generic file type


class Status(Enum):
    """Publication status of a copyright item."""

    PUBLISHED = "Published"
    UNPUBLISHED = "Unpublished"
    DELETED = "Deleted"


class WorkflowStatus(Enum):
    """Workflow status for internal processing of copyright items."""

    ToDo = "ToDo"
    Done = "Done"
    InProgress = "InProgress"


class Infringement(Enum):
    """Assessment of copyright infringement status."""

    YES = "yes"
    NO = "no"
    MAYBE = "maybe"
    UNDETERMINED = "undetermined"


# Dynamically generated Enums
# These might not be ideal for static analysis tools if SETTINGS isn't fully available at import time for them.
# However, Tortoise ORM will generate them at runtime.

_period_enum_values: dict[str, str] = {
    f"{year}_{period}": f"{year}-{period}"
    for year in range(2020, 2031)  # Example range, adjust as needed
    for period in ["1A", "1B", "2A", "2B", "3", "SEM1", "SEM2", "JAAR"]
}
Period = Enum("Period", _period_enum_values)
Period.__doc__ = "Dynamically generated enum for academic periods (e.g., 2023-1A)."


_department_enum_values: dict[str, str] = {
    # Ensure department names are valid enum member names (e.g., replace spaces, special chars)
    # For now, using a simplified approach; robust generation might need name sanitization.
    # Or, if SETTINGS.university_settings.department_mapping keys are simple enough:
    # department.replace(" ", "_").replace(":", "").replace(",", ""): department
    # For this example, assuming keys are simple or pre-sanitized if used directly.
    # This dynamic enum is problematic if keys are complex.
    # A fixed list or a separate table might be more robust for departments if names are complex.
    # For now, if SETTINGS are loaded, it will work. If not, this will be empty.
    # Using a placeholder if SETTINGS isn't fully resolved at this point for type checkers.
    # department: department
    # for department in (SETTINGS.university_settings.department_mapping if hasattr(SETTINGS, 'university_settings') else {"Unknown": "Unknown"})
    "Unknown": "Unknown"  # Placeholder if dynamic generation is tricky at import time for linters
}
if hasattr(SETTINGS, "university_settings"):  # Check if SETTINGS is initialized
    _department_enum_values = {
        # Sanitize keys for Enum member names
        name.replace(" ", "_")
        .replace(":", "")
        .replace(",", "")
        .replace("-", "_")
        .upper(): name
        for name in SETTINGS.university_settings.department_mapping.keys()
    }
    if not _department_enum_values:  # Ensure at least one member if mapping is empty
        _department_enum_values = {"UNKNOWN_DEPT": "Unknown Department"}

Department = Enum("Department", _department_enum_values)
Department.__doc__ = (
    "Dynamically generated enum for university departments, based on settings."
)


class CopyrightItem(Model, TimestampMixin):
    """
    Core model representing a copyright item, typically from a Canvas export.
    It includes metadata from the export and fields for internal workflow and analysis.
    """

    material_id: fields.IntField = fields.IntField(
        primary_key=True,
        description="Unique material identifier from the source system.",
    )
    period: fields.CharEnumField = fields.CharEnumField(
        enum_type=Period,
        max_length=255,
        description="Academic period of the course item.",
    )
    # department field might be better as a ForeignKey to a Programme or Department table if these are standardized.
    department: fields.CharField = fields.CharField(
        max_length=2048,
        db_index=True,
        description="Department associated with the course (often programme name).",
    )
    course_code: fields.CharField = fields.CharField(
        max_length=255,
        db_index=True,
        description="Course code, may not always be the Osiris code.",
    )
    course_name: fields.CharField = fields.CharField(
        max_length=2048, db_index=True, description="Name of the course."
    )
    url: fields.CharField = fields.CharField(
        max_length=2048,
        unique=True,
        null=True,
        description="URL of the copyrighted material.",
    )  # Max length increased
    filename: fields.CharField = fields.CharField(
        max_length=2048,
        db_index=True,
        null=True,
        description="Filename of the material.",
    )
    title: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="Title of the material, if available."
    )
    owner: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="Owner/uploader of the material."
    )
    filetype: fields.CharEnumField = fields.CharEnumField(
        enum_type=Filetype,
        max_length=255,
        default=Filetype.UNKNOWN,
        description="Detected file type of the material.",
    )
    # 'classification' might be the original system's classification.
    classification: fields.CharEnumField = fields.CharEnumField(
        enum_type=Classification,
        max_length=255,
        default=Classification.LANGE_OVERNAME,
        description="Original classification from the source system.",
    )
    ml_prediction: fields.CharEnumField = fields.CharEnumField(
        enum_type=Classification,
        max_length=255,
        db_index=True,
        null=True,
        description="Machine Learning based classification prediction.",
    )
    manual_classification: fields.CharField = fields.CharField(
        max_length=2048,
        null=True,
        db_index=True,
        description="Manually assigned classification by a user.",
    )
    manual_identifier: fields.CharField = fields.CharField(
        max_length=2048,
        null=True,
        description="Manual identifier like ISBN or DOI if applicable.",
    )
    scope: fields.CharField = fields.CharField(
        max_length=255, null=True, description="Scope of use (e.g., 'Always', 'Once')."
    )
    remarks: fields.CharField = fields.CharField(
        max_length=5128, null=True, db_index=True, description="User remarks or notes."
    )
    auditor: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="User who audited/checked the item."
    )  # Max length increased
    last_change: fields.DateField = fields.DateField(
        null=True, description="Date of last change in the source system."
    )
    status: fields.CharEnumField = fields.CharEnumField(
        enum_type=Status,
        max_length=255,
        default=Status.PUBLISHED,
        db_index=True,
        description="Publication status in the source system.",
    )
    isbn: fields.CharField = fields.CharField(
        max_length=255, null=True, description="ISBN of the material, if applicable."
    )
    doi: fields.CharField = fields.CharField(
        max_length=255, null=True, description="DOI of the material, if applicable."
    )
    in_collection: fields.BooleanField = fields.BooleanField(
        null=True,
        description="Indicates if the item is part of a specific collection (e.g., library).",
    )
    pagecount: fields.IntField = fields.IntField(
        description="Number of pages in the material."
    )
    wordcount: fields.IntField = fields.IntField(
        description="Word count of the material."
    )
    picturecount: fields.IntField = fields.IntField(
        description="Number of pictures in the material."
    )
    author: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="Author(s) of the material."
    )
    publisher: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="Publisher of the material."
    )
    reliability: fields.IntField = fields.IntField(
        description="A score indicating data reliability or confidence."
    )  # Clarify meaning
    pages_x_students: fields.IntField = fields.IntField(
        description="Product of pages and number of students."
    )
    count_students_registered: fields.IntField = fields.IntField(
        description="Number of students registered for the course/material."
    )

    # Workflow data, added by the tool
    retrieved_from_copyright_on: fields.DatetimeField = fields.DatetimeField(
        null=True,
        db_index=True,
        description="Timestamp when the item was last retrieved from the source system.",
    )
    workflow_status: fields.CharEnumField = fields.CharEnumField(
        enum_type=WorkflowStatus,
        max_length=255,
        default=WorkflowStatus.ToDo,
        db_index=True,
        description="Internal workflow status for processing.",
    )
    possible_fine: fields.FloatField = fields.FloatField(
        null=True, description="Calculated possible fine amount related to copyright."
    )
    infringement: fields.CharEnumField = fields.CharEnumField(
        enum_type=Infringement,
        max_length=255,
        default=Infringement.UNDETERMINED,
        description="Assessed copyright infringement status.",
    )

    # Relations
    courses: fields.ManyToManyRelation["Course"] = fields.ManyToManyField(
        "models.Course",
        related_name="course_items",
        description="Courses this item is associated with.",
    )
    faculty: fields.ForeignKeyRelation["Faculty"] = fields.ForeignKeyField(
        "models.Faculty",
        related_name="faculty_items",
        to_field="abbreviation",
        description="Faculty associated with this item.",
    )
    llm_classification: fields.OneToOneRelation["LLMClassification"] = (
        fields.OneToOneField(  # Changed to OneToOneRelation
            "models.LLMClassification",
            related_name="item",
            null=True,
            description="LLM-generated classification data for this item.",
        )
    )
    changes: fields.ManyToManyRelation["ItemUpdate"] = fields.ManyToManyField(
        "models.ItemUpdate",
        related_name="item_updates_for",
        description="Log of changes made to this item.",
    )  # Changed related_name

    is_duplicate: fields.BooleanField = fields.BooleanField(
        db_index=True,
        null=True,
        description="True if this item is a duplicate of another.",
    )
    replacement_id: fields.IntField = fields.IntField(
        null=True,
        db_index=True,
        description="Material ID of the original item if this is a duplicate.",
    )

    class Meta:
        table = "copyright_data"
        ordering = ["-retrieved_from_copyright_on", "-material_id"]

    def __str__(self) -> str:
        """String representation of the CopyrightItem."""
        return f"{self.filename or 'N/A'} ({self.material_id})"


class ItemUpdate(Model, TimestampMixin):
    """
    Stores a log of changes made to a CopyrightItem.
    Each instance represents a snapshot of modifications at a point in time.
    """

    id: fields.IntField = fields.IntField(primary_key=True)
    change_details: fields.JSONField = fields.JSONField(
        description="JSON blob detailing the changes made (e.g., field, old_value, new_value)."
    )
    # material_id here is not a ForeignKey but a direct integer to allow logging changes even if the item is deleted.
    # Or, it could be a non-constrained FK: item = fields.ForeignKeyField("models.CopyrightItem", null=True, on_delete=fields.SET_NULL)
    material_id: fields.IntField = fields.IntField(
        description="Material ID of the CopyrightItem this update refers to."
    )

    class Meta:
        table = "item_updates"
        ordering = ["-created_at"]

    def __str__(self) -> str:
        """String representation of the ItemUpdate."""
        return f"Update for {self.material_id} at {self.created_at.strftime('%Y-%m-%d %H:%M:%S') if self.created_at else 'N/A'}"


class MissingCourse(Model, TimestampMixin):
    """
    Tracks OSIRIS course codes that were encountered but not found in the Course table,
    to be potentially retrieved and added later.
    """

    cursuscode: fields.IntField = fields.IntField(
        primary_key=True, description="The OSIRIS course code that is missing."
    )

    class Meta:
        table = "missing_courses"

    def __str__(self) -> str:
        """String representation of the MissingCourse."""
        return str(self.cursuscode)


class Course(Model, TimestampMixin):
    """
    Represents a university course, typically imported from a system like OSIRIS.
    """

    cursuscode: fields.IntField = fields.IntField(
        primary_key=True, description="Primary OSIRIS course code."
    )
    internal_id: fields.IntField = fields.IntField(
        unique=True,
        description="Internal system ID for the course, if different from cursuscode.",
    )
    year: fields.IntField = fields.IntField(
        description="Academic year of the course (e.g., 2024 for 2024-2025)."
    )
    name: fields.CharField = fields.CharField(
        max_length=2048, description="Full name of the course."
    )
    short_name: fields.CharField = fields.CharField(
        max_length=255,
        null=True,
        description="Short name or abbreviation for the course.",
    )
    faculty: fields.ForeignKeyRelation["Faculty"] = fields.ForeignKeyField(
        "models.Faculty",
        related_name="faculty_courses",
        null=True,
        to_field="abbreviation",
        description="Faculty offering the course.",
    )
    ec: fields.IntField = fields.IntField(
        null=True, description="Number of ECTS credits for the course."
    )
    # 'programme' could be a ForeignKey to a Programme table if programmes are managed entities.
    programme: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="Programme(s) this course belongs to."
    )
    notes: fields.TextField = fields.TextField(
        null=True, description="Additional notes or description for the course."
    )  # Changed to TextField
    category: fields.CharField = fields.CharField(
        max_length=2048,
        null=True,
        description="Category of the course (e.g., Bachelor, Master).",
    )

    # M2M relation defined via CourseEmployee
    # teachers: fields.ManyToManyRelation["Person"]

    class Meta:
        table = "course_data"
        ordering = ["year", "name"]

    def __str__(self) -> str:
        """String representation of the Course."""
        return f"{self.name} ({self.cursuscode}) - {self.year}"


class CourseEmployee(Model, TimestampMixin):
    """
    Through-table for the many-to-many relationship between Course and Person,
    specifying the role of a person in a course (e.g., teacher, contact).
    """

    id: fields.IntField = fields.IntField(primary_key=True)
    course: fields.ForeignKeyRelation[Course] = fields.ForeignKeyField(
        "models.Course", related_name="course_employees_roles"
    )  # Changed related_name
    person: fields.ForeignKeyRelation["Person"] = fields.ForeignKeyField(
        "models.Person", related_name="person_courses_roles"
    )  # Changed related_name
    role: fields.CharField = fields.CharField(
        max_length=255,
        null=True,
        description="Role of the person in the course (e.g., teacher, contact, tutor).",
    )  # Max length reduced

    class Meta:
        table = "course_employee"
        unique_together = (
            "course",
            "person",
            "role",
        )  # Ensure a person doesn't have same role twice for same course
        ordering = ["course__cursuscode", "person__main_name", "role"]

    def __str__(self) -> str:
        """String representation of the CourseEmployee relation."""
        return (
            f"Course: {self.course_id} - Person: {self.person_id} - Role: {self.role}"
        )


class Person(Model, TimestampMixin):
    """
    Represents a person (e.g., employee, staff) with data often sourced from
    university directories or people pages.
    """

    id: fields.IntField = fields.IntField(primary_key=True)
    input_name: fields.CharField = fields.CharField(
        max_length=2048,
        db_index=True,
        unique=True,
        description="The name as it originally appeared in the source data.",
    )
    main_name: fields.CharField = fields.CharField(
        max_length=2048,
        null=True,
        description="Standardized or primary name of the person.",
    )
    match_confidence: fields.FloatField = fields.FloatField(
        null=True, description="Confidence score if name matching was performed."
    )
    first_name: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="First name of the person."
    )
    email: fields.CharField = fields.CharField(
        max_length=2048,
        null=True,
        unique=True,
        description="Email address of the person.",
    )  # Added unique=True
    faculty: fields.ForeignKeyRelation["Faculty"] = fields.ForeignKeyField(
        "models.Faculty",
        related_name="faculty_employees",
        null=True,
        to_field="abbreviation",
        description="Faculty affiliation of the person.",
    )
    people_page_url: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="URL to the person's profile page."
    )
    orgs: fields.ManyToManyRelation["Organization"] = fields.ManyToManyField(
        "models.Organization",
        related_name="org_employees",
        description="Organizational units this person belongs to.",
    )

    # M2M relation defined via CourseEmployee
    courses_roles: fields.ReverseRelation[
        "CourseEmployee"
    ]  # Access through related_name 'person_courses_roles'

    class Meta:
        table = "person_data"
        ordering = ["main_name", "input_name"]

    def __str__(self) -> str:
        """String representation of the Person."""
        if self.main_name:
            return f"{self.main_name} ({self.faculty_id if self.faculty_id else 'N/A Faculty'})"
        return self.input_name


class LLMClassification(Model, TimestampMixin):
    """
    Stores additional classification data for a CopyrightItem, typically generated by an LLM.
    """

    id: fields.IntField = fields.IntField(primary_key=True)
    allowed_usage: fields.CharEnumField = fields.CharEnumField(
        enum_type=AllowedUsageByUT,
        max_length=255,
        default=AllowedUsageByUT.UNDETERMINED,
        description="LLM assessment of allowed usage.",
    )
    allowed_usage_reasoning: fields.TextField = fields.TextField(
        description="LLM reasoning for allowed usage assessment."
    )  # Changed to TextField
    copyright_status: fields.CharEnumField = fields.CharEnumField(
        enum_type=CopyrightStatus,
        max_length=255,
        default=CopyrightStatus.OTHER,
        description="LLM assessment of copyright status.",
    )
    copyright_classification_reason: fields.TextField = fields.TextField(
        description="LLM reasoning for copyright status assessment."
    )  # Changed to TextField
    item_type: fields.CharEnumField = fields.CharEnumField(
        enum_type=ItemType,
        max_length=255,
        default=ItemType.UNKNOWN,
        description="LLM assessment of item type.",
    )
    item_type_classification_reason: fields.TextField = fields.TextField(
        description="LLM reasoning for item type assessment."
    )  # Changed to TextField
    pdf_name: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="Filename of the PDF processed by LLM."
    )
    publisher_name: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="Publisher name identified by LLM."
    )
    copyright_holder: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="Copyright holder identified by LLM."
    )
    item_title: fields.CharField = fields.CharField(
        max_length=2048,
        null=True,
        description="Item title identified or refined by LLM.",
    )
    pdf_page_count: fields.IntField = fields.IntField(
        null=True, description="Page count of PDF as determined by LLM processing."
    )
    remarks: fields.TextField = fields.TextField(
        null=True, description="Additional remarks from LLM."
    )  # Changed to TextField

    # JSONFields can store lists or structured data extracted by LLM
    author_names: fields.JSONField = fields.JSONField(
        null=True, description="Author names identified by LLM (stored as JSON list)."
    )
    doi: fields.JSONField = fields.JSONField(
        null=True, description="DOI(s) identified by LLM (stored as JSON list)."
    )
    isbn: fields.JSONField = fields.JSONField(
        null=True, description="ISBN(s) identified by LLM (stored as JSON list)."
    )
    source_url: fields.JSONField = fields.JSONField(
        null=True, description="Source URL(s) identified by LLM (stored as JSON list)."
    )
    license: fields.JSONField = fields.JSONField(
        null=True,
        description="License information identified by LLM (stored as JSON list).",
    )
    topic: fields.JSONField = fields.JSONField(
        null=True, description="Topic(s) identified by LLM (stored as JSON list)."
    )

    used_material_id: fields.IntField = fields.IntField(
        unique=True,
        description="Material ID of the CopyrightItem this classification refers to.",
    )  # Added unique=True
    item: fields.OneToOneNullableRelation[
        CopyrightItem
    ]  # Defined in CopyrightItem as 'llm_classification'

    class Meta:
        table = "llm_classification_data"
        ordering = ["-modified_at"]

    def __str__(self) -> str:
        """String representation of the LLMClassification."""
        return f"LLM Data for Item ID {self.used_material_id}"


class Organization(Model, TimestampMixin):
    """
    Represents an organizational unit within the university hierarchy (e.g., university, faculty, department).
    """

    id: fields.IntField = fields.IntField(primary_key=True)
    parent_organization: fields.ForeignKeyNullableRelation["Organization"] = (
        fields.ForeignKeyField(
            "models.Organization",
            related_name="child_organizations",
            null=True,
            on_delete=fields.SET_NULL,
            description="Parent organizational unit in the hierarchy.",
        )
    )  # Added on_delete
    child_organizations: fields.ReverseRelation["Organization"]
    hierarchy_level: fields.IntField = fields.IntField(
        description="Level in the hierarchy (e.g., 0 for university, 1 for faculty)."
    )
    name: fields.CharField = fields.CharField(
        max_length=2048,
        db_index=True,
        description="Full name of the organizational unit.",
    )
    abbreviation: fields.CharField = fields.CharField(
        max_length=255,
        db_index=True,
        description="Abbreviation for this specific unit (e.g., HMI).",
    )
    full_abbreviation: fields.CharField = fields.CharField(
        max_length=2048,
        db_index=True,
        unique=True,
        null=True,
        description="Full path abbreviation (e.g., EEMCS-CS-HMI).",
    )  # Added null=True as root org might not have this in same way

    class Meta:
        table = "organization_data"
        unique_together = (
            "name",
            "abbreviation",
            "parent_organization",
        )  # More robust uniqueness
        ordering = ["hierarchy_level", "name"]

    def __str__(self) -> str:
        """String representation of the Organization."""
        return f"{self.name} ({self.abbreviation or self.full_abbreviation or 'N/A'})"


class Faculty(Organization):
    """
    Represents a university faculty, inheriting from Organization.
    The 'abbreviation' field is made unique for Faculties as they are often primary identifiers.
    """

    # Tortoise inherits fields, but we can override them or add constraints if needed.
    # Making abbreviation unique specifically for Faculty if it wasn't in Organization.
    # However, Tortoise doesn't directly support changing constraints of inherited fields this way.
    # This is more of a conceptual distinction; the table 'faculty' would be 'organization_data'
    # unless table_inheritance=False is used (which is not default).
    # For simplicity, if Faculty is just an Organization with hierarchy_level=1, this model might be redundant
    # unless specific Faculty-only fields or relations are added.
    # For now, keeping it as a proxy model.
    # If 'abbreviation' should be unique for Faculties, it's better enforced at Organization level with checks or a different structure.
    # For this pass, I'll assume it's a conceptual alias.
    # No, Tortoise creates a separate table for inherited models by default.

    # faculty_id = fields.IntField(primary_key=True) # If it needs its own PK
    # organization_ptr = fields.OneToOneField("models.Organization", pk=True, related_name="faculty_profile") # Explicit inheritance

    # If abbreviation needs to be unique for Faculty, it should be defined here.
    # However, if it shares the same table via table_inheritance=True (default for concrete inheritance),
    # then the constraint is on the base table.
    # If it's a separate table (multi-table inheritance), then this works:
    abbreviation: fields.CharField = fields.CharField(
        max_length=255,
        db_index=True,
        unique=True,
        description="Unique abbreviation for the faculty.",
    )

    class Meta:
        table = "faculty_data"  # Explicitly define a different table name for clarity or if needed

    def __str__(self) -> str:
        return f"Faculty: {self.name} ({self.abbreviation})"


class Programme(Model, TimestampMixin):
    """
    Represents an academic programme within a faculty.
    """

    id: fields.IntField = fields.IntField(primary_key=True)  # Added primary key
    faculty: fields.ForeignKeyRelation[Faculty] = fields.ForeignKeyField(
        "models.Faculty",
        related_name="faculty_programmes",
        to_field="abbreviation",
        null=True,
        description="Faculty this programme belongs to.",
    )
    cluster: fields.CharField = fields.CharField(
        max_length=2048,
        null=True,
        description="Cluster or grouping this programme might belong to.",
    )
    name: fields.CharField = fields.CharField(
        max_length=2048, db_index=True, description="Full name of the programme."
    )
    abbreviation: fields.CharField = fields.CharField(
        max_length=255,
        db_index=True,
        null=True,
        description="Abbreviation for the programme.",
    )  # Made null=True
    programme_type: fields.CharField = fields.CharField(
        max_length=30,
        null=True,
        description="Type of programme (e.g., Bachelor, Master).",
    )  # Max length reduced

    class Meta:
        table = "programme_data"
        unique_together = (
            "name",
            "faculty",
        )  # Programme name should be unique within a faculty
        ordering = ["faculty__abbreviation", "name"]

    def __str__(self) -> str:
        """String representation of the Programme."""
        return f"{self.name} ({self.abbreviation or 'N/A'})"


class PDF(Model, TimestampMixin):
    """
    Stores metadata and extracted information for a PDF file linked to a CopyrightItem.
    """

    material_id: fields.IntField = fields.IntField(
        primary_key=True,
        description="Material ID of the CopyrightItem this PDF belongs to.",
    )
    current_file_name: fields.CharField = fields.CharField(
        max_length=2048,
        description="Current filename on disk (may include path elements or be relative to a base PDF directory).",
    )

    replace_with: fields.ForeignKeyNullableRelation["PDF"] = fields.ForeignKeyField(
        "models.PDF",
        related_name="replacement_for_set",
        null=True,
        on_delete=fields.SET_NULL,
        description="If this PDF is a duplicate, this points to the original PDF item.",
    )  # Changed related_name
    replacement_for_set: fields.ReverseRelation["PDF"]  # PDFs that this PDF replaces

    extracted_text: fields.TextField = fields.TextField(
        null=True, description="Full extracted text from the PDF."
    )
    extracted_text_max_pages: fields.IntField = fields.IntField(
        null=True, description="Number of pages processed for text extraction."
    )
    extracted_text_max_length: fields.IntField = fields.IntField(
        null=True, description="Maximum length of extracted text stored (if truncated)."
    )

    original_file_name: fields.CharField = fields.CharField(
        max_length=2048,
        null=True,
        description="Original filename as uploaded or provided.",
    )
    original_page_count: fields.IntField = fields.IntField(
        null=True, description="Original page count from document metadata."
    )
    author: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="Author from PDF metadata."
    )
    file_modification_date: fields.DatetimeField = fields.DatetimeField(
        null=True, description="Modification date from PDF metadata."
    )
    file_creation_date: fields.DatetimeField = fields.DatetimeField(
        null=True, description="Creation date from PDF metadata."
    )
    producer: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="PDF producer from metadata."
    )
    creator: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="PDF creator tool from metadata."
    )
    subject: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="Subject from PDF metadata."
    )
    title: fields.CharField = fields.CharField(
        max_length=2048, null=True, description="Title from PDF metadata."
    )

    parsing_failed: fields.BooleanField = fields.BooleanField(
        null=True,
        default=False,
        description="True if parsing the PDF metadata or text failed.",
    )

    class Meta:
        table = "pdf_data"
        ordering = ["material_id"]

    @property
    def path(self) -> Path:
        """Constructs the full path to the PDF file on disk."""
        return SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full / self.current_file_name

    @property
    def age(self) -> int:
        """
        Calculates the age of the file in seconds based on available timestamps.
        Order of preference: file_modification_date, file_creation_date, record's modified_at.
        """
        now_ts: float = datetime.now().timestamp()
        file_date_ts: float | None = None

        if self.file_modification_date:
            file_date_ts = self.file_modification_date.timestamp()
        elif self.file_creation_date:
            file_date_ts = self.file_creation_date.timestamp()
        elif self.modified_at:  # Fallback to record modification time
            file_date_ts = self.modified_at.timestamp()

        if file_date_ts is not None:
            return int(now_ts - file_date_ts)
        return -1  # Indicate age cannot be determined

    def as_file(self) -> "File":
        """
        Returns a `File` utility object representing this PDF.
        Requires `easy_access.utils.File` to be importable.
        """
        from easy_access.utils import (
            File,  # Local import to avoid circularity at top level
        )

        return File(self.path)

    def __str__(self) -> str:
        """String representation of the PDF model instance."""
        return f"{self.current_file_name} (Material ID: {self.material_id})"
