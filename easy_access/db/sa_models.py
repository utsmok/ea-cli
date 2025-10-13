from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from sqlalchemy import (
    JSON,
    Boolean,
    Date,
    DateTime,
    Float,
    ForeignKey,
    Integer,
    String,
    Text,
    UniqueConstraint,
)
from sqlalchemy import (
    Enum as SAEnum,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from easy_access.settings import SETTINGS

from .enums import (
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
from .models_base import Base


class Organization(Base):
    """
    'base' class for organizations, can be used for faculties, departments, etc.
    """

    __tablename__ = "organization_data"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    parent_organization_id: Mapped[int | None] = mapped_column(
        Integer, ForeignKey("organization_data.id"), nullable=True
    )
    hierarchy_level: Mapped[int] = mapped_column(
        Integer
    )  # how much levels of parent orgs are above this one. E.g. 0 for the university, 1 for faculty, 2 for departments, 3 for groups.
    name: Mapped[str] = mapped_column(String(2048), index=True)
    abbreviation: Mapped[str] = mapped_column(
        String(255), index=True
    )  # the standalone abbreviation of this org, e.g. HMI
    full_abbreviation: Mapped[str] = mapped_column(
        String(2048), index=True, unique=True
    )  # including the parent orgs abbreviations, e.g. EEMCS-CS-HMI

    __table_args__ = (
        UniqueConstraint("name", "abbreviation", name="uq_organization_name_abbr"),
    )

    # relationships
    parent_organization = relationship(
        "Organization", remote_side=[id], backref="child_organizations"
    )

    def __str__(self) -> str:  # pragma: no cover - helper
        return f"{self.name} ({self.abbreviation})"


class Faculty(Organization):
    """
    Faculty data
    Same as Organization, just using a distinct name as it's used a lot.
    """

    # Inherits all fields from Organization including abbreviation

    __mapper_args__ = {
        "polymorphic_identity": "faculty",
    }


class CourseEmployee(Base):
    __tablename__ = "course_employee"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    course_cursuscode: Mapped[int] = mapped_column(
        Integer, ForeignKey("course_data.cursuscode")
    )
    person_id: Mapped[int] = mapped_column(Integer, ForeignKey("person_data.id"))
    role: Mapped[str | None] = mapped_column(String(2048), nullable=True)


class Course(Base):
    __tablename__ = "course_data"

    cursuscode: Mapped[int] = mapped_column(Integer, primary_key=True)
    internal_id: Mapped[int] = mapped_column(Integer, unique=True)
    year: Mapped[int] = mapped_column(Integer, nullable=False)
    name: Mapped[str] = mapped_column(String(2048), nullable=False)
    short_name: Mapped[str | None] = mapped_column(String(255), nullable=True)
    ec: Mapped[float | None] = mapped_column(Float, nullable=True)
    programme: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    notes: Mapped[str | None] = mapped_column(String(10000), nullable=True)
    category: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    modified_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    teachers = relationship(
        "Person", secondary="course_employee", back_populates="courses"
    )

    def __str__(self) -> str:  # pragma: no cover - simple helper
        return f"{self.name} ({self.cursuscode})"


class Person(Base):
    __tablename__ = "person_data"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    input_name: Mapped[str] = mapped_column(String(2048), unique=True, index=True)
    main_name: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    match_confidence: Mapped[float | None] = mapped_column(Float, nullable=True)
    first_name: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    email: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    people_page_url: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    modified_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    courses = relationship(
        "Course", secondary="course_employee", back_populates="teachers"
    )

    def __str__(self) -> str:  # pragma: no cover - helper
        return self.main_name or self.input_name


class ItemUpdate(Base):
    __tablename__ = "item_updates"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    change_details: Mapped[dict] = mapped_column(JSON, nullable=False)
    material_id: Mapped[int] = mapped_column(Integer, nullable=False)
    created_at: Mapped[datetime] = mapped_column(
        DateTime, nullable=True, default=datetime.utcnow
    )


class MissingCourse(Base):
    __tablename__ = "missing_courses"

    cursuscode: Mapped[int] = mapped_column(Integer, primary_key=True)
    modified_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)


class Programme(Base):
    __tablename__ = "programme_data"

    id: Mapped[int] = mapped_column(
        Integer, primary_key=True
    )  # Assuming added, as not in Tortoise
    cluster: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    name: Mapped[str] = mapped_column(String(2048), index=True)
    abbreviation: Mapped[str] = mapped_column(String(255), index=True)
    programme_type: Mapped[str | None] = mapped_column(String(255), nullable=True)
    modified_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    __table_args__ = (
        UniqueConstraint("name", "abbreviation", name="uq_programme_name_abbr"),
    )


class PDFCanvasMetadata(Base):
    __tablename__ = "pdf_canvas_metadata"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    uuid: Mapped[str] = mapped_column(String(255), nullable=False)
    folder_id: Mapped[int | None] = mapped_column(Integer, nullable=True)

    display_name: Mapped[str] = mapped_column(String(2048), nullable=False)
    filename: Mapped[str] = mapped_column(String(2048), nullable=False)

    upload_status: Mapped[str] = mapped_column(String(255), nullable=False)

    content_type: Mapped[str] = mapped_column(String(255), nullable=False)
    mime_class: Mapped[str] = mapped_column(String(255), nullable=False)
    category: Mapped[str] = mapped_column(String(255), nullable=False)

    download_url: Mapped[str] = mapped_column(String(2048), nullable=False)
    size: Mapped[int] = mapped_column(Integer, nullable=False)  # in bytes
    thumbnail_url: Mapped[str | None] = mapped_column(String(2048), nullable=True)

    canvas_created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False)
    canvas_updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False)
    canvas_modified_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    locked: Mapped[bool] = mapped_column(Boolean, nullable=False)
    hidden: Mapped[bool] = mapped_column(Boolean, nullable=False)
    lock_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    unlock_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    visibility_level: Mapped[str] = mapped_column(String(255), nullable=False)

    # user fields
    user_id: Mapped[int | None] = mapped_column(Integer, nullable=True)
    user_anonymous_id: Mapped[str | None] = mapped_column(String(255), nullable=True)
    user_display_name: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    user_avatar_image_url: Mapped[str | None] = mapped_column(
        String(2048), nullable=True
    )
    user_html_url: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    user_pronouns: Mapped[str | None] = mapped_column(String(255), nullable=True)


class PDFText(Base):
    __tablename__ = "pdf_text_data"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    extracted_text: Mapped[str | None] = mapped_column(Text, nullable=True)
    num_pages: Mapped[int | None] = mapped_column(Integer, nullable=True)
    text_quality: Mapped[float] = mapped_column(Float, nullable=False, default=0)


class Entity(Base):
    __tablename__ = "pdf_entity_data"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    label: Mapped[str] = mapped_column(String(255), nullable=False)
    raw_text: Mapped[str] = mapped_column(String(2048), nullable=False)
    canonical_form: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    recognized: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    recognition_type: Mapped[EntityTypes] = mapped_column(
        SAEnum(EntityTypes), nullable=False
    )
    confidence: Mapped[float | None] = mapped_column(Float, nullable=True)


class PDFEntity(Base):
    __tablename__ = "pdf_entity"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)  # Assuming added
    pdf_id: Mapped[int] = mapped_column(Integer, ForeignKey("pdf_data.id"))
    entity_id: Mapped[int] = mapped_column(Integer, ForeignKey("pdf_entity_data.id"))


class PDF(Base):
    __tablename__ = "pdf_data"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)

    copyright_item_id: Mapped[int | None] = mapped_column(
        Integer, ForeignKey("copyright_data.material_id"), nullable=True
    )
    v1_copyright_item_id: Mapped[int | None] = mapped_column(
        Integer, ForeignKey("v1_copyright_item.material_id"), nullable=True
    )
    canvas_metadata_id: Mapped[int] = mapped_column(
        Integer, ForeignKey("pdf_canvas_metadata.id"), nullable=False
    )

    filename: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    url: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    file_size: Mapped[int | None] = mapped_column(Integer, nullable=True)
    retrieved_on: Mapped[datetime | None] = mapped_column(
        DateTime, nullable=True, default=datetime.utcnow
    )
    current_file_name: Mapped[str] = mapped_column(String(2048), nullable=False)

    author: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    title: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    subject: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    keywords: Mapped[list | None] = mapped_column(JSON, nullable=True)
    producer: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    creation_date: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    mod_date: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    creator: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    summary: Mapped[str | None] = mapped_column(String(10000), nullable=True)
    description: Mapped[str | None] = mapped_column(String(10000), nullable=True)

    filehash: Mapped[str | None] = mapped_column(String(255), nullable=True, index=True)

    extraction_attempted: Mapped[bool] = mapped_column(
        Boolean, nullable=False, default=False
    )
    extraction_successful: Mapped[bool] = mapped_column(
        Boolean, nullable=False, default=False
    )

    num_pages: Mapped[int | None] = mapped_column(Integer, nullable=True)
    num_words: Mapped[int | None] = mapped_column(Integer, nullable=True)
    num_images: Mapped[int | None] = mapped_column(Integer, nullable=True)

    extracted_text_id: Mapped[int | None] = mapped_column(
        Integer, ForeignKey("pdf_text_data.id"), nullable=True
    )

    # relationships
    canvas_metadata = relationship("PDFCanvasMetadata", backref="pdf")
    extracted_text = relationship("PDFText", backref="pdf")
    entities = relationship("Entity", secondary="pdf_entity", backref="pdfs")

    @property
    def path(self) -> Path:
        """Compute file path from current_file_name."""
        from easy_access.settings import SETTINGS, DirSetting

        return SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full / self.current_file_name


class v1_CopyrightItem(Base):
    __tablename__ = "v1_copyright_item"

    material_id: Mapped[int] = mapped_column(Integer, primary_key=True)
    matching_copyright_item_id: Mapped[int | None] = mapped_column(
        Integer, ForeignKey("copyright_data.material_id"), nullable=True
    )
    workflow_status: Mapped[WorkflowStatus] = mapped_column(
        SAEnum(WorkflowStatus), nullable=False, default=WorkflowStatus.ToDo
    )
    retrieved_from_copyright_on: Mapped[datetime | None] = mapped_column(
        DateTime, nullable=True
    )
    url: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    manual_classification: Mapped[str | None] = mapped_column(
        String(2048), nullable=True, index=True
    )
    remarks: Mapped[str | None] = mapped_column(String(10000), nullable=True)
    scope: Mapped[str | None] = mapped_column(String(255), nullable=True)
    faculty: Mapped[str | None] = mapped_column(String(255), nullable=True)
    ml_prediction: Mapped[Classification | None] = mapped_column(
        SAEnum(Classification, values_callable=lambda x: [e.value for e in x]),
        nullable=True,
    )
    filename: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    title: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    filehash: Mapped[str | None] = mapped_column(String(255), nullable=True)
    owner: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    period: Mapped[Period | None] = mapped_column(
        SAEnum(Period, values_callable=lambda x: [e.value for e in x]), nullable=True
    )  # type: ignore
    department: Mapped[str | None] = mapped_column(
        String(2048), nullable=True, index=True
    )
    course_code: Mapped[str | None] = mapped_column(
        String(255), nullable=True, index=True
    )
    course_name: Mapped[str | None] = mapped_column(
        String(2048), nullable=True, index=True
    )
    filetype: Mapped[Filetype] = mapped_column(
        SAEnum(Filetype, values_callable=lambda x: [e.value for e in x]),
        nullable=False,
        default=Filetype.UNKNOWN,
    )
    classification: Mapped[Classification] = mapped_column(
        SAEnum(Classification, values_callable=lambda x: [e.value for e in x]),
        nullable=False,
        default=Classification.LANGE_OVERNAME,
    )
    manual_identifier: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    auditor: Mapped[str | None] = mapped_column(String(10000), nullable=True)
    last_change: Mapped[date | None] = mapped_column(Date, nullable=True)
    status: Mapped[Status] = mapped_column(
        SAEnum(Status, values_callable=lambda x: [e.value for e in x]),
        nullable=False,
        default=Status.PUBLISHED,
        index=True,
    )
    isbn: Mapped[str | None] = mapped_column(String(255), nullable=True)
    doi: Mapped[str | None] = mapped_column(String(255), nullable=True)
    in_collection: Mapped[bool | None] = mapped_column(Boolean, nullable=True)
    pagecount: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    wordcount: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    picturecount: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    author: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    publisher: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    reliability: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    pages_x_students: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    count_students_registered: Mapped[int] = mapped_column(
        Integer, nullable=False, default=0
    )
    cursuscodes: Mapped[str | None] = mapped_column(String(2048), nullable=True)


class CopyrightItemCourse(Base):
    __tablename__ = "copyright_item_course"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    copyright_item_id: Mapped[int] = mapped_column(
        Integer, ForeignKey("copyright_data.material_id")
    )
    course_cursuscode: Mapped[int] = mapped_column(
        Integer, ForeignKey("course_data.cursuscode")
    )


class CopyrightItemUpdate(Base):
    __tablename__ = "copyright_item_update"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    copyright_item_id: Mapped[int] = mapped_column(
        Integer, ForeignKey("copyright_data.material_id")
    )
    item_update_id: Mapped[int] = mapped_column(Integer, ForeignKey("item_updates.id"))


class CopyrightItem(Base):
    __tablename__ = "copyright_data"

    material_id: Mapped[int] = mapped_column(Integer, primary_key=True)
    period: Mapped[Period] = mapped_column(
        SAEnum(Period, values_callable=lambda x: [e.value for e in x]), nullable=False
    )  # type: ignore
    department: Mapped[str] = mapped_column(String(2048), nullable=False, index=True)
    course_code: Mapped[str] = mapped_column(String(255), nullable=False, index=True)
    course_name: Mapped[str] = mapped_column(String(2048), nullable=False, index=True)
    url: Mapped[str | None] = mapped_column(String(255), nullable=True, unique=True)
    filename: Mapped[str | None] = mapped_column(
        String(2048), nullable=True, index=True
    )
    title: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    owner: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    filetype: Mapped[Filetype] = mapped_column(
        SAEnum(Filetype, values_callable=lambda x: [e.value for e in x]),
        nullable=False,
        default=Filetype.UNKNOWN,
    )
    classification: Mapped[Classification] = mapped_column(
        SAEnum(Classification, values_callable=lambda x: [e.value for e in x]),
        nullable=False,
        default=Classification.LANGE_OVERNAME,
    )
    ml_prediction: Mapped[Classification | None] = mapped_column(
        SAEnum(Classification, values_callable=lambda x: [e.value for e in x]),
        nullable=True,
        index=True,
    )
    manual_classification: Mapped[str | None] = mapped_column(
        String(2048), nullable=True, index=True
    )
    manual_identifier: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    v2_manual_classification: Mapped[ClassificationV2 | None] = mapped_column(
        SAEnum(ClassificationV2, values_callable=lambda x: [e.value for e in x]),
        nullable=True,
        index=True,
        default=ClassificationV2.ONBEKEND,
    )
    v2_overnamestatus: Mapped[OvernameStatus | None] = mapped_column(
        SAEnum(OvernameStatus, values_callable=lambda x: [e.value for e in x]),
        nullable=True,
        index=True,
        default=OvernameStatus.ONBEKEND,
    )
    v2_lengte: Mapped[Lengte | None] = mapped_column(
        SAEnum(Lengte, values_callable=lambda x: [e.value for e in x]),
        nullable=True,
        index=True,
        default=Lengte.ONBEKEND,
    )
    scope: Mapped[str | None] = mapped_column(String(255), nullable=True)
    remarks: Mapped[str | None] = mapped_column(
        String(10000), nullable=True, index=True
    )
    auditor: Mapped[str | None] = mapped_column(String(10000), nullable=True)
    last_change: Mapped[date | None] = mapped_column(Date, nullable=True)
    status: Mapped[Status] = mapped_column(
        SAEnum(Status, values_callable=lambda x: [e.value for e in x]),
        nullable=False,
        default=Status.PUBLISHED,
        index=True,
    )
    isbn: Mapped[str | None] = mapped_column(String(255), nullable=True)
    doi: Mapped[str | None] = mapped_column(String(255), nullable=True)
    in_collection: Mapped[bool | None] = mapped_column(Boolean, nullable=True)
    pagecount: Mapped[int] = mapped_column(Integer, nullable=False)
    wordcount: Mapped[int] = mapped_column(Integer, nullable=False)
    picturecount: Mapped[int] = mapped_column(Integer, nullable=False)
    author: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    publisher: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    reliability: Mapped[int] = mapped_column(Integer, nullable=False)
    pages_x_students: Mapped[int] = mapped_column(Integer, nullable=False)
    count_students_registered: Mapped[int] = mapped_column(Integer, nullable=False)
    filehash: Mapped[str | None] = mapped_column(String(255), nullable=True)
    last_scan_date_university: Mapped[date | None] = mapped_column(Date, nullable=True)
    last_scan_date_course: Mapped[date | None] = mapped_column(Date, nullable=True)

    retrieved_from_copyright_on: Mapped[datetime | None] = mapped_column(
        DateTime, nullable=True, index=True
    )
    workflow_status: Mapped[WorkflowStatus] = mapped_column(
        SAEnum(WorkflowStatus), nullable=False, default=WorkflowStatus.ToDo, index=True
    )
    possible_fine: Mapped[float | None] = mapped_column(Float, nullable=True)
    infringement: Mapped[Infringement] = mapped_column(
        SAEnum(Infringement, values_callable=lambda x: [e.value for e in x]),
        nullable=False,
        default=Infringement.UNDETERMINED,
    )
    file_exists: Mapped[bool | None] = mapped_column(Boolean, nullable=True)
    last_canvas_check: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    canvas_course_id: Mapped[int | None] = mapped_column(
        Integer, nullable=True, index=True
    )

    faculty_id: Mapped[str] = mapped_column(
        String(255), ForeignKey("organization_data.abbreviation"), nullable=False
    )
    is_duplicate: Mapped[bool | None] = mapped_column(
        Boolean, nullable=True, index=True
    )

    # relationships
    courses = relationship("Course", secondary="copyright_item_course", backref="items")
    faculty = relationship("Organization", foreign_keys=[faculty_id], backref="items")
    changes = relationship(
        "ItemUpdate", secondary="copyright_item_update", backref="items"
    )
    pdf = relationship("PDF", backref="copyright_item", uselist=False)

    def actual_status(self) -> Status:
        if not self.url or self.url.strip() == "":
            return Status.DELETED
        elif self.file_exists in [True, 1, "1"]:
            return self.status
        else:
            return Status.DELETED

    def misaligned_status(self) -> bool:
        return self.status != self.actual_status()

    def status_details(self) -> dict[str, str | bool | datetime | int | Status | None]:
        return {
            "material_id": self.material_id,
            "filename": self.filename,
            "url": self.url,
            "status": self.status,
            "actual_status": self.actual_status(),
            "file_exists": self.file_exists,
            "last_canvas_check": self.last_canvas_check,
        }

    @property
    def course_link(self) -> str:
        if not self.canvas_course_id or not self.filename:
            return ""
        base_url = SETTINGS.university_settings.lms.url

        return f"{base_url}/courses/{self.canvas_course_id}/files?search_term={self.filename.replace(' ', '%20')}"


class StagedCopyrightItem(Base):
    """
    Staging table for raw data ingested from copyright export files.
    Fields are kept as simple as possible to accommodate raw data.
    """

    __tablename__ = "staged_copyright_item"

    material_id: Mapped[int] = mapped_column(Integer, primary_key=True)
    period: Mapped[str | None] = mapped_column(String(255), nullable=True)
    department: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    course_code: Mapped[str | None] = mapped_column(String(255), nullable=True)
    course_name: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    url: Mapped[str | None] = mapped_column(String(255), nullable=True)
    filename: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    title: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    owner: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    filetype: Mapped[str | None] = mapped_column(String(255), nullable=True)
    classification: Mapped[str | None] = mapped_column(String(255), nullable=True)
    manual_classification: Mapped[str | None] = mapped_column(
        String(2048), nullable=True
    )
    manual_identifier: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    scope: Mapped[str | None] = mapped_column(String(255), nullable=True)
    remarks: Mapped[str | None] = mapped_column(String(10000), nullable=True)
    ml_prediction: Mapped[str | None] = mapped_column(String(255), nullable=True)
    isbn: Mapped[str | None] = mapped_column(String(255), nullable=True)
    doi: Mapped[str | None] = mapped_column(String(255), nullable=True)
    in_collection: Mapped[str | None] = mapped_column(String(255), nullable=True)
    pagecount: Mapped[str | None] = mapped_column(String(255), nullable=True)
    wordcount: Mapped[str | None] = mapped_column(String(255), nullable=True)
    picturecount: Mapped[str | None] = mapped_column(String(255), nullable=True)
    author: Mapped[str | None] = mapped_column(String(255), nullable=True)
    publisher: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    auditor: Mapped[str | None] = mapped_column(String(10000), nullable=True)
    last_change: Mapped[date | None] = mapped_column(Date, nullable=True)
    status: Mapped[str | None] = mapped_column(String(255), nullable=True)
    reliability: Mapped[str | None] = mapped_column(String(255), nullable=True)
    pages_x_students: Mapped[str | None] = mapped_column(String(255), nullable=True)
    count_students_registered: Mapped[str | None] = mapped_column(
        String(255), nullable=True
    )
    retrieved_from_copyright_on: Mapped[datetime | None] = mapped_column(
        DateTime, nullable=True
    )
    workflow_status: Mapped[str | None] = mapped_column(String(255), nullable=True)
    faculty: Mapped[str | None] = mapped_column(String(255), nullable=True)
    file_exists: Mapped[str | None] = mapped_column(String(255), nullable=True)
    # Additional fields found in raw data
    id_course: Mapped[str | None] = mapped_column(String(255), nullable=True)
    id_material: Mapped[str | None] = mapped_column(String(255), nullable=True)
    last_scan_date_university: Mapped[str | None] = mapped_column(
        String(255), nullable=True
    )
    filehash: Mapped[str | None] = mapped_column(String(255), nullable=True)
    manual_classification_report: Mapped[str | None] = mapped_column(
        String(2048), nullable=True
    )
    count_downloads_material: Mapped[str | None] = mapped_column(
        String(255), nullable=True
    )
    last_scan_date_course: Mapped[str | None] = mapped_column(
        String(255), nullable=True
    )


class StagedFacultyUpdate(Base):
    """
    Staging table for updates from faculty sheets.
    """

    __tablename__ = "staged_faculty_update"

    material_id: Mapped[int] = mapped_column(Integer, primary_key=True)
    manual_classification: Mapped[str | None] = mapped_column(
        String(2048), nullable=True
    )
    remarks: Mapped[str | None] = mapped_column(String(10000), nullable=True)
    workflow_status: Mapped[str | None] = mapped_column(String(255), nullable=True)


class StagedProcessingFailure(Base):
    """
    Stores failures encountered while processing staged rows.
    Each row references the staged material_id (if available), the raw payload
    (as JSON), and an error message to aid debugging/retry.
    """

    __tablename__ = "staged_processing_failures"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    material_id: Mapped[int | None] = mapped_column(Integer, nullable=True, index=True)
    staged_payload: Mapped[dict | None] = mapped_column(JSON, nullable=True)
    error_message: Mapped[str | None] = mapped_column(String(2000), nullable=True)
