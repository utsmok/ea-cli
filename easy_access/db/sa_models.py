from __future__ import annotations

from sqlalchemy import (
    Float,
    ForeignKey,
    Integer,
    String,
    UniqueConstraint,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from .models_base import Base


class Organization(Base):
    __tablename__ = "organization_data"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    parent_organization_id: Mapped[int | None] = mapped_column(
        Integer, ForeignKey("organization_data.id"), nullable=True
    )
    hierarchy_level: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    name: Mapped[str] = mapped_column(String(2048), index=True)
    abbreviation: Mapped[str] = mapped_column(String(255), unique=True, index=True)
    full_abbreviation: Mapped[str] = mapped_column(String(2048), unique=True)

    __table_args__ = (
        UniqueConstraint("name", "abbreviation", name="uq_organization_name_abbr"),
    )

    # relationships
    children = relationship("Organization", backref="parent", remote_side=[id])


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
    faculty_abbreviation: Mapped[str | None] = mapped_column(
        String(255), ForeignKey("organization_data.abbreviation"), nullable=True
    )
    ec: Mapped[float | None] = mapped_column(Float, nullable=True)
    programme: Mapped[str | None] = mapped_column(String(2048), nullable=True)
    notes: Mapped[str | None] = mapped_column(String(10000), nullable=True)
    category: Mapped[str | None] = mapped_column(String(2048), nullable=True)

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
    faculty_abbreviation: Mapped[str | None] = mapped_column(
        String(255), ForeignKey("organization_data.abbreviation"), nullable=True
    )
    people_page_url: Mapped[str | None] = mapped_column(String(2048), nullable=True)

    courses = relationship(
        "Course", secondary="course_employee", back_populates="teachers"
    )

    def __str__(self) -> str:  # pragma: no cover - helper
        return self.main_name or self.input_name
