"""
functions to manage the db, add items, cleanup, etc.
"""

from tortoise import Tortoise
from easy_access.orm.models import Organization, Programme, Person, Course
from easy_access.settings import SETTINGS, DirSetting, SettingsFaculty, SettingsProgramme

async def init()    :

    await Tortoise.init(
        db_url='sqlite://db.sqlite3',
        modules={'models': ['easy_access.orm.models']}
    )

    await Tortoise.generate_schemas(safe=True)



async def add_orgs() -> None:
    """
    From SETTINGS, grab data for the university organization: faculties, departments, programmes, etc.
    Then load them into the db.
    Adds Organizations and Programmes.
    """
    ...
    faculties: list[SettingsFaculty] = SETTINGS.university_settings.faculties
    department_mapping: dict[str, str] = SETTINGS.university_settings.department_mapping
    programmes: set[SettingsProgramme] = SETTINGS.university_settings.programmes
    course_mapping: dict[str, dict[str, str]] = SETTINGS.university_settings.course_mapping

    """
    class Programme:
        name: str | None = None
        abbreviation: str | None = None
        programme_type: Literal['b', 'm', 'o'] = 'o'
        cluster: str | None = None
        faculty_name: str | None = None
        faculty_abbreviation: str | None = None


    class Faculty:
        name: str
        abbreviation: str = ""
        programmes: list[Programme] = field(default_factory=list)
    """

async def load_osiris_data() -> None:
    """
    Load data from the osiris_data.json file into the db.
    Adds Courses.
    """
    ...

async def load_person_data() -> None:
    """
    Load data from the person_data.json file into the db.
    Adds Persons.
    """
    ...

async def load_llm_data() -> None:
    """
    Load data from the script_data / classifications / *.json files into the db.
    Adds LLMClassifications.
    Note: this can only be done after the CopyrightItems have been added.
    """
    ...
