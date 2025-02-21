"""
functions to manage the db, add items, cleanup, etc.
"""
from collections import Counter
from tortoise import Tortoise
from easy_access.orm.models import (
    Organization, Programme, Person, Course, Faculty, LLMClassification, MissingCourse, CopyrightItem
    )
from easy_access.settings import SETTINGS, DirSetting, SettingsFaculty, SettingsProgramme, FileSetting
import json
from easy_access.utils import warn, info, cool
from easy_access.sheets.sheet import read_copyright_export
import polars as pl
from datetime import datetime

async def init()    :

    await Tortoise.init(
        db_url='sqlite://db.sqlite3',
        modules={'models': ['easy_access.orm.models']}
    )

async def create() -> None:

    await Tortoise.generate_schemas(safe=True)


async def load_base_data() -> None:
    """
    Load the base data into the db: orgs, courses, persons.
    """
    async def load_org_data() -> None:
        """
        From SETTINGS.university_settings.faculties, create Faculty and Programme objects.
        """
        faculties: list[SettingsFaculty] = SETTINGS.university_settings.faculties
        # first retrieve or create the university org
        university, _ = await Organization.get_or_create(defaults={"name":"University of Twente", "abbreviation":"UT", "full_abbreviation":"UT", "parent_organization":None, "hierarchy_level":0}, abbreviation="UT")

        for faculty in faculties:
            faculty_obj, _ = await Faculty.get_or_create(defaults={
                    "name": faculty.name,
                    "abbreviation": faculty.abbreviation,
                    "full_abbreviation": faculty.abbreviation,
                    "parent_organization": university,
                    "hierarchy_level": 1
                }, abbreviation=faculty.abbreviation
            )
            progamme_list = []
            existing_programme_names = await Programme().all().values("name")
            existing_programme_names = {p['name'] for p in existing_programme_names}
            for programme in faculty.programmes:
                if programme.name in existing_programme_names:
                    continue
                if programme.abbreviation:
                    abbr = programme.abbreviation
                else:
                    abbr = ""
                    if "master" in programme.name.lower():
                        abbr = "M-"
                    elif "bachelor" in programme.name.lower():
                        abbr = "B-"
                    else:
                        abbr = "O-"

                    removed_prefix = programme.name.lower().replace('bachelor', '').replace('master', '').strip()
                    if " " in removed_prefix:
                        parts = removed_prefix.split(' ')
                        if len(parts) == 2:
                            abbr += parts[0][0] + parts[1][0]
                        elif len(parts) >= 3:
                            abbr += parts[0][0] + parts[1][0] + parts[2][0]
                    else:
                        abbr += removed_prefix[:3]


                programme_dict = {
                    "name": programme.name,
                    "abbreviation": abbr,
                    "programme_type": programme.programme_type if programme.programme_type else None,
                    "faculty": faculty_obj,
                    "cluster": programme.cluster if programme.cluster else None
                }
                progamme_list.append(programme_dict)



            await Programme.bulk_create(objects=[Programme(**p) for p in progamme_list])

    async def load_osiris_data() -> None:
        """
        Creates Courses from the osiris_data.json file.
        Staff data is added later once people data has been loaded.
        """
        if not SETTINGS.files[FileSetting.OSIRIS_DATA].exists:
            warn("No osiris_data.json file found; data not loaded to DB.")
            return

        try:
            with open(SETTINGS.files[FileSetting.OSIRIS_DATA].path, 'r', encoding='utf-8') as f:
                osiris_data: dict[str, dict[str,str|list[str]]] = json.load(f)
        except Exception as e:
            warn(f"Error reading osiris_data.json: {e}")
            return


        course_dicts = []
        existing_course_codes = await Course().all().values("cursuscode")
        existing_course_codes = {int(c['cursuscode']) for c in existing_course_codes}
        for course_data in osiris_data.values():
            if int(course_data.get('cursuscode',0)) in existing_course_codes or not course_data.get('cursuscode'):
                continue
            course_dict: dict[str, str | list[str] | None] = {
                "cursuscode": int(course_data.get('cursuscode')),
                "internal_id": int(course_data.get('internal_id')),
                "name": course_data.get('name', None),
                "short_name": course_data.get('short_name', None),
                "ec": int(round(float(course_data.get('ec',0).replace(',','.')),0)),
                "programme": course_data.get('programme', None),
                "notes": course_data.get('notes', None),
                "category": course_data.get('category', None),
            }

            existing_course_codes.add(int(course_data.get('cursuscode')))

            year: str = course_data.get('year')
            if '-' in year:
                year = int(year.split('-')[0])
                course_dict['year'] = year

            faculty = await Faculty.get_or_none(abbreviation=course_data.get('faculty'))
            if faculty:
                course_dict['faculty'] = faculty

            course_dicts.append(course_dict)

        await Course.bulk_create(objects=[Course(**c) for c in course_dicts])

    async def load_person_data() -> None:
        """
        Load data from the person_data.json file into the db.
        Adds Persons and Orgs, creates MissingCourses where necessary.
        """


        if not SETTINGS.files[FileSetting.PERSON_DATA].exists:
            warn("No person_data.json file found; data not loaded to DB.")
            return
        try:
            with open(SETTINGS.files[FileSetting.PERSON_DATA].path, 'r', encoding='utf-8') as f:
                person_data: list[dict[str, str | float | list[str]]] = json.load(f)
        except Exception as e:
            warn(f"Error reading person_data.json: {e}")
            return

        person_dicts = []
        existing_person_names = await Person().all().values("input_name")
        existing_person_names = {p['input_name'] for p in existing_person_names}
        for person in person_data:
            if person.get('input_name') in existing_person_names:
                continue
            person_dict = {
                "input_name": person.get('input_name'),
                "main_name": person.get('main_name', None),
                "match_confidence": person.get('match_confidence', None),
                "first_name": person.get('other_names', [None])[0],
                "email": person.get('email', None),
                "faculty": await Faculty.get_or_none(abbreviation=person.get('faculty', None)),
                "people_page_url": person.get('people_page_url', None)
            }

            orgs: list[dict[str, str]] = person.get('orgs', [])
            orgs.sort(key=lambda x: x.get('abbr').count('-'))
            orglist = []
            for org in orgs:
                try:
                    hierarchy_level = org.get('abbr').count('-')+1
                    org_obj, _ = await Organization.get_or_create(defaults={
                        "name":org.get('name'),
                        "abbreviation":org.get('abbr',"-").split('-')[-1],
                        "full_abbreviation":org.get('abbr'),
                        "hierarchy_level": hierarchy_level,
                        }, full_abbreviation=org.get('abbr'))
                    if hierarchy_level == 1:
                        org_obj.parent_organization, _ = await Organization.get_or_create(defaults={"name":"University of Twente", "abbreviation":"UT", "full_abbreviation":"UT", "parent_organization":None, "hierarchy_level":0}, abbreviation="UT")
                    elif hierarchy_level >= 2:
                        parent_org_abbr = org.get('abbr').split('-')[-2]
                        parent_org_full_abbreviation = org.get('abbr').rsplit('-',1)[0]
                        org_obj.parent_organization = await Organization.get(abbreviation=parent_org_abbr, full_abbreviation=parent_org_full_abbreviation, hierarchy_level=hierarchy_level-1)
                    await org_obj.save()
                    orglist.append(org_obj)
                except Exception as e:
                    warn(f"Error adding org to person: {e}")
                    continue
            person_dict['orgs'] = orglist

            courses: list[dict[str, str]] = person.get('courses', [])
            course_list = []
            for course in courses:
                try:
                    course_obj = await Course.get_or_none(cursuscode=course.get('course_code'))
                    if course_obj:
                        course_list.append(course_obj)
                    else:
                        await MissingCourse.get_or_create(defaults={"cursuscode":course.get('course_code')}, cursuscode=course.get('course_code'))
                except Exception as e:
                    warn(f"Error adding course to person: {e}")
                    continue
            if course_list:
                person_dict['courses'] = course_list

            programmes: list[dict[str, str]] = person.get('programmes', [])
            programme_list = []
            for programme in programmes:
                name = programme.get('name').replace('Master', "").replace('Bachelor', "").strip()
                if not name:
                    continue
                programme_obj = await Programme.get_or_none(name=name, abbreviation=programme.get('abbr'))
                if programme_obj:
                    programme_list.append(programme_obj)

            if programme_list:
                person_dict['programmes'] = programme_list

            person_dicts.append(person_dict)

        await Person.bulk_create(objects=[Person(**p) for p in person_dicts])

    async def link_persons_to_courses() -> Counter:
        """
        Link the Persons to the Courses they are involved in.
        """

        # load osiris data
        if not SETTINGS.files[FileSetting.OSIRIS_DATA].exists:
            warn("No osiris_data.json file found; data not loaded to DB.")
            return

        try:
            with open(SETTINGS.files[FileSetting.OSIRIS_DATA].path, 'r', encoding='utf-8') as f:
                osiris_data: dict[str, dict[str,str|list[str]]] = json.load(f)
        except Exception as e:
            warn(f"Error reading osiris_data.json: {e}")
            return

        counter = Counter()

        for course_data in osiris_data.values():
            course_obj = await Course.get_or_none(cursuscode=course_data.get('cursuscode'))
            if not course_obj:
                continue


            teachers = course_data.get('teachers', [])
            for teacher_name in teachers:
                teacher_obj = await Person.get_or_none(input_name=teacher_name)
                if teacher_obj:
                    await course_obj.teachers.add(teacher_obj)
                    counter['teachers'] += 1

            contacts = course_data.get('contacts', [])
            for contact_name in contacts:
                contact_obj = await Person.get_or_none(input_name=contact_name)
                if contact_obj:
                    await course_obj.contacts.add(contact_obj)
                    counter['contacts'] += 1

            tutors = course_data.get('tutors', [])
            for tutor_name in tutors:
                tutor_obj = await Person.get_or_none(input_name=tutor_name)
                if tutor_obj:
                    await course_obj.tutors.add(tutor_obj)
                    counter['tutors'] += 1

            docenten = course_data.get('contacts', [])
            for docent_name in docenten:
                docent_obj = await Person.get_or_none(input_name=docent_name)
                if docent_obj:
                    await course_obj.docenten.add(docent_obj)
                    counter['docenten'] += 1

            examinators = course_data.get('examinators', [])
            for examinators_name in examinators:
                examinator_obj = await Person.get_or_none(input_name=examinators_name)
                if examinator_obj:
                    await course_obj.examinators.add(examinator_obj)
                    counter['examinators'] += 1

            unknown_roles = course_data.get('unknown_role', [])
            for unknown_role_name in unknown_roles:
                unknown_role_obj = await Person.get_or_none(input_name=unknown_role_name)
                if unknown_role_obj:
                    await course_obj.unknown_roles.add(unknown_role_obj)
                    counter['unknown_roles'] += 1

        return counter

    try:
        faculty_count = await Faculty.all().count()
        programme_count = await Programme.all().count()
        info(f"# of Faculties present in DB before load_org_data: {faculty_count}")
        info(f"# of Programmes present in DB before load_org_data: {programme_count}")

        await load_org_data()


        faculty_count = await Faculty.all().count()
        programme_count = await Programme.all().count()
        info(f"# of Faculties present in DB after load_org_data: {faculty_count}")
        info(f"# of Programmes present in DB after load_org_data: {programme_count}")

        course_count = await Course.all().count()
        info(f"# of Courses present in DB before load_osiris_data: {course_count}")

        await load_osiris_data()

        course_count = await Course.all().count()
        info(f"# of Courses present in DB after load_osiris_data: {course_count}")

        person_count = await Person.all().count()
        org_count = await Organization.all().count()
        missing_orgs = await MissingCourse.all().count()
        info(f"# of Persons present in DB before load_person_data: {person_count}")
        info(f"# of Organizations present in DB before load_person_data: {org_count}")
        info(f"# of MissingCourses present in DB before load_person_data: {missing_orgs}")

        await load_person_data()

        org_count = await Organization.all().count()
        person_count = await Person.all().count()
        missing_orgs = await MissingCourse.all().count()
        info(f"# of Persons present in DB after load_person_data: {person_count}")
        info(f"# of Organizations present in DB after load_person_data: {org_count}")
        info(f"# of MissingCourses present in DB after load_person_data: {missing_orgs}")

        results = await link_persons_to_courses()
        info(f"Overview of new relations added to Courses: {results}")
    except Exception as e:
        warn(f"Error loading base data: {e}")

    await Tortoise.close_connections()

def standardize_dataframe(df: pl.DataFrame) -> pl.DataFrame:
    """
    rename cols to standard format
    cast all cols to str
    replace '-' with None
    drop useless cols
    """
    df = df.with_columns(
            pl.exclude(pl.String).cast(str)
        ).rename(
            lambda col: col.replace(" ", "_")
            .replace("#", "count_")
            .replace("*", "x")
            .lower()
        ).with_columns(
            pl.when(pl.col(pl.String) != "-")
            .then(pl.col(pl.String))
            .name.keep()
        )
    if 'type' in df.columns:
        df = df.drop('type')
    if 'google_search_file' in df.columns:
        df = df.drop('google_search_file')
    return df
async def parse_raw_copyright_item(item: dict[str, str]) -> CopyrightItem:
    try:
        if item.get('faculty') == 'Unmapped':
            abbr = 'UNM'
        else:
            abbr = item.get('faculty')
        faculty = await Faculty.get(abbreviation=abbr)
    except Exception as e:
        print(f"Error getting faculty with {item.get('faculty')}: {e}")
    item['faculty'] = faculty
    item['material_id'] = int(item.get('material_id'))
    item['last_change'] = datetime.strptime(item.get('last_change'), "%Y-%m-%d")
    item['retrieved_from_copyright_on'] = datetime.strptime(item.get('retrieved_from_copyright_on'), "%Y-%m-%d")
    item['pagecount'] = int(item.get('pagecount'))
    item['wordcount'] = int(item.get('wordcount'))
    item['picturecount'] = int(item.get('picturecount'))
    item['reliability'] = int(item.get('reliability'))
    item['pages_x_students'] = int(item.get('pages_x_students'))
    item['count_students_registered'] = int(item.get('count_students_registered'))
    item['filetype'] = item.get('filetype', 'unknown') if item.get('filetype') else 'unknown'
    try:
        final_item = CopyrightItem(**item)
        return final_item

    except Exception as e:
        print(f'Error while trying to create CopyrightItem with mat_id {item['material_id']}:{e}')
        return None
async def load_raw_items() -> None:
    """
    Loads in new items from copyright export raw data
    """
    error = None
    info(f'# of items in db before loading raw items: {await CopyrightItem.all().count()}')
    try:
        file_date, df = read_copyright_export()
        items = standardize_dataframe(df).to_dicts()

        item_list = []

        existing_mat_ids = await CopyrightItem.all().values("material_id")
        existing_mat_ids = {int(m['material_id']) for m in existing_mat_ids}

        info(f'Read in {len(items)} raw copyright items. {len(existing_mat_ids)} items already in db.')
        for item in items:
            if int(item.get('material_id')) in existing_mat_ids:
                continue
            created_item: CopyrightItem = await parse_raw_copyright_item(item)
            if not created_item:
                continue
            item_list.append(created_item)

        await CopyrightItem.bulk_create(objects=item_list)
    except Exception as e:
        warn(f"Error loading raw items: {e}")
        error = e
    finally:
        info(f'# of items in db after loading raw items: {await CopyrightItem.all().count()}')
        if error:
            raise error
    await Tortoise.close_connections()


async def update_copyright_items(data: pl.DataFrame) -> None:
    """
    for a given dataframe with copyright items, compare to db and update where needed
    TODO: Determine when to update and when to keep existing data
    """

    # standardize df
    # loop over items
    # if item is not in db: add it
    # else compare values in specific fields to determine if we need to update

