"""
functions to manage the db, add items, cleanup, etc.
"""
from sqlalchemy import create_engine
from enum import Enum
from collections import Counter
from tortoise import Tortoise
from easy_access.orm.models import (
    Organization, Programme, Person, Course, Faculty, LLMClassification, MissingCourse, CopyrightItem, CourseEmployee,
    Infringement, WorkflowStatus, Filetype, Status, Classification, Period
    )
from easy_access.settings import SETTINGS, DirSetting, SettingsFaculty, SettingsProgramme, FileSetting
import json
from easy_access.utils import warn, info, cool, File
from easy_access.sheets.sheet import read_copyright_export
import polars as pl
from datetime import datetime, timezone
from loguru import logger
from easy_access.sheets.enrichment import determine_course_code

async def init() -> None:
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

        existing_person_names = await Person().all().values("input_name")
        existing_person_names = {p['input_name'] for p in existing_person_names}
        for person in person_data:
            if person.get('input_name') in existing_person_names:
                continue
            person_dict = {
                "input_name": person.get('input_name').strip(),
                "main_name": person.get('main_name', None),
                "match_confidence": person.get('match_confidence', None),
                "first_name": person.get('other_names', [None])[0],
                "email": person.get('email', None),
                "faculty": await Faculty.get_or_none(abbreviation=person.get('faculty', None)),
                "people_page_url": person.get('people_page_url', None)
            }
            for k, v in person_dict.items():
                if isinstance(v, str):
                    person_dict[k] = v.strip()
            orgs: list[dict[str, str]] = person.get('orgs', [])
            orgs.sort(key=lambda x: x.get('abbr').count('-'))
            orglist = []
            for org in orgs:
                try:
                    hierarchy_level = org.get('abbr').count('-')+1
                    orgname = org.get('name').strip()
                    org_sole_abbr = org.get('abbr').split('-')[-1].strip()
                    org_full_abbr = org.get('abbr').strip()
                    org_obj, _ = await Organization.get_or_create(defaults={
                        "name":orgname,
                        "abbreviation":org_sole_abbr,
                        "full_abbreviation": org_full_abbr,
                        "hierarchy_level": hierarchy_level,
                        }, full_abbreviation=org_full_abbr)
                    if hierarchy_level == 1:
                        org_obj.parent_organization, _ = await Organization.get_or_create(defaults={"name":"University of Twente", "abbreviation":"UT", "full_abbreviation":"UT", "parent_organization":None, "hierarchy_level":0}, abbreviation="UT")
                    elif hierarchy_level >= 2:
                        parent_org_abbr = org_full_abbr.split('-')[-2].strip()
                        parent_org_full_abbreviation = org_full_abbr.rsplit('-',1)[0].strip()
                        org_obj.parent_organization = await Organization.get(abbreviation=parent_org_abbr, full_abbreviation=parent_org_full_abbreviation, hierarchy_level=hierarchy_level-1)
                    await org_obj.save()
                    orglist.append(org_obj)
                except Exception as e:
                    warn(f"Error adding org {org} to person: {e}")
                    continue

            person = await Person.create(**person_dict)
            if orglist:
                await person.orgs.add(*orglist)

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
                    await CourseEmployee.get_or_create(course=course_obj, person=teacher_obj, role="teacher")
                    counter['teachers'] += 1

            contacts = course_data.get('contacts', [])
            for contact_name in contacts:
                contact_obj = await Person.get_or_none(input_name=contact_name)
                if contact_obj:
                    await CourseEmployee.get_or_create(course=course_obj, person=contact_obj, role="contact")
                    counter['contacts'] += 1

            tutors = course_data.get('tutors', [])
            for tutor_name in tutors:
                tutor_obj = await Person.get_or_none(input_name=tutor_name)
                if tutor_obj:
                    await CourseEmployee.get_or_create(course=course_obj, person=tutor_obj, role="tutor")
                    counter['tutors'] += 1

            docenten = course_data.get('contacts', [])
            for docent_name in docenten:
                docent_obj = await Person.get_or_none(input_name=docent_name)
                if docent_obj:
                    await CourseEmployee.get_or_create(course=course_obj, person=docent_obj, role="docent")
                    counter['docenten'] += 1

            examinators = course_data.get('examinators', [])
            for examinators_name in examinators:
                examinator_obj = await Person.get_or_none(input_name=examinators_name)
                if examinator_obj:
                    await CourseEmployee.get_or_create(course=course_obj, person=examinator_obj, role="examinator")
                    counter['examinators'] += 1

            unknown_roles = course_data.get('unknown_role', [])
            for unknown_role_name in unknown_roles:
                unknown_role_obj = await Person.get_or_none(input_name=unknown_role_name)
                if unknown_role_obj:
                    await CourseEmployee.get_or_create(course=course_obj, person=unknown_role_obj, role="unknown_role")
                    counter['unknown_roles'] += 1

        return counter

    await init()
    await create()
    try:
        faculty_count = await Faculty.all().count()
        programme_count = await Programme.all().count()
        info(f"# of Faculties present in DB before load_org_data: {faculty_count}")
        info(f"# of Programmes present in DB before load_org_data: {programme_count}")

        await load_org_data()


        faculty_count = await Faculty.all().count()
        programme_count = await Programme.all().count()
        cool(f"# of Faculties present in DB after load_org_data: {faculty_count}")
        cool(f"# of Programmes present in DB after load_org_data: {programme_count}")

        course_count = await Course.all().count()
        info(f"# of Courses present in DB before load_osiris_data: {course_count}")

        await load_osiris_data()

        course_count = await Course.all().count()
        cool(f"# of Courses present in DB after load_osiris_data: {course_count}")

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
        cool(f"# of Persons present in DB after load_person_data: {person_count}")
        cool(f"# of Organizations present in DB after load_person_data: {org_count}")
        cool(f"# of MissingCourses present in DB after load_person_data: {missing_orgs}")

        results = await link_persons_to_courses()
        cool(f"Overview of new relations added to Courses: {results}")
    except Exception as e:
        warn(f"Error loading base data: {e}")

    await Tortoise.close_connections()
def standardize_dataframe(df: pl.DataFrame) -> pl.DataFrame:
    """
    rename cols to standard format
    cast all cols to str
    replace '-' with None
    filter missing material_ids
    filter to select only relevant itemtypes
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
        ).filter(
            (pl.col("material_id").is_not_null())
        ).filter(
            (pl.col("filetype").is_in(["pdf", "ppt", "doc", "-"])) |
            (pl.col("filetype").is_null())
        )

    if 'type' in df.columns:
        df = df.drop('type')
    if 'google_search_file' in df.columns:
        df = df.drop('google_search_file')
    return df
async def dict_to_copyright_item(item: dict[str, str]) -> CopyrightItem:
    copyright_item_keys = {
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
        "retrieved_from_copyright_on",
        "workflow_status",
        "possible_fine",
        "infringement",
        "faculty",
    }
    try:
        if item.get('faculty') == 'Unmapped':
            abbr = 'UNM'
        else:
            abbr = item.get('faculty')
        faculty = await Faculty.get(abbreviation=abbr)
    except Exception as e:
        warn(f"Error getting faculty with {item.get('faculty')}: {e}")
    try:

        item['faculty'] = faculty
        item['material_id'] = int(item.get('material_id'))
        item['last_change'] = datetime.strptime(item.get('last_change'), "%Y-%m-%d") if item.get('last_change') else None
        item['retrieved_from_copyright_on'] = datetime.strptime(item.get('retrieved_from_copyright_on'), "%Y-%m-%d")
        item['pagecount'] = int(item.get('pagecount'))
        item['wordcount'] = int(item.get('wordcount'))
        item['picturecount'] = int(item.get('picturecount'))
        item['reliability'] = int(item.get('reliability'))
        item['pages_x_students'] = int(item.get('pages_x_students'))
        item['count_students_registered'] = int(item.get('count_students_registered'))
        item['filetype'] = item.get('filetype', 'unknown') if item.get('filetype') else 'unknown'
        final_dict = {}
        for key in item.keys():
            if key in copyright_item_keys:
                final_dict[key] = item[key]

        final_item = CopyrightItem(**final_dict)
        return final_item

    except Exception as e:
        warn(f'Error while trying to create CopyrightItem with mat_id {item['material_id']}:{e}')
        return None
async def load_raw_items(file: File | None = None) -> None:
    """
    Loads in new items from copyright export raw data.
    Either give a specific file to read in, or default to newest regular raw copyright export.
    """
    error = None
    info(f'# of items in db before loading raw items: {await CopyrightItem.all().count()}')
    if file:
        info(f'loading raw items from {file}')
    else:
        info(f'loading raw items from most recent raw copyright export')
    try:
        file_date, df = read_copyright_export(file)
        items = standardize_dataframe(df).to_dicts()

        item_list = []

        existing_mat_ids = await CopyrightItem.all().values("material_id")
        existing_mat_ids = {int(m['material_id']) for m in existing_mat_ids}

        info(f'Read in {len(items)} raw copyright items. {len(existing_mat_ids)} items already in db.')
        num_total = len(items)

        for item in items:
            if int(item.get('material_id')) in existing_mat_ids:
                continue
            created_item: CopyrightItem = await dict_to_copyright_item(item)
            if not created_item:
                inp = input('Error parsing raw item. enter x to stop, anything else to continue')
                if inp.lower() == 'x':
                    break
                continue
            item_list.append(created_item)
            if num_total % 100 == 0:
                info(f'Parsed {len(item_list)}/{num_total} items.')
    except Exception as e:
        warn(f"Error loading raw items: {e}")
        error = e
    finally:
        if item_list:
            await CopyrightItem.bulk_create(objects=item_list)
        cool(f'# of items in db after loading raw items: {await CopyrightItem.all().count()}')
        if error:
            raise error
    await Tortoise.close_connections()
async def update_copyright_items(data: pl.DataFrame) -> None:
    """
    Update the db with copyrightitems from the dataframe.
    Adds new if they don't exist, or updates if they do.
    For updates, see `compare_items` and the dicts added_fields, changeable_fields, core_fields for details
    """

    def compare_fields(new_item:dict, db_item:CopyrightItem, fielddict: dict, changed:bool, changed_fields:set) -> tuple[bool, set, CopyrightItem]:
        for field, ordering in fielddict.items():
            new_value = new_item.get(field)
            old_value = getattr(db_item, field)
            try:

                if isinstance(old_value, datetime):
                        new_value = datetime.strptime(new_value, "%Y-%m-%d").replace(tzinfo=timezone.utc)
                        old_value = old_value.replace(tzinfo=timezone.utc)
                if isinstance(old_value, Enum):
                    old_value = old_value.value
                if isinstance(old_value, float):
                    new_value = round(float(new_value),2)
                    old_value = round(old_value,2)
                if isinstance(old_value, int):
                    new_value = int(new_value)
            except Exception as e:
                logger.debug(f'error {e} while typecasting data for field comparison of {field}')
                continue
            if new_value:
                if old_value == new_value:
                    continue

                if not old_value:
                    logger.debug(f'[No old value] [{field}] {old_value} --> {new_value}')
                    setattr(db_item, field, new_value)
                    changed = True
                    changed_fields.add(field)

                elif isinstance(ordering,list):

                    new_rank = 20
                    old_rank = 20
                    if new_value in ordering:
                        new_rank = ordering.index(new_value)
                    if old_value in ordering:
                        old_rank = ordering.index(old_value)

                    if new_rank < old_rank:
                        logger.debug(f'[new rank > old rank] [{field}] {old_value} --> {new_value}')

                        setattr(db_item, field, new_value)
                        changed = True
                        changed_fields.add(field)

                else:
                    if isinstance(new_value, str) and isinstance(old_value, str):
                        new_value = new_value.strip()
                        old_value = old_value.strip()
                        if len(new_value) > len(old_value):
                            logger.debug(f'[len(new str) > len(old str)] [{field}] {old_value} --> {new_value}')

                            setattr(db_item, field, new_value)
                            changed = True
                            changed_fields.add(field)

                    else:
                        if type(new_value) == type(old_value):
                            if new_value > old_value:
                                logger.debug(f'[new > old] [{field}] {old_value} --> {new_value}')

                                setattr(db_item, field, new_value)
                                changed = True
                                changed_fields.add(field)
                        else:
                            logger.debug(f'[incomparable types] [{field}] {type(new_value)=}, {type(old_value)=}')


        return changed,changed_fields, db_item

    await init()
    # Fields added by script.
    # dict with field name as key,
    # value are the ordered possible values:in case of conflict, take the earliest value
    # for values that are None, sort and take the highest/latest/... or implement some other logic
    added_fields = {
        'workflow_status':[WorkflowStatus.Done.value, WorkflowStatus.InProgress.value, WorkflowStatus.ToDo.value],
        'retrieved_from_copyright_on': None,
        'possible_fine': None,
        'infringement':[Infringement.YES.value, Infringement.NO.value, Infringement.UNDETERMINED.value],
    }

    # fields changable by checkers. See above for details
    changeable_fields = {
        'manual_classification':[
            Classification.OPEN_ACCESS.value,
            Classification.KORTE_OVERNAME.value,
            Classification.MIDDELLANGE_OVERNAME.value,
            Classification.LANGE_OVERNAME.value,
            Classification.EIGEN_MATERIAAL_POWERPOINT.value,
            Classification.EIGEN_MATERIAAL_TITELINDICATIE.value,
            Classification.EIGEN_MATERIAAL_OVERIG.value,
            Classification.EIGEN_MATERIAAL.value,
            Classification.ONBEKEND.value,
            Classification.LICENTIE_BESCHIKBAAR.value,
            Classification.NIET_GEANALYSEERD.value,
            Classification.IN_ONDERZOEK.value,
            Classification.VERWIJDERVERZOEK_VERSTUURD.value,
            ],
        'manual_identifier':None,
        'remarks': None,
        'scope': None,
    }

    core_fields = {
        "title",
        "classification",
        "ml_prediction",
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
    }
    # standardize df
    # loop over items
    # if item is not in db: add it
    # else compare values in specific fields to determine if we need to update
    info(f'Received {len(data)} raw copyright items as input for an update.')

    data = standardize_dataframe(data)
    existing_mat_ids = await CopyrightItem.all().values("material_id")
    existing_mat_ids = {int(m['material_id']) for m in existing_mat_ids}

    new_items = data.with_columns(pl.col('material_id').cast(int)).filter(~pl.col('material_id').is_in(existing_mat_ids)).to_dicts()
    info(f'# of new items: {len(new_items)}')
    new_objects = []
    if new_items:
        new_objects = [await dict_to_copyright_item(item) for item in new_items]
        new_objects = [item for item in new_objects if item]
        try:
            await CopyrightItem.bulk_create(objects=new_objects)
        except Exception as e:
            warn(f'error while trying to bulk save items. Error: {e}. Trying one-by-one.')
            for item in new_objects:
                try:
                    await item.save()
                except Exception as e:
                    logger.error(e)
                    warn(f'error {e} while trying to save item {item} with material_id {item.material_id}. Skipping for now.')

    cool(f'Created {len(new_objects)} new copyright items in db.')

    update_items = data.with_columns(pl.col('material_id').cast(int)).filter(pl.col('material_id').is_in(existing_mat_ids)).to_dicts()
    info(f'Updating {len(update_items)} existing items.')
    changelist = []
    changed_fields = set()
    for new_item in update_items:
        changed = False
        try:
            db_item = await CopyrightItem.get(material_id=new_item.get('material_id'))

            changed, changed_fields, db_item = compare_fields(new_item, db_item, added_fields, changed, changed_fields)
            changed, changed_fields, db_item = compare_fields(new_item, db_item, changeable_fields, changed, changed_fields)

            if new_item.get('last_change'):
                # if new_item has a newer last_change value, we need to update the core CopyRight fields
                item_last_changed = None
                db_last_changed = None
                try:
                    item_last_changed = datetime.strptime(new_item.get('last_change'), "%Y-%m-%d").replace(tzinfo=timezone.utc)
                    db_last_changed = db_item.last_change.replace(tzinfo=timezone.utc) if db_item.last_change else None
                except Exception as e:
                    pass

                if isinstance(item_last_changed, datetime) and isinstance(db_last_changed, datetime):
                    if item_last_changed > db_last_changed:
                        # replace the core fields in the db item with new values
                        for field in core_fields:
                            if new_item.get('field'):
                                if new_item.get('field') != getattr(db_item, field):
                                    setattr(db_item, field, new_item.get(field))
                                    changed = True
                                    changed_fields.add(field)
                                    logger.debug(f'[core field] Changing field {field} for item {db_item.material_id} from {getattr(db_item, field)} to {new_item.get(field)}')
        except Exception as e:
            warn(f'Could not update item {new_item.get("material_id")}: {e}')
        finally:
            if changed:
                changelist.append(db_item)

    if changelist:
        info(f'Updating {len(changelist)} items in db for fields: {changed_fields}.')
        await CopyrightItem.bulk_update(changelist, fields=changed_fields)

    cool(f'Done updating!')
    await Tortoise.close_connections()

async def load_new_llm_classifications() -> None:
    """
    load llm classifications from .json files in the classifications dir
    """
    await init()
    existing_classifications = await LLMClassification.all().values("used_material_id")
    existing_classifications = {int(m['used_material_id']) for m in existing_classifications}
    info(f'# of existing llm classifications: {len(existing_classifications)}')
    # load jsons to list of dicts
    data_list: list[dict] = []
    try:
        all_files = [f for f in SETTINGS.dirs[DirSetting.CLASSIFICATIONS].files if f.name.endswith('.json')]
        old_files = [f for f in all_files if '_old' in f.name]
        if old_files:
            info(f'Found {len(old_files)} old json files. Removing...')
            for f in old_files:
                f.delete()

        all_files = [f for f in SETTINGS.dirs[DirSetting.CLASSIFICATIONS].files if f.name.endswith('.json')]
        json_mat_ids = {int(f.name.replace('.json','').strip()):f for f in all_files}
        info(f'# of jsons with data found: {len(json_mat_ids)}')
        remaining_jsons = [json_mat_ids.get(m) for m in json_mat_ids if m not in existing_classifications]
        info(f'# of jsons with data not in db: {len(remaining_jsons)}')
        deletelist: list[File] = []
        for file in remaining_jsons:
            try:
                with open(file.path, 'r', encoding='utf-8', errors='ignore') as f:
                    data = json.load(f)
                    data['used_material_id'] = int(file.name.replace('.json','').strip())
                    if not data.get('allowed_usage') or data.get('allowed_usage') == '':
                        deletelist.append(file)
                    else:
                        data_list.append(data)
            except Exception as e:
                warn(f'Error loading json file {file.name}: {e}')
                continue

        cool(f'Retrieved {len(data_list)} new llm classifications, now adding to db.')
        if deletelist:
            for f in deletelist:
                f.delete()
    except Exception as e:
        warn(f'Error loading llm classifications: {e}')


    if not data_list:
        info('No new llm classifications found.')
    else:
        new_objects = []
        for d in data_list:
            try:
                new_objects.append(LLMClassification(
                    allowed_usage = d.get("allowed_usage",""),
                    allowed_usage_reasoning = d.get("allowed_usage_reasoning",""),
                    copyright_status = d.get("copyright_status",""),
                    copyright_classification_reason = d.get("copyright_classification_reason",""),
                    item_type = d.get("item_type",""),
                    item_type_classification_reason = d.get("item_type_classification_reason",""),
                    pdf_name = d.get("pdf_name",""),
                    publisher_name = d.get("publisher_name",""),
                    copyright_holder = d.get("copyright_holder",""),
                    item_title = d.get("item_title",""),
                    pdf_page_count = int(d.get("pdf_page_count",0)),
                    remarks = d.get("remarks",""),

                    author_names = d.get("author_names",[""]) if isinstance(d.get("author_names",[""]),list) else [d.get('author_names',"")],
                    doi = d.get('doi',[""]) if isinstance(d.get('doi',[""]),list) else [d.get('doi',"")],
                    isbn = d.get('isbn',[""]) if isinstance(d.get('isbn',[""]),list) else [d.get('isbn',"")],
                    source_url = d.get('source_url',[""]) if isinstance(d.get('source_url',[""]),list) else [d.get('source_url',"")],
                    license = d.get('license',[""]) if isinstance(d.get('license',[""]),list) else [d.get('license',"")],
                    topic = d.get('topic',[""]) if isinstance(d.get('topic',[""]),list) else [d.get('topic',"")],
                    used_material_id = int(d.get('used_material_id',0))
                ))
            except Exception as e:
                warn(f'error while trying to create llm classification object. Error: {e}.')
        try:
            await LLMClassification.bulk_create(new_objects)
        except Exception as e:
            warn(f'error while trying to bulk save items. Error: {e}. Trying one-by-one.')
            for item in new_objects:
                try:
                    await item.save()
                except Exception as e:
                    logger.error(e)
                    warn(f'error {e} while trying to save item {item}. Skipping for now.')

        cool(f'Created {len(new_objects)} new llm classifications in db. Now linking to copyright items in db.')
    await link_llm_classifications_to_copyright_items()
    await Tortoise.close_connections()

async def link_llm_classifications_to_copyright_items() -> None:
    """
    Link llm classifications to copyright items
    """
    # get all copyright items that do not have a llm classification
    items_to_update = await CopyrightItem.filter(llm_classification=None).all()
    missing_classifications = []
    # load dedupe info from .replace files
    replace_files = {f.name.split('_')[0]:f.name.split('_')[1] for f in SETTINGS.dirs[DirSetting.CLASSIFICATIONS].files if f.name.endswith('.replace')}

    for item in items_to_update:
        try:
            material_id = item.material_id
            if material_id in replace_files:
                material_id = replace_files[material_id]
            llm_classification = await LLMClassification.get_or_none(used_material_id=material_id)
            if llm_classification:
                item.llm_classification = llm_classification
                await item.save()
            else:
                missing_classifications.append(item.material_id)
        except Exception as e:
            warn(f'Error while trying to get llm classification for item {item.material_id}: {e}')
            continue

    info(f'Missing {len(missing_classifications)} llm classifications of {len(items_to_update)} total items.')
async def link_courses_to_copyright_items() -> None:

    items_w_prefetch = await CopyrightItem.all()
    info(f'got {len(items_w_prefetch)} items from db')

    # for each of the items, extract the course code (see enrichment.py)
    # then match with existing course item in db
    # if missing, add to list to retrieve later
    links_added = 0
    course_codes_found = 0
    for item in items_w_prefetch:
        course_codes = determine_course_code(item.course_code, item.course_name)
        if not course_codes or len(course_codes) == 0:
            warn(f'Could not determine course code for item {item.material_id} with input course code {item.course_code} and course name {item.course_name}.')
        course_codes = list(course_codes)

        for course_code in course_codes:
            if not course_code:
                continue
            try:
                cursuscode = int(course_code)
                course_codes_found += 1

                course = await Course.get_or_none(cursuscode=cursuscode)
                if course:
                    await item.courses.add(course)
                    links_added += 1
            except Exception as e:
                warn(f'Error while trying to get course {course_code} for item {item.material_id}: {e}')
    cool(f'Added {links_added} links to {course_codes_found} found coursecodes.')
async def update_copyright_relations() -> None:
    """
    Go through the copyright items in the db
    use the values of the item fields to find links to other tables.
    """
    await init()
    await load_new_llm_classifications()
    await link_courses_to_copyright_items()
    await Tortoise.close_connections()

def retrieve_full_data() -> pl.DataFrame:
    """
    Retrieve all copyright items from db
    includes all relevant data from other models linked to the items
    """
    query="""WITH CourseDataAggregated AS (
        SELECT
            cdcd.copyright_data_id,
            (SELECT GROUP_CONCAT(cursuscode, ' | ') FROM (SELECT DISTINCT CAST(cd.cursuscode AS TEXT) as cursuscode FROM course_data cd WHERE cd.cursuscode = cdcd.course_id)) AS cursuscodes,
            (SELECT GROUP_CONCAT(programme, ' | ') FROM (SELECT DISTINCT cd.programme FROM course_data cd WHERE cd.cursuscode = cdcd.course_id)) AS programmes,
            (SELECT GROUP_CONCAT(name, ' | ') FROM (SELECT DISTINCT cd.name FROM course_data cd WHERE cd.cursuscode = cdcd.course_id)) AS course_names
        FROM copyright_data_course_data cdcd
    ), PersonDataAggregated AS (
        SELECT
        ce.course_id,
        (SELECT GROUP_CONCAT(email, ' | ') FROM (SELECT DISTINCT pd.email FROM person_data pd WHERE pd.id = ce.person_id)) as course_contacts_emails,
        (SELECT GROUP_CONCAT(main_name, ' | ') FROM (SELECT DISTINCT pd.main_name FROM person_data pd WHERE pd.id = ce.person_id)) as course_contacts_names,
        (SELECT GROUP_CONCAT(abbreviation, ' | ') FROM (SELECT DISTINCT f.abbreviation FROM faculty f LEFT JOIN person_data pd ON pd.faculty_id = f.abbreviation WHERE pd.id = ce.person_id)) as course_contacts_faculties,
        (SELECT GROUP_CONCAT(full_abbreviation, ' | ') FROM (
            SELECT DISTINCT org.full_abbreviation
            FROM organization_data org
            LEFT JOIN person_data_organization_data pdod on pdod.organization_id = org.id
            LEFT JOIN person_data pd ON pd.id = pdod.person_data_id
            WHERE pd.id = ce.person_id
        )) as course_contacts_organizations
        FROM course_employee ce
        WHERE ce.role = 'contact'
    )
    SELECT
        cd.*,
        llm.allowed_usage as llm_allowed_usage,
        llm.allowed_usage_reasoning as llm_allowed_usage_reason,
        llm.copyright_status as llm_copyright,
        llm.copyright_classification_reason as llm_copyright_reason,
        llm.item_type as llm_item_type,
        llm.remarks as llm_remarks,
        llm.item_title as llm_title,
        llm.copyright_holder as llm_copyright_holder,
        llm.publisher_name as llm_publisher,
        llm.isbn as llm_isbn,
        llm.doi as llm_doi,
        llm.source_url as llm_source_url,
        llm.license as llm_license,
        llm.author_names as llm_authors,
        cda.cursuscodes,
        cda.programmes,
        cda.course_names,
        pda.course_contacts_names,
        pda.course_contacts_emails,
        pda.course_contacts_faculties,
        pda.course_contacts_organizations
    FROM copyright_data cd
    LEFT JOIN llm_classification_data llm ON cd.llm_classification_id = llm.id
    LEFT JOIN CourseDataAggregated cda ON cd.material_id = cda.copyright_data_id
    LEFT JOIN copyright_data_course_data cdcd ON cd.material_id = cdcd.copyright_data_id
    LEFT JOIN PersonDataAggregated pda ON cdcd.course_id = pda.course_id;
    """
    engine = create_engine("sqlite:///db.sqlite3")
    df: pl.DataFrame = pl.read_database(query=query, connection=engine.connect(), infer_schema_length=None)
    df = df.drop(
            [
                "llm_classification_id",
                'created_at',
                'modified_at',
                'possible_fine',
                'infringement'
            ]
        ).rename(
            mapping={
                "faculty_id":"faculty"
            }
        ).with_columns(
            [
                pl.col(name=colname).
                str.json_decode(infer_schema_length=None).
                list.join(separator=' | ')
                for colname in [
                    "llm_isbn",
                    "llm_doi",
                    "llm_source_url",
                    "llm_license",
                    "llm_authors"
                ]
            ]
        )
    return df
