from easy_access.utils import info, warn, print
from easy_access.settings import SETTINGS, FileSetting, OSIRIS_DATA
import polars as pl
import json
import bs4
import re
import httpx
import asyncio
from dataclasses import dataclass, field
from loguru import logger
from nameparser import HumanName
import Levenshtein
# Dataclasses for osiris_contact parsing / matching



@dataclass(frozen=True)
class Faculty:
    abbreviation: str = field(default="", compare=True)
    name: str = field(default="", compare=True)

    def __str__(self):
        return self.abbreviation

DEFAULT_FACULTIES = {
        'EEMCS': Faculty(abbreviation='EEMCS', name='Electrical Engineering, Mathematics and Computer Science'),
        'BMS': Faculty(abbreviation='BMS', name='Behavioural, Management and Social Sciences'),
        'TNW': Faculty(abbreviation='TNW', name='Science and Technology'),
        'ET': Faculty(abbreviation='ET', name='Engineering Technology'),
        'ITC': Faculty(abbreviation='ITC', name='ITC Faculty'),
    }

@dataclass(frozen=True)
class Department:
    abbreviation: str = ""
    name: str = ""
    faculty: Faculty = None

    def __str__(self):
        if self.faculty:
            return f"{self.faculty.abbreviation}-{self.abbreviation}"
        else:
            return self.abbreviation

@dataclass(frozen=True)
class Group:
    abbreviation: str = ""
    name: str = ""
    department: Department = None

    def __str__(self):
        if self.department:
            if self.department.faculty:
                return f"{self.department.faculty.abbreviation}-{self.department.abbreviation}-{self.abbreviation}"
            else:
                return f"{self.department.abbreviation}-{self.abbreviation}"
        else:
            return self.abbreviation

@dataclass
class Contact:
    raw_input_data: list[dict[str,str]] = field(default_factory=list, compare=False)
    name: str = ""
    email: str = ""
    groups: set[Group] = field(default_factory=set)
    departments: dict[str, Department] = field(default_factory=dict)
    faculties: dict[str, Faculty] = field(default_factory=dict)

    def parse_raw_input(self) -> None:
        # parse all raw_input into Group/Department/Faculty instances
        # and assign them to the correct list(s)

        def add_faculty(abbr: str, name: str = "") -> Faculty | None:
            try:
                if abbr in self.faculties:
                    return self.faculties[abbr]
                if abbr in DEFAULT_FACULTIES:
                    faculty = DEFAULT_FACULTIES[abbr]
                elif abbr == "Department":
                    faculty = None
                else:
                    faculty = Faculty(abbreviation=abbr, name=name)
                self.faculties[abbr] = faculty
                return faculty
            except Exception as e:
                warn(f"Error parsing faculty from raw input {abbr}, {name}: {e}")
                return None

        def add_department(abbr: str, name: str, faculty: Faculty | None, full_abbr: str) -> Department | None:
            try:
                if full_abbr in self.departments:
                    return self.departments[full_abbr]
                department = Department(abbreviation=abbr, name=name, faculty=faculty)
                self.departments[full_abbr] = department
                return department
            except Exception as e:
                warn(f"Error parsing department from raw input {abbr}, {name}, {faculty}: {e}")
                return None

        def add_group(abbr: str, name: str, department: Department) -> Group | None:
            try:
                group = Group(abbreviation=abbr, name=name, department=department)
                self.groups.add(group)
                return group
            except Exception as e:
                warn(f"Error parsing group from raw input {raw_input}: {e}")
                return None

        for raw_input in self.raw_input_data:
            abbr = raw_input.get("abbr", "")
            name = raw_input.get("name", "")

            if '-' not in abbr:
                # should be a faculty
                add_faculty(abbr, name)
            elif abbr.count('-') == 1:
                # abbr should be FACULTYABBR - DEPARTMENTABBR
                # name should be department name
                faculty_abbr, dept_abbr = abbr.split('-')
                dept_name = name
                faculty = add_faculty(faculty_abbr)
                add_department(dept_abbr, dept_name, faculty, abbr)
            elif abbr.count('-') == 2:
                # abbr should be FACULTYABBR - DEPARTMENTABBR - GROUPABBR
                # name should be group name
                faculty_abbr, dept_abbr, group_abbr = abbr.split('-')
                full_dept_abbr = f"{faculty_abbr}-{dept_abbr}"
                group_name = name
                faculty = add_faculty(faculty_abbr)
                department = add_department(dept_abbr, "", faculty, full_dept_abbr)
                add_group(group_abbr, group_name, department)

    def add_raw_input(self, raw_input: dict[str,str]) -> None:
        if raw_input not in self.raw_input_data:
            self.raw_input_data.append(raw_input)

    def update(self) -> tuple[list[Group],list[Department],list[Faculty]]:
        for group in self.groups:
            if group.department:
                if str(group.department) not in self.departments:
                    self.departments[str(group.department)] = group.department

        for dept in self.departments.values():
            if dept.faculty:
                if str(dept.faculty) not in self.faculties:
                    self.faculties[str(dept.faculty)]=dept.faculty

    def get_data(self) -> dict[str,str|list[Group]|list[Department]|list[Faculty]]:
        """
        Returns a dict with sheet colnames as keys and deduplicated data as values
        """
        self.parse_raw_input()
        self.update()

        return {
            "contact_name": self.name,
            "contact_email": self.email,
            "contact_groups": list(self.groups),
            "contact_departments": list(self.departments.values()),
            "contact_faculties": list(self.faculties.values()),
            }



@dataclass
class ItemContacts:
    """
    Keeps track of all Contacts for a single item
    Mainly used to output strs with found unique depts/groups/faculties
    """
    contacts: list[Contact] = field(default_factory=list)

    def get_merged_contact_data(self) -> dict[str,str]:
        """
        Returns a dict with sheet colnames as keys with merged unique strings as values
        """
        data: dict[str, set[str|Group|Department|Faculty]] = {
            "contact_name": set(),
            "contact_email": set(),
            "contact_groups": set(),
            "contact_departments": set(),
            "contact_faculties": set()
        }
        for contact in self.contacts:
            contact_data = contact.get_data()
            for key, value in contact_data.items():
                if isinstance(value, list):
                    data[key].update(value)
                else:
                    data[key].add(value)

        final_data: dict[str,str] = dict()
        for key, value in data.items():
            if not value:
                value = ""
            elif len(value) == 1:
                value = str(value.pop())
            else:
                value = " | ".join([str(x) for x in list(value) if x])
            final_data[key] = value

        return final_data

def determine_course_code(code: str, name: str) -> set[str|None]:
    """
    For a given course code and name (cols of a copyright item; canvas data), determine the correct osiris course code(s).
    Returns a set of course codes; if no valid course code could be found it will be empty.
    Heuristic is as follows:

    STEP 1: attempt to parse canvas course code into osiris course code(s)
        - from column 'course_code', get the course code as a string
        - Should look like YYYY - XXXXXXXXXXX - 1A, where YYYY is the year, XXXXXXXXXXX is the course code, and 1A is the period.
        - split on '-', select the second part.
        - course code should be numeric and (probably?) 9 digits long.
        - period is (probably) one value from: JAAR, 1A, 1B, 2A, 2B, 3A, SEM1, SEM2, SEM3

        EXAMPLES:
            should result in extracted course code + period:
                2024-191158500-JAAR
                    --> Course code: 191158500, Period: JAAR
                    --> return {191158500}
                2024-201800005-1A
                    --> Course code: 201800005, Period: 1A
                    --> return {201800005}
                2024-202400157-1A
                    --> Course code: 202400157, Period: 1A
                    --> return {202400157}
                2024-201800236-SEM1
                    --> Course code: 201800236, Period: SEM1
                    --> return {201800236}

            should be processed further:
                2024-IDVWI-1A
                    --> Course code: IDVWI, Period: 1A
                    --> ERROR: not a valid course code
                    --> continue to step 2
                2024-ELECMSE-1B
                    --> Course code: ELECMSE, Period: 1B
                    --> ERROR: not a valid course code
                    --> continue to step 2

    If step 1 fails:
    STEP 2: attempt to parse canvas course name into osiris course code(s)
    in cases where the 'course code' is a string with only letters, it is likely this course has multiple course codes attached to it.
    in this case, the set of related course codes should be extracted from the 'course name' column.
        1. retrieve the string to parse from the 'course name' column.
        2. split the string on ';'. Split the second item of result on '(', select the first item of that result. This should give a string of course codes separated by commas.
        3. split on ',' and loop over results
        4. For each: if str with only digits and len >= 8: add to result set, set found to True.

        EXAMPLES:
            should result in extracted course codes:
                Circuit Analysis 1 and 2; 202001116,202200163 (2024-JAAR)
                    --> return {202001116, 202200163}
                Characterization of Nanostructures 2023; 193700010,201600043 (2024-1A)
                    --> return {193700010, 201600043}

            should not result in extracted course codes:
                Circuit Analysis 1 and 2; CA12,CA34 (2024-JAAR)
                    --> Course codes found: [CA12, CA34]
                    --> ERROR: invalid course codes
                    --> return empty list

    """
    def is_valid_course_code(check_code) -> bool:
        try:
            check_code = str(check_code).strip()

            if check_code.isdigit() and len(check_code) >= 8:
                return True
            else:
                return False
        except Exception as e:
            return False

    tempresults = set()
    first_try = ""
    second_try = ""

    try:
        first_try = code.split("-")[1].strip()
        tempresults.add(first_try) if is_valid_course_code(first_try) else None
        if (';' in name) and ('(' in name):
            second_try = name.split(";")[1].split("(")[0]
            [tempresults.add(c.strip()) for c in second_try.split(",") if is_valid_course_code(c)]

        if not tempresults:
            warn(f"No valid course code found for {code} - {name}")
            info(
                f"code extraction results: {first_try}, name extraction results: {second_try}"
            )
        return tempresults
    except Exception as e:
        warn(f"Error in determine_course_code for input: code={code}, name={name}: {e}")
        return tempresults

def enrich_df_with_osiris_data(df: pl.DataFrame, group:str = "all items") -> pl.DataFrame:
    """
    Read in OSIRIS/people page data.
    Enrich the supplied df with the information contained in the jsons.
    Return the enriched dataframe.
    """

    item_data = df.select(
        pl.col("course_code"), pl.col("course_name"), pl.col("material_id")
    ).to_dicts()
    enriched_item_data = []
    osiris_cat_link = "https://utwente.osiris-student.nl/onderwijscatalogus/extern/cursus/zoek?trefwoord="
    info(f"Enriching {len(item_data)} items for {group} with OSIRIS data.")
    total = len(item_data)
    updated = 0
    already_enriched = 0
    no_course_code_found = set()
    no_osiris_data_found = set()
    for item in item_data:
        already_found_codes = []
        if item.get('osiris_course_codes_found'):
            if isinstance(item['osiris_course_codes_found'], str):
                if ' | ' in item['osiris_course_codes_found']:
                    already_found_codes = item['osiris_course_codes_found'].split(" | ")
                else:
                    already_found_codes = already_found_codes.append(item['osiris_course_codes_found'])
                already_found_codes = [i.strip() for i in already_found_codes]

        if already_found_codes:
            already_enriched += 1
            continue

        course_codes = determine_course_code(
            item["course_code"], item["course_name"]
        )

        if not course_codes:
            no_course_code_found.add(item["course_code"])
            continue

        course_codes = list(course_codes)


        course_codes = [i.strip() for i in course_codes]


        new_item = dict()
        new_item["material_id"] = item["material_id"]
        if len(course_codes) == 1:
            new_item["osiris_course_codes_found"] = course_codes[0]
        if len(course_codes) > 1:
            new_item["osiris_course_codes_found"] = " | ".join(course_codes)
        used_codes = []
        found_osiris_data = []
        maybe_used_codes = []
        maybe_found_data = []
        for code in course_codes:
            new_osiris_data = OSIRIS_DATA.get(code, None) if OSIRIS_DATA else None
            if not new_osiris_data:
                no_osiris_data_found.add(code)
            else:
                if new_osiris_data.get('faculty','') == item.get('faculty'):
                    used_codes.append(code)
                    found_osiris_data.append(new_osiris_data)
                else:
                    maybe_used_codes.append(code)
                    maybe_found_data.append(new_osiris_data)

        if not found_osiris_data:
            if maybe_found_data and maybe_used_codes:
                used_codes = maybe_used_codes
                found_osiris_data = maybe_found_data

        if len(used_codes) == 1:
            used_code = used_codes[0]
            new_item["osiris_course_codes_used"] = used_code
            new_item["osiris_course_code_data_selected"] = used_code

            new_item["osiris_catalogue_url"] = osiris_cat_link + used_code
            new_item["osiris_programme"] = found_osiris_data[0].get("programme")

        elif len(used_codes) > 1:
            new_item["osiris_course_codes_used"] = " | ".join(used_codes)
            new_item["osiris_course_code_data_selected"] = used_codes[0]
            new_item["osiris_catalogue_url"] = osiris_cat_link + used_codes[0]
            new_item["osiris_programme"] = " | ".join([osiris_data.get("programme", "") for osiris_data in found_osiris_data if osiris_data.get("programme")])
            used_code = " | ".join(used_codes)

        course_contacts = ItemContacts()
        for osiris_data in found_osiris_data:
            if osiris_data.get("contacts"):
                contacts: dict[str,dict[str, str|list[dict[str,str]]]] = osiris_data.get("contacts")
                for contact_name, contact in contacts.items():
                    faculty_dict_start = {}
                    if contact.get('faculty'):
                        if contact.get('faculty') in DEFAULT_FACULTIES:
                            faculty_dict_start = {contact.get("faculty"): DEFAULT_FACULTIES.get(contact.get("faculty"))}
                    cur_contact = Contact(name=contact_name, email=contact.get("email"), faculties=faculty_dict_start)
                    if contact.get('orgs'):
                        [cur_contact.add_raw_input(org) for org in contact.get('orgs')]
                    course_contacts.contacts.append(cur_contact)

        new_item.update(course_contacts.get_merged_contact_data())
        updated += 1
        enriched_item_data.append(new_item)

    enriched_items_df = pl.DataFrame(enriched_item_data)
    df = df.join(enriched_items_df, on="material_id", how="left")

    for col in df.columns:
        # TODO: improve handling of conflicting values

        if col.endswith("_left") or col.endswith("_right"):
            base = col.replace("_left","").replace("_right","")
            # Drop if suffix column is all null
            if df.select(pl.col(col).is_null().all()).item(0,0):
                df = df.drop(col)
                continue
            # If base column exists, compare
            if base in df.columns:

                same_vals = df.select((pl.col(base).fill_null(value='') == pl.col(col).fill_null(value='')).all()).item(0,0)
                if same_vals:
                    df = df.drop(col)
                else:

                    # If base is all null, replace it
                    if df.select(pl.col(base).is_null().all()).item(0,0):
                        df = df.drop(base)
                        df = df.rename({col: base})
                    # if not: we have 2 cols that do not match. Merge them in some way.
                    else:
                        final_data = []
                        #warn(f"Conflicting values for column {base} in {group} enrichment results. Overwriting with new values.")
                        all_base_vals = df.select(pl.col(base)).to_series().to_list()
                        all_new_vals = df.select(pl.col(col)).to_series().to_list()
                        material_ids = df.select(pl.col("material_id")).to_series().to_list()
                        mismatched_cols = set()
                        if base == 'osiris_catalogue_url':
                            #info('Not printing comparison for osiris_catalogue_urls.')
                            ...
                        elif len(all_base_vals) != len(all_new_vals):
                            #warn(f'Cannot print comparison: cols are not the same length. \nlen({base})={len(all_base_vals)} != len({col})={len(all_new_vals)}')
                            ...
                        else:
                            #info("Mismatches found:")
                            printstr =""
                            for material_id, orig_val, new_val in zip(material_ids, all_base_vals, all_new_vals):
                                if not orig_val:
                                    orig_val = ""
                                if not new_val:
                                    new_val = ""
                                orig_data = orig_val
                                new_data = new_val

                                if ' | ' in orig_val:
                                    orig_val = orig_val.split(' | ')
                                    if isinstance(orig_val,list):
                                        if len(orig_val) > 1:
                                            orig_val = set(orig_val)
                                if ' | ' in new_val:
                                    new_val = new_val.split(' | ')
                                    if isinstance(new_val,list):
                                        if len(new_val) > 1:
                                            new_val = set(new_val)

                                if orig_val != new_val:
                                    if not orig_val or orig_val == "":
                                        final_data.append({'material_id':material_id, base:new_data})
                                        printstr += (f'\nmaterial_id: {material_id}\n  {str(orig_val)}\n    --> {str(new_val)}')
                                    elif not new_val or new_val == "":
                                        final_data.append({'material_id':material_id, base:orig_data})
                                        printstr += (f'\nmaterial_id: {material_id}\n  Keeping  {str(orig_val)}\n  NOT replacing with {str(new_val)}')
                                    else:
                                        final_data.append({'material_id':material_id, base:new_data})
                                        printstr += (f'\nmaterial_id: {material_id}\n  Keeping  {str(orig_val)}\n  NOT replacing with {str(new_val)}')
                                    mismatched_cols.add(material_id)
                                else:
                                    # no new val
                                    final_data.append({'material_id':material_id, base:orig_data})
                            #info(printstr)

                        if not final_data:
                            df = df.drop(base).rename({col: base})
                        else:
                            # drop base and col from orig df
                            # create temp dataframe from final_data -- has cols material_id and base
                            # join temp dataframe to df on material_id
                            #info(f'replacing {base} with col constructed from {len(final_data)} items')
                            tmp_df = pl.from_dicts(final_data, strict=False, infer_schema_length=1000)
                            #info(f'tmp_df details: {tmp_df.shape[0]} rows, cols: {tmp_df.columns}')
                            #info(tmp_df)
                            df = df.drop([base, col]).join(tmp_df, on="material_id", how="left")
            else:
                # base doesn't exist? weird, just rename suffix column to base and done
                df = df.rename({col: base})
    info(f"{group} enrichment results\n----------------------------\nUpdated:          {updated}/{total}\nAlready enriched: {already_enriched}/{total}\nNot in osiris:    {len(no_osiris_data_found)}/{total}\nNo coursecode:    {len(no_course_code_found)}/{total}")
    return df

async def update_osiris_data(df: pl.DataFrame, only_retrieve_missing: bool = False) -> None:
    """
    For a given df with copyright items, retrieve all OSIRIS course data + person data from people pages.

    Stores the data as 3 jsons in the ea-cli dir root; to be used for enriching later.
    """

    async def get_data_from_osiris(
        input_number: int,
        httpx_client: httpx.AsyncClient,
        semaphore: asyncio.Semaphore,
        jaar: int = 2024,
    ) -> dict[str, dict[str, str | list | set]]:
        print_details = False
        startstring: str = '{"from":0,"size":25,"sort":[{"cursus_lange_naam.raw":{"order":"asc"}},{"cursus":{"order":"asc"}},{"collegejaar":{"order":"desc"}}],"aggs":{"agg_terms_collegejaar":{"filter":{"bool":{"must":[]}},"aggs":{"agg_collegejaar_buckets":{"terms":{"field":"collegejaar","size":2500,"order":{"_term":"desc"}}}}},"agg_terms_blokken_nested.periode_omschrijving":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_blokken_nested.periode_omschrijving":{"terms":{"field":"blokken_nested.periode_omschrijving","size":2500,"order":{"_term":"asc"},"exclude":"Periode: [0-9][0-9]-[0-9][0-9]-[0-9][0-9][0-9][0-9]"}},"nested_aggs":{"nested":{"path":"blokken_nested"},"aggs":{"nested_aggs":{"filter":{"bool":{"must":[]}},"aggs":{"agg_blokken_nested.periode_omschrijving_buckets":{"terms":{"field":"blokken_nested.periode_omschrijving","size":2500,"order":{"_term":"asc"},"exclude":"Periode: [0-9][0-9]-[0-9][0-9]-[0-9][0-9][0-9][0-9]"},"aggs":{"items":{"reverse_nested":{}}}}}}}}}},"agg_terms_faculteit_naam":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_faculteit_naam_buckets":{"terms":{"field":"faculteit_naam","size":2500,"order":{"_term":"asc"}}}}},"agg_terms_coordinerend_onderdeel_oms":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_coordinerend_onderdeel_oms_buckets":{"terms":{"field":"coordinerend_onderdeel_oms","size":2500,"order":{"_term":"asc"}}}}},"agg_terms_categorie_omschrijving":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_categorie_omschrijving_buckets":{"terms":{"field":"categorie_omschrijving","size":2500,"order":{"_term":"asc"}}}}},"agg_terms_voertalen.voertaal_omschrijving":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_voertalen.voertaal_omschrijving_buckets":{"terms":{"field":"voertalen.voertaal_omschrijving","size":2500,"order":{"_term":"asc"}}}}}},"post_filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"query":{"bool":{"must":[{"multi_match":{"query":'
        if jaar != 2024:
            if isinstance(jaar, int):
                startstring.replace('"2024-2025"', f'"{jaar}-{jaar + 1}"')
            elif jaar == "":
                startstring.replace('"2024-2025"', "")

        code: str = f'"{input_number}"'
        endstring: str = ',"type":"phrase_prefix","fields":["cursus","cursus_korte_naam","cursus_lange_naam"],"max_expansions":200}}]}}}'
        body: str = startstring + code + endstring
        url: str = "https://utwente.osiris-student.nl/student/osiris/student/cursussen/zoeken"
        headers: dict[str, str] = {
            "host": "utwente.osiris-student.nl",
            "connection": "keep-alive",
            "content-length": "2183",
            "sec-ch-ua-platform": '"Windows"',
            "authorization": "undefined undefined",
            "cache-control": "no-cache, no-store, must-revalidate, private",
            "pragma": "no-cache",
            "client_type": "web",
            "release_version": "c0d3b6a1d72bf1610166027c903b46fc10580f30",
            "manifest": "24.46_B346_c0d3b6a1",
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua": '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
            "accept": "application/json, text/plain, */*",
            "content-type": "application/json",
            "taal": "NL",
            "origin": "https//utwente.osiris-student.nl",
            "sec-fetch-site": "same-origin",
            "sec-fetch-mode": "cors",
            "sec-fetch-dest": "empty",
            "referer": "https//utwente.osiris-student.nl/onderwijscatalogus/extern/cursussen",
            "accept-encoding": "gzip, deflate, br, zstd",
            "accept-language": "en-GB,en-US;q=0.9,en;q=0.8",
        }
        retry = False
        try:
            async with semaphore:
                x = await httpx_client.post(url=url, headers=headers, data=body)
                results = x.json().get("hits", {}).get("hits")
                datadict = dict()
                if not results:
                    if not jaar:
                        warn(f'No data found for code {input_number}.')
                        return
                    elif str(jaar) == "2018":
                        jaar = ""
                        retry = True
                    else:
                        jaar = jaar - 1
                        retry = True
                else:
                    if len(results) != 1:
                        info(
                            str(len(results))
                            + f" hit(s) for code {input_number} for year {jaar} - {jaar + 1}."
                        )
                        print_details = True

                    def process_teacher_items(items: str | list | set) -> set[str]:
                        """Helper function to process teacher items into a consistent set format"""
                        if isinstance(items, list):
                            return set(items)
                        elif isinstance(items, str):
                            return {items}
                        elif isinstance(items, set):
                            return items
                        return set()

                    for result in results:
                        rawdata: dict = result.get("_source")
                        teachers = set()
                        for key, value in rawdata.items():
                            if (
                                value == ""
                                or not value
                                or value == []
                                or value == {}
                            ):
                                continue
                            if isinstance(value, list):
                                if not value:
                                    continue
                                items = (list(value[0].values())[0] if len(value) == 1
                                        else list({list(item.values())[0] for item in value}))

                                if key == "docenten":
                                    teachers = process_teacher_items(items)

                        datadict[rawdata.get("cursus")] = {
                            "cursuscode": rawdata.get("cursus"),
                            "internal_id": rawdata.get("id_cursus"),
                            "year": rawdata.get("collegejaar"),
                            "short_name": rawdata.get("cursus_korte_naam"),
                            "name": rawdata.get("cursus_lange_naam"),
                            "faculty": rawdata.get("faculteit"),
                            "faculty_long": rawdata.get("faculteit_naam"),
                            "programme": rawdata.get("coordinerend_onderdeel_oms"),
                            "ec": rawdata.get("punten"),
                            "language": [
                                x.get("voertaal_omschrijving")
                                for x in rawdata.get("voertalen")
                            ],
                            "notes": rawdata.get("opmerking_cursus"),
                            "category": rawdata.get("categorie_omschrijving"),
                            "teachers": teachers,
                            "contacts": set(),
                            "docenten": set(),
                            "examinators": set(),
                            "unknown_role": set(),
                            "tutors": set(),
                        }
                        print("\n") if print_details else None

                    headers_course = {
                        "accept": "application/json, text/plain, */*",
                        "accept-language": "en-US,en;q=0.9,nl-NL;q=0.8,nl;q=0.7",
                        "authorization": "undefined undefined",
                        "cache-control": "no-cache, no-store, must-revalidate, private",
                        "client_type": "web",
                        "content-type": "application/json",
                        "dnt": "1",
                        "manifest": "24.46_B346_c0d3b6a1",
                        "pragma": "no-cache",
                        "priority": "u=1, i",
                        "referer": "https://utwente.osiris-student.nl/onderwijscatalogus/extern/cursussen",
                        "release_version": "c0d3b6a1d72bf1610166027c903b46fc10580f30",
                        "sec-ch-ua": '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
                        "sec-ch-ua-mobile": "?0",
                        "sec-ch-ua-platform": '"Windows"',
                        "sec-fetch-dest": "empty",
                        "sec-fetch-mode": "cors",
                        "sec-fetch-site": "same-origin",
                        "taal": "NL",
                        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
                    }
                    newdatadict = datadict.copy()
                    for course, data in datadict.items():
                        internal_id = data.get("internal_id")
                        url_course = f"https://utwente.osiris-student.nl/student/osiris/owc/cursussen/{internal_id}"
                        course_details = httpx.get(
                            url=url_course,
                            headers=headers_course,
                        )

                        if course_details.status_code == 200:
                            course_data = course_details.json()
                            for datapoint in course_data.get("items"):
                                if datapoint.get("rubriek") == "rubriek-docenten":
                                    docentdata = datapoint.get("velden")
                            if docentdata:
                                for docentitem in docentdata:
                                    if docentitem.get("waarde"):
                                        for docenttype in docentitem.get("waarde"):
                                            for persoon in docenttype.get("velden"):
                                                if (
                                                    docenttype.get("omschrijving")
                                                    == "Contactpersoon"
                                                ):
                                                    newdatadict[course][
                                                        "contacts"
                                                    ].add(persoon.get("docent"))
                                                elif (
                                                    docenttype.get("omschrijving")
                                                    == "Docent"
                                                ):
                                                    newdatadict[course][
                                                        "docenten"
                                                    ].add(persoon.get("docent"))
                                                elif (
                                                    docenttype.get("omschrijving")
                                                    == "Examinator"
                                                ):
                                                    newdatadict[course][
                                                        "examinators"
                                                    ].add(persoon.get("docent"))
                                                elif (
                                                    docenttype.get("omschrijving")
                                                    == "Tutor"
                                                ):
                                                    newdatadict[course][
                                                        "tutors"
                                                    ].add(persoon.get("docent"))
                                                else:
                                                    try:
                                                        newdatadict[course][
                                                            "unknown_role"
                                                        ].add(persoon.get("docent"))
                                                    except Exception as e:
                                                        pass
                                for field in [
                                    "teachers",
                                    "docenten",
                                    "examinators",
                                    "tutors",
                                    "unknown_role",
                                    "contacts",
                                ]:
                                    if isinstance(
                                        newdatadict[course].get(field, None), set
                                    ):
                                        newdatadict[course][field] = list(
                                            newdatadict[course][field]
                                        )
                                        if len(
                                            newdatadict[course][field]
                                        ) > 8 and all(
                                            len(x) == 1
                                            for x in newdatadict[course][field]
                                        ):
                                            newdatadict[course][field] = []

                        else:
                            print("Error!")
                            print(course_details.status_code)

                    print(newdatadict) if print_details else None
                    return newdatadict
        except Exception as e:
            print("exception when getting course details")
            logger.exception(e)
            return
        if retry:
            return await get_data_from_osiris(input_number, httpx_client, semaphore, jaar)
    async def get_data_from_people_page(
        name: str, httpx_client: httpx.AsyncClient, semaphore: asyncio.Semaphore
    ) -> dict:
        url: str = "https://people.utwente.nl/overview"
        headers: dict = {
            "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "accept-language": "en-US,en;q=0.9",
            "priority": "u=0, i",
            "sec-ch-ua": '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"',
            "sec-fetch-dest": "document",
            "sec-fetch-mode": "navigate",
            "sec-fetch-site": "same-origin",
            "sec-fetch-user": "?1",
            "upgrade-insecure-requests": "1",
        }
        def strip_name(name: str) -> str:
            stripped_name = name.strip()
            titles = ['ing.', 'dr.', 'prof.', 'ir.', "rer.", 'nat.', ', MSc', ', PhD', ', BSc']
            for title in titles:
                stripped_name=stripped_name.replace(title,'').strip()
            return stripped_name
        def remove_dot_and_lower(name: str) -> str:
            return str(name).strip().replace(".","").lower()
        async with semaphore:
            url = f"https://people.utwente.nl/overview?query={name}"
            r = await httpx_client.get(url, headers=headers)
            print(f"{name} --> {r.request}")
            # print(r.text)
            data = r.text
            pattern = r'data-link="([^"]+)"'
            name_parsed_str = strip_name(name)
            compare_name = remove_dot_and_lower(name_parsed_str)

            if data:
                matches = re.findall(pattern, data)
                if matches:
                    if len(matches) >= 10:
                        matches = matches[:5]
                    best_match = matches[0]
                    ratio = Levenshtein.ratio(compare_name, remove_dot_and_lower(best_match))
                    if ratio < 0.8:
                        for match in matches:
                            if 'business' in match or '/' in match:
                                continue
                            new_ratio = Levenshtein.ratio(compare_name, remove_dot_and_lower(match))

                            if new_ratio > ratio:
                                best_match = match
                                ratio = new_ratio
                                if ratio > 0.8:
                                    break
                    if ratio < 0.7:
                        warn(f'Low match confidence {ratio}: best match for {name} is {best_match}. Actual comparison:\nfound: {remove_dot_and_lower(best_match)} vs input {compare_name})')


                    new_url: str = "https://people.utwente.nl/" + best_match
                    try:
                        r = await httpx_client.get(new_url, headers=headers)
                        page_data = None
                        if r.status_code in [500, 502]:
                            return await get_data_from_people_page(name, httpx_client, semaphore)
                        data = r.text
                        page_data = bs4.BeautifulSoup(data, "lxml")
                        main_name = ""
                        other_names = []
                        email = ""
                        if not page_data:
                            warn(f'No page data found for {name} at {new_url}')
                            return
                        found_name = page_data.find(
                            "h1", class_="pageheader__title"
                        )
                        if found_name:
                            for possible_name in found_name.strings:
                                if not main_name:
                                    main_name = possible_name
                                else:
                                    other_names.append(
                                        str(possible_name)
                                        .strip()
                                        .replace("(", "")
                                        .replace(")", "")
                                    )
                            final_ratio = Levenshtein.ratio(name_parsed_str, strip_name(main_name))

                            if final_ratio < 0.7:
                                print(f"found name {main_name} differs from input name: {name} with ratio {final_ratio}. Actually compared strings: found: {strip_name(main_name)} | input: {strip_name(name)}")
                                print("still processing...")

                        try:
                            for link in page_data.find_all("a"):
                                if "mailto:" in link.get("href"):
                                    email = link.get("href").replace("mailto:", "")
                        except Exception as e:
                            logger.exception(e)
                            email = ""

                        orgs = []
                        found_orgs = []
                        faculty = ""
                        facultyabbr = ""
                        org_data = page_data.find_all(
                            class_="widget-linklist--smallicons"
                        )
                        if not org_data:
                            org_data = []
                        if len(org_data) >= 1:
                            org_data = org_data[0].find_all(
                                class_="widget-linklist__text"
                            )
                        else:
                            org_data = []
                        for org in org_data:
                            text = org.string
                            if "(" in text:
                                try:
                                    orgname = text.split("(")[0]
                                    orgabbr = text.split("(")[1].split(")")[0]
                                    if orgabbr in [
                                        "BMS",
                                        "ET",
                                        "EEMCS",
                                        "ITC",
                                        "TNW",
                                    ]:
                                        faculty = orgname
                                        facultyabbr = orgabbr
                                    else:
                                        found_orgs.append(
                                            {"name": orgname, "abbr": orgabbr}
                                        )
                                except Exception as e:
                                    logger.exception(f'error while processing org {text}: {e}')

                        if faculty and facultyabbr and found_orgs:
                            orgs.append({"name": faculty, "abbr": facultyabbr})
                            for org in found_orgs:
                                if facultyabbr in org.get("abbr"):
                                    cleaned_abbr = org.get("abbr").replace(
                                        "-" + facultyabbr, ""
                                    )
                                    orgs.append(
                                        {
                                            "name": org.get("name"),
                                            "abbr": cleaned_abbr,
                                        }
                                    )
                                    continue
                                orgs.append(
                                    {
                                        "name": org.get("name"),
                                        "abbr": org.get("abbr"),
                                    }
                                )

                        education_tab = page_data.find(
                            "div", id="tabpanel-education"
                        )
                        courses = []
                        programmes = []
                        if education_tab:
                            for link in education_tab.find_all("a"):
                                if "https://utwente.osiris-student.nl" in link.get(
                                    "href"
                                ):
                                    # course
                                    linktext = link.string if link.string else ""
                                    if " - " in linktext:
                                        code, coursename = linktext.split(" - ", 1)
                                        courses.append(
                                            {"course_code": code, "course_name": coursename}
                                        )
                                if "https://www.utwente.nl/" in link.get("href"):
                                    # programme
                                    url = link.get("href")
                                    programme = link.string
                                    if url and programme:
                                        programmes.append({"name": programme, "url": url})

                        person_data = {
                            "input_name": name,
                            "main_name": main_name,
                            "match_confidence": final_ratio,
                            "other_names": other_names,
                            "email": email,
                            "orgs": orgs,
                            "courses": courses,
                            "programmes": programmes,
                            "faculty": facultyabbr,
                            "people_page_url": new_url,
                        }
                        return person_data

                    except Exception as e:
                        print(
                            f"error while retrieving / processing {new_url} for person {name}"
                        )
                        logger.exception(e)
                        raise e


    '''
    each row in the df should have 1 or multiple osiris course codes attached to it.
    Extract them using determine_course_code(code, name) with
        code: canvas code from column course_code
        name: canvas course name from column course_name
    '''
    course_data_dict = df.select(
        pl.col("course_code"), pl.col("course_name")
    ).to_dict()
    course_code_list = course_data_dict.get("course_code").to_list()
    course_name_list = course_data_dict.get("course_name").to_list()
    lookup_values = set()
    for code, name in zip(course_code_list, course_name_list):
        result = determine_course_code(code, name)
        lookup_values.update(result)

    if len(lookup_values) == 0:
        info("No course codes found, skipping OSIRIS data enrichment")
        return
    else:
        info(f"Found {len(lookup_values)} course codes to look up in OSIRIS")


    osiris_data_w_contacts_file = {}
    try:
        osiris_data_w_contacts_file = json.load(
            open(SETTINGS.files[FileSetting.OSIRIS_DATA_W_CONTACTS].path, "r", encoding="utf-8")
        )
    except Exception as e:
        warn(f'couldnt load {SETTINGS.files[FileSetting.OSIRIS_DATA_W_CONTACTS].path}')
        print(e)
        pass

    course_codes_already_retrieved = set(osiris_data_w_contacts_file.keys())
    retrieve_course_data = True
    if only_retrieve_missing:
        cur_osiris_data = json.load(open(SETTINGS.files[FileSetting.OSIRIS_DATA].path, "r", encoding="utf-8"))
        cur_osiris_data = {k:v for k,v in cur_osiris_data.items() if v}
        lookup_values = lookup_values - course_codes_already_retrieved
        info(f"{len(lookup_values)} remaining course codes to look up in OSIRIS after filtering out already retrieved course codes")
        if len(lookup_values) == 0:
            info("All course data already retrieved!")
            retrieve_course_data = False
            course_data_dict = cur_osiris_data

    if retrieve_course_data:
        # then retrieve data from OSIRIS for each of the values in lookup_values
        course_data_dict = {}
        not_found = set()
        found_amount = 0
        max_concurrent = 10
        semaphore = asyncio.Semaphore(max_concurrent)  # Rate limiting with semaphore
        async with httpx.AsyncClient(timeout=60) as client:
            tasks = []
            for code in lookup_values:
                if code in course_data_dict:
                    continue
                task1 = asyncio.create_task(
                    get_data_from_osiris(
                        httpx_client=client, input_number=code, semaphore=semaphore
                    )
                )
                course_data_dict[code] = {}
                tasks.append((code, task1))

            for code, task in tasks:
                result = await task  # Get the result of the task
                if result:
                    course_data_dict.update(result)
                    found_amount += 1
                else:
                    result = await get_data_from_osiris(
                        httpx_client=client,
                        input_number=code,
                        jaar="",
                        semaphore=semaphore,
                    )
                    if result:
                        course_data_dict.update(result)
                        found_amount += 1
                    else:
                        not_found.add(code)

        info(
            f"Found {found_amount} course codes in OSIRIS from {len(lookup_values)} starting course codes."
        )
        # store course_data_dict as a json file
        if only_retrieve_missing:
            course_data_dict.update(cur_osiris_data)
        with open(SETTINGS.files[FileSetting.OSIRIS_DATA].path, "w") as f:
            json.dump(course_data_dict, f, indent=4)
        if len(not_found) > 0:
            info(f"{len(not_found)} course codes not found: ")
            for code in not_found:
                print("            " + str(code))
    # now look up all the person data
    persons_to_retrieve = set()
    extended_persons_to_retrieve = set()
    person_data = list()
    for data in course_data_dict.values():
        if data.get("contacts"):
            persons_to_retrieve.update(data.get("contacts"))
        for data_field in ["docenten", "examinators"]:
            if data.get(data_field):
                extended_persons_to_retrieve.update(data.get(data_field))
    info(f'{len(persons_to_retrieve)} persons in current osiris data to enrich')

    if only_retrieve_missing:
        try:
                cur_person_data = json.load(open(SETTINGS.files[FileSetting.PERSON_DATA].path, "r", encoding="utf-8"))
                cur_persons = {x.get('input_name') for x in cur_person_data}
                persons_to_retrieve = persons_to_retrieve - set(cur_persons)
                extended_persons_to_retrieve = extended_persons_to_retrieve - set(cur_persons)
                info(f'{len(persons_to_retrieve)} persons remaining after filtering out already retrieved persons')
        except Exception as e:
            warn(f'error while loading {SETTINGS.files[FileSetting.PERSON_DATA].path}: {e}')
            ...
    if len(persons_to_retrieve) > 0:
        info(f"now retrieving person data for {len(persons_to_retrieve)} people.")
        person_data = []
        persontasks = []
        async with httpx.AsyncClient(timeout=30) as client:
            for person in persons_to_retrieve | extended_persons_to_retrieve:
                persontasks.append(
                    asyncio.create_task(
                        get_data_from_people_page(
                            person, httpx_client=client, semaphore=semaphore
                        )
                    )
                )
            for task in persontasks:
                try:
                    parsed_data = await task
                    if parsed_data:
                        person_data.append(parsed_data)
                except Exception as e:
                    print(e)
                    pass

        info(f"got data for {len(person_data)} persons")
        try:
            if only_retrieve_missing:
                current_person_data = json.load(open(SETTINGS.files[FileSetting.PERSON_DATA].path, "r", encoding="utf-8"))
                person_data.extend(current_person_data)

            json.dump(person_data, open(SETTINGS.files[FileSetting.PERSON_DATA].path, "w", encoding="utf-8"), indent=4)
        except Exception as e:
            print('error while dumping person data')
            print(e)
            pass
    if len(person_data) == 0:
        try:
            person_data = json.load(open(SETTINGS.files[FileSetting.PERSON_DATA].path, "r", encoding="utf-8"))
        except Exception as e:
            warn(f'couldnt load {SETTINGS.files[FileSetting.PERSON_DATA].path}: {e}')
            person_dict = {}

    person_dict = {a.get("input_name"): a for a in person_data}

    # finally, combine the two by adding the contact details to the course data
    info("Now enriching each osiris course with detailed contact data.")
    osiris_data_w_contacts = dict()
    for code, entry in course_data_dict.items():
        if not entry:
            print(f'No osiris data found for course code {code}')
            continue
        contactdetails = {}
        if entry.get("contacts"):
            for contact in entry.get("contacts"):
                details = person_dict.get(contact)

                if details:
                    contactdetails[contact] = {
                        "name": details.get("main_name"),
                        "first_name": details.get("other_names",[""])[0],
                        "email": details.get("email"),
                        "faculty": details.get("faculty"),
                        "orgs": details.get("orgs"),
                        "programmes": details.get("programmes"),
                        "people_page": details.get("people_page_url"),
                    }
                    if not details.get("orgs"):
                        warn(f'No orgs found for contact {contact} with details:')
                        info(details)
                else:
                    warn(f'No details found for contact {contact}')

        entry["contacts"] = contactdetails
        osiris_data_w_contacts[code] = entry
        if entry.get("contacts") == {}:
            print(f'No contact details found for course code {code}')
            print(f'osiris course data:')
            print(entry)

    try:
        if only_retrieve_missing:
            current_osiris_data_w_contacts = json.load(
                open(SETTINGS.files[FileSetting.OSIRIS_DATA_W_CONTACTS].path, "r", encoding="utf-8")
            )
            osiris_data_w_contacts.update(current_osiris_data_w_contacts)
        json.dump(
            osiris_data_w_contacts,
            open(SETTINGS.files[FileSetting.OSIRIS_DATA_W_CONTACTS].path, "w", encoding="utf-8"),
            indent=4,
        )
    except Exception as e:
        print(e)
        pass

    info(
        f"Done. Stored data in json files:\n    {SETTINGS.files[FileSetting.OSIRIS_DATA]}\n    {SETTINGS.files[FileSetting.PERSON_DATA]}\n    {SETTINGS.files[FileSetting.OSIRIS_DATA_W_CONTACTS]}"
    )
