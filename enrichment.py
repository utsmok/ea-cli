from utils import info, warn, print
import polars as pl
import json
import bs4
import re
import httpx
import asyncio
from constants import OSIRIS_DATA

def determine_course_code(code: str, name: str) -> set | None:
    """
    For a given course code and name (cols of a copyright item), determine the correct course code(s).
    Returns a set of course codes or None if no valid course code could be found.
    """
    try:
        found = False
        tempresults = set()
        first_try = code.split("-")[1].strip()
        if len(first_try) >= 8 and first_try.isdigit():
            tempresults.add(first_try)
            found = True
        else:
            second_try = name.split(";")[1].split("(")[0]
            for c in second_try.split(","):
                c = c.strip()
                if c.isdigit() and len(c) >= 8:
                    tempresults.add(c)
                    found = True
        if not found:
            warn(f"No valid course code found for {code} - {name}")
            info(
                f"code extraction results: {first_try}, name extraction results: {second_try}"
            )
        return tempresults
    except Exception as e:
        warn(f"Error in determine_course_code: {e}")
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
    not_found = 0
    for item in item_data:
        already_found_codes = []
        if item.get('osiris_course_codes_found'):
            if isinstance(item['osiris_course_codes_found'], str):
                if ' | ' in item['osiris_course_codes_found']:
                    already_found_codes = item['osiris_course_codes_found'].split(" | ")
                else:
                    already_found_codes = already_found_codes.append(item['osiris_course_codes_found'])
                already_found_codes = [i.strip() for i in already_found_codes]

        course_codes = determine_course_code(
            item["course_code"], item["course_name"]
        )

        if not course_codes:
            not_found += 1
            continue

        course_codes = list(course_codes)

        if len(course_codes) < 1:
            not_found += 1
            continue

        course_codes = [i.strip() for i in course_codes]

        proceed = False
        for cur_code in course_codes:
            if cur_code not in already_found_codes:
                proceed = True
                break

        if not proceed:
            already_enriched += 1
            continue

        new_item = dict()
        new_item["material_id"] = item["material_id"]
        if len(course_codes) == 1:
            new_item["osiris_course_codes_found"] = course_codes[0]
        if len(course_codes) > 1:
            new_item["osiris_course_codes_found"] = " | ".join(course_codes)
        found_osiris_data = OSIRIS_DATA.get(course_codes[0], None)

        if not found_osiris_data:
            not_found += 1
            continue

        new_item["osiris_course_code_data_selected"] = course_codes[0]
        new_item["osiris_catalogue_url"] = osiris_cat_link + course_codes[0]
        new_item["osiris_programme"] = found_osiris_data.get("programme")
        if found_osiris_data.get("contacts"):
            contacts: dict[str,dict[str, str|list[dict[str,str]]]] = found_osiris_data.get("contacts")
            if len(contacts) == 1:
                new_item["contact_name"] = list(contacts.keys())[0]
                new_item["contact_email"] = list(contacts.values())[0].get("email")
                if list(contacts.values())[0].get("orgs"):
                    maxlen = 0
                    curabbr = ""
                    for org in list(contacts.values())[0].get("orgs"):
                        if len(org.get("abbr")) > maxlen and any(
                            org.get("abbr").startswith(x)
                            for x in ["EEMCS", "BMS", "TNW", "ET", "ITC"]
                        ):
                            maxlen = len(org.get("abbr"))
                            curabbr = org.get("abbr")
                    if maxlen > 0:
                        new_item["contact_org"] = curabbr
        updated += 1
        enriched_item_data.append(new_item)

    enriched_items_df = pl.DataFrame(enriched_item_data)
    df = df.join(enriched_items_df, on="material_id", how="left")

    for col in df.columns:
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
            else:
                # Rename suffix column to base
                df = df.rename({col: base})
    info(f"{group} enrichment results\n----------------------------\nUpdated:          {updated}/{total}\nAlready enriched: {already_enriched}/{total}\nNot found:        {not_found}/{total}")
    return df

async def update_osiris_data(df: pl.DataFrame) -> None:
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
        jaar: int = 2024  # startjaar academisch jaar, 2024 = 2024-2025
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
        try:
            async with semaphore:
                x = await httpx_client.post(url=url, headers=headers, data=body)

                results = x.json().get("hits", {}).get("hits")
                datadict = dict()
                if not results:
                    return
                else:
                    if len(results) != 1:
                        print(
                            str(len(results))
                            + f" hit(s) for code {input_number} for year {jaar} - {jaar + 1}."
                        )
                        print_details = True

                    for h, result in enumerate(results):
                        print(
                            f"------- Result {h} -----------\n"
                        ) if print_details else None
                        rawdata: dict = result.get("_source")
                        teachers = []
                        # pretty print the raw data
                        print(rawdata.keys()) if print_details else None
                        for key, value in rawdata.items():
                            if (
                                value == ""
                                or not value
                                or value == []
                                or value == {}
                            ):
                                continue
                            gaplen = 25 - len(key)
                            if gaplen <= 0:
                                gaplen = 1
                                key = key[:21] + "..."
                            gap = " " + "─" * (gaplen - 1)
                            if isinstance(value, list):
                                if len(value) == 0:
                                    continue
                                if len(value) == 1:
                                    print(
                                        f"{key}{gap}─ {list(value[0].values())[0]}"
                                    ) if print_details else None
                                    items = list(value[0].values())[0]
                                else:
                                    gap = f"{key}{gap}┬ "
                                    i = 0
                                    items = [
                                        list(item.values())[0] for item in value
                                    ]
                                    itemset = set(items)
                                    items = list(itemset)
                                    for item in items:
                                        i += 1
                                        if i - (len(items)) == 0:
                                            gap = " " * (len(key) + gaplen) + "└ "
                                        elif i == 2:
                                            gap = " " * (len(key) + gaplen) + "├ "
                                        if isinstance(item, dict):
                                            print(
                                                f"{gap}{list(item.values())[0]}"
                                            ) if print_details else None
                                        else:
                                            print(
                                                f"{gap}{item}"
                                            ) if print_details else None
                                if key == "docenten":
                                    if isinstance(items, list):
                                        if len(items) == 1:
                                            teachers = set()
                                            teachers.add(items[0])
                                        else:
                                            teachers = set(items)
                                    elif isinstance(items, set):
                                        teachers = items
                                    elif isinstance(items, str):
                                        teachers = set()
                                        teachers.add(items)

                            else:
                                if "\n" not in str(value):
                                    print(
                                        f"{key}{gap}─ {value}"
                                    ) if print_details else None
                                else:
                                    lines = value.split("\n")
                                    printer = f"{key}{gap}┬ "
                                    i = 0
                                    for line in lines:
                                        i = i + 1
                                        if i - len(lines) == 0:
                                            printer = (
                                                " " * (len(key) + gaplen) + "└ "
                                            )
                                        elif i > 1:
                                            printer = (
                                                f"{' ' * (len(key) + gaplen)}├ "
                                            )
                                        print(
                                            f"{printer}{line}"
                                        ) if print_details else None

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
            print("excption when getting course details")
            print(e)
            return

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
        async with semaphore:
            url = f"https://people.utwente.nl/overview?query={name}"
            r = await httpx_client.get(url, headers=headers)
            print(f"{name} --> {r.request}")
            # print(r.text)
            data = r.text
            pattern = r'data-link="([^"]+)"'

            if data:
                matches = re.findall(pattern, data)
                if matches:
                    new_url: str = "https://people.utwente.nl/" + matches[0]
                    try:
                        r = await httpx_client.get(new_url, headers=headers)
                        page_data = None
                        r.raise_for_status()
                        data = r.text
                        page_data = bs4.BeautifulSoup(data, "lxml")
                        found_name = page_data.find(
                            "h1", class_="pageheader__title"
                        ).strings
                        main_name = ""
                        other_names = []
                        for possible_name in found_name:
                            if not main_name:
                                main_name = possible_name
                            else:
                                other_names.append(
                                    str(possible_name)
                                    .strip()
                                    .replace("(", "")
                                    .replace(")", "")
                                )

                        if not main_name.strip().lower() == name.strip().lower():
                            print(f"{main_name} != input name: {name}")
                            print("still processing")
                        for link in page_data.find_all("a"):
                            if "mailto:" in link.get("href"):
                                email = link.get("href").replace("mailto:", "")

                        orgs = []
                        found_orgs = []
                        faculty = ""
                        facultyabbr = ""
                        org_data = page_data.find_all(
                            class_="widget-linklist--smallicons"
                        )
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
                                    pass

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
                        for link in education_tab.find_all("a"):
                            if "https://utwente.osiris-student.nl" in link.get(
                                "href"
                            ):
                                # course
                                linktext = link.string
                                code, coursename = linktext.split(" - ", 1)
                                courses.append(
                                    {"course_code": code, "course_name": coursename}
                                )
                            if "https://www.utwente.nl/" in link.get("href"):
                                # programme
                                url = link.get("href")
                                programme = link.string
                                programmes.append({"name": programme, "url": url})

                        person_data = {
                            "input_name": name,
                            "main_name": main_name,
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
                        raise e

    # step 1: determine list of courseids to search for
    # each row in the df should have 1 or multiple course codes attached to it.
    # we are going to search for each of these course codes in OSIRIS.
    # we will need to extract these codes first.

    # heuristic:

    # 1. FROM COLUMN COURSE_CODE
    # - from column 'course_code', get the course code as a string
    # - Should look like YYYY - XXXXXXXXXXX - 1A, where YYYY is the year, XXXXXXXXXXX is the course code, and 1A is the period.
    # - split on '-', select the second part.
    # - course code should be numeric and (probably?) 9 digits long.
    # - period is (probably) one value from: JAAR, 1A, 1B, 2A, 2B, 3A, SEM1, SEM2, SEM3

    # example values that should result in extracted course code + period:
    # 2024-191158500-JAAR --> Course code: 191158500, Period: JAAR
    # 2024-201800005-1A --> Course code: 201800005, Period: 1A
    # 2024-202400157-1A --> Course code: 202400157, Period: 1A
    # 2024-201800236-SEM1 --> Course code: 201800236, Period: SEM1
    #
    # example values that should be processed further:
    # 2024-IDVWI-1A --> Course code: IDVWI, Period: 1A --> ERROR: not a valid course code
    # 2024-ELECMSE-1B --> Course code: ELECMSE, Period: 1B --> ERROR: not a valid course code

    # 2. IF NO COURSE CODE FOUND: EXTRACT FROM COLUMN COURSE_NAME
    # - in cases where the 'course code' is a string with only letters, it is likely this course has multiple course codes attached to it.
    # - in this case, a list of all related course codes should be extracted from the 'course name' column.
    # - retrieve the string to parse from the 'course name' column.
    # - split the string on ';'. Retrieve the second part. Split this on '(', keep only the first part. This should give you the course codes separated by commas.
    # - Each course code should consist solely of digits w/ len >= 8.
    # - if no valid codes are found, mark as 'no code found'.

    # example values that should result in extracted course codes:
    # Circuit Analysis 1 and 2; 202001116,202200163 (2024-JAAR) --> Course codes: [202001116, 202200163]
    # Characterization of Nanostructures 2023; 193700010,201600043 (2024-1A) --> Course codes: [193700010, 201600043]
    #
    # example values that should not result in extracted course codes:
    # Circuit Analysis 1 and 2; CA12,CA34 (2024-JAAR) --> Course codes: [CA12, CA34] --> ERROR: no valid course codes -> return empty list

    # first we extract the cols as lists using to_dict()

    course_data_dict = df.select(
        pl.col("course_code"), pl.col("course_name")
    ).to_dict()
    course_code_list = course_data_dict.get("course_code").to_list()
    course_name_list = course_data_dict.get("course_name").to_list()

    # then we build a set of all the course codes we need to look up
    lookup_values = set()
    for code, name in zip(course_code_list, course_name_list):
        result = determine_course_code(code, name)
        lookup_values.update(result)

    if len(lookup_values) == 0:
        info("No course codes found, skipping OSIRIS data enrichment")
        return
    else:
        info(f"Found {len(lookup_values)} course codes to look up in OSIRIS")

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
    with open("osiris_data.json", "w") as f:
        json.dump(course_data_dict, f, indent=4)
    if len(not_found) > 0:
        info(f"{len(not_found)} course codes not found: ")
        for code in not_found:
            print("            " + str(code))

    # now look up all the person data
    persons_to_retrieve = set()
    extended_persons_to_retrieve = set()
    for data in course_data_dict.values():
        if data.get("contacts"):
            persons_to_retrieve.update(data.get("contacts"))
        for field in ["docenten", "examinators"]:
            if data.get(field):
                extended_persons_to_retrieve.update(data.get(field))

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
        json.dump(person_data, open("person_data.json", "w"), indent=4)
    except Exception as e:
        print(e)
        pass
    person_dict = {a.get("input_name"): a for a in person_data}

    # finally, combine the two by adding the contact details to the course data
    info(f"Now enriching each osiris course with detailed contact data.")
    osiris_data_w_contacts = dict()
    for code, entry in course_data_dict.items():
        contactdetails = {}
        if entry.get("contacts"):
            for contact in entry.get("contacts"):
                details = person_dict.get(contact)
                if details:
                    contactdetails[contact] = {
                        "name": details.get("main_name"),
                        "first_name": details.get("other_names")[0],
                        "email": details.get("email"),
                        "faculty": details.get("faculty"),
                        "orgs": details.get("orgs"),
                        "programmes": details.get("programmes"),
                        "people_page": details.get("people_page_url"),
                    }
        entry["contacts"] = contactdetails
        osiris_data_w_contacts[code] = entry
    try:
        json.dump(
            osiris_data_w_contacts,
            open("osiris_data_w_contacts.json", "w"),
            indent=4,
        )
    except Exception as e:
        print(e)
        pass

    info(
        "Done. Stored data in json files:\n    osiris_data.json\n    person_data.json\n    osiris_data_w_contacts.json"
    )
