"""
This module is responsible for enriching copyright item data by fetching
additional information from external university systems, specifically Osiris
(for course details) and the university's people pages (for staff details).

It uses asynchronous HTTP requests (`httpx`) to query these sources,
parses the results (JSON for Osiris, HTML for people pages using BeautifulSoup),
and structures the data for integration into the main dataset. The enriched
data is typically stored in JSON files specified in the application settings.
"""

import asyncio
import json
import logging
import re
from typing import Any, Coroutine, Dict, List, Set, Tuple # Corrected type hints

import bs4
import httpx
import Levenshtein
import polars as pl

from easy_access.db.ingest import load_base_data # Used by update_osiris_data at the end
from easy_access.settings import SETTINGS, FileSetting
from easy_access.utils import determine_course_code # Assuming this is the correct location

logger = logging.getLogger(__name__)

# --- Configuration Candidates (to be moved to settings.yaml ideally) ---
OSIRIS_SEARCH_URL: str = "https://utwente.osiris-student.nl/student/osiris/student/cursussen/zoeken"
OSIRIS_COURSE_DETAIL_URL_TEMPLATE: str = "https://utwente.osiris-student.nl/student/osiris/owc/cursussen/{internal_id}"
PEOPLE_PAGE_OVERVIEW_URL_TEMPLATE: str = "https://people.utwente.nl/overview?query={name}"
PEOPLE_PAGE_BASE_URL: str = "https://people.utwente.nl/"
# The large Osiris query body ('startstring' below) is also a major candidate for external configuration.
ACADEMIC_TITLES_TO_STRIP: List[str] = ["ing.", "dr.", "prof.", "ir.", "rer.", "nat.", ", MSc", ", PhD", ", BSc"]
# --- End Configuration Candidates ---


async def _get_data_from_osiris(
    input_number: int,
    httpx_client: httpx.AsyncClient,
    semaphore: asyncio.Semaphore,
    jaar: int | str = 2024, # Allow empty string for specific retry logic
) -> Dict[str, Dict[str, Any]]: # Return type reflects dict of course_code -> course_data
    """
    Fetches detailed course data from Osiris for a given course code and year.

    It first queries the Osiris search endpoint, then fetches detailed contact
    information from the course-specific page. It includes retry logic for
    different academic years if initial attempts yield no results.

    Args:
        input_number (int): The course code (cursuscode) to search for.
        httpx_client (httpx.AsyncClient): An active httpx client for making requests.
        semaphore (asyncio.Semaphore): Semaphore for limiting concurrent requests.
        jaar (int | str, optional): The academic year to query. Defaults to 2024.
                                    Can be an empty string for a specific retry query.

    Returns:
        Dict[str, Dict[str, Any]]: A dictionary where keys are course codes (as strings)
                                   and values are dictionaries of course data.
                                   Returns an empty dictionary if no data is found or an error occurs.
    """
    # Osiris search query body parts (highly specific and brittle)
    # TODO: Externalize this query structure if possible, or make more robust.
    osiris_query_start: str = ('{"from":0,"size":25,"sort":[{"cursus_lange_naam.raw":{"order":"asc"}},'
                               '{"cursus":{"order":"asc"}},{"collegejaar":{"order":"desc"}}],"aggs":{'
                               '"agg_terms_collegejaar":{"filter":{"bool":{"must":[]}},"aggs":{'
                               '"agg_collegejaar_buckets":{"terms":{"field":"collegejaar","size":2500,'
                               '"order":{"_term":"desc"}}}}},"agg_terms_blokken_nested.periode_omschrijving":{'
                               '"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{'
                               '"agg_blokken_nested.periode_omschrijving":{"terms":{"field":'
                               '"blokken_nested.periode_omschrijving","size":2500,"order":{"_term":"asc"},'
                               '"exclude":"Periode: [0-9][0-9]-[0-9][0-9]-[0-9][0-9][0-9][0-9]"}},"nested_aggs":{'
                               '"nested":{"path":"blokken_nested"},"aggs":{"nested_aggs":{"filter":{"bool":{'
                               '"must":[]}},"aggs":{"agg_blokken_nested.periode_omschrijving_buckets":{'
                               '"terms":{"field":"blokken_nested.periode_omschrijving","size":2500,"order":{'
                               '"_term":"asc"},"exclude":"Periode: [0-9][0-9]-[0-9][0-9]-[0-9][0-9][0-9][0-9]"},'
                               '"aggs":{"items":{"reverse_nested":{}}}}}}}}}},"agg_terms_faculteit_naam":{'
                               '"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{'
                               '"agg_faculteit_naam_buckets":{"terms":{"field":"faculteit_naam","size":2500,'
                               '"order":{"_term":"asc"}}}}},"agg_terms_coordinerend_onderdeel_oms":{'
                               '"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{'
                               '"agg_coordinerend_onderdeel_oms_buckets":{"terms":{"field":'
                               '"coordinerend_onderdeel_oms","size":2500,"order":{"_term":"asc"}}}}},'
                               '"agg_terms_categorie_omschrijving":{"filter":{"bool":{"must":[{"terms":{'
                               '"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_categorie_omschrijving_buckets":{'
                               '"terms":{"field":"categorie_omschrijving","size":2500,"order":{"_term":"asc"}}}}}},'
                               '"agg_terms_voertalen.voertaal_omschrijving":{"filter":{"bool":{"must":[{'
                               '"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{'
                               '"agg_voertalen.voertaal_omschrijving_buckets":{"terms":{"field":'
                               '"voertalen.voertaal_omschrijving","size":2500,"order":{"_term":"asc"}}}}}}},'
                               '"post_filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},'
                               '"query":{"bool":{"must":[{"multi_match":{"query":')

    osiris_query_code_str: str = f'"{input_number}"'
    osiris_query_end: str = (',"type":"phrase_prefix","fields":["cursus","cursus_korte_naam",'
                             '"cursus_lange_naam"],"max_expansions":200}}]}}}')

    current_year_filter = f"{jaar}-{int(jaar) + 1}" if isinstance(jaar, int) else "2024-2025" # Default if jaar is empty string
    if jaar == "": # Special case for retry to query all years
        final_query_body_str = (osiris_query_start.replace(f'[{{"terms":{{"collegejaar":["2024-2025"]}}}}]','[]') +
                                osiris_query_code_str + osiris_query_end)
    else:
        final_query_body_str = (osiris_query_start.replace('"2024-2025"', f'"{current_year_filter}"') +
                                osiris_query_code_str + osiris_query_end)

    # Headers for Osiris requests (mimicking browser)
    # TODO: Some headers might be unnecessary or could be simplified. User-Agent is important.
    osiris_headers: Dict[str, str] = {
        "host": "utwente.osiris-student.nl", "connection": "keep-alive",
        "sec-ch-ua-platform": '"Windows"', "authorization": "undefined undefined",
        "cache-control": "no-cache, no-store, must-revalidate, private", "pragma": "no-cache",
        "client_type": "web", "release_version": "c0d3b6a1d72bf1610166027c903b46fc10580f30", # These might change
        "manifest": "24.46_B346_c0d3b6a1", # These might change
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua": '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "accept": "application/json, text/plain, */*", "content-type": "application/json", "taal": "NL",
        "origin": "https//utwente.osiris-student.nl", "sec-fetch-site": "same-origin",
        "sec-fetch-mode": "cors", "sec-fetch-dest": "empty",
        "referer": "https//utwente.osiris-student.nl/onderwijscatalogus/extern/cursussen",
        "accept-encoding": "gzip, deflate, br, zstd", "accept-language": "en-GB,en-US;q=0.9,en;q=0.8",
    }

    should_retry_year: bool = False
    processed_course_data: Dict[str, Dict[str, Any]] = {}

    try:
        async with semaphore: # Limit concurrency
            response = await httpx_client.post(url=OSIRIS_SEARCH_URL, headers=osiris_headers, data=final_query_body_str)
            response.raise_for_status() # Raise HTTPStatusError for bad responses (4xx or 5xx)

            results_json = response.json()
            hits = results_json.get("hits", {}).get("hits", [])

            if not hits:
                if jaar and str(jaar) == "2018": # Specific year to switch to all-years query
                    jaar = "" # Trigger all-years search
                    should_retry_year = True
                elif isinstance(jaar, int) and jaar > 2018 : # Try previous year
                    jaar -= 1
                    should_retry_year = True
                else: # No more retry options for year
                    logger.debug(f"No Osiris data found for course code {input_number} (year: {current_year_filter}).")
            else:
                if len(hits) > 1:
                    logger.info(f"{len(hits)} Osiris search hits for code {input_number} (year {current_year_filter}). Processing all.")

                for hit in hits:
                    raw_data = hit.get("_source", {})
                    course_code_from_hit = raw_data.get("cursus")
                    if not course_code_from_hit: continue

                    # Helper to process lists of teacher/staff names
                    def _extract_names(items: Any) -> Set[str]:
                        if isinstance(items, list): return {str(item).strip() for item in items if item}
                        if isinstance(items, str): return {items.strip()} if items.strip() else set()
                        return set()

                    teachers_set: Set[str] = set()
                    if isinstance(raw_data.get("docenten"), list): # Assuming 'docenten' is a list of dicts
                        for docente_item_list in raw_data.get("docenten", []):
                             # This structure was complex, assuming 'docenten' contains list of names directly or needs parsing
                             if isinstance(docente_item_list, dict): # e.g. {"naam": "Name"}
                                 teachers_set.add(str(list(docente_item_list.values())[0]))
                             elif isinstance(docente_item_list, str) : # e.g. ["Name1", "Name2"]
                                 teachers_set.add(docente_item_list)

                    current_course_details: Dict[str, Any] = {
                        "cursuscode": course_code_from_hit,
                        "internal_id": raw_data.get("id_cursus"), "year": raw_data.get("collegejaar"),
                        "short_name": raw_data.get("cursus_korte_naam"), "name": raw_data.get("cursus_lange_naam"),
                        "faculty": raw_data.get("faculteit"), "faculty_long": raw_data.get("faculteit_naam"),
                        "programme": raw_data.get("coordinerend_onderdeel_oms"), "ec": raw_data.get("punten"),
                        "language": [lang_item.get("voertaal_omschrijving") for lang_item in raw_data.get("voertalen", []) if lang_item],
                        "notes": raw_data.get("opmerking_cursus"), "category": raw_data.get("categorie_omschrijving"),
                        "teachers": list(teachers_set), # Store as list
                        "contacts": [], "docenten": [], "examinators": [], "tutors": [], "unknown_role": [] # Initialize for detail page
                    }

                    # Fetch detailed contact page
                    internal_id = current_course_details.get("internal_id")
                    if internal_id:
                        detail_url = OSIRIS_COURSE_DETAIL_URL_TEMPLATE.format(internal_id=internal_id)
                        # Re-use headers, or define specific ones if needed
                        details_resp = await httpx_client.get(url=detail_url, headers=osiris_headers)
                        if details_resp.status_code == 200:
                            course_detail_json = details_resp.json()
                            for section in course_detail_json.get("items", []):
                                if section.get("rubriek") == "rubriek-docenten": # Staff section
                                    for field_group in section.get("velden", []):
                                        role_description = field_group.get("omschrijving")
                                        for staff_entry_list in field_group.get("waarde", []):
                                            for staff_member_field in staff_entry_list.get("velden",[]):
                                                staff_name = staff_member_field.get("docent")
                                                if staff_name:
                                                    if role_description == "Contactpersoon": current_course_details["contacts"].append(staff_name)
                                                    elif role_description == "Docent": current_course_details["docenten"].append(staff_name)
                                                    elif role_description == "Examinator": current_course_details["examinators"].append(staff_name)
                                                    elif role_description == "Tutor": current_course_details["tutors"].append(staff_name)
                                                    else: current_course_details["unknown_role"].append(staff_name)
                        else:
                            logger.debug(f"Failed to get course details from {detail_url}, status: {details_resp.status_code}")
                    processed_course_data[str(course_code_from_hit)] = current_course_details

    except httpx.HTTPStatusError as e_http:
        logger.warning(f"HTTP error fetching Osiris data for {input_number} (year {jaar}): {e_http.response.status_code} - {e_http.request.url}")
    except json.JSONDecodeError as e_json:
        logger.warning(f"JSON decode error for Osiris data {input_number} (year {jaar}): {e_json}")
    except Exception as e_generic:
        logger.error(f"Unexpected error in _get_data_from_osiris for {input_number} (year {jaar}): {e_generic}")
        logger.debug(traceback.format_exc())
        return {} # Return empty on error to prevent cascading failures

    if should_retry_year: # If a retry was triggered by previous logic
        logger.debug(f"Retrying _get_data_from_osiris for {input_number} with year parameter: '{jaar}'")
        return await _get_data_from_osiris(input_number, httpx_client, semaphore, jaar)

    return processed_course_data


def _strip_academic_titles(name: str) -> str:
    """Removes common academic titles from a name string for better matching."""
    stripped_name = name.strip()
    for title in ACADEMIC_TITLES_TO_STRIP: # Use configured list
        stripped_name = stripped_name.replace(title, "").strip()
    return stripped_name

def _normalize_name_for_matching(name: str) -> str:
    """Removes dots and converts to lowercase for robust name matching."""
    return str(name).strip().replace(".", "").lower()


async def _get_data_from_people_page(
    name: str, httpx_client: httpx.AsyncClient, semaphore: asyncio.Semaphore
) -> Dict[str, Any] | None:
    """
    Fetches staff details from the university's people pages for a given name.

    It searches for the name, attempts to find the best match from search results
    using Levenshtein distance, then scrapes details (name, email, organizations,
    courses, programmes, faculty) from the individual's profile page.

    Args:
        name (str): The name of the person to search for.
        httpx_client (httpx.AsyncClient): An active httpx client.
        semaphore (asyncio.Semaphore): Semaphore for limiting concurrent requests.

    Returns:
        Dict[str, Any] | None: A dictionary containing scraped staff data,
                               or None if no match is found or an error occurs.
    """
    # Headers for people page requests
    people_page_headers: Dict[str, str] = {
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "accept-language": "en-US,en;q=0.9", "priority": "u=0, i",
        "sec-ch-ua": '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
        "sec-ch-ua-mobile": "?0", "sec-ch-ua-platform": '"Windows"',
        "sec-fetch-dest": "document", "sec-fetch-mode": "navigate",
        "sec-fetch-site": "same-origin", "sec-fetch-user": "?1", "upgrade-insecure-requests": "1",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    }

    normalized_input_name = _strip_academic_titles(name)
    compare_name_normalized = _normalize_name_for_matching(normalized_input_name)

    async with semaphore:
        search_url = PEOPLE_PAGE_OVERVIEW_URL_TEMPLATE.format(name=name) # Use configured template
        try:
            search_response = await httpx_client.get(search_url, headers=people_page_headers)
            search_response.raise_for_status()
            search_html = search_response.text

            # Regex to find profile links from search results
            profile_link_pattern = r'data-link="([^"]+)"' # Assumes links are in data-link attribute
            matches = re.findall(profile_link_pattern, search_html)

            if not matches:
                logger.debug(f"No search results on people page for name: '{name}'")
                return None

            # Find best match using Levenshtein distance
            best_match_link_suffix: str | None = None
            best_ratio: float = 0.0
            # Limit number of matches to check for performance
            for match_suffix in matches[:10]: # Check top 10 results
                # Skip common non-person links if any patterns are known
                if "business" in match_suffix or "/" not in match_suffix: # Example filter
                    continue

                # Extract name from suffix for comparison (e.g. "j.p.smith" from "j.p.smith/profile")
                # This part is brittle if URL structure changes.
                name_from_suffix = match_suffix.split("/")[0] if "/" in match_suffix else match_suffix
                current_ratio = Levenshtein.ratio(compare_name_normalized, _normalize_name_for_matching(name_from_suffix))

                if current_ratio > best_ratio:
                    best_match_link_suffix = match_suffix
                    best_ratio = current_ratio
                    if best_ratio > 0.95: break # Confident match

            if not best_match_link_suffix or best_ratio < 0.7: # Adjust threshold as needed
                logger.info(f"Low match confidence ({best_ratio:.2f}) for '{name}'. Best guess: '{best_match_link_suffix}'. Skipping.")
                return None

            profile_url = PEOPLE_PAGE_BASE_URL + best_match_link_suffix # Use configured base
            profile_response = await httpx_client.get(profile_url, headers=people_page_headers)
            profile_response.raise_for_status() # Check for errors on profile page

            profile_html = profile_response.text
            soup = bs4.BeautifulSoup(profile_html, "lxml")

            # Scrape details from profile page (selectors are highly site-specific)
            # TODO: These selectors need to be verified against current people.utwente.nl structure.
            main_name_tag = soup.find("h1", class_="pageheader__title") # Example selector
            main_name_scraped = str(main_name_tag.find(string=True, recursive=False) if main_name_tag else "").strip()

            other_names_scraped: List[str] = []
            if main_name_tag and main_name_tag.find_all("span"): # Example for other names/titles
                 other_names_scraped = [s.text.strip().replace("(", "").replace(")", "") for s in main_name_tag.find_all("span")]


            email_scraped: str | None = None
            email_tag = soup.find("a", href=lambda href: href and href.startswith("mailto:"))
            if email_tag: email_scraped = email_tag.get("href", "").replace("mailto:", "")

            # Organization scraping (example, highly dependent on HTML structure)
            orgs_scraped: List[Dict[str, str]] = []
            faculty_scraped: str = ""
            faculty_abbr_scraped: str = ""
            # Example: find a div with class 'organisation' then list items
            org_section = soup.find("div", class_="widget-linklist--smallicons") # This class was in original code
            if org_section:
                for org_item_tag in org_section.find_all("li"): # Assuming orgs are in list items
                    org_text_tag = org_item_tag.find(class_="widget-linklist__text") # Original class
                    if org_text_tag and org_text_tag.string:
                        org_full_text = org_text_tag.string.strip()
                        org_name_part, org_abbr_part = "", ""
                        if "(" in org_full_text and org_full_text.endswith(")"):
                            org_name_part = org_full_text.split("(",1)[0].strip()
                            org_abbr_part = org_full_text.split("(",1)[1][:-1].strip()
                        else:
                            org_name_part = org_full_text

                        # Basic faculty identification (very heuristic)
                        # TODO: Improve faculty identification, perhaps via known abbreviations list from settings
                        if org_abbr_part in ["BMS", "ET", "EEMCS", "ITC", "TNW", "ES", "PP"]: # Example faculty abbrs
                            faculty_scraped = org_name_part
                            faculty_abbr_scraped = org_abbr_part
                        orgs_scraped.append({"name": org_name_part, "abbr": org_abbr_part})

            # Course/Programme scraping (example, highly dependent on HTML)
            courses_scraped: List[Dict[str, str]] = []
            programmes_scraped: List[Dict[str, str]] = []
            education_tab = soup.find("div", id="tabpanel-education") # Original ID
            if education_tab:
                for link_tag in education_tab.find_all("a", href=True):
                    href = link_tag["href"]
                    link_text = (link_tag.string or "").strip()
                    if "utwente.osiris-student.nl" in href and " - " in link_text:
                        code, course_name_text = link_text.split(" - ", 1)
                        courses_scraped.append({"course_code": code.strip(), "course_name": course_name_text.strip()})
                    elif "www.utwente.nl/" in href and link_text: # Assuming these are programme links
                        programmes_scraped.append({"name": link_text, "url": href})

            return {
                "input_name": name, "main_name": main_name_scraped or normalized_input_name,
                "match_confidence": best_ratio, "other_names": other_names_scraped,
                "email": email_scraped, "orgs": orgs_scraped, "courses": courses_scraped,
                "programmes": programmes_scraped, "faculty": faculty_abbr_scraped, # Use abbreviation
                "people_page_url": profile_url,
            }

        except httpx.HTTPStatusError as e_http_profile:
            logger.warning(f"HTTP error fetching profile for '{name}' (URL: {e_http_profile.request.url}): {e_http_profile.response.status_code}")
        except Exception as e_profile:
            logger.error(f"Error processing profile page for '{name}': {e_profile}")
            logger.debug(traceback.format_exc())
        return None


async def update_osiris_data(
    df_copyright_items: pl.DataFrame, only_retrieve_missing: bool = False
) -> None:
    """
    Enriches copyright item data by fetching related information from Osiris and
    university people pages. Stores the fetched data in JSON files.

    Args:
        df_copyright_items (pl.DataFrame): DataFrame of copyright items. Expected to have
                                         'course_code' and 'course_name' columns.
        only_retrieve_missing (bool, optional): If True, only fetches data for courses/persons
                                               not already present in existing JSON files.
                                               Defaults to False.
    """
    if df_copyright_items.is_empty():
        logger.info("No copyright items provided to update_osiris_data. Skipping enrichment.")
        return

    logger.info(f"Starting Osiris data update. Only retrieve missing: {only_retrieve_missing}")

    # Determine unique course codes to look up from the input DataFrame
    unique_course_keys_df = df_copyright_items.select(["course_code", "course_name"]).unique()
    lookup_course_codes: Set[str] = set()
    for row in unique_course_keys_df.iter_rows(named=True):
        # determine_course_code expects str, str -> set[str]
        codes_from_row = determine_course_code(str(row.get("course_code","")), str(row.get("course_name","")))
        lookup_course_codes.update(codes_from_row)

    if not lookup_course_codes:
        logger.info("No valid course codes determined from input DataFrame. Skipping Osiris data fetch.")
        # Fall through to person data processing if that's independent or uses other criteria
    else:
        logger.info(f"Found {len(lookup_course_codes)} unique course codes to query in Osiris.")

    # Load existing Osiris data with contact details to check what's missing
    osiris_data_w_contacts_path: Path = SETTINGS.files[FileSetting.OSIRIS_DATA_W_CONTACTS].path
    osiris_data_w_contacts_dict: Dict[str, Any] = {}
    if osiris_data_w_contacts_path.exists():
        try:
            with open(osiris_data_w_contacts_path, "r", encoding="utf-8") as f:
                osiris_data_w_contacts_dict = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            logger.warning(f"Could not load existing Osiris data with contacts from {osiris_data_w_contacts_path}: {e}")

    course_codes_already_retrieved: Set[str] = set(osiris_data_w_contacts_dict.keys())

    # Filter lookup_course_codes if only_retrieve_missing
    if only_retrieve_missing:
        original_lookup_count = len(lookup_course_codes)
        lookup_course_codes.difference_update(course_codes_already_retrieved) # Remove already known
        logger.info(f"{len(lookup_course_codes)} course codes remain after filtering "
                    f"({original_lookup_count - len(lookup_course_codes)} already present).")

    # Fetch Osiris course data
    all_fetched_osiris_courses: Dict[str, Dict[str, Any]] = {}
    if lookup_course_codes: # Only fetch if there are codes to look up
        max_concurrent_osiris: int = 10 # Configurable?
        semaphore_osiris = asyncio.Semaphore(max_concurrent_osiris)
        async with httpx.AsyncClient(timeout=60.0) as client_osiris: # Increased timeout
            osiris_tasks: List[Coroutine[Any, Any, Dict[str, Dict[str, Any]]]] = []
            for code_str in lookup_course_codes:
                if code_str.isdigit(): # Ensure it's a digit string before int conversion
                    osiris_tasks.append(_get_data_from_osiris(int(code_str), client_osiris, semaphore_osiris))

            for result_dict in await asyncio.gather(*osiris_tasks, return_exceptions=True):
                if isinstance(result_dict, dict):
                    all_fetched_osiris_courses.update(result_dict)
                elif isinstance(result_dict, Exception):
                    logger.warning(f"An Osiris data fetching task failed: {result_dict}")
        logger.info(f"Fetched data for {len(all_fetched_osiris_courses)} new course codes from Osiris.")
    else:
        logger.info("No new course codes to fetch from Osiris based on current settings.")

    # Merge new data with existing, then save
    if only_retrieve_missing:
        # Load existing osiris_data.json (without contacts, as per original logic)
        osiris_data_path = SETTINGS.files[FileSetting.OSIRIS_DATA].path
        if osiris_data_path.exists():
            try:
                with open(osiris_data_path, "r", encoding="utf-8") as f_curr:
                    current_osiris_data_no_contacts = json.load(f_curr)
                all_fetched_osiris_courses.update(current_osiris_data_no_contacts)
            except (json.JSONDecodeError, OSError) as e:
                 logger.warning(f"Could not load current osiris_data.json for merging: {e}")

    if all_fetched_osiris_courses: # Save if any data (new or merged)
        try:
            with open(SETTINGS.files[FileSetting.OSIRIS_DATA].path, "w", encoding="utf-8") as f_out:
                json.dump(all_fetched_osiris_courses, f_out, indent=4, ensure_ascii=False)
            logger.info(f"Osiris course data saved to {SETTINGS.files[FileSetting.OSIRIS_DATA].path}")
        except OSError as e:
            logger.error(f"Failed to save Osiris course data: {e}")


    # --- Person Data Enrichment ---
    persons_to_retrieve_set: Set[str] = set()
    # Collect all unique staff names from the *currently available* Osiris course data (newly fetched + existing)
    for course_detail in all_fetched_osiris_courses.values():
        for role_key in ["teachers", "contacts", "docenten", "examinators", "tutors", "unknown_role"]:
            persons_in_role = course_detail.get(role_key, [])
            if isinstance(persons_in_role, list):
                persons_to_retrieve_set.update(str(p).strip() for p in persons_in_role if isinstance(p, str) and str(p).strip())

    logger.info(f"Found {len(persons_to_retrieve_set)} unique person names from Osiris data for potential People Page lookup.")

    current_person_data_list: List[Dict[str, Any]] = []
    person_data_path: Path = SETTINGS.files[FileSetting.PERSON_DATA].path
    if person_data_path.exists():
        try:
            with open(person_data_path, "r", encoding="utf-8") as f_person:
                current_person_data_list = json.load(f_person)
        except (json.JSONDecodeError, OSError) as e:
            logger.warning(f"Could not load existing person data from {person_data_path}: {e}")

    if only_retrieve_missing:
        already_retrieved_persons: Set[str] = {str(p.get("input_name","")).strip() for p in current_person_data_list if p.get("input_name")}
        original_person_lookup_count = len(persons_to_retrieve_set)
        persons_to_retrieve_set.difference_update(already_retrieved_persons)
        logger.info(f"{len(persons_to_retrieve_set)} person names remain for People Page lookup "
                    f"({original_person_lookup_count - len(persons_to_retrieve_set)} already present).")

    newly_fetched_person_data: List[Dict[str, Any]] = []
    if persons_to_retrieve_set:
        max_concurrent_people: int = 5 # Be gentle with people pages
        semaphore_people = asyncio.Semaphore(max_concurrent_people)
        async with httpx.AsyncClient(timeout=30.0) as client_people:
            people_tasks: List[Coroutine[Any, Any, Dict[str, Any] | None]] = [
                _get_data_from_people_page(person_name, client_people, semaphore_people)
                for person_name in persons_to_retrieve_set
            ]
            for result_person_data in await asyncio.gather(*people_tasks, return_exceptions=True):
                if isinstance(result_person_data, dict):
                    newly_fetched_person_data.append(result_person_data)
                elif isinstance(result_person_data, Exception):
                    logger.warning(f"A People Page fetching task failed: {result_person_data}")
        logger.info(f"Fetched data for {len(newly_fetched_person_data)} new persons from People Pages.")

    # Merge new person data with existing and save
    final_person_data_list = current_person_data_list + newly_fetched_person_data
    # Deduplicate based on input_name, keeping the latest entry (though order isn't guaranteed here easily)
    # A simple way is to convert to dict by input_name, then back to list
    deduplicated_person_data_map: Dict[str, Dict[str, Any]] = {}
    for p_data in final_person_data_list:
        if p_data.get("input_name"): # Ensure there's a name to key by
            deduplicated_person_data_map[str(p_data["input_name"])] = p_data

    if deduplicated_person_data_map: # Check if there's anything to save
        try:
            with open(person_data_path, "w", encoding="utf-8") as f_person_out:
                json.dump(list(deduplicated_person_data_map.values()), f_person_out, indent=4, ensure_ascii=False)
            logger.info(f"Person data saved to {person_data_path} ({len(deduplicated_person_data_map)} unique persons).")
        except OSError as e:
            logger.error(f"Failed to save person data: {e}")

    # --- Combine Osiris course data with new contact details ---
    logger.info("Enriching Osiris course data with detailed contact information.")
    final_osiris_data_w_contacts: Dict[str, Any] = {}
    person_details_map: Dict[str, Dict[str, Any]] = {str(p.get("input_name","")): p for p in deduplicated_person_data_map.values()}

    for course_code_str, course_info_dict in all_fetched_osiris_courses.items():
        if not course_info_dict: continue # Skip if course info is empty for some reason

        enriched_contacts: Dict[str, Dict[str, Any]] = {}
        # The original logic used 'contacts' key from Osiris for this detailed enrichment
        # but other roles like 'teachers', 'docenten' also contain names.
        # For now, sticking to enriching only the people listed under "contacts" key from Osiris search.
        # This might need expansion if all roles need full details.

        # Example: if 'contacts' field in course_info_dict contains list of names
        contact_names_in_course = course_info_dict.get("contacts", [])
        if isinstance(contact_names_in_course, list):
            for contact_name in contact_names_in_course:
                contact_name_str = str(contact_name).strip()
                person_detail = person_details_map.get(contact_name_str)
                if person_detail:
                    enriched_contacts[contact_name_str] = {
                        "name": person_detail.get("main_name"),
                        "first_name": person_detail.get("first_name"), # Assuming first_name is directly available
                        "email": person_detail.get("email"),
                        "faculty": person_detail.get("faculty"), # This is faculty abbr
                        "orgs": person_detail.get("orgs"),
                        "programmes": person_detail.get("programmes"), # From people page scraping
                        "people_page": person_detail.get("people_page_url"),
                    }
                else:
                    logger.debug(f"No detailed person data found for contact '{contact_name_str}' in course '{course_code_str}'.")

        course_info_dict["contacts_detailed"] = enriched_contacts # Add new key for enriched data
        final_osiris_data_w_contacts[course_code_str] = course_info_dict

    # Merge with existing osiris_data_w_contacts if only_retrieve_missing
    if only_retrieve_missing:
        final_osiris_data_w_contacts.update(osiris_data_w_contacts_dict) # Add old entries not re-fetched

    if final_osiris_data_w_contacts:
        try:
            with open(osiris_data_w_contacts_path, "w", encoding="utf-8") as f_contacts_out:
                json.dump(final_osiris_data_w_contacts, f_contacts_out, indent=4, ensure_ascii=False)
            logger.info(f"Osiris data with enriched contacts saved to {osiris_data_w_contacts_path}")
        except OSError as e:
            logger.error(f"Failed to save Osiris data with contacts: {e}")

    logger.info(
        f"Osiris data enrichment process finished. Files updated:\n"
        f"  Course Data: {SETTINGS.files[FileSetting.OSIRIS_DATA].path}\n"
        f"  Person Data: {SETTINGS.files[FileSetting.PERSON_DATA].path}\n"
        f"  Courses with Contacts: {SETTINGS.files[FileSetting.OSIRIS_DATA_W_CONTACTS].path}"
    )

    # Final step: ensure base data (like new Persons, Courses from enrichment) is in DB
    # This load_base_data call might be redundant if all individual load_* functions were called
    # or if the data is only used for JSON files and not directly DB ingested here.
    # The original code had this, so keeping it for now.
    logger.info("Running load_base_data to ensure all new entities from enrichment are in DB.")
    await load_base_data() # This handles its own Tortoise init/close.
```
