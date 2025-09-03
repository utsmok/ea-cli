"""
Enrichment module for fetching and persisting OSIRIS course and person data.

This module provides DB-centric enrichment functionality that:
- Fetches missing or stale course/person data from OSIRIS
- Uses TTL-based freshness policies
- Stores data directly in the database
- Integrates with the pipeline for automated enrichment
"""

from typing import Dict, List, Set
import asyncio
import httpx
import bs4
from bs4 import Tag
import Levenshtein
from loguru import logger

from easy_access.db.base import ensure_db_inited, close_connections
from easy_access.db.models import CopyrightItem, Course, Person
from easy_access.settings import Settings
from easy_access.utils import determine_course_code, safe_int


async def gather_target_course_codes(settings: Settings) -> Set[int]:
    """
    Gather all unique course codes from copyright items that need enrichment.

    Returns:
        Set of course codes (integers) that exist in the database
    """
    logger.info("Gathering target course codes for enrichment...")

    # Query all unique course codes from copyright items
    items = await CopyrightItem.all().distinct()
    all_course_codes: Set[str] = set()

    for item in items:
        course_codes = determine_course_code(item.course_code or "", item.course_name or "")
        if course_codes:
            all_course_codes.update(course_codes)

    # Convert to integers and filter valid ones
    valid_course_codes: Set[int] = set()
    for code in all_course_codes:
        if code:
            int_code = safe_int(code)
            if int_code is not None:
                valid_course_codes.add(int_code)

    logger.info(f"Found {len(valid_course_codes)} unique course codes")
    return valid_course_codes


async def select_missing_or_stale_courses(
    settings: Settings,
    course_codes: Set[int],
    ttl_days: int | None = None
) -> Set[int]:
    """
    Select course codes that are missing or stale based on TTL policy.

    Args:
        settings: Application settings
        course_codes: Set of course codes to check
        ttl_days: TTL in days (None means no TTL check, only missing courses)

    Returns:
        Set of course codes that need fetching
    """
    logger.info("Selecting courses that need enrichment...")

    # Get existing courses
    existing_courses = await Course.filter(cursuscode__in=course_codes)
    existing_codes = {course.cursuscode for course in existing_courses}

    # Find missing courses
    missing_codes = course_codes - existing_codes
    logger.info(f"Found {len(missing_codes)} missing courses")

    # If no TTL specified, return only missing courses
    if ttl_days is None:
        return missing_codes

    # Check for stale courses based on TTL
    stale_codes: Set[int] = set()
    for course in existing_courses:
        if course.modified_at is None:
            # No modification date, consider stale
            stale_codes.add(course.cursuscode)
        else:
            # Check if older than TTL
            from datetime import datetime
            age_days = (datetime.now() - course.modified_at).days
            if age_days > ttl_days:
                stale_codes.add(course.cursuscode)

    logger.info(f"Found {len(stale_codes)} stale courses (TTL: {ttl_days} days)")
    return missing_codes | stale_codes


async def gather_target_person_names(settings: Settings) -> Set[str]:
    """
    Gather all unique person names from copyright items that need enrichment.

    Returns:
        Set of person names (strings) that exist in the database
    """
    logger.info("Gathering target person names for enrichment...")

    # Query all unique person names from copyright items
    items = await CopyrightItem.all().distinct()
    all_person_names: Set[str] = set()

    for item in items:
        # Collect names from various fields
        if item.author:
            all_person_names.add(item.author)
        if item.auditor:
            all_person_names.add(item.auditor)

    # Filter out empty names
    valid_person_names = {name for name in all_person_names if name and name.strip()}

    logger.info(f"Found {len(valid_person_names)} unique person names")
    return valid_person_names


async def select_missing_or_stale_persons(
    settings: Settings,
    person_names: Set[str],
    ttl_days: int | None = None
) -> Set[str]:
    """
    Select person names that are missing or stale based on TTL policy.

    Args:
        settings: Application settings
        person_names: Set of person names to check
        ttl_days: TTL in days (None means no TTL check, only missing persons)

    Returns:
        Set of person names that need fetching
    """
    logger.info("Selecting persons that need enrichment...")

    # Get existing persons
    existing_persons = await Person.filter(input_name__in=person_names)
    existing_names = {person.input_name for person in existing_persons}

    # Find missing persons
    missing_names = person_names - existing_names
    logger.info(f"Found {len(missing_names)} missing persons")

    # If no TTL specified, return only missing persons
    if ttl_days is None:
        return missing_names

    # Check for stale persons based on TTL
    stale_names: Set[str] = set()
    for person in existing_persons:
        if person.modified_at is None:
            # No modification date, consider stale
            stale_names.add(person.input_name)
        else:
            # Check if older than TTL
            from datetime import datetime
            age_days = (datetime.now() - person.modified_at).days
            if age_days > ttl_days:
                stale_names.add(person.input_name)

    logger.info(f"Found {len(stale_names)} stale persons (TTL: {ttl_days} days)")
    return missing_names | stale_names


async def fetch_and_parse_courses(
    settings: Settings,
    course_codes: Set[int],
    max_concurrent: int = 10
) -> Dict[int, Dict]:
    """
    Fetch and parse course data concurrently for multiple course codes.

    Args:
        settings: Application settings
        course_codes: Set of course codes to fetch
        max_concurrent: Maximum number of concurrent requests

    Returns:
        Dictionary mapping course codes to parsed course data
    """
    logger.info(f"Fetching {len(course_codes)} courses concurrently (max {max_concurrent} at a time)")

    # Create semaphore to limit concurrent requests
    semaphore = asyncio.Semaphore(max_concurrent)
    results = {}

    async def fetch_single_course(course_code: int) -> None:
        async with semaphore:
            try:
                async with httpx.AsyncClient(timeout=30.0) as client:
                    course_data = await fetch_course_data(course_code, client)
                    if course_data:
                        results[course_code] = course_data
                        logger.debug(f"Successfully fetched course {course_code}")
                    else:
                        logger.warning(f"No data found for course {course_code}")
            except Exception as e:
                logger.error(f"Error fetching course {course_code}: {e}")

    # Create tasks for all course codes
    tasks = [fetch_single_course(code) for code in course_codes]

    # Execute all tasks concurrently
    await asyncio.gather(*tasks, return_exceptions=True)

    logger.info(f"Completed fetching {len(results)}/{len(course_codes)} courses")
    return results


async def fetch_and_parse_persons(
    settings: Settings,
    person_names: Set[str],
    max_concurrent: int = 5
) -> Dict[str, Dict]:
    """
    Fetch and parse person data concurrently for multiple person names.

    Args:
        settings: Application settings
        person_names: Set of person names to fetch
        max_concurrent: Maximum number of concurrent requests

    Returns:
        Dictionary mapping person names to parsed person data
    """
    logger.info(f"Fetching {len(person_names)} persons concurrently (max {max_concurrent} at a time)")

    # Create semaphore to limit concurrent requests
    semaphore = asyncio.Semaphore(max_concurrent)
    results = {}

    async def fetch_single_person(person_name: str) -> None:
        async with semaphore:
            try:
                async with httpx.AsyncClient(timeout=30.0) as client:
                    person_data = await fetch_person_data(person_name, client)
                    if person_data:
                        results[person_name] = person_data
                        logger.debug(f"Successfully fetched person {person_name}")
                    else:
                        logger.warning(f"No data found for person {person_name}")
            except Exception as e:
                logger.error(f"Error fetching person {person_name}: {e}")

    # Create tasks for all person names
    tasks = [fetch_single_person(name) for name in person_names]

    # Execute all tasks concurrently
    await asyncio.gather(*tasks, return_exceptions=True)

    logger.info(f"Completed fetching {len(results)}/{len(person_names)} persons")
    return results


async def persist_courses(courses_data: Dict[int, Dict]) -> None:
    """
    Persist course data to the database with bulk upsert operations.

    Args:
        courses_data: Dictionary mapping course codes to course data
    """
    logger.info(f"Persisting {len(courses_data)} courses to database...")

    if not courses_data:
        logger.info("No course data to persist")
        return

    # Prepare bulk operations
    courses_to_create = []
    courses_to_update = []

    # Check existing courses
    existing_codes = set()
    existing_courses = await Course.filter(cursuscode__in=list(courses_data.keys()))
    for course in existing_courses:
        existing_codes.add(course.cursuscode)

    # Prepare data for bulk operations
    for course_code, course_data in courses_data.items():
        if course_code in existing_codes:
            courses_to_update.append(course_data)
        else:
            courses_to_create.append(course_data)

    # Bulk create new courses
    if courses_to_create:
        logger.info(f"Creating {len(courses_to_create)} new courses")
        for course_data in courses_to_create:
            try:
                await Course.create(**course_data)
            except Exception as e:
                logger.error(f"Error creating course {course_data.get('cursuscode')}: {e}")

    # Bulk update existing courses
    if courses_to_update:
        logger.info(f"Updating {len(courses_to_update)} existing courses")
        for course_data in courses_to_update:
            try:
                course_code = course_data.get('cursuscode')
                if course_code:
                    await Course.filter(cursuscode=course_code).update(**course_data)
            except Exception as e:
                logger.error(f"Error updating course {course_data.get('cursuscode')}: {e}")

    logger.info(f"Successfully persisted {len(courses_data)} courses")


async def persist_persons(persons_data: Dict[str, Dict]) -> None:
    """
    Persist person data to the database with bulk upsert operations.

    Args:
        persons_data: Dictionary mapping person names to person data
    """
    logger.info(f"Persisting {len(persons_data)} persons to database...")

    if not persons_data:
        logger.info("No person data to persist")
        return

    # Prepare bulk operations
    persons_to_create = []
    persons_to_update = []

    # Check existing persons
    existing_names = set()
    existing_persons = await Person.filter(input_name__in=list(persons_data.keys()))
    for person in existing_persons:
        existing_names.add(person.input_name)

    # Prepare data for bulk operations
    for person_name, person_data in persons_data.items():
        if person_name in existing_names:
            persons_to_update.append(person_data)
        else:
            persons_to_create.append(person_data)

    # Bulk create new persons
    if persons_to_create:
        logger.info(f"Creating {len(persons_to_create)} new persons")
        for person_data in persons_to_create:
            try:
                await Person.create(**person_data)
            except Exception as e:
                logger.error(f"Error creating person {person_data.get('input_name')}: {e}")

    # Bulk update existing persons
    if persons_to_update:
        logger.info(f"Updating {len(persons_to_update)} existing persons")
        for person_data in persons_to_update:
            try:
                input_name = person_data.get('input_name')
                if input_name:
                    await Person.filter(input_name=input_name).update(**person_data)
            except Exception as e:
                logger.error(f"Error updating person {person_data.get('input_name')}: {e}")

    logger.info(f"Successfully persisted {len(persons_data)} persons")


async def enrich_async(settings: Settings) -> None:
    """
    Main enrichment orchestrator that fetches and persists missing/stale OSIRIS data.

    This is the pipeline stage that coordinates:
    - Course data fetching and persistence
    - Person data fetching and persistence
    - Course-person relationship linking

    Args:
        settings: Application settings
    """
    logger.info("Starting OSIRIS enrichment...")

    # Ensure database is initialized
    await ensure_db_inited(settings)

    try:
        # Gather target course codes
        course_codes = await gather_target_course_codes(settings)

        if not course_codes:
            logger.info("No course codes found for enrichment")
            return

        # Get TTL settings (default to None if not configured)
        course_ttl = getattr(settings.enrichment_settings, 'course_ttl_days', None)
        person_ttl = getattr(settings.enrichment_settings, 'person_ttl_days', None)

        # Select courses that need fetching
        courses_to_fetch = await select_missing_or_stale_courses(
            settings, course_codes, course_ttl
        )

        if not courses_to_fetch:
            logger.info("All courses are fresh, skipping enrichment")
            return

        logger.info(f"Will fetch {len(courses_to_fetch)} courses")

        # Fetch course data concurrently
        courses_data = await fetch_and_parse_courses(settings, courses_to_fetch)

        if not courses_data:
            logger.info("No course data retrieved")
            return

        # Extract person names from course data
        person_names = set()
        for course_data in courses_data.values():
            # Add teachers, contacts, etc. from course data
            for field in ['teachers', 'contacts', 'docenten', 'examinators', 'tutors']:
                if field in course_data and course_data[field]:
                    if isinstance(course_data[field], list):
                        person_names.update(course_data[field])
                    elif isinstance(course_data[field], set):
                        person_names.update(course_data[field])

        # Filter out empty names
        person_names = {name for name in person_names if name and name.strip()}

        if not person_names:
            logger.info("No person names found in course data")
        else:
            # Select persons that need fetching
            persons_to_fetch = await select_missing_or_stale_persons(
                settings, person_names, person_ttl
            )

            if persons_to_fetch:
                logger.info(f"Will fetch {len(persons_to_fetch)} persons")

                # Fetch person data concurrently
                persons_data = await fetch_and_parse_persons(settings, persons_to_fetch)

                # Persist person data
                await persist_persons(persons_data)
            else:
                logger.info("All persons are fresh, skipping person fetching")

        # Persist course data
        await persist_courses(courses_data)

        logger.info("Enrichment completed successfully")

    finally:
        # Close database connections
        await close_connections()


def _process_teacher_items(items) -> set[str]:
    """Helper function to process teacher items into a consistent set format"""
    if isinstance(items, list):
        return set(items)
    elif isinstance(items, str):
        return {items}
    elif isinstance(items, set):
        return items
    return set()


def _extract_languages(voertalen_data) -> list[str]:
    """Extract language information from voertalen data"""
    if isinstance(voertalen_data, list):
        return [x.get("voertaal_omschrijving") for x in voertalen_data if x.get("voertaal_omschrijving")]
    return []


async def _fetch_course_details(course_data: dict, httpx_client: httpx.AsyncClient) -> None:
    """Fetch detailed course information including contacts from OSIRIS"""
    internal_id = course_data.get("internal_id")
    if not internal_id:
        return

    url = f"https://utwente.osiris-student.nl/student/osiris/owc/cursussen/{internal_id}"
    headers = {
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
        "sec-fetch-user": "?1",
        "taal": "NL",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    }

    try:
        response = await httpx_client.get(url, headers=headers)
        if response.status_code == 200:
            course_details = response.json()
            for datapoint in course_details.get("items", []):
                if datapoint.get("rubriek") == "rubriek-docenten":
                    docent_data = datapoint.get("velden", [])
                    if docent_data:
                        for docent_item in docent_data:
                            if docent_item.get("waarde"):
                                for docent_type in docent_item.get("waarde", []):
                                    for persoon in docent_type.get("velden", []):
                                        person_name = persoon.get("docent")
                                        if person_name:
                                            role_type = docent_type.get("omschrijving")
                                            if role_type == "Contactpersoon":
                                                course_data["contacts"].add(person_name)
                                            elif role_type == "Docent":
                                                course_data["docenten"].add(person_name)
                                            elif role_type == "Examinator":
                                                course_data["examinators"].add(person_name)
                                            elif role_type == "Tutor":
                                                course_data["tutors"].add(person_name)
                                            else:
                                                course_data["unknown_role"].add(person_name)

            # Convert sets to lists for JSON serialization
            for field in ["teachers", "contacts", "docenten", "examinators", "tutors", "unknown_role"]:
                if isinstance(course_data.get(field), set):
                    course_data[field] = list(course_data[field])
                    # Filter out single-character entries (likely parsing errors)
                    if len(course_data[field]) > 8 and all(len(x) == 1 for x in course_data[field]):
                        course_data[field] = []

        else:
            logger.error(f"Error retrieving course details: HTTP {response.status_code}")

    except Exception as e:
        logger.error(f"Error fetching course details: {e}")


async def fetch_course_data(course_code: int, httpx_client: httpx.AsyncClient) -> Dict:
    """
    Fetch course data from OSIRIS API.

    Args:
        course_code: Course code to fetch
        httpx_client: HTTP client for making requests

    Returns:
        Dictionary containing course data
    """
    # OSIRIS API endpoint and headers based on existing implementation
    url = "https://utwente.osiris-student.nl/student/osiris/student/cursussen/zoeken"

    # Build the search query for the course code
    startstring = '{"from":0,"size":25,"sort":[{"cursus_lange_naam.raw":{"order":"asc"}},{"cursus":{"order":"asc"}},{"collegejaar":{"order":"desc"}}],"aggs":{"agg_terms_collegejaar":{"filter":{"bool":{"must":[]}},"aggs":{"agg_collegejaar_buckets":{"terms":{"field":"collegejaar","size":2500,"order":{"_term":"desc"}}}}},"agg_terms_blokken_nested.periode_omschrijving":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_blokken_nested.periode_omschrijving":{"terms":{"field":"blokken_nested.periode_omschrijving","size":2500,"order":{"_term":"asc"},"exclude":"Periode: [0-9][0-9]-[0-9][0-9]-[0-9][0-9][0-9][0-9]"}},"nested_aggs":{"nested":{"path":"blokken_nested"},"aggs":{"nested_aggs":{"filter":{"bool":{"must":[]}},"aggs":{"agg_blokken_nested.periode_omschrijving_buckets":{"terms":{"field":"blokken_nested.periode_omschrijving","size":2500,"order":{"_term":"asc"},"exclude":"Periode: [0-9][0-9]-[0-9][0-9]-[0-9][0-9][0-9][0-9]"},"aggs":{"items":{"reverse_nested":{}}}}}}}}}},"agg_terms_faculteit_naam":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_faculteit_naam_buckets":{"terms":{"field":"faculteit_naam","size":2500,"order":{"_term":"asc"}}}}},"agg_terms_coordinerend_onderdeel_oms":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_coordinerend_onderdeel_oms_buckets":{"terms":{"field":"coordinerend_onderdeel_oms","size":2500,"order":{"_term":"asc"}}}}},"agg_terms_categorie_omschrijving":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_categorie_omschrijving_buckets":{"terms":{"field":"categorie_omschrijving","size":2500,"order":{"_term":"asc"}}}}},"agg_terms_voertalen.voertaal_omschrijving":{"filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"aggs":{"agg_voertalen.voertaal_omschrijving_buckets":{"terms":{"field":"voertalen.voertaal_omschrijving","size":2500,"order":{"_term":"asc"}}}}}},"post_filter":{"bool":{"must":[{"terms":{"collegejaar":["2024-2025"]}}]}},"query":{"bool":{"must":[{"multi_match":{"query":'
    code = f'"{course_code}"'
    endstring = ',"type":"phrase_prefix","fields":["cursus","cursus_korte_naam","cursus_lange_naam"],"max_expansions":200}}]}}}'
    body = startstring + code + endstring

    headers = {
        "host": "utwente.osiris-student.nl",
        "connection": "keep-alive",
        "content-length": str(len(body)),
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
        response = await httpx_client.post(url=url, headers=headers, content=body)
        results = response.json().get("hits", {}).get("hits", [])

        if not results:
            logger.warning(f"No OSIRIS data found for course code {course_code}")
            return {}

        # Process the first result (most relevant)
        rawdata = results[0].get("_source", {})

        # Extract teacher information
        teachers = set()
        for key, value in rawdata.items():
            if key == "docenten" and value:
                teachers = _process_teacher_items(value)

        # Build course data structure
        course_data = {
            "cursuscode": course_code,
            "internal_id": rawdata.get("id_cursus"),
            "year": rawdata.get("collegejaar"),
            "short_name": rawdata.get("cursus_korte_naam"),
            "name": rawdata.get("cursus_lange_naam"),
            "faculty": rawdata.get("faculteit"),
            "faculty_long": rawdata.get("faculteit_naam"),
            "programme": rawdata.get("coordinerend_onderdeel_oms"),
            "ec": rawdata.get("punten"),
            "language": _extract_languages(rawdata.get("voertalen", [])),
            "notes": rawdata.get("opmerking_cursus"),
            "category": rawdata.get("categorie_omschrijving"),
            "teachers": teachers,
            "contacts": set(),
            "docenten": set(),
            "examinators": set(),
            "unknown_role": set(),
            "tutors": set(),
        }

        # Fetch detailed course information including contacts
        await _fetch_course_details(course_data, httpx_client)

        logger.info(f"Successfully fetched course data for {course_code}")
        return course_data

    except Exception as e:
        logger.error(f"Error fetching course data for {course_code}: {e}")
        return {}


def _strip_name(name: str) -> str:
    """Strip academic titles from person names for better matching"""
    stripped_name = name.strip()
    titles = [
        "ing.",
        "dr.",
        "prof.",
        "ir.",
        "rer.",
        "nat.",
        ", MSc",
        ", PhD",
        ", BSc",
    ]
    for title in titles:
        stripped_name = stripped_name.replace(title, "").strip()
    return stripped_name


def _remove_dot_and_lower(name: str) -> str:
    """Normalize name for comparison by removing dots and converting to lowercase"""
    return str(name).strip().replace(".", "").lower()


async def fetch_person_data(person_name: str, httpx_client: httpx.AsyncClient) -> Dict:
    """
    Fetch person data from people.utwente.nl.

    Args:
        person_name: Person name to search for
        httpx_client: HTTP client for making requests

    Returns:
        Dictionary containing person data
    """
    url = "https://people.utwente.nl/overview"
    headers = {
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

    try:
        # Search for the person
        search_url = f"https://people.utwente.nl/overview?query={person_name}"
        response = await httpx_client.get(search_url, headers=headers)

        if response.status_code != 200:
            logger.warning(f"Failed to search for person {person_name}: HTTP {response.status_code}")
            return {}

        # Parse search results
        soup = bs4.BeautifulSoup(response.text, "lxml")
        data_links = soup.find_all("a", {"data-link": True})

        if not data_links:
            logger.warning(f"No search results found for person {person_name}")
            return {}

        # Limit to first 5 results to avoid too many matches
        matches = []
        for link in data_links[:5]:
            if isinstance(link, Tag) and hasattr(link, 'get'):
                data_link = link.get("data-link")
                if isinstance(data_link, str):
                    matches.append(data_link)

        if not matches:
            logger.warning(f"No valid data links found for person {person_name}")
            return {}

        # Find best match using Levenshtein distance
        name_parsed_str = _strip_name(person_name)
        compare_name = _remove_dot_and_lower(name_parsed_str)

        best_match = matches[0]
        best_ratio = Levenshtein.ratio(compare_name, _remove_dot_and_lower(best_match))

        # Check other matches for better similarity
        for match in matches[1:]:
            if isinstance(match, str) and ("business" in match or "/" in match):
                continue
            if isinstance(match, str):
                ratio = Levenshtein.ratio(compare_name, _remove_dot_and_lower(match))
                if ratio > best_ratio:
                    best_match = match
                    best_ratio = ratio

        # Skip if match confidence is too low
        if best_ratio < 0.7:
            logger.warning(
                f"Low match confidence {best_ratio}: best match for {person_name} is {best_match}. "
                f"Actual comparison: found {_remove_dot_and_lower(best_match)} vs input {compare_name}"
            )
            return {}

        # Fetch detailed person page
        person_url = "https://people.utwente.nl/" + best_match
        response = await httpx_client.get(person_url, headers=headers)

        if response.status_code in [500, 502]:
            logger.warning(f"Server error for person {person_name}, retrying...")
            # Could implement retry logic here
            return {}

        if response.status_code != 200:
            logger.warning(f"Failed to fetch person page for {person_name}: HTTP {response.status_code}")
            return {}

        # Parse person page
        soup = bs4.BeautifulSoup(response.text, "lxml")

        # Extract main name
        name_tag = soup.find("h1", class_="pageheader__title")
        main_name = ""
        other_names = []

        if name_tag and isinstance(name_tag, Tag):
            for string_part in name_tag.strings:
                if not main_name:
                    main_name = str(string_part).strip()
                else:
                    other_names.append(str(string_part).strip().replace("(", "").replace(")", ""))

        if not main_name:
            logger.warning(f"No name found for person {person_name} at {person_url}")
            return {}

        # Calculate match confidence
        final_ratio = Levenshtein.ratio(name_parsed_str, _strip_name(main_name))
        if final_ratio < 0.7:
            logger.debug(
                f"Found name {main_name} differs from input name: {person_name} "
                f"with ratio {final_ratio}. Compared: found {_strip_name(main_name)} | input {_strip_name(person_name)}"
            )

        # Extract email
        email = ""
        for link_tag in soup.find_all("a"):
            if isinstance(link_tag, Tag):
                href = link_tag.get("href")
                if href and isinstance(href, str) and "mailto:" in href:
                    email = href.replace("mailto:", "")
                    break

        # Extract organization data
        orgs = []
        faculty = ""
        faculty_abbr = ""

        org_containers = soup.find_all(class_="widget-linklist--smallicons")
        if org_containers:
            first_container = org_containers[0]
            if isinstance(first_container, Tag):
                org_tags = first_container.find_all(class_="widget-linklist__text")

                for org_tag in org_tags:
                    if isinstance(org_tag, Tag):
                        text_content = org_tag.string
                        if text_content and isinstance(text_content, str) and "(" in text_content:
                            try:
                                org_name = text_content.split("(")[0].strip()
                                org_abbr = text_content.split("(")[1].split(")")[0].strip()

                                # Check if this is a faculty
                                if org_abbr in ["BMS", "ET", "EEMCS", "ITC", "TNW"]:
                                    faculty = org_name
                                    faculty_abbr = org_abbr
                                else:
                                    orgs.append({"name": org_name, "abbr": org_abbr})
                            except Exception as e:
                                logger.exception(f"Error processing org {text_content}: {e}")

        # Add faculty to orgs if found
        if faculty and faculty_abbr:
            orgs.insert(0, {"name": faculty, "abbr": faculty_abbr})

            # Link non-faculty orgs to faculty
            for org in orgs[1:]:
                if faculty_abbr in org.get("abbr", ""):
                    cleaned_abbr = org["abbr"].replace("-" + faculty_abbr, "")
                    org["abbr"] = cleaned_abbr

        # Extract education data (courses and programmes)
        courses = []
        programmes = []

        education_tab = soup.find("div", id="tabpanel-education")
        if education_tab and isinstance(education_tab, Tag):
            for link_tag in education_tab.find_all("a"):
                if isinstance(link_tag, Tag):
                    href = link_tag.get("href")
                    link_text = link_tag.string

                    if href and isinstance(href, str) and link_text and isinstance(link_text, str):
                        if "https://utwente.osiris-student.nl" in href:
                            # This is a course
                            link_text = str(link_text).strip()
                            if " - " in link_text:
                                code, course_name = link_text.split(" - ", 1)
                                courses.append({
                                    "course_code": code.strip(),
                                    "course_name": course_name.strip(),
                                })
                        elif "https://www.utwente.nl/" in href:
                            # This is a programme
                            programmes.append({
                                "name": str(link_text).strip(),
                                "url": href,
                            })

        # Build person data structure
        person_data = {
            "input_name": person_name,
            "main_name": main_name,
            "match_confidence": final_ratio,
            "other_names": other_names,
            "email": email,
            "orgs": orgs,
            "courses": courses,
            "programmes": programmes,
            "faculty": faculty_abbr,
            "people_page_url": person_url,
        }

        logger.info(f"Successfully fetched person data for {person_name}")
        return person_data

    except Exception as e:
        logger.error(f"Error fetching person data for {person_name}: {e}")
        return {}
