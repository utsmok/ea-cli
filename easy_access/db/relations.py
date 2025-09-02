"""
Database relations management module.

This module handles updating relationships between copyright items and other entities:
- Course linking based on course codes
- Duplicate status detection and management

Optimized to reduce N+1 query patterns using batch operations.
"""

from typing import Dict, List, Set

import polars as pl
from loguru import logger
from tortoise.transactions import in_transaction

from easy_access.db.base import ensure_db_inited, close_connections
from easy_access.db.models import CopyrightItem, Course, PDF
from easy_access.settings import Settings
from easy_access.utils import determine_course_code, safe_int


async def update_duplicates(settings: Settings) -> None:
    """
    Update duplicate status for all copyright items.

    Uses batch operations to minimize database queries:
    1. Fetch all PDFs with replace_with_id in one query
    2. Build replacement mapping in memory
    3. Bulk update items that have duplicates

    Args:
        settings: Application settings
    """
    logger.info("Updating duplicate statuses...")

    # Fetch all PDFs that have replacements in one query
    pdfs_with_replacements = await PDF.filter(replace_with_id__not_isnull=True).prefetch_related("replace_with")

    if not pdfs_with_replacements:
        logger.info("No PDFs with replacements found")
        return

    # Build mapping of material_id -> replacement_material_id
    replacement_map: Dict[int, int] = {}
    for pdf in pdfs_with_replacements:
        if pdf.replace_with:
            replacement_map[pdf.material_id] = pdf.replace_with.material_id

    if not replacement_map:
        logger.info("No valid replacements found")
        return

    # Get all items that might be duplicates
    material_ids = list(replacement_map.keys())
    items_to_update = await CopyrightItem.filter(material_id__in=material_ids)

    # Update items in memory
    updated_count = 0
    for item in items_to_update:
        if item.material_id in replacement_map:
            item.is_duplicate = True
            item.replacement_id = replacement_map[item.material_id]
            updated_count += 1

        # Bulk update only the changed items
        if updated_count > 0:
            async with in_transaction():
                for item in items_to_update:
                    if item.is_duplicate:
                        await item.save(update_fields=["is_duplicate", "replacement_id"])

        logger.success(f"Updated {updated_count} duplicate statuses")
    else:
        logger.info("No items needed duplicate status updates")


async def link_courses(settings: Settings) -> None:
    """
    Link copyright items to courses based on course codes.

    Optimized to reduce N+1 queries:
    1. Query items missing course links
    2. Extract course codes using determine_course_code
    3. Batch fetch all relevant Course objects
    4. Build relationships in memory
    5. Bulk create CourseItem links

    Args:
        settings: Application settings
    """
    logger.info("Linking courses to copyright items...")

    # Query items that don't have course links yet
    # For ManyToMany fields, we need to use a different approach than isnull
    all_items = await CopyrightItem.all().prefetch_related("courses")

    # Filter items that have no courses
    items_without_courses = []
    for item in all_items:
        if not item.courses:
            items_without_courses.append(item)

    if not items_without_courses:
        logger.info("All items already have course links")
        return

    logger.info(f"Found {len(items_without_courses)} items without course links")

    # Extract all potential course codes
    all_course_codes: Set[str] = set()
    item_course_map: Dict[int, List[str]] = {}

    for item in items_without_courses:
        course_codes = determine_course_code(item.course_code or "", item.course_name or "")
        if course_codes:
            course_codes_list = list(course_codes)
            item_course_map[item.material_id] = course_codes_list
            all_course_codes.update(course_codes_list)

    if not all_course_codes:
        logger.warning("No course codes found to link")
        return

    # Convert course codes to integers and filter valid ones
    valid_course_codes: Set[int] = set()
    for code in all_course_codes:
        if code:
            int_code = safe_int(code)
            if int_code is not None:
                valid_course_codes.add(int_code)

    if not valid_course_codes:
        logger.warning("No valid course codes found")
        return

    # Batch fetch all relevant courses
    courses = await Course.filter(cursuscode__in=valid_course_codes)
    course_map: Dict[int, Course] = {course.cursuscode: course for course in courses}

    logger.info(f"Fetched {len(courses)} courses for {len(valid_course_codes)} course codes")

    # Build relationships
    links_to_create = []
    links_added = 0

    for item in items_without_courses:
        item_id = item.material_id
        if item_id not in item_course_map:
            continue

        for course_code in item_course_map[item_id]:
            int_code = safe_int(course_code)
            if int_code and int_code in course_map:
                course = course_map[int_code]
                # Check if link already exists (shouldn't but safety check)
                if course not in await item.courses:
                    links_to_create.append((item, course))
                    links_added += 1

    # Bulk create relationships
    if links_to_create:
        async with in_transaction():
            for item, course in links_to_create:
                await item.courses.add(course)

        logger.success(f"Added {links_added} course links")
    else:
        logger.info("No new course links to create")


async def update_relations_async(settings: Settings) -> None:
    """
    Orchestrate all relations updates.

    Runs duplicate detection and course linking in sequence.

    Args:
        settings: Application settings
    """
    logger.info("Starting relations update...")

    # Ensure database is initialized
    await ensure_db_inited(settings)

    await update_duplicates(settings)
    await link_courses(settings)

    logger.info("Relations update completed")

    # Close database connections
    await close_connections()
