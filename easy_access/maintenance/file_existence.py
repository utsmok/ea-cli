"""
File existence verification with TTL-based freshness policies.

This module provides incremental file existence checking that:
- Uses TTL policies based on last_canvas_check timestamps
- Leverages existing file existence utilities
- Updates database records with fresh status
- Integrates with the maintenance pipeline
"""

import asyncio
import time
from datetime import datetime, timedelta
from typing import List, Dict, Any

import httpx
from loguru import logger

from easy_access.db.base import ensure_db_inited, close_connections
from easy_access.db.models import CopyrightItem
from easy_access.settings import Settings
from easy_access.utilities.file_exists import check_file_exists


async def select_items_needing_file_check(
    settings: Settings,
    ttl_days: int | None = None,
    batch_size: int = 1000,
    force: bool = False
) -> List[Dict[str, Any]]:
    """
    Select copyright items that need file existence verification.

    Args:
        settings: Application settings
        ttl_days: TTL in days (None means check all unchecked items)
        batch_size: Maximum number of items to return
        force: If True, check all items regardless of TTL

    Returns:
        List of item dictionaries with material_id and url
    """
    logger.info("Selecting items that need file existence verification...")

    # Build query conditions
    conditions = []

    if not force:
        if ttl_days is not None:
            # Check items that are either unchecked or older than TTL
            cutoff_date = datetime.now() - timedelta(days=ttl_days)
            conditions.append(
                f"(file_exists IS NULL OR last_canvas_check < '{cutoff_date.isoformat()}')"
            )
        else:
            # Check only unchecked items
            conditions.append("file_exists IS NULL")

    # Build the query
    where_clause = " AND ".join(conditions) if conditions else "1=1"

    # Query items needing check
    items = await CopyrightItem.raw(
        f"""
        SELECT material_id, url
        FROM copyright_data
        WHERE {where_clause} AND url IS NOT NULL AND url != ''
        ORDER BY last_canvas_check ASC NULLS FIRST
        LIMIT {batch_size}
        """
    )

    result = []
    for item in items:
        # Access attributes from raw query result
        material_id = getattr(item, 'material_id', None)
        url = getattr(item, 'url', None)
        if material_id and url:
            result.append({
                "material_id": material_id,
                "url": url,
            })

    logger.info(f"Selected {len(result)} items for file existence verification")
    return result


async def check_single_file_existence(
    item_data: Dict[str, Any],
    session: httpx.AsyncClient
) -> Dict[str, Any]:
    """
    Check file existence for a single item.

    Args:
        item_data: Dictionary with material_id and url
        session: HTTP client session

    Returns:
        Dictionary with material_id, file_exists, and last_canvas_check
    """
    material_id = item_data["material_id"]
    url = item_data["url"]

    try:
        # Extract file_id from URL
        if "/files/" in url:
            file_id = url.split("/files/")[1].split("/")[0].split("?")[0]
            file_url = f"https://utwente.instructure.com/api/v1/files/{file_id}"

            # Check file existence
            response = await session.get(file_url)
            file_exists = response.status_code == 200

            return {
                "material_id": material_id,
                "file_exists": file_exists,
                "last_canvas_check": datetime.now(),
            }
        else:
            logger.warning(f"Invalid URL format for material_id {material_id}: {url}")
            return {
                "material_id": material_id,
                "file_exists": False,
                "last_canvas_check": datetime.now(),
            }

    except Exception as e:
        logger.error(f"Error checking file existence for material_id {material_id}: {e}")
        return {
            "material_id": material_id,
            "file_exists": False,
            "last_canvas_check": datetime.now(),
        }


async def update_file_existence_batch(
    results: List[Dict[str, Any]]
) -> None:
    """
    Update file existence status for a batch of items using bulk operations.

    Args:
        results: List of result dictionaries with material_id, file_exists, last_canvas_check
    """
    if not results:
        return

    logger.info(f"Updating file existence for {len(results)} items using bulk operations")

    # Prepare data for bulk update
    material_ids = []
    file_exists_values = []
    last_check_values = []

    for result in results:
        material_ids.append(result["material_id"])
        file_exists_values.append(result["file_exists"])
        last_check_values.append(result["last_canvas_check"].isoformat())

    # Use raw SQL for bulk update to avoid N+1 queries
    # Create temporary table for bulk update
    temp_table_name = f"temp_file_existence_{int(time.time())}"

    try:
        # Create temporary table
        await CopyrightItem.raw(f"""
            CREATE TEMP TABLE {temp_table_name} (
                material_id INTEGER PRIMARY KEY,
                file_exists BOOLEAN,
                last_canvas_check TIMESTAMP
            )
        """)

        # Bulk insert into temporary table
        values_list = []
        for i, result in enumerate(results):
            values_list.append(f"({result['material_id']}, {result['file_exists']}, '{result['last_canvas_check'].isoformat()}')")

        if values_list:
            values_str = ", ".join(values_list)
            await CopyrightItem.raw(f"""
                INSERT INTO {temp_table_name} (material_id, file_exists, last_canvas_check)
                VALUES {values_str}
            """)

            # Bulk update from temporary table
            await CopyrightItem.raw(f"""
                UPDATE copyright_data
                SET file_exists = t.file_exists,
                    last_canvas_check = t.last_canvas_check
                FROM {temp_table_name} t
                WHERE copyright_data.material_id = t.material_id
            """)

        logger.info(f"Successfully bulk updated {len(results)} items")

    except Exception as e:
        logger.error(f"Error in bulk update: {e}")
        # Fallback to individual updates
        logger.info("Falling back to individual updates")
        for result in results:
            material_id = result["material_id"]
            file_exists = result["file_exists"]
            last_canvas_check = result["last_canvas_check"]

            await CopyrightItem.filter(material_id=material_id).update(
                file_exists=file_exists,
                last_canvas_check=last_canvas_check
            )
        logger.info(f"Successfully updated {len(results)} items using fallback method")

    finally:
        # Clean up temporary table
        try:
            await CopyrightItem.raw(f"DROP TABLE IF EXISTS {temp_table_name}")
        except Exception as e:
            logger.warning(f"Could not drop temporary table {temp_table_name}: {e}")


async def refresh_file_existence_async(
    settings: Settings,
    ttl_days: int | None = None,
    batch_size: int = 1000,
    max_concurrent: int = 50,
    force: bool = False,
    rate_limit_delay: float = 0.1  # Add configurable rate limiting
) -> Dict[str, Any]:
    """
    Refresh file existence status for copyright items based on TTL policy.

    Args:
        settings: Application settings
        ttl_days: TTL in days for file existence checks
        batch_size: Number of items to process in each batch
        max_concurrent: Maximum concurrent requests
        force: If True, check all items regardless of TTL
        rate_limit_delay: Delay in seconds between requests to avoid rate limiting

    Returns:
        Dictionary with statistics about the operation
    """
    logger.info("Starting file existence verification...")

    # Ensure database is initialized
    await ensure_db_inited(settings)

    try:
        # Get API token from settings
        api_token = getattr(settings.university_settings, 'canvas_api_token', None)
        if not api_token:
            logger.error("Canvas API token not found in settings")
            return {"error": "No API token", "checked": 0, "exists": 0, "not_exists": 0}

        # Select items needing verification
        items_to_check = await select_items_needing_file_check(
            settings, ttl_days, batch_size, force
        )

        if not items_to_check:
            logger.info("No items need file existence verification")
            return {"checked": 0, "updated": 0}

        # Set up HTTP client
        headers = {"Authorization": f"Bearer {api_token}"}
        async with httpx.AsyncClient(
            headers=headers,
            follow_redirects=True,
            timeout=20.0
        ) as session:

            logger.info(f"Checking file existence for {len(items_to_check)} items")

            # Check files concurrently with rate limiting
            semaphore = asyncio.Semaphore(max_concurrent)
            results = []

            async def check_with_semaphore_and_rate_limit(item_data: Dict[str, Any]) -> None:
                async with semaphore:
                    result = await check_single_file_existence(item_data, session)
                    results.append(result)
                    # Add rate limiting delay between requests
                    if rate_limit_delay > 0:
                        await asyncio.sleep(rate_limit_delay)

            # Create tasks
            tasks = [check_with_semaphore_and_rate_limit(item) for item in items_to_check]

            # Execute concurrently
            start_time = time.time()
            await asyncio.gather(*tasks, return_exceptions=True)
            duration = time.time() - start_time

            logger.info(
                f"Completed file existence checks in {duration:.2f} seconds "
                f"({len(results)/max(duration, 0.001):.1f} req/sec)"
            )

            # Update database using optimized bulk operations
            await update_file_existence_batch(results)

            # Calculate statistics
            exists_count = sum(1 for r in results if r["file_exists"])
            not_exists_count = sum(1 for r in results if not r["file_exists"])

            logger.info(
                f"File existence verification complete: "
                f"{exists_count} exist, {not_exists_count} not found"
            )

            return {
                "checked": len(results),
                "exists": exists_count,
                "not_exists": not_exists_count,
                "duration_seconds": int(duration),
            }

    finally:
        # Close database connections
        await close_connections()
