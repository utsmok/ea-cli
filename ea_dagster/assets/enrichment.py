"""
Enrichment assets wrapping existing OSIRIS enrichment logic.

This module provides a Dagster asset that wraps the existing
`easy_access.enrichment.osiris.enrich_async` function.
"""

from dagster import AssetExecutionContext, AssetIn, Output, asset

from ea_dagster.resources import TortoiseDBResource
from easy_access.settings import SETTINGS


@asset(
    group_name="enrichment",
    ins={"processed_items": AssetIn(key="processed_copyright_items")},
)
async def osiris_enrichment(
    context: AssetExecutionContext,
    db: TortoiseDBResource,
    processed_items: str,
) -> Output:
    """
    Enriches copyright items with OSIRIS course and person data.

    This asset wraps the existing enrichment logic in
    `easy_access.enrichment.osiris.enrich_async`.

    Args:
        context: Dagster execution context
        db: Tortoise database resource
        processed_items: Output from processed_copyright_items asset

    Returns:
        str: Summary of enrichment results
    """
    async with db.yield_for_execution(context):
        from easy_access.enrichment.osiris import enrich_async

        context.log.info("Starting OSIRIS enrichment...")

        try:
            await enrich_async(SETTINGS)
            context.log.info("OSIRIS enrichment completed successfully")
            return Output(
                value="OSIRIS enrichment completed",
                metadata={"enrichment_status": "completed"},
            )
        except Exception as e:
            context.log.error(f"OSIRIS enrichment failed: {e}")
            return Output(
                value=f"OSIRIS enrichment failed: {e}",
                metadata={"enrichment_status": "failed"},
            )


@asset(
    group_name="enrichment",
    ins={"processed_items": AssetIn(key="processed_copyright_items")},
)
async def file_existence_check(
    context: AssetExecutionContext,
    db: TortoiseDBResource,
    processed_items: str,
) -> Output:
    """
    Checks file existence for copyright items based on TTL policies.

    This asset wraps the existing file existence logic in
    `easy_access.maintenance.file_existence.refresh_file_existence_async`.

    Args:
        context: Dagster execution context
        db: Tortoise database resource
        processed_items: Output from processed_copyright_items asset

    Returns:
        Output: Summary of file existence check results
    """
    async with db.yield_for_execution(context):
        from easy_access.maintenance.file_existence import refresh_file_existence_async

        context.log.info("Starting file existence check...")

        # Get settings from SETTINGS
        ttl_days = getattr(SETTINGS, "file_exists_ttl_days", 7)
        rate_limit_delay = getattr(SETTINGS, "file_exists_rate_limit_delay", 0.05)

        try:
            result = await refresh_file_existence_async(
                SETTINGS,
                ttl_days=ttl_days,
                rate_limit_delay=rate_limit_delay,
            )

            if "error" in result:
                context.log.error(f"File existence check failed: {result['error']}")
                return Output(
                    value=f"File existence check failed: {result['error']}",
                    metadata={"status": "failed", **result},
                )
            else:
                checked = result.get("checked", 0)
                exists = result.get("exists", 0)
                not_exists = result.get("not_exists", 0)
                value = (
                    f"File existence check completed: "
                    f"{checked} checked, {exists} exist, {not_exists} not found"
                )
                context.log.info(value)
                return Output(value=value, metadata=result)

        except Exception as e:
            context.log.error(f"File existence check failed: {e}")
            return Output(
                value=f"File existence check failed: {e}",
                metadata={"status": "failed"},
            )
