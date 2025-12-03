"""
Enrichment assets wrapping existing OSIRIS enrichment logic.

This module provides a Dagster asset that wraps the existing
`easy_access.enrichment.osiris.enrich_async` function.
"""

from dagster import AssetExecutionContext, AssetIn, asset

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
) -> str:
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
            return "OSIRIS enrichment completed"
        except Exception as e:
            context.log.error(f"OSIRIS enrichment failed: {e}")
            return f"OSIRIS enrichment failed: {e}"
