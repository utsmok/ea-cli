"""
Export assets wrapping existing Excel report generation logic.

This module provides a Dagster asset that wraps the existing
export functionality in `easy_access.sheets.export`.
"""

from dagster import AssetExecutionContext, AssetIn, asset

from ea_dagster.resources import TortoiseDBResource
from easy_access.settings import SETTINGS


@asset(
    group_name="export",
    ins={"enrichment": AssetIn(key="osiris_enrichment")},
)
async def excel_reports(
    context: AssetExecutionContext,
    db: TortoiseDBResource,
    enrichment: str,
) -> str:
    """
    Exports processed data to Excel reports (faculty sheets, programme sheets, etc.).

    This asset wraps the existing export logic in
    `easy_access.sheets.export.export_reports_async`.

    Args:
        context: Dagster execution context
        db: Tortoise database resource
        enrichment: Output from osiris_enrichment asset

    Returns:
        str: Summary of export results
    """
    async with db.yield_for_execution(context):
        from easy_access.sheets.export import (
            export_faculty_workflow_files,
            gather_faculty_data,
        )

        context.log.info("Starting Excel report export...")

        try:
            # Gather faculty data
            faculty_data = await gather_faculty_data(SETTINGS)

            if not faculty_data:
                context.log.warning("No faculty data to export")
                return "No faculty data to export"

            # Export using workflow files (inbox/in_progress/done structure)
            style_iter = 9  # Default style iteration
            await export_faculty_workflow_files(SETTINGS, faculty_data, style_iter)

            faculty_count = len(faculty_data)
            total_items = sum(len(df) for df in faculty_data.values())

            result_msg = (
                f"Exported reports for {faculty_count} faculties, "
                f"{total_items} total items"
            )
            context.log.info(result_msg)
            return result_msg

        except Exception as e:
            context.log.error(f"Excel export failed: {e}")
            return f"Excel export failed: {e}"
