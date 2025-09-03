"""
This module contains the main data processing pipeline for the Easy Access tool.
"""

import asyncio

from loguru import logger


class DataPipeline:
    """
    Orchestrates the data processing workflow.

    Provides both synchronous and asynchronous interfaces:
    - Use sync methods (run_sync, ingest_raw_data_sync, etc.) for simple synchronous usage
    - Use async methods (run_async, ingest_raw_data_async, etc.) for async contexts
    """

    def __init__(self, settings, ea_settings=None):
        self.settings = settings
        self.ea_settings = ea_settings

    # Synchronous interface (backwards compatible)
    def run(self) -> None:
        """
        Runs the full data processing pipeline synchronously.
        This is the backwards-compatible synchronous entrypoint.
        """
        asyncio.run(self.run_async())

    def ingest_raw_data(self, file_path: str | None = None) -> None:
        """
        Synchronous wrapper for ingesting raw data.
        """
        asyncio.run(self.ingest_raw_data_async(file_path))

    def ingest_faculty_updates(self) -> None:
        """
        Synchronous wrapper for ingesting faculty updates.
        """
        asyncio.run(self.ingest_faculty_updates_async())

    def process_data(self) -> None:
        """
        Synchronous wrapper for processing staged data.
        """
        asyncio.run(self.process_data_async())

    def update_relations(self) -> None:
        """
        Synchronous wrapper for updating relations.
        """
        asyncio.run(self.update_relations_async())

    def export_reports(self) -> None:
        """
        Synchronous wrapper for exporting reports.
        """
        asyncio.run(self.export_reports_async())

    def enrich_data(self) -> None:
        """
        Synchronous wrapper for enriching data with OSIRIS information.
        """
        asyncio.run(self.enrich_data_async())

    def verify_file_existence(self) -> None:
        """
        Synchronous wrapper for verifying file existence.
        """
        asyncio.run(self.verify_file_existence_async())

    # Asynchronous interface
    async def run_async(self) -> None:
        """
        Runs the full data processing pipeline asynchronously.
        """
        logger.info("Starting data processing pipeline...")
        await self.ingest_raw_data_async()
        await self.ingest_faculty_updates_async()
        await self.process_data_async()
        await self.update_relations_async()
        await self.enrich_data_async()

        # Conditionally run file existence verification
        if self.ea_settings and not self.ea_settings.no_file_exists:
            await self.verify_file_existence_async()
        else:
            logger.info("File existence verification disabled, skipping...")

        await self.export_reports_async()
        logger.info("Data processing pipeline finished.")

    async def ingest_raw_data_async(self, file_path: str | None = None) -> None:
        """
        Ingests raw data from a copyright export Excel file into the staging table.
        """
        from easy_access.db.ingest import load_raw_copyright_data_to_staging
        from easy_access.sheets.sheet import read_copyright_export

        logger.info("Ingesting raw copyright data...")
        if file_path:
            from easy_access.utils import File

            file = File(file_path)
            _, df = read_copyright_export(self.settings, file=file)
        else:
            _, df = read_copyright_export(self.settings)

        if df.is_empty():
            logger.warning("No new copyright data found to ingest.")
            return

        await load_raw_copyright_data_to_staging(self.settings, df)
        logger.info("Raw copyright data ingested into staging table.")

    async def ingest_faculty_updates_async(self) -> None:
        """
        Ingests data from faculty Excel sheets into the staging table.
        """
        from easy_access.db.ingest import load_faculty_updates_to_staging
        from easy_access.sheets.sheet import read_faculty_sheets

        logger.info("Ingesting faculty updates...")
        df = read_faculty_sheets(self.settings)

        if df.is_empty():
            logger.warning("No faculty updates found to ingest.")
            return

        await load_faculty_updates_to_staging(self.settings, df)
        logger.info("Faculty updates ingested into staging table.")

    async def process_data_async(self) -> None:
        """
        Processes the staged data and updates the main CopyrightItem table.
        """
        from easy_access.db.update import (
            process_staged_faculty_updates,
            process_staged_raw_data,
        )

        logger.info("Processing staged data...")
        await process_staged_raw_data(self.settings)
        await process_staged_faculty_updates(self.settings)
        logger.info("Staged data processed.")

    async def update_relations_async(self) -> None:
        """
        Updates database relations (duplicates, course links).
        """
        from easy_access.db.relations import update_relations_async

        logger.info("Updating relations...")
        await update_relations_async(self.settings)
        logger.info("Relations updated.")

    async def export_reports_async(self) -> None:
        """
        Exports processed data to Excel reports (faculty sheets, programme sheets, etc.).
        """
        from easy_access.sheets.export import export_reports_async

        logger.info("Exporting reports...")
        await export_reports_async(self.settings)
        logger.info("Reports exported.")

    async def enrich_data_async(self) -> None:
        """
        Enriches data with OSIRIS course and person information.
        """
        from easy_access.enrichment.osiris import enrich_async

        logger.info("Enriching data with OSIRIS information...")
        await enrich_async(self.settings)
        logger.info("Data enrichment completed.")

    async def verify_file_existence_async(self) -> None:
        """
        Verifies file existence for copyright items based on TTL policies.
        """
        from easy_access.maintenance.file_existence import refresh_file_existence_async

        logger.info("Verifying file existence...")
        ttl_days = getattr(self.settings, "file_exists_ttl_days", 30)
        # Get rate limit delay from settings or use default
        rate_limit_delay = getattr(self.settings, "file_exists_rate_limit_delay", 0.1)

        result = await refresh_file_existence_async(
            self.settings,
            ttl_days=ttl_days,
            batch_size=1000,
            max_concurrent=50,
            rate_limit_delay=rate_limit_delay,
        )

        if "error" in result:
            logger.error(f"File existence verification failed: {result['error']}")
        else:
            logger.info(
                f"File existence verification completed: "
                f"{result.get('checked', 0)} checked, "
                f"{result.get('exists', 0)} exist, "
                f"{result.get('not_exists', 0)} not found"
            )
