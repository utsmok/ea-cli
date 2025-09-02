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

    def __init__(self, settings):
        self.settings = settings

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

    # Asynchronous interface
    async def run_async(self) -> None:
        """
        Runs the full data processing pipeline asynchronously.
        """
        logger.info("Starting data processing pipeline...")
        await self.ingest_raw_data_async()
        await self.ingest_faculty_updates_async()
        await self.process_data_async()
        # await self.export_reports_async()  # To be implemented
        logger.info("Data processing pipeline finished.")

    async def ingest_raw_data_async(self, file_path: str | None = None) -> None:
        """
        Ingests raw data from a copyright export Excel file into the staging table.
        """
        from easy_access.sheets.sheet import read_copyright_export
        from easy_access.db.ingest import load_raw_copyright_data_to_staging

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
        from easy_access.sheets.sheet import read_faculty_sheets
        from easy_access.db.ingest import load_faculty_updates_to_staging

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
        from easy_access.db.update import process_staged_raw_data, process_staged_faculty_updates

        logger.info("Processing staged data...")
        await process_staged_raw_data(self.settings)
        await process_staged_faculty_updates(self.settings)
        logger.info("Staged data processed.")
