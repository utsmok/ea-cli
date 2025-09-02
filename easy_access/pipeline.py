"""
This module contains the main data processing pipeline for the Easy Access tool.
"""

from loguru import logger

class DataPipeline:
    """
    Orchestrates the data processing workflow.
    """

    def __init__(self, settings):
        self.settings = settings

    def run(self):
        """
        Runs the full data processing pipeline.
        """
        logger.info("Starting data processing pipeline...")
        self.ingest_raw_data()
        # The following methods will be implemented in subsequent steps
        # self.ingest_faculty_updates()
        # self.process_data()
        # self.export_reports()

    def ingest_raw_data(self, file_path: str | None = None) -> None:
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

        import asyncio
        asyncio.run(load_raw_copyright_data_to_staging(self.settings, df))
        logger.info("Raw copyright data ingested into staging table.")
        logger.info("Data processing pipeline finished.")
