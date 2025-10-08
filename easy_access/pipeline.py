"""
This module contains the main data processing pipeline for the Easy Access tool.
"""

from loguru import logger

from easy_access.settings import EasyAccessSettings, Settings
from easy_access.utils import run_sync


class DataPipeline:
    """
    Orchestrates the data processing workflow.

    Provides both synchronous and asynchronous interfaces.
    """

    def __init__(
        self, settings: Settings, ea_settings: EasyAccessSettings | None = None
    ):
        self.settings = settings
        self.ea_settings = ea_settings

    # Synchronous wrappers (loop-aware)
    def run(self) -> None:
        """Run full pipeline synchronously."""
        return run_sync(self.run_async())

    def ingest_raw_data(self, file_path: str | None = None) -> None:
        """Synchronous wrapper for ingesting raw data."""
        return run_sync(self.ingest_raw_data_async(file_path))

    def ingest_faculty_updates(self) -> None:
        """Synchronous wrapper for ingesting faculty updates."""
        return run_sync(self.ingest_faculty_updates_async())

    def process_data(self) -> None:
        """Synchronous wrapper for processing staged data."""
        return run_sync(self.process_data_async())

    def update_relations(self) -> None:
        """Synchronous wrapper for updating relations."""
        return run_sync(self.update_relations_async())

    def export_reports(self) -> None:
        """Synchronous wrapper for exporting reports."""
        return run_sync(self.export_reports_async())

    def enrich_data(self) -> None:
        """Synchronous wrapper for enriching data with OSIRIS information."""
        return run_sync(self.enrich_data_async())

    def verify_file_existence(self) -> None:
        """Synchronous wrapper for verifying file existence."""
        return run_sync(self.verify_file_existence_async())

    def download_pdfs(self) -> None:
        """Synchronous wrapper for downloading PDFs."""
        return run_sync(self.download_pdfs_async())

    def parse_pdfs(self) -> None:
        """Synchronous wrapper for parsing PDFs."""

        return run_sync(self.parse_pdfs_async())

    def process_db_changes(self) -> None:
        """Synchronous wrapper for processing DB changes."""
        return run_sync(self.process_db_changes_async())

    def close_connections(self) -> None:
        """
        Closes any open connections, such as database connections.
        """
        from easy_access.db.base import close_connections

        logger.info("Closing database connections...")
        return run_sync(close_connections())
        logger.info("Database connections closed.")

    # Asynchronous interface
    async def run_async(self) -> None:
        """Runs the full data processing pipeline asynchronously."""
        logger.info("Starting data processing pipeline...")
        await self.ingest_raw_data_async()
        await self.ingest_faculty_updates_async()
        await self.process_data_async()
        await self.update_relations_async()
        await self.enrich_data_async()

        # Conditionally run file existence verification
        if self.ea_settings and not getattr(self.ea_settings, "no_file_exists", False):
            await self.verify_file_existence_async()
        else:
            logger.info("File existence verification disabled, skipping...")

        if self.ea_settings and getattr(self.ea_settings, "no_pdf_download", False):
            logger.info("PDF downloading disabled, skipping...")
        else:
            await self.download_pdfs_async()

        if self.ea_settings and getattr(self.ea_settings, "no_pdf_parse", False):
            logger.info("PDF parsing disabled, skipping...")
        else:
            await self.parse_pdfs_async()

        await self.process_db_changes_async()

        await self.export_reports_async()
        logger.info("Data processing pipeline finished.")

    async def ingest_raw_data_async(self, file_path: str | None = None) -> None:
        """Ingests raw data from a copyright export Excel file into the staging table."""
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
        """Ingests data from faculty Excel sheets into the staging table."""
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
        """Processes the staged data and updates the main CopyrightItem table."""
        from easy_access.db.update import (
            process_staged_faculty_updates,
            process_staged_raw_data,
        )

        logger.info("Processing staged data...")
        await process_staged_raw_data(self.settings)
        await process_staged_faculty_updates(self.settings)
        logger.info("Staged data processed.")

    async def update_relations_async(self) -> None:
        """Updates database relations (duplicates, course links)."""
        from easy_access.db.relations import update_relations_async

        logger.info("Updating relations...")
        await update_relations_async(self.settings)
        logger.info("Relations updated.")

    async def export_reports_async(self) -> None:
        """Exports processed data to Excel reports (faculty sheets, programme sheets, etc.)."""
        from easy_access.sheets.export import (
            export_faculty_workflow_files,
            export_reports_async,
        )

        logger.info("Exporting reports...")

        # Run either the legacy top-level exporter OR the new workflow exporter
        # depending on the runtime flag in ea_settings. Previously we always ran
        # the legacy exporter and then optionally the workflow exporter which
        # resulted in duplicate/undesired outputs. Choose one path here.
        if self.ea_settings and getattr(self.ea_settings, "export_workflow", False):
            # gather faculty data and call the workflow writer
            from easy_access.sheets.export import gather_faculty_data

            faculty_data = await gather_faculty_data(self.settings)
            if faculty_data:
                style_iter = 9
                await export_faculty_workflow_files(
                    self.settings, faculty_data, style_iter
                )
        else:
            # Default: run the legacy top-level exporter which handles full orchestration
            await export_reports_async(self.settings)
        logger.info("Reports exported.")

    async def enrich_data_async(self) -> None:
        """Enriches data with OSIRIS course and person information."""
        from easy_access.enrichment.osiris import enrich_async

        logger.info("Enriching data with OSIRIS information...")
        await enrich_async(self.settings)

        await self.update_relations_async()
        logger.info("Data enrichment completed.")

    async def verify_file_existence_async(self) -> None:
        """Verifies file existence for copyright items based on TTL policies."""
        from easy_access.maintenance.file_existence import refresh_file_existence_async

        logger.info("Verifying file existence...")
        ttl_days = getattr(self.settings, "file_exists_ttl_days", 7)
        # Get rate limit delay from settings or use default
        rate_limit_delay = getattr(self.settings, "file_exists_rate_limit_delay", 0.05)
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

    async def download_pdfs_async(self) -> None:
        """Downloads undownloaded PDFs from Canvas and creates PDF entries in the database."""
        from easy_access.pdf.download import download_pdfs

        logger.info("Downloading undownloaded PDFs...")
        await download_pdfs(self.settings)
        logger.info("PDF downloading completed.")

    async def parse_pdfs_async(self) -> None:
        """Parses undparsed PDFs to extract text, filehashes, metadata..."""
        from easy_access.pdf.parse import ocr_pdfs, parse_pdfs

        logger.info("Parsing unparsed PDFs...")
        await parse_pdfs()
        await ocr_pdfs()
        logger.info("PDF parsing completed.")

    async def process_db_changes_async(self) -> None:
        """
        Processes database changes to ensure all fields are up to date.
        First map entered V1 classifications to V2,
        then go through the DB and ensure all fields are updated based on current status.
        e.g. set workflow status to Done if a file has a proper classification and no action is needed, etc"""
        from easy_access.db.update import (
            map_v1_to_v2_classifications,
            update_workflow_status_from_db,
        )

        await map_v1_to_v2_classifications(self.settings)

        await update_workflow_status_from_db(self.settings)
