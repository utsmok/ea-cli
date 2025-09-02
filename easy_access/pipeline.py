def ingest_faculty_updates(self) -> None:
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

        import asyncio
        asyncio.run(load_faculty_updates_to_staging(self.settings, df))
        logger.info("Faculty updates ingested into staging table.")