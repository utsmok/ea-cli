"""
Main orchestrator for the Easy Access tool.

This module defines the `EasyAccessTool` class, which is responsible for
coordinating the various data processing workflows, including:
- Ingesting raw copyright data.
- Synchronizing data between Excel sheets and the SQLite database.
- Generating weekly faculty sheets and overview reports.
- Enriching data with external sources like Osiris.
"""



from easy_access.pipeline import DataPipeline
from easy_access.settings import (
    EasyAccessSettings,
    Settings,
)


class EasyAccessTool:
    """
    Orchestrates the Easy Access tool's data processing workflows.

    This class handles the main sequence of operations, such as ingesting
    copyright data, updating the database from various sheet sources,
    and generating output sheets for faculties and programs.
    """

    settings: Settings
    ea_settings: EasyAccessSettings

    def __init__(self, settings_obj: Settings, ea_settings: EasyAccessSettings) -> None:
        """
        Initializes the EasyAccessTool.

        Args:
            settings_obj: The main Settings object for the application.
            ea_settings: EasyAccessSettings object containing runtime/CLI settings.
        """
        self.settings = settings_obj
        self.ea_settings = ea_settings

    def run(self) -> None:
        """
        Executes the configured sequence of processing functions.
        """
        pipeline = DataPipeline(settings=self.settings)
        pipeline.run()