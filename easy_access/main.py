"""
Main orchestrator for the Easy Access tool.

This module defines the `EasyAccessTool` class, which is responsible for
coordinating the various data processing workflows, including:
- Ingesting raw copyright data.
- Synchronizing data between Excel sheets and the SQLite database.
- Generating weekly faculty sheets and overview reports.
- Enriching data with external sources like Osiris.
"""

import asyncio
import contextlib
import datetime
import os
from collections.abc import Callable
from pathlib import Path

import polars as pl
from loguru import logger

from easy_access.db.base import ensure_db_inited
from easy_access.db.ingest import load_base_data, load_raw_copyright_data
from easy_access.db.retrieve import retrieve_copyright_items, retrieve_full_data
from easy_access.db.update import update_copyright_items
from easy_access.settings import (
    DirSetting,
    EasyAccessSettings,
    Settings,
)
from easy_access.sheets.analysis import create_faculty_overviews
from easy_access.sheets.enrichment import update_osiris_data
from easy_access.sheets.sheet import (
    create_export_sheet,
    finalize_sheet,
    read_copyright_export,
    store_complete_data,
)
from easy_access.utils import Directory, File
from easy_access.utilities.file_exists import check_file_exists
from easy_access.api_keys import canvas as api_token
from easy_access.pipeline import DataPipeline


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