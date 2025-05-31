"""
Easy Access Sheet Toolkit - Command Line Interface
April 2025
Samuel Mok / s.mok@utwente.nl / cip@utwente.nl
Homepage: https://github.com/utsmok/ea-cli

This script provides a Command Line Interface (CLI) to run the Easy Access toolkit.
It uses Typer for CLI argument parsing and command definition.
The main functionalities include processing copyright data, managing backups,
and optionally serving a web dashboard.

Refer to `readme.md` for detailed usage instructions and `settings.yaml` for
configuration parameters.

Quickstart:
1. Ensure Python and pip are installed.
2. Install uv: `pip install uv` (or see https://docs.astral.sh/uv/getting-started/installation/)
3. Run the CLI: `uv run python run.py --help`
   (If not using `uv run`, ensure dependencies from `pyproject.toml` are installed in your environment.)
"""

import logging
from pathlib import Path
from typing import Annotated, Optional # Optional for Python <3.10 compatibility with Path | None

import typer
import uvicorn

from easy_access.main import EasyAccessTool
from easy_access.settings import SETTINGS, EasyAccessSettings # SETTINGS needed for backup default
from easy_access.sheets.backup import (
    BackupFlag,
    Backupper,
    RestoreOptions,
    RestoreStrategy,
)
# create_export_sheet is part of EasyAccessTool now, direct import might be for a special CLI mode.
# from easy_access.sheets.sheet import create_export_sheet

cli_app = typer.Typer(help="Easy Access Sheet Toolkit: Process copyright data and generate reports.")
logger = logging.getLogger(__name__)


@cli_app.command()
def cli(
    dashboard: Annotated[
        bool,
        typer.Option(
            help="Serve the Easy Access web dashboard. Ignores all other parameters.",
            rich_help_panel="Mode", # Changed panel name
        ),
    ] = False,
    export_mode: Annotated[ # Renamed from 'export' to avoid conflict with settings attribute
        bool,
        typer.Option(
            "--export-only", # CLI flag is more explicit
            help="Run in standalone export mode. Creates export sheets based on current DB data. "
                 "This mode does not run the main data processing pipeline. "
                 "Use with --single-faculty option if needed.",
            rich_help_panel="Mode",
        ),
    ] = False,
    include_export_in_run: Annotated[ # New option to include export as part of a normal run
        bool,
        typer.Option(
            "--include-export",
            help="Include creation of export sheets at the end of a normal processing run.",
            rich_help_panel="Functions",
        ),
    ] = False,
    changes: Annotated[
        bool,
        typer.Option(
            "--only-changes/--all-items",
            help="Processing focuses on new/changed items vs. all items for faculty sheets.",
            rich_help_panel="Functions",
        ),
    ] = True,
    osiris_update: Annotated[
        bool,
        typer.Option(
            help="Refresh Osiris data (courses, people pages).",
            rich_help_panel="Enrichment",
        ),
    ] = False,
    osiris_full_refresh: Annotated[
        bool,
        typer.Option(
            help="If refreshing Osiris data, re-retrieve for ALL items, not just those missing data.",
            rich_help_panel="Enrichment",
        ),
    ] = True,
    other_sheet: Annotated[
        Optional[Path], # Use Optional for older Python compatibility
        typer.Option(
            help="Path to an XLSX sheet to use as the primary data source instead of default raw Copyright data.",
            rich_help_panel="Input Data", # Changed panel name
            exists=True, # Typer will check if path exists
            file_okay=True,
            dir_okay=False,
            resolve_path=True, # Typer will resolve to absolute path
        ),
    ] = None,
    backup: Annotated[
        BackupFlag,
        typer.Option(
            help="Backup/restore data: 'backup', 'restore', 'none', or 'default' (uses settings.yaml).",
            rich_help_panel="Backup/Restore",
            case_sensitive=False, # Allow lowercase enum values
        ),
    ] = BackupFlag.DEFAULT.value, # Default to enum member's value
    disable_writes: Annotated[
        bool,
        typer.Option(
            help="Disable all file writing operations (e.g., Excel sheets, backups).",
            rich_help_panel="Safety", # Changed panel name
        ),
    ] = False,
    restore_dir: Annotated[
        RestoreOptions,
        typer.Option(
            help="Specify which backup to restore: 'latest', 'oldest', or a specific YYYY-MM-DD_HH-MM-SS folder name.",
            rich_help_panel="Backup/Restore",
            case_sensitive=False,
        ),
    ] = RestoreOptions.LATEST.value,
    restore_strategy: Annotated[
        RestoreStrategy,
        typer.Option(
            help="Restore strategy: 'replace' (full replacement) or 'merge' (overwrite conflicts).",
            rich_help_panel="Backup/Restore",
            case_sensitive=False,
        ),
    ] = RestoreStrategy.REPLACE.value,
    # Removed download, classify, deduplicate as they were not implemented
    # download: Annotated[bool, typer.Option(help="Download PDFs from Canvas.", rich_help_panel="Experimental")] = False,
    # classify: Annotated[bool, typer.Option(help="Classify PDFs by LLM.", rich_help_panel="Experimental")] = False,
    # deduplicate: Annotated[bool, typer.Option(help="Deduplicate PDFs.", rich_help_panel="Experimental")] = False,
    single_faculty: Annotated[
        Optional[str], # Use Optional for older Python
        typer.Option(
            help="Run the tool for a single faculty only (e.g., 'BMS').",
            rich_help_panel="Functions",
        ),
    ] = None,
    log_level: Annotated[
        str,
        typer.Option(
            help="Set logging level (e.g., DEBUG, INFO, WARNING).",
            rich_help_panel="Advanced",
            case_sensitive=False,
        )
    ] = "INFO",
) -> None:
    """
    Easy Access Toolkit for managing faculty copyright sheet data.

    This CLI allows you to process copyright data, generate various Excel reports,
    manage backups, and (optionally) run a web dashboard for data interaction.
    Configure primary paths and behaviors in `settings.yaml`.
    """

    # Setup standard Python logging
    # The level set here will be the base; modules use getLogger(__name__)
    log_level_upper = log_level.upper()
    numeric_level = getattr(logging, log_level_upper, None)
    if not isinstance(numeric_level, int):
        print(f"Warning: Invalid log level '{log_level}'. Defaulting to INFO.") # Use print before logging fully set up
        numeric_level = logging.INFO

    # Basic configuration for all loggers.
    # Adding process ID can be useful if multiple instances run.
    logging.basicConfig(
        level=numeric_level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    logger.info(f"Log level set to {log_level_upper}")


    if dashboard:
        logger.info("Starting web dashboard mode...")
        logger.info("Dashboard will be available at http://localhost:8000 (or your configured host/port).")
        logger.info("Press Ctrl+C to stop the server.")
        try:
            uvicorn.run( # Hardcoded app string is common for entry points
                "easy_access.dashboard.dash:app", host="0.0.0.0", port=8000, reload=True
            )
        except Exception as e_dash:
            logger.error(f"Failed to start dashboard: {e_dash}", exc_info=True)
            raise typer.Exit(code=1)
        logger.info("Dashboard server stopped. Exiting.")
        raise typer.Exit(code=0)

    if export_mode:
        logger.info("Running in standalone EXPORT ONLY mode.")
        # This mode requires data to be present in the DB. It creates export sheets directly.
        # It does not run the main processing pipeline.
        # `EasyAccessTool` is not used in this specific mode; it calls a function that
        # would need to initialize DB connection and fetch data itself.
        # The refactored `create_export_sheet` is now a method of `EasyAccessTool`.
        # To make this mode work, we would need to:
        # 1. Initialize settings (minimal, for DB path)
        # 2. Create an EasyAccessTool instance.
        # 3. Call tool.process_raw_copyright_data() to load data into tool.copyright_data
        #    (or a more lightweight data loading method if available).
        # 4. Then call tool.create_export_sheet().
        logger.warning("Standalone export mode (--export-only) currently implies data is already processed and in DB.")
        logger.warning("This CLI option might need further refinement to manage data loading for export.")

        # Simplified temporary approach: Create settings, tool, load data, then export
        temp_export_settings = EasyAccessSettings(
            export=True, # This flag tells the tool to include export in its run.
            other_sheet=other_sheet, # Use the CLI provided other_sheet if any
            faculty=single_faculty,
            # Other settings can be default or from settings.yaml via global SETTINGS
            dirs=SETTINGS.dirs,
            disable_writes=disable_writes # Respect disable_writes for safety
        )
        export_tool_instance = EasyAccessTool(temp_export_settings)
        logger.info("Loading data for export...")
        export_tool_instance.process_raw_copyright_data() # Load data into export_tool_instance.copyright_data
        if not export_tool_instance.copyright_data.is_empty():
            export_tool_instance.create_export_sheet() # Call the method
            logger.info("Standalone export sheet creation process finished.")
        else:
            logger.error("No data loaded for standalone export. Export sheets not created.")
        raise typer.Exit(code=0)

    # Backup and Restore Logic (before main processing)
    # Enum instances are now correctly created by Typer from string values.
    backupper = Backupper()
    if backup == BackupFlag.BACKUP:
        backupper.backup_files()
    elif backup == BackupFlag.RESTORE:
        backupper.restore_backup(strategy=restore_strategy, select=restore_dir)
    elif backup == BackupFlag.DEFAULT and SETTINGS.backup_settings.backup_all:
        backupper.backup_files()
    elif backup == BackupFlag.NONE:
        logger.info("Backup/restore skipped as per '--backup none' option.")
    # No 'else' needed as Typer handles invalid enum choices.

    # `other_sheet` is already a resolved Path object due to `resolve_path=True` in Typer.Option
    # No further validation needed here unless specific content checks are required.

    # This block for pre-processing seems to be a special case.
    # It was noted as not fully implemented. For now, it's a placeholder.
    # if any([download, deduplicate, classify]): # These flags were removed
    #     logger.info("Running pre-processing steps (download/deduplicate/classify)...")
    #     logger.warning("The pre-processing steps are currently placeholders and not fully implemented.")
    #     # ... (rest of the pre-processing logic if it were to be implemented) ...

    # Main tool execution
    logger.info("Initializing EasyAccessTool for main processing run...")
    runtime_settings = EasyAccessSettings(
        export=include_export_in_run, # Use the new flag for including export in normal run
        only_changes=changes,
        refresh_osiris_data=osiris_update,
        other_sheet=other_sheet, # Already a Path object or None
        only_retrieve_missing_osiris_data=not osiris_full_refresh, # Note: True means only missing, False means full
        disable_writes=disable_writes,
        faculty=single_faculty,
        dirs=SETTINGS.dirs # Ensure dirs from global SETTINGS are passed
    )

    tool_instance = EasyAccessTool(runtime_settings)
    tool_instance.run()

    logger.info("Easy Access tool run completed successfully!")


if __name__ == "__main__":
    # This basicConfig is for when the script is run directly, not as part of `uv run`.
    # It's good practice for the main entry point to configure logging.
    # However, the one inside cli() will take precedence if cli() is called.
    # Consider moving basicConfig to the top level of the script if it should always apply.
    # For now, cli() handles its own config.
    cli_app()
