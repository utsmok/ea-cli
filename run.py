"""
Easy Access Sheet Toolkit
Apr 2025
Samuel Mok / s.mok@utwente.nl / cip@utwente.nl
homepage: https://github.com/utsmok/ea-cli
Note: only tested on windows systems

This script runs the Easy Access tool for you.
All code can be found in folder 'easy_access', with easy_access_cli.py containing the main functionality.
See readme.md for more info, and the settings.yaml example file for specific parameters.

quickstart:
1. install uv (https://docs.astral.sh/uv/getting-started/installation/)
2. > uv run run.py --help
3. you'll probably see a lot of error messages, try to fix them and run again! :)
"""

from pathlib import Path
from typing import Annotated

import typer
import uvicorn

from easy_access.main import EasyAccessTool
from easy_access.settings import (  # Import the main Settings class
    SETTINGS,
    EasyAccessSettings,
)
from easy_access.sheets.backup import (
    Backupper,
    RestoreOptions,
    RestoreStrategy,
)
from easy_access.sheets.sheet import create_export_sheet
from loguru import logger

# INIT typer apps

app = typer.Typer(
    name="ea-cli",
    help="Easy Access toolkit for managing faculty sheet data.",
    add_completion=False,
)

preprocess_app = typer.Typer(
    name="preprocess",
    help="Pre-processing: PDF download, classification, deduplication.",
)
app.add_typer(preprocess_app)

backup_app = typer.Typer(name="backup", help="Backup and restore operations.")
app.add_typer(backup_app)

# Commands for main app

@app.command(name="process")
def process_data(
    changes: Annotated[
        bool,
        typer.Option(
            help="Only add items that have been changed to new faculty sheets.",
            rich_help_panel="Processing Options",
        ),
    ] = True,
    osiris_update: Annotated[
        bool,
        typer.Option(
            help="If enabled, will retrieve fresh osiris data for all course + people page data.",
            rich_help_panel="Enrichment Options",
        ),
    ] = False,
    osiris_full_refresh: Annotated[
        bool,
        typer.Option(
            help="If osiris_update is enabled, this flag will toggle retrieval of fresh osiris data for either ALL data, or only data currently missing osiris info.",
            rich_help_panel="Enrichment Options",
        ),
    ] = True,
    other_sheet: Annotated[
        Path | None,
        typer.Option(
            help="Path to a xlsx sheet to read instead of CopyRight Data.",
            rich_help_panel="Input Options",
            exists=True,
            file_okay=True,
            dir_okay=False,
        ),
    ] = None,
    disable_writes: Annotated[
        bool,
        typer.Option(
            help="Disable all write operations.",
            rich_help_panel="Processing Options",
        ),
    ] = False,
    single_faculty: Annotated[
        str | None,
        typer.Option(
            help="Only run the tool for a single faculty. use the faculty abbreviation as the parameter (e.g. 'BMS').",
            rich_help_panel="Processing Options",
        ),
    ] = None,
) -> None:
    """Runs the main Easy Access data processing workflow."""


    if other_sheet:
        try:
            other_sheet = Path(other_sheet)  # Ensure it's a Path object
            logger.info(f"Reading in data from other sheet: {other_sheet.absolute()}")
        except Exception as e:
            logger.warning(f"Failed to parse path to other sheet: {e}")
            other_sheet = None

    # Load settings from CLI params, using main SETTINGS for base dir config
    ea_settings = EasyAccessSettings.create_for_runtime(  # Renamed method
        main_settings=SETTINGS,  # Pass the main SETTINGS object
        export=False,  # 'export' is now a separate command
        only_changes=changes,
        refresh_osiris_data=osiris_update,
        other_sheet=other_sheet,
        only_retrieve_missing_osiris_data=not osiris_full_refresh,  # Corrected logic
        disable_writes=disable_writes,
        faculty=single_faculty,
    )

    # The main SETTINGS object is loaded globally in easy_access.settings
    tool = EasyAccessTool(settings_obj=SETTINGS, ea_settings=ea_settings)
    tool.run()

    logger.success("Main processing done!")


@app.command(name="dashboard")
def run_dashboard(port: Annotated[int, typer.Option(help="Port to serve the dashboard on.")] = 8000, host: Annotated[str, typer.Option(help="Host to serve the dashboard on.")] = "0.0.0.0") -> None:
    """Serves the easy_access dashboard.
    """
    logger.info("Serving the easy_access dashboard.")
    logger.info(f"Once launched, it will be available at http://{host}:{port}.")
    logger.info("Press Ctrl+C or close this terminal window to stop the server.")
    uvicorn.run(
        "easy_access.dashboard.dash:app",
        host=host,
        port=port,
        reload=SETTINGS.dashboard_reload,
    )
    logger.success("Dashboard server stopped.")


@app.command(name="export")
def run_export(
    single_faculty: Annotated[
        str | None,
        typer.Option(
            help="Only export data for a single faculty. Use the faculty abbreviation (e.g. 'BMS').",
        ),
    ] = None,
) -> None:
    """Creates export sheets."""
    if single_faculty:
        logger.info(f"Exporting data for faculty: {single_faculty}")
        create_export_sheet(settings=SETTINGS, faculty=single_faculty)
    else:
        logger.info("Creating export sheets for all faculties.")
        create_export_sheet(
            settings=SETTINGS
        )
    logger.success("Done creating export sheets.")

# Commands for backup app

@backup_app.command(name="create")
def create_backup_command() -> None:
    """Creates a backup of the current data based on settings.yaml."""
    backupper = Backupper()
    if SETTINGS.backup_settings.backup_all:  # Check main settings
        logger.info("Creating backup as per settings.yaml (backup_all: true).")
        backupper.backup_files()
    else:
        logger.info("Backup not created as per settings.yaml (backup_all: false or not set).")


@backup_app.command(name="restore")
def restore_backup_command(
    restore_dir: Annotated[
        RestoreOptions,
        typer.Option(help="Set which backup to restore."),
    ] = RestoreOptions.LATEST,
    restore_strategy: Annotated[
        RestoreStrategy,
        typer.Option(help="Set the strategy for restoring the backup."),
    ] = RestoreStrategy.REPLACE,
) -> None:
    """Restores data from a backup."""
    backupper = Backupper()
    logger.info(
        f"Restoring backup from '{restore_dir.value}' with strategy '{restore_strategy.value}'."
    )
    backupper.restore_backup(
        strategy=restore_strategy,
        select=restore_dir,
    )
    logger.success("Backup restoration process finished.")


# Commands for pre-processing app

@preprocess_app.command(name="run_all")
def run_all_preprocess(
    osiris_update: Annotated[
        bool,
        typer.Option(
            help="If enabled, will retrieve fresh osiris data for all course + people page data.",
        ),
    ] = False,
    osiris_full_refresh: Annotated[
        bool,
        typer.Option(
            help="If osiris_update is enabled, this flag will toggle retrieval of fresh osiris data for either ALL data, or only data currently missing osiris info.",
        ),
    ] = True,
    single_faculty: Annotated[
        str | None,
        typer.Option(
            help="Only run the tool for a single faculty. use the faculty abbreviation as the parameter (e.g. 'BMS').",
        ),
    ] = None,
    # download: Annotated[
    #     bool,
    #     typer.Option(help="Download pdfs from canvas."),
    # ] = False,
    # classify: Annotated[
    #     bool,
    #     typer.Option(help="Classify the pdfs by LLM."),
    # ] = False,
    # deduplicate: Annotated[
    #     bool,
    #     typer.Option(help="Deduplicate the pdfs."),
    # ] = False,
) -> None:
    """(Currently Stubs) Runs all pre-processing steps: PDF download, classification, deduplication."""
    logger.info("Running pre-processing steps (download, deduplicate, classify)...")
    logger.info("First, running the tool in read-only mode to update DB data if needed.")

    ea_temp_settings = EasyAccessSettings.create_for_runtime(  # Renamed method
        main_settings=SETTINGS,  # Pass the main SETTINGS object
        export=False,
        only_changes=True,  # Assuming this is a sensible default for pre-processing
        refresh_osiris_data=osiris_update,
        other_sheet=None,  # Pre-processing typically works on existing DB data
        only_retrieve_missing_osiris_data=not osiris_full_refresh,
        disable_writes=True,
        faculty=single_faculty,
    )
    temp_tool = EasyAccessTool(settings_obj=SETTINGS, ea_settings=ea_temp_settings)
    temp_tool.run()
    logger.success("Done updating data in read-only mode for pre-processing.")

    logger.info("Actual pre-processing steps (download, deduplicate, classify) follow.")
    logger.warning(
        "Download, deduplication, and classification steps are currently stubs and not implemented."
    )
    # if download:
    #     logger.info("Downloading PDFs...")
    #     # ... downloader logic ...
    # if deduplicate:
    #     logger.info("Deduplicating PDFs...")
    #     # ... deduplicator logic ...
    # if classify:
    #     logger.info("Classifying PDFs...")
    #     # ... classifier logic ...
    logger.success("Pre-processing steps finished")




if __name__ == "__main__":
    app()
