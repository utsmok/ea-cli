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

# from easy_access.classification.classifier_api import main
# from easy_access.classification.downloader import Downloader
# from easy_access.classification.pdf_handling import enrich_pdfs
from easy_access.main import EasyAccessTool
from easy_access.settings import SETTINGS, EasyAccessSettings
from easy_access.sheets.backup import (
    BackupFlag,
    Backupper,
    RestoreOptions,
    RestoreStrategy,
)
from easy_access.sheets.sheet import create_export_sheet
from easy_access.utils import cool, info, warn

cli_app = typer.Typer()


@cli_app.command()
def cli(
    dashboard: Annotated[
        bool,
        typer.Option(
            help="Serves the easy_access dashboard. Ignores all other parameters.",
            rich_help_panel="Frontend",
        ),
    ] = False,
    export: Annotated[
        bool,
        typer.Option(
            case_sensitive=False,
            help="Also create export sheets?",
            rich_help_panel="Functions",
        ),
    ] = False,
    changes: Annotated[
        bool,
        typer.Option(
            help="Only add items that have been changed to new faculty sheets.",
            rich_help_panel="Functions",
        ),
    ] = True,
    osiris_update: Annotated[
        bool,
        typer.Option(
            help="If enabled, will retrieve fresh osiris data for all course + people page data.",
            rich_help_panel="Enrichment",
        ),
    ] = False,
    osiris_full_refresh: Annotated[
        bool,
        typer.Option(
            help="If osiris_update is enabled, this flag will toggle retrieval of fresh osiris data for either ALL data, or only data currently missing osiris info.",
            rich_help_panel="Enrichment",
        ),
    ] = True,
    other_sheet: Annotated[
        Path | None,
        typer.Option(
            help="Path to a xlsx sheet to read instead of CopyRight Data.",
            rich_help_panel="Read in data from alternate source",
            exists=True,
            file_okay=True,
            dir_okay=False,
        ),
    ] = None,
    backup: Annotated[
        BackupFlag,
        typer.Option(
            help="Backup/restore data before starting, neither, or based on settings.yaml (default).",
            rich_help_panel="Backup/Restore",
        ),
    ] = BackupFlag.DEFAULT.value,
    disable_writes: Annotated[
        bool,
        typer.Option(
            help="Disable all write operations.",
            rich_help_panel="Functions",
        ),
    ] = False,
    restore_dir: Annotated[
        RestoreOptions,
        typer.Option(
            help="Set which backup to restore.",
            rich_help_panel="Backup/Restore",
        ),
    ] = RestoreOptions.LATEST.value,
    restore_strategy: Annotated[
        RestoreStrategy,
        typer.Option(
            help="Set the strategy for restoring the backup. 'replace' will fully replace the faculties dir, 'merge' will only overwrite conflicts (with the prioritized source file) and keep the rest",
            rich_help_panel="Backup/Restore",
        ),
    ] = RestoreStrategy.REPLACE.value,
    download: Annotated[
        bool,
        typer.Option(
            help="Download pdfs from canvas.",
            rich_help_panel="Functions",
        ),
    ] = False,
    classify: Annotated[
        bool,
        typer.Option(
            help="Classify the pdfs by LLM.",
            rich_help_panel="Enrichment",
        ),
    ] = False,
    deduplicate: Annotated[
        bool,
        typer.Option(
            help="Deduplicate the pdfs.",
            rich_help_panel="Enrichment",
        ),
    ] = False,
    single_faculty: Annotated[
        str | None,
        typer.Option(
            help="Only run the tool for a single faculty. use the faculty abbreviation as the parameter (e.g. 'BMS').",
            rich_help_panel="Functions",
        ),
    ] = None,
) -> None:
    """Easy Access toolkit for managing faculty sheet data."""

    if dashboard:
        info("Serving the easy_access dashboard.")
        info("Once launched, it will be available at http://localhost:8000.")
        info("Press Ctrl+C or close this terminal window to stop the server.")
        uvicorn.run(
            "easy_access.dashboard.dash:app",
            host="0.0.0.0",
            port=8000,
            reload=True,
        )

        cool("Done serving the dashboard. Exiting tool!")
        raise typer.Exit(code=0)

    if export:
        if single_faculty:
            info(f"Exporting data for faculty: {single_faculty}")
            create_export_sheet(faculty=single_faculty)
        else:
            info("Creating export sheets.")
            create_export_sheet()
        cool("Done creating export sheets. Exiting tool!")
        raise typer.Exit(code=0)
    backupper = Backupper()
    backup = BackupFlag(backup)
    match backup:
        case BackupFlag.BACKUP:
            backupper.backup_files()
        case BackupFlag.RESTORE:
            backupper.restore_backup(
                strategy=RestoreStrategy(restore_strategy),
                select=RestoreOptions(restore_dir),
            )
        case BackupFlag.DEFAULT:
            if SETTINGS.backup_settings.backup_all:
                backupper.backup_files()
        case BackupFlag.NONE:
            pass
        case _:
            warn("Unrecognized backup flag. Skipping backup/restore.")

    if other_sheet:
        try:
            other_sheet = Path(other_sheet)
            info(f"Reading in data from other sheet: {other_sheet.absolute()}")
        except Exception as e:
            warn(f"Failed to parse path to other sheet: {e}")
            other_sheet = None

    if any([download, deduplicate, classify]):
        info("Running the tool without writes to update db data")
        temp_tool = EasyAccessTool(
            settings=EasyAccessSettings.from_env(
                functions="read",
                only_changes=True,
                refresh_osiris_data=osiris_update,
                other_sheet=None,
                only_retrieve_missing_osiris_data=not osiris_full_refresh,
                disable_writes=True,
            )
        )
        temp_tool.run()
        cool("Done updating data without writes to excels. ")

        info(
            "Doing the rest of the preprocessing steps: download files, deduplication, and classification."
        )
        warn("currently not implemented")

    # Load settings from env and CLI params
    ea_settings = EasyAccessSettings.from_env(
        export=export,
        only_changes=changes,
        refresh_osiris_data=osiris_update,
        other_sheet=other_sheet,
        only_retrieve_missing_osiris_data=not osiris_full_refresh,
        disable_writes=disable_writes,
        faculty=single_faculty,
    )

    tool = EasyAccessTool(ea_settings)
    tool.run()

    cool("All done! Thank you for using the Easy Access tool!")


if __name__ == "__main__":
    cli_app()
