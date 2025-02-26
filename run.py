# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "bs4",
#     "httpx",
#     "lxml",
#     "openpyxl",
#     "polars",
#     "rich",
#     "typer",
#     "fastexcel",
#     "xlsxwriter",
#     "pyyaml",
#     "colorama",
#     "loguru",
#     "levenshtein",
#     "selenium",
#     "google-genai",
#     "tortoise-orm[accel]",
#     "aiometer",
#     "pdfminer-six",
#     "pydantic",
#     "sqlalchemy",
#     "connectorx",
#     "pikepdf",
#     "qdrant_client",
#     "fastembed",
#      "xxhash",
# ]
# ///

"""
Easy Access Sheet Toolkit
Feb 2025
Samuel Mok / s.mok@utwente.nl / cip@utwente.nl
homepage: https://github.com/utsmok/ea-cli
Note: only tested on windows systems

This script runs the Easy Access tool for you.
All code can be found in folder 'easy_access', with easy_access_cli.py containing the main functionality.
See readme.md for more info, and the settings.yaml example file for specific parameters.

quickstart:
1. install uv (https://docs.astral.sh/uv/getting-started/installation/)
2. make sure settings.yaml is present in the same dir as run.py and the contents are correct
3. > uv run run.py --help
"""

import time
import typer
from typing import Annotated
from easy_access.utils import cool, warn, info
from easy_access.settings import Functions, EasyAccessSettings, SETTINGS
from easy_access.sheets.backup import Backupper, BackupFlag, RestoreOptions, RestoreStrategy
from pathlib import Path
from easy_access.main import EasyAccessTool
from easy_access.classification.downloader import Downloader
from easy_access.classification.pdf_handling import enrich_pdfs
from easy_access.classification.classifier_api import main
import asyncio
cli_app = typer.Typer()

@cli_app.command()
def cli(
    do: Annotated[
        Functions,
        typer.Option(
            case_sensitive=False,
            help="Which tool to run: read in new data, export current data, or both.",
            rich_help_panel="Functions",
        ),
    ] = "read",
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
    download_files: Annotated[
        bool,
        typer.Option(
            help="Download pdfs from canvas.",
            rich_help_panel="Functions",
        ),
    ] = False,
    max_amount_to_download: Annotated[
        int,
        typer.Option(
            help="Maximum amount of files to download.",
            rich_help_panel="Functions",
        ),
    ] = None,
    classify: Annotated[
        bool,
        typer.Option(
            help="Classify the pdfs by LLM.",
            rich_help_panel="Functions",
        ),
    ] = False,
    deduplicate: Annotated[
        bool,
        typer.Option(
            help="Deduplicate the pdfs.",
            rich_help_panel="Functions",
        ),
    ] = False,
) -> None:
    """Easy Access toolkit for managing faculty sheet data."""

    backupper = Backupper()
    backup = BackupFlag(backup)
    match backup:
        case BackupFlag.BACKUP:
            backupper.backup_files()
        case BackupFlag.RESTORE:
            backupper.restore_backup(strategy=RestoreStrategy(restore_strategy), select=RestoreOptions(restore_dir))
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

    # Load settings from env and CLI params
    ea_settings = EasyAccessSettings.from_env(
        functions=do,
        only_changes=changes,
        refresh_osiris_data=osiris_update,
        other_sheet=other_sheet,
        only_retrieve_missing_osiris_data=not osiris_full_refresh,
    )

    if do not in [Functions.both, Functions.read, Functions.export]:
        warn("No functions selected! Aborting. Run ea-cli --help for details.")
        cool("Thank you for using the Easy Access tool!")
        raise typer.Exit(code=1)

    # Initialize and run tool with settings
    tool = EasyAccessTool(ea_settings)
    driver = None
    if download_files:
        print(f'downloading files. Will use Chrome do so. Please make sure you have disabled all extensions, are logged in to Canvas, and have the correct permissions to download the files.\n Then completely close Chrome before continueing.')
        input(f'Press any key to continue...')
        downloader = Downloader()
        asyncio.get_event_loop().run_until_complete(downloader.download_pdfs(subset=None, max_amount=max_amount_to_download))
        print(f'done downloading, waiting for download to finish...')
        time.sleep(5)
    if deduplicate:
        print(f'Enriching/deduplicating PDFs.')
        asyncio.get_event_loop().run_until_complete(enrich_pdfs())

    if classify:
        print(f'Classifying PDFS.')
        asyncio.get_event_loop().run_until_complete(main())


    tool.run()

    cool("All done! Thank you for using the Easy Access tool!")

if __name__ == "__main__":
    cli_app()
