# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "bs4",
#     "python-dotenv",
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
#     "nameparser",
#     "levenshtein",
#     "selenium",
#     "docling",
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
from easy_access.utils import cool, warn
from easy_access.settings import Functions, EasyAccessSettings, SETTINGS
from easy_access.backup import Backupper, BackupFlag, RestoreOptions, RestoreStrategy
from pathlib import Path
from easy_access.main import EasyAccessTool
from easy_access.downloader import Downloader

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
    save: Annotated[
        bool,
        typer.Option(
            help="If enabled, will store results in excel files. If disabled will only print to console.",
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
    retrieve_all: Annotated[
        bool,
        typer.Option(
            help="Retrieve all data from data entry folders and store as parquet file.",
            rich_help_panel="Backup/Restore",
        ),
    ] = True,
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
    ] = False
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

    # Load settings from env and CLI params
    ea_settings = EasyAccessSettings.from_env(
        functions=do,
        only_changes=changes,
        save_files=save,
        refresh_osiris_data=osiris_update,
        retrieve_all=retrieve_all,
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
        print(f'downloading files')
        downloader = Downloader()
        driver = downloader.download_pdfs()
        print(f'done downloading, waiting for download to finish...')
        time.sleep(5)
        if driver:
            driver.close()
    else:
        tool.run()

    cool("All done! Thank you for using the Easy Access tool!")

if __name__ == "__main__":
    cli_app()
