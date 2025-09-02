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

# Delay importing project modules that may perform work at import-time.
# Import them inside the command functions to avoid side-effects when
# the module is imported just to show --help.
from loguru import logger

# Compatibility shim: some combinations of Typer and Click/Rich have a
# small signature mismatch where Typer's rich help calls
# `param.make_metavar()` without providing the `ctx` argument while newer
# Click versions require `ctx`. Detect that case at runtime and wrap
# `click.Parameter.make_metavar` so it accepts an optional `ctx` (default
# None). This is a minimal, local compatibility fix that avoids editing
# third-party packages or forcing package downgrades.
try:
    import inspect

    import click

    _orig_make_metavar = click.Parameter.make_metavar
    _sig = inspect.signature(_orig_make_metavar)
    _params = list(_sig.parameters.values())
    # If the original has a required ctx parameter (positional, no default),
    # replace it with a thin wrapper that provides a default None so calls
    # without ctx don't raise a TypeError.
    if len(_params) >= 2 and _params[1].default is inspect._empty:

        def _make_metavar_compat(self, ctx=None):
            return _orig_make_metavar(self, ctx)

        click.Parameter.make_metavar = _make_metavar_compat
except Exception:
    # If anything goes wrong here, fall back to normal behavior.
    pass

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

    # Import project modules here to avoid import-time side-effects when showing --help
    from easy_access.main import EasyAccessTool
    from easy_access.settings import SETTINGS, EasyAccessSettings

    # Load settings from CLI params, using main SETTINGS for base dir config
    ea_settings = EasyAccessSettings.create_for_runtime(
        main_settings=SETTINGS,
        export=False,
        only_changes=changes,
        refresh_osiris_data=osiris_update,
        other_sheet=other_sheet,
        only_retrieve_missing_osiris_data=not osiris_full_refresh,
        disable_writes=disable_writes,
        faculty=single_faculty,
    )

    tool = EasyAccessTool(settings_obj=SETTINGS, ea_settings=ea_settings)
    tool.run()

    logger.success("Main processing done!")
    typer.Exit()


@app.command(name="dashboard")
def run_dashboard(
    port: Annotated[int, typer.Option(help="Port to serve the dashboard on.")] = 8000,
    host: Annotated[
        str, typer.Option(help="Host to serve the dashboard on.")
    ] = "0.0.0.0",
) -> None:
    """Serves the easy_access dashboard."""
    logger.info("Serving the easy_access dashboard.")
    logger.info(f"Once launched, it will be available at http://{host}:{port}.")
    logger.info("Press Ctrl+C or close this terminal window to stop the server.")
    from easy_access.settings import SETTINGS

    uvicorn.run(
        "dashboard.dash:app",
        host=host,
        port=port,
        reload=SETTINGS.dashboard_reload,
    )
    logger.success("Dashboard server stopped.")
    typer.Exit()


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
    from easy_access.settings import SETTINGS
    from easy_access.sheets.sheet import create_export_sheet

    if single_faculty:
        logger.info(f"Exporting data for faculty: {single_faculty}")
        create_export_sheet(settings=SETTINGS, faculty=single_faculty)
    else:
        logger.info("Creating export sheets for all faculties.")
        create_export_sheet(settings=SETTINGS)
    logger.success("Done creating export sheets.")
    typer.Exit()


# Commands for backup app


@backup_app.command(name="create")
def create_backup_command() -> None:
    """Creates a backup of the current data based on settings.yaml."""
    from easy_access.settings import SETTINGS
    from easy_access.sheets.backup import Backupper

    backupper = Backupper()
    if SETTINGS.backup_settings.backup_all:
        logger.info("Creating backup as per settings.yaml (backup_all: true).")
        backupper.backup_files()
    else:
        logger.info(
            "Backup not created as per settings.yaml (backup_all: false or not set)."
        )
    typer.Exit()


@backup_app.command(name="restore")
def restore_backup_command(
    restore_dir: Annotated[
        str,
        typer.Option(
            help="Set which backup to restore. Options: 'latest','oldest','manual'"
        ),
    ] = "latest",
    restore_strategy: Annotated[
        str,
        typer.Option(
            help="Set the strategy for restoring the backup. Options: 'replace','merge_prefer_existing','merge_prefer_backup'"
        ),
    ] = "replace",
) -> None:
    """Restores data from a backup."""
    from easy_access.sheets.backup import Backupper, RestoreOptions, RestoreStrategy

    # Map string inputs to enum values
    try:
        select_enum = RestoreOptions(restore_dir)
    except Exception:
        logger.warning(
            f"Invalid restore option '{restore_dir}', defaulting to 'latest'."
        )
        select_enum = RestoreOptions.LATEST

    try:
        strategy_enum = RestoreStrategy(restore_strategy)
    except Exception:
        logger.warning(
            f"Invalid restore strategy '{restore_strategy}', defaulting to 'replace'."
        )
        strategy_enum = RestoreStrategy.REPLACE

    backupper = Backupper()
    logger.info(
        f"Restoring backup from '{select_enum.value}' with strategy '{strategy_enum.value}'."
    )
    backupper.restore_backup(
        strategy=strategy_enum,
        select=select_enum,
    )
    logger.success("Backup restoration process finished.")
    typer.Exit()


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
    ] = False,
    dry_run: Annotated[
        bool,
        typer.Option(
            help="If enabled, will run the tool in dry-run mode to update the DB without changing .xlsx files.",
        ),
    ] = True,
    single_faculty: Annotated[
        str | None,
        typer.Option(
            help="Only run the tool for a single faculty. use the faculty abbreviation as the parameter (e.g. 'BMS').",
        ),
    ] = None,
    download: Annotated[
        bool,
        typer.Option(help="Download pdfs from canvas."),
    ] = False,
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
    logger.info(
        "First, running the tool in read-only mode to update DB data if needed."
    )

    import asyncio

    from classification.httpx_downloader import main_download_all
    from easy_access.db.retrieve import retrieve_unmarked_deleted_items
    from easy_access.main import EasyAccessTool
    from easy_access.settings import SETTINGS, EasyAccessSettings

    ea_temp_settings = EasyAccessSettings.create_for_runtime(
        main_settings=SETTINGS,
        export=False,
        only_changes=True,
        refresh_osiris_data=osiris_update,
        other_sheet=None,
        only_retrieve_missing_osiris_data=not osiris_full_refresh,
        disable_writes=True,
        faculty=single_faculty,
    )
    if dry_run:
        temp_tool = EasyAccessTool(settings_obj=SETTINGS, ea_settings=ea_temp_settings)
        temp_tool.run()
        logger.success("Done updating data in read-only mode for pre-processing.")

        logger.info(
            "Actual pre-processing steps (download, deduplicate, classify) follow."
        )
        logger.warning(
            "deduplication and classification steps are currently stubs and not implemented."
        )
    if download:
        logger.info("Downloading PDFs...")
        downloaded, failed = asyncio.run(
            main_download_all(settings=SETTINGS, max_concurrent=15)
        )
        logger.info(f"\nDownloaded Files ({len(downloaded)})")
        logger.info(f"Failed Files ({len(failed)})")

    # now use retrieve_unmarked_deleted_items function to see if any failed downloads correspond to unmarked deleted items
    failed_items = asyncio.run(retrieve_unmarked_deleted_items(settings=SETTINGS))
    if failed_items:
        logger.info(
            f"Found {len(failed_items)} failed downloads corresponding to unmarked deleted items."
        )
        skip = 0
        for item in failed_items:
            if not item.manual_classification:
                skip += 1
                continue
            else:
                logger.info(
                    f" - {item.material_id} | {item.last_change} | {item.title} | {item.status} | {item.remarks} | {item.manual_classification}"
                )
        logger.info(
            f"Skipped {skip}/{len(failed_items)} items without manual classification."
        )

    # if deduplicate:
    #     logger.info("Deduplicating PDFs...")
    #     # ... deduplicator logic ...
    # if classify:
    #     logger.info("Classifying PDFs...")
    #     # ... classifier logic ...
    logger.success("Pre-processing steps finished")
    typer.Exit()


if __name__ == "__main__":
    app()
