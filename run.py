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

import asyncio
from pathlib import Path
from typing import Annotated

import typer
from loguru import logger

from easy_access.db.base import close_connections
from easy_access.db.update import (
    map_v1_to_v2_classifications,
    update_workflow_status_from_db,
)
from easy_access.maintenance.v1_items import match_v1_to_copyright_items
from easy_access.settings import Settings

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


app = typer.Typer(
    name="ea-cli",
    help="Easy Access toolkit for the University of Twente.",
    add_completion=False,
)


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
    # Stage selection options
    ingest_only: Annotated[
        bool,
        typer.Option(
            help="Only run the data ingestion stages (raw data and faculty updates).",
            rich_help_panel="Stage Selection",
        ),
    ] = False,
    process_only: Annotated[
        bool,
        typer.Option(
            help="Only run the data processing stage (update database from staged data).",
            rich_help_panel="Stage Selection",
        ),
    ] = False,
    export_only: Annotated[
        bool,
        typer.Option(
            help="Only run the export stage (generate faculty sheets).",
            rich_help_panel="Stage Selection",
        ),
    ] = False,
    enrich_only: Annotated[
        bool,
        typer.Option(
            help="Only run the enrichment stage (fetch OSIRIS data).",
            rich_help_panel="Stage Selection",
        ),
    ] = False,
    file_exists_only: Annotated[
        bool,
        typer.Option(
            help="Only run the file existence verification stage.",
            rich_help_panel="Stage Selection",
        ),
    ] = False,
    pdf_download_only: Annotated[
        bool,
        typer.Option(
            help="Only run the PDF downloading stage.",
            rich_help_panel="Stage Selection",
        ),
    ] = False,
    parse_only: Annotated[
        bool,
        typer.Option(
            help="Only run the PDF parsing stage.",
            rich_help_panel="Stage Selection",
        ),
    ] = False,
    no_file_exists: Annotated[
        bool,
        typer.Option(
            help="Skip file existence verification stage.",
            rich_help_panel="Processing Options",
        ),
    ] = False,
    no_pdf_download: Annotated[
        bool,
        typer.Option(
            help="Skip PDF downloading stage.",
            rich_help_panel="Processing Options",
        ),
    ] = False,
    no_pdf_parse: Annotated[
        bool,
        typer.Option(
            help="Skip PDF parsing stage.",
            rich_help_panel="Processing Options",
        ),
    ] = False,
    no_ingest: Annotated[
        bool,
        typer.Option(
            help="Skip ingestion of data from sheets (raw crc data / faculty sheets) when running the workflow.",
            rich_help_panel="Processing Options",
        ),
    ] = False,
    new_workflow: Annotated[
        bool,
        typer.Option(
            help="Enable new workflow-based exporter (writes inbox/in_progress/done per faculty).",
            rich_help_panel="Stage Selection",
        ),
    ] = True,
) -> None:
    """Runs the main Easy Access data processing workflow."""
    if other_sheet:
        try:
            other_sheet = Path(other_sheet)  # Ensure it's a Path object
            logger.info(f"Reading in data from other sheet: {other_sheet.absolute()}")
        except Exception as e:
            logger.warning(f"Failed to parse path to other sheet: {e}")
            other_sheet = None

    # Validate stage selection options
    stage_options = [
        ingest_only,
        process_only,
        export_only,
        enrich_only,
        file_exists_only,
        pdf_download_only,
    ]
    if sum(stage_options) > 1:
        logger.error(
            "Cannot specify multiple stage options. Choose only one: --ingest-only, --process-only, --export-only, --enrich-only, or --file-exists-only."
        )
        typer.Exit(1)

    # Determine which stages to run
    run_ingest = (ingest_only or not any(stage_options)) and not no_ingest
    run_process = process_only or not any(stage_options)
    run_export = export_only or not any(stage_options)
    run_enrich = enrich_only or not any(stage_options)
    run_file_exists = file_exists_only or not any(stage_options)
    run_pdf_download = pdf_download_only or not any(stage_options)
    run_parse = parse_only or not any(stage_options)
    run_db_changes = not any(stage_options)

    # Import project modules here to avoid import-time side-effects when showing --help
    from easy_access.main import EasyAccessTool
    from easy_access.settings import SETTINGS, EasyAccessSettings

    # Load settings from CLI params, using main SETTINGS for base dir config
    ea_settings = EasyAccessSettings.create_for_runtime(
        main_settings=SETTINGS,
        export=run_export,
        only_changes=changes,
        refresh_osiris_data=osiris_update,
        other_sheet=other_sheet,
        only_retrieve_missing_osiris_data=not osiris_full_refresh,
        disable_writes=disable_writes,
        faculty=single_faculty,
        no_file_exists=no_file_exists,
        export_workflow=new_workflow,
        no_pdf_download=no_pdf_download,
        no_pdf_parse=no_pdf_parse,
    )

    tool = EasyAccessTool(settings_obj=SETTINGS, ea_settings=ea_settings)

    # Run selected stages
    if run_ingest:
        logger.info("Running ingest stages...")
        tool.run_ingest()
    if run_process:
        logger.info("Running process stage...")
        tool.run_process()
    if run_enrich:
        logger.info("Running enrichment stage...")
        tool.run_enrich()
    tool.run_relations()
    if run_file_exists and not no_file_exists:
        logger.info("Running file existence verification stage...")
        tool.run_verify_file_existence()
    if run_pdf_download and not no_pdf_download:
        logger.info("Running PDF downloading stage...")
        tool.run_download_pdfs()
    if run_parse and not no_pdf_parse:
        logger.info("Running PDF parsing stage...")
        tool.run_parse_pdfs()
    if run_db_changes:
        logger.info("Processing DB changes...")
        tool.process_db_changes()
    if run_export:
        logger.info("Running export stage...")
        tool.run_export()

    logger.success("Processing done!")
    typer.Exit()


@app.command(name="update-from-v1")
def update_from_v1(
    path: Annotated[
        Path, typer.Option(help="Path to the dir containing the v1 faculty sheets.")
    ] = Path("faculty_sheets/"),
) -> None:
    """Updates the database from Easy Access V1 faculty sheets, focusing on ingesting the old manual classifications."""

    from easy_access.maintenance.v1_items import ingest_v1_data

    async def run_ingest_pipeline(settings: Settings, path: Path):
        await ingest_v1_data(settings, path)
        await match_v1_to_copyright_items(settings)
        await map_v1_to_v2_classifications(settings)
        await update_workflow_status_from_db(settings)
        await close_connections()

    settings = Settings()
    logger.info(
        f"Running v1 to v2 update pipeline on sheets in: {path.resolve()}, {type(path)}"
    )
    asyncio.run(run_ingest_pipeline(settings, path))
    logger.success("v1 to v2 update pipeline completed.")
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
    import uvicorn

    from easy_access.settings import SETTINGS

    uvicorn.run(
        "dashboard.dash:app",
        host=host,
        port=port,
        reload=SETTINGS.dashboard_reload,
    )
    logger.success("Dashboard server stopped.")
    typer.Exit()


if __name__ == "__main__":
    # Try to initialize Trogon TUI if available
    try:
        from trogon.typer import init_tui

        init_tui(app)
    except ImportError:
        # Trogon not installed, continue with regular CLI
        pass

    app()
