"""
Easy Access Sheet Toolkit
Sept 2025
Samuel Mok / s.mok@utwente.nl / cip@utwente.nl
homepage: https://github.com/utsmok/ea-cli
Note: only tested on windows systems

This script runs the Easy Access tool for you.
All code can be found in folder 'easy_access'
See readme.md for more info, and the settings.yaml example file for specific parameters.

quickstart:
1. install uv (https://docs.astral.sh/uv/getting-started/installation/)
2. run the command `uv run run.py --help`
"""

import asyncio
from pathlib import Path
from typing import Annotated

import typer
from loguru import logger

from easy_access.db.base import close_connections
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
            return _orig_make_metavar(self, ctx)  # type: ignore

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
    verbose: Annotated[
        bool,
        typer.Option(
            help="Enable verbose error logging.",
            rich_help_panel="Options",
        ),
    ] = False,
    changes: Annotated[
        bool,
        typer.Option(
            help="[DEPRECATED?] Only add items that have been changed to new faculty sheets.",
            rich_help_panel="Deprecated",
        ),
    ] = True,
    osiris_update: Annotated[
        bool,
        typer.Option(
            help="If enabled, will retrieve fresh osiris data for all course + people page data.",
            rich_help_panel="Enrichment",
        ),
    ] = True,
    osiris_full_refresh: Annotated[
        bool,
        typer.Option(
            help="If osiris_update is enabled, this flag will toggle retrieval of fresh osiris data for either ALL data, or only data currently missing osiris info.",
            rich_help_panel="Enrichment",
        ),
    ] = False,
    input_sheet: Annotated[
        Path | None,
        typer.Option(
            help="Path to a xlsx sheet with raw CopyRight Data to read in. If not provided, will use the most recent sheet in the raw_copyright_data/ folder.",
            rich_help_panel="Processing",
            exists=True,
            file_okay=True,
            dir_okay=False,
        ),
    ] = None,
    write_xlsx_files: Annotated[
        bool,
        typer.Option(
            help="Disable all xlsx write operations, only update the sql db.",
            rich_help_panel="Options",
        ),
    ] = True,
    single_faculty: Annotated[
        str | None,
        typer.Option(
            help="Only run the tool for a single faculty. use the faculty abbreviation as the parameter (e.g. 'BMS').",
            rich_help_panel="Processing",
        ),
    ] = None,
    # Stage selection options
    ingest_only: Annotated[
        bool,
        typer.Option(
            help="Only run the data ingestion stages (raw data and faculty updates).",
            rich_help_panel="Stage",
        ),
    ] = False,
    process_only: Annotated[
        bool,
        typer.Option(
            help="Only run the data processing stage (update database from staged data).",
            rich_help_panel="Stage",
        ),
    ] = False,
    export_only: Annotated[
        bool,
        typer.Option(
            help="Only run the export stage (generate faculty sheets).",
            rich_help_panel="Stage",
        ),
    ] = False,
    enrich_only: Annotated[
        bool,
        typer.Option(
            help="Only run the enrichment stage (fetch OSIRIS data).",
            rich_help_panel="Stage",
        ),
    ] = False,
    file_exists_only: Annotated[
        bool,
        typer.Option(
            help="Only run the file existence verification stage.",
            rich_help_panel="Stage",
        ),
    ] = False,
    pdf_download_only: Annotated[
        bool,
        typer.Option(
            help="Only run the PDF downloading stage.",
            rich_help_panel="Stage",
        ),
    ] = False,
    parse_only: Annotated[
        bool,
        typer.Option(
            help="Only run the PDF parsing stage.",
            rich_help_panel="Stage",
        ),
    ] = False,
    file_exists: Annotated[
        bool,
        typer.Option(
            help="Enable file existence verification stage.",
            rich_help_panel="Enrichment",
        ),
    ] = True,
    pdf_download: Annotated[
        bool,
        typer.Option(
            help="Enable PDF downloading stage.",
            rich_help_panel="Enrichment",
        ),
    ] = False,
    parse_pdf: Annotated[
        bool,
        typer.Option(
            help="Enable PDF parsing stage.",
            rich_help_panel="Enrichment",
        ),
    ] = False,
    ingest: Annotated[
        bool,
        typer.Option(
            help="Enable ingestion of data from sheets (raw crc data / faculty sheets) when running the workflow.",
            rich_help_panel="Processing",
        ),
    ] = True,
) -> None:
    """Runs the main Easy Access data processing workflow. See --help for all options. Use --verbose to enable verbose error logging."""

    # Validate and parse input sheet path
    if input_sheet:
        try:
            input_sheet = Path(input_sheet)  # Ensure it's a Path object
            logger.info(f"Reading in data from input sheet: {input_sheet.absolute()}")
        except Exception as e:
            logger.warning(f"Failed to parse path to input sheet: {e}")
            input_sheet = None

    # STAGE SELECTION

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

    # INIT TOOL

    from easy_access.main import EasyAccessTool
    from easy_access.settings import SETTINGS, EasyAccessSettings

    # create settings object for this run
    # TODO: this logic is getting messy and also duplicated here+main.py+pipeline.py
    # clean / refactor this later to disentangle settings vs runtime options better
    ea_settings = EasyAccessSettings.create_for_runtime(
        main_settings=SETTINGS,
        export=export_only or not any(stage_options),
        only_changes=changes,
        refresh_osiris_data=osiris_update,
        other_sheet=input_sheet,
        only_retrieve_missing_osiris_data=not osiris_full_refresh,
        disable_writes=not write_xlsx_files,
        faculty=single_faculty,
        no_file_exists=not file_exists,
        export_workflow=True,
        no_pdf_download=not pdf_download,
        no_pdf_parse=parse_pdf,
    )
    tool = EasyAccessTool(settings_obj=SETTINGS, ea_settings=ea_settings)

    # Determine which stages to run based on options

    stages = {
        "ingest": {
            "enabled": (ingest_only or not any(stage_options)) and ingest,
            "func": tool.run_ingest,
        },
        "process": {
            "enabled": process_only or not any(stage_options),
            "func": tool.run_process,
        },
        "enrich": {
            "enabled": enrich_only or not any(stage_options),
            "func": tool.run_enrich,
        },
        "file_exists": {
            "enabled": (file_exists_only or not any(stage_options)) and file_exists,
            "func": tool.run_verify_file_existence,
        },
        "pdf_download": {
            "enabled": (pdf_download_only or not any(stage_options)) and pdf_download,
            "func": tool.run_download_pdfs,
        },
        "parse": {
            "enabled": (parse_only or not any(stage_options)) and parse_pdf,
            "func": tool.run_parse_pdfs,
        },
        "db_changes": {
            "enabled": not any(stage_options),
            "func": tool.process_db_changes,
        },
        "export": {
            "enabled": (export_only or not any(stage_options)) and write_xlsx_files,
            "func": tool.run_export,
        },
    }

    # now we run the selected stages, wrapped in try/except to catch errors and log them nicely
    # ending with a finally block to always close the async db connections to prevent hanging processes
    # if verbose is enabled, we log full tracebacks, otherwise just the error message
    # if a critical error occurs in ingest/process/db_changes, we abort the workflow
    # otherwise we skip the failed stage and continue with the next one
    try:
        for stage_name, stage_info in stages.items():
            if stage_info["enabled"]:
                try:
                    logger.info(f"Running {stage_name} stage...")
                    stage_info["func"]()
                    logger.success(f"{stage_name} stage complete.")
                except Exception as e:
                    logger.error(f"Error during {stage_name} stage: {e}")
                    if not verbose:
                        logger.info(
                            "Non-verbose mode active. For more error details, rerun with the verbose flag enabled."
                        )
                    else:
                        logger.exception(e)
                    if stage_name in ["ingest", "process", "db_changes"]:
                        logger.critical(
                            "Critical error in core stage. Aborting further processing."
                        )
                        tool.close_connections()
                        typer.Exit(1)
                    else:
                        logger.warning(
                            f"Non-critical stage, skipping {stage_name} and moving on."
                        )
            else:
                logger.warning(f"{stage_name} stage disabled. Skipping!")
    except Exception as e:
        logger.critical(f"Critical error in workflow: {e}")
    finally:
        tool.close_connections()
        logger.success("Tool successfully closed.")
        typer.Exit(1)


@app.command(name="update-from-v1")
def update_from_v1(
    path: Annotated[
        Path, typer.Option(help="Path to the dir containing the v1 faculty sheets.")
    ] = Path("faculty_sheets_old/"),
) -> None:
    """Updates the database from Easy Access V1 faculty sheets, focusing on ingesting the old manual classifications."""

    from easy_access.db.relations import match_v1_to_copyright_items
    from easy_access.db.update import (
        map_v1_to_v2_classifications,
        update_workflow_status_from_db,
    )
    from easy_access.maintenance.v1_items import add_v1_hashes, ingest_v1_data

    async def run_ingest_pipeline(settings: Settings, path: Path):
        await ingest_v1_data(settings, path)
        await add_v1_hashes(settings)
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
