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

admin_app = typer.Typer(
    name="admin", help="Administrative operations and failure management."
)
app.add_typer(admin_app)


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
    no_file_exists: Annotated[
        bool,
        typer.Option(
            help="Skip file existence verification stage.",
            rich_help_panel="Processing Options",
        ),
    ] = False,
    export_workflow: Annotated[
        bool,
        typer.Option(
            help="Enable new workflow-based exporter (writes inbox/in_progress/done per faculty).",
            rich_help_panel="Stage Selection",
        ),
    ] = False,
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
    ]
    if sum(stage_options) > 1:
        logger.error(
            "Cannot specify multiple stage options. Choose only one: --ingest-only, --process-only, --export-only, --enrich-only, or --file-exists-only."
        )
        typer.Exit(1)

    # Determine which stages to run
    run_ingest = ingest_only or not any(
        stage_options
    )  # Default to all if no stage specified
    run_process = process_only or not any(stage_options)
    run_export = export_only or not any(stage_options)
    run_enrich = enrich_only or not any(stage_options)
    run_file_exists = file_exists_only or not any(stage_options)

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
        export_workflow=export_workflow,
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
    if run_file_exists and not no_file_exists:
        logger.info("Running file existence verification stage...")
        tool.run_verify_file_existence()
    if run_export:
        logger.info("Running export stage...")
        tool.run_export()

    logger.success("Processing done!")
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


# Commands for admin app


@admin_app.command(name="inspect-failures")
def inspect_failures(
    limit: Annotated[int, typer.Option(help="Limit number of records to show")] = 50,
    material_id: Annotated[
        int | None, typer.Option(help="Filter by specific material ID")
    ] = None,
    show_payload: Annotated[
        bool, typer.Option(help="Show full staged payload in output")
    ] = False,
) -> None:
    """Inspect StagedProcessingFailure records for debugging."""
    import asyncio
    import json

    from tortoise import Tortoise

    from easy_access.db.models import StagedProcessingFailure
    from easy_access.settings import SETTINGS

    async def run_inspect():
        await Tortoise.init(
            db_url=f"sqlite://{SETTINGS.db_path}",
            modules={"models": ["easy_access.db.models"]},
        )

        try:
            query = StagedProcessingFailure.all().order_by("-created_at")

            if material_id:
                query = query.filter(material_id=material_id)

            failures = await query.limit(limit)

            if failures:
                typer.echo(f"\nFound {len(failures)} failure records:")
                typer.echo("-" * 80)

                for failure in failures:
                    typer.echo(f"ID: {failure.id}")
                    typer.echo(f"Material ID: {failure.material_id}")
                    typer.echo(f"Created: {failure.created_at}")
                    typer.echo(f"Error: {failure.error_message}")

                    if show_payload and failure.staged_payload:
                        typer.echo(
                            f"Payload: {json.dumps(failure.staged_payload, indent=2)}"
                        )

                    typer.echo("-" * 80)
            else:
                typer.echo("No failure records found.")
        finally:
            await Tortoise.close_connections()

    asyncio.run(run_inspect())
    typer.Exit()


@admin_app.command(name="failure-stats")
def failure_stats() -> None:
    """Show statistics about StagedProcessingFailure records."""
    import asyncio
    from datetime import UTC, datetime, timedelta

    from tortoise import Tortoise

    from easy_access.db.models import StagedProcessingFailure
    from easy_access.settings import SETTINGS

    def categorize_error(error_message: str) -> str:
        """Categorize error messages into common patterns."""
        error_lower = error_message.lower()

        if "faculty" in error_lower and (
            "not found" in error_lower or "does not exist" in error_lower
        ):
            return "Faculty Lookup Error"
        elif "material_id" in error_lower and (
            "invalid" in error_lower or "missing" in error_lower
        ):
            return "Invalid Material ID"
        elif "classification" in error_lower:
            return "Classification Error"
        elif "database" in error_lower or "connection" in error_lower:
            return "Database Error"
        elif "permission" in error_lower or "access" in error_lower:
            return "Permission Error"
        elif "timeout" in error_lower:
            return "Timeout Error"
        elif "validation" in error_lower:
            return "Validation Error"
        else:
            return "Other Error"

    async def run_stats():
        await Tortoise.init(
            db_url=f"sqlite://{SETTINGS.db_path}",
            modules={"models": ["easy_access.db.models"]},
        )

        try:
            total_failures = await StagedProcessingFailure.all().count()

            # Group by error patterns
            failures = await StagedProcessingFailure.all()
            error_patterns = {}

            for failure in failures:
                if failure.error_message:
                    error_key = categorize_error(failure.error_message)
                    error_patterns[error_key] = error_patterns.get(error_key, 0) + 1

            # Get failures by material_id
            material_failures = await StagedProcessingFailure.filter(
                material_id__not_isnull=True
            ).count()
            unknown_material_failures = await StagedProcessingFailure.filter(
                material_id__isnull=True
            ).count()

            # Get recent failures (last 24 hours)
            yesterday = datetime.now(UTC) - timedelta(days=1)
            recent_failures = await StagedProcessingFailure.filter(
                created_at__gte=yesterday
            ).count()

            typer.echo("\nFailure Statistics:")
            typer.echo("-" * 40)
            typer.echo(f"Total failures: {total_failures}")
            typer.echo(f"Failures with material_id: {material_failures}")
            typer.echo(f"Failures without material_id: {unknown_material_failures}")
            typer.echo(f"Recent failures (24h): {recent_failures}")

            if error_patterns:
                typer.echo("\nError Patterns:")
                for pattern, count in sorted(
                    error_patterns.items(), key=lambda x: x[1], reverse=True
                ):
                    typer.echo(f"  {pattern}: {count}")
        finally:
            await Tortoise.close_connections()

    asyncio.run(run_stats())
    typer.Exit()


@admin_app.command(name="retry-failures")
def retry_failures(
    material_id: Annotated[
        int | None, typer.Option(help="Retry specific material ID only")
    ] = None,
    dry_run: Annotated[
        bool, typer.Option(help="Show what would be done without making changes")
    ] = True,
) -> None:
    """Retry processing of failed StagedProcessingFailure records."""
    import asyncio

    from tortoise import Tortoise

    from easy_access.db.models import (
        StagedCopyrightItem,
        StagedFacultyUpdate,
        StagedProcessingFailure,
    )
    from easy_access.db.update import (
        process_staged_faculty_updates,
        process_staged_raw_data,
    )
    from easy_access.settings import SETTINGS

    async def retry_single_failure(failure):
        """Retry processing a single failure."""
        try:
            material_id_val = failure.material_id
            payload = failure.staged_payload

            if not payload or not material_id_val:
                return False

            # Check if the original staged record still exists
            staged_raw = await StagedCopyrightItem.filter(
                material_id=material_id_val
            ).first()
            staged_faculty = await StagedFacultyUpdate.filter(
                material_id=material_id_val
            ).first()

            if staged_raw:
                # Retry raw data processing
                if not dry_run:
                    await process_staged_raw_data(SETTINGS)
                return True
            elif staged_faculty:
                # Retry faculty update processing
                if not dry_run:
                    await process_staged_faculty_updates(SETTINGS)
                return True
            else:
                typer.echo(
                    f"Warning: No staged record found for material_id {material_id_val}"
                )
                return False

        except Exception as e:
            typer.echo(
                f"Error retrying failure for material_id {failure.material_id}: {str(e)}"
            )
            return False

    async def run_retry():
        await Tortoise.init(
            db_url=f"sqlite://{SETTINGS.db_path}",
            modules={"models": ["easy_access.db.models"]},
        )

        try:
            query = StagedProcessingFailure.all()

            if material_id:
                query = query.filter(material_id=material_id)

            failures = await query

            successful_retries = 0
            failed_retries = 0
            errors = []

            for failure in failures:
                try:
                    if failure.staged_payload and failure.material_id:
                        # Try to reprocess the staged data
                        success = await retry_single_failure(failure)
                        if success:
                            successful_retries += 1
                            if not dry_run:
                                await failure.delete()  # Remove successful retry
                        else:
                            failed_retries += 1
                            errors.append(
                                f"Failed to retry material_id {failure.material_id}"
                            )
                    else:
                        failed_retries += 1
                        errors.append(
                            f"Missing payload or material_id for failure {failure.id}"
                        )

                except Exception as e:
                    failed_retries += 1
                    errors.append(f"Error retrying failure {failure.id}: {str(e)}")

            typer.echo(f"\nRetry Results ({'DRY RUN' if dry_run else 'LIVE'}):")
            typer.echo("-" * 40)
            typer.echo(f"Total attempted: {len(failures)}")
            typer.echo(f"Successful retries: {successful_retries}")
            typer.echo(f"Failed retries: {failed_retries}")

            if errors:
                typer.echo("\nErrors:")
                for error in errors[:10]:  # Show first 10 errors
                    typer.echo(f"  {error}")
                if len(errors) > 10:
                    typer.echo(f"  ... and {len(errors) - 10} more errors")
        finally:
            await Tortoise.close_connections()

    asyncio.run(run_retry())
    typer.Exit()


@admin_app.command(name="cleanup-failures")
def cleanup_failures(
    days_old: Annotated[
        int, typer.Option(help="Delete records older than N days")
    ] = 30,
    dry_run: Annotated[
        bool, typer.Option(help="Show what would be done without making changes")
    ] = True,
) -> None:
    """Clean up old StagedProcessingFailure records."""
    import asyncio
    from datetime import UTC, datetime, timedelta

    from tortoise import Tortoise

    from easy_access.db.models import StagedProcessingFailure
    from easy_access.settings import SETTINGS

    async def run_cleanup():
        await Tortoise.init(
            db_url=f"sqlite://{SETTINGS.db_path}",
            modules={"models": ["easy_access.db.models"]},
        )

        try:
            cutoff_date = datetime.now(UTC) - timedelta(days=days_old)

            query = StagedProcessingFailure.filter(created_at__lt=cutoff_date)
            old_failures_count = await query.count()

            typer.echo(f"\nCleanup Results ({'DRY RUN' if dry_run else 'LIVE'}):")
            typer.echo("-" * 40)
            typer.echo(f"Cutoff date: {cutoff_date}")
            typer.echo(f"Failures to delete: {old_failures_count}")

            if not dry_run and old_failures_count > 0:
                deleted_count = await query.delete()
                typer.echo(f"Actually deleted: {deleted_count}")
            else:
                typer.echo("Actually deleted: 0")
        finally:
            await Tortoise.close_connections()

    asyncio.run(run_cleanup())
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
