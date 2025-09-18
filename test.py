import asyncio
from pathlib import Path

import polars as pl

from easy_access.db.base import close_connections, ensure_db_inited
from easy_access.db.update import (
    map_v1_to_v2_classifications,
    update_workflow_status_from_db,
)
from easy_access.maintenance.v1_items import (
    import_v1_sheets,
    ingest_v1_items,
    match_v1_to_copyright_items,
    merge_all_sheets,
    process_v1_sheets,
)
from easy_access.settings import Settings


async def test_db_updates():
    settings = Settings()
    await ensure_db_inited(settings)
    await update_workflow_status_from_db(settings)
    await map_v1_to_v2_classifications(settings)
    await update_workflow_status_from_db(settings)
    await close_connections()


async def matcher():
    settings = Settings()
    await ensure_db_inited(settings)
    await match_v1_to_copyright_items(settings)
    await close_connections()


async def store_data(final_results: dict[str, pl.DataFrame]):
    settings = Settings()
    await ensure_db_inited(settings)
    for faculty, df in final_results.items():
        await ingest_v1_items(df)

    await close_connections()


asyncio.run(test_db_updates())

if False:
    base_dir = Path("E:/ea-cli/faculty_sheets_from_teams_sept_2025")
    individual_results = {}
    processed_results = {}
    final_results: dict[str, pl.DataFrame] = {}
    for subdir in base_dir.iterdir():
        if subdir.is_dir():
            print(f"Processing faculty directory: {subdir.name}")

            result = import_v1_sheets(subdir)
            individual_results[subdir.name] = result
            final = process_v1_sheets(result)
            processed_results[subdir.name] = final
            main = merge_all_sheets(final)
            final_results[subdir.name] = main

    asyncio.run(store_data(final_results))
