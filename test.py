import asyncio
from pathlib import Path

import polars as pl

from easy_access.db.base import close_connections, ensure_db_inited
from easy_access.maintenance.v1_items import (
    import_v1_sheets,
    ingest_v1_items,
    merge_all_sheets,
    process_v1_sheets,
)
from easy_access.settings import Settings


async def store_data(final_results: dict[str, pl.DataFrame]):
    settings = Settings()
    await ensure_db_inited(settings)
    for faculty, df in final_results.items():
        await ingest_v1_items(df)

    await close_connections()


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
