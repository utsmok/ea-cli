from easy_access.utilities.file_exists import check_file_exists
from easy_access.api_keys import canvas as api_token
import polars as pl
from rich import print
import asyncio

from easy_access.db.retrieve import retrieve_tortoise_copyright_items
from easy_access.settings import SETTINGS
async def main():
    df = pl.read_excel("C:\\dev\\ea-cli\\full_backups\\backup_2025-03-06_11-06-04\\ET\\ET_2024-11-06.xlsx", sheet_name="Complete data")

    result_df = await check_file_exists(
                    api_token=api_token,
                    df=df,
                    excel_file=None,
                    sheetname="Complete data",
                    url_col="url"
        )

    print(f"input len: {df.shape[0]}")
    print(f"result len: {result_df.shape[0]}")

    print(f'Results:')
    result_data = result_df.select(['title', 'material_id', 'url', 'file_exists']).to_dicts()

    # replace 'title'== None with untitled
    new_results = [{k: (v if v is not None else '--untitled--') for k, v in row.items()} for row in result_data]
    # sort
    new_results.sort(key=lambda x: (x['file_exists'], x['title']), reverse=True)
    for row in new_results:

        if row['file_exists']:
            prefix = f"✅ [green]"
            suffix = "[/green]"

        else:
            prefix = f"❌ [red]"
            suffix = "[/red]"

        print(f"  {prefix} {row['title']} {suffix}")
        print(f"        ↳ {row['material_id']} | {row['url']}")

    material_ids = [row['material_id'] for row in new_results]
    items = await retrieve_tortoise_copyright_items(material_ids=material_ids, settings=SETTINGS)

    for item in items:
        if item.misaligned_status():
            print(item.actual_status(), item.status, item.file_exists, item.last_canvas_check, item.material_id, item.filename)

async def test_misalign():
    items = await retrieve_tortoise_copyright_items(settings=SETTINGS)
    details = []
    for item in items:
        if item.misaligned_status():
            details.append(item.status_details())

    print(f'Got {len(details)} misaligned items:')
    for detail in details:
        print(detail)
asyncio.run(test_misalign())