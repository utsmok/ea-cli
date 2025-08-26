from easy_access.utilities.file_exists import check_file_exists
from easy_access.api_keys import canvas as api_token
import polars as pl
from rich import print
df = pl.read_excel("C:\\dev\\ea-cli\\full_backups\\backup_2025-03-06_11-06-04\\ET\\ET_2024-11-06.xlsx", sheet_name="Complete data")

result_df = check_file_exists(
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