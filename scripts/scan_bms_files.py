from pathlib import Path

import polars as pl
from rich.console import Console

console = Console()

input_material_ids: list[int] = [
    17294859,
    17293817,
    17271751,
    17272958,
    17271997,
    17277356,
    17301720,
    17272061,
    17293934,
    17272142,
    17287736,
    17272177,
    17274948,
    17271769,
    17294832,
    17271774,
    20360379,
    17271742,
    17271771,
    17287786,
    17272911,
    17301702,
    17297618,
    17274765,
    17287683,
    17272844,
    17275060,
    17278184,
    17271754,
    17294989,
    17272132,
    17299590,
    17296836,
    17288018,
    17273033,
    17294879,
    17277338,
    17301707,
    17278237,
    17272038,
    17277161,
    17278230,
    17271730,
    17274768,
    17299593,
    17272400,
    17293853,
    17293832,
    17711371,
    17299279,
    17278259,
    17277343,
    17271776,
    17273476,
    17272037,
    17294966,
    17272849,
    17301731,
    17272076,
    17287668,
    17287819,
    17297607,
]

select_cols = [
    "material_id",
    "manual_classification",
    "workflow_status",
    "remarks",
    "url",
    "title",
]
bms_dir = Path("scripts/BMS")

xlsx_files = list(bms_dir.glob("*.xlsx"))
console.print(f"Found {len(xlsx_files)} xlsx files in {bms_dir}.")
# remove files with "overview" or "llm" in the name
xlsx_files = [
    file
    for file in xlsx_files
    if ("overview" not in file.name) and ("llm" not in file.name)
]
console.print(
    f"After filtering out overviews & llm classification sheets, {len(xlsx_files)} xlsx files remain."
)

all_dfs: dict[str, dict[str, pl.DataFrame]] = {}
for file in xlsx_files:
    all_dfs[file.name] = {
        "data_entry": pl.read_excel(file, sheet_name="Data entry").with_columns(
            pl.col("material_id").cast(pl.Int64)
        )
    }

console.print(
    f"Loaded {len(all_dfs)} files with {sum(len(v['data_entry']) for v in all_dfs.values())} total rows in 'Data entry' sheets."
)


# now load in the overview sheet
overview_file = list(bms_dir.glob("*overview*.xlsx"))
assert len(overview_file) == 1, "There should be exactly one overview file."
overview_sheet = {
    "data_entry": pl.read_excel(overview_file[0], sheet_name="Data entry").with_columns(
        pl.col("material_id").cast(pl.Int64)
    )
}

console.print(
    f"Loaded overview file with {len(overview_sheet['data_entry'])} rows in 'Data entry' sheet ."
)

# compare lengths of overview sheet to sum of lengths of other sheets
console.print(
    f' "Data entry" overview vs sum of other sheets: {len(overview_sheet["data_entry"])} vs {sum(len(v["data_entry"]) for v in all_dfs.values())}'
)

# now check that all material_ids in input_material_ids are present in the overview sheet, and only one of the other sheets
overview_material_ids = set(overview_sheet["data_entry"]["material_id"].to_list())
missing_material_ids = set(input_material_ids) - overview_material_ids
if missing_material_ids:
    console.print(f"Missing material_ids in overview sheet: {missing_material_ids}")
else:
    console.print("All input material_ids are present in the overview sheet.")

for material_id in input_material_ids:
    count = sum(
        material_id in v["data_entry"]["material_id"].to_list()
        for v in all_dfs.values()
    )
    if count == 0:
        console.print(
            f"material_id {material_id} is missing from all individual sheets."
        )
    elif count > 1:
        console.print(
            f"material_id {material_id} is present in {count} individual sheets."
        )
    else:
        pass  # all good


# loop over all material ids
# look up all rows with that material id in all dataframes
# then print the 'select_cols' columns for each row found, along with the filename it was found in

material_id_data: dict[int, list[tuple[str, str, pl.DataFrame]]] = {}
for material_id in input_material_ids:
    console.print(f"\nMaterial ID: {material_id}")
    found = False
    for filename, dfs in all_dfs.items():
        for sheet_name, df in dfs.items():
            rows = df.filter(pl.col("material_id") == material_id)
            if len(rows) > 0:
                found = True
                console.print(f"Found in file: {filename}, sheet: {sheet_name}")
                material_id_data.setdefault(material_id, []).append(
                    (filename, sheet_name, rows.select(select_cols))
                )
    if not found:
        console.print("Not found in any individual sheets.")
    # also check the overview sheet
    rows = overview_sheet["data_entry"].filter(pl.col("material_id") == material_id)
    if len(rows) > 0:
        material_id_data.setdefault(material_id, []).append(
            (
                "BMS_total_overview_updated_2025-09-11.xlsx",
                "data_entry",
                rows.select(select_cols),
            )
        )
    else:
        console.print("Not found in overview sheet data entry.")


# then for each material id, print the found rows nicely

for material_id, entries in material_id_data.items():
    console.rule(f"Material ID: {material_id}")
    for filename, sheet_name, df in entries:
        console.print(f"[cyan]{filename}[/cyan]:")
        for row in df.to_dicts():
            console.print(row)
    console.print("\n")
