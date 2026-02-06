"""
Compare Excel files between faculty_sheets/ and from_teams/ directories item-by-item.

Collects all data from 'Data entry' sheets across all Excel files in each directory,
adds provenance via 'from_file' column, and compares items by material_id.
"""

import sys
from pathlib import Path

import polars as pl
import openpyxl


# Use ASCII-safe checkmarks for Windows console
CHECK = "[OK]"
CROSS = "[DIFF]"


def get_excel_files(root: Path) -> dict[str, Path]:
    """Get all Excel files by relative path (faculty/filename.xlsx)."""
    files = {}
    for excel_file in root.rglob("*.xlsx"):
        # Get relative path from root, excluding the root itself
        rel_path = excel_file.relative_to(root)
        files[str(rel_path)] = excel_file
    return files


def read_data_entry_sheet(file_path: Path) -> pl.DataFrame | None:
    """Read the 'Data entry' sheet from an Excel file."""
    try:
        # First, use openpyxl to get sheet names and find the "Data entry" sheet
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        wb.close()

        # Try to find "Data entry" sheet (case-insensitive)
        sheet_to_read = None
        for name in sheet_names:
            if name.lower() == "data entry":
                sheet_to_read = name
                break

        if sheet_to_read is None:
            return None

        # Read the sheet with polars
        df = pl.read_excel(file_path, sheet_name=sheet_to_read)
        return df
    except Exception as e:
        print(f"  Warning: Could not read 'Data entry' from {file_path}: {e}")
        return None


def collect_data(root: Path, source_name: str) -> pl.DataFrame:
    """Collect all 'Data entry' sheets from Excel files in root, adding from_file column."""
    all_dfs = []
    for excel_file in root.rglob("*.xlsx"):
        if excel_file.name == 'overview.xlsx':
            continue
        df = read_data_entry_sheet(excel_file)
        if df is not None:
            rel_path = excel_file.relative_to(root.parent)
            df = df.with_columns(pl.lit(str(rel_path)).alias("from_file"))
            all_dfs.append(df)
    if all_dfs:
        combined = pl.concat(all_dfs)
        if combined.select("material_id").is_duplicated().any():
            print(f"Warning: Duplicate material_ids in {source_name}")
        return combined
    else:
        return pl.DataFrame()


def compare_dataframes(df1: pl.DataFrame, df2: pl.DataFrame, name: str) -> pl.DataFrame:
    diff_data = []

    # Common rows with differences
    in_both = df1.join(df2, on="material_id", how="inner", suffix="_teams")
    for row in in_both.to_dicts():
        mid = row["material_id"]
        diffs = {}
        for col in ['remarks', 'v1_manual_classification', 'v2_manual_classification']:
            val1 = row[col] if col in df1.columns else None
            val2 = row[f"{col}_teams"] if f"{col}_teams" in row else None
            if val1 != val2:
                diffs[col] = (val1, val2)
        if diffs:
            # Get from_files
            from1 = df1.filter(pl.col("material_id") == mid)["from_file"].to_list()
            from2 = df2.filter(pl.col("material_id") == mid)["from_file"].to_list()
            for col, (val1, val2) in diffs.items():
                diff_data.append({
                    "material_id": mid,
                    "diff_type": "diff",
                    "column": col,
                    "new_version_value": val1,
                    "previous_version_value": val2,
                    "new_version_files": from1,
                    "previous_version_files": from2
                })

    return pl.DataFrame(diff_data)
def main() -> None:
    """Main entry point."""
    # Define paths
    faculty_dir = Path("faculty_sheets")
    from_teams_dir = Path("full_backups") / "backup 23 jan 2026" / "faculty_sheets"

    if not faculty_dir.exists():
        print(f"Error: Directory '{faculty_dir}' not found")
        sys.exit(1)

    if not from_teams_dir.exists():
        print(f"Error: Directory '{from_teams_dir}' not found")
        sys.exit(1)

    # Get all Excel files from both directories
    faculty_files = get_excel_files(faculty_dir)
    from_teams_files = get_excel_files(from_teams_dir)

    print(f"New version sheets: {len(faculty_files)} files")
    print(f"Previous version sheets: {len(from_teams_files)} files")

    # Collect data
    faculty_df = collect_data(faculty_dir, "new_version")
    teams_df = collect_data(from_teams_dir, "previous_version")

    print(f"Collected {faculty_df.height} rows from new_version")
    print(f"Collected {teams_df.height} rows from previous_version")

    # Compare
    diff_df = compare_dataframes(faculty_df, teams_df, "new_version vs previous_version")
    if diff_df.height > 0:
        diff_df.write_excel("differences.xlsx")
        print(f"Exported {diff_df.height} differences to differences.xlsx")

    # Summary
    print(f"\n{'='*70}")
    print("SUMMARY")
    print(f"{'='*70}")
    if diff_df.height > 0:
        print(f"{CROSS} Differences were found between the directories")
    else:
        print(f"{CHECK} All compared files are identical")


if __name__ == "__main__":
    main()
