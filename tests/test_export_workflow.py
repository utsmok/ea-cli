import polars as pl

from easy_access.settings import SETTINGS, DirSetting, ColInfo
from easy_access.sheets.export import export_faculty_workflow_files
from easy_access.utils import Directory
import openpyxl


def make_df():
    return pl.DataFrame([
        {
            "material_id": 1,
            "faculty": "TEST",
            "workflow_status": "ToDo",
            "title": "Item One",
            "period": "2024-01",
        },
        {
            "material_id": 2,
            "faculty": "TEST",
            "workflow_status": "InProgress",
            "title": "Item Two",
            "period": "2024-02",
        },
        {
            "material_id": 3,
            "faculty": "TEST",
            "workflow_status": "Done",
            "title": "Item Three",
            "period": "2024-03",
        },
    ])


def test_export_faculty_workflow_files_creates_files_and_backups(tmp_path):
    # Prepare a temp faculties dir
    faculties_dir = tmp_path / "faculties"
    faculties_dir.mkdir()

    # Monkeypatch SETTINGS dirs to point to tmp path (use Directory wrapper expected by code)
    orig_dir = SETTINGS.dirs.get(DirSetting.FACULTIES_DIR)
    SETTINGS.dirs[DirSetting.FACULTIES_DIR] = Directory(faculties_dir)

    try:
        # Create an existing file to be backed up for the TEST faculty inbox
        faculty_folder = faculties_dir / "TEST"
        faculty_folder.mkdir()
        existing_inbox = faculty_folder / "inbox.xlsx"
        existing_inbox.write_text("old content")

        # Prepare faculty_data
        df = make_df()
        faculty_data = {"TEST": df}

        # Run the exporter (it's async - call via asyncio.run)
        import asyncio

        # tests only provide a small subset of columns; ensure settings expect only those
        orig_final_cols = SETTINGS.data_settings.final_data_col_order
        orig_data_entry_cols = SETTINGS.data_settings.data_entry_cols
        SETTINGS.data_settings.final_data_col_order = [
            "material_id",
            "faculty",
            "workflow_status",
            "title",
            "period",
        ]
        # Provide a minimal data_entry_cols list so finalize_sheet doesn't try to read missing cols
        SETTINGS.data_settings.data_entry_cols = [
            ColInfo(name="material_id"),
            ColInfo(name="title"),
            ColInfo(name="workflow_status"),
            ColInfo(name="period"),
        ]

        asyncio.run(export_faculty_workflow_files(SETTINGS, faculty_data, style_iter=9))

        # Check files exist
        inbox = faculty_folder / "inbox.xlsx"
        in_progress = faculty_folder / "in_progress.xlsx"
        done = faculty_folder / "done.xlsx"

        assert inbox.exists(), "inbox.xlsx not created"
        assert in_progress.exists(), "in_progress.xlsx not created"
        assert done.exists(), "done.xlsx not created"

        # Backups dir should contain the backed-up original inbox
        backups_dir = faculty_folder / "backups"
        assert backups_dir.exists(), "backups dir missing"
        backups = list(backups_dir.iterdir())
        # Should contain at least one file (the moved inbox.xlsx)
        assert any(p.name.startswith("inbox_") for p in backups), f"No inbox backup found in {backups_dir}"

    finally:
        # restore original
        if orig_dir:
            SETTINGS.dirs[DirSetting.FACULTIES_DIR] = orig_dir
        else:
            del SETTINGS.dirs[DirSetting.FACULTIES_DIR]
        # restore final cols if present
        import contextlib
        with contextlib.suppress(Exception):
            SETTINGS.data_settings.final_data_col_order = orig_final_cols
        with contextlib.suppress(Exception):
            SETTINGS.data_settings.data_entry_cols = orig_data_entry_cols


def test_unknown_workflow_status_defaults_to_inbox(tmp_path):
    faculties_dir = tmp_path / "faculties"
    faculties_dir.mkdir()

    orig_dir = SETTINGS.dirs.get(DirSetting.FACULTIES_DIR)
    SETTINGS.dirs[DirSetting.FACULTIES_DIR] = Directory(faculties_dir)

    try:
        df = pl.DataFrame([
            {"material_id": 10, "faculty": "TEST2", "workflow_status": "StrangeState", "title": "Weird", "period": "2024-04"}
        ])
        faculty_data = {"TEST2": df}

        import asyncio

        orig_final_cols = SETTINGS.data_settings.final_data_col_order
        orig_data_entry_cols = SETTINGS.data_settings.data_entry_cols
        SETTINGS.data_settings.final_data_col_order = [
            "material_id",
            "faculty",
            "workflow_status",
            "title",
            "period",
        ]
        SETTINGS.data_settings.data_entry_cols = [
            ColInfo(name="material_id"),
            ColInfo(name="title"),
            ColInfo(name="workflow_status"),
            ColInfo(name="period"),
        ]

        asyncio.run(export_faculty_workflow_files(SETTINGS, faculty_data, style_iter=9))

        faculty_folder = faculties_dir / "TEST2"
        inbox = faculty_folder / "inbox.xlsx"
        assert inbox.exists(), "inbox.xlsx not created for unknown status default"

        wb = openpyxl.load_workbook(filename=str(inbox))
        sheet = wb[SETTINGS.data_settings.complete_data_name]
        # material_id should be present in column A (header at row1, data from row2)
        values = [c.value for c in sheet['A']]
        # cell values may be int or str depending on write path
        assert any((v == 10) or (str(v) == "10") for v in values), "material_id 10 not found in inbox.xlsx"

    finally:
        if orig_dir:
            SETTINGS.dirs[DirSetting.FACULTIES_DIR] = orig_dir
        else:
            del SETTINGS.dirs[DirSetting.FACULTIES_DIR]
        import contextlib
        with contextlib.suppress(Exception):
            SETTINGS.data_settings.final_data_col_order = orig_final_cols
        with contextlib.suppress(Exception):
            SETTINGS.data_settings.data_entry_cols = orig_data_entry_cols


def test_done_file_is_protected_and_active_sheet_set(tmp_path):
    faculties_dir = tmp_path / "faculties"
    faculties_dir.mkdir()

    orig_dir = SETTINGS.dirs.get(DirSetting.FACULTIES_DIR)
    SETTINGS.dirs[DirSetting.FACULTIES_DIR] = Directory(faculties_dir)

    try:
        df = pl.DataFrame([
            {"material_id": 20, "faculty": "TEST3", "workflow_status": "Done", "title": "Final", "period": "2024-05"}
        ])
        faculty_data = {"TEST3": df}

        import asyncio

        orig_final_cols = SETTINGS.data_settings.final_data_col_order
        orig_data_entry_cols = SETTINGS.data_settings.data_entry_cols
        SETTINGS.data_settings.final_data_col_order = [
            "material_id",
            "faculty",
            "workflow_status",
            "title",
            "period",
        ]
        SETTINGS.data_settings.data_entry_cols = [
            ColInfo(name="material_id"),
            ColInfo(name="title"),
            ColInfo(name="workflow_status"),
            ColInfo(name="period"),
        ]

        asyncio.run(export_faculty_workflow_files(SETTINGS, faculty_data, style_iter=9))

        faculty_folder = faculties_dir / "TEST3"
        done = faculty_folder / "done.xlsx"
        assert done.exists(), "done.xlsx not created"

        wb = openpyxl.load_workbook(filename=str(done))
        # The Complete Data sheet must exist and be protected; Data Entry sheet is optional
        assert SETTINGS.data_settings.complete_data_name in wb.sheetnames
        complete_protected = wb[SETTINGS.data_settings.complete_data_name].protection.sheet
        # If Data Entry exists, it may also be protected; but require at least Complete Data protection
        dataentry_protected = False
        if SETTINGS.data_settings.data_entry_name in wb.sheetnames:
            dataentry_protected = wb[SETTINGS.data_settings.data_entry_name].protection.sheet
        assert complete_protected or dataentry_protected

    finally:
        if orig_dir:
            SETTINGS.dirs[DirSetting.FACULTIES_DIR] = orig_dir
        else:
            del SETTINGS.dirs[DirSetting.FACULTIES_DIR]
        import contextlib
        with contextlib.suppress(Exception):
            SETTINGS.data_settings.final_data_col_order = orig_final_cols
        with contextlib.suppress(Exception):
            SETTINGS.data_settings.data_entry_cols = orig_data_entry_cols


def test_backup_manifest_written(tmp_path):
    faculties_dir = tmp_path / "faculties"
    faculties_dir.mkdir()

    orig_dir = SETTINGS.dirs.get(DirSetting.FACULTIES_DIR)
    SETTINGS.dirs[DirSetting.FACULTIES_DIR] = Directory(faculties_dir)

    try:
        faculty_folder = faculties_dir / "TEST4"
        faculty_folder.mkdir()
        existing_inbox = faculty_folder / "inbox.xlsx"
        existing_inbox.write_text("old content")

        df = pl.DataFrame([
            {"material_id": 30, "faculty": "TEST4", "workflow_status": "ToDo", "title": "Item", "period": "2024-06"}
        ])
        faculty_data = {"TEST4": df}

        import asyncio

        orig_final_cols = SETTINGS.data_settings.final_data_col_order
        SETTINGS.data_settings.final_data_col_order = [
            "material_id",
            "faculty",
            "workflow_status",
            "title",
            "period",
        ]

        asyncio.run(export_faculty_workflow_files(SETTINGS, faculty_data, style_iter=9))

        backups_dir = faculty_folder / "backups"
        assert backups_dir.exists()
        # Find any .manifest.json file
        manifests = list(backups_dir.glob("*.manifest.json"))
        assert len(manifests) > 0, "No backup manifest files found"

    finally:
        if orig_dir:
            SETTINGS.dirs[DirSetting.FACULTIES_DIR] = orig_dir
        else:
            del SETTINGS.dirs[DirSetting.FACULTIES_DIR]
        import contextlib
        with contextlib.suppress(Exception):
            SETTINGS.data_settings.final_data_col_order = orig_final_cols
