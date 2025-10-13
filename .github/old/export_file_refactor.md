## Replace weekly exports with workflow-based faculty sheets — plan



Checklist

- [ ] Inventory current export & sheet utilities (done)
- [ ] Finalise design decisions & edge-case rules
- [ ] Add backup + atomic-swap helpers (reusable)
- [ ] Add workbook protection helper
- [ ] Add status-based export writer (per-faculty)
- [ ] Update export orchestrator to call status writers + overview
- [ ] Add/adjust import/read logic notes so row moves are ignored (use workflow_status cell)
- [ ] Add unit + integration tests (export names, backups, protection, status grouping)
- [ ] Docs & changelog entry
- [ ] Optional: small follow-ups (CLI flag, settings.yaml docs)

Design summary
- Replace the single `weekly` update sheet with three per-faculty workbooks that mirror the current layout and contain the same tabs: `Complete Data` and `Data Entry`.
- Per faculty, produce these files in the faculty folder: `inbox.xlsx` (ToDo), `in_progress.xlsx` (InProgress), `done.xlsx` (Done).
- Keep the existing `total overview` export, named `overview_[YYYY-MM-DD].xlsx` and protected.
- When exporting, move any existing target file(s) into a faculty backup directory and append a timestamp to the filename (e.g. `inbox_20250910_153200.xlsx`). Use an atomic write pattern already present in `easy_access/sheets/sheet.py`.
- Import/ingest semantics: never rely on which workbook a row is in; use the `workflow_status` column value in the `Data Entry` sheet to determine grouping. This enforces a single source-of-truth and allows the script to ignore manual movement of rows between files.

Why this is better (advantages vs current system)
- Clear workflow separation: users see only the items relevant to their stage (less noise), which reduces accidental edits and manual re-sorting.
- More robust: canonical `workflow_status` field (already enforced in DB merge logic) becomes the single grouping key — removes reliance on sheet order or file naming.
- Safer exports: `done.xlsx` + `overview` protected to avoid accidental edits to finalised items.
- Simpler merge heuristics: because each workbook represents a stage, merging rules can treat `workflow_status` as authoritative for grouping; canonical priority enforcement remains for status changes (Done > InProgress > ToDo).

Potential drawbacks & mitigations
- Increased number of files per faculty — mitigated by stable naming (`inbox/in_progress/done`) and automatic backups.
- Users may still try to move rows; script explicitly ignores file placement and reads the `workflow_status` cell, log any anomalies.
- Unknown custom workflow_status values: treat them as InProgress and log a warning; optionally add a `custom.xlsx` bucket later.

Contract / acceptance criteria
- Inputs: Settings object, existing DB CopyrightItem rows.
- Outputs: For each faculty directory, three XLSX files named `inbox.xlsx`, `in_progress.xlsx`, `done.xlsx`; plus `overview_[date].xlsx` in the overview destination.
- Behavior: Files contain `Complete Data` and `Data Entry` sheets with the same columns/layout as current exports; `done.xlsx` and `overview_[date].xlsx` are sheet-protected from edits; existing target files moved to a backup dir with timestamp; grouping based on `workflow_status` column values, not file placement.
- Error modes: If write disabled via Settings or CLI flag, skip writes and log; if file move/backup fails, leave old file untouched and log an error; unknown workflow_status values are logged and put in `inbox.xlsx` by default.

Edge cases
- Missing `workflow_status` column: populate default `ToDo` before export (current behavior in `sheet.py` sets default).
- Mixed/unknown string values: map case-insensitively against canonical options (Done/InProgress/ToDo); extras go to `inbox` and produce a WARN log.
- Very large exports: reuse existing atomic-write and style helpers; if memory pressure becomes an issue, consider chunked writes (future).
- Concurrent runs: atomic-swaps + unique timestamped backups reduce corruption risk; recommend scheduling to avoid overlapping runs.

Concrete code changes (files and responsibilities)
- easy_access/sheets/sheet.py
  - Reuse existing DataEntrySheet and atomic write helpers. Add a small helper: `protect_workbook(path, sheet_names=None, lock=True)` that sets sheet protection on the written workbook (openpyxl supports sheet.protection). If builder uses xlsxwriter for formatting, reopen with openpyxl for protection step.
  - Ensure exported workbook opens on `Data Entry` sheet (there's a TODO in `.github/todo.md`); set active sheet accordingly using openpyxl after the atomic save.

- easy_access/sheets/export.py
  - Add a new or extend existing function: `export_faculty_workflow_files(settings, faculty, df)` — accepts the faculty dataframe and creates the three per-status workbooks plus the faculty overview movement logic (backup existing matching filenames). Use existing `finalize_sheet`/style helpers for layout consistency.
  - Modify `export_reports_async` to call the new writer per faculty instead of producing a `weekly` sheet.

- easy_access/sheets/analysis.py
  - Reuse `create_faculty_overviews` for the `overview_[date].xlsx` generation. Keep its current backup logic for overview files. Ensure it sets protection on the workbook (or call the new protect helper).

- easy_access/pipeline.py
  - No major changes other than ensuring `export_reports_async` still invoked; optionally add a settings flag to enable 'workflow-mode' but default to enabled when this change is merged.

- easy_access/db/update.py + easy_access/db/ingest.py
  - Mostly unchanged. Ensure import/ingest logic continues to read `workflow_status` from uploaded Data Entry sheet and trust the existing canonical priority logic that prevents downgrades (Done->ToDo) — tests already exist (see `tests/test_workflow_status_regression.py`). Add a small note in docstrings to explicitly ignore sheet filename and rely on `workflow_status` cell.

- New helpers (proposed):
  - easy_access/sheets/backup.py (small module)
    - backup_existing_file(target_path, backups_dir, timestamp_fmt) -> new_path
    - This centralizes the move+rename logic already used for overviews and keeps export code tidy.

- Tests to add/modify (tests/):
  - tests/test_export_workflow.py
    - Test that for a small dataframe with workflow_status set to ToDo/InProgress/Done, the correct files are written and contain the expected items.
    - Test that existing files are moved to a backup dir and renamed with timestamp.
    - Test that `done.xlsx` and `overview_[date].xlsx` are protected (openpyxl sheet.protection.enabled or workbook.active index points at Data Entry and protection set).
    - Integration test: simulate an item moved by a user to a different file but with unchanged `workflow_status` — export should keep item in its canonical file.

Merge-rules simplifications (impact on `easy_access/merge_rules.py`)
- We can keep the current canonical ordering and downgrade guard for `workflow_status` (already implemented in `db/update.py`). The refactor gives us stronger guarantees so we can:
  - Remove heuristics that relied on file ordering or spreadsheet option ordering (this has already been partly addressed by build_merge_rules_from_settings in `merge_rules.py`).
  - Make `workflow_status` a strictly authoritative field when exported items are re-ingested: the import logic updates DB workflow_status only when the cell differs and respects canonical priorities (already in place).

Logging & observability
- Log each faculty export with counts per file (e.g., "Exported 123 ToDo -> inbox.xlsx, 45 InProgress -> in_progress.xlsx, 1 Done -> done.xlsx") and backup file paths.
- Log unknown workflow_status values and the file they were written to.

Rollout strategy
1. Create the planning PR with tests and helpers (no behaviour change yet). Mark as feature branch `export-workflow`.
2. Implement writers + tests, keep a feature-flag in settings (e.g., settings.export.workflow_mode = True) defaulting to False. Run in CI.
3. Deploy to staging; run a manual export, inspect backups & outputs.
4. Flip default to True after one successful staged run and document change in `.github/changelog.md`.

Open questions
- Do we want to treat unknown/custom workflow_status values by default as "ToDo/inbox" or as "InProgress"? Yes: put unknowns into `inbox.xlsx` and surface a WARN so staff can inspect. This should not happen, as users will input via dropdowns w/ validation in excel, but it's safer to handle gracefully.
- Should the backup directory be a single faculty-level `backups/` (per-faculty) or a central `overviews_backup` style directory? Yes, implement recommend per-faculty `backups/` to keep faculty folders self-contained.
- Should the exporter also write a small manifest JSON in the faculty backup directory describing the backup (source filename, timestamp, item counts)? Yes, good idea.

Small follow-ups (low-risk improvements)
- Add a CLI flag `--export-workflow` or a Settings boolean to toggle the new behavior (useful while we roll out).
- Add a small audit log file per run summarising counts and warnings.

References in repo (investigation notes)
- core export orchestration: `easy_access/sheets/export.py`
- faculty overview generation: `easy_access/sheets/analysis.py`
- sheet utilities: `easy_access/sheets/sheet.py` (atomic write helper and data-entry builder)
- existing canonical workflow guard & tests: `easy_access/db/update.py`, `tests/test_workflow_status_regression.py`
- settings and data column definitions: `easy_access/settings.py` and `settings.yaml`

Next steps
1. Add `easy_access/sheets/backup.py` and the `protect_workbook()` helper in `sheet.py`.
2. Implement `export_faculty_workflow_files()` in `sheets/export.py` that selects items by `workflow_status` and writes the three files, using atomic-save and backup helper. Reuse existing `finalize_sheet` and DataEntrySheet mechanics to preserve layout.
3. Add unit tests `tests/test_export_workflow.py` and run the test suite locally.
