## Easy Access (ea-cli) — Processing pipeline overview

This document explains what happens when you run `run.py process`, how data flows through the codebase, where inputs/outputs live, and which modules/tables are involved. It is intended to be a single self-contained reference for developers and operators.

### Checklist (what this file contains)
- [x] End-to-end description of `run.py process` (ingest → process → enrich → verify → export)
- [x] Inputs and outputs (files and DBs)
- [x] Directory structure focused on data folders
- [x] Mapping to key modules and DB models
- [x] Simple pipeline diagram you can read without other files
- [x] Command flags and behaviour notes

## High-level: what `run.py process` does

1. CLI parses flags and builds a runtime `EasyAccessSettings` using `easy_access.settings.SETTINGS`.
2. `EasyAccessTool` (in `easy_access/main.py`) is instantiated with the main settings and runtime flags.
3. In default mode it runs these stages in order:
   - ingest (`run_ingest`) — read raw export sheet(s) into staging
   - process (`run_process`) — convert staged rows to canonical DB records
   - enrich (`run_enrich`) — fetch OSIRIS course/person data where needed
   - verify file existence (`run_verify_file_existence`) — check Canvas URLs and update flags
   - export (`run_export`) — create per-faculty Excel sheets
4. Each stage updates the repository SQLite DB files and may write output files (Excel, PDFs, backups) depending on flags.

## Simple pipeline diagram

Text flow (left-to-right):

COPYRIGHT exports (XLSX) in repo folders
  --> ingest (staging table `StagedCopyrightItem`)
  --> processing (canonical tables: `CopyrightItem`, `PDF`, `Course`, `Person`, `Programme`, `Organization`)
  --> enrichment (OSIRIS → `Course`, `Person`)
  --> file-existence checks (update `file_exists`, `last_canvas_check`)
  --> export (Excel files in `faculty_sheets/`)

Mermaid (optional renderers):

```mermaid
flowchart LR
  A[cip_sheets/, raw_copyright_data/] --> B[Ingest]
  B --> C[StagedCopyrightItem (staging DB)]
  C --> D[Process]
  D --> E[CopyrightItem, PDF, Course, Person]
  E --> F[Enrich (OSIRIS)]
  E --> G[Verify file existence]
  E & F & G --> H[Export -> faculty_sheets/]
```

## Inputs (where data comes from)
- Raw copyright export spreadsheets :
  - `raw_copyright_data/`
  - CLI: `--other-sheet <path>` to ingest one xlsx file
-  application DB file (SQLite) at repo root: `db.sqlite3`

## Outputs (what the pipeline produces)
- Canonical DB updates: tables under the `easy_access/db/models.py` mapping (see Models section).
- Per-faculty Excel workbook(s) written to `faculty_sheets/<FACULTY>/`.
- Downloaded PDF files (if downloader runs) to `pdf_downloads/`.
- Backups under `full_backups/` (when backup operations executed).
- Logs in `logs/` and console output.

## Key stages and where code lives
- Ingest
  - Code: `easy_access/db/ingest.py` (called by `EasyAccessTool.run_ingest`)
  - Action: normalize XLSX columns, insert into `StagedCopyrightItem`.
  - Errors: recorded to `StagedProcessingFailure`.

- Process
  - Code: `easy_access/db/update.py`, `easy_access/pipeline.py`, `easy_access/merge_rules.py` and helpers
  - Action: map staged rows to canonical entities; create/update `CopyrightItem` and `PDF`; link to `Course`, `Faculty`, `Programme`, `Person`; write `ItemUpdate` history; mark duplicates and apply merge rules.
  - Errors: per-row failures are saved to `StagedProcessingFailure` (admin commands to inspect/retry).

- Enrich
  - Code: `easy_access/enrichment/osiris.py`
  - Action: fetch course and people data from OSIRIS and update `Course` and `Person` tables; controlled by `--osiris-update` and `--osiris-full-refresh` flags.

- File existence verification
  - Code: `easy_access/maintenance/file_existence.py`
  - Action: HTTP/Canvas checks on `CopyrightItem.url` (or derived paths), update `file_exists` and `last_canvas_check` on `CopyrightItem`, and possibly `PDF` download flags.

- Export
  - Code: `easy_access/sheets/sheet.py`, `easy_access/sheets/export.py`
  - Action: build Excel workbook(s) per faculty using canonical DB state; controlled by `--changes`, `--single-faculty`, and `--disable-writes`.

## Important DB models (where data is stored)
- Staging tables
  - `StagedCopyrightItem` (table `staged_copyright_item`) — raw, normalized rows
  - `StagedFacultyUpdate` (table `staged_faculty_update`) — faculty sheet updates
  - `StagedProcessingFailure` (table `staged_processing_failures`) — captures per-row failures

- Canonical tables
  - `CopyrightItem` (table `copyright_data`) — core item (material_id, url, filename, status, file_exists, workflow_status, etc.)
  - `PDF` (table `pdf_data`) — per-item PDF metadata and download flags
  - `Course`, `Person`, `Organization`, `Programme`, `ItemUpdate` — related entities and history

## File & folder map (most relevant)
- Data inputs
  - `cip_sheets/` — main COPYRIGHT export XLSX files
  - `raw_copyright_data/` — saved raw exports
  - `copyright_data_for_SURF_import/` — example import files

- Generated outputs
  - `faculty_sheets/` — generated Excel exports (per-faculty subfolders)
  - `faculty_sheets_old/` — archived exports
  - `pdf_downloads/` — downloaded PDFs
  - `full_backups/` — backup snapshots

- Database & runtime
  - `db.sqlite3`, `db_bakkup.sqlite3`, `db_teams.sqlite3` — SQLite DB files
  - `logs/` — runtime logs

## Command flags and behavior notes
- Stage selection (mutually exclusive): `--ingest-only`, `--process-only`, `--export-only`, `--enrich-only`, `--file-exists-only`.
- `--disable-writes` prevents write operations (read-only / dry-run behaviour depending on implementation).
- `--changes` limits exports to changed items only.
- `--single-faculty <ABBR>` restricts processing/export to that faculty.
- `--other-sheet <path>` replaces default ingestion sources with a single XLSX.

## Contract (inputs / outputs / errors)
- Inputs: XLSX export files (raw data from `raw_copyright_data/`, or via `--other-sheet`; and `faculty_sheets/` for user data) and existing SQLite DB files.
- Outputs: updated canonical DB table and Excel files in `faculty_sheets/`; optional downloaded PDFs and backups.
- Errors: per-row failures do not abort the run; they are stored in `StagedProcessingFailure` for admin inspection and retry.

## Where to look for code-level traces
- Entrypoint and CLI wiring: `run.py` (creates `EasyAccessSettings`, instantiates `EasyAccessTool`, calls `run_*` methods).
- Tool implementation and stage orchestration: `easy_access/main.py`.
- Ingest: `easy_access/db/ingest.py`.
- Processing/updating DB: `easy_access/db/update.py`, `easy_access/pipeline.py`, `easy_access/merge_rules.py`.
- Enrichment: `easy_access/enrichment/osiris.py`.
- File checks: `easy_access/maintenance/file_existence.py`.
- Export generation: `easy_access/sheets/sheet.py`, `easy_access/sheets/export.py`.

## Edge cases & notes for maintainers
- `--disable-writes` and `--changes` alter side effects; test a run with small data first.
- Staging rows are normalized before storage; look at `db.base` helpers (standardize dataframe) if you need the exact normalization rules.
- Admin CLI commands in `run.py` let you inspect/retry/cleanup `StagedProcessingFailure` entries.
