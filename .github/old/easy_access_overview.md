## Easy Access (ea-cli) — Processing pipeline overview

This document explains what happens when you run `run.py process`, how data flows through the codebase, where inputs/outputs live, and which modules/tables are involved. It is intended to be a concise, developer-friendly reference updated to reflect the post-refactor state (branch `main`).

### What this file contains
- End-to-end description of `run.py process` (ingest → process → enrich → verify → export)
- Inputs and outputs (files and DBs)
- Directory and data folder mapping
- Mapping to key modules and DB models
- Current runtime flags and CLI commands
- Notes on implemented vs pending items after the refactor

## High-level: what `run.py process` does (runtime wiring)

1. CLI parses flags and builds a runtime `EasyAccessSettings` using `easy_access.settings.SETTINGS`.
2. `EasyAccessTool` (in `easy_access/main.py`) is instantiated with the main settings and runtime flags.
3. By default it runs the staged pipeline in order. The typical run sequence invoked by `run.py process` is:
   - ingest (`run_ingest`) — read and normalize raw export sheet(s) into staging tables
   - process (`run_process`) — convert staged rows to canonical DB records and apply merge rules
   - enrich (`run_enrich`) — fetch OSIRIS course/person data where missing or stale
   - relations (`run_relations`) — link courses/programmes in bulk (optimized relations code)
   - verify file existence (`run_verify_file_existence`) — check Canvas URLs and update flags
   - export (`run_export`) — create per-faculty Excel sheets
4. Each stage updates the repository SQLite DB files and may write output files (Excel, PDFs, backups) depending on flags and settings.

## Simple pipeline diagram

Text flow (left-to-right):

COPYRIGHT exports (XLSX) in repo folders
  --> ingest (staging table `StagedCopyrightItem`)
  --> processing (canonical tables: `CopyrightItem`, `PDF`, `Course`, `Person`, `Programme`, `Organization`)
  --> enrichment (OSIRIS → `Course`, `Person`)
  --> relations (bulk linking)
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
  E --> R[Relations (bulk linking)]
  E --> G[Verify file existence]
  E & F & G & R --> H[Export -> faculty_sheets/]
```

## Inputs (where data comes from)
- Raw copyright export spreadsheets:
  - `raw_copyright_data/` and `cip_sheets/`
  - CLI: `--other-sheet <path>` to ingest a single xlsx file for ad-hoc runs
- Application DB files (SQLite) at repo root: `db.sqlite3` (and optional backups)

## Outputs (what the pipeline produces)
- Canonical DB updates: tables under `easy_access/db/models.py`.
- Per-faculty Excel workbook(s) written to `faculty_sheets/<FACULTY>/`.
- Downloaded PDF files (when downloader runs) to `pdf_downloads/`.
- Backups and manifests under `full_backups/` (backup helpers implemented; pre/post pipeline integration is configurable and partially available via CLI).
- Logs in `logs/` and console output.

## Key stages and where code lives (quick map)
- Ingest
  - Code: `easy_access/db/ingest.py` (used by `EasyAccessTool.run_ingest`)
  - Action: normalize XLSX columns and insert into `StagedCopyrightItem`.
  - Errors: normalized/staging issues are recorded to `StagedProcessingFailure` for inspection and retry.

- Process
  - Code: `easy_access/db/update.py`, `easy_access/pipeline.py`, `easy_access/merge_rules.py`
  - Action: map staged rows to canonical entities; create/update `CopyrightItem` and `PDF`; run merge rules; mark duplicates.
  - Notes: processing is transactional and records per-row failures to `StagedProcessingFailure` rather than aborting the whole run.

- Enrich
  - Code: `easy_access/enrichment/osiris.py`
  - Action: concurrent fetches for course and person data; persists via bulk upserts; TTL controls staleness selection.

- Relations
  - Code: `easy_access/db/relations.py`
  - Action: optimized bulk linking for courses/programmes to avoid N+1 updates; deterministic `bulk_update` semantics.

- File existence verification
  - Code: `easy_access/maintenance/file_existence.py`
  - Action: Canvas HTTP checks on item URLs (or derived file paths), update `file_exists`, `last_canvas_check`, and optionally trigger PDF downloads.

- Export
  - Code: `easy_access/sheets/sheet.py`, `easy_access/sheets/export.py`
  - Action: build Excel workbook(s) per faculty using canonical DB state. Supports atomic writes, dropdowns and conditional formatting driven by `settings.yaml`.
  - Modes: legacy exporter and optional workflow exporter (`--export-workflow`) that writes separate inbox/in_progress/done workbooks or sheets per-faculty.

## Important DB models (where data is stored)
- Staging tables
  - `StagedCopyrightItem` — raw, normalized rows
  - `StagedFacultyUpdate` — user-driven faculty updates staged for processing
  - `StagedProcessingFailure` — captures per-row failures and payload for retries

- Canonical tables
  - `CopyrightItem` — core item (material_id, url, filename, status, file_exists, workflow_status, filehash, last_scan_date_university, last_scan_date_course, ...)
  - `PDF` — per-item PDF metadata and download flags
  - `Course`, `Person`, `Organization`, `Programme`, `ItemUpdate` — related entities and change history

## File & folder map (most relevant)
- Data inputs
  - `cip_sheets/`, `raw_copyright_data/` — incoming export XLSX files
- Generated outputs
  - `faculty_sheets/` — generated Excel exports (per-faculty subfolders)
  - `faculty_sheets_old/` — archived exports
  - `pdf_downloads/` — downloaded PDFs
  - `full_backups/` — backup snapshots and manifests
- Database & runtime
  - `db.sqlite3`, additional DB backups in repo root
  - `logs/` — runtime logs and diagnostics (tests may write `hang_diagnostics.txt` on teardown issues)

## CLI flags and behaviour notes (runtime)
- Stage selection (mutually exclusive): `--ingest-only`, `--process-only`, `--export-only`, `--enrich-only`, `--file-exists-only`.
- `--disable-writes`: prevents (most) write operations and is useful for dry-run inspections; export is read-only by default when this is set.
- `--changes`: limit export to items that changed since last export (useful for incremental runs).
- `--single-faculty <ABBR>`: restrict processing/export to a single faculty.
- `--other-sheet <path>`: use a single XLSX input instead of repository input folders.
- `--export-workflow`: enable workflow-mode exporter (separate inbox/in_progress/done outputs); legacy exporter still available by default.

Run-time admin and backup commands (exposed in `run.py`):
- `backup create` / `backup restore` — CLI wrappers around `easy_access/sheets/backup.Backupper` (backup helpers implemented; manifested restores supported).
- `admin inspect-failures` / `admin retry-failures` / `admin cleanup-failures` — inspect and operate on `StagedProcessingFailure` entries.

## Contract (inputs / outputs / error handling)
- Inputs: XLSX export files and existing SQLite DB files.
- Outputs: updated canonical DB tables and Excel files in `faculty_sheets/`; optional PDFs and backups.
- Errors: per-row failures are persisted to `StagedProcessingFailure` so runs continue; admin CLI supports inspection and retries.

## Where to look for code-level traces
- Entrypoint and CLI wiring: `run.py` (builds `EasyAccessSettings`, instantiates `EasyAccessTool`).
- Tool implementation and stage orchestration: `easy_access/main.py` and `easy_access/pipeline.py`.
- Ingest: `easy_access/db/ingest.py`.
- Processing/updating DB: `easy_access/db/update.py`, `easy_access/merge_rules.py`.
- Enrichment: `easy_access/enrichment/osiris.py`.
- Relations and bulk linking: `easy_access/db/relations.py`.
- File checks: `easy_access/maintenance/file_existence.py`.
- Export generation: `easy_access/sheets/sheet.py`, `easy_access/sheets/export.py`.
- Backup helpers: `easy_access/sheets/backup.py`.

## Current status notes (what is implemented vs pending)
- Implemented / Available:
  - Full pipeline orchestration with async entrypoints and thin sync wrappers.
  - OSIRIS enrichment with TTL selection and concurrent fetches.
  - File existence TTL checks and bulk persistence.
  - Export generation with atomic writes, dropdowns and conditional formatting helpers.
  - Backup helper module and CLI backup/restore commands.
  - Admin CLI for failure inspection and retry (`StagedProcessingFailure`).

- Partially implemented / Pending / To do:
  - Pre/post pipeline automatic backup hooks: backup module is available, but fully automated pre/post pipeline wiring is configurable and still being refined.
  - Reactive faculty sheet workflow (new / to_check / checked) is available as an optional export-mode but operational conventions and reconcile scripts remain a roadmap item.
  - Some fields recently added to models (`filehash`, `last_scan_date_university`, `last_scan_date_course`) need thorough validation in ingestion and export flows — tests and edge-case handling are on the TODO list.
  - Documentation, types and docstrings: coverage needs improvement across the codebase.

## Notes for maintainers
- When changing export behaviour, verify `--disable-writes` and `--changes` semantics to avoid accidental DB writes.
- Use the admin CLI to triage per-row failures before re-running full pipeline.
- If adding or modifying enrichment parsers, add TTL selection unit tests to avoid accidental refresh storms.

## Quick checklist for a safe run
1. Commit or snapshot DB if you want to be able to rollback.
2. Run `uv sync` to ensure dependencies are installed.
3. Try a read-only dry run: `uv run run.py process --disable-writes`.
4. Inspect `admin inspect-failures` after a run if any issues are suspected.

---

This file is intentionally concise — consult individual modules listed above for detailed implementation and tests.
