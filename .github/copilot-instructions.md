# GitHub Copilot Instructions for Easy Access CLI

This repository contains the Easy Access Sheet Toolkit, a Python application with a built-in CLI for automating the processing, enrichment, and export of copyright data from university systems.

## Coding Preferences

- **Formatting**: Use Ruff with an 88-character line length targeting Python 3.12 (see `pyproject.toml`). Keep formatting consistent with existing config.
- **Type Hints**: Add or improve type hints pragmatically; many dataclasses are typed but full coverage is not required. Prefer readable, explicit types for public functions.
- **Logging**: Use `loguru`; `configure_logger()` is invoked from `easy_access.settings` on startup.
- **Async-First**: I/O and pipeline stages are async; expose sync wrappers where needed with `run_sync` (present in `easy_access.utils`).
- **CLI**: `run.py` (Typer) is the canonical CLI. README suggests `uv run run.py process` as a common invocation.
- **Testing**: Pytest (and pytest-asyncio) exist but tests are not guaranteed green; use tests as guidance and add small, focused tests for new code.

## Project Architecture

- **Language/Runtime**: Python 3.12.2 (pinned in `pyproject.toml`)
- **Root Package**: `easy_access/` — contains pipeline, DB models, enrichment, PDF handling, maintenance, and exports
- **Key Components**:
  - `pipeline.py`: async pipeline stages and sync wrappers
  - `settings.py`: dataclasses (`Settings`, `EasyAccessSettings`), YAML parsing, logger setup
  - `db/`: Tortoise ORM models and DB helpers (ingest/update/retrieve/relations)
  - `enrichment/`: OSIRIS enrichment and PDF-backed enrichment helpers
  - `sheets/`: Excel sheet builders, backups, overviews and export orchestration
  - `maintenance/`: Canvas file existence checks, v1 ingestion helpers, cleanup utilities
  - `pdf/`: Canvas downloads, extraction/parsing, and PDF-related models
- **Configuration**: `settings.yaml` at repo root is authoritative; per-faculty and per-run overrides supported
- **Database**: SQLite (default `db.sqlite3`) with Tortoise ORM; several DB snapshots in repo for restoration/testing
- **Performance**: Uses bulk DB ops, `polars` for DataFrame work, TTL policies, semaphores and async concurrency

## Solutions & Common Patterns

- **Async + Sync Wrappers**: Implement async logic and expose a `run_sync` wrapper for CLI consumers (pattern used across pipeline stages)
- **Settings Dataclasses**: `Settings`/`EasyAccessSettings` parse YAML and provide typed access; prefer small, targeted additions rather than sweeping changes
- **File Helpers**: `Directory`/`File` wrappers centralize path handling, creation and listing logic — use these rather than ad-hoc pathlib use
- **Exports**: `ColInfo`/`StyleInfo` dataclasses plus `polars` + openpyxl/xlsxwriter are used for fast processing and precise Excel formatting

## Constraints & Environment

- **Python Version**: Pinned to 3.12.2 — keep CI and local dev interpreters aligned
- **Package Manager**: `uv` is used for environment management/running; follow existing scripts and README guidance
- **Optional Dependencies**: Heavy optional deps (OCR/LLM/PyTorch) are grouped; avoid installing unless working on related features

## Conventions & Practical Constraints

- **Settings**: Canonical `settings.yaml` is authoritative; per-faculty/run overrides supported via `OverrideSettings` and `EasyAccessSettings.create_for_runtime`
- **Secret Discovery Order**: environment variables → `.env`/`.secret` files (search up to 2 parents) → `api_keys.py` module (repo root or `easy_access/api_keys.py`). Prefer env vars for CI
- **Runtime**: Windows-first manual runs; `run.py` (Typer) is the CLI entrypoint; recommended `uv run run.py process` in README
- **DB Migrations**: Ad-hoc/manual; datasets can be reconstructed from Excel exports when necessary
- **Typing/Linting**: Use ruff/pyright; prefer pragmatic incremental typing and clear inline docs
- **Priorities**: Performance and documentation first; tests and full typing second

## Repository Summary

- **Scope**: `easy_access/` implements the full data pipeline for copyright items: ingestion, enrichment, DB persistence, PDF download & parsing, maintenance tasks, and Excel exports
- **Config & Secrets**: `settings.yaml` is authoritative; secret lookup order explained above; prefer environment variables for CI and ephemeral runs
- **Runtime & Platform**: Windows-focused usage, CLI via `run.py` (Typer), async-first with sync wrappers for CLI usage
- **Database**: SQLite + Tortoise ORM; polars + SQLAlchemy engine used for read-heavy exports and grouping
- **Exports**: polars → Excel (openpyxl/xlsxwriter), protected `done`/`overview` workbooks, backup strategy per-faculty
- **Enrichment/PDF/Maintenance/Classification**: OSIRIS enrichment, GLiNER-based NER, kreuzberg extraction, Canvas file checks are all present and integrated
- **Tests**: Present but may not pass; useful for examples and regression checks but not strict gating

## Key Files & Responsibilities

- `run.py`: Typer CLI entrypoint; constructs runtime settings and runs selected pipeline stages
- `easy_access/settings.py`: typed settings and secret discovery logic (`SETTINGS` global)
- `easy_access/pipeline.py`: async pipeline orchestration and sync bridges
- `easy_access/utils.py`: `run_sync`, safe parsers, `Directory`/`File`, `standardize_dataframe`
- `easy_access/db/*`: models, ingest/update/retrieve/relations and merge rules
- `easy_access/enrichment/*`: OSIRIS enrichment, PDF-backed enrichment helpers
- `easy_access/pdf/*`: Canvas download, parsing, and PDF-related models
- `easy_access/sheets/*`: sheet builders, backups, overviews, and export orchestration
- `easy_access/maintenance/*`: Canvas file checks, v1 ingestion helpers, cleaning utilities
- `easy_access/classification/*`: NER helpers and experimental classification/LLM code
- `easy_access/api_keys.py`: optional per-device token fallback; do not store production secrets here

## Development Guidance

- Implement async logic with `run_sync` wrappers for CLI usage
- Preserve dataclass shapes and add focused unit tests for public helpers
- Use `polars` for large dataframe transformations and prefer bulk DB updates/inserts to avoid N+1 queries
- Use `Directory`/`File` wrappers to abstract filesystem differences and keep code cross-platform

## Utilities & Common Helpers (from `easy_access/utils.py`)

- **run_sync(coro)**: Async/sync bridge that detects running event loop and, if present, executes the coroutine in a background thread via ThreadPoolExecutor to avoid "event loop already running" errors. Use this pattern for CLI wrappers.
- **safe_* helpers**: `safe_int`, `safe_float`, `safe_date`, `safe_enum`, `safe_compare_greater` — robust parsing helpers to normalize heterogeneous inputs from Excel/Canvas/OSIRIS
- **determine_course_code(code, name)**: Heuristic parser to extract numeric Osiris course codes from Canvas course_code and course_name fields. It validates codes by being numeric and length >= 8
- **standardize_dataframe(df: pl.DataFrame)**: Normalizes column names (lowercase, replace spaces and special chars), casts non-string columns to str, replaces '-' with None, filters missing material_id and undesirable filetypes, drops unwanted columns. Used before ingestion
- **Directory & File classes**: High-level wrappers around pathlib for directory/file manipulation with convenience methods (files, files_r, newest_file, copy/move, create/delete). They encapsulate create-on-init semantics and return typed `File` objects. Use them across export/backup logic

## Database Models & Schema (from `easy_access/db/models.py`)

- **ORM**: Uses Tortoise ORM with many models; `TimestampMixin` provides `created_at` and `modified_at`
- **Core Entities**:
  - `CopyrightItem` (core table): workflow fields (file_exists, last_canvas_check, workflow_status), relations to `Course`, `Faculty`, `ItemUpdate` and `PDF`
  - `PDF`, `PDFCanvasMetadata`, `PDFText`, `Entity` — detailed PDF storage and extracted text/entities, with `filehash` and parsing metadata
  - `Course`, `Person`, `Faculty`/`Organization`, `Programme` — enrichment data with relations and M2M relationships (teachers → Person)
  - Staging tables: `StagedCopyrightItem`, `StagedFacultyUpdate`, `StagedProcessingFailure` used for robust staged ingestion and persisted failures
- **Indexing & Data Types**: Many fields use CharEnumField for controlled enums; text fields often have large max lengths; some fields are indexed (db_index=True). File-related fields include unique url constraint on `CopyrightItem.url`
- **Schema Evolution**: Schema evolves ad-hoc; keep migrations simple. New fields like `filehash`, `last_scan_date_university/course` exist on models and are part of the expected dataflow

## Export Patterns & Behavior (from `easy_access/sheets/export.py`)

- **Data Retrieval**: `gather_faculty_data` calls `retrieve_full_data(settings)` (DB-first) and performs column normalizations (e.g., `file_exists` → Yes/No and `course_link` construction using `canvas_course_id`). Grouping is by `faculty`
- **Export Modes**:
  - Legacy per-faculty sheet generation (`export_faculty_sheets`) — writes a date-stamped file per faculty, determines "new items" by comparing material_id to existing Complete Data sheets in the faculty directory
  - Workflow-based exports (`export_faculty_workflow_files`) — produces `inbox.xlsx`, `in_progress.xlsx`, `done.xlsx`, plus an `overview.xlsx`. Existing files are backed up to a faculty-specific backups folder. `done` and `overview` are protected workbooks after writing
  - Overview exports (`export_faculty_overviews`) — delegated to `create_faculty_overviews` in `sheets/analysis.py`
- **Key Helpers**: `store_complete_data`, `finalize_sheet`, `protect_workbook`, `_get_unique_filepath`, and `backup_existing_file` are the key building blocks; they use `Directory`/`File` wrappers and `polars` for fast DataFrame processing
- **Per-Faculty Overrides**: The exporter checks for `.yml` override files in faculty directories and will create an `OverrideSettings` object if found before writing that faculty's exports

## Migration Notes

**Note**: There is a planned migration from SQLite + Tortoise ORM to PostgreSQL + SQLAlchemy. See `.agents/migration_to_postgres_sqlalchemy.md` for the detailed migration plan. Key decisions:
- Target DB: PostgreSQL 18
- Driver: asyncpg
- Migration tool: pgloader
- Migrations location: `migrations/` in repo root
- No SQLite fallback (standardize on PostgreSQL)

When working on database-related code, be aware of this upcoming migration and prefer patterns that will be compatible with SQLAlchemy where practical.
