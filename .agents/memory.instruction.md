---
applyTo: '**'
---

# Coding preferences
- Ruff/formatting: repo uses ruff with an 88-character line length and targets Python 3.12 (see `pyproject.toml`). Keep formatting consistent with existing config.
- Typing: Add or improve type hints pragmatically; many dataclasses are typed but full coverage is not required. Prefer readable, explicit types for public functions.
- Logging: `loguru` is used; `configure_logger()` is invoked from `easy_access.settings` on startup.
- Async-first: I/O and pipeline stages are async; expose sync wrappers where needed with `run_sync` (present in `easy_access.utils`).
- CLI: `run.py` (Typer) is the canonical CLI. README suggests `uv run run.py process` as a common invocation.
- Testing: Pytest (and pytest-asyncio) exist but tests are not guaranteed green; use tests as guidance and add small, focused tests for new code.

# Project architecture
- Language/runtime: Python 3.12.2 (pinned in `pyproject.toml`).
- Root package: `easy_access/` — contains pipeline, DB models, enrichment, PDF handling, maintenance, and exports.
- Key components:
  - `pipeline.py`: async pipeline stages and sync wrappers.
  - `settings.py`: dataclasses (`Settings`, `EasyAccessSettings`), YAML parsing, logger setup.
  - `db/`: Tortoise ORM models and DB helpers (ingest/update/retrieve/relations). **New: SQLAlchemy models in `sa_models.py`, session management in `session.py`, base in `models_base.py`.**
  - `enrichment/`: OSIRIS enrichment and PDF-backed enrichment helpers.
  - `sheets/`: Excel sheet builders, backups, overviews and export orchestration.
  - `maintenance/`: Canvas file existence checks, v1 ingestion helpers, cleanup utilities.
  - `pdf/`: Canvas downloads, extraction/parsing, and PDF-related models.
- Configuration: `settings.yaml` at repo root is authoritative; per-faculty and per-run overrides supported.
- DB: SQLite (default `db.sqlite3`) with Tortoise ORM; several DB snapshots in repo for restoration/testing. **Migration in progress to PostgreSQL 18 with SQLAlchemy 2.0 (async) and Alembic migrations.**
- Performance: uses bulk DB ops, `polars` for DataFrame work, TTL policies, semaphores and async concurrency.

## Solutions & common patterns
- Async + sync wrappers: implement async logic and expose a `run_sync` wrapper for CLI consumers (pattern used across pipeline stages).
- Settings dataclasses: `Settings`/`EasyAccessSettings` parse YAML and provide typed access; prefer small, targeted additions rather than sweeping changes.
- File helpers: `Directory`/`File` wrappers centralize path handling, creation and listing logic — use these rather than ad-hoc pathlib use.
- Exports: `ColInfo`/`StyleInfo` dataclasses plus `polars` + openpyxl/xlsxwriter are used for fast processing and precise Excel formatting.

# Constraints & environment
- Python pinned to 3.12.2 — keep CI and local dev interpreters aligned.
- `uv` is used for environment management/running; follow existing scripts and README guidance.
- NEVER run python or alembic or pip directly; always use `uv` to ensure the correct environment is used. use `uv run <command>` to run commands in the correct environment.
- Heavy optional deps (OCR/LLM/PyTorch) are grouped; avoid installing unless working on related features.

# Files inspected (representative)
- `pyproject.toml`, `README.md` — project config and usage notes.
- `easy_access/main.py`, `easy_access/settings.py`, `easy_access/pipeline.py`, `easy_access/utils.py` — core orchestration & helpers.
- `easy_access/db/*` — models, ingest/update/retrieve/relations.
- `easy_access/sheets/*`, `easy_access/enrichment/*`, `easy_access/pdf/*`, `easy_access/maintenance/*`, `easy_access/classification/*` — full pipeline coverage.

## Conventions & practical constraints
- Canonical `settings.yaml` is authoritative; per-faculty/run overrides supported via `OverrideSettings` and `EasyAccessSettings.create_for_runtime`.
- Secret discovery order: environment variables → `.env`/`.secret` files (search up to 2 parents) → `api_keys.py` module (repo root or `easy_access/api_keys.py`). Prefer env vars for CI.
- Runtime: Windows-first manual runs; `run.py` (Typer) is the CLI entrypoint; recommended `uv run run.py process` in README.
- DB migrations: ad-hoc/manual; datasets can be reconstructed from Excel exports when necessary.
- Typing/linting: Use ruff/pyright; prefer pragmatic incremental typing and clear inline docs.
- Priorities: performance and documentation first; tests and full typing second.

## Repository summary
- Scope: `easy_access/` implements the full data pipeline for copyright items: ingestion, enrichment, DB persistence, PDF download & parsing, maintenance tasks, and Excel exports.
- Config & secrets: `settings.yaml` is authoritative; secret lookup order explained above; prefer environment variables for CI and ephemeral runs.
- Runtime & platform: Windows-focused usage, CLI via `run.py` (Typer), async-first with sync wrappers for CLI usage.
- DB: SQLite + Tortoise ORM; polars + SQLAlchemy engine used for read-heavy exports and grouping.
- Exports: polars -> Excel (openpyxl/xlsxwriter), protected `done`/`overview` workbooks, backup strategy per-faculty.
- Enrichment/PDF/maintenance/classification: OSIRIS enrichment, GLiNER-based NER, kreuzberg extraction, Canvas file checks are all present and integrated.
- Tests: present but may not pass; useful for examples and regression checks but not strict gating.

## Files & responsibilities (final checklist)
- `run.py`: Typer CLI entrypoint; constructs runtime settings and runs selected pipeline stages.
- `easy_access/settings.py`: typed settings and secret discovery logic (`SETTINGS` global).
- `easy_access/pipeline.py`: async pipeline orchestration and sync bridges.
- `easy_access/utils.py`: `run_sync`, safe parsers, `Directory`/`File`, `standardize_dataframe`.
- `easy_access/db/*`: models, ingest/update/retrieve/relations and merge rules.
- `easy_access/enrichment/*`: OSIRIS enrichment, PDF-backed enrichment helpers.
- `easy_access/pdf/*`: Canvas download, parsing, and PDF-related models.
- `easy_access/sheets/*`: sheet builders, backups, overviews, and export orchestration.
- `easy_access/maintenance/*`: Canvas file checks, v1 ingestion helpers, cleaning utilities.
- `easy_access/classification/*`: NER helpers and experimental classification/LLM code.
- `easy_access/api_keys.py`: optional per-device token fallback; do not store production secrets here.

## Development guidance
- Implement async logic with `run_sync` wrappers for CLI usage.
- Preserve dataclass shapes and add focused unit tests for public helpers.
- Use `polars` for large dataframe transformations and prefer bulk DB updates/inserts to avoid N+1 queries.
- Use `Directory`/`File` wrappers to abstract filesystem differences and keep code cross-platform.

## Additional findings from scanned files (ingest/update/retrieve/relations/sheets/enrichment/maintenance/pdf)

Below are compact notes from the remaining files I scanned (no code changes were made):

1) run.py
- CLI entrypoint for daily usage. Provides `process`, `update-from-v1`, and `dashboard` commands.
- Uses Typer and includes a small compatibility shim for Click/Rich parameter signature mismatches.
- Stage selection is handled with explicit boolean flags and constructs an `EasyAccessSettings` runtime object before invoking `EasyAccessTool`.
- Error handling: critical errors in `ingest`, `process`, or `db_changes` abort the workflow; other stages are skipped on failure.

2) `easy_access/db/ingest.py`
- Functions to load base/org data, load PDFs from the PDF directory, and load raw copyright/faculty data into staging tables.
- Uses `standardize_dataframe` before creating `Staged*` objects. Bulk upserts use `bulk_create(..., on_conflict=["material_id"], update_fields=...)`.
- Defensive guards against missing fields for creation of `CopyrightItem`.

3) `easy_access/db/update.py`
- Complex merge/update logic for CopyrightItems, using strategy pattern (ranked/string/numeric/date/enum/file_exists).
- Preprocessing separates new vs. existing items, enforces required fields for creation, and records small changes to `ItemUpdate` changelogs.
- Provides `persist_courses` and `persist_persons` helpers used by enrichment (test-friendly with dependency injection).
- Includes `map_v1_to_v2_classifications`, `calculate_derived_fields`, and staged processing functions that record persistent failures in `StagedProcessingFailure`.

4) `easy_access/db/retrieve.py`
- DB retrieval helpers centered on `polars` reads via a SQLAlchemy engine created with `init_engine(settings)`.
- `retrieve_full_data` has an optimized variant using CTEs and GROUP_CONCAT / JSON aggregation to pre-join course/person/contact info; falls back to original query on failure.
- Many retrieval helpers require `settings` to construct engine and validate faculties.

5) `easy_access/db/relations.py`
- Orchestrates relation updates (course linking, v1 matching, person-course linking) with a test-aware resolution helper `_resolve_queryset_candidate()`.
- Course linking extracts possible course codes via `determine_course_code`, batches DB fetches, and uses either ORM M2M `add()` or (in tests) `bulk_update` mocks.
- Person-course linking is batched and avoids N+1 via bulk prefetches. All operations are defensive for test mocks and missing DB initialization.

6) `easy_access/sheets/sheet.py`, `backup.py`, `analysis.py`
- `sheet.py` contains Excel helpers: quiet Excel reads, `DataEntrySheet` helper for building data entry sheets with dropdowns, `finalize_sheet`, `store_complete_data`, and `protect_workbook`.
- `backup.py` centralizes timestamped backup moves and optional manifest writing.
- `analysis.py` (overview generation) coordinates creation of faculty overviews, backs up older overview files, and optionally persists export-derived updates to the DB (gated by `disable_writes`).

7) `easy_access/enrichment/osiris.py`
- OSIRIS enrichment orchestrator: gathers course/person targets, selects missing/stale entities using TTLs, fetches concurrently with `httpx` + semaphores, and persists results using `persist_courses`/`persist_persons`.
- Robust parsing for course search and detailed fetch; uses Levenshtein scoring to match person tiles and has retry/backoff mechanics.

8) `easy_access/maintenance/file_existence.py`
- TTL-based file existence verifier against Canvas API. Selects items needing checks, queries Canvas `files` and `folders` endpoints, determines `canvas_course_id` from folder metadata, and updates `CopyrightItem.file_exists` + last check using bulk updates with a test-aware fallback.

9) `easy_access/pdf/download.py` and `easy_access/pdf/parse.py`
- Download: token-authenticated Canvas API downloads, creation/upsert of `PDFCanvasMetadata`, and storing files under configured `pdf_downloads` directory. Uses streaming downloads and rate-limit checks.
- Parse: Uses `kreuzberg` for extraction, stores extracted text in `PDFText`, computes `filehash` with xxhash, and records extraction metadata. OCR path currently disabled/placeholder.

These notes are reflected in the "Files inspected" and the utilities / schema / export sections above.


## Utilities & common helpers (from `easy_access/utils.py`)
- run_sync(coro): Async<->sync bridge that detects running event loop and, if present, executes the coroutine in a background thread via ThreadPoolExecutor to avoid "event loop already running" errors. Use this pattern for CLI wrappers.
- safe_* helpers: `safe_int`, `safe_float`, `safe_date`, `safe_enum`, `safe_compare_greater` — robust parsing helpers to normalize heterogeneous inputs from Excel/Canvas/OSIRIS.
- determine_course_code(code, name): Heuristic parser to extract numeric Osiris course codes from Canvas course_code and course_name fields. It validates codes by being numeric and length >= 8.
- standardize_dataframe(df: pl.DataFrame): Normalizes column names (lowercase, replace spaces and special chars), casts non-string columns to str, replaces '-' with None, filters missing material_id and undesirable filetypes, drops unwanted columns. Used before ingestion.
- Directory & File classes: High-level wrappers around pathlib for directory/file manipulation with convenience methods (files, files_r, newest_file, copy/move, create/delete). They encapsulate create-on-init semantics and return typed `File` objects. Use them across export/backup logic.

## Database models & schema notes (from `easy_access/db/models.py`)
- ORM: Uses Tortoise ORM with many models; `TimestampMixin` provides `created_at` and `modified_at`.
- Core entities:
  - `CopyrightItem` (core table): workflow fields (file_exists, last_canvas_check, workflow_status), relations to `Course`, `Faculty`, `ItemUpdate` and `PDF`.
  - `PDF`, `PDFCanvasMetadata`, `PDFText`, `Entity` — detailed PDF storage and extracted text/entities, with `filehash` and parsing metadata.
  - `Course`, `Person`, `Faculty`/`Organization`, `Programme` — enrichment data with relations and M2M relationships (teachers -> Person).
  - Staging tables: `StagedCopyrightItem`, `StagedFacultyUpdate`, `StagedProcessingFailure` used for robust staged ingestion and persisted failures.
- Indexing & datatypes: Many fields use CharEnumField for controlled enums; text fields often have large max lengths; some fields are indexed (db_index=True). File-related fields include unique url constraint on `CopyrightItem.url`.
- Notes: Schema evolves ad-hoc; keep migrations simple. New fields like `filehash`, `last_scan_date_university/course` exist on models and are part of the expected dataflow.

## Export patterns & behavior (from `easy_access/sheets/export.py`)
- Data retrieval: `gather_faculty_data` calls `retrieve_full_data(settings)` (DB-first) and performs column normalizations (e.g., `file_exists` -> Yes/No and `course_link` construction using `canvas_course_id`). Grouping is by `faculty`.
- Export modes:
  - Legacy per-faculty sheet generation (`export_faculty_sheets`) — writes a date-stamped file per faculty, determines "new items" by comparing material_id to existing Complete Data sheets in the faculty directory.
  - Workflow-based exports (`export_faculty_workflow_files`) — produces `inbox.xlsx`, `in_progress.xlsx`, `done.xlsx`, plus an `overview.xlsx`. Existing files are backed up to a faculty-specific backups folder. `done` and `overview` are protected workbooks after writing.
  - Overview exports (`export_faculty_overviews`) — delegated to `create_faculty_overviews` in `sheets/analysis.py`.
- Key helpers: `store_complete_data`, `finalize_sheet`, `protect_workbook`, `_get_unique_filepath`, and `backup_existing_file` are the key building blocks; they use `Directory`/`File` wrappers and `polars` for fast DataFrame processing.
- Per-faculty overrides: The exporter checks for `.yml` override files in faculty directories and will create an `OverrideSettings` object if found before writing that faculty's exports.

## Practical constraints & conventions (incorporating your answers)
- Canonical settings file: `settings.yaml` (root) is authoritative; per-faculty or per-run overrides are supported via override YAMLs and `OverrideSettings`.
- Secrets & API keys: `easy_access.settings` supports env vars, `.env/.secret` files (searched up to 2 parent levels), and `api_keys.py` modules (either in repo root or `easy_access/api_keys.py`). Prefer env vars for CI, per-device `api_keys.py` is acceptable locally.
- Platform: Primary development/run platform is Windows (manual runs). Keep Windows path semantics and file creation behavior in mind (but code mostly uses pathlib + wrappers so is cross-platform friendly).
- Tests & migrations: Tests are low-priority; schema changes are performed manually and datasets can be rebuilt from Excel when necessary. Keep migrations simple and backwards-compatible where possible.
- Typing & tooling: Continue using type hints where practical and run pyright/ruff/ty; avoid spending excessive time on tricky typing issues.

## Recent migration preferences (recorded Oct 13, 2025)

- Target DB: PostgreSQL (standardize; no SQLite fallback)
- Postgres version: 18
- Driver: asyncpg
- Data migration tool: pgloader (user chose pgloader; a Docker wrapper script is recommended)
- Alembic migrations location: `migrations/` in the repo root
- Test scaffold for this migration: skip (user requested no test scaffold)

### Migration Progress & Validation (Oct 13, 2025 - Updated Session 2)

**Completed Phases:**
- Phase 1: Dependencies, Docker environment, .env setup ✅
- Phase 2: SQLAlchemy foundation (session.py, models_base.py, compat.py) ✅
- Phase 3: Model conversion (sa_models.py with full SQLAlchemy equivalents) ✅
- Phase 4: Alembic setup (alembic.ini, initial migration generated and applied) ✅
- Phase 5: Data migration (pgloader migrated 27,343 rows from SQLite to PostgreSQL) ✅

**Data Migration Verification Test:**
- **Test Performed:** Ran complex SQL query `retrieve_full_data` (with CTEs for course/person aggregations) on both SQLite (`db.sqlite3`) and PostgreSQL databases, adding `ORDER BY cd.material_id` for consistent ordering.
- **Query Adaptation:** For PostgreSQL, replaced `GROUP_CONCAT` with `STRING_AGG` and added `::text` casts to avoid type errors.
- **Results:**
  - **Rows:** 2,512 (identical)
  - **Columns:** 54 (identical)
  - **Types:** Compatible (object vs datetime64[ns] for dates)
  - **Data:** Identical except for expected ordering differences in aggregated strings (no ORDER BY in STRING_AGG/GROUP_CONCAT)
- **Conclusion:** Migration successful - all data integrity preserved.

**Phase 6 - Application Code Migration (In Progress - 29% Complete):**

**✅ Completed Modules (4/14):**
1. **compat.py** - Comprehensive compatibility layer with 20+ functions
   - All Tortoise patterns have SQLAlchemy equivalents
   - Key functions: bulk_create/bulk_update (with upsert), get_or_create, update_or_create, transaction context, filter/count helpers
   - Handles race conditions, batching, PostgreSQL-specific upserts

2. **session.py** - Async session management
   - Modified init_db() to accept Settings object
   - Extracts DATABASE_URL from environment
   - Provides get_session() async generator pattern

3. **base.py** - Database initialization
   - Converted from Tortoise.init() to SQLAlchemy init_db()
   - Replaced generate_schemas() with Base.metadata.create_all()
   - Minor cleanup pending: copyright_item_from_dict() type hint

4. **ingest.py** - Data ingestion (all 5 functions)
   - load_org_data_from_settings(), load_base_data(), load_pdfs(), load_raw_copyright_data_to_staging(), load_faculty_updates_to_staging()
   - All Tortoise calls replaced with compat layer functions
   - 3 minor non-critical lint warnings remain (type checking)

**🔄 In Progress (1/14):**
5. **relations.py** - M2M relationship management (500+ lines)
   - Conversion guide created in `.agents/relations_conversion_guide.md`
   - Needs: selectinload for prefetch, association table inserts for M2M, transaction context
   - Functions: link_courses(), link_persons_to_courses(), match_v1_to_copyright_items()

**📋 Remaining Modules (9/14):**
6. **update.py** - Complex merge logic (1,734 lines) - highest priority
7. **retrieve.py** - Mixed Tortoise/SQLAlchemy (827 lines) - cleanup mostly
8. **pdf/download.py** - Simple CRUD (straightforward)
9. **pdf/parse.py** - Simple CRUD (straightforward)
10. **enrichment/osiris.py** - M2M operations (similar to relations.py)
11. **maintenance/file_existence.py** - Simple CRUD (straightforward)
12. **maintenance/v1_items.py** - Simple CRUD (straightforward)
13. **Dependency cleanup** - Remove tortoise-orm from pyproject.toml
14. **Documentation** - Update README.md with PostgreSQL setup instructions

**Session 2 Deliverables:**
- Created comprehensive conversion guide for relations.py with code examples
- Updated memory and TODO tracking
- Documented M2M patterns (association table inserts vs ORM .add())
- Documented prefetch patterns (selectinload vs prefetch_related)
- Preserved project state with detailed status documentation

### TortoiseORM Usage Analysis (Oct 13, 2025)

**Files with Tortoise Usage:**
- `db/base.py`: Tortoise.init/generate_schemas/close_connections, Model.get_or_create
- `db/retrieve.py`: Tortoise, Q expressions, Model.filter/all/get/values_list
- `db/update.py`: Tortoise, Q, in_transaction, bulk_create/bulk_update, CRUD operations
- `db/relations.py`: in_transaction, filter, prefetch_related, M2M .add()
- `db/ingest.py`: Tortoise, bulk_create, get_or_create, all/values
- `db/models.py`: tortoise fields, Model class inheritance
- `pdf/download.py`: get_or_none, create, filter
- `pdf/parse.py`: all, create, save
- `maintenance/v1_items.py`: update_or_create, all, values_list, filter
- `enrichment/osiris.py`: all, filter, distinct, get_or_none, M2M .add(), delete
- `maintenance/file_existence.py`: filter, get, update

**Common Patterns:**
- **Querying:** Model.filter().all(), Model.get/get_or_none()
- **Bulk Ops:** bulk_create, bulk_update
- **Transactions:** in_transaction (inconsistent usage)
- **M2M Relations:** .add()/.remove()
- **Expressions:** Q objects for complex queries
- **CRUD:** create, save, update, delete

**Areas for Improvement During Migration:**
1. **Repeated Queries:** Abstract common filter/all patterns into repository methods
2. **Transaction Management:** Standardize async transaction usage with SQLAlchemy
3. **Error Handling:** Add comprehensive try/catch with proper logging
4. **Business Logic Separation:** Extract DB queries from business logic into service layers
5. **N+1 Prevention:** Ensure all queries use proper joins/prefetching
6. **Batch Processing:** Standardize batch sizes and processing patterns
7. **Logging:** Add consistent operation logging and metrics

**Specific Refactor Opportunities:**
- `update.py`: Break down complex staged processing logic into smaller functions
- `relations.py`: Simplify queryset resolution logic
- `ingest.py`: Improve error handling in bulk operations
- `retrieve.py`: Unify SQLAlchemy engine vs Tortoise ORM usage patterns
