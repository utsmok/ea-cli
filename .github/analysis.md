# Code Analysis & Background

This file contains the background, analysis, and rationale behind the refactor and dataflow redesign.

## Current Data Flow

The application processes copyright exports and faculty sheets, enriches data from external sources, computes derived fields, and generates reports. Historically, logic was spread across a monolithic `old_main.py`, `db/update.py`, and sheet helpers which caused side-effects and unclear priorities.

Key stages:
- Ingestion (raw exports, faculty sheets)
- Staging (staging tables in DB)
- Processing (merge/validation into main tables)
- Enrichment (external sources)
- Export (reports)

## What was refactored

- Centralized pipeline orchestration in `easy_access/pipeline.py` and simplified entrypoint in `easy_access/main.py`.
- Introduced staging tables (`StagedCopyrightItem`, `StagedFacultyUpdate`) and staging ingestion helpers in `easy_access/db/ingest.py`.
- Implemented minimal staged processors `process_staged_raw_data` and `process_staged_faculty_updates` in `easy_access/db/update.py` to provide a safe initial path from staging to main tables.

## Key design goals

- Make the DB the single source of truth.
- Unidirectional flow: Ingest -> Stage -> Process -> Export.
- Preserve legacy merge/priority rules but implement them in a testable, isolated module.
- Improve safety (transactions, atomic updates) and observability.

## Legacy code notes

- `old_main.py` contains the authoritative merge heuristics; preserve these by extracting them into `easy_access/merge_rules.py`.
- `db/base.py::copyright_item_from_dict` contains important normalization logic — keep and test this as the canonical factory.

## Tortoise ORM considerations

- Use `Tortoise.init(...)` and `Tortoise.generate_schemas(safe=True)` for idempotent setup.
- Use `async with in_transaction():` and savepoints for atomic staged processing.
- Prefer `QuerySet.update(...)` and `F` expressions for bulk DB-side updates.
- Use `prefetch_related(...)` for retrieval performance in FK/M2M loops.

## Recommendations (short)

- Harden staged processing, integrate `update_copyright_items` for full merges, and wrap processing in transactions.
- Replace `asyncio.run` in library code with async entrypoints and awrapper for CLI.
- Add unit & integration tests covering ingest -> stage -> process flows.

## Recent changes and operational notes (2025-09-02)

- Implemented staged-processing hardening in `easy_access/db/update.py`:
	- Replaced unsafe `__dict__` usage with an explicit `staged_fields` mapping.
	- Process staged rows in batches inside `async with in_transaction()` blocks.
	- Per-row try/except: only successfully processed staged rows are deleted after commit.
	- Per-row failures are now persisted to the DB in `staged_processing_failures` (model `StagedProcessingFailure`) for inspection and retry.

- Moved small parsing/normalization helpers into `easy_access/utils.py` (`safe_int`, `safe_float`, `safe_date`, `safe_enum`, `safe_compare_greater`) and updated modules to import them from there.

- Added unit tests for the helpers (`tests/test_safe_parsers.py`) and an integration test for staged processing (`tests/test_integration_staging.py`). Both pass locally in the development environment.

## Phase D Performance Optimizations (2025-03-01)

### Bulk M2M Linking Optimization
- **Problem**: N+1 query problem in `relations.py::link_courses` function causing individual database calls for each relationship
- **Solution**: Implemented bulk operations using raw SQL with temporary tables and `INSERT OR IGNORE` statements
- **Impact**: Significant performance improvement for large datasets with many course-copyright relationships
- **Files Modified**: `easy_access/db/relations.py`

### Export Memory Usage Optimization
- **Problem**: Complex correlated subqueries in `retrieve_full_data` causing memory issues and slow performance
- **Solution**: Replaced correlated subqueries with pre-aggregated CTEs (Common Table Expressions) and JOINs
- **Features Added**:
  - Memory usage monitoring and logging
  - Fallback to original method if optimization fails
  - Pre-aggregated course and contact data
- **Files Modified**: `easy_access/db/retrieve.py`

### File Exists Persistence Optimization
- **Problem**: Individual database updates for each file existence check causing N+1 query pattern
- **Solution**: Implemented bulk updates using temporary tables and raw SQL
- **Features Added**:
  - Temporary table-based bulk updates
  - Fallback to individual updates on error
  - Proper cleanup of temporary tables
- **Files Modified**: `easy_access/maintenance/file_existence.py`

### Rate Scheduling for API Calls
- **Problem**: No rate limiting for Canvas API calls risking rate limit violations
- **Solution**: Added configurable rate limiting with asyncio.sleep between requests
- **Configuration**: `file_exists_rate_limit_delay` setting (default: 0.1 seconds)
- **Files Modified**: `easy_access/maintenance/file_existence.py`, `easy_access/pipeline.py`, `easy_access/settings.py`

## Performance Optimization Patterns

### Bulk Database Operations
- Use temporary tables for complex bulk updates
- Prefer raw SQL over ORM for performance-critical bulk operations
- Implement fallback mechanisms for error handling
- Clean up temporary resources properly

### Memory-Efficient Data Processing
- Use CTEs to pre-aggregate data and avoid repeated computations
- Implement streaming/chunked processing for large datasets
- Monitor memory usage and log performance metrics
- Provide fallback methods for complex optimizations

### API Rate Limiting
- Implement configurable delays between API calls
- Use semaphores for concurrent request management
- Monitor request rates and adjust accordingly
- Provide settings for different environments

## Short-term next steps (priority order)

- Repo-wide safety sweep: replace remaining ad-hoc casts and `__dict__` usages with `safe_*` helpers and explicit mappings. Add unit tests for changed areas. (High priority)
- Implement a lightweight admin CLI to list `StagedProcessingFailure` rows and allow re-queuing for processing. (Medium)
- Add CI after the repo-wide safety sweep to avoid introducing regressions into mainline. (Low)

## Quick review of diffs vs starting commit (849628b9...)

I compared the current branch against commit `849628b965f4bd23b91400f1a5034eaf40787334` and found the following noteworthy changes:

- Files added: `easy_access/old_main.py`, `easy_access/pipeline.py` — legacy orchestrator preserved and new pipeline introduced.
- Files modified: `easy_access/main.py`, `easy_access/db/update.py`, `easy_access/db/ingest.py`, `easy_access/db/base.py`, `easy_access/db/models.py`, `easy_access/retrieve.py`, multiple sheet helpers and utilities.

Summary of risks and follow-ups discovered by the diff:

- `easy_access/main.py` was heavily simplified; ensure the new CLI entrypoint still calls `ensure_db_inited()` and closes Tortoise on exit.
- `process_staged_raw_data` exists but is conservative; route complex staged rows into the richer `update_copyright_items` path.
- Some files reference `__dict__` on model objects; replace with explicit safe mapping helpers to avoid leaking internal state or unserializable objects.
- Add tests that assert staging is only cleared after successful processing (per-row savepoints or batch atomicity).

I added a top-level todo item in `.github/todo.md` ("Critical review of refactor changes") to formalize this review and track any follow-ups. Add further follow-up items to that list as you find concrete code fixes.
