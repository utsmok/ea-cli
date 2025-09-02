# Export, Enrichment & File Existence Integration Plan

Date: 2025-09-02

This document analyses the missing functional areas (Excel exports, OSIRIS & course enrichment, file existence verification) after the dataflow refactor and defines a concrete, testable reintegration plan aligned with the new staging -> processing architecture.

## 1. Current State (Post‑Refactor) - Updated Analysis
Pipeline stages implemented:
- ingest_raw_data_async -> loads raw export into staging
- ingest_faculty_updates_async -> loads data-entry edits into staging
- process_data_async -> merges staged rows into main tables via refactored merge logic

**Critical Findings from Code Review**:
- **Relations functions exist** but need extraction: `link_courses_to_copyright_items()` and `update_duplicate_status()` are in `db/update.py` with N+1 query patterns (await per item in loop).
- **Export functions completely missing**: Legacy `old_main.py` references `create_export_sheet` but function doesn't exist in current codebase; all export orchestration removed.
- **Pipeline export stage stubbed**: `pipeline.py` shows `# await self.export_reports_async() # To be implemented`.
- **Sheet formatting mature**: `DataEntrySheet` class fully functional with dropdown validation, width calculation, table styling.
- **Legacy export workflow complex**: `old_main.py` shows sophisticated sheet types (faculty, program, overview, all_items) with file uniqueness handling.

Not yet implemented / incomplete:
- **Complete export stage missing**: No faculty sheets, programme sheets, overview sheets, or all_items sheet generation.
- **No export orchestrator module**: Legacy had `create_faculty_sheets`, `create_programme_sheets`, `create_overviews`, `create_all_items_sheet` functions.
- OSIRIS enrichment logic (`update_osiris_data`) still a large legacy function; not orchestrated by pipeline.
- Relations linking exists but optimizations needed and pipeline integration missing.
- File existence check lives in `utilities/file_exists.py`, only invoked in legacy flow; not incremental or TTL-aware.

## 2. Design Principles
1. Deterministic, idempotent stages: Each stage can be re-run safely.
2. Separation of concerns: Retrieval, transformation, persistence, and export are distinct modules.
3. Observability: Structured logging (counts, durations, deltas), minimal broad exception masking.
4. Incremental where possible: Re-check only missing/expired file existence or OSIRIS data.
5. Settings-driven: Concurrency limits, TTLs, enable/disable flags.
6. Testability: Unit + integration tests with small fixtures, no reliance on external network for core logic (mock httpx).

## 3. Proposed New/Updated Modules
| Area | Module | Responsibility |
|------|--------|---------------|
| Export Orchestration | `easy_access/sheets/export.py` | High-level export orchestrator (faculties, programmes, all items) invoking existing sheet helpers. |
| OSIRIS Enrichment | `easy_access/enrichment/osiris.py` (new subpkg) | Split retrieval (courses), person enrichment, merge, persistence, TTL logic. |
| Relations Update | `easy_access/db/relations.py` | Duplicate detection, course linking, future: staff/course person linkage. |
| File Existence | `easy_access/maintenance/file_existence.py` | Incremental file existence refresh + TTL policy wrapper around existing core checker. |
| Pipeline | `easy_access/pipeline.py` | Add async stages: `enrich_async()`, `update_relations_async()`, `refresh_file_existence_async()`, `export_reports_async()` with sync wrappers. |

## 4. Export Stage Detailed Plan
### Inputs
- DB (post-processing) via retrieval functions (`retrieve_full_data`, or a new optimized retrieval returning per-faculty partitions).

### Outputs
- Per faculty overview sheet (Complete Data + Data Entry sheet) – file naming: `{FACULTY}_total_overview_updated_{DATE}.xlsx`.
- Per programme sheets under `faculty/per_programme/` (grouping by course_mapping).
- All items sheet (optionally only new items since last run – later improvement).

### Implementation Steps
1. Create `sheets/export.py` with functions:
   - `gather_faculty_data(settings) -> dict[str, pl.DataFrame]` (DB fetch + fac filtering).
   - `export_faculty_overviews(settings, faculty_data)` (leverages existing `create_faculty_overviews`).
   - `export_all_items(settings)`.
2. Refactor duplicated logic now in legacy `old_main.py` but ensure use of standardized functions in `sheets/analysis.py` & `sheets/sheet.py` (minimal changes – adapt to DB‑first flow).
3. Add `export_reports_async()` to pipeline calling the above.
4. Add CLI flag(s) (already partial: `--export-only`) to include new stage; adjust run sequence ordering: enrichment -> relations -> file existence -> export.
5. Tests:
   - Unit: ensure file naming uniqueness helper works; ensure exported DataFrames have required columns.
   - Integration: run minimal pipeline on sample dataset; assert files created under temp directory; open a sheet and validate presence of data entry sheet with dropdown columns.

### Edge Cases
- Faculty with zero rows: skip and log.
- Programme course mapping referencing department with no rows: skip (current logic already covers this).
- Re-run same day: unique file naming adds suffix `_1`, `_2`, etc.

## 5. OSIRIS / People Enrichment (Revised – DB Centric)
### Key Update
The repository already contains fully normalized models for enrichment data (`Course`, `Person`, `CourseEmployee`, `Programme`, `Organization`, `MissingCourse`) plus ingestion helpers (`load_osiris_data`, `load_person_data`, `load_linked_persons_for_courses`). Therefore we DO NOT need JSON cache centric logic for steady‑state operation. JSON files can be treated as an optional import source (legacy) only; enrichment should query & update the relational tables directly.

### Current Gaps
1. The large legacy function `update_osiris_data` both scrapes and writes JSON; it is not integrated with the staging/pipeline and duplicates logic now represented in models.
2. Linking between `CopyrightItem.course_code` values and `Course` objects is partially handled by `link_courses_to_copyright_items` but not optimized (N+1 queries) and not part of the pipeline.
3. Person ↔ Course linkage relies on `load_linked_persons_for_courses` but is not surfaced as an explicit pipeline stage.
4. No standardized enrichment freshness policy. However all enrichment models inherit `TimestampMixin` (`created_at`, `modified_at`) which we can exploit as implicit freshness timestamps.

### Revised Enrichment Strategy
Pipeline enrichment stage will:
1. (Optional) Scrape missing course / person data ONLY for course codes (derived via `determine_course_code`) that lack a `Course` row OR whose `modified_at` is older than a TTL (configurable, default disabled).
2. Persist new/updated `Course` rows directly (no JSON). Reuse or refactor scraping logic into small functions: `fetch_course(code) -> dict`, `fetch_person(name) -> dict`.
3. Ensure course-person-role relationships by upserting `Person` then `CourseEmployee` rows (role classification logic preserved from legacy parser; roles: contact, docent, examiner, tutor, unknown_role).
4. Link newly (or previously) ingested `Course` objects to `CopyrightItem`s using numeric `cursuscode` extraction as part of the relations stage.

### Decomposition (DB Aware)
`gather_target_course_codes(settings) -> set[int]` (distinct from items or explicit override)
`select_missing_or_stale_courses(settings, codes, ttl_days|None) -> set[int]`
`async fetch_and_parse_courses(codes) -> list[CourseData]`
`derive_person_names(course_payloads) -> set[str]`
`select_missing_or_stale_persons(names, ttl_days|None) -> set[str]`
`async fetch_and_parse_persons(names) -> list[PersonData]`
`persist_courses(courses)` / `persist_persons(persons)` / `persist_course_employees(relations)` (bulk create with `on_conflict` if supported; otherwise prefetch existing -> diff -> create)

### Freshness (TTL) Policy
Use `modified_at` from the corresponding table; TTL comparison done in SQL (or Polars after retrieval) to reduce Python filtering overhead. Config keys (added to Settings):
`enrichment.course_ttl_days`, `enrichment.person_ttl_days` (None or 0 disables refresh based on age).

### Error Handling & Observability
- Count: requested, fetched, created, updated for each entity type.
- Structured log per stage (courses, persons, relations) with duration & throughput.

### Testing
- Unit: parser for course payload, parser for person page HTML (pure input → dict).
- Unit: selection logic for stale vs fresh given synthetic timestamps.
- Integration: mocked httpx scenario with mixed existing & missing records verifying incremental persistence.

### Removal of JSON Tasks
Tasks related to JSON cache diffing, atomic JSON writes, and JSON-based TTL have been removed from the TODO in favor of direct DB usage.

## 6. Relations Update Stage (Code-Informed)
**Existing Functions Found**: `link_courses_to_copyright_items()` and `update_duplicate_status()` in `db/update.py` (lines 737+) with current N+1 patterns.

### Current Implementation Issues
- `link_courses_to_copyright_items()`: Loops through all items, awaits `Course.get_or_none()` per course code per item
- `update_duplicate_status()`: Loops through all items, awaits `PDF.get_or_none()` per item, then saves each item individually
- Both functions called by `update_copyright_relations()` but not integrated into new pipeline

### Optimized Relations Module (`db/relations.py`)
Extract and optimize existing functions:
`async def update_duplicates(settings)` –
1. Batch prefetch all PDFs with `replace_with_id` (single query with filter)
2. Build replacement mapping dict in memory
3. Bulk update only items where duplicate status changes (compare before/after)

`async def link_courses(settings)` –
1. Query items missing course links using anti-join or LEFT JOIN WHERE course_id IS NULL
2. Extract all potential course codes using `determine_course_code()` in Polars (vectorized)
3. Batch fetch all relevant Course objects (single query with IN clause)
4. Build M2M relationships in memory, then bulk create CourseItem links
5. Log link counts per course for observability

`async def update_relations_async(settings)` – orchestrator calling both functions sequentially
Expose `async def update_relations_async(settings)` orchestrating both; call after enrichment (so new course codes may exist) and before export.

## 7. File Existence Verification
### Leverage Existing Schema
Model already exposes `file_exists` (bool) and `last_canvas_check` (datetime). So TTL requires no schema change; we compute age using `last_canvas_check` and re-check when NULL or older than `file_exists.ttl_days` (settings key to add).

### Proposed Flow
`refresh_file_existence_async(settings, force: bool=False)`:
1. Select candidate items via query filters (NULL `file_exists` OR `last_canvas_check < now()-ttl` OR force) limiting batch size.
2. Use existing concurrency pattern (aiometer/httpx); reuse extraction logic for file_id but move it to a pure helper so both DB and Excel paths share code.
3. Persist updates with one of:
   - Direct bulk update using Tortoise `.bulk_update` if supported for changed fields.
   - Fallback: per-item update only when status changed (compare old vs new in memory first).
4. Log hit ratio & mean requests/sec.

### Future Optimization (Optional)
Maintain rolling schedule (e.g. cap daily checks) – not in immediate phase.

### Testing
- Mock httpx client returning 200/404 pattern; verify join logic and counts.

## 8. Pipeline Ordering (Revised)
1. ingest_raw_data_async
2. ingest_faculty_updates_async
3. process_data_async
4. enrich_async (OSIRIS) [optional via flag]
5. update_relations_async (duplicates, course links)
6. refresh_file_existence_async (optional or scheduled)
7. export_reports_async

CLI flags to control inclusion (already partial for ingest/process/export): add `--no-enrich`, `--no-file-exists`, `--no-relations` if needed.

## 9. Incremental Delivery Phases
Phase A (High Priority / Minimal Viable):
- Implement export stage (faculties, programmes, all_items) using existing helpers.
- Integrate relations update stage (move logic, ensure idempotent).

Phase B:
- Implement DB-centric enrichment stage (course/person fetch & persistence) with optional TTL (skip if ttl_days unset).
- Add tests for stale selection & parser accuracy with mocked HTTP.

Phase C:
- Add file existence stage (TTL using existing `last_canvas_check`).
- Add tests (mock httpx) for selection & persistence.

Phase D:
- Performance tuning (bulk update strategies, reduce N+1 in linking functions).
- Optional: create consolidated retrieval view/function for export to minimize Python-side joins.

## 10. Risks & Mitigations
| Risk | Impact | Mitigation |
|------|--------|------------|
| Large OSIRIS requests rate-limited | Slow / failures | Bounded semaphore, backoff retry wrapper. |
| Excel writing race conditions (parallel runs) | Corrupted files | Use unique temp file & atomic rename; maintain per-run output dir optional (future). |
| Network dependency in tests | Flaky CI | Mock httpx layer; isolate network code behind small functions. |
| Inefficient sheet generation (repeated formatting) | Slow exports | Refactor sheet formatting into reusable style registry & vectorized width calc. |
| Memory spikes exporting very large datasets | OOM risk | Stream partitioned faculty exports; avoid holding all faculty DataFrames simultaneously (process sequentially). |

## 11. Excel Sheet Creation & Formatting Analysis / Refactor Targets
### Observations (from `sheets/sheet.py` & legacy `old_main.py`)
1. Column width & word-wrap logic executed per column; width heuristics okay but can be isolated into helper and unit tested.
2. DataEntrySheet currently writes cell-by-cell (acceptable for moderate size, but may be slow for very large sets). Potential optimization: build a 2D list and use `worksheet.append` in a loop or `write_only` mode (openpyxl) if performance becomes an issue.
3. Repeated style creation risk: `NamedStyle` reuse should check existence before add (openpyxl will raise on duplicate names). Current code implicitly reuses name but not guarded.
4. Dropdown validations added individually; fine, but we could consolidate by building all DV ranges then adding once.
5. No atomic write: writing directly risks partial file on crash.
6. Lack of structural validation before write (e.g., ensuring required columns present) – this is implicit; clearer pre-flight check desirable.
7. Partial code sections in current file (placeholders / blanks) suggest refactor interrupted – need restoration.

### Refactor Plan
Introduce utility functions:
`build_data_entry_matrix(data: pl.DataFrame, col_infos: list[ColInfo]) -> list[list[Any]]`
`apply_column_dimensions(ws, col_infos)`
`create_or_reuse_style(wb, style_name, **attrs)`
`add_dropdown_validations(ws, col_infos, max_row)`

Add `atomic_save_workbook(wb, target_path)` writing to `target_path.tmp` then renaming.

Add validation helper: `validate_export_dataframe(df, required_cols: set[str])` raising early with counts of missing columns.

### Testing
- Unit: width calculation logic & dropdown list injection.
- Integration: generate a small sheet; reopen with openpyxl; assert presence of data entry sheet, table style, dropdown DV count.

## 12. Updated TODO Items
See updated `.github/todo.md` for revised, DB‑centric tasks and sheet refactor items.

## 13. Acceptance Criteria Summary
- Export stage produces correct set of files for sample dataset & passes structural validation tests.
- Relations stage logs added/linked counts deterministically; rerun yields zero-delta metrics.
- Enrichment stage stores JSON caches, skips already fresh entries, and exposes enrichment counts.
- File existence stage processes only selected items & persists results correctly.

---
Prepared to guide implementation phases; update this document as phases complete.
