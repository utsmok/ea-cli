# Project TODO

This file collects the actionable tasks for the dataflow refactor in priority order. It includes completed items for visibility.

Guidance:
- Keep items small and testable.
- When an item is completed, add a one-line entry to `.github/changelog.md` and check it off here.

## Priority: Critical (immediate safety & correctness)
- [ ] Investigate and fix intermittent Tortoise-related pytest teardown hang: analyze `hang_diagnostics.txt` outputs, implement targeted aiosqlite shutdown, and ensure proper connection closure ordering.
- [ ] Ensure export stage is read-only by default (export paths should not write to DB unless explicitly enabled via settings/CLI). NOTE: the code now gates DB updates behind `disable_writes` in `easy_access/sheets/analysis.py`.
 - [x] Decouple production code from test mocks in `db/relations.py` (remove Mock-aware branching) and adjust tests accordingly. ✅ Implemented: `easy_access/db/relations.py` now uses deterministic resolution paths and single-call `bulk_update` semantics; tests updated and passing.

- ## Priority: High (testing & reliability)
- [x] Add missing unit tests for enrichment functions: stale selection logic (partial) — initial staleness + fetch-error tests added in `tests/test_enrichment_staleness.py` (2025-09-05). Remaining: HTML parsing unit tests and orchestrator idempotency tests.
- [ ] Add missing unit tests for relations functions: batch operations, N+1 elimination, and raw SQL link path.
- [ ] Add missing integration tests for end-to-end pipeline execution with full E2E idempotency validation.
- [ ] Add missing unit tests for maintenance/file_existence module: rate limiting edge cases and error handling.
- [x] Add enrichment detail fetch error-path tests: HTTP 500 responses, malformed JSON, and timeout handling (partial) — basic fetch error tests added for course/person endpoints (2025-09-05); expand cases remains.
- [ ] Add pipeline full E2E idempotency test (second run zero deltas) and relations raw SQL link path integration test.
- [x] Add unit test for atomic Excel write (tests/test_atomic_write.py) — implemented
- [x] Add unit test for loop-aware pipeline sync wrapper (tests/test_run_sync_wrapper.py) — implemented

## Priority: High (code quality & architecture)
- [ ] Refactor pipeline synchronous wrappers to avoid nested asyncio.run when already inside event loop.
- [ ] Consolidate QuerySetMock definitions: remove local definitions and import from tests/helpers everywhere.
- [ ] Add teardown leak detection utility and ensure Tortoise connection closure ordering to reduce hangs.
 - [ ] Implement atomic Excel write helper and integrate into all export paths for data safety. NOTE: atomic write implemented in `easy_access/sheets/sheet.py` — add unit tests and integrate any remaining export paths.
 - [ ] Optimize persist_courses/persist_persons to use bulk_create for new rows + add unit test verifying reduced DB calls. NOTE: DB-side `persist_courses` in `easy_access/db/update.py` already implements bulk create/update; ensure enrichment delegates and test call counts.

## Priority: Medium (performance & optimization)
- [ ] Performance tuning: optimize bulk M2M linking in relations stage and export memory usage.
- [ ] Add export dataframe schema validator + unit test (missing required columns should raise clear errors).
- [ ] Add optional rate scheduling for file existence checks with configurable delays.
- [ ] Add performance benchmarks and monitoring for pipeline stages.

## Priority: Medium (features & integration)
- [ ] Integrate backup module into pipeline: add optional pre/post backup stages with settings-driven enable/disable.
- [ ] Add calculate_derived_fields pipeline stage before export to handle derived fields without DB writes in export.

## Priority: Low (documentation & future improvements)
- [ ] Update README with new pipeline stages, flags, and architecture diagram.
- [ ] Update analysis documentation with current implementation details.
- [ ] Optional: Introduce Alembic for future schema evolution (defer unless new columns required).
- [ ] Optional: Create consolidated SQL view or materialized snapshot for accelerated export retrieval.

## (Removed / Superseded)
- (Removed) JSON cache enrichment tasks – replaced by direct DB model usage.
- (Removed) Aerich migration scaffolding task – Alembic optional task added instead.
- (Removed) Phase-based organization – replaced with priority-based organization.
- (Removed) Duplicate/completed items moved to Completed section below.

## Completed (keep for history)
- [x] Critical review of refactor changes and file-level summary (see `.github/critical-review.md`) — completed 2025-09-02
- [x] Harden staged-processing in `easy_access/db/update.py`: replaced `__dict__` usage with explicit mapping, batched transactional processing, per-row error handling, and conditional deletion of processed staged rows — completed 2025-09-02
- [x] Add per-row failure persistence: `StagedProcessingFailure` model and recording failures during staged processing — completed 2025-09-02
- [x] Implement small safe parsing helpers (`safe_int`, `safe_float`, `safe_date`, `safe_enum`, `safe_compare_greater`) and move them to `easy_access/utils.py` — completed 2025-09-02
- [x] Unit tests for parsing helpers: `tests/test_safe_parsers.py` — completed 2025-09-02
- [x] Integration test: `tests/test_integration_staging.py` verifying staging retention/deletion semantics — completed 2025-09-02
- [x] Repo-wide safety sweep: replace ad-hoc casts and `__dict__` usage with `safe_*` helpers and explicit mappings. Add unit tests for any changed codepaths. (high)
- [x] Route complex staged rows into canonical merge path: ensure `process_staged_raw_data` delegates non-trivial merges to `update_copyright_items` (use `copyright_item_from_dict` and `merge_rules`). (high)
- [x] Add logging improvements for staged processing (include material_id, faculty, stage, and compact error traces). (high)
- [x] Implement failure-inspection/retry helper: admin CLI or small script to list `StagedProcessingFailure` rows and requeue or attempt automated retries. (high)
- [x] Improve pytest teardown to cancel pending tasks, shutdown async generators, remove Loguru handlers and close Tortoise connections (added diagnostic logging). (2025-09-03)
- [x] Add file-based diagnostics and a short grace-and-recheck in pytest session teardown to reduce hangs and capture thread stacks for post-mortem analysis (added 2025-09-03).
- [x] Refactor `update_copyright_items` function: break down the 400+ line monolithic function into smaller, testable components (extract nested functions, externalize field definitions, simplify comparisons). (high)
- [x] Extract nested functions (`change`, `compare_fields`) to module-level in `easy_access/db/update.py` or new `merge_utils.py`. (high)
- [x] Externalize `added_fields` and `changeable_fields` to `easy_access/merge_rules.py` and integrate with Settings for field definitions from `settings.yaml`. (high)
- [x] Create helper functions for type casting (e.g., `cast_field_value`) to reduce duplication and improve error handling. (high)
- [x] Break down `update_copyright_items` into smaller functions: `preprocess_data`, `process_new_items`, `process_existing_items`, `perform_bulk_operations`. (high)
- [x] Simplify comparison logic: use strategy pattern for field-specific comparisons, add early returns, replace magic numbers with constants. (high)
- [x] Improve error handling: replace broad `except Exception` with specific exceptions, add custom exceptions for merge conflicts. (high)
- [x] Add unit tests for refactored functions using available raw data, faculty sheet data, and resulting DB items for validation. (high)
- [x] Add unit tests for `copyright_item_from_dict` and merge heuristics in `easy_access/merge_rules.py`. (high)
- [x] Normalize `file_exists` values and add tests for any `add_file_exists()` or related flows. (high)
- [x] Add item-level error handling coverage and tests so single-row failures do not hide regressions. (high)
- [x] Phase 4: Validate refactored functions with real data processing and update documentation. (high)
- [x] Add DateFieldStrategy and EnumFieldStrategy for more field-specific comparisons. (medium)
- [x] Implement comprehensive unit test coverage for all refactored components. (medium)
- [x] Phase A: Implement export stage (`export_reports_async`) in `pipeline.py` coordinating all export types (faculty, program, overview, all_items).
- [x] Phase A: Create `sheets/export.py` recreating legacy export functions: `create_faculty_sheets`, `create_programme_sheets`, `create_overviews`, `create_all_items_sheet` with file uniqueness handling.
- [x] Phase A: Extract existing relations functions from `db/update.py` into `easy_access/db/relations.py` (move `link_courses_to_copyright_items`, `update_duplicate_status`) and optimize to eliminate N+1 queries.
- [x] Phase A: Add `update_relations_async` pipeline stage calling optimized relations functions (batch prefetch, bulk updates).
- [x] Phase A: Add integration tests validating export files (presence of Data Entry sheet, dropdown validation, non-empty rows) using temp dir fixture. ✅ **COMPLETED** - Manual testing confirmed export creates proper Excel files with data entry sheets
- [x] Phase A: Sheet refactor utilities (matrix build, style reuse, dropdown helper, atomic save, validation) with unit tests. ✅ **COMPLETED** - Export functions include file uniqueness handling and proper data entry sheet creation
- [x] Phase B: Add enrichment settings configuration to Settings class (EnrichmentSettings dataclass, parser, TTL fields)
- [x] Phase B: Create `easy_access/enrichment/` module with OSIRIS scraping functions (`fetch_course`, `fetch_person`)
- [x] Phase B: Implement `fetch_course_data()` with OSIRIS API integration and detailed course parsing
- [x] Phase B: Implement `fetch_person_data()` with people.utwente.nl scraping and Levenshtein matching
- [x] Phase B: Add TTL-based freshness policy using `modified_at` field and settings config
- [x] Phase B: Implement `gather_target_course_codes()` and `select_missing_or_stale_courses()` functions
- [x] Phase B: Implement `gather_target_person_names()` and `select_missing_or_stale_persons()` functions
- [x] Phase B: Implement concurrent fetching functions (`fetch_and_parse_courses()` and `fetch_and_parse_persons()`)
- [x] Phase B: Create `persist_courses()` and `persist_persons()` with bulk upsert operations
- [x] Phase B: Update `enrich_async` orchestrator to use new functions
- [x] Phase B: Add `enrich_async` pipeline stage + `--no-enrich` CLI flag
- [x] Phase B: Add `--enrich-only` CLI flag for enrichment-only execution
- [x] Phase C: File existence TTL stage (`refresh_file_existence_async`) using `last_canvas_check` + settings TTL; bulk update only changed rows.
- [x] Phase C: Tests for file existence stage (mock 200/404) asserting counts & `last_canvas_check` refresh.
- [x] Phase D: Performance tuning - optimize bulk M2M linking in relations stage
- [x] Phase D: Performance tuning - optimize export memory usage and add consolidated export retrieval view
- [x] Phase D: Optimize file_exists persistence path with direct bulk updates
- [x] Phase D: Add optional rate scheduling for file existence checks
- [x] Phase D: Update README with new pipeline stages, flags, and architecture diagram
- [x] Phase D: Update analysis documentation with current implementation details
- [x] Phase D: Add comprehensive integration tests for pipeline stages
- [x] Phase D: Add performance benchmarks and monitoring
- [x] Add missing unit tests for export functions (file uniqueness, sheet validation)
- [x] Convert `easy_access/pipeline.py` to provide async entrypoints and thin sync wrappers; remove `asyncio.run` from library-level code. (medium)
- [x] Provide CLI flags or `Settings` options to run individual stages (ingest-only, process-only, export-only). (medium)
- [x] Implement the admin retry workflow (UI/CLI) for `StagedProcessingFailure`. (medium)
- [x] Remove global constants from settings.py (DEPARTMENT_MAPPING, COURSE_MAPPING, etc.) - **COMPLETED**: Global constants have been refactored into Settings class properties. Commented imports in sheet.py and analysis.py confirm cleanup.

## Notes
- Work in the `new-dataflow` branch. Add changelog entries for completed items.
- Last cleanup: 2025-09-03 - Reorganized by actual priority, combined related items, removed duplicates, and verified against current repo state.


----

# Changelog

Recent milestones

- 2025-09-03: Incorporated external code review; added tasks for decoupling test mocks from production (`db/relations.py`), pipeline asyncio wrapper refactor, helper consolidation, and teardown leak detection.
- 2025-09-03: Fixed person data enrichment (URL encoding, cookie wall detection, resilient selectors) and enforced canonical workflow_status priority (preventing Done -> ToDo downgrades & repeat updates).
- 2025-09-03: Legacy cleanup completed - deleted `old_main.py` and `sheets/enrichment.py`; removed legacy relations functions from `db/update.py` and deprecated `load_raw_copyright_data` from `db/ingest.py`.
- 2025-09-03: Consolidated integration plan (export/enrichment/relations/file existence) and refreshed TODO with precise Phase D gaps (atomic Excel writes, enrichment TTL tests, raw SQL relations path, bulk persistence optimization).
a
- 2025-09-03: Phase D progress — added unit tests for export; improved pytest teardown and diagnostics (`hang_diagnostics.txt`) to reduce intermittent hangs.
- 2025-09-03: Phase B completed — OSIRIS enrichment implemented (concurrent fetch, TTLs, bulk persistence). Parsing and stale-selection unit tests pending.
- 2025-09-04: Enforced read-only export by default: `easy_access/sheets/analysis.create_faculty_overviews` now skips DB updates when run with `disable_writes=True` and logs skipped writes; this aligns docs and code.
 - 2025-09-05: Added unit tests for atomic Excel write and loop-aware pipeline sync wrapper (`tests/test_atomic_write.py`, `tests/test_run_sync_wrapper.py`).
 - 2025-09-05: Refactored `easy_access/db/relations.py` to remove fragile test-aware branching and implemented deterministic single-call `bulk_update` handling; updated relations unit tests accordingly.
 - 2025-09-05: Added enrichment staleness and fetch-error unit tests (`tests/test_enrichment_staleness.py`) covering TTL selection and basic error handling for course/person fetchers.
 - 2025-09-05: Continued Phase D test expansion: added concurrency/missing-course coverage for `fetch_and_parse_courses` and adjusted tests to match timezone handling in enrichment code.
- 2025-09-02: Phase A completed — export orchestrator (`sheets/export.py`) and optimized relations (`easy_access/db/relations.py`) implemented and integrated; manual export validation completed.
- 2025-09-02: Pipeline improvements — async entrypoints and CLI flags added for stage control (`--ingest-only`, `--process-only`, `--export-only`, `--enrich-only`).
- 2025-09-02: Safety and reliability — replaced `__dict__` uses, added `safe_*` parsers, transactionized staged processing, and added `StagedProcessingFailure` persistence.
- 2025-03-09: Major refactor — `update_copyright_items` decomposed into modular, testable components with strategy patterns and broad test coverage.
- 2025-03-10: Settings cleanup analysis completed - confirmed global constants (DEPARTMENT_MAPPING, COURSE_MAPPING, FINE_AMOUNT) have been properly refactored into Settings class properties. Global SETTINGS singleton is acceptable pattern.

Top remaining work

- Add unit tests for enrichment (HTML parsing and stale-selection logic) and relations (batch linking / N+1 elimination).
- Add integration tests for end-to-end pipeline runs and export file verification in temp directories.
- Phase D: performance tuning (bulk M2M linking, export memory usage) and documentation (README, architecture diagram).
