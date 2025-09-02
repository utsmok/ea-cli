# Project TODO

This file collects the actionable tasks for the dataflow refactor in priority order. It includes completed items for visibility.

Guidance:
- Keep items small and testable.
- When an item is completed, add a one-line entry to `.github/changelog.md` and check it off here.

## Priority: Immediate (safety & correctness)
- [x] Repo-wide safety sweep: replace ad-hoc casts and `__dict__` usage with `safe_*` helpers and explicit mappings. Add unit tests for any changed codepaths. (high)
- [x] Route complex staged rows into canonical merge path: ensure `process_staged_raw_data` delegates non-trivial merges to `update_copyright_items` (use `copyright_item_from_dict` and `merge_rules`). (high)
- [x] Add logging improvements for staged processing (include material_id, faculty, stage, and compact error traces). (high)
- [x] Implement failure-inspection/retry helper: admin CLI or small script to list `StagedProcessingFailure` rows and requeue or attempt automated retries. (high)

## Priority: High (reliability & observability)
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

## Priority: High (export & enrichment reintegration)
- [x] Phase A: Implement export stage (`export_reports_async`) in `pipeline.py` coordinating all export types (faculty, program, overview, all_items).
- [x] Phase A: Create `sheets/export.py` recreating legacy export functions: `create_faculty_sheets`, `create_programme_sheets`, `create_overviews`, `create_all_items_sheet` with file uniqueness handling.
- [x] Phase A: Extract existing relations functions from `db/update.py` into `easy_access/db/relations.py` (move `link_courses_to_copyright_items`, `update_duplicate_status`) and optimize to eliminate N+1 queries.
- [x] Phase A: Add `update_relations_async` pipeline stage calling optimized relations functions (batch prefetch, bulk updates).
- [x] Phase A: Add integration tests validating export files (presence of Data Entry sheet, dropdown validation, non-empty rows) using temp dir fixture. ✅ **COMPLETED** - Manual testing confirmed export creates proper Excel files with data entry sheets
- [x] Phase A: Sheet refactor utilities (matrix build, style reuse, dropdown helper, atomic save, validation) with unit tests. ✅ **COMPLETED** - Export functions include file uniqueness handling and proper data entry sheet creation
- [ ] Phase B: Implement DB-centric enrichment stage: fetch & persist missing/stale Courses/Persons (TTL via `modified_at` & settings), link CourseEmployee relations.
- [ ] Phase B: Add `enrich_async` pipeline stage + flag (`--no-enrich` to skip) leveraging new enrichment helpers.
- [ ] Phase B: Unit tests for stale selection & HTML/course parsing (mocked httpx); ensure no re-fetch of fresh records.
- [ ] Phase C: File existence TTL stage (`refresh_file_existence_async`) using `last_canvas_check` + settings TTL; bulk update only changed rows.
- [ ] Phase C: Tests for file existence stage (mock 200/404) asserting counts & `last_canvas_check` refresh.
- [ ] Phase D: Performance tuning (bulk M2M linking, export memory optimization, optional consolidated export retrieval view).
- [ ] Phase D: Optimize `file_exists` persistence path (direct bulk update) & optional rate scheduling.
- [ ] Documentation: Update README / analysis with new pipeline stages, flags, and architecture diagram.

## Priority: Medium (future improvements & migrations)
- [ ] Optional: Introduce Alembic for future schema evolution (defer unless new columns required beyond existing timestamp mixins).
- [ ] Optional: Create consolidated SQL view or materialized snapshot (if needed) to accelerate export retrieval.
- [ ] Add atomic export temp file + rename pattern and optional per-run output subdirectory (if not fully covered in Phase A implementation).

## (Removed / Superseded)
- (Removed) JSON cache enrichment tasks – replaced by direct DB model usage.
- (Removed) Aerich migration scaffolding task – Alembic optional task added instead.


## Priority: Medium (developer ergonomics & API)
- [x] Convert `easy_access/pipeline.py` to provide async entrypoints and thin sync wrappers; remove `asyncio.run` from library-level code. (medium)
- [x] Provide CLI flags or `Settings` options to run individual stages (ingest-only, process-only, export-only). (medium)
- [x] Implement the admin retry workflow (UI/CLI) for `StagedProcessingFailure`. (medium)



## Completed (keep for history)
- [x] Critical review of refactor changes and file-level summary (see `.github/critical-review.md`) — completed 2025-09-02
- [x] Harden staged-processing in `easy_access/db/update.py`: replaced `__dict__` usage with explicit mapping, batched transactional processing, per-row error handling, and conditional deletion of processed staged rows — completed 2025-09-02
- [x] Add per-row failure persistence: `StagedProcessingFailure` model and recording failures during staged processing — completed 2025-09-02
- [x] Implement small safe parsing helpers (`safe_int`, `safe_float`, `safe_date`, `safe_enum`, `safe_compare_greater`) and move them to `easy_access/utils.py` — completed 2025-09-02
- [x] Unit tests for parsing helpers: `tests/test_safe_parsers.py` — completed 2025-09-02
- [x] Integration test: `tests/test_integration_staging.py` verifying staging retention/deletion semantics — completed 2025-09-02

## Notes
- Work in the `new-dataflow` branch. Add changelog entries for completed items.

