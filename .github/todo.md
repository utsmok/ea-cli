# Project TODO

This file lists the remaining high-impact tasks required to finish the refactor. Items are small, testable, and ordered by impact.

Guidance:
- Keep items small and testable.
- When an item is completed, add a one-line entry to `.github/changelog.md` and check it off here.

Branch / context:
- Working branch: `new-dataflow` (keep changelog entries for any completed items).

## Priority: Critical (immediate safety & correctness)
- [ ] Ensure export stage is read-only by default (export paths should not write to DB unless explicitly enabled via settings/CLI). Confirm `disable_writes` gating across export flows.


## (testing, reliability & correctness)

- [ ] Add relations raw-SQL integration test: exercise `link_courses` raw-SQL path by running without mocking `CopyrightItem.bulk_update` and assert expected links are created.
- [ ] Decouple test-aware logic from production code in these modules: `easy_access/db/relations.py`, `easy_access/maintenance/file_existence.py`, and `easy_access/enrichment/osiris.py` (where safe). Implement a small test adapter used only by tests rather than scattered mock-detection heuristics.
- [ ] Consolidate `QuerySetMock` and related test helpers into `tests/helpers.py` and update tests to import the shared helper.

##  (code quality & performance)
- [ ] Refactor `easy_access/pipeline._run_sync` to a documented, loop-safe dispatcher (avoid relying on nested `asyncio.run` in threads); add regression tests for loop-aware behavior.

##  (export safety & validation)
- [ ] Add export dataframe schema validator + unit test: ensure required columns (e.g., `material_id`) exist and fail early with clear errors.
- [ ] Verify atomic Excel writes are used across all export paths and add failure-simulation tests (simulate write error leaving tmp file and ensure final file not corrupted).

 - [ ] Generalize conditional formatting helper in `sheet.py`: refactor `_add_conditional_formatting` so it can be applied to any column and support all current condition types. Drive conditions and targets from `settings.yaml` and `ColInfo` in `settings.py`. Define a small set of premade styles (also configurable from `settings.yaml`) for reuse across sheets.

 - [ ] Ingest and persist newly-added `CopyrightItem` fields: `filehash`, `last_scan_date_university`, and `last_scan_date_course`.
	 - Ensure these fields are read from raw inputs, validated, and included in processing and exports.
	 - Use `filehash` for deduplication: items with identical `filehash` should be marked with `is_duplicate == true` (investigate downstream handling and sheet presentation for duplicates).
	 - Treat `last_scan_date_university` and `last_scan_date_course` as date-only values; validate Excel-derived dates carefully to ensure correct day/month/year parsing.

 - [ ] Ensure exported Excel files open showing the `Data Entry` sheet by default (both the `Complete Data` and `Data Entry` sheets must still be present).

## (features & integration)
- [ ] Integrate backup module into pipeline: add optional pre/post backup stages with settings-driven enable/disable.
- [ ] Add calculate_derived_fields pipeline stage before export to compute non-persistent derived fields used only for reporting.

 - [ ] Long-term: implement reactive faculty sheets workflow and monitor script.
	 - Replace static `overview`/`weekly` sheets with three sheet states per export file: `new`, `to_check`, and `checked` (keep `checked` initially empty).
	 - New items land in `new`. Items manually marked `Done` (workflowstatus) move to `checked`. Any other user edits move items to `to_check`.
	 - Items in `checked` are considered locked and persisted to the DB. `new` and `to_check` remain live and should reflect upstream changes (raw copyright imports, file existence updates) but not overwrite user-entered data until reviewed.
	 - Provide a periodic script that scans `faculty_sheets/`, reconciles state transitions, and updates sheets a few times per day.

## (maintenance & reliability)
## NOTE: Only execute these if hanging-issues keep cropping up.
- [ ] Add teardown leak-detection utility and harden Tortoise / aiosqlite shutdown ordering; run tests to verify no intermittent hangs.
- [ ] If pytest keeps hanging: investigate and fix intermittent Tortoise-related pytest teardown hang (see `hang_diagnostics.txt`) and implement targeted aiosqlite shutdown and connection closure ordering.

##  (docs & housekeeping)
- [ ] Update README with new pipeline stages, flags, and a short architecture diagram (PNG/SVG). Include quick-run commands for local testing.

## (Removed / Superseded)
- (Removed) JSON cache enrichment tasks – replaced by direct DB model usage.
- (Removed) Aerich migration scaffolding task – Alembic optional task added instead.
- (Removed) Phase-based organization – replaced with priority-based organization.

## Completed (keep for history)
The list below is a merged, de-duplicated history combining both versions. When items are marked completed, add an entry to `.github/changelog.md` (one-line) and leave the checkbox.

- [x] Critical review of refactor changes and file-level summary (see `.github/critical-review.md`) — completed 2025-09-02
- [x] Harden staged-processing in `easy_access/db/update.py`: replaced `__dict__` usage with explicit mapping, batched transactional processing, per-row error handling, and conditional deletion of processed staged rows — completed 2025-09-02
- [x] Add per-row failure persistence: `StagedProcessingFailure` model and recording failures during staged processing — completed 2025-09-02
- [x] Implement small safe parsing helpers (`safe_int`, `safe_float`, `safe_date`, `safe_enum`, `safe_compare_greater`) and move them to `easy_access/utils.py` — completed 2025-09-02
- [x] Unit tests for parsing helpers: `tests/test_safe_parsers.py` — completed 2025-09-02
- [x] Integration test: `tests/test_integration_staging.py` verifying staging retention/deletion semantics — completed 2025-09-02
- [x] Repo-wide safety sweep: replace ad-hoc casts and `__dict__` usage with `safe_*` helpers and explicit mappings. Add unit tests for any changed codepaths.
- [x] Route complex staged rows into canonical merge path: ensure `process_staged_raw_data` delegates non-trivial merges to `update_copyright_items` (use `copyright_item_from_dict` and `merge_rules`).
- [x] Add logging improvements for staged processing (include material_id, faculty, stage, and compact error traces).
- [x] Implement failure-inspection/retry helper: admin CLI or small script to list `StagedProcessingFailure` rows and requeue or attempt automated retries.
- [x] Improve pytest teardown to cancel pending tasks, shutdown async generators, remove Loguru handlers and close Tortoise connections (added diagnostic logging). — 2025-09-03
- [x] Add file-based diagnostics and a short grace-and-recheck in pytest session teardown to reduce hangs and capture thread stacks for post-mortem analysis (added 2025-09-03).
- [x] Refactor `update_copyright_items` function: break down the 400+ line monolithic function into smaller, testable components (extract nested functions, externalize field definitions, simplify comparisons).
- [x] Extract nested functions (`change`, `compare_fields`) to module-level in `easy_access/db/update.py` or new `merge_utils.py`.
- [x] Externalize `added_fields` and `changeable_fields` to `easy_access/merge_rules.py` and integrate with Settings for field definitions from `settings.yaml`.
- [x] Create helper functions for type casting (e.g., `cast_field_value`) to reduce duplication and improve error handling.
- [x] Break down `update_copyright_items` into smaller functions: `preprocess_data`, `process_new_items`, `process_existing_items`, `perform_bulk_operations`.
- [x] Simplify comparison logic: use strategy pattern for field-specific comparisons, add early returns, replace magic numbers with constants.
- [x] Improve error handling: replace broad `except Exception` with specific exceptions, add custom exceptions for merge conflicts.
- [x] Add unit tests for refactored functions using available raw data, faculty sheet data, and resulting DB items for validation.
- [x] Add unit tests for `copyright_item_from_dict` and merge heuristics in `easy_access/merge_rules.py`.
- [x] Normalize `file_exists` values and add tests for any `add_file_exists()` or related flows.
- [x] Add item-level error handling coverage and tests so single-row failures do not hide regressions.
- [x] Phase 4: Validate refactored functions with real data processing and update documentation.
- [x] Add DateFieldStrategy and EnumFieldStrategy for more field-specific comparisons.
- [x] Implement comprehensive unit test coverage for all refactored components.
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
- [x] Convert `easy_access/pipeline.py` to provide async entrypoints and thin sync wrappers; remove `asyncio.run` from library-level code.
- [x] Provide CLI flags or `Settings` options to run individual stages (ingest-only, process-only, export-only).
- [x] Implement the admin retry workflow (UI/CLI) for `StagedProcessingFailure`.
- [x] Remove global constants from settings.py (DEPARTMENT_MAPPING, COURSE_MAPPING, etc.) - **COMPLETED**: Global constants have been refactored into Settings class properties. Commented imports in sheet.py and analysis.py confirm cleanup.

## Notes / Assumptions
- We keep test-aware fallbacks for minimal test payloads where removing them would be high risk; the longer-term plan is to replace ad-hoc detection with explicit test adapters.
- HTML parsing unit tests are intentionally omitted per request; we will validate parser outputs instead and debug changes if/when parsing fails.

## Notes
- Work in the `new-dataflow` branch. Add changelog entries for completed items.
- Last cleanup: 2025-09-03 - Reorganized by actual priority, combined related items, removed duplicates, and verified against current repo state.
