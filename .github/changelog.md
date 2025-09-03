# Changelog

- 2025-09-03: **PHASE D PROGRESS** - Completed comprehensive unit tests for export functions: created `tests/test_export.py` with 17 test cases covering faculty data gathering, faculty/programme/all-items sheet export, faculty overviews, main export orchestrator, file uniqueness handling, and integration testing. All tests pass with proper mocking of database operations, file system interactions, and Excel generation functions.

- 2025-09-03: **PHASE B COMPLETED** - Successfully implemented complete OSIRIS enrichment system: course/person fetching with concurrent HTTP requests, TTL-based freshness policies, bulk database persistence, pipeline integration with CLI flags, and robust error handling. All core functionality working, unit tests marked as future enhancement.

- 2025-09-02: **PHASE B COMPLETED** - Successfully integrated enrichment pipeline stage: added `enrich_data_async()` to DataPipeline, integrated into main processing workflow, added `--enrich-only` CLI flag for selective execution, and updated EasyAccessTool to conditionally run enrichment based on `refresh_osiris_data` setting.

- 2025-09-02: **PHASE B PROGRESS** - Completed concurrent fetching and persistence implementation: added `fetch_and_parse_courses()` and `fetch_and_parse_persons()` with asyncio.gather for concurrent HTTP requests, implemented `persist_courses()` and `persist_persons()` with bulk upsert operations, and updated `enrich_async` orchestrator to coordinate the complete enrichment pipeline.

- 2025-09-02: **PHASE B PROGRESS** - Completed person fetching implementation: added `fetch_person_data()` function with people.utwente.nl scraping, Levenshtein distance matching for best person selection, detailed HTML parsing for name/email/organization/education data, and proper type checking for BeautifulSoup elements to resolve all lint errors.

- 2025-09-02: **FIXED** RuntimeError "This event loop is already running" in export pipeline: converted `create_faculty_overviews` to async function, replaced `asyncio.get_event_loop().run_until_complete()` with direct await, added sync wrapper for backward compatibility with legacy code.

# Changelog

- 2025-09-02: **PHASE A COMPLETED** - Successfully implemented export & relations reintegration: created `sheets/export.py` with complete export orchestrator (faculty sheets, programme sheets, all items, overviews), optimized `db/relations.py` with batch operations eliminating N+1 queries, integrated both into pipeline with proper async handling. Fixed critical RuntimeError "This event loop is already running" by converting `create_faculty_overviews` to async. Manual testing confirmed Excel file generation with data entry sheets and proper file uniqueness handling.

- 2025-09-02: **PHASE B PROGRESS** - Implemented OSIRIS course fetching logic: extracted `fetch_course_data()` function with API integration, detailed course parsing including contacts/docents/examinators, and helper functions for data processing. Added `_fetch_course_details()` for retrieving detailed course information and `_process_teacher_items()` for consistent teacher data handling.

- 2025-09-02: Added CLI flags for individual pipeline stages: `--ingest-only`, `--process-only`, `--export-only` to `run.py process` command, allowing developers to run specific stages of the data processing pipeline.
- 2025-09-02: Implemented comprehensive admin CLI for `StagedProcessingFailure` management with commands: `inspect-failures`, `failure-stats`, `retry-failures`, `cleanup-failures`; includes Trogon TUI support for enhanced user experience.
- 2025-09-02: Converted `easy_access/pipeline.py` to provide async entrypoints (`run_async`, `ingest_raw_data_async`, etc.) and thin sync wrappers; removed `asyncio.run` from library-level code for better async compatibility.
- 2025-03-09: Added detailed refactoring plan for `update_copyright_items` function to `.github/update_copyright_items_refactor_plan.md` and integrated subtasks into `.github/todo.md`, including integration with Settings for dynamic field definitions and testing with available data.
- 2025-03-06: Initial analysis and refactor plan added to `.github/code_analysis.md`.
- 2025-03-07: Extracted `todo.md` into `.github/todo.md` (canonical checklist).
- 2025-03-08: Added `.github/analysis.md` containing background, design goals, and Tortoise ORM recommendations.
- 2025-09-02: Added "Critical review of refactor changes" todo and summarized diffs against commit `849628b965f4bd23b91400f1a5034eaf40787334` (files added/modified and follow-up recommendations).
- 2025-09-02: Completed Phase 2 of `update_copyright_items` refactor: extracted functions, externalized field definitions, implemented strategy pattern, improved error handling with custom exceptions.
- 2025-09-02: Added comprehensive integration plan for export, OSIRIS enrichment, relations update, and file existence stages (`export_enrichment_integration_plan.md`) and populated new TODO items for phased implementation.
- 2025-09-02: Revised integration plan to reflect existing normalized DB models for courses/persons (removed JSON cache approach), added sheet export refactor analysis & updated TODO/memory accordingly.
- 2025-09-02: Final code review confirms export functions completely missing (must recreate from legacy patterns), relations functions exist but need N+1 optimization, pipeline export stage stubbed. Updated plan/todo/memory to reflect critical path.
 - 2025-09-02: Hardened staged-processing in `easy_access/db/update.py`: removed `staged_item.__dict__` usage, added batched transactional processing, per-row error handling, and conditional deletion of processed staged rows; updated `.github/todo.md` with follow-ups.
- 2025-03-09: Repo-wide safety sweep: replaced `__dict__` usage with `vars()`, ad-hoc casts with `safe_*` helpers in `easy_access/db/update.py`, `easy_access/settings.py`, `easy_access/db/ingest.py`, and `dashboard/data.py`; updated imports.
- 2025-03-09: Routed complex staged rows into canonical merge path: modified `process_staged_raw_data` to detect non-trivial fields and delegate to `update_copyright_items` using extracted `merge_rules.py`.
 - 2025-09-02: Added parsing helpers (`safe_int`, `safe_float`, `safe_date`, `safe_enum`, `safe_compare_greater`) and unit tests `tests/test_safe_parsers.py`.
 - 2025-09-02: Added integration test `tests/test_integration_staging.py` that verifies staged-processing semantics (failed-first-run preserves staged rows; success run deletes processed rows).
 - 2025-09-02: Added `StagedProcessingFailure` model to `easy_access/db/models.py` and persisted per-row processing failures from `process_staged_raw_data` into this table for later inspection/retry.
 - 2025-09-02: Moved small safe parsing helpers (`safe_int`, `safe_float`, `safe_date`, `safe_enum`, `safe_compare_greater`) into `easy_access/utils.py` for reuse across the codebase; updated `easy_access/db/update.py` to use these helpers.
 - 2025-09-02: Completed repo-wide safety sweep: replaced `__dict__` usage with `vars()` in dashboard/components.py and easy_access/settings.py; replaced ad-hoc casts (int, float) with `safe_*` helpers in easy_access/settings.py, easy_access/db/ingest.py, dashboard/data.py; updated imports accordingly. No new unit tests added as existing safe_* tests cover the changes.
- 2025-03-09: Completed Phase 2 refactor: extracted nested functions (`change`, `compare_fields`) to module-level in `easy_access/db/update.py` or new `merge_utils.py`.
- 2025-03-09: Externalized `added_fields` and `changeable_fields` to `easy_access/merge_rules.py` and integrated with Settings for field definitions from `settings.yaml`.
- 2025-03-09: Created helper functions for type casting (e.g., `cast_field_value`) to reduce duplication and improve error handling.
- 2025-03-09: Broke down `update_copyright_items` into smaller functions: `preprocess_data`, `process_new_items`, `process_existing_items`, `perform_bulk_operations`.
- 2025-03-09: Simplified comparison logic: used strategy pattern for field-specific comparisons, added early returns, replaced magic numbers with constants.
- 2025-03-09: Improved error handling: replaced broad `except Exception` with specific exceptions, added custom exceptions for merge conflicts.
- 2025-03-09: Added unit tests for refactored functions using available raw data, faculty sheet data, and resulting DB items for validation.
- 2025-03-09: Added unit tests for `copyright_item_from_dict` and merge heuristics in `easy_access/merge_rules.py`.
- 2025-03-09: Normalized `file_exists` values and added tests for any `add_file_exists()` or related flows.
- 2025-03-09: Added item-level error handling coverage and tests so single-row failures do not hide regressions.
- 2025-03-09: Phase 4: Validated refactored functions with real data processing and updated documentation.
- 2025-03-09: Added DateFieldStrategy and EnumFieldStrategy for enhanced field comparisons.
- 2025-03-09: Implemented comprehensive unit test coverage for all refactored components.
- 2025-03-09: **MAJOR MILESTONE**: Complete `update_copyright_items` refactor finished: All 5 phases implemented including comprehensive integration tests with real data (7 tests in `tests/test_integration_real_data.py`), 50 unit tests covering all refactored components, strategy pattern for field comparisons, custom exceptions, Settings integration, and full validation. The 400+ line monolithic function has been successfully broken down into modular, testable components.
- 2025-09-02: Implemented comprehensive admin CLI for `StagedProcessingFailure` management with commands: `inspect-failures`, `failure-stats`, `retry-failures`, `cleanup-failures`; includes Trogon TUI support for enhanced user experience.
- 2025-09-02: Added CLI flags for individual pipeline stages: `--ingest-only`, `--process-only`, `--export-only` to `run.py process` command, allowing developers to run specific stages of the data processing pipeline.
- 2025-09-02: Converted `easy_access/pipeline.py` to provide async entrypoints (`run_async`, `ingest_raw_data_async`, etc.) and thin sync wrappers; removed `asyncio.run` from library-level code for better async compatibility.
- 2025-09-02: **PHASE A COMPLETED**: Implemented export stage with `sheets/export.py` recreating legacy export functions (`create_faculty_sheets`, `create_programme_sheets`, `create_overviews`, `create_all_items_sheet`) with file uniqueness handling and DB-first architecture.
- 2025-09-02: **PHASE A COMPLETED**: Extracted relations functions from `db/update.py` into optimized `easy_access/db/relations.py` with batch operations to eliminate N+1 queries.
- 2025-09-02: **PHASE A COMPLETED**: Added `update_relations_async` pipeline stage calling optimized relations functions (batch prefetch, bulk updates).
- 2025-09-02: **PHASE A COMPLETED**: Added `export_reports_async` to pipeline coordinating all export types (faculty, program, overview, all_items).hangelog

Th- 2025-03-09: **COMPLETED** comprehensive refactoring of `update_copyright_items` function: fully implemented all 5 phases including integration tests with real data validation, comprehensive unit test coverage (50 tests), and Settings integration. All refactoring goals achieved.s file records high-level repository changes and analysis edits.

- 2025-09-02: Added CLI flags for individual pipeline stages: `--ingest-only`, `--process-only`, `--export-only` to `run.py process` command, allowing developers to run specific stages of the data processing pipeline.
- 2025-09-02: Implemented comprehensive admin CLI for `StagedProcessingFailure` management with commands: `inspect-failures`, `failure-stats`, `retry-failures`, `cleanup-failures`; includes Trogon TUI support for enhanced user experience.
- 2025-09-02: Converted `easy_access/pipeline.py` to provide async entrypoints (`run_async`, `ingest_raw_data_async`, etc.) and thin sync wrappers; removed `asyncio.run` from library-level code for better async compatibility.
- 2025-03-09: Added detailed refactoring plan for `update_copyright_items` function to `.github/update_copyright_items_refactor_plan.md` and integrated subtasks into `.github/todo.md`, including integration with Settings for dynamic field definitions and testing with available data.
- 2025-03-06: Initial analysis and refactor plan added to `.github/code_analysis.md`.
- 2025-03-07: Extracted `todo.md` into `.github/todo.md` (canonical checklist).
- 2025-03-08: Added `.github/analysis.md` containing background, design goals, and Tortoise ORM recommendations.
- 2025-09-02: Added "Critical review of refactor changes" todo and summarized diffs against commit `849628b965f4bd23b91400f1a5034eaf40787334` (files added/modified and follow-up recommendations).
- 2025-09-02: Completed Phase 2 of `update_copyright_items` refactor: extracted functions, externalized field definitions, implemented strategy pattern, improved error handling with custom exceptions.
- 2025-09-02: Added comprehensive integration plan for export, OSIRIS enrichment, relations update, and file existence stages (`export_enrichment_integration_plan.md`) and populated new TODO items for phased implementation.
- 2025-09-02: Revised integration plan to reflect existing normalized DB models for courses/persons (removed JSON cache approach), added sheet export refactor analysis & updated TODO/memory accordingly.
- 2025-09-02: Final code review confirms export functions completely missing (must recreate from legacy patterns), relations functions exist but need N+1 optimization, pipeline export stage stubbed. Updated plan/todo/memory to reflect critical path.
 - 2025-09-02: Hardened staged-processing in `easy_access/db/update.py`: removed `staged_item.__dict__` usage, added batched transactional processing, per-row error handling, and conditional deletion of processed staged rows; updated `.github/todo.md` with follow-ups.
- 2025-03-09: Completed repo-wide safety sweep: replaced `__dict__` usage with `vars()`, ad-hoc casts with `safe_*` helpers in `easy_access/db/update.py`, `easy_access/settings.py`, `easy_access/db/ingest.py`, and `dashboard/data.py`; updated imports.
- 2025-03-09: Routed complex staged rows into canonical merge path: modified `process_staged_raw_data` to detect non-trivial fields and delegate to `update_copyright_items` using extracted `merge_rules.py`.
 - 2025-09-02: Added parsing helpers (`safe_int`, `safe_float`, `safe_date`, `safe_enum`, `safe_compare_greater`) and unit tests `tests/test_safe_parsers.py`.
 - 2025-09-02: Added integration test `tests/test_integration_staging.py` that verifies staged-processing semantics (failed-first-run preserves staged rows; success run deletes processed rows).
 - 2025-09-02: Added `StagedProcessingFailure` model to `easy_access/db/models.py` and persisted per-row processing failures from `process_staged_raw_data` into this table for later inspection/retry.
 - 2025-09-02: Moved small safe parsing helpers (`safe_int`, `safe_float`, `safe_date`, `safe_enum`, `safe_compare_greater`) into `easy_access/utils.py` for reuse across the codebase; updated `easy_access/db/update.py` to use these helpers.
 - 2025-09-02: Completed repo-wide safety sweep: replaced `__dict__` usage with `vars()` in dashboard/components.py and easy_access/settings.py; replaced ad-hoc casts (int, float) with `safe_*` helpers in easy_access/settings.py, easy_access/db/ingest.py, dashboard/data.py; updated imports accordingly. No new unit tests added as existing safe_* tests cover the changes.
