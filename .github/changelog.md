# Changelog

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
