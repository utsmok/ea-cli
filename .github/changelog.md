# Changelog

This file records high-level repository changes and analysis edits.

- 2025-03-06: Initial analysis and refactor plan added to `.github/code_analysis.md`.
- 2025-03-07: Extracted `todo.md` into `.github/todo.md` (canonical checklist).
- 2025-03-08: Added `.github/analysis.md` containing background, design goals, and Tortoise ORM recommendations.
- 2025-09-02: Added "Critical review of refactor changes" todo and summarized diffs against commit `849628b965f4bd23b91400f1a5034eaf40787334` (files added/modified and follow-up recommendations).
- 2025-09-02: Completed critical review; added `.github/critical-review.md` with per-file findings and created follow-up tasks in `.github/todo.md`.
 - 2025-09-02: Hardened staged-processing in `easy_access/db/update.py`: removed `staged_item.__dict__` usage, added batched transactional processing, per-row error handling, and conditional deletion of processed staged rows; updated `.github/todo.md` with follow-ups.
 - 2025-09-02: Added parsing helpers (`safe_int`, `safe_float`, `safe_date`, `safe_enum`, `safe_compare_greater`) and unit tests `tests/test_safe_parsers.py`.
 - 2025-09-02: Added integration test `tests/test_integration_staging.py` that verifies staged-processing semantics (failed-first-run preserves staged rows; success run deletes processed rows).
 - 2025-09-02: Added `StagedProcessingFailure` model to `easy_access/db/models.py` and persisted per-row processing failures from `process_staged_raw_data` into this table for later inspection/retry.
 - 2025-09-02: Moved small safe parsing helpers (`safe_int`, `safe_float`, `safe_date`, `safe_enum`, `safe_compare_greater`) into `easy_access/utils.py` for reuse across the codebase; updated `easy_access/db/update.py` to use these helpers.
 - 2025-09-02: Completed repo-wide safety sweep: replaced `__dict__` usage with `vars()` in dashboard/components.py and easy_access/settings.py; replaced ad-hoc casts (int, float) with `safe_*` helpers in easy_access/settings.py, easy_access/db/ingest.py, dashboard/data.py, and pdf_downloads/moondream_test.py (if exists); updated imports accordingly. No new unit tests added as existing safe_* tests cover the changes.
