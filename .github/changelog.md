# Changelog

Recent milestones

- 2025-09-03: Incorporated external code review; added tasks for decoupling test mocks from production (`db/relations.py`), pipeline asyncio wrapper refactor, helper consolidation, and teardown leak detection.
- 2025-09-03: Fixed person data enrichment (URL encoding, cookie wall detection, resilient selectors) and enforced canonical workflow_status priority (preventing Done -> ToDo downgrades & repeat updates).
- 2025-09-03: Legacy cleanup completed - deleted `old_main.py` and `sheets/enrichment.py`; removed legacy relations functions from `db/update.py` and deprecated `load_raw_copyright_data` from `db/ingest.py`.
- 2025-09-03: Consolidated integration plan (export/enrichment/relations/file existence) and refreshed TODO with precise Phase D gaps (atomic Excel writes, enrichment TTL tests, raw SQL relations path, bulk persistence optimization).

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
- 2025-09-03: Major refactor — `update_copyright_items` decomposed into modular, testable components with strategy patterns and broad test coverage.
- 2025-09-03: Settings cleanup analysis completed - confirmed global constants (DEPARTMENT_MAPPING, COURSE_MAPPING, FINE_AMOUNT) have been properly refactored into Settings class properties. Global SETTINGS singleton is acceptable pattern.

Top remaining work

- Add unit tests for enrichment (HTML parsing and stale-selection logic) and relations (batch linking / N+1 elimination).
- Add integration tests for end-to-end pipeline runs and export file verification in temp directories.
- Phase D: performance tuning (bulk M2M linking, export memory usage) and documentation (README, architecture diagram).

- 2025-09-08: Consolidated TODO items and archived `export_enrichment_integration_plan.md` to `.github/completed_or_old/`.
