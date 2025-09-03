# Changelog

Recent milestones

- 2025-09-03: Phase D progress — added unit tests for export; improved pytest teardown and diagnostics (`hang_diagnostics.txt`) to reduce intermittent hangs.
- 2025-09-03: Phase B completed — OSIRIS enrichment implemented (concurrent fetch, TTLs, bulk persistence). Parsing and stale-selection unit tests pending.
- 2025-09-02: Phase A completed — export orchestrator (`sheets/export.py`) and optimized relations (`easy_access/db/relations.py`) implemented and integrated; manual export validation completed.
- 2025-09-02: Pipeline improvements — async entrypoints and CLI flags added for stage control (`--ingest-only`, `--process-only`, `--export-only`, `--enrich-only`).
- 2025-09-02: Safety and reliability — replaced `__dict__` uses, added `safe_*` parsers, transactionized staged processing, and added `StagedProcessingFailure` persistence.
- 2025-03-09: Major refactor — `update_copyright_items` decomposed into modular, testable components with strategy patterns and broad test coverage.

Top remaining work

- Add unit tests for enrichment (HTML parsing and stale-selection logic) and relations (batch linking / N+1 elimination).
- Add integration tests for end-to-end pipeline runs and export file verification in temp directories.
- Phase D: performance tuning (bulk M2M linking, export memory usage) and documentation (README, architecture diagram).
