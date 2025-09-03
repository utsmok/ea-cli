# Export, Enrichment & File Existence Integration Plan (Concise)

Date: 2025-09-03 (refreshed)

Purpose: Single source of truth for the post‑refactor export, enrichment, relations & file‑existence stages, their status, remaining gaps, and acceptance criteria.

## 1. Current State
Implemented pipeline stages (async, idempotent):
1. ingest_raw_data_async
2. ingest_faculty_updates_async
3. process_data_async
4. enrich_async (OSIRIS courses + persons)
5. update_relations_async (duplicates + course links)
6. refresh_file_existence_async (Canvas file API)
7. export_reports_async (faculty/programme/all-items + overviews)

Status summary:
- Phase A (export & relations): COMPLETE ✅
- Phase B (OSIRIS enrichment core): COMPLETE ✅ (selection & parsing tests still missing)
- Phase C (file existence): COMPLETE ✅ (broad unit + integration tests present)
- Phase D (performance, documentation & residual test gaps): IN PROGRESS 🚧

## 2. Design Principles (Retained)
- Idempotent reruns; zero side‑effect second pass.
- Separation of retrieval / transform / persist / export.
- Settings‑driven concurrency + TTLs.
- Incremental refresh (only stale / missing).
- Observability via structured log counts + durations.
- Testability with mocked IO (httpx) & small fixtures.

## 3. Implemented Modules
| Module | Role | Notable Gaps |
|--------|------|-------------|
| `easy_access/sheets/export.py` | Orchestrates all Excel exports | Lacks atomic write + dataframe validation helper |
| `easy_access/enrichment/osiris.py` | Course & person enrichment (concurrent) | TTL selection + `_fetch_course_details` path tests missing; per-row create (no bulk) |
| `easy_access/db/relations.py` | Duplicate detection & course linking (batch) | Raw SQL link path not covered by tests |
| `easy_access/maintenance/file_existence.py` | TTL-based Canvas file existence | Mostly covered; add rate limiting edge test (0 delay) |
| `easy_access/pipeline.py` | Stage orchestrator + CLI flags | Full end‑to‑end idempotency test missing |

## 4. Confirmed Test Coverage (Snapshot)
- Export: unit + integration (uniqueness, orchestration) ✅
- Relations: unit tests for duplicate + link + orchestration (mocked bulk paths) ✅ (missing raw-SQL path)
- File existence: selection, single check, batch update, refresh orchestration (incl. concurrency & errors) ✅
- Enrichment: fetch_course_data, fetch_person_data, persist_* basics ✅ (selection/TTL + detail fetch + orchestrator not explicitly asserted)

## 5. Remaining High-Value Gaps
1. Enrichment selection TTL logic tests (courses & persons) + negative cases (no stale, all stale, mixed age).
2. Enrichment `_fetch_course_details` branch & error handling (HTTP 500 / malformed JSON) tests.
3. Enrichment orchestrator end-to-end test with mocked HTTP + ensuring persons fetched only from returned course data.
4. Relations raw SQL link creation path integration test (no mocked `bulk_update`).
5. Pipeline full E2E test: run all stages twice; assert second run performs zero new links / zero new enrich ops (log or DB delta assertion).
6. Export atomic write helper (temp file + rename) with failure simulation test.
7. Export dataframe schema validator + test (required columns present; raise early if missing).
8. Bulk insertion optimization for `persist_courses` / `persist_persons` (replace per-item `create` with `bulk_create`) + perf assertion (time or call count).
9. Optional: add lightweight performance benchmark harness (timing a mid-size dataset) to guard regressions (behind marker, not default).
10. Documentation: architecture diagram (PNG/SVG), README section on stage idempotency & CLI flags, minimal troubleshooting table.
 11. Decouple production code from test mocks in `db/relations.py` (remove Mock-aware branching) + adjust tests.
 12. Refactor pipeline synchronous wrappers to avoid nested `asyncio.run` misuse; provide loop-safe dispatcher.
 13. Consolidate duplicated `QuerySetMock` (remove local definition in `tests/test_enrichment.py`).
 14. Strengthen teardown reliability: isolate aiosqlite / Tortoise connection closure ordering; add leak detection utility.

## 6. Risks / Technical Debt
- Per-row persistence may become bottleneck at scale (optimize before large dataset adoption).
- Lack of atomic Excel writes risks partial files on crash / interruption.
- Absence of idempotency regression test could hide duplicate linking regressions.
- Enrichment scraping selectors brittle to upstream HTML changes (need selector health test / fallback strategy).
- Test-contaminated production logic in `db/relations.py` increases complexity & obscures true runtime paths.
- Repeated `asyncio.run` in library code limits embedding in other async systems (potential event loop errors).
- Duplicate test helpers risk divergent behavior (QuerySetMock variants).

## 7. Acceptance Criteria (Final State)
| Area | Criteria |
|------|----------|
| Enrichment | TTL selection & detail fetch paths fully unit-tested; bulk create optimization implemented |
| Relations | Both mocked and raw-SQL paths covered; idempotent second-run verified |
| File Existence | Concurrency + rate-limiting edge (0 delay) tested; metrics logged |
| Export | Atomic write + schema validation; uniqueness & rerun behavior tested |
| Pipeline | Full E2E (all stages) green twice in a row; second run zero deltas |
| Docs | Updated README + diagram + troubleshooting; plan & TODO reflect reality |

## 8. Immediate Next Steps (Ordered)
1. Add enrichment selection & TTL tests (courses/persons) + orchestrator test.
2. Implement atomic_excel_save(file) utility + integrate into export paths.
3. Decouple `db/relations.py` from mocks; refactor tests accordingly (maintain coverage for duplicate + linking + raw SQL path).
4. Refactor pipeline sync wrappers to be loop-aware (avoid nested asyncio.run) and add regression test.
5. Add idempotent pipeline E2E double-run test (assert zero deltas).
6. Add bulk_create optimization for new courses/persons (retain safe fallback) + unit test verifying call counts.
7. Cover relations raw SQL path by disabling / not mocking `bulk_update` in integration test.
8. Consolidate `QuerySetMock` usage (remove duplication).
9. Enhance teardown diagnostics to isolate any lingering connections and threads; add automated leak assertion.

## 9. Deferred / Optional Enhancements
- Structured metrics emitter (JSON logs -> future dashboard).
- Materialized view for export retrieval if row counts grow (>100k) to reduce memory.
- Retry / backoff strategy for transient 5xx in enrichment detail fetch.

## 10. Maintenance Notes
- All new tests should avoid real network; mock `httpx.AsyncClient`.
- Keep test data small; prefer deterministic timestamps via `freezegun` (future) or manual patching of `datetime.now`.

## Current state
- Ingest and processing pipeline stages are implemented (staging -> processing -> main tables).
- Phase A (export + relations) completed and manually validated.
- Phase B (OSIRIS enrichment) implemented; unit tests for parsing/staleness selection are still outstanding.
- Phase C (file-existence checking) implemented.
- Phase D (performance, docs, comprehensive tests) is remaining.

## Design principles
- Deterministic, idempotent stages.
- Clear separation of retrieval, transformation, persistence, and export.
- Settings-driven concurrency and TTLs.
- Observability: counts, durations, deltas.

## Modules (implemented or planned)
- `easy_access/sheets/export.py` — export orchestrator (implemented).
- `easy_access/enrichment/` — OSIRIS/person enrichment (implemented; tests pending).
- `easy_access/db/relations.py` — optimized relations (implemented).
- `easy_access/maintenance/file_existence.py` — file existence TTL checking (implemented).
- `easy_access/pipeline.py` — pipeline stages: ingest, process, enrich, relations, file-existence, export (implemented with async entrypoints and CLI flags).

## (Superseded Older Section) Remaining work (see sections above instead)
This section is kept for reference; should not be used directly.

## Export notes
- Export files follow the existing, validated patterns: per-faculty overviews, per-programme sheets, and an all-items sheet. File-uniqueness and data-entry sheet formatting implemented.

## Enrichment notes
- Enrichment is DB-centric: fetch missing/stale Course and Person records, persist directly to models, then run relations linking.
- Settings keys used: `enrichment.course_ttl_days`, `enrichment.person_ttl_days`.

## File existence notes
- Uses `file_exists` and `last_canvas_check` fields; re-check when NULL or older than TTL.

## Pipeline ordering (final)
1. ingest_raw_data_async
2. ingest_faculty_updates_async
3. process_data_async
4. enrich_async (optional via flag)
5. update_relations_async
6. refresh_file_existence_async
7. export_reports_async

## Tests and verification (updated)
- See sections 4–8 above for precise gaps and priorities.

### Implementation Steps
1. Add missing unit tests for all modules
2. Implement performance optimizations
3. Add comprehensive integration tests
4. Update documentation
5. Add monitoring and benchmarking

### Acceptance Criteria
- All critical functions have unit test coverage
- Performance benchmarks meet requirements
- Documentation is complete and accurate
- Data persistence is reliable and atomic
- Integration tests pass consistently

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

