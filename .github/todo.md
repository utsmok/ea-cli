# Project TODO

This file lists the remaining high-impact tasks required to finish the refactor.
Items are small, testable, and ordered by impact. Per your instruction, we do NOT
change how `SETTINGS` is instantiated, and we DO NOT add HTML parsing unit tests
(we will instead validate parser outputs and debug if problems occur).

Guidance:
- Keep items small and testable.
- When an item is completed, add a one-line entry to `.github/changelog.md` and check it off here.

## Priority: Critical (immediate safety & correctness)
- [ ] Ensure export stage is read-only by default (export paths should not write to DB unless explicitly enabled via settings/CLI). Confirm `disable_writes` gating across export flows.

## Priority: High (testing, reliability & correctness)
- [ ] Add pipeline full E2E idempotency test: run full pipeline (ingest -> process -> enrich -> relations -> file-existence -> export) twice in an isolated test DB and assert the second run produces zero net changes (no duplicate links, no new enrich writes).
- [ ] Add relations raw-SQL integration test: exercise `link_courses` raw-SQL path by running without mocking `CopyrightItem.bulk_update` and assert expected links are created.
- [ ] Decouple test-aware logic from production code in these modules: `easy_access/db/relations.py`, `easy_access/maintenance/file_existence.py`, and `easy_access/enrichment/osiris.py` (where safe). Implement a small test adapter used only by tests rather than scattered mock-detection heuristics.
- [ ] Consolidate `QuerySetMock` and related test helpers into `tests/helpers.py` and update tests to import the shared helper.

## Priority: High (code quality & performance)
- [ ] Refactor `easy_access/pipeline._run_sync` to a documented, loop-safe dispatcher (avoid relying on nested `asyncio.run` in threads); add regression tests for loop-aware behavior.
- [ ] Optimize enrichment persistence: ensure `persist_courses` and `persist_persons` use DB-level bulk upserts for full payloads (keep a clear, documented test fallback path). Add unit tests asserting the DB-level bulk path is called when payloads are complete.
- [ ] Optimize `easy_access/db/relations.py` linking to reduce memory & N+1 risks (bulk M2M / grouped updates). Add micro-benchmarks or call-count tests to show improvements.

## Priority: Medium (export safety & validation)
- [ ] Add export dataframe schema validator + unit test: ensure required columns (e.g., `material_id`) exist and fail early with clear errors.
- [ ] Verify atomic Excel writes are used across all export paths and add failure-simulation tests (simulate write error leaving tmp file and ensure final file not corrupted).
- [ ] Improve `sheets/sheet.py` I/O performance for large exports (consider write_only mode or batched table writes); add benchmark.

## Priority: Medium (maintenance & reliability)
- [ ] Add unit tests for file-existence rate limit edge cases (including zero delay) and error handling for Canvas API responses.

## Priority: Medium (features & integration)
- [ ] Integrate backup module into pipeline: add optional pre/post backup stages with settings-driven enable/disable.
- [ ] Add calculate_derived_fields pipeline stage before export to compute non-persistent derived fields used only for reporting.

## Priority: Medium (maintenance & reliability)
- [ ] Add teardown leak-detection utility and harden Tortoise / aiosqlite shutdown ordering; run tests to verify no intermittent hangs.
- [ ] If pytest keeps hangs return [has not happened so far]: investigate and fix intermittent Tortoise-related pytest teardown hang: analyze `hang_diagnostics.txt` outputs, implement targeted aiosqlite shutdown, and ensure proper connection closure ordering.

## Priority: Low (docs & housekeeping)
- [ ] Update README with new pipeline stages, flags, and a short architecture diagram (PNG/SVG). Include quick-run commands for local testing.

## Completed / Already addressed (keep for history)
- [x] Atomic Excel write implemented in `easy_access/sheets/sheet.py` (temp file + rename).
- [x] Enrichment staleness and basic fetch-error tests added (`tests/test_enrichment_staleness.py`).
- [x] Unit tests for atomic write and loop-aware pipeline sync wrapper present.

## Notes / Assumptions
- We keep test-aware fallbacks for minimal test payloads where removing them would be high risk; the longer-term plan is to replace ad-hoc detection with explicit test adapters.
- HTML parsing unit tests are intentionally omitted per request; we will validate parser outputs instead and debug changes if/when parsing fails.

```

## Notes
- Work in the `new-dataflow` branch. Add changelog entries for completed items.
- Last cleanup: 2025-09-03 - Reorganized by actual priority, combined related items, removed duplicates, and verified against current repo state.
