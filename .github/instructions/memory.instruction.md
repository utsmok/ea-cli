---
applyTo: '**'
---

# Short project memory for ea-cli

This file captures key repo anchors, goals, and the current status of the new-dataflow refactor.

- repo: ea-cli (branch: new-dataflow)
- baseline commit referenced: 849628b965f4bd23b91400f1a5034eaf40787334

Key invariants:
- DB is the single source of truth; pipeline is: Ingest -> Stage -> Process -> Export.
- Preserve legacy merge rules via `easy_access/merge_rules.py` and normalize inputs with `copyright_item_from_dict`.

Recent milestones (trimmed):
- Phase A (export & relations): implemented `sheets/export.py` and `easy_access/db/relations.py`; integrated into pipeline; manual export validation completed.
- Phase B (enrichment): OSIRIS-person enrichment implemented; unit tests for parsing/stale-selection remain.
- Phase C (file-existence): TTL-based file existence checks implemented.
- Ongoing: Phase D – performance tuning, comprehensive tests, and documentation.
 - 2025-09-05: Added enrichment staleness + fetch-error tests (`tests/test_enrichment_staleness.py`) validating TTL selection for courses/persons and basic network error handling for fetchers. Also added a test for `fetch_and_parse_courses` to ensure MissingCourse is recorded for missing data.
 - Pending enrichment TTL tests (additional HTML parsing + orchestrator idempotency), raw SQL relations path test, atomic Excel write (unit tests in place), bulk_create optimization.
 - Current Focus: Completed initial enrichment staleness and error tests; next actions when resuming:
	 - Add HTML parsing unit tests for `fetch_person_data` and `_fetch_course_details` covering selector fallback and cookie-wall cases.
	 - Add orchestrator idempotency test for `enrich_async` (run twice -> no duplicate DB writes, MissingCourse backoff respected).
	 - Extend fetch-error tests to include HTTP 500, malformed JSON, and timeouts.
 - How to resume: checkout branch `new-dataflow`, run `uv run pytest tests/test_enrichment_staleness.py` to validate the small suite, then add tests in `tests/test_enrichment_parsing.py` focusing on HTML fixtures in `tests/fixtures/`.

Operational notes:
- Prefer `new-dataflow` branch when acting on the repo.
- When adding todo items, also add a one-line changelog entry.

External review anchors (2025-09-03):
- Production/test coupling detected in `db/relations.py` (Mock-aware logic) slated for removal.
- Need pipeline asyncio wrapper refactor to avoid nested `asyncio.run` in library code.
 - Production/test coupling detected in `db/relations.py` (Mock-aware logic) slated for removal. ✅ Partially addressed: relations module refactored to remove fragile branching and use deterministic resolution + single-call bulk_update.
- Pipeline asyncio wrapper refactor: implemented `_run_sync` helper and updated sync wrappers in `easy_access/pipeline.py`.
- Duplicate `QuerySetMock` (tests/helpers vs test_enrichment) to be consolidated.
- Pending enrichment TTL tests, raw SQL relations path test, atomic Excel write, bulk_create optimization.
- Backup module integration: add pre/post stages to pipeline for file backup/restore.
- Legacy cleanup: remove globals from settings.py **COMPLETED** - global constants refactored to Settings properties. Note: the export stage still contained a DB update call; it has been changed to be read-only by default (gated by `disable_writes`), so documentation and code are now aligned.

This memory is intentionally concise and focused on actionable anchors.
