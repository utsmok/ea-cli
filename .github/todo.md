# Project TODO

Guidance: keep entries short, and move completed items to the changelog with a short note.

When you complete an item, add a one-line entry to `.github/changelog.md` (date,  summary).

## 0 Critical review of refactor changes
- [x] Perform a thorough, critical review of the refactor changes already made comparing current branch to commit `849628b965f4bd23b91400f1a5034eaf40787334`:
  - [x] Produce a file-level summary of changed files and a short rationale for each change. (see `.github/critical-review.md`)
  - [x] Identify required follow-up code changes and add them to this todo list. (see follow-ups below)
  - [x] Update `.github/analysis.md` and `.github/changelog.md` with review findings and follow-up actions.
  - Owner/ETA: (unassigned) — review completed 2025-09-02

### Follow-ups from critical review
- [ ] Replace `__dict__` usage in `easy_access/db/update.py::process_staged_raw_data` with an explicit safe mapping and typed conversion. (high)
- [ ] Wrap staged processing in transactions/savepoints and process in batches; only clear staging after successful commit. (high)
 - [x] Replace `__dict__` usage in `easy_access/db/update.py::process_staged_raw_data` with an explicit safe mapping and typed conversion. (high) — implemented 2025-09-02
 - [x] Wrap staged processing in transactions/savepoints and process in batches; only clear staging after successful commit. (high) — implemented 2025-09-02
- [ ] Route complex staged rows to `update_copyright_items` for canonical merge logic. (high)
- [ ] Add per-row error capture and persist failed staged rows to a CSV/table for later retry. (high)
- [ ] Extract merge heuristics from `easy_access/old_main.py` into `easy_access/merge_rules.py` and add unit tests that encode the priority rules. (medium)
- [ ] Convert `easy_access/pipeline.py` to provide async entrypoints and a thin sync wrapper; remove `asyncio.run` from library-level methods. (medium)
- [ ] Add unit and integration tests: staging success, staging failure (staging retained), `copyright_item_from_dict`.
 - [ ] Add unit and integration tests: staging success, staging failure (staging retained), `copyright_item_from_dict`.
 - [ ] Implement small `safe_*` parsing helpers used by staged processors and add unit tests for the helpers (next immediate task).
- [ ] Add logging improvements for staged processing (include material_id, faculty, error trace).
- [ ] Normalize `file_exists` values and add tests for `add_file_exists()` flows.
- [ ] Create a developer note in README describing the new pipeline and where to find legacy heuristics.

This file holds the canonical, actionable to-do list for the dataflow refactor.
Update this file when you make changes: check off items, add details, or create new items.


## 1 Core correctness & safety
- [ ] Harden staged processing (`easy_access/db/update.py::process_staged_raw_data`)
  - [ ] Replace `staged_item.__dict__` usage with an explicit safe mapping / `to_dict()` helper.
  - [ ] Route complex rows to the existing `update_copyright_items` merge logic instead of ad-hoc updates.
  - [ ] Ensure staged rows are validated/normalized (reuse `copyright_item_from_dict` where appropriate).
  - [ ] Only clear staging on full success (transactional / atomic behavior or compensating rollback).
- [ ] Add item-level error handling so a single failing row does not abort the whole run without safe reporting.

## 2 Merge rules & legacy behavior preservation
- [ ] Extract merge heuristics from `easy_access/old_main.py` into `easy_access/merge_rules.py`.
- [ ] Add unit tests codifying the priority rules (timestamp precedence, workflow-status ranking, manual-classification ordering, remark merging).

## 3 API stability & runtime behavior
- [ ] Remove/replace `asyncio.run` calls in library helpers; provide async entrypoints and thin sync wrappers for CLI use.
- [ ] Add CLI flags or `EasyAccessSettings` options to run individual stages (ingest-only, process-only, export-only, full-run).

## 4 Observability & error handling
- [ ] Improve logging (include material_id, faculty, and stage) for success and failure cases.
- [ ] Add retries/backoff for transient I/O or DB contention errors.

## 5 Testing and quality gates
- [ ] Unit tests for ingestion helpers: `load_raw_copyright_data_to_staging`, `load_faculty_updates_to_staging` (use small polars DataFrames).
- [ ] Integration tests for `DataPipeline.run()` against a temporary sqlite DB with a small sample dataset.
- [ ] Tests that assert staging tables are only cleared on successful processing.
- [ ] Add linting/type checks and include them in CI pipeline.

## 6 Export / Reports
- [ ] Implement `DataPipeline.export_reports` and wire existing report-generation code into it.
- [ ] Add smoke tests for generated reports (columns present, row counts, basic values).

## 7 Performance & bulk operations
- [ ] Optimize bulk create/update (batch sizes, use of `on_conflict` where supported by ORM/bulk helpers).
- [ ] Profile pipeline on representative dataset and address hotspots.

## 8 Documentation & developer ergonomics
- [ ] Update README with pipeline flow, CLI usage, and development notes.
- [ ] Add a short migration note for developers about `old_main.py` and where to find the new entrypoints.

## Small actionable items / quick wins
- [ ] Implement async/sync wrapper helper for running coroutines from CLI code.
- [ ] Add unit tests for `copyright_item_from_dict`.
- [ ] Add a README note marking `load_raw_copyright_data` as deprecated and pointing to staging loaders.

