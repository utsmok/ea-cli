# Critical review of refactor changes

Date: 2025-09-02

Summary
- Compared branch `new-dataflow` against baseline commit `849628b965f4bd23b91400f1a5034eaf40787334`.
- Diff: 15 files changed, ~1.6k insertions, ~1.4k deletions.
- Purpose: capture per-file rationale, identify risks, and list concrete follow-up tasks.

Key findings (short)
- The refactor centralizes orchestration and introduces staging tables and a simple pipeline. Good separation intent.
- High-risk area: staged -> main processing (`easy_access/db/update.py`) uses unsafe model introspection and clears staging unconditionally.
- Secondary risks: mixing sync/async (use of `asyncio.run` inside library code), legacy heuristics preserved in `old_main.py` but not extracted to testable modules.

Per-file notes & follow-ups

- `easy_access/db/update.py` (HIGH)
  - Rationale: contains staged processors and core update flow; currently uses `staged_item.__dict__`, performs per-row DB saves without transactions, and clears staging at the end.
  - Risk: data loss on partial failure, leaked ORM internals, unexpected types in merges.
  - Follow-ups:
    - Replace `__dict__` usage with an explicit safe mapping helper (use the model fields list).
    - Wrap processing in transactions/savepoints and process in batches. Only clear staging after successful commit.
    - Route complex staged rows to `update_copyright_items` for canonical merging.
    - Add per-row error handling and persist failed rows for later inspection.

- `easy_access/pipeline.py` (MEDIUM)
  - Rationale: new pipeline orchestration uses `asyncio.run` inside methods and calls DB functions; lacks async/sync contract.
  - Risk: broken event loop usage when called from async contexts and inconsistent lifecycle handling.
  - Follow-ups: provide async entrypoints + thin sync wrapper for CLI; ensure DB init/close is handled by pipeline.

- `easy_access/main.py` (MEDIUM)
  - Rationale: simplified to instantiate `DataPipeline` and call `run()`.
  - Follow-ups: ensure DB init (`ensure_db_inited`) and `Tortoise.close_connections()` are called in pipeline start/stop.

- `easy_access/old_main.py` (INFO / MEDIUM)
  - Rationale: preserved legacy heuristics and design notes—valuable reference.
  - Follow-ups: extract merge heuristics into `easy_access/merge_rules.py` and add unit tests that codify priority ordering and remark merging.

- `easy_access/db/ingest.py` (MEDIUM)
  - Rationale: ingestion should enforce normalization before writing to staging.
  - Follow-ups: validate/normalize at ingest time (cast `material_id`, required fields) and mark/reject incomplete rows.

- `easy_access/db/base.py` (INFO)
  - Rationale: contains `copyright_item_from_dict` canonical normalizer.
  - Follow-ups: ensure staged processors reuse this normalizer for new-item creation.

- `easy_access/db/models.py` (INFO)
  - Rationale: staging models keep fields as strings; helpful for mapping the safe keys.
  - Follow-ups: implement a small helper (e.g., `staged_to_dict`) using the model definition to build safe, typed dicts.

- `easy_access/sheets/*` and `easy_access/utilities/file_exists.py` (LOW/MEDIUM)
  - Rationale: changes around file-exists logic and sheet comparisons.
  - Follow-ups: ensure boolean normalization for `file_exists` and add tests for `add_file_exists()` path.

- `easy_access/utils.py`, `run.py` (LOW)
  - Follow-ups: sanity-check CLI wiring and Directory/File helpers after pipeline changes.

Immediate recommended next work (in order)
1. Fix `process_staged_raw_data` safety: remove `__dict__`, add transaction/savepoint, per-row error handling, and only clear staging on success.
2. Add unit/integration tests: staging -> processing happy path and failure path (staging retained on failure).
3. Extract `merge_rules.py` from `old_main.py` and add unit tests encoding merge priorities.
4. Convert `DataPipeline` methods to async with a sync wrapper; ensure DB lifecycle calls are present.
5. Improve logging and add a persistent failed-rows report for staged-processing.

Where to find more details
- The unified diff was written to `.github/critical-review.patch` and the file list to `.github/critical-diff-files.txt`.

End of review
