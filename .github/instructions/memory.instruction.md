---
applyTo: '**'
---

# Short project memory for ea-cli

This memory captures the current refactor state, priorities, and a few repo anchors derived from the analysis, todo and changelog artifacts.

- repo: ea-cli (branch: new-dataflow)
- refactor baseline commit provided by user: 849628b965f4bd23b91400f1a5034eaf40787334

Key goals and invariants
- Make the DB the single source of truth; adopt a unidirectional flow: Ingest -> Stage -> Process -> Export.
- Preserve legacy merge/priority rules (source: `easy_access/old_main.py`) and encapsulate them in a testable module (planned `easy_access/merge_rules.py`).
- Use `copyright_item_from_dict` in `easy_access/db/base.py` as the canonical normalizer.

Top-priority todo items (from `.github/todo.md`)
- Critical review of refactor changes (compare current branch to commit above, generate per-file rationales, update analysis/changelog) — added 2025-09-02.
- Critical review of refactor changes (compare current branch to commit above, generate per-file rationales, update analysis/changelog) — completed 2025-09-02; see `.github/critical-review.md`.
- Harden staged processing: remove `__dict__` usage, validate/normalize staged rows, route complex merges into `update_copyright_items`, and wrap processing in transactions/savepoints.
- Add tests ensuring staging is only cleared on successful processing.
 - Hardened staged processing (2025-09-02): `process_staged_raw_data` now uses an explicit staged field mapping, batched `in_transaction()` blocks, per-row error handling, and deletes only successfully-processed staged rows. Further type/validation helper work remains.
 - Next: add `safe_*` parsing helpers and unit tests for staged-processing; route complex merges into canonical `update_copyright_items` after validation.
 - Hardened staged processing (2025-09-02): `process_staged_raw_data` now uses an explicit staged field mapping, batched `in_transaction()` blocks, per-row error handling, and deletes only successfully-processed staged rows. Per-row failures are persisted in `StagedProcessingFailure` for inspection and retry.
 - Implemented small `safe_*` parsing helpers and moved them to `easy_access/utils.py` (2025-09-02). Unit tests `tests/test_safe_parsers.py` and integration test `tests/test_integration_staging.py` were added and pass locally.
 - CI/GitHub Actions is deprioritized until a repo-wide safety sweep is completed.

Tortoise ORM guidance recorded
- Use `Tortoise.init(...)` and `Tortoise.generate_schemas(safe=True)` for idempotent setup; call `Tortoise.close_connections()` on exit.
- Use `async with in_transaction()` and per-row/batch savepoints for atomic staged processing.
- Prefer `QuerySet.update(...)`, F-expressions and `prefetch_related(...)` to reduce N+1 and leverage DB-side operations.
- Use Aerich for migrations and `tortoise.contrib.pydantic` for serialization in APIs/tests.

Changelog anchors
- 2025-03-06: initial analysis added.
- 2025-03-07: todo extracted.
- 2025-03-08: analysis.md added.
- 2025-09-02: critical review todo and diff summary added (this memory references that review).

Notes for future interactions
- When acting on the repo, prefer the `new-dataflow` branch and consider the baseline commit above for comparisons.
- Treat the staging tables and `update_copyright_items` as authoritative integration points.
- When adding tasks to the todo, also add a one-line changelog entry.
- `.github/analysis.md` contains a high level overview of the refactor
- `.github/changelog.md` contains a changelog for this project; when you log changes here in your memory also update the changelog with a more detailed update
- `.github/todo.md` contains a detailed todo list that should be kept updated alongside your memory
 - `.github/critical-review.md` contains the per-file findings and concrete follow-ups produced during the 2025-09-02 review.
 - New (2025-09-02 Final): Code review confirms export functions completely removed from refactor; relations functions exist but need N+1 optimization & pipeline integration; export stage stubbed in pipeline. Updated plan reflects legacy `old_main.py` patterns: faculty/program/overview/all_items sheets with file uniqueness, mature `DataEntrySheet` formatting. Phase A critical path: recreate export orchestrator + optimize relations.
- Repo-wide safety sweep completed 2025-09-02: replaced `__dict__` with `vars()`, ad-hoc casts with `safe_*` helpers in key files; no new tests needed as existing ones cover.
- Phase 2 refactor completed 2025-09-02: `update_copyright_items` broken down into modular functions with strategy pattern, custom exceptions, and Settings integration.
- Phase 4 completed 2025-03-09: Added DateFieldStrategy and EnumFieldStrategy for enhanced field comparisons, implemented comprehensive unit test coverage (50 tests in `tests/test_update_refactor.py`), and validated all refactored functions with real data processing. All 63 tests pass.
- **MAJOR MILESTONE**: Complete `update_copyright_items` refactor finished 2025-03-09: All 5 phases implemented including comprehensive integration tests with real data (7 tests in `tests/test_integration_real_data.py`), 50 unit tests covering all refactored components, strategy pattern for field comparisons, custom exceptions, Settings integration, and full validation. The 400+ line monolithic function has been successfully broken down into modular, testable components.
- 2025-09-02: Implemented comprehensive admin CLI for `StagedProcessingFailure` management with commands: `inspect-failures`, `failure-stats`, `retry-failures`, `cleanup-failures`; includes Trogon TUI support for enhanced user experience.
- 2025-09-02: Added CLI flags for individual pipeline stages: `--ingest-only`, `--process-only`, `--export-only` to `run.py process` command, allowing developers to run specific stages of the data processing pipeline.
- 2025-09-02: Converted `easy_access/pipeline.py` to provide async entrypoints (`run_async`, `ingest_raw_data_async`, etc.) and thin sync wrappers; removed `asyncio.run` from library-level code for better async compatibility.

(End of memory)

(End of memory)
