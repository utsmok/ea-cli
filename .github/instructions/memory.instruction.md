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

(End of memory)
