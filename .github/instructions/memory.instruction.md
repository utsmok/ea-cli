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

Testing / safety notes:
- Staged processing hardened: explicit field mapping, batched transactions, per-row failure persistence (`StagedProcessingFailure`).
- `safe_*` parsing helpers added in `easy_access/utils.py` with unit tests `tests/test_safe_parsers.py`.
- Pytest teardown improvements added to reduce hangs; diagnostics written to `hang_diagnostics.txt` when needed.

Operational notes:
- Prefer `new-dataflow` branch when acting on the repo.
- When adding todo items, also add a one-line changelog entry.

External review anchors (2025-09-03):
- Production/test coupling detected in `db/relations.py` (Mock-aware logic) slated for removal.
- Need pipeline asyncio wrapper refactor to avoid nested `asyncio.run` in library code.
- Duplicate `QuerySetMock` (tests/helpers vs test_enrichment) to be consolidated.
- Pending enrichment TTL tests, raw SQL relations path test, atomic Excel write, bulk_create optimization.

This memory is intentionally concise and focused on actionable anchors.
