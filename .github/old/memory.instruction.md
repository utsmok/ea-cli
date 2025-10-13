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
- 2025-09-05: Added unit tests for atomic Excel write and loop-aware pipeline sync wrapper. Pipeline sync wrapper (`_run_sync`) implemented to avoid nested event loop errors.
- 2025-09-03 Fixes: Person enrichment robustness (URL encoding, cookie wall detection, selector fallback) and workflow_status canonical priority & downgrade guard (prevent repeated Done->ToDo updates). Regression tests added.
- **2025-09-XX: Comprehensive code review completed** - Individual file analyses and holistic review reports created in `.github/instructions/`

Code Quality Findings (from recent review):
- **Type Hints**: Inconsistent coverage - needs completion across all files
- **Documentation**: Variable quality - critical functions lack docstrings
- **Security**: Hardcoded credentials in `api_keys.py` - HIGH PRIORITY FIX
- **File Size**: Several files >800 lines should be decomposed (`settings.py`, `models.py`, `sheet.py`)
- **Error Handling**: Generic exception handling needs standardization
- **Testing**: Limited coverage - comprehensive test suite needed
- **Performance**: Good foundation but monitoring needed for large datasets

Implementation Priorities (from code review):
2. **HIGH**: Complete type hint coverage across all files
3. **HIGH**: Add comprehensive docstrings to public functions
4. **MEDIUM**: Decompose large files (>500 lines)
6. **MEDIUM**: Standardize error handling patterns

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
 - Production/test coupling detected in `db/relations.py` (Mock-aware logic) slated for removal. ✅ Partially addressed: relations module refactored to remove fragile branching and use deterministic resolution + single-call bulk_update.
- Pipeline asyncio wrapper refactor: implemented `_run_sync` helper and updated sync wrappers in `easy_access/pipeline.py`.
- Duplicate `QuerySetMock` (tests/helpers vs test_enrichment) to be consolidated.
- Pending enrichment TTL tests, raw SQL relations path test, atomic Excel write, bulk_create optimization.
- Backup module integration: add pre/post stages to pipeline for file backup/restore.
 - 2025-09-10: `easy_access/sheets/backup.py` added with timestamped move and manifest helpers; pipeline integration is pending.
 - 2025-09-10: Added `export_faculty_workflow_files` in `easy_access/sheets/export.py` and wired an opt-in CLI flag `--export-workflow` that enables per-faculty workflow exports while keeping the legacy exporter as default. Changes committed on branch `export-workflow`.
- Legacy cleanup: remove globals from settings.py **COMPLETED** - global constants refactored to Settings properties. Note: the export stage still contained a DB update call; it has been changed to be read-only by default (gated by `disable_writes`), so documentation and code are now aligned.

This memory is intentionally concise and focused on actionable anchors.
