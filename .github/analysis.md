# Code Analysis & Background

This file contains the background, analysis, and rationale behind the refactor and dataflow redesign.

## Current Data Flow

The application processes copyright exports and faculty sheets, enriches data from external sources, computes derived fields, and generates reports. Historically, logic was spread across a monolithic `old_main.py`, `db/update.py`, and sheet helpers which caused side-effects and unclear priorities.

Key stages:
- Ingestion (raw exports, faculty sheets)
- Staging (staging tables in DB)
- Processing (merge/validation into main tables)
- Enrichment (external sources)
- Export (reports)

## What was refactored

- Centralized pipeline orchestration in `easy_access/pipeline.py` and simplified entrypoint in `easy_access/main.py`.
- Introduced staging tables (`StagedCopyrightItem`, `StagedFacultyUpdate`) and staging ingestion helpers in `easy_access/db/ingest.py`.
- Implemented minimal staged processors `process_staged_raw_data` and `process_staged_faculty_updates` in `easy_access/db/update.py` to provide a safe initial path from staging to main tables.

## Key design goals

- Make the DB the single source of truth.
- Unidirectional flow: Ingest -> Stage -> Process -> Export.
- Preserve legacy merge/priority rules but implement them in a testable, isolated module.
- Improve safety (transactions, atomic updates) and observability.

## Legacy code notes

- `old_main.py` contains the authoritative merge heuristics; preserve these by extracting them into `easy_access/merge_rules.py`.
- `db/base.py::copyright_item_from_dict` contains important normalization logic — keep and test this as the canonical factory.

## Tortoise ORM considerations

- Use `Tortoise.init(...)` and `Tortoise.generate_schemas(safe=True)` for idempotent setup.
- Use `async with in_transaction():` and savepoints for atomic staged processing.
- Prefer `QuerySet.update(...)` and `F` expressions for bulk DB-side updates.
- Use `prefetch_related(...)` for retrieval performance in FK/M2M loops.

## Recommendations (short)

- Harden staged processing, integrate `update_copyright_items` for full merges, and wrap processing in transactions.
- Replace `asyncio.run` in library code with async entrypoints and awrapper for CLI.
- Add unit & integration tests covering ingest -> stage -> process flows.

## Quick review of diffs vs starting commit (849628b9...)

I compared the current branch against commit `849628b965f4bd23b91400f1a5034eaf40787334` and found the following noteworthy changes:

- Files added: `easy_access/old_main.py`, `easy_access/pipeline.py` — legacy orchestrator preserved and new pipeline introduced.
- Files modified: `easy_access/main.py`, `easy_access/db/update.py`, `easy_access/db/ingest.py`, `easy_access/db/base.py`, `easy_access/db/models.py`, `easy_access/retrieve.py`, multiple sheet helpers and utilities.

Summary of risks and follow-ups discovered by the diff:

- `easy_access/main.py` was heavily simplified; ensure the new CLI entrypoint still calls `ensure_db_inited()` and closes Tortoise on exit.
- `process_staged_raw_data` exists but is conservative; route complex staged rows into the richer `update_copyright_items` path.
- Some files reference `__dict__` on model objects; replace with explicit safe mapping helpers to avoid leaking internal state or unserializable objects.
- Add tests that assert staging is only cleared after successful processing (per-row savepoints or batch atomicity).

I added a top-level todo item in `.github/todo.md` ("Critical review of refactor changes") to formalize this review and track any follow-ups. Add further follow-up items to that list as you find concrete code fixes.
