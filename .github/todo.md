# Project TODO

This file collects the actionable tasks for the dataflow refactor in priority order. It includes completed items for visibility.

Guidance:
- Keep items small and testable.
- When an item is completed, add a one-line entry to `.github/changelog.md` and check it off here.

## Priority: Immediate (safety & correctness)
- [x] Repo-wide safety sweep: replace ad-hoc casts and `__dict__` usage with `safe_*` helpers and explicit mappings. Add unit tests for any changed codepaths. (high)
- [ ] Route complex staged rows into canonical merge path: ensure `process_staged_raw_data` delegates non-trivial merges to `update_copyright_items` (use `copyright_item_from_dict` and `merge_rules`). (high)
- [ ] Add logging improvements for staged processing (include material_id, faculty, stage, and compact error traces). (high)
- [ ] Implement failure-inspection/retry helper: admin CLI or small script to list `StagedProcessingFailure` rows and requeue or attempt automated retries. (high)

## Priority: High (reliability & observability)
- [ ] Add unit tests for `copyright_item_from_dict` and merge heuristics in `easy_access/merge_rules.py`. (high)
- [ ] Normalize `file_exists` values and add tests for any `add_file_exists()` or related flows. (high)
- [ ] Add item-level error handling coverage and tests so single-row failures do not hide regressions. (high)

## Priority: Medium (developer ergonomics & API)
- [ ] Convert `easy_access/pipeline.py` to provide async entrypoints and thin sync wrappers; remove `asyncio.run` from library-level code. (medium)
- [ ] Provide CLI flags or `Settings` options to run individual stages (ingest-only, process-only, export-only). (medium)
- [ ] Implement the admin retry workflow (UI/CLI) for `StagedProcessingFailure`. (medium)

## Priority: Low (infrastructure & performance)
- [ ] Add GitHub Actions CI to run `uv run pytest` on push/PR after the repo-wide safety sweep. (low)
- [ ] Profile and optimize bulk create/update paths (batch sizes, DB-side upserts where supported). (low)

## Documentation & housekeeping
- [ ] Add a short developer note in README describing the pipeline flow and where to find legacy heuristics (`old_main.py`).
- [ ] Add a migration/developer note listing notable model additions (`StagedProcessingFailure`) and where to find failure records.

## Completed (keep for history)
- [x] Critical review of refactor changes and file-level summary (see `.github/critical-review.md`) — completed 2025-09-02
- [x] Harden staged-processing in `easy_access/db/update.py`: replaced `__dict__` usage with explicit mapping, batched transactional processing, per-row error handling, and conditional deletion of processed staged rows — completed 2025-09-02
- [x] Add per-row failure persistence: `StagedProcessingFailure` model and recording failures during staged processing — completed 2025-09-02
- [x] Implement small safe parsing helpers (`safe_int`, `safe_float`, `safe_date`, `safe_enum`, `safe_compare_greater`) and move them to `easy_access/utils.py` — completed 2025-09-02
- [x] Unit tests for parsing helpers: `tests/test_safe_parsers.py` — completed 2025-09-02
- [x] Integration test: `tests/test_integration_staging.py` verifying staging retention/deletion semantics — completed 2025-09-02

## Notes
- Work in the `new-dataflow` branch. Add changelog entries for completed items.

