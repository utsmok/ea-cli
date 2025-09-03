# Export, Enrichment & File Existence Integration Plan

Date: 2025-09-02

This document analyses the missing functional areas (Excel exports, OSIRIS & course enrichment, file existence verification) after the dataflow refactor and defines a concrete, testable reintegration plan aligned with the new staging -> processing architecture.

## 1. Current State (Post‑Refactor) - Updated Analysis
Pipeline stages implemented:
- ingest_raw_data_async -> loads raw export into staging
- ingest_faculty_updates_async -> loads data-entry edits into staging
- process_data_async -> merges staged rows into main tables via refactored merge logic

**Critical Findings from Code Review**:
- **Relations functions exist** but need extraction: `link_courses_to_copyright_items()` and `update_duplicate_status()` are in `db/update.py` with N+1 query patterns (await per item in loop).
- **Export functions completely missing**: Legacy `old_main.py` references `create_export_sheet` but function doesn't exist in current codebase; all export orchestration removed.
- **Pipeline export stage stubbed**: `pipeline.py` shows `# await self.export_reports_async() # To be implemented`.
- **Sheet formatting mature**: `DataEntrySheet` class fully functional with dropdown validation, width calculation, table styling.
- **Legacy export workflow complex**: `old_main.py` shows sophisticated sheet types (faculty, program, overview, all_items) with file uniqueness handling.

**✅ PHASE A COMPLETED**: Successfully implemented export stage with `sheets/export.py`, optimized relations stage with `db/relations.py`, integrated both into pipeline with proper database connection management. Manual testing confirmed Excel file generation with data entry sheets and proper file uniqueness handling. Fixed RuntimeError "This event loop is already running" by converting `create_faculty_overviews` to async and adding sync wrapper for backward compatibility.

**✅ PHASE B COMPLETED**: Successfully implemented complete OSIRIS enrichment system with concurrent HTTP requests, TTL-based freshness policies, bulk database persistence, pipeline integration with CLI flags, and robust error handling. All core functionality working, unit tests marked as future enhancement.

**✅ PHASE C COMPLETED**: Successfully implemented file existence verification with TTL-based freshness policies, concurrent processing, database integration, pipeline integration, and CLI flags. Tested CLI functionality and confirmed no errors.

Not yet implemented / incomplete:
- **Phase D: Performance & Documentation** - Optimize performance, improve persistence, add comprehensive tests, and update documentation.

## 2. Design Principles
1. Deterministic, idempotent stages: Each stage can be re-run safely.
2. Separation of concerns: Retrieval, transformation, persistence, and export are distinct modules.
3. Observability: Structured logging (counts, durations, deltas), minimal broad exception masking.
4. Incremental where possible: Re-check only missing/expired file existence or OSIRIS data.
5. Settings-driven: Concurrency limits, TTLs, enable/disable flags.
6. Testability: Unit + integration tests with small fixtures, no reliance on external network for core logic (mock httpx).

## 3. Proposed New/Updated Modules
| Area | Module | Responsibility |
|------|--------|---------------|
| Export Orchestration | `easy_access/sheets/export.py` | High-level export orchestrator (faculties, programmes, all items) invoking existing sheet helpers. |
| OSIRIS Enrichment | `easy_access/enrichment/osiris.py` (new subpkg) | Split retrieval (courses), person enrichment, merge, persistence, TTL logic. |
| Relations Update | `easy_access/db/relations.py` | Duplicate detection, course linking, future: staff/course person linkage. |
| File Existence | `easy_access/maintenance/file_existence.py` | Incremental file existence refresh + TTL policy wrapper around existing core checker. |
| Pipeline | `easy_access/pipeline.py` | Add async stages: `enrich_async()`, `update_relations_async()`, `refresh_file_existence_async()`, `export_reports_async()` with sync wrappers. |

## 4. Export Stage Detailed Plan
### Inputs
- DB (post-processing) via retrieval functions (`retrieve_full_data`, or a new optimized retrieval returning per-faculty partitions).

### Outputs
- Per faculty overview sheet (Complete Data + Data Entry sheet) – file naming: `{FACULTY}_total_overview_updated_{DATE}.xlsx`.
- Per programme sheets under `faculty/per_programme/` (grouping by course_mapping).
- All items sheet (optionally only new items since last run – later improvement).

### Implementation Steps
1. Create `sheets/export.py` with functions:
   - `gather_faculty_data(settings) -> dict[str, pl.DataFrame]` (DB fetch + fac filtering).
   - `export_faculty_overviews(settings, faculty_data)` (leverages existing `create_faculty_overviews`).
   - `export_all_items(settings)`.
2. Refactor duplicated logic now in legacy `old_main.py` but ensure use of standardized functions in `sheets/analysis.py` & `sheets/sheet.py` (minimal changes – adapt to DB‑first flow).
3. Add `export_reports_async()` to pipeline calling the above.
4. Add CLI flag(s) (already partial: `--export-only`) to include new stage; adjust run sequence ordering: enrichment -> relations -> file existence -> export.
5. Tests:
   - Unit: ensure file naming uniqueness helper works; ensure exported DataFrames have required columns.
   - Integration: run minimal pipeline on sample dataset; assert files created under temp directory; open a sheet and validate presence of data entry sheet with dropdown columns.

### Edge Cases
- Faculty with zero rows: skip and log.
- Programme course mapping referencing department with no rows: skip (current logic already covers this).
- Re-run same day: unique file naming adds suffix `_1`, `_2`, etc.

# Export, Enrichment & File Existence Integration Plan

Date: 2025-09-03

This concise plan documents the current state of the export/enrichment/file-existence reintegration, the design principles, and the minimal, testable implementation plan for the remaining work.

## Current state
- Ingest and processing pipeline stages are implemented (staging -> processing -> main tables).
- Phase A (export + relations) completed and manually validated.
- Phase B (OSIRIS enrichment) implemented; unit tests for parsing/staleness selection are still outstanding.
- Phase C (file-existence checking) implemented.
- Phase D (performance, docs, comprehensive tests) is remaining.

## Design principles
- Deterministic, idempotent stages.
- Clear separation of retrieval, transformation, persistence, and export.
- Settings-driven concurrency and TTLs.
- Observability: counts, durations, deltas.

## Modules (implemented or planned)
- `easy_access/sheets/export.py` — export orchestrator (implemented).
- `easy_access/enrichment/` — OSIRIS/person enrichment (implemented; tests pending).
- `easy_access/db/relations.py` — optimized relations (implemented).
- `easy_access/maintenance/file_existence.py` — file existence TTL checking (implemented).
- `easy_access/pipeline.py` — pipeline stages: ingest, process, enrich, relations, file-existence, export (implemented with async entrypoints and CLI flags).

## Remaining work (high priority)
1. Add unit tests for enrichment functions (HTML parsing, stale-selection logic).
2. Add unit tests for relations functions (batch linking / N+1 elimination).
3. Add integration tests for end-to-end pipeline runs in a temp dir (export file assertions).
4. Performance tuning (bulk M2M linking, export memory usage) — Phase D.

## Export notes
- Export files follow the existing, validated patterns: per-faculty overviews, per-programme sheets, and an all-items sheet. File-uniqueness and data-entry sheet formatting implemented.

## Enrichment notes
- Enrichment is DB-centric: fetch missing/stale Course and Person records, persist directly to models, then run relations linking.
- Settings keys used: `enrichment.course_ttl_days`, `enrichment.person_ttl_days`.

## File existence notes
- Uses `file_exists` and `last_canvas_check` fields; re-check when NULL or older than TTL.

## Pipeline ordering (final)
1. ingest_raw_data_async
2. ingest_faculty_updates_async
3. process_data_async
4. enrich_async (optional via flag)
5. update_relations_async
6. refresh_file_existence_async
7. export_reports_async

## Tests and verification
- Priorities: add unit tests for enrichment and relations, then integration tests for exports. Use mocked HTTP for enrichment and temp directories for export verification.

---
Status: concise plan saved. Phase D remains for performance and documentation work.
2. **Architecture Diagrams**: Data flow, module relationships
3. **API Documentation**: Function signatures, usage examples
4. **Deployment Guide**: Setup, configuration, troubleshooting

### Implementation Steps
1. Add missing unit tests for all modules
2. Implement performance optimizations
3. Add comprehensive integration tests
4. Update documentation
5. Add monitoring and benchmarking

### Acceptance Criteria
- All critical functions have unit test coverage
- Performance benchmarks meet requirements
- Documentation is complete and accurate
- Data persistence is reliable and atomic
- Integration tests pass consistently

## 11. Excel Sheet Creation & Formatting Analysis / Refactor Targets
### Observations (from `sheets/sheet.py` & legacy `old_main.py`)
1. Column width & word-wrap logic executed per column; width heuristics okay but can be isolated into helper and unit tested.
2. DataEntrySheet currently writes cell-by-cell (acceptable for moderate size, but may be slow for very large sets). Potential optimization: build a 2D list and use `worksheet.append` in a loop or `write_only` mode (openpyxl) if performance becomes an issue.
3. Repeated style creation risk: `NamedStyle` reuse should check existence before add (openpyxl will raise on duplicate names). Current code implicitly reuses name but not guarded.
4. Dropdown validations added individually; fine, but we could consolidate by building all DV ranges then adding once.
5. No atomic write: writing directly risks partial file on crash.
6. Lack of structural validation before write (e.g., ensuring required columns present) – this is implicit; clearer pre-flight check desirable.
7. Partial code sections in current file (placeholders / blanks) suggest refactor interrupted – need restoration.

### Refactor Plan
Introduce utility functions:
`build_data_entry_matrix(data: pl.DataFrame, col_infos: list[ColInfo]) -> list[list[Any]]`
`apply_column_dimensions(ws, col_infos)`
`create_or_reuse_style(wb, style_name, **attrs)`
`add_dropdown_validations(ws, col_infos, max_row)`

Add `atomic_save_workbook(wb, target_path)` writing to `target_path.tmp` then renaming.

Add validation helper: `validate_export_dataframe(df, required_cols: set[str])` raising early with counts of missing columns.

### Testing
- Unit: width calculation logic & dropdown list injection.
- Integration: generate a small sheet; reopen with openpyxl; assert presence of data entry sheet, table style, dropdown DV count.

## 12. Updated TODO Items
See updated `.github/todo.md` for revised, DB‑centric tasks and sheet refactor items.

## 13. Acceptance Criteria Summary
- Export stage produces correct set of files for sample dataset & passes structural validation tests.
- Relations stage logs added/linked counts deterministically; rerun yields zero-delta metrics.
- Enrichment stage stores JSON caches, skips already fresh entries, and exposes enrichment counts.
- File existence stage processes only selected items & persists results correctly.

---
Prepared to guide implementation phases; update this document as phases complete.
