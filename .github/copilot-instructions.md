# GitHub Copilot Instructions for EA-CLI

This repository contains the Easy Access Sheet Toolkit, a Python application for automating copyright data processing and enrichment.

## Project Overview

The Easy Access Sheet Toolkit processes copyright data from university systems through a multi-stage pipeline: ingestion → processing → enrichment → file existence checking → export generation. The application uses async/await patterns, Tortoise ORM with SQLite, and generates Excel exports with conditional formatting.

## Key Technologies

- **Language**: Python 3.12.2 (pinned)
- **Package Manager**: uv
- **Database**: SQLite with Tortoise ORM
- **Async Runtime**: asyncio
- **CLI Framework**: Typer
- **Data Processing**: polars, pandas
- **Excel Generation**: openpyxl, xlsxwriter
- **HTTP Client**: httpx
- **Logging**: loguru

## Code Style & Conventions

### Formatting & Linting
- Use **ruff** with 88-character line length
- Target Python 3.12
- Run `uv run ruff check` before committing
- Configuration in `pyproject.toml`

### Type Hints
- Add type hints to all public functions and methods
- Many dataclasses are already typed
- Prefer readable, explicit types
- Use `from __future__ import annotations` for forward references

### Async Patterns
- All I/O operations and pipeline stages are async
- Use `run_sync` wrapper (from `easy_access.utils`) for sync compatibility
- Never use nested `asyncio.run()` - use `run_sync` instead
- Example:
  ```python
  # In utils.py
  def run_sync(coro):
      """Execute async coroutine in sync context"""
      # Implementation handles event loop detection
  ```

### Logging
- Use `loguru` for all logging
- Logger is configured via `configure_logger()` in `easy_access.settings`
- Do not use `print()` statements; use `logger.info()`, `logger.debug()`, etc.

### Error Handling
- Use specific exception types when possible
- Add context to error messages
- Use `safe_*` parsing helpers from `easy_access.utils` for robust data parsing:
  - `safe_int()`, `safe_float()`, `safe_date()`, `safe_enum()`, `safe_compare_greater()`

## Project Structure

```
ea-cli/
├── easy_access/              # Main package
│   ├── main.py              # Module entry point
│   ├── pipeline.py          # Pipeline orchestrator (async stages + sync wrappers)
│   ├── settings.py          # Configuration dataclasses and YAML parsing
│   ├── utils.py             # Safe parsers, run_sync, Directory/File helpers
│   ├── merge_rules.py       # Data merging logic
│   ├── db/                  # Database layer
│   │   ├── models.py        # Tortoise ORM models
│   │   ├── ingest.py        # Data ingestion
│   │   ├── update.py        # Data processing and merging
│   │   ├── relations.py     # M2M relationship management
│   │   └── retrieve.py      # Data retrieval
│   ├── enrichment/          # External data fetching
│   │   └── osiris.py        # OSIRIS API client (course/person data)
│   ├── maintenance/         # Ongoing tasks
│   │   └── file_existence.py # Canvas file verification
│   ├── sheets/              # Excel export generation
│   │   ├── export.py        # Main export orchestration
│   │   ├── backup.py        # Backup operations
│   │   └── sheet.py         # Sheet utilities and formatting
│   ├── pdf/                 # PDF handling
│   └── classification/      # NER and classification helpers
├── tests/                   # Test suite (pytest + pytest-asyncio)
├── run.py                   # CLI entry point (Typer)
├── settings.yaml            # Main configuration file
└── pyproject.toml           # Dependencies and tool config
```

## Key Components

### Configuration (`settings.py`)
- `Settings` and `EasyAccessSettings` dataclasses parse `settings.yaml`
- Supports per-faculty and per-run overrides via `OverrideSettings`
- Secret discovery order: environment variables → `.env`/`.secret` files → `api_keys.py`
- Use `SETTINGS` global for configuration access

### Database (`db/models.py`)
- Core models: `CopyrightItem`, `CourseData`, `PersonData`, `Faculty`, `Programme`
- Staging tables: `StagedCopyrightItem`, `StagedFacultyUpdate`, `StagedProcessingFailure`
- Use bulk operations (`bulk_create`, `bulk_update`) for performance
- Use `TimestampMixin` for created_at/modified_at tracking

### Pipeline (`pipeline.py`)
- Implements async stages: ingest → process → enrich → file_exists → export
- Each stage has async implementation + sync wrapper
- Use `run_sync` for CLI consumers
- Error handling: critical errors in ingest/process abort; other stages skip on failure

### Utilities (`utils.py`)
- `run_sync(coro)`: Execute async code in sync context
- `Directory` / `File` classes: Path handling with convenience methods
- `standardize_dataframe(df)`: Normalize DataFrames before ingestion
- `determine_course_code()`: Extract numeric course codes from strings
- Safe parsers: Handle heterogeneous inputs from Excel/APIs

## Common Patterns

### Implementing New Features

1. **Add Async Logic**: Implement core logic as async functions
2. **Add Sync Wrapper**: Use `run_sync` for CLI/sync compatibility
3. **Update Settings**: Add configuration to `Settings` dataclass if needed
4. **Add Tests**: Write focused pytest tests (see `tests/`)
5. **Update Documentation**: Add docstrings and update README if needed

### Database Operations

```python
# Bulk operations for performance
from easy_access.db.models import CopyrightItem

# Bulk create with conflict handling
await CopyrightItem.bulk_create(
    items,
    on_conflict=["material_id"],
    update_fields=["field1", "field2"]
)

# Bulk update
await CopyrightItem.filter(id__in=ids).update(status="processed")

# Use prefetch_related for M2M to avoid N+1
items = await CopyrightItem.all().prefetch_related("courses", "faculty")
```

### File Operations

```python
from easy_access.utils import Directory, File

# Directory operations
dir = Directory("/path/to/dir")
dir.create()  # Creates if not exists
files = dir.files()  # List files
newest = dir.newest_file("*.xlsx")  # Get newest matching file

# File operations
file = File("/path/to/file.txt")
file.copy(dest)
file.move(dest)
```

### DataFrame Processing

```python
import polars as pl
from easy_access.utils import standardize_dataframe

# Read and standardize
df = pl.read_excel("file.xlsx")
df = standardize_dataframe(df)  # Normalizes columns, handles nulls

# Use polars for large data transformations
# Use pandas only when necessary for openpyxl compatibility
```

## Development Workflow

### Setup
```bash
# Clone and install
git clone https://github.com/utsmok/ea-cli.git
cd ea-cli
uv sync
```

### Running
```bash
# Full pipeline
uv run run.py process

# Individual stages
uv run run.py process --ingest-only
uv run run.py process --enrich-only
uv run run.py process --export-only

# Single faculty export
uv run run.py export --single-faculty BMS

# Dashboard
uv run run.py dashboard --port 8000
```

### Testing
```bash
# Run all tests
uv run pytest

# Run specific tests
uv run pytest tests/test_utils.py -v

# With coverage
uv run pytest --cov=easy_access --cov-report=html
```

### Linting
```bash
# Check code
uv run ruff check

# Fix auto-fixable issues
uv run ruff check --fix

# Format code
uv run ruff format
```

## Important Constraints

### Do NOT
- Remove or modify working code unnecessarily
- Use `print()` for logging (use `logger` instead)
- Create nested `asyncio.run()` calls (use `run_sync`)
- Add dependencies without checking existing alternatives
- Ignore type hints in new code
- Write files >800 lines (decompose into smaller modules)

### DO
- Use bulk database operations for performance
- Add type hints to new functions
- Write docstrings for public APIs
- Use `safe_*` parsers for external data
- Preserve existing merge rules in `merge_rules.py`
- Follow async-first design with sync wrappers
- Use `Directory`/`File` wrappers for filesystem operations
- Add focused unit tests for new features

## Known Issues & Priorities

### High Priority
- Complete type hint coverage across all files
- Add comprehensive docstrings to public functions
- Remove hardcoded credentials from `api_keys.py` (use env vars)

### Medium Priority
- Decompose large files: `settings.py`, `models.py`, `sheet.py` (>800 lines)
- Standardize error handling patterns
- Expand test coverage (especially integration tests)
- Remove production/test coupling in `db/relations.py`

### Under Development
- PostgreSQL migration (currently using SQLite)
- Reactive workflow mode (new/to_check/checked sheets)
- File hash and scan date tracking
- Performance monitoring for large datasets

## Secrets Management

**Order of precedence** (highest to lowest):
1. Environment variables
2. `.env` or `.secret` files (searched up to 2 parent directories)
3. `api_keys.py` module (repo root or `easy_access/api_keys.py`)

**For CI/CD**: Always use environment variables
**For local dev**: Use `.env` file or environment variables (do not commit secrets)

## Additional Resources

- **Main branch**: `main` (merged from `new-dataflow` after large refactor)
- **Documentation**: See `.agents/` for detailed architecture notes
- **Changelog**: See `.github/old/changelog.md` for recent updates
- **Legacy docs**: See `.github/old/` for historical planning documents

## Migration Notes

**PostgreSQL Migration** (planned): The project is planning to migrate from SQLite + Tortoise ORM to PostgreSQL + SQLAlchemy. See `.agents/migration_to_postgres_sqlalchemy.md` for the detailed migration plan. When working on database-related features, keep this future migration in mind.

---

*Last updated: 2025-10-31*
