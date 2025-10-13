# Easy Access Sheet Toolkit
*September 2025*

The Easy Access Sheet Toolkit is a comprehensive Python application with a built-in CLI designed to automate the processing, enrichment, and export of copyright data from university systems. It provides a complete pipeline for transforming raw copyright data into enriched, faculty-organized Excel sheets.

## Features

- Multi-stage Pipeline: Modular processing pipeline with independent stages (ingest, process, enrich, file-existence, export)
- Data Enrichment: Automatic enrichment with OSIRIS course and person data using TTL-based freshness
- File Existence Verification: TTL-based Canvas API file existence checking with rate limiting
- Bulk Operations: Optimized database operations for performance using SQLAlchemy ORM
- Export Generation: Multiple export formats (faculty sheets, overview, all items) with conditional formatting
- Backup & Restore: Automated backup of faculty sheets with configurable retention
- Admin Tools: Failure inspection, retry mechanisms, and cleanup utilities
- Modern Architecture: Async/await, dependency injection, comprehensive testing
- Workflow Mode: Optional reactive workflow exports (new/to_check/checked sheets)
- Database: PostgreSQL with asyncpg driver for production deployments

## Pipeline Stages

The toolkit's main processing function operates through several configurable pipeline stages:

### 1. Data Ingestion (`--ingest-only`)
- Reads raw copyright data from SURF CopyRight exports
- Processes data into standardized format using `copyright_item_from_dict`
- Handles duplicate detection and merging via `merge_rules.py`
- Stores processed data in PostgreSQL database using SQLAlchemy ORM

### 2. Data Processing (`--process-only`)
- Applies business rules and transformations
- Updates copyright item relationships and faculty mappings
- Performs data validation and cleanup
- Prepares data for enrichment with staged processing

### 3. Data Enrichment (`--enrich-only`)
- Fetches course data from OSIRIS API with concurrent requests
- Retrieves person/contact information from people pages
- Links courses to copyright items with bulk M2M operations
- TTL-based freshness policies (configurable in settings.yaml)

### 4. File Existence Check (`--file-exists-only`)
- Verifies file existence via Canvas API with rate limiting
- TTL-based checking to avoid unnecessary API calls
- Bulk database updates for efficiency
- Handles 404s and other API responses gracefully

### 5. Export Generation (`--export-only`)
- Creates faculty-specific Excel sheets with data entry and complete data sheets
- Generates overview and summary sheets
- Applies conditional formatting and dropdowns from settings.yaml
- Handles file uniqueness and versioning
- Optional workflow mode: new/to_check/checked sheets for reactive processing

## Quick Start

### Prerequisites
- Python 3.11+
- PostgreSQL 18+ with asyncpg driver
- [uv](https://docs.astral.sh/uv/) package manager (recommended)

### Database Setup

1. Install and start PostgreSQL:
   ```bash
   # Using Docker (recommended for development)
   docker run --name ea-postgres -e POSTGRES_PASSWORD=password -e POSTGRES_DB=ea_db -p 5432:5432 -d postgres:18
   ```

2. Create `.env` file in project root:
   ```
   DATABASE_URL=postgresql+asyncpg://postgres:password@localhost:5432/ea_db
   ```

3. Run database migrations:
   ```bash
   uv run alembic upgrade head
   ```

### Installation

1. Install uv (recommended):
   ```bash
   # Windows PowerShell
   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
   ```

2. Clone and setup:
   ```bash
   git clone https://github.com/utsmok/ea-cli.git
   cd ea-cli
   uv sync
   ```

### Configuration

1. Settings File: The main configuration is in `settings.yaml`. Key sections include:

   - Directories: Paths for input/output folders
   - University: Faculty hierarchy, LMS details, OSIRIS URLs
   - Data Settings: Column orders, new fields, dropdown options
   - Backup: Auto-backup settings and retention
   - Files: Script data and enrichment file paths

   Example configuration:
   ```yaml
   university:
     name: University of Twente
     abbreviation: UT
     lms:
       name: Canvas
       url: https://canvas.utwente.nl
     course_catalogue:
       name: OSIRIS
       query_url: https://utwente.osiris-student.nl/student/osiris/owc/cursussen/
     faculties:
       - name: Faculty of Behavioural, Management and Social Sciences
         abbreviation: BMS
         programmes: [...]

   data_settings:
     data_entry_cols:
       - name: "workflow_status"
         dropdown_options: '"ToDo,Done,InProgress"'
         default_val: "ToDo"
       - name: "manual_classification"
         dropdown_options: '"open access,eigen materiaal - powerpoint,..."'
     final_data_col_order: [material_id, is_duplicate, ...]
   ```

2. Add Copyright Data: Place SURF CopyRight exports in `raw_copyright_data/`

### Running the Pipeline

Full pipeline (default):
```bash
uv run run.py process
```

Individual stages:
```bash
# Only ingest new data
uv run run.py process --ingest-only

# Only process existing data
uv run run.py process --process-only

# Only enrich with external data
uv run run.py process --enrich-only

# Only check file existence
uv run run.py process --file-exists-only

# Only generate exports
uv run run.py process --export-only
```

Other commands:
```bash
# Run dashboard
uv run run.py dashboard --port 8000

# Export for single faculty
uv run run.py export --single-faculty BMS

# Create backup
uv run run.py backup create

# Restore backup
uv run run.py backup restore --restore-dir latest

# Inspect processing failures
uv run run.py admin inspect-failures

# Retry failed items
uv run run.py admin retry-failures --material-id 12345
```

## Architecture

### Core Components

- `pipeline.py`: Main orchestrator coordinating all stages with async entrypoints
- `db/`: Database models and operations
  - `sa_models.py`: SQLAlchemy ORM models (CopyrightItem, CourseData, PersonData, etc.)
  - `ingest.py`: Raw data ingestion and processing
  - `update.py`: Data processing and merging logic with staged failure handling
  - `relations.py`: M2M relationship management with bulk operations
  - `retrieve.py`: Optimized data retrieval with aggregation
- `enrichment/`: External data fetching
  - `osiris.py`: Course and person data APIs with concurrent fetching
- `maintenance/`: Ongoing data maintenance
  - `file_existence.py`: Canvas API file verification with TTL
- `sheets/`: Export generation
  - `export.py`: Excel sheet creation and formatting
  - `backup.py`: Backup and restore operations
  - `sheet.py`: Sheet utilities and conditional formatting
- `settings.py`: Configuration management with Settings dataclass
- `utils.py`: Safe parsing helpers and utilities

### Database Schema

Key entities:
- CopyrightItem: Core copyright data with filehash, scan dates
- CourseData: Course information from OSIRIS
- PersonData: Contact information
- Faculty: Organizational hierarchy
- StagedProcessingFailure: Failure tracking for retries
- StagedCopyrightItem/StagedFacultyUpdate: Staging tables

Database: PostgreSQL with SQLAlchemy 2.0 async ORM and Alembic migrations

### Performance Optimizations

- Bulk Operations: Raw SQL for efficient batch updates
- Memory Management: Streaming/chunked data processing
- Rate Limiting: Configurable delays for API calls
- Connection Pooling: Optimized SQLAlchemy async connections
- Async Processing: Concurrent API calls with semaphores
- TTL Caching: Avoid redundant API calls

## Configuration Options

See `settings.yaml` for full configuration options.
Some example snippets:

Directory Settings:

```yaml
directories:
  raw_copyright_data: raw_copyright_data
  faculties_dir: faculty_sheets
  all_items_dir: cip_sheets
  script_data: script_data
  pdf_downloads: pdf_downloads
```

University Settings:

```yaml
university:
  faculties:
    - name: Faculty of Behavioural, Management and Social Sciences
      abbreviation: BMS
      programmes:
        - abbreviation: B-COM
          name: Communication Science
```

Data Settings:

```yaml
data_settings:
  data_entry_cols:
    - name: "workflow_status"
      dropdown_options: '"ToDo,Done,InProgress"'
    - name: "manual_classification_v2"
      dropdown_options: 'ENUM:ClassificationV2'
  final_data_col_order: [material_id, is_duplicate, ...]
```

## Development

### Project Structure
```
ea-cli/
├── easy_access/
│   ├── main.py              # module entry point
│   ├── pipeline.py          # Main pipeline orchestrator
│   ├── settings.py          # Configuration management
│   ├── utils.py             # Safe parsing helpers
│   ├── merge_rules.py       # Data merging logic
│   ├── read_data_and_update.py # Legacy data reading
│   ├── db/
│   │   ├── sa_models.py     # SQLAlchemy ORM models
│   │   ├── models.py        # Legacy Tortoise ORM models (deprecated)
│   │   ├── ingest.py        # Data ingestion
│   │   ├── update.py        # Data processing
│   │   ├── relations.py     # M2M relationships
│   │   └── retrieve.py      # Data retrieval
│   ├── enrichment/
│   │   └── osiris.py        # OSIRIS API client
│   ├── maintenance/
│   │   └── file_existence.py # File existence checks
│   └── sheets/
│       ├── export.py        # Export generation
│       ├── backup.py        # Backup operations
│       └── sheet.py         # Sheet utilities
├── tests/                   # Unit and integration tests
├── raw_copyright_data/      # Input data directory
├── cip_sheets/              # Generated faculty sheets
├── faculty_sheets/          # Faculty-specific exports
├── pdf_downloads/           # Downloaded PDF files
├── script_data/             # Intermediate data files
├── settings.yaml            # Configuration file
├── pyproject.toml           # Project dependencies
├── uv.lock                  # Lock file
└── run.py                   # CLI runner, main entry point
```

### Testing

Run the test suite:
```bash
uv run pytest
```

Run specific test categories:
```bash
# Unit tests only
uv run pytest tests/ -k "test_" -v

# Integration tests
uv run pytest tests/ -k "integration" -v

# With coverage
uv run pytest --cov=easy_access --cov-report=html
```

### Branch Information
- Current Branch: `main` (merged from `new-dataflow` after large refactor)
- Working Branch: Create feature branches from `main`
- Changelog: See `.github/changelog.md` for recent updates

### Known Issues & Under Development

- Type Hints: Complete coverage across all files (currently inconsistent)
- Documentation: Add comprehensive docstrings to public functions
- File Size: Decompose large files (>800 lines): `settings.py`, `models.py`, `sheet.py`

- Testing: Expand test coverage, especially integration tests
- Performance: Monitor and optimize for large datasets
- Error Handling: Standardize exception handling patterns
- Conditional Formatting: Make configurable per column via settings.yaml
- New Fields: Implement `filehash`, `last_scan_date_university`, `last_scan_date_course`
- Reactive Workflow: Implement new/to_check/checked sheet states
- Backup Integration: Add pre/post pipeline backup stages

- Docs: Update README with architecture diagram
- Maintenance: Add teardown leak detection for tests

#### Completed
- Pipeline refactor with async entrypoints
- Export read-only by default with `--disable-writes`
- Staged processing with failure persistence
- OSIRIS enrichment with TTL and concurrency
- File existence TTL-based checks
- Backup and restore functionality
- Admin tools for failure management

### Contributing

1. Branching: Create feature branches from `main`
2. Testing: Add tests for new functionality
3. Documentation: Update README and `.github/changelog.md`
4. Code Style: Follow existing patterns, add type hints


## License

See LICENSE file for details.
