# Easy Access Sheet Toolkit
*March 2025*

The Easy Access Sheet Toolkit is a comprehensive Python application with a built-in CLI designed to automate the processing, enrichment, and export of copyright data from university systems. It provides a complete pipeline for transforming raw copyright data into enriched, faculty-organized Excel sheets.

## Features

- **Multi-stage Pipeline**: Modular processing pipeline with independent stages
- **Data Enrichment**: Automatic enrichment with OSIRIS course and person data
- **File Existence Verification**: TTL-based Canvas API file existence checking
- **Bulk Operations**: Optimized database operations for performance
- **Export Generation**: Multiple export formats (faculty sheets, overview, all items)
- **Modern Architecture**: Async/await, dependency injection, comprehensive testing

## Pipeline Stages

The toolkit operates through several configurable pipeline stages:

### 1. Data Ingestion (`--ingest-only`)
- Reads raw copyright data from SURF CopyRight exports
- Processes data into standardized format
- Handles duplicate detection and merging
- Stores processed data in database

### 2. Data Processing (`--process-only`)
- Applies business rules and transformations
- Updates copyright item relationships
- Performs data validation and cleanup
- Prepares data for enrichment

### 3. Data Enrichment (`--enrich-only`)
- Fetches course data from OSIRIS API
- Retrieves person/contact information
- Links courses to copyright items
- TTL-based freshness policies

### 4. File Existence Check (`--file-exists-only`)
- Verifies file existence via Canvas API
- TTL-based checking (configurable)
- Rate-limited API calls
- Bulk database updates

### 5. Export Generation (`--export-only`)
- Creates faculty-specific Excel sheets
- Generates overview and summary sheets
- Applies formatting and styling
- Handles file uniqueness and versioning

## Quick Start

### Prerequisites
- Python 3.11+
- [uv](https://docs.astral.sh/uv/) package manager (recommended)

### Installation

1. **Install uv** (recommended):
   ```bash
   # Windows PowerShell
   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
   uv python install
   ```

2. **Clone and setup**:
   ```bash
   git clone <repository-url>
   cd ea-cli
   uv sync
   ```

### Configuration

1. **Settings File**: Copy and modify `settings.yaml`:
   ```yaml
   university_settings:
     canvas_api_token: "your_canvas_token"
     osiris_base_url: "https://osiris.utwente.nl"

   enrichment_settings:
     course_ttl_days: 30
     person_ttl_days: 30
     file_exists_ttl_days: 30
     file_exists_rate_limit_delay: 0.1

   data_settings:
     raw_data_col_order: [...]
   ```

2. **Add Copyright Data**: Place SURF CopyRight exports in `raw_copyright_data/`

### Running the Pipeline

**Full pipeline** (default):
```bash
uv run run.py
```

**Individual stages**:
```bash
# Only ingest new data
uv run run.py --ingest-only

# Only process existing data
uv run run.py --process-only

# Only enrich with external data
uv run run.py --enrich-only

# Only check file existence
uv run run.py --file-exists-only

# Only generate exports
uv run run.py --export-only
```

**Skip stages**:
```bash
# Skip file existence checks
uv run run.py --no-file-exists

# Skip enrichment
uv run run.py --no-enrich
```

## Architecture

### Core Components

- **`pipeline.py`**: Main orchestrator coordinating all stages
- **`db/`**: Database models and operations
  - `models.py`: Tortoise ORM models
  - `update.py`: Data processing and merging logic
  - `relations.py`: M2M relationship management
  - `retrieve.py`: Optimized data retrieval with aggregation
- **`enrichment/`**: External data fetching
  - `osiris.py`: Course and person data APIs
- **`maintenance/`**: Ongoing data maintenance
  - `file_existence.py`: Canvas API file verification
- **`sheets/`**: Export generation
  - `export.py`: Excel sheet creation and formatting

### Database Schema

Key entities:
- **CopyrightItem**: Core copyright data
- **CourseData**: Course information from OSIRIS
- **PersonData**: Contact information
- **Faculty**: Organizational hierarchy
- **OrganizationData**: Department affiliations

### Performance Optimizations

- **Bulk Operations**: Raw SQL for efficient batch updates
- **Memory Management**: Streaming/chunked data processing
- **Rate Limiting**: Configurable delays for API calls
- **Connection Pooling**: Optimized database connections
- **Async Processing**: Concurrent API calls with semaphores

## Configuration Options

### CLI Flags

| Flag | Description |
|------|-------------|
| `--ingest-only` | Run only data ingestion stage |
| `--process-only` | Run only data processing stage |
| `--enrich-only` | Run only data enrichment stage |
| `--export-only` | Run only export generation stage |
| `--file-exists-only` | Run only file existence verification |
| `--no-enrich` | Skip data enrichment stage |
| `--no-file-exists` | Skip file existence verification |
| `--force` | Force reprocessing of all data |

### Settings Configuration

**Enrichment Settings** (`settings.yaml`):
```yaml
enrichment_settings:
  course_ttl_days: 30          # Days before course data is considered stale
  person_ttl_days: 30          # Days before person data is considered stale
  file_exists_ttl_days: 30     # Days before file existence is rechecked
  file_exists_rate_limit_delay: 0.1  # Seconds between API calls
```

**Data Settings**:
```yaml
data_settings:
  raw_data_col_order: [...]     # Column ordering for exports
  faculty_hierarchy: {...}      # Faculty/program hierarchy
```

## Development

### Project Structure
```
ea-cli/
├── easy_access/
│   ├── db/                    # Database operations
│   ├── enrichment/           # External data fetching
│   ├── maintenance/          # Ongoing maintenance tasks
│   ├── sheets/              # Export generation
│   ├── settings.py          # Configuration management
│   ├── pipeline.py          # Main orchestration
│   └── main.py              # CLI entry point
├── tests/                   # Unit and integration tests
├── raw_copyright_data/      # Input data directory
├── cip_sheets/             # Generated faculty sheets
├── pdf_downloads/          # Downloaded PDF files
└── settings.yaml           # Configuration file
```

### Testing

Run the test suite:
```bash
uv run pytest
```

Run specific test categories:
```bash
# Unit tests only
uv run pytest tests/test_*.py -v

# Integration tests
uv run pytest tests/test_integration_*.py -v

# With coverage
uv run pytest --cov=easy_access --cov-report=html
```

### Contributing

1. **Branching**: Create feature branches from `new-dataflow`
2. **Testing**: Add tests for new functionality
3. **Documentation**: Update README for new features
4. **Code Style**: Follow existing patterns and add type hints

## Troubleshooting

### Common Issues

1. **API Token Missing**: Ensure `canvas_api_token` is set in `settings.yaml`
2. **Database Errors**: Check database file permissions and disk space
3. **Memory Issues**: Reduce batch sizes in settings for large datasets
4. **Rate Limiting**: Increase `file_exists_rate_limit_delay` if hitting API limits

### Logs

Check logs in the console output or enable debug logging:
```bash
uv run run.py --verbose
```

### Performance Tuning

For large datasets, adjust these settings:
```yaml
# Reduce memory usage
batch_size: 500

# Reduce API load
max_concurrent: 25
file_exists_rate_limit_delay: 0.2

# Database optimization
pool_size: 10
```

## License

See LICENSE file for details.
