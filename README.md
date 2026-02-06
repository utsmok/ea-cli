# Easy Access Sheet Toolkit
*February 2026*

A Python CLI tool for automated copyright data processing, enrichment, and Excel export generation. Transforms raw SURF CopyRight exports into enriched, faculty-organized spreadsheets.

## Features

- **Multi-stage Pipeline**: Ingest → Process → Enrich → Verify → Export
- **Data Enrichment**: OSIRIS course/person data with TTL-based freshness
- **File Verification**: Canvas API file existence checks with rate limiting
- **Bulk Operations**: Optimized database operations using Tortoise ORM
- **Export Formats**: Faculty sheets, overviews, and all-items exports
- **Backup System**: Automated backup with configurable retention
- **Admin Tools**: Failure inspection, retry mechanisms, and cleanup utilities

## Quick Start

### Prerequisites
- Python 3.12+
- [uv](https://docs.astral.sh/uv/) package manager

### Installation

```bash
# Clone repository
git clone https://github.com/utsmok/ea-cli.git
cd ea-cli

# Install dependencies
uv sync

# Configure settings.yaml (see Configuration section below)
# Place SURF CopyRight exports in raw_copyright_data/
```

### Running the Pipeline

**Full pipeline (all stages):**
```bash
uv run run.py process
```

**Individual stages:**
```bash
uv run run.py process --ingest-only          # Ingest raw data
uv run run.py process --process-only         # Process and validate
uv run run.py process --enrich-only          # Enrich with OSIRIS data
uv run run.py process --file-exists-only     # Verify file existence
uv run run.py process --export-only          # Generate exports
```

**Other commands:**
```bash
# Export for single faculty
uv run run.py export --single-faculty BMS

# Backup management
uv run run.py backup create
uv run run.py backup restore --restore-dir latest

# Admin tools
uv run run.py admin inspect-failures
uv run run.py admin retry-failures --material-id 12345
```

## Configuration

Main configuration is in `settings.yaml`. Key sections:

### Directories
```yaml
directories:
  raw_copyright_data: raw_copyright_data    # Input from SURF
  faculties_dir: faculty_sheets             # Faculty-specific exports
  script_data: script_data                  # Intermediate files
  pdf_downloads: pdf_downloads              # Downloaded PDFs
```

### University Settings
```yaml
university:
  name: University of Twente
  abbreviation: UT
  lms:
    name: Canvas
    url: https://canvas.utwente.nl
  faculties:
    - name: Faculty of Behavioural, Management and Social Sciences
      abbreviation: BMS
```

### Data Settings
```yaml
data_settings:
  data_entry_cols:
    - name: "workflow_status"
      dropdown_options: '"ToDo,Done,InProgress"'
    - name: "v2_manual_classification"
      dropdown_options: 'ENUM:ClassificationV2'
```

## Architecture

### Core Components
```
easy_access/
├── pipeline.py          # Main orchestrator
├── db/                  # Database layer (models, ingest, update, retrieve)
├── enrichment/          # OSIRIS API client
├── maintenance/         # File existence checks
├── sheets/              # Export generation & backup
├── classification/      # ML classification
├── pdf/                 # PDF handling
├── settings.py          # Configuration
└── utils.py             # Utilities
```

### Database Schema
- **CopyrightItem**: Core copyright data with filehash and scan dates
- **CourseData**: Course information from OSIRIS
- **PersonData**: Contact information
- **Faculty**: Organizational hierarchy
- **Staging tables**: For batch processing and failure tracking

## Development

### Testing
```bash
# Run all tests
uv run pytest

# With coverage
uv run pytest --cov=easy_access --cov-report=html
```

### Project Structure
```
ea-cli/
├── easy_access/              # Main package
├── tests/                    # Test suite
├── raw_copyright_data/       # Input data
├── faculty_sheets/           # Generated exports
├── settings.yaml             # Configuration
├── pyproject.toml            # Dependencies
└── run.py                    # CLI entry point
```

## Contributing

1. Create feature branches from `main`
2. Add tests for new functionality
3. Follow existing code patterns and add type hints
4. Update `settings.yaml` if adding configuration options

## License

See LICENSE file for details.
