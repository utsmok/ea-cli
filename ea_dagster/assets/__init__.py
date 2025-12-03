"""
Dagster asset definitions for the Easy Access pipeline.

This package contains:
- ingestion: dlt-based raw data loading from Excel files
- processing: Merge logic from staging to main tables
- enrichment: OSIRIS data enrichment wrapper
- export: Excel report generation wrapper
"""

from ea_dagster.assets.enrichment import file_existence_check, osiris_enrichment
from ea_dagster.assets.export import excel_reports
from ea_dagster.assets.ingestion import faculty_updates_data, raw_copyright_data
from ea_dagster.assets.processing import processed_copyright_items

__all__ = [
    "raw_copyright_data",
    "faculty_updates_data",
    "processed_copyright_items",
    "osiris_enrichment",
    "file_existence_check",
    "excel_reports",
]
