"""
Main Dagster definitions entry point.

This module defines the Dagster Definitions object that includes:
- All pipeline assets (ingestion, processing, enrichment, export)
- Resource configurations (database, settings)

Usage:
    Run `dagster dev` from the project root to start the Dagster UI.
    The UI will be available at http://localhost:3000
"""

from dagster import Definitions, load_assets_from_modules

from ea_dagster.assets import enrichment, export, ingestion, processing
from ea_dagster.resources import SettingsResource, TortoiseDBResource

# Load all assets from asset modules
all_assets = load_assets_from_modules([ingestion, processing, enrichment, export])

# Define the Dagster application
defs = Definitions(
    assets=all_assets,
    resources={
        "db": TortoiseDBResource(),
        "settings": SettingsResource(),
    },
)
