"""
Enrichment module for fetching and persisting OSIRIS course and person data.

This module provides DB-centric enrichment functionality that:
- Fetches missing or stale course/person data from OSIRIS
- Uses TTL-based freshness policies
- Stores data directly in the database
- Integrates with the pipeline for automated enrichment
"""
