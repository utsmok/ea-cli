async def load_faculty_updates_to_staging(settings: Settings, data: pl.DataFrame) -> None:
    """
    Loads faculty updates into the staging table.
    """
    await ensure_db_inited(settings)
    data = data.select(["material_id", "manual_classification", "remarks", "workflow_status"])
    items = standardize_dataframe(data).to_dicts()
    staged_items = [StagedFacultyUpdate(**item) for item in items]
    await StagedFacultyUpdate.bulk_create(staged_items, on_conflict=["material_id"], update_fields=["manual_classification", "remarks", "workflow_status"])