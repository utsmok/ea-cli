"""
Processing assets for merging staged data into the main database.

This module replaces `pipeline.process_data_async` by:
1. Reading from dlt staging tables
2. Running merge logic via services
3. Updating the main CopyrightItem table using Tortoise ORM
"""

from dagster import (
    AssetExecutionContext,
    AssetIn,
    MetadataValue,
    Output,
    TableColumn,
    TableSchema,
    asset,
)

from ea_dagster.resources import TortoiseDBResource

# Table names created by dlt (follows naming convention: {dataset}__{table_name})
STAGING_RAW_COPYRIGHT_TABLE = "staging__raw_copyright_items"
STAGING_FACULTY_UPDATES_TABLE = "staging__faculty_updates"


@asset(
    group_name="processing",
    ins={
        "raw_data": AssetIn(key="raw_copyright_data"),
        "faculty_data": AssetIn(key="faculty_updates_data"),
    },
)
async def processed_copyright_items(
    context: AssetExecutionContext,
    db: TortoiseDBResource,
    raw_data: str,
    faculty_data: str,
) -> Output:
    """
    Reads from dlt staging tables, runs merge logic, and updates main CopyrightItem table.

    This asset:
    1. Reads raw data from the dlt-created staging table
    2. Splits data into new vs update items
    3. Bulk creates new items
    4. Calculates and applies changes for existing items
    5. Creates changelog entries for updates

    Args:
        context: Dagster execution context
        db: Tortoise database resource
        raw_data: Output from raw_copyright_data asset
        faculty_data: Output from faculty_updates_data asset

    Returns:
        str: Summary of processing results
    """
    async with db.yield_for_execution(context):
        from tortoise import Tortoise

        from easy_access.db.base import copyright_item_from_dict
        from easy_access.db.models import (
            CopyrightItem,
            ItemUpdate,
        )
        from easy_access.merge_rules import build_merge_rules_from_settings
        from easy_access.services.merge import calculate_changes
        from easy_access.settings import SETTINGS

        # 1. Read Raw Data from the dlt staging table
        conn = Tortoise.get_connection("default")

        try:
            rows = await conn.execute_query_dict(
                f"SELECT * FROM {STAGING_RAW_COPYRIGHT_TABLE}"
            )
            context.log.info(f"Fetched {len(rows)} rows from staging.")
        except Exception as e:
            context.log.warning(
                f"Could not read staging table (may not exist yet): {e}"
            )
            rows = []

        if not rows:
            context.log.info("No staged data to process")
            return Output(value="No staged data to process")

        # 2. Get existing items for comparison
        existing_items_list = await CopyrightItem.all()
        existing_ids = {item.material_id for item in existing_items_list}
        existing_map = {item.material_id: item for item in existing_items_list}

        context.log.info(f"Found {len(existing_ids)} existing items in database")

        # 3. Split into new vs update
        new_items_data = []
        update_items_data = []

        for row in rows:
            mat_id = row.get("material_id")
            if mat_id is None:
                continue

            try:
                mat_id = int(mat_id)
            except (ValueError, TypeError):
                context.log.warning(f"Invalid material_id: {mat_id}")
                continue

            if mat_id in existing_ids:
                update_items_data.append(row)
            else:
                new_items_data.append(row)

        context.log.info(
            f"Split: {len(new_items_data)} new items, {len(update_items_data)} updates"
        )

        # 4. Bulk Create New Items
        created_count = 0
        if new_items_data:
            new_items = []
            for item_dict in new_items_data:
                item = await copyright_item_from_dict(item_dict)
                if item:
                    new_items.append(item)

            if new_items:
                await CopyrightItem.bulk_create(new_items, ignore_conflicts=True)
                created_count = len(new_items)
                context.log.info(f"Created {created_count} new items")

        # 5. Process Updates (Business Logic)
        updated_count = 0
        changelog_entries = []

        # Get merge rules from settings (proper dict format with priority lists)
        added_fields, changeable_fields = build_merge_rules_from_settings(SETTINGS)

        for row in update_items_data:
            mat_id = int(row.get("material_id"))
            if mat_id not in existing_map:
                continue

            current_item = existing_map[mat_id]

            # Calculate changes using the merge service
            try:
                changes, _ = calculate_changes(
                    new_data=row,
                    current_item=current_item,
                    added_fields=added_fields,
                    changeable_fields=changeable_fields,
                )
            except Exception as e:
                context.log.warning(f"Error calculating changes for {mat_id}: {e}")
                continue

            if changes:
                # Apply changes to the item
                for field, change_info in changes.items():
                    if hasattr(current_item, field):
                        # changes dict contains {field: {old: X, new: Y}}
                        new_val = (
                            change_info.get("new")
                            if isinstance(change_info, dict)
                            else change_info
                        )
                        setattr(current_item, field, new_val)

                await current_item.save()
                updated_count += 1

                # Create changelog entry
                changelog_entries.append(
                    ItemUpdate(
                        material_id=mat_id,
                        field_name="merged_update",
                        old_value=str(changes),
                        new_value="applied",
                    )
                )

        # Bulk create changelog entries
        if changelog_entries:
            await ItemUpdate.bulk_create(changelog_entries, ignore_conflicts=True)
            context.log.info(f"Created {len(changelog_entries)} changelog entries")

        # 6. Process faculty updates
        faculty_updates_count = 0
        try:
            faculty_rows = await conn.execute_query_dict(
                f"SELECT * FROM {STAGING_FACULTY_UPDATES_TABLE}"
            )
            context.log.info(f"Fetched {len(faculty_rows)} faculty update rows")

            for row in faculty_rows:
                mat_id = row.get("material_id")
                if mat_id is None:
                    continue

                try:
                    mat_id = int(mat_id)
                except (ValueError, TypeError):
                    continue

                if mat_id not in existing_map:
                    continue

                item = existing_map[mat_id]
                updated = False

                # Apply faculty updates
                for field in ["manual_classification", "remarks", "workflow_status"]:
                    if field in row and row[field]:
                        old_val = getattr(item, field, None)
                        new_val = row[field]
                        if old_val != new_val:
                            setattr(item, field, new_val)
                            updated = True

                if updated:
                    await item.save()
                    faculty_updates_count += 1

        except Exception as e:
            context.log.warning(f"Could not process faculty updates: {e}")

        result_msg = (
            f"Processed: {created_count} new items created, "
            f"{updated_count} items updated, "
            f"{faculty_updates_count} faculty updates applied"
        )
        context.log.info(result_msg)

        # Metadata
        columns = [TableColumn(name=f) for f in CopyrightItem._meta.fields]
        schema = TableSchema(columns=columns)
        metadata = {
            "table_name": "copyrightitem",
            "table_schema": MetadataValue.table_schema(schema),
            "created_count": created_count,
            "updated_count": updated_count,
            "faculty_updates_count": faculty_updates_count,
        }
        return Output(value=result_msg, metadata=metadata)
