import pytest

from easy_access.db.models import (
    WorkflowStatus,
)
from easy_access.db.update import compare_and_update_fields
from easy_access.merge_rules import build_merge_rules_from_settings
from easy_access.settings import Settings


@pytest.mark.asyncio
async def test_workflow_status_never_downgrades(monkeypatch):
    # Minimal settings mock
    settings = Settings()
    added, changeable = build_merge_rules_from_settings(settings)

    # Force a misordered workflow list to simulate spreadsheet input
    added["workflow_status"] = [
        WorkflowStatus.ToDo.value,
        WorkflowStatus.InProgress.value,
        WorkflowStatus.Done.value,
    ]

    # Create mock db item
    class DummyItem:  # runtime stub mimicking ORM field access
        def __init__(self):
            self.workflow_status = WorkflowStatus.Done.value

    db_item = DummyItem()  # type: ignore
    new_item = {"material_id": 1, "workflow_status": WorkflowStatus.ToDo.value}

    changes = {}
    changes, _ = compare_and_update_fields(
        new_item, db_item, {"workflow_status": added["workflow_status"]}, changes
    )

    # Should not contain workflow_status change because downgrade blocked
    assert "workflow_status" not in changes


@pytest.mark.asyncio
async def test_workflow_status_upgrade(monkeypatch):
    settings = Settings()
    added, _ = build_merge_rules_from_settings(settings)
    # Misordered list still should allow upgrade from ToDo -> Done
    added["workflow_status"] = [
        WorkflowStatus.ToDo.value,
        WorkflowStatus.InProgress.value,
        WorkflowStatus.Done.value,
    ]

    class DummyItem:
        def __init__(self):
            self.workflow_status = WorkflowStatus.ToDo.value

    db_item = DummyItem()  # type: ignore
    new_item = {"material_id": 2, "workflow_status": WorkflowStatus.Done.value}
    changes = {}
    changes, _ = compare_and_update_fields(
        new_item, db_item, {"workflow_status": added["workflow_status"]}, changes
    )
    assert "workflow_status" in changes
