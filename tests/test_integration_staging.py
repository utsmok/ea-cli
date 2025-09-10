import asyncio
import contextlib
from pathlib import Path

import pytest
from tortoise import Tortoise

from easy_access.db import base as db_base
from easy_access.db.models import (
    CopyrightItem,
    Faculty,
    StagedCopyrightItem,
)
from easy_access.db.update import process_staged_raw_data
from easy_access.settings import Settings


@pytest.mark.asyncio
async def test_process_staged_raw_data_respects_partial_failures():
    # Use a temporary sqlite file for stable multi-connection behavior
    import tempfile

    tf = tempfile.NamedTemporaryFile(delete=False)  # noqa: SIM115
    tf.close()
    db_path = Path(tf.name)

    settings = object.__new__(Settings)
    settings.db_path = db_path

    # Initialize Tortoise directly for the test to avoid interaction with module memoization
    print("[test] initializing Tortoise directly")
    await asyncio.wait_for(
        Tortoise.init(
            db_url=f"sqlite:///{db_path}", modules={"models": ["easy_access.db.models"]}
        ),
        timeout=10,
    )
    await asyncio.wait_for(Tortoise.generate_schemas(safe=True), timeout=10)
    # mark module-level init flag true so ensure_db_inited won't try to re-init
    with contextlib.suppress(Exception):
        db_base._DB_INITIALIZED = True
    print("[test] db initialized via Tortoise")

    try:
        # ensure no Faculty exists so canonical creation will initially fallback
        await Faculty.all().delete()

        print("[test] inserting staged rows")
        # insert two staged rows with required fields missing faculty but including required fields for creation
        await StagedCopyrightItem.create(
            material_id=101,
            period="2020-1A",
            department="DEPT",
            course_code="1001",
            course_name="Intro Test",
            classification="lange overname",
        )
        await StagedCopyrightItem.create(
            material_id=102,
            period="2020-1A",
            department="DEPT",
            course_code="1002",
            course_name="Intro Test 2",
            classification="lange overname",
        )

        print("[test] running first processing (should not clear rows)")
        await asyncio.wait_for(process_staged_raw_data(settings), timeout=10)

        remaining = await StagedCopyrightItem.all().values_list(
            "material_id", flat=True
        )
        assert set(remaining) == {101, 102}

        print("[test] creating fallback Faculty UNM")
        await Faculty.create(
            name="Unmapped",
            abbreviation="UNM",
            full_abbreviation="UNM",
            hierarchy_level=0,
        )

        print("[test] running second processing (should clear rows)")
        await asyncio.wait_for(process_staged_raw_data(settings), timeout=10)

        remaining_after = await StagedCopyrightItem.all().values_list(
            "material_id", flat=True
        )
        assert list(remaining_after) == []

        # also ensure copyright items were created
        ci101 = await CopyrightItem.get_or_none(material_id=101)
        ci102 = await CopyrightItem.get_or_none(material_id=102)
        assert ci101 is not None or ci102 is not None
    finally:
        # cleanup connections and remove temp file; always run
        with contextlib.suppress(Exception):
            await Tortoise.close_connections()
        with contextlib.suppress(Exception):
            db_path.unlink()
