# from easy_access.classification.pdf_parser import deduplicate_pdfs
import asyncio

import polars as pl

from easy_access.classification.classifier_api import main
from easy_access.db.models import PDF, CopyrightItem


async def compare_pdfs():
    delete = False
    from easy_access.db.base import ensure_db_inited

    await ensure_db_inited()

    all_pdfs = (
        await PDF.all()
        .prefetch_related("replace_with")
        .values("material_id", "author", "title", "replace_with__material_id")
    )
    all_pdfs = pl.from_dicts(all_pdfs)
    pdfs_with_duplicates = all_pdfs.filter(
        pl.col("replace_with__material_id").is_not_null()
    ).to_dicts()
    print(f"Got {len(pdfs_with_duplicates)}/{len(all_pdfs)} pdfs are duplicates")

    for dupe in pdfs_with_duplicates:
        original = all_pdfs.filter(
            pl.col("material_id") == dupe.get("replace_with__material_id")
        ).to_dicts()[0]
        compare = {
            "pdf_material_id": (dupe.get("material_id"), original.get("material_id")),
            "pdf_author": (dupe.get("author"), original.get("author")),
            "pdf_title": (dupe.get("title"), original.get("title")),
        }
        dupe_copyright_item = await CopyrightItem.get_or_none(
            material_id=dupe.get("material_id")
        )

        original_copyright_item = await CopyrightItem.get_or_none(
            material_id=original.get("material_id")
        )
        if dupe_copyright_item and original_copyright_item:
            compare["ci_author"] = (
                dupe_copyright_item.author,
                original_copyright_item.author,
            )
            compare["ci_title"] = (
                dupe_copyright_item.title,
                original_copyright_item.title,
            )
            compare["ci_filename"] = (
                dupe_copyright_item.filename,
                original_copyright_item.filename,
            )
            compare["ci_course_name"] = (
                dupe_copyright_item.course_name,
                original_copyright_item.course_name,
            )
            compare["ci_man_class"] = (
                dupe_copyright_item.manual_classification,
                original_copyright_item.manual_classification,
            )

        print(f"Comparing {dupe.get('material_id')} with {original.get('material_id')}")

        not_same = False
        for k, v in compare.items():
            print(f"     {k}:\n          {v[0]}\n          {v[1]}")
            if k in ["ci_course", "pdf_title"]:
                if v[0] != v[1]:
                    not_same = True

        if not_same and delete:
            print(f"Not same, deleting replace_with from {dupe.get('material_id')}")

            pdf = await PDF.get_or_none(material_id=dupe.get("material_id"))
            if pdf:
                pdf.replace_with = None
                await pdf.save()
        input("Press enter to continue")


async def run():
    from easy_access.db.base import ensure_db_inited

    await ensure_db_inited()
    await main()


asyncio.run(run())
