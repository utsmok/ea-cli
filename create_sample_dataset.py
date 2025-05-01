"""
This script will output an excel file with a sample dataset for the easy access group.

"""

import asyncio
from collections import defaultdict
from pathlib import Path
from typing import Annotated

import polars as pl
import typer
from sqlalchemy import Engine, create_engine

from easy_access.classification.classifier_api import main
from easy_access.classification.pdf_handling import enrich_pdfs
from easy_access.classification.to_delete_downloader import Downloader
from easy_access.db.base import init, set_db_path
from easy_access.db.ingest import load_pdfs
from easy_access.db.retrieve import retrieve_full_data
from easy_access.db.update import update_copyright_relations
from easy_access.main import EasyAccessTool
from easy_access.settings import SAMPLESETTINGS, SETTINGS, EasyAccessSettings, Functions
from easy_access.utils import cool, info, warn

cli_app = typer.Typer()


async def download_files(subset=None, max_amount=None):
    info("downloading files. Will use Chrome do so.")
    warn(
        "Please make sure you have disabled all extensions, are logged in to Canvas, and have the correct permissions to download the files.\n Then completely close Chrome before continueing."
    )
    input("Press any key to continue...")
    downloader = Downloader()
    await downloader.download_pdfs(subset=None, max_amount=None)
    cool("done downloading!")


async def enrich(max_pages=50, str_limit=50000, pdfs=None, input_mat_ids=None):
    info("Loading existing PDFs into database.")
    await load_pdfs()
    cool("done loading existing PDFs into database.")
    info("Enriching & deduplicating PDFs.")
    await enrich_pdfs(max_pages, str_limit, pdfs, input_mat_ids)
    cool("done enriching & deduplicating PDFs.")


async def classify_items():
    info("Classifying PDFS.")
    await main()
    cool("done classifying PDFs.")


async def run_preprocessing(download, deduplicate, classify):
    info(
        "Running preprocessing steps: download files, deduplication, and classification."
    )
    if download:
        await download_files()
    if deduplicate:
        await enrich()
    if classify:
        await classify_items()


@cli_app.command()
def cli(
    separate_db: Annotated[
        bool,
        typer.Option(
            help="Use a separate database to create the sample data.",
            rich_help_panel="Data",
        ),
    ] = True,
    data_from: Annotated[
        str,
        typer.Option(
            help="Path to the data file to use for creating the sample dataset.",
            rich_help_panel="Data",
            show_default=False,
        ),
    ] = "",
    osiris: Annotated[
        bool,
        typer.Option(
            help="Include enrichment with Osiris data.",
            rich_help_panel="Enrichment",
        ),
    ] = False,
    download: Annotated[
        bool,
        typer.Option(
            help="Download pdfs from canvas.",
            rich_help_panel="Functions",
        ),
    ] = False,
    classify: Annotated[
        bool,
        typer.Option(
            help="Classify the pdfs by LLM.",
            rich_help_panel="Enrichment",
        ),
    ] = False,
    deduplicate: Annotated[
        bool,
        typer.Option(
            help="Deduplicate the pdfs.",
            rich_help_panel="Enrichment",
        ),
    ] = False,
) -> None:
    """cli to create sample datasets from easy access data."""

    if False:
        base_settings = SETTINGS
        sample_settings = SAMPLESETTINGS

        if data_from:
            if isinstance(data_from, str):
                data_from = Path(data_from)
            if not data_from.exists():
                warn(f"Data file not found at {data_from}.")
                raise typer.Exit(1)
        else:
            data_from = Path(sample_settings.input.get("file"))

        if separate_db:
            db_path = Path("sample_db.sqlite3")
            set_db_path(db_path)
            asyncio.get_event_loop().run_until_complete(init())
            # asyncio.get_event_loop().run_until_complete(create())
            # asyncio.get_event_loop().run_until_complete(load_base_data())
            # asyncio.get_event_loop().run_until_complete(
            #    load_raw_copyright_data(File(data_from))
            # )
        else:
            asyncio.get_event_loop().run_until_complete(init())

        if any([osiris, download, classify, deduplicate]):
            info(
                "Preprocessing steps selected. Doing a dry importing run of the data first, so it can be enriched."
            )
            ea_settings = EasyAccessSettings(
                functions=Functions.read,
                only_changes=False,
                refresh_osiris_data=osiris,
                other_sheet=data_from,  # path to data to read in?
                only_retrieve_missing_osiris_data=True,
                disable_writes=True,
                dirs=base_settings.dirs,
            )

            temp_tool = EasyAccessTool(settings=ea_settings)
            temp_tool.run()

            if any([download, deduplicate, classify]):
                info(
                    "Doing the rest of the preprocessing steps: download files, deduplication, and classification."
                )
                asyncio.get_event_loop().run_until_complete(
                    run_preprocessing(download, deduplicate, classify)
                )

    engine: Engine = create_engine("sqlite:///sample_db.sqlite3")
    db_path = Path("sample_db.sqlite3")
    set_db_path(db_path)
    asyncio.get_event_loop().run_until_complete(init())
    append = False
    create = True

    try:
        final_df = pl.read_excel(
            "sample_dataset.xlsx", infer_schema_length=None
        ).with_columns(pl.exclude(pl.String).cast(str))
        cool("Successfully loaded sample dataset.")
        create = False
    except FileNotFoundError:
        cool("Creating sample dataset.")
        create = True
    if not create:
        info(f"Loaded existing sample dataset with {len(final_df)} records.")
        # drop duplicate rows based on material_id
        final_df = final_df.unique("material_id")
        info(f"After dropping duplicates: {len(final_df)} records.")
        # drop rows without llm_classification_id
        # final_df = final_df.filter(pl.col("llm_classification_id").is_not_null())

        if len(final_df) < 500:
            append = True
            info(f"Only {len(final_df)} records found. Will append to the dataset.")

    input("Press any key to continue...")
    if create or append:
        # retrieve filtered data from the database for a new list of items
        # or append to the existing dataset
        query = """
        SELECT * FROM copyright_data
        WHERE faculty_id IN ('BMS', 'EEMCS', 'ET', 'ITC', 'TNW')
        AND classification NOT IN ('lange overname', 'eigen materiaal')
        AND filetype = 'pdf'
        AND status NOT IN ('Deleted')
        GROUP BY filename
        """
        # Execute with parameters
        with engine.connect() as conn:
            df = pl.read_database(
                query=query,
                connection=conn,
                infer_schema_length=None,
            )

        info(f"Retrieved {len(df)} records matching the filter criteria")

        # make a selection from the results:
        # -> For each faculty select 100 records with the following criteria:
        #    - there are at least 10 records for each classification in the selection
        #    - a max of 5 records per course_code

        if "faculty_id" in df.columns:
            df = df.rename({"faculty_id": "faculty"})
        if "created_at" in df.columns:
            df = df.drop("created_at")
        if "modified_at" in df.columns:
            df = df.drop("modified_at")

        df = df.unique("material_id")
        df = df.unique("filename")
        df = df.filter(pl.col("is_duplicate") == 0)
        info(
            f"{len(df)} records remaining after filtering out duplicate material_ids, filenames, and is_duplictate"
        )
        if append:
            # remove mat_ids already in the dataset
            existing_mat_ids = final_df.select("material_id").to_series().to_list()
            existing_mat_ids = {int(x) for x in existing_mat_ids}
            df = df.filter(~pl.col("material_id").is_in(existing_mat_ids))
            info(
                f"{len(df)} records remaining after filtering out material_ids already in the dataset"
            )
            for col in [
                "llm_allowed_usage",
                "llm_allowed_usage_reason",
                "llm_copyright",
                "llm_copyright_reason",
                "llm_item_type",
                "llm_remarks",
                "llm_title",
                "llm_copyright_holder",
                "llm_publisher",
                "llm_isbn",
                "llm_doi",
                "llm_source_url",
                "llm_license",
                "llm_authors",
                "cursuscodes",
                "programmes",
                "course_names",
                "course_contacts_names",
                "course_contacts_emails",
                "course_contacts_faculties",
                "course_contacts_organizations",
            ]:
                if col in df.columns:
                    final_df = final_df.drop(col)

        else:
            existing_mat_ids = set()

        num_per_faculty = 150
        num_per_course = 10
        faculties = ["BMS", "EEMCS", "ET", "ITC", "TNW"]
        classifications = df.select("classification").unique().to_series().to_list()
        num_per_classification = num_per_faculty // len(classifications)

        info(
            f"Found {len(classifications)} unique classifications. \nWill select {num_per_classification} records for each classification per faculty."
        )

        # Process each faculty separately
        final_selection: list[pl.DataFrame] = []
        per_course = defaultdict(int)

        for faculty in faculties:
            info(f"Processing faculty: {faculty}")

            # Filter for this faculty
            faculty_df = df.filter(pl.col("faculty") == faculty)

            if len(faculty_df) == 0:
                warn(f"No data found for faculty {faculty}")
                continue

            faculty_selection = []
            per_classification = defaultdict(int)

            if append:
                faculty_selection = final_df.filter(
                    pl.col("faculty") == faculty
                ).to_dicts()
                for row in faculty_selection:
                    code = row["course_code"]
                    if isinstance(code, list):
                        code = code[0]
                    classification = row["classification"]
                    if isinstance(classification, list):
                        classification = classification[0]
                    per_course[code] += 1
                    per_classification[classification] += 1

            for classification in classifications:
                class_df = faculty_df.filter(pl.col("classification") == classification)
                course_groups = class_df.group_by("course_code").all()

                for course_group in course_groups.to_dicts():
                    code = course_group["course_code"]
                    if isinstance(code, list):
                        code = code[0]
                    name = course_group["course_name"]
                    if isinstance(name, list):
                        name = name[0]

                    if (
                        per_classification[classification] < num_per_classification
                        and per_course[code] < num_per_course
                    ):
                        limit = min(
                            num_per_course - per_course[code],
                            num_per_classification - per_classification[classification],
                        )
                        selection = (
                            class_df.filter(pl.col("course_name") == name)
                            .filter(~pl.col("material_id").is_in(existing_mat_ids))
                            .to_dicts()
                        )
                        if len(selection) > limit:
                            selection = selection[:limit]
                        per_course[code] += len(selection)
                        per_classification[classification] += len(selection)
                        existing_mat_ids.update(
                            [int(x["material_id"]) for x in selection]
                        )
                        faculty_selection.extend(selection)

            info(
                f"[{faculty}] {len(faculty_selection)} records retrieved before limiting"
            )
            for var in ["classification", "course_name"]:
                info(f"{var}:")
                varcounts = (
                    pl.DataFrame(faculty_selection)
                    .select(var)
                    .group_by(var)
                    .agg(pl.len().alias("count"))
                    .sort("count", descending=True)
                    .to_dicts()
                )
                data = {list(k.values())[0]: list(k.values())[1] for k in varcounts}
                info(f"Count          {var}")
                for k, v in data.items():
                    print(f"{v}           {k}")
            if len(faculty_selection) > num_per_faculty:
                # Select random subset
                faculty_selection = pl.DataFrame(faculty_selection).sample(
                    num_per_faculty
                )
            else:
                # Convert to DataFrame
                faculty_result = pl.DataFrame(faculty_selection)

            # Distribute last_change dates evenly

            info(f"[{faculty}] final set: {len(faculty_result)} records")
            final_selection.append(
                faculty_result.with_columns(pl.exclude(pl.String).cast(str))
            )

        # for each df in final_selection, cast all cols to str
        for i, df in enumerate(final_selection):
            final_selection[i] = df

        final_df: pl.DataFrame = pl.concat(final_selection)
        final_df.write_excel("sample_dataset.xlsx")
        cool(
            f'Wrote base sample dataset with {len(final_selection)} to "sample_dataset.xlsx"'
        )

    selected_material_ids = (
        final_df.select("material_id").cast(pl.Int32).to_series().to_list()
    )

    selected_material_ids = list({int(x) for x in selected_material_ids})
    info(f"Selected {len(selected_material_ids)} material_ids for the dataset.")
    # retrieve the full data and see which items need additional enrichment
    data: pl.DataFrame = retrieve_full_data(selected_material_ids)
    # select material ids where llm.allowed_usage is empty
    material_ids_missing_classification = (
        data.filter(pl.col("llm_allowed_usage").is_null())
        .select("material_id")
        .to_series()
        .to_list()
    )

    material_ids_missing_classification = list(
        {int(x) for x in material_ids_missing_classification}
    )

    info(f"{len(material_ids_missing_classification)} material_ids needing enrichment.")
    info("Downloading PDFs for selected material_ids.")
    downloader = Downloader()
    try:
        asyncio.get_event_loop().run_until_complete(
            downloader.download_pdfs(material_ids_missing_classification)
        )
    except Exception as e:
        warn(f"Error downloading PDFs: {e}")

    try:
        info("Enriching PDFs for selected material_ids.")
        asyncio.get_event_loop().run_until_complete(
            enrich_pdfs(
                input_mat_ids=material_ids_missing_classification,
                max_pages=5,
                str_limit=5000,
            )
        )
    except Exception as e:
        warn(f"Error enriching PDFs: {e}")
    try:
        info("Classifying PDFs")
        asyncio.get_event_loop().run_until_complete(
            main(material_ids_missing_classification)
        )
    except Exception as e:
        warn(f"Error classifying PDFs: {e}")

    # now that all is processed, retrieve the detailed data from the db for the selected material_ids
    info(
        "Done with enrichment. Now:\n -> updating all relations\n -> retrieving detailed data for all selected material_ids\n -> writing to excel."
    )
    asyncio.get_event_loop().run_until_complete(update_copyright_relations())
    data: pl.DataFrame = retrieve_full_data(selected_material_ids)
    if Path("sample_dataset_full.xlsx").exists():
        Path("sample_dataset_full.xlsx").unlink()
    data.write_excel("sample_dataset_full.xlsx")
    cool("All done!")


if __name__ == "__main__":
    cli_app()
