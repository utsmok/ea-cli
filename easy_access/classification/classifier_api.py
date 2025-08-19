"""
This module uses an api-based service to classify documents.
"""

import asyncio
import io
import os
import time
from functools import partial

import pikepdf
from aiometer import amap
from google import genai
from rich.console import Console

from easy_access.classification.classifier_models import Classification
from easy_access.classification.pdf_handling import extract_pdf_text
from easy_access.db.base import init
from easy_access.db.ingest import load_llm_classifications
from easy_access.db.models import PDF, CopyrightItem
from easy_access.settings import SETTINGS, DirSetting
from loguru import logger

console = Console(emoji=True, markup=True)

client = None
prompt = """From the included document, first extract and determine a list of metadata, then determine the copyright status and item type for this item.
Finally determine the most important classification: if the item is allowed to be shared with students in the context of the University of Twente learning environment.
Use all available (meta)data in the file or that you extracted earlier (e.g. the author name, publisher name, and copyright holder name, license statements, etc.) to help determine these statuses.
The copyright status should be focused on the overall document. You can ignore any possible copyrighted elements included inside the work (like images from other works).
For determining allowed use, take into account that the works are being shared internally at the University of Twente, a Dutch public institute, for educational purposes only, in a closed environment.
This means that clearly copyrighted commercial works cannot be used, except if educational use is explicitly allowed for instance.
There will never be any commercial use in this context. Assume attribution is always given.
If the detected 'publisher' or 'author' is the University of Twente or is employed by the University of Twente, the work should be classified as OWN_MATERIAL.
Include reasoning for the classification in the response in the corresponding fields.

The requested output format is replicated here as a set of Python classes, including additional details, hints, and suggestions.
    class ItemType(str, Enum):
        Possible item types of the item.
        PRESENTATION = "presentation" # a powerpoint in pdf format for example. By definition, this should have CopyrightStatus.OWN_MATERIAL.
        READER = "reader" # often self-written information by teachers for students for this specific course. By definition, this should have CopyrightStatus.OWN_MATERIAL.
        BOOK = "book" # Often COPYRIGHTED_MATERIAL or OPEN_ACCESS.
        ARTICLE = "article" # Often COPYRIGHTED_MATERIAL or OPEN_ACCESS.
        REPORT = "report" # Often COPYRIGHTED_MATERIAL or OPEN_ACCESS.
        ASSIGNMENT = "assignment" # an assignment description for this course. By definition, this should have CopyrightStatus.OWN_MATERIAL
        THESIS = "thesis" # Often COPYRIGHTED_MATERIAL or OPEN_ACCESS.
        MANUAL = "manual" # e.g. for a measuring device. Often COPYRIGHTED_MATERIAL or OPEN_ACCESS, but can be OWN_MATERIAL.
        UNKNOWN = "unknown" # if not possible to determine.
    class CopyrightStatus(str, Enum):
        The possible copyright classifications
        OPEN_ACCESS = "open access" # free to use
        OWN_MATERIAL = "own material" # made for or by an employee of the university of Twente
        COPYRIGHTED_MATERIAL = "copyrighted material" # not free to use, owned by a publisher for instance
        OTHER = "other" # should not be used? maybe if unable to classify otherwise.
    class AllowedUsageByUT(str, Enum):
        These possible classifications denote if the item is allowed to be shared with students in the context of the University of Twente learning environment.
        ALLOWED = "allowed" # the item can be shared with students without further limitations, e.g. it is open access or own material by a UT employee.
        RESTRICTED = "restricted" # the item has limitations on sharing; e.g. only this year, only if the uploader is the author, or only with specific permissions/acknowledgements etc.
        NOT_ALLOWED = "not allowed" # the item cannot be shared without further permissions, e.g. it is fully copyrighted without any other routes to obtain permission
        UNDETERMINED = "undetermined" # the item cannot be classified as allowed or not allowed, e.g. if the classification is not possible due to missing or conflicting information.
    class Classification(BaseModel):
        allowed_usage: AllowedUsageByUT = AllowedUsageByUT.UNDETERMINED
        allowed_usage_reasoning: str # add a 1 to 2 sentence explanation on why this allowed usage was chosen
        copyright_status: CopyrightStatus = CopyrightStatus.OTHER
        copyright_classification_reason: str  # add a 1 to 2 sentence explanation on why this copyright status was chosen
        item_type: ItemType = ItemType.UNKNOWN
        item_type_classification_reason: str # add a 1 to 2 sentence explanation on why this item type was chosen
        pdf_name: str

        # metadata fields -- if not possible to determine from the pdf store an empty string instead
        author_name: list[str] # the name of the author(s) that created the item, if possible to determine
        publisher_name: str # who published the item, if possible to determine
        copyright_holder: str # who holds the copyright, if possible to determine
        item_title: str # the title of the item, if possible to determine
        doi: list[str] # the DOI(s) for the item if included in the document itself
        isbn: list[str]  # the ISBN(s) for the item if included in the document itself
        source_url: list[str] # the source URL(s) for the item if included in the document itself
        license: list[str] # the license(s) for the item if included in the document itself, or any license-related statement like 'all rights reserved', 'creative commons', 'Reproduction is allowed with acknowledgement'.
        topic: str # the topic of the item, what it covers
        pdf_page_count: int # the amount of pages in the pdf
        remarks: str # any additional remarks on the item relevant to copyright status, metadata, and item type
"""


def activate_client():
    global client
    client = genai.Client(api_key="gemini_api_key")


async def classify_pdf(
    pdf: PDF, full_pdf: bool = False
) -> tuple[PDF, Classification | None]:
    async def send_pdf_to_gemini(pdf: PDF, max_pages: int = 20) -> io.BytesIO:
        """
        Sends the first 'max_pages' of a PDF file to the Gemini API.

        Args:
            pdf_path (Path): Path to the PDF file.
            max_pages (int): The maximum number of pages to send. Defaults to 20.

        Returns:
            bytes: The content of the first 'max_pages' of the PDF as bytes.
                Returns an empty bytes object if there's an error.
        """
        try:
            with pikepdf.open(pdf.path) as opened_pdf:
                # Create a new PDF to hold the first 'max_pages' pages
                new_pdf = pikepdf.Pdf.new()
                num_to_copy = min(max_pages, len(opened_pdf.pages))
                if num_to_copy > 0:
                    # opened_pdf.pages is a PageList, which supports slicing.
                    # The slice itself is iterable and can be used with extend.
                    new_pdf.pages.extend(opened_pdf.pages[0:num_to_copy])
                # Store the new PDF in memory as bytes
                temp_stream = io.BytesIO()
                new_pdf.save(temp_stream)
                return temp_stream.getvalue()
        except Exception as e:
            print(f"Error processing PDF: {e}")
            return b""

    contents = None
    try:
        mat_id = pdf.material_id
        if pdf.parsing_failed:
            full_pdf = True

        # Combined SIM102: if not full_pdf and not pdf.extracted_text
        if not full_pdf and not pdf.extracted_text:
            logger.warning("pdf has no extracted text. Trying to extract.")
            await extract_pdf_text(pdf)
            pdf = await PDF.get(material_id=mat_id) # Re-fetch pdf after potential modification
            if not pdf.extracted_text: # Check again after trying to extract
                logger.warning(
                    f"pdf still has no extracted text. Sending full pdf instead for {pdf.current_file_name}"
                )
                full_pdf = True

        if pdf.extracted_text and not full_pdf:
            pdf_text = pdf.extracted_text
            if len(pdf_text) > 10_000:
                pdf_text = pdf_text[:10_000]
            if len(pdf_text) < 100:
                logger.warning(f"pdf text is too short for {pdf.current_file_name}")
                full_pdf = True
            if not full_pdf:
                contents = f"\n | text content of pdf file {pdf.current_file_name} is as follows: |\n".join(
                    [prompt, pdf_text]
                )
        if full_pdf:
            print(f"Uploading {pdf.current_file_name} to gemini storage.")
            try:
                # select only the first 20 pages
                pdf_bytes = await send_pdf_to_gemini(pdf, max_pages=20)
                if not pdf_bytes:
                    logger.warning("Could not send pdf to gemini storage.")
                    return pdf, None

                pdf_file = client.files.upload(
                    file=io.BytesIO(pdf_bytes),
                    config={"mime_type": "application/pdf", "name": str(mat_id)},
                )
                contents = [
                    pdf_file,
                    f"You received the pdf file {pdf.current_file_name}.\n" + prompt,
                ]
            except Exception as e:
                print(e)
                return pdf, None

        print(f"sent request for {mat_id}")
        if not mat_id:
            print(f"Could not extract material id from {pdf.current_file_name}")
            return pdf, None
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=contents,
            config={
                "response_mime_type": "application/json",
                "response_schema": Classification,
            },
        )
        parsed = None
        parsed = response.parsed
        if full_pdf:
            client.files.delete(name=str(mat_id))

        if parsed:
            if isinstance(parsed, Classification):
                parsed.pdf_name = pdf.current_file_name
                print(f"returned response for {mat_id}")
                return pdf, parsed
            else:
                print(f"Error parsing response for {mat_id}")
                return pdf, None
        else:
            if response.candidates:
                for candidate in response.candidates:
                    if candidate.finish_reason:
                        reason = candidate.finish_reason.value

            if reason:
                print(f"No result for {mat_id}, reason: {reason}")
            else:
                print(
                    f"No response for {mat_id},\n\n parsed: {parsed}.\n\n response:{response}"
                )
            return pdf, None
    except Exception as e:
        print(f"Error while classifying {pdf.current_file_name}: {e}")
        console.print(e)
        return


def delete_files():
    console.print("Deleting files from gemini storage.")
    for f in client.files.list():
        console.print("Deleting: ", f.name)
        client.files.delete(name=str(f.name))


async def classify_items(files: list[PDF]) -> int:
    async with amap(
        partial(classify_pdf, full_pdf=False),
        files,
        max_at_once=5,  # Limit maximum number of concurrently running tasks.
        max_per_second=1,  # Limit request rate to not overload the server.
    ) as classifications:
        async for details in classifications:
            if isinstance(details, tuple):
                pdf, classification = details
            else:
                continue
            if not classification:
                continue
            console.print(classification)
            mat_id = pdf.material_id
            json_name = f"{mat_id}.json"
            console.print(f"Storing results as {json_name}")
            if os.path.exists(
                SETTINGS.dirs[DirSetting.CLASSIFICATIONS].full / f"{json_name}"
            ):
                os.rename(
                    SETTINGS.dirs[DirSetting.CLASSIFICATIONS].full / f"{json_name}",
                    SETTINGS.dirs[DirSetting.CLASSIFICATIONS].full
                    / f"{json_name.rstrip('.json')}_old.json",
                )

            try:
                with open(
                    SETTINGS.dirs[DirSetting.CLASSIFICATIONS].full / f"{json_name}",
                    "w",
                    encoding="utf-8",
                ) as f:
                    f.write(classification.model_dump_json(indent=2, warnings="warn"))
            except Exception as e:
                console.print(f"Could not store {json_name}: {e}")
                continue

    return 1


async def main(subset: list[int] | list[str] | None = None):
    # PAGE_COUNT_REGEX = re.compile(rb"/Type\s*/Page([^s]|$)", re.MULTILINE | re.DOTALL) # Unused

    # def get_page_count(file: File, regex=PAGE_COUNT_REGEX): # Unused function
    #     """Count number of pages in a pdf"""
    #     with open(file.path, "rb") as f:
    #         return len(regex.findall(f.read()))

    await init()
    activate_client()
    if not subset:
        all_copyright_items = (
            await CopyrightItem.all()
            .prefetch_related("llm_classification")
            .values("material_id", "llm_classification_id")
        )
        items_without_classifications = [
            item["material_id"]
            for item in all_copyright_items
            if not item["llm_classification_id"]
        ]
        pdfs = await PDF.filter(material_id__in=items_without_classifications).all()
    else:
        subset = [int(item) for item in subset]
        selected_items = (
            await CopyrightItem()
            .filter(material_id__in=subset)
            .prefetch_related("llm_classification")
            .values("material_id", "llm_classification_id")
        )
        items_without_classifications = [
            item["material_id"]
            for item in selected_items
            if not item["llm_classification_id"]
        ]
        pdfs = await PDF.filter(material_id__in=items_without_classifications).all()

    console.print(f"{len(pdfs)} files found requiring classification.")

    pdf_batch = []
    batch_start_time = time.time()
    for total, pdf in enumerate(pdfs, 1): # Start enumeration from 1
        pdf_batch.append(pdf)
        if len(pdf_batch) == 10:  # rate limit to 10 requests per minute
            print("awaiting results for a batch of 10 files...")
            await classify_items(pdf_batch)
            print(f"\n           Processed {total}/{len(pdfs)} files.\n\n")
            if time.time() - batch_start_time < 120:
                console.print(
                    f"Sleeping for {60 - (time.time() - batch_start_time)} seconds to avoid rate limit."
                )
                await asyncio.sleep(60 - (time.time() - batch_start_time))
            delete_files()
            pdf_batch = []
            batch_start_time = time.time()

    await load_llm_classifications()
