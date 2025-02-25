"""
This module uses an api-based service to classify documents.
"""
import datetime
import os
from easy_access.settings import SETTINGS, DirSetting
from google import genai
from easy_access.utils import  File, warn
from enum import Enum
from pydantic import BaseModel
import asyncio
import time
from rich.console import Console
import re
from easy_access.classification.api_keys import gemini
from aiometer import amap
from pdfminer.high_level import extract_text
from functools import partial

console = Console(emoji=True, markup=True)

class CopyrightStatus(str, Enum):
    """
    The possible classifications
    """
    OPEN_ACCESS = "open access" # free to use
    OWN_MATERIAL = "own material" # made for or by an employee of the university of Twente
    COPYRIGHTED_MATERIAL = "copyrighted material" # not free to use, owned by a publisher for instance
    OTHER = "other" # should not be used? maybe if unable to classify otherwise.

class AllowedUsageByUT(str, Enum):
    """
    These possible classifications denote if the item is allowed to be shared with students in the context of the University of Twente learning environment.
    """
    ALLOWED = "allowed" # the item can be shared with students without further limitations, e.g. it is open access or own material by a UT employee.
    RESTRICTED = "restricted" # the item has limitations on sharing; e.g. only this year, only if the uploader is the author, or only with specific permissions/acknowledgements etc.
    NOT_ALLOWED = "not allowed" # the item cannot be shared without further permissions, e.g. it is fully copyrighted without any other routes to obtain permission
    UNDETERMINED = "undetermined" # the item cannot be classified as allowed or not allowed, e.g. if the classification is not possible due to missing or conflicting information.

class ItemType(str, Enum):
    """
    Possible item types of the item
    """
    PRESENTATION = "presentation" # a powerpoint in pdf format for example. By definition, this should have CopyrightStatus.OWN_MATERIAL.
    READER = "reader" # often self-written information by teachers for students for this specific course. By definition, this should have CopyrightStatus.OWN_MATERIAL.
    BOOK = "book" # Often COPYRIGHTED_MATERIAL or OPEN_ACCESS.
    ARTICLE = "article" # Often COPYRIGHTED_MATERIAL or OPEN_ACCESS.
    REPORT = "report" # Often COPYRIGHTED_MATERIAL or OPEN_ACCESS.
    ASSIGNMENT = "assignment" # an assignment description for this course. By definition, this should have CopyrightStatus.OWN_MATERIAL
    THESIS = "thesis" # Often COPYRIGHTED_MATERIAL or OPEN_ACCESS.
    MANUAL = "manual" # e.g. for a measuring device. Often COPYRIGHTED_MATERIAL or OPEN_ACCESS, but can be OWN_MATERIAL.
    UNKNOWN = "unknown" # if not possible to determine.


class Classification(BaseModel):
    """
    contains the allowed usage status, itemtype, and copyright status of a pdf file.
    """
    allowed_usage: AllowedUsageByUT = AllowedUsageByUT.UNDETERMINED
    allowed_usage_reasoning: str # add a 1 to 2 sentence explanation on why this allowed usage was chosen
    copyright_status: CopyrightStatus = CopyrightStatus.OTHER
    copyright_classification_reason: str  # add a 1 to 2 sentence explanation on why this copyright status was chosen
    item_type: ItemType = ItemType.UNKNOWN
    item_type_classification_reason: str # add a 1 to 2 sentence explanation on why this item type was chosen
    pdf_name: str

    # metadata fields -- if not possible to determine from the pdf store an empty string instead
    author_names: list[str] # the name of the author(s) that created the item, if possible to determine
    publisher_name: str # who published the item, if possible to determine
    copyright_holder: str # who holds the copyright, if possible to determine
    item_title: str # the title of the item, if possible to determine
    doi: list[str] # the DOI(s) for the item if included in the document itself
    isbn: list[str] # the ISBN(s) for the item if included in the document itself
    source_url: list[str] # the source URL(s) for the item if included in the document itself
    license: list[str]  # the license(s) for the item if included in the document itself
    topic: str # the topic of the item, what it covers
    pdf_page_count: int # the amount of pages in the pdf

    remarks: str # any additional remarks on the item relevant to copyright status, metadata, and item type

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
    client = genai.Client(api_key=gemini)

async def classify_pdf(file: File, full_pdf:bool = False) -> Classification:
    try:
        mat_id =file.name.split('_')[0]
        if full_pdf:
            print(f'Uploading {file.name} to gemini storage.')
            try:
                pdf = client.files.upload(
                    file=file.path,
                    config= {'mime_type': 'application/pdf',
                            'name': mat_id},
                )
                contents = [pdf,f"You received either the first 20 pages (or complete contents if =< 20 pgs) of the pdf file {file.name}.\n"+prompt]
            except Exception as e:
                print(e)
                ...
        else:
            warn(f'TODO: use the extracted text from the PDF object in db instead!')
            print(f'Extracting text from {file.name}')
            pdf: str = extract_text(pdf_file=file.path, maxpages=8, codec='utf-8')
            if len(pdf)> 10_000:
                pdf = pdf[:10_000]
            contents = f"\n | text content of pdf file {file.name} is as follows: |\n".join([prompt,pdf])
        print(f'sent request for {mat_id}')
        if not mat_id:
            print(f'Could not extract material id from {file.name}')
            return
        response = client.models.generate_content(
            model='gemini-2.0-flash',
            contents=contents,
            config={
                'response_mime_type': 'application/json',
                'response_schema': Classification,
            },
        )

        parsed = response.parsed
        if full_pdf:
            client.files.delete(name=mat_id)

        if parsed:
            if isinstance(parsed, Classification):
                parsed.pdf_name = file.name
                print(f'returned response for {mat_id}')
                return parsed

        return
    except Exception as e:
        print(f'Error while classifying {file.name}: {e}')
        console.print(e)
        return


def delete_files():
    console.print('Deleting files from gemini storage.')
    for f in client.files.list():
        console.print("Deleting: ", f.name)
        client.files.delete(name=f.name)

async def classify_items(files: list[File]) -> int:
    async with amap(
        partial(classify_pdf,
        full_pdf=True),
        files,
        max_at_once=5, # Limit maximum number of concurrently running tasks.
        max_per_second=1,  # Limit request rate to not overload the server.
    ) as classifications:
        async for classification in classifications:
            warn(f'TODO: change to store in DB directly instead of writing to jsons? Or both?')
            if not classification:
                continue
            console.print(classification)
            mat_id = classification.pdf_name.split('_')[0]
            json_name = f'{mat_id}.json'
            console.print(f'Storing results as {json_name}')
            if os.path.exists(SETTINGS.dirs[DirSetting.CLASSIFICATIONS].full / f'{json_name}'):
                os.rename(SETTINGS.dirs[DirSetting.CLASSIFICATIONS].full / f'{json_name}',
                          SETTINGS.dirs[DirSetting.CLASSIFICATIONS].full / f'{json_name.rstrip(".json")}_old.json')

            try:
                with open(SETTINGS.dirs[DirSetting.CLASSIFICATIONS].full / f'{json_name}', 'w', encoding='utf-8') as f:
                    f.write(classification.model_dump_json(indent=2, warnings='warn'))
            except Exception as e:
                console.print(f'Could not store {json_name}: {e}')
                continue
    return 1
async def main():
    PAGE_COUNT_REGEX = re.compile(
        rb"/Type\s*/Page([^s]|$)",
        re.MULTILINE|re.DOTALL
    )

    def get_page_count(file:File, regex=PAGE_COUNT_REGEX):
        """Count number of pages in a pdf"""
        with open(file.path, "rb") as f:
            return len(regex.findall(f.read()))

    all_files = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].files
    pdfs = [f for f in all_files if f.extension == '.pdf']
    pdfs_found = len(pdfs)

    pdfs_for_mat_ids = [f for f in pdfs if "_" in f.name]
    pdf_material_ids: dict[str, File] = {f.name.split(sep='_')[0]:f for f in pdfs_for_mat_ids}

    activate_client()

    # existing_classifications = [f.name.rstrip('.json') for f in SETTINGS.dirs[DirSetting.CLASSIFICATIONS].files if f.extension == '.json']
    existing_classifications = [f.name.rstrip('.json') for f in SETTINGS.dirs[DirSetting.CLASSIFICATIONS].files if f.extension == '.json' and f.created >= datetime.datetime(year=2025,month=2,day=17, hour=11) and "_old" not in f.name]

    if existing_classifications:
        pdfs = [pdf_material_ids.get(f) for f in pdf_material_ids if f not in existing_classifications]

    console.print(f'{pdfs_found} files found, with {len(existing_classifications)} already classified. {len(pdfs)} files remaining.')
    pdf_batch = []
    batch_start_time = time.time()
    batch_size = 0

    for pdf in pdfs:
        pdf_batch.append(pdf)
        if len(pdf_batch) == 10: # rate limit to 10 requests per minute
            print('awaiting results for a batch of 10 files...')
            result = await classify_items(pdf_batch)
            while result != 1:
                pass
            if time.time()-batch_start_time < 120:
                console.print(f'Sleeping for {120-(time.time()-batch_start_time)} seconds to avoid rate limit.')
                await asyncio.sleep(120-(time.time()-batch_start_time))
            delete_files()
            pdf_batch = []
            batch_start_time = time.time()

    result = await classify_items(pdf_batch)
