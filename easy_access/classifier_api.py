"""
This module uses an api-based service to classify documents.
"""
from easy_access.settings import SETTINGS, DirSetting
from google import genai
from easy_access.utils import  File
from enum import Enum
from pydantic import BaseModel
import asyncio
import time
from rich.console import Console
import re
from easy_access.api_keys import gemini
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
    contains the itemtype and copyright status of a pdf file.
    """
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

client = genai.Client(api_key=gemini)
prompt = """First extract and determine a list of metadata, then determine the copyright status and item type for the included pdf file or plain text parsed from pdf file.
The metadata should help determine the copyright status and item type,  e.g. the author name, publisher name, and copyright holder name, license statements, etc.
The copyright status should be focused on the overall document, please ignore any possible copyrighted elements included inside the work.
If the detected 'publisher' is the University of Twente, or an 'author' is employed by the University of Twente, please consider the work as OWN_MATERIAL.
Take into account that the works are being used by a public institute for educational purposes in a closed environment: there is never any commercial use, and attribution is always given.
Include the final reason for the classification in the the response.

The requested output classes are replicated here including more details, hints, and suggestions:
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
    class Classification(BaseModel):
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


async def classify_pdf(file: File, full_pdf:bool = True) -> Classification:
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
                contents = [pdf,f"You received the complete contents of the pdf file {file.name}.\n"+prompt]
            except Exception as e:
                print(e)
                ...
        else:
            print(f'Extracting text from {file.name}')
            pdf: str = extract_text(pdf_file=file.path)
            if len(pdf)> 1_000_000:
                pdf = pdf[:1_000_000]
            contents = f"\n | text content of pdf file {file.name} is as follows: |\n".join([prompt,pdf])
        print(f'sent request for {mat_id}')
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
        full_pdf=False),
        files,
        max_at_once=5, # Limit maximum number of concurrently running tasks.
        max_per_second=1,  # Limit request rate to not overload the server.
    ) as classifications:
        async for classification in classifications:
            if not classification:
                continue
            console.print(classification)
            name = classification.pdf_name.rstrip('.pdf')
            name = "".join([c for c in name if re.match(r'\w', c)])
            classification.pdf_name = name+".pdf"
            json_name = name+"_gemini_classification.json"
            console.print(f'Storing results as {json_name}')
            try:
                with open(SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full / f'{json_name}', 'w') as f:
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

    delete_files()
    all_files = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].files
    pdfs = [f for f in all_files if f.extension == '.pdf']
    pdfs_found = len(pdfs)
    existing_classifications = [f.name.rstrip('_gemini_classification.json')+'.pdf' for f in all_files if f.extension == '.json']
    if existing_classifications:
        pdfs = [f for f in pdfs if f.name not in existing_classifications]
    # filter out all files larger than 10 MB (in bytes)
    console.print(f'{pdfs_found} files found, with {len(existing_classifications)} already classified. Starting classification of {len(pdfs)} files.')
    # remove pdfs with more than 1000 pages
    pdfs = [f for f in pdfs if get_page_count(f) < 1000]
    pdfs = [f for f in pdfs if not any(['scipy' in f.name, 'numpy' in f.name])]
    console.print(f'Also removed items > 1000 pages and some specifically selected items. {len(pdfs)} items remaining.')
    pdf_batch = []
    batch_start_time = time.time()
    batch_size = 0

    for pdf in pdfs:
        pdf_batch.append(pdf)
        if len(pdf_batch) == 15: # rate limit to 30 requests per minute
            print('awaiting results for a batch of 15 files...')
            result = await classify_items(pdf_batch)
            while result != 1:
                pass
            if time.time()-batch_start_time < 60:
                console.print(f'Sleeping for {60-(time.time()-batch_start_time)} seconds to avoid rate limit.')
                await asyncio.sleep(60-(time.time()-batch_start_time))
            delete_files()
            pdf_batch = []
            batch_start_time = time.time()

    result = await classify_items(pdf_batch)
