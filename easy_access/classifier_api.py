"""
This module uses an api-based service to classify documents.
"""
from easy_access.settings import SETTINGS, DirSetting
from dataclasses import dataclass
from google import genai
from easy_access.utils import info, warn, cool, File, Directory
from enum import Enum
from pydantic import BaseModel, TypeAdapter
import asyncio
import time
from rich.console import Console
import re
from easy_access.api_keys import gemini
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
    remarks: str # any additional remarks on the item relevant to copyright status, metadata, and item type
client = genai.Client(api_key=gemini)



async def classify_pdf(file: File) -> Classification:
    mat_id =file.name.split('_')[0]
    pdf = client.files.upload(
        file=file.path,
        config= {'mime_type': 'application/pdf',
                 'name': mat_id},
    )
    response = client.models.generate_content(
        model='gemini-2.0-flash',
        contents=[pdf,
"""Determine the copyright status and item type of this pdf file.
Pay close attention to any author affiliation, publisher names, and copyright statements you might find in the document, like 'all rights reserved', 'creative commons', 'Reproduction is allowed with acknowledgement', etcetera.
Determining the copyright status should be focused on the entire document. You can ignore copyrighted elements in a larger work, e.g. if a presentation includes copyrighted images or excerpts from a book: that is allowed under local copyright laws.
Take into account that the works are being used by a public institute for educational purposes in a closed environment: no commercial use and attribution is always given.
Include the final reason for the classification in the the response, as well as metadata like the author name, publisher name, and copyright holder name, etc where possible.
The requested output class contains hints & instructions as well, replicated here:
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
        author_name: list[str] # tthe name of the author(s) that created the item, if possible to determine
        publisher_name: str # who published the item, if possible to determine
        copyright_holder: str # who holds the copyright, if possible to determine
        item_title: str # the title of the item, if possible to determine
        doi: list[str] # the DOI(s) for the item if included in the document itself
        isbn: list[str]  # the ISBN(s) for the item if included in the document itself
        source_url: list[str] # the source URL(s) for the item if included in the document itself
        license: list[str] # the license(s) for the item if included in the document itself
        topic: str # the topic of the item, what it covers
        remarks: str # any additional remarks on the item relevant to copyright status, metadata, and item type

"""
],
        config={
            'response_mime_type': 'application/json',
            'response_schema': Classification,
        },
    )

    parsed = response.parsed
    client.files.delete(name=mat_id)

    if parsed:
        if isinstance(parsed, Classification):
            parsed.pdf_name = file.name
            return parsed

    return


def delete_files():
    console.print('Deleting files from gemini storage.')
    for f in client.files.list():
        console.print("Deleting: ", f.name)
        client.files.delete(name=f.name)

async def classify_items(files: list[File]) -> int:
    tasks = []
    for pdf in files:
        tasks.append(asyncio.create_task(classify_pdf(pdf)))
    classifications = await asyncio.gather(*tasks)
    classifications: list[Classification] = [c for c in classifications if c]
    console.print(f'{len(classifications)} files classified. Storing results as json files.')
    for classification in classifications:
        console.print(classification)
        name = classification.pdf_name.rstrip('.pdf')
        name = "".join([c for c in name if re.match(r'\w', c)])
        classification.pdf_name = name+".pdf"
        json_name = name+"_gemini_classification.json"
        console.print(f'Storing results as {json_name}')
        try:
            with open(SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full / f'{json_name}_gemini_classification.json', 'w') as f:
                f.write(classification.model_dump_json(indent=2, warnings='warn'))
        except Exception as e:
            console.print(f'Could not store {json_name}: {e}')
            continue
    return 1
async def main():

    all_files = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].files
    pdfs = [f for f in all_files if f.extension == '.pdf']
    pdfs_found = len(pdfs)
    existing_classifications = [f.name.rstrip('_gemini_classification.json')+'.pdf' for f in all_files if f.extension == '.json']
    if existing_classifications:
        pdfs = [f for f in pdfs if f.name not in existing_classifications]
    # filter out all files larger than 10 MB (in bytes)
    console.print(f'{pdfs_found} files found, with {len(existing_classifications)} already classified. Starting classification of {len(pdfs)} files.')

    pdfs = [f for f in pdfs if f.size < 10*1024*1024]
    pdfs = [f for f in pdfs if not any(['scipy' in f.name, 'numpy' in f.name])]
    console.print(f'Also removed items > 10MB and some specifically selected items. {len(pdfs)} items remaining.')
    pdf_batch = []
    batch_start_time = time.time()
    batch_size = 0
    for pdf in pdfs:
        pdf_batch.append(pdf)
        if len(pdf_batch) == 15: # rate limit to 15 requests per minute
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
