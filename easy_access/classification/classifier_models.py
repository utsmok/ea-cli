from enum import Enum
from pydantic import BaseModel

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
