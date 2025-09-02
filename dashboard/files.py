# functions related to files like pdfs, extracted text, etc.
import json
from collections import defaultdict
from dataclasses import dataclass, field

from loguru import logger

from easy_access.settings import SETTINGS, DirSetting


# dataclasses for entities
@dataclass
class Entity:
    """
    Represents an entity extracted from the text.
    Attributes:
        start (int): The starting index of the entity in the text.
        end (int): The ending index of the entity in the text.
        label (str): The label of the entity.
        text (str): The text of the entity.
        score (float | None): The confidence score of the entity (optional).
    """

    start: int
    end: int
    label: str
    text: str
    score: float | None = None


@dataclass
class Entities:
    """
    Represents a collection of entities extracted from the text.
    Attributes:
        items (list[Entity]): A list of Entity objects.

    Functions:
        group_by_label: Groups entities by their labels.
        group_and_sort: Sorts the grouped entities by label and alphabetically within each group.
    """

    items: list[Entity] = field(init=False, default_factory=list)
    grouped_by_label: dict[str, list[Entity]] = field(init=False, default=None)

    def __init__(self, entities: list[dict[str, int | str | float]]):
        self.items = [Entity(**ent) for ent in entities]

    def group_by_label(self):
        self.grouped_by_label: dict[str, list[Entity]] = defaultdict(list)

        for ent in self.items:
            self.grouped_by_label[ent.label].append(ent)

    def group_and_sort(self):
        if not self.grouped_by_label:
            self.group_by_label()

        sorted_labels = sorted(
            self.grouped_by_label.keys(),
            key=lambda x: (
                0
                if "recognized" in x.lower()
                # recognized first
                else 1
                if "copyright" in x.lower() or "license" in x.lower()
                # copyright or license next
                else 2
                if "university" in x.lower()
                # university next
                else 3  # everything else last
            ),
        )

        for label in sorted_labels:
            # sort alphabetically within each label group
            self.grouped_by_label[label] = sorted(
                self.grouped_by_label[label], key=lambda x: x.label.lower()
            )

        return self.grouped_by_label, sorted_labels


def get_entities(material_id: int) -> Entities | None:
    """
    This function retrieves the entities extracted from the annotated text for the given material_id.
    Entities are stored in a JSON file named "{material_id}_annotated.json" in the "pdf_downloads" folder.
    """
    entities_path = (
        SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full / f"{material_id}_annotated.json"
    )
    if not entities_path.exists():
        return None
    try:
        with open(entities_path, encoding="utf-8") as f:
            entities = Entities(json.load(f))
    except Exception as e:
        logger.error(f"Error processing entities file {entities_path}: {e}")
        return None

    return entities


def get_extracted_text(material_id: int) -> str:
    """
    Returns extracted text from the PDF file for the given material_id.
    If annotated text is available, it will be used; otherwise, the plain text will be returned.
    If neither are available, an str with an error message will be returned.
    """
    # suffixes to check in order of priority
    suffixes = ["_annotated.md", "_paddle.txt", ".md", ".txt"]
    for suff in suffixes:
        extracted_text_path = (
            SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full / f"{material_id}{suff}"
        )
        if extracted_text_path.exists():
            break
    if not extracted_text_path.exists():
        return "No extracted text found."

    logger.debug(
        f"retrieving extracted text for {material_id} from {extracted_text_path}"
    )

    text_element = f"Error loading text from {extracted_text_path.name}."
    if extracted_text_path.exists():
        try:
            with open(extracted_text_path, encoding="utf-8") as f:
                text_element = f.read()

        except Exception as e:
            logger.error(f"Error processing text file {extracted_text_path}: {e}")

    return text_element
