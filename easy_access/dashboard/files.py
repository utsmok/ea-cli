"""
This module provides utility functions for the dashboard to access and process
file-based data associated with copyright items. This includes retrieving
extracted text content (plain, annotated Markdown, or OCR'd) and structured
entity information (from JSON files).
"""

import json
import logging
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

logger = logging.getLogger(__name__)

# Attempt to import Entity from the classification module to avoid duplication
# If this path changes (e.g., Entity moved to classifier_models), this import needs update.
try:
    from easy_access.classification.entity_recognition import Entity
except ImportError:
    # Fallback or definition if it's decided Entity should live here or shared differently.
    # For now, assuming it should be imported. If not, the old definition would be here.
    logger.warning(
        "Could not import Entity from easy_access.classification.entity_recognition. Dashboard display of entities might be affected."
    )

    # Minimal fallback Entity definition if needed for the dashboard to run without full classification features.
    @dataclass
    class Entity:
        start: int
        end: int
        label: str
        text: str
        score: float | None = None


from easy_access.settings import SETTINGS, DirSetting

logger = logging.getLogger(__name__)


@dataclass
class Entities:
    """
    Represents and processes a collection of Entity objects extracted from a text.

    Attributes:
        items (list[Entity]): A list of Entity objects.
        grouped_by_label (Optional[dict[str, list[Entity]]]): Entities grouped by their labels.
                                                              Populated by `group_by_label`.
    """

    items: list[Entity] = field(default_factory=list)
    grouped_by_label: dict[str, list[Entity]] | None = field(init=False, default=None)

    def __init__(
        self, entities_data: list[dict[str, int | str | float | None]]
    ) -> None:
        """
        Initializes Entities with a list of dictionaries, converting them to Entity objects.

        Args:
            entities_data (list[dict[str, Union[int, str, float, None]]]):
                A list of dictionaries, where each dictionary represents an entity's data.
        """
        self.items = [Entity(**ent_data) for ent_data in entities_data]  # type: ignore
        # The type ignore might be needed if the fallback Entity is used and doesn't perfectly match

    def group_by_label(self) -> None:
        """Groups entities by their labels and stores them in `self.grouped_by_label`."""
        self.grouped_by_label = defaultdict(list)
        for ent in self.items:
            self.grouped_by_label[ent.label].append(ent)

    def get_grouped_and_sorted_entities(
        self,
    ) -> tuple[dict[str, list[Entity]], list[str]]:
        """
        Groups entities by label and sorts them based on predefined label priority and then alphabetically.

        Returns:
            Tuple[dict[str, list[Entity]], list[str]]:
                A tuple containing:
                - dictionary of entities grouped by sorted labels.
                - list of sorted label keys.
        """
        if self.grouped_by_label is None:  # Ensure grouping is done
            self.group_by_label()

        # Ensure self.grouped_by_label is not None for type checkers after the call above
        if self.grouped_by_label is None:  # Should not happen if group_by_label works
            return {}, []

        # Define sorting key for labels (prioritize certain labels)
        def sort_key_for_labels(label: str) -> int:
            label_lower = label.lower()
            if "recognized" in label_lower:
                return 0
            if "copyright" in label_lower or "license" in label_lower:
                return 1
            if "university" in label_lower:
                return 2
            return 3  # Everything else

        sorted_label_keys = sorted(
            self.grouped_by_label.keys(), key=sort_key_for_labels
        )

        # Sort items within each label group alphabetically by entity text
        # (The original sorted by label, which is already the key; changed to sort by text)
        sorted_grouped_entities: dict[str, list[Entity]] = {}
        for label_key in sorted_label_keys:
            sorted_grouped_entities[label_key] = sorted(
                self.grouped_by_label[label_key], key=lambda x: x.text.lower()
            )

        return sorted_grouped_entities, sorted_label_keys


def get_entities(material_id: int) -> Entities | None:
    """
    Retrieves and structures entity data for a given material_id from its JSON file.
    The JSON file is expected to be named `{material_id}_entities.json` (previously `_annotated.json`)
    and located in the PDF_DOWNLOADS directory.

    Args:
        material_id (int): The material ID for which to retrieve entities.

    Returns:
        Optional[Entities]: An Entities object if the file is found and parsed successfully,
                            None otherwise.
    """
    if (
        not SETTINGS.dirs.get(DirSetting.PDF_DOWNLOADS)
        or not SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full.exists()
    ):
        logger.error(
            f"PDF_DOWNLOADS directory not configured or does not exist. Cannot get entities for {material_id}."
        )
        return None

    # Updated filename based on changes in entity_recognition.py
    entities_file_path = (
        SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full / f"{material_id}_entities.json"
    )

    if not entities_file_path.exists():
        logger.debug(
            f"Entities file not found for material_id {material_id} at {entities_file_path}"
        )
        return None

    try:
        with open(entities_file_path, encoding="utf-8") as f:  # Specify read mode
            entities_data_list: list[dict[str, int | str | float | None]] = json.load(f)
        return Entities(entities_data_list)
    except json.JSONDecodeError as e:
        logger.error(
            f"Error decoding JSON from entities file {entities_file_path}: {e}"
        )
    except (
        Exception
    ) as e:  # Catch other potential errors like file read issues or Entities init error
        logger.error(f"Error processing entities file {entities_file_path}: {e}")
    return None


def get_extracted_text(material_id: int) -> str:
    """
    Retrieves extracted text for a given material_id.
    It searches for text files in a specific order of preference based on suffixes:
    1. `{material_id}_annotated_text.md` (HTML annotated text)
    2. `{material_id}_paddle.txt` (PaddleOCR raw text)
    3. `{material_id}.md` (Plain Markdown text from kreuzberg direct extraction)
    4. `{material_id}.txt` (Plain text, possibly older format)

    Args:
        material_id (int): The material ID for which to retrieve text.

    Returns:
        str: The extracted text content if found, otherwise an error message string.
    """
    if (
        not SETTINGS.dirs.get(DirSetting.PDF_DOWNLOADS)
        or not SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full.exists()
    ):
        logger.error(
            f"PDF_DOWNLOADS directory not configured or does not exist. Cannot get text for {material_id}."
        )
        return "Error: Text directory not configured."

    # Suffixes to check in order of priority (new filenames from entity_recognition.py)
    # TODO: Consider making these suffixes configurable if they change often.
    preferred_suffixes: list[str] = ["_annotated_text.md", "_paddle.txt", ".md", ".txt"]

    extracted_text_content: str = (
        f"No extracted text file found for material ID {material_id}."
    )
    found_file_path: Path | None = None

    for suffix in preferred_suffixes:
        potential_file_path = (
            SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full / f"{material_id}{suffix}"
        )
        if potential_file_path.exists():
            found_file_path = potential_file_path
            break  # Found the highest priority file

    if found_file_path:
        logger.info(
            f"Retrieving extracted text for {material_id} from {found_file_path.name}"
        )
        try:
            with open(found_file_path, encoding="utf-8") as f:  # Specify read mode
                extracted_text_content = f.read()
        except Exception as e:
            logger.error(f"Error reading text file {found_file_path}: {e}")
            extracted_text_content = f"Error loading text from {found_file_path.name}."
    else:
        logger.info(
            f"No text file found for material ID {material_id} with any of the expected suffixes."
        )

    return extracted_text_content
