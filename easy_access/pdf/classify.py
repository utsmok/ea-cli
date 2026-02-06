# entity extraction, document classification, etcetera.

import re
from re import Pattern
from typing import Literal

from flashtext import KeywordProcessor
from gliner import GLiNER
from langchain_text_splitters import RecursiveCharacterTextSplitter
from loguru import logger
from rich import print

from easy_access.pdf.models import Entity, ExtractedEntities

MODEL = GLiNER.from_pretrained(
    "gliner-community/gliner_large-v2.5", load_tokenizer=True
).to("cuda")
LABELS = [
    "author",
    "professor",
    "faculty",
    "publisher",
    "university",
    "license",
    "copyright statement",
    "copyright holder",
    "email",
]

SPLITTER = RecursiveCharacterTextSplitter(
    chunk_size=700,
    chunk_overlap=0,
    length_function=len,
    is_separator_regex=False,
)


def find_entities(
    text: str, skip_extraction: bool = False, merged: list[dict] | None = None
) -> tuple[list[Entity], str]:
    """
    uses GLiNER to find entities in the given text
    returns a list of found Entities and the text as used by the model
    """

    def merge_entities(entities):
        if not entities:
            return []
        merged = []
        current = entities[0]
        for next_entity in entities[1:]:
            if next_entity["label"] == current["label"] and (
                next_entity["start"] == current["end"] + 1
                or next_entity["start"] == current["end"]
            ):
                current["text"] = text[current["start"] : next_entity["end"]].strip()
                current["end"] = next_entity["end"]
            else:
                merged.append(current)
                current = next_entity
        merged.append(current)
        return merged

    results = []
    passed_chars = 0
    new_text = ""
    for chunk in SPLITTER.split_text(text):
        if not skip_extraction:
            result = MODEL.predict_entities(chunk, LABELS, threshold=0.8)
            if result:
                for ent in result:
                    ent["start"] += passed_chars
                    ent["end"] += passed_chars
                results.extend(result)
        new_text += chunk
        passed_chars += len(chunk)

    if not skip_extraction:
        merged = merge_entities(results)
        for ent in merged:
            print(ent)
        entities = [Entity(**ent) for ent in merged]

    return entities, new_text


def regex_extraction(
    text: str, specific_entitites: dict[str, list[str]] | None = None
) -> ExtractedEntities:
    """
    use regex to extract specific entities:
    - email + url
    - dois
    - isbns
    - issns
    - orcid ids
    - exact matches for names

    for each match, create an entity object and add it to the list of entities
    """
    isbn_regex: tuple[Pattern[str], Literal["isbn"]] = (
        re.compile(
            r"(ISBN[-]*(1[03])*[ ]*(: ){0,1})*(([0-9Xx][- ]*){13}|([0-9Xx][- ]*){10})"
        ),
        "isbn",
    )

    url_regex: tuple[Pattern[str], Literal["url"]] = (
        re.compile(
            r"^(?:(?:http|https|ftp|telnet|gopher|ms\-help|file|notes)://)?(?:(?:[a-z][\w~%!&amp;',;=\-\.$\(\)\*\+]*):.*@)?(?:(?:[a-z0-9][\w\-]*[a-z0-9]*\.)*(?:(?:(?:(?:[a-z0-9][\w\-]*[a-z0-9]*)(?:\.[a-z0-9]+)?)|(?:(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)))(?::[0-9]+)?))?(?:(?:(?:/(?:[\w`~!$=;\-\+\.\^\(\)\|\{\}\[\]]|(?:%\d\d))+)*/(?:[\w`~!$=;\-\+\.\^\(\)\|\{\}\[\]]|(?:%\d\d))*)(?:\?[^#]+)?(?:#[a-z0-9]\w*)?)?$"
        ),
        "url",
    )
    doi_regex: tuple[Pattern[str], Literal["doi"]] = (
        re.compile(r"^(10\.\d{4,5}\/[\S]+[^;,.\s])$"),
        "doi",
    )

    issn_regex: tuple[Pattern[str], Literal["issn"]] = (
        re.compile(r"\d{4}-\d{3}(\d|x|X)"),
        "issn",
    )
    orcid_regex: tuple[Pattern[str], Literal["orcid"]] = (
        re.compile(r"^(\d{4}-){3}\d{3}(\d|X)$"),
        "orcid",
    )

    # TODO: add more specific names/entities?
    keyword_dict = {
        "University of Twente": [
            "universiteit twente",
            "twente university",
        ],
        "@utwente.nl": ["@utwente", "utwente.nl"],
        "Faculty of Behavioral, Management and Social Sciences": [
            "Behavioral, Management and Social Sciences"
        ],
        "Faculty of Electrical Engineering, Mathematics and Computer Science": [
            "Electrical Engineering, Mathematics and Computer Science",
            "EEMCS",
        ],
        "Faculty of Engineering Technology": ["engineering technology"],
        "Faculty of Geo-Information Science and Earth Observation": [
            "Geo-Information Science and Earth Observation"
        ],
        "Faculty of Science and Technology": ["science and technology"],
    }
    # join with specific_entities if provided
    if specific_entitites:
        keyword_dict.update(specific_entitites)

    keyword_processor = KeywordProcessor(case_sensitive=False)
    keyword_processor.add_keywords_from_dict(keyword_dict)
    entities: list[Entity] = []

    # first fasttext keyword search
    found_keywords = keyword_processor.extract_keywords(text, span_info=True)
    for result in found_keywords:
        keyword, start, end = result
        text = text[start:end]
        entity = Entity(
            text=text,
            label="specific_name",
            entity_name=keyword,
            start=start,
            end=end,
            confidence=1.0,
        )
        entities.append(entity)

    # then regex searches
    for regex, label in [isbn_regex, url_regex, doi_regex, issn_regex, orcid_regex]:
        matches = re.finditer(regex, text)
        for match in matches:
            start = match.start()
            end = match.end()
            entity_text = text[start:end]
            entities.append(
                Entity(
                    text=entity_text,
                    label=label,
                    entity_name=entity_text,
                    start=start,
                    end=end,
                    confidence=1.0,
                )
            )

    return ExtractedEntities(entities)


def format_labels(labels: set[str]) -> str:
    """Sorts and joins a set of labels into a comma-separated string."""
    return ", ".join(sorted(list(labels)))


def filter_and_merge_overlapping_entities(entities: list[Entity]) -> list[Entity]:
    """
    Filters and merges overlapping entities based on span and labels.

    Sorts entities by start position, then end position descending (longer entities first).
    Iterates through the sorted list. If an overlap occurs, the entity with the
    larger span (which comes first due to sorting) is kept, and the labels of
    overlapping entities are merged into it.

    Args:
        entities: A list of Entity objects.

    Returns:
        A list of Entity objects with overlaps resolved and labels merged.
    """
    if not entities:
        return []

    # Sort by start index ascending, then by end index descending (prioritize longer spans)
    entities.sort(key=lambda e: (e.start, -e.end))

    merged_entities = []
    # Use a temporary structure to hold entities and their collected labels during processing
    # Stores tuples of: (dominant_entity, set_of_labels_for_this_entity)
    temp_merged_info: list[tuple[Entity, set[str]]] = []

    for current_entity in entities:
        if not temp_merged_info:
            # First entity, add it with its label in a set
            temp_merged_info.append((current_entity, {current_entity.label}))
            continue

        # Get the last dominant entity added and its current label set
        last_entity, last_labels = temp_merged_info[-1]

        # Check for overlap: current entity starts before the last dominant one ends
        if current_entity.start < last_entity.end:
            # Overlap detected. Add current entity's label to the dominant one's set.
            # The current_entity itself is effectively discarded because the last_entity
            # has the larger or equal span and starts earlier or at the same position (due to sorting).
            last_labels.add(current_entity.label)
        else:
            # No overlap, add the current entity as a new dominant entity
            temp_merged_info.append((current_entity, {current_entity.label}))

    # Finalize the entities: update their label field with the formatted string from the collected set
    for entity, labels in temp_merged_info:
        entity.label = format_labels(labels)
        merged_entities.append(entity)

    return merged_entities


def update_text(original_text: str, entities: list[Entity]) -> str:
    """
    Wraps identified entities in the original text with HTML <mark> tags.

    Handles potential overlaps by filtering and merging them first.
    Builds the annotated text segment by segment based on original indices.

    Args:
        original_text: The original text string.
        entities: A list of Entity objects found in the original_text.

    Returns:
        A new string with entities wrapped in HTML tags, using merged labels for overlaps.
    """
    if not entities:
        return original_text

    # Filter and merge overlapping entities using the new logic
    processed_entities = filter_and_merge_overlapping_entities(entities)

    # Sort the processed entities by start position (should be mostly sorted, but ensures correctness)
    processed_entities.sort(key=lambda x: x.start)

    result_parts = []
    current_pos = 0

    for entity in processed_entities:
        # Add the text segment before the current entity
        if entity.start > current_pos:
            result_parts.append(original_text[current_pos : entity.start])
        elif entity.start < current_pos:
            # This condition should ideally not be met after filtering/merging. Log if it happens.
            logger.warning(
                "Skipping entity due to unexpected overlap/order after filtering: {}",
                entity,
            )
            continue

        # Add the wrapped entity text - using the potentially merged label
        entity_text_segment = original_text[entity.start : entity.end]
        wrapped_entity = (
            f'<mark x-data x-tooltip="{entity.label}">{entity_text_segment}</mark>'
        )
        result_parts.append(wrapped_entity)

        # Update the current position to the end of the current entity
        current_pos = entity.end

    # Add any remaining text after the last entity
    if current_pos < len(original_text):
        result_parts.append(original_text[current_pos:])

    return "".join(result_parts)
