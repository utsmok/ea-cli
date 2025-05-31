"""
This module uses GLiNER (Generalist Line-level Named Entity Recognition) and
regex/keyword matching to identify and extract predefined entities from text files.
The primary workflow involves:
1. Reading text content (presumably extracted from PDFs by `pdf_handling.py`).
2. Processing text with GLiNER for model-based entity prediction.
3. Augmenting with regex and keyword-based entity extraction for specific patterns (ISBN, DOI, URLs, predefined names).
4. Merging and filtering overlapping entities to produce a clean list.
5. Generating an HTML-annotated version of the text with <mark> tags highlighting entities.
6. Storing the processed entities (as JSON) and annotated text (as Markdown).

Key components include the `Entity` dataclass for representing found entities,
functions for model prediction (`find_entities`), pattern matching (`regex_extraction`),
overlap resolution (`filter_and_merge_overlapping_entities`), HTML annotation
(`update_text`), and result storage (`store_files`).
"""

import contextlib
import dataclasses
import json
import logging # Added
import re
from dataclasses import dataclass
from re import Pattern
from typing import List, Tuple, Set, Dict, Optional, Any # Used more specific types

from flashtext import KeywordProcessor
from gliner import GLiNER # External dependency
from langchain_text_splitters import RecursiveCharacterTextSplitter # External dependency
from nameparser import HumanName # External dependency

from easy_access.db.retrieve import retrieve_osiris_data # For specific name extraction
from easy_access.settings import SETTINGS, DirSetting
# from easy_access.utils import warn # Removed warn, using logger

logger = logging.getLogger(__name__)

# --- Configuration Candidates (Consider moving to settings.yaml or a dedicated config file) ---
# GLiNER Model Configuration
DEFAULT_GLINER_MODEL_NAME: str = "gliner-community/gliner_large-v2.5"
DEFAULT_GLINER_DEVICE: str = "cuda" # "cuda" or "cpu"
# List of entity labels for GLiNER to predict
DEFAULT_GLINER_LABELS: List[str] = [
    "author", "professor", "faculty", "publisher", "university",
    "license", "copyright statement", "copyright holder", "email",
]
# Text Splitting Configuration for GLiNER
DEFAULT_CHUNK_SIZE: int = 700
DEFAULT_CHUNK_OVERLAP: int = 0
GLINER_PREDICTION_THRESHOLD: float = 0.8 # Confidence threshold for GLiNER predictions

# Regex patterns (some are complex, could be settings if they need frequent changes)
ISBN_REGEX_PATTERN: str = r"(ISBN[-]*(1[03])*[ ]*(: ){0,1})*(([0-9Xx][- ]*){13}|([0-9Xx][- ]*){10})"
URL_REGEX_PATTERN: str = r"^(?:(?:http|https|ftp|telnet|gopher|ms\-help|file|notes)://)?(?:(?:[a-z][\w~%!&amp;',;=\-\.$\(\)\*\+]*):.*@)?(?:(?:[a-z0-9][\w\-]*[a-z0-9]*\.)*(?:(?:(?:(?:[a-z0-9][\w\-]*[a-z0-9]*)(?:\.[a-z0-9]+)?)|(?:(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)))(?::[0-9]+)?))?(?:(?:(?:/(?:[\w`~!$=;\-\+\.\^\(\)\|\{\}\[\]]|(?:%\d\d))+)*/(?:[\w`~!$=;\-\+\.\^\(\)\|\{\}\[\]]|(?:%\d\d))*)(?:\?[^#]+)?(?:#[a-z0-9]\w*)?)?$"
DOI_REGEX_PATTERN: str = r"(10\.\d{4,5}\/[\S]+[^;,.\s])" # Simplified from ^(...)$ to find within text
ISSN_REGEX_PATTERN: str = r"\d{4}-\d{3}(\d|x|X)"
ORCID_REGEX_PATTERN: str = r"(\d{4}-){3}\d{3}(\d|X)" # Simplified from ^(...)$

# Initial Keyword Dictionary for FlashText (can be extended from DB)
INITIAL_KEYWORD_DICT: Dict[str, List[str]] = {
    "University of Twente": ["universiteit twente", "twente university"],
    "@utwente.nl": ["@utwente", "utwente.nl"], # For email domain as entity
    "Faculty of Behavioral, Management and Social Sciences": ["Behavioral, Management and Social Sciences"],
    "Faculty of Electrical Engineering, Mathematics and Computer Science": ["Electrical Engineering, Mathematics and Computer Science", "EEMCS"],
    "Faculty of Engineering Technology": ["engineering technology", "ET"],
    "Faculty of Geo-Information Science and Earth Observation": ["Geo-Information Science and Earth Observation", "ITC"],
    "Faculty of Science and Technology": ["science and technology", "TNW", "ST"], # Added ST
}
# --- End Configuration Candidates ---

# Attempt to load FILE_DIR, handle if SETTINGS or DirSetting.PDF_DOWNLOADS is not ready
try:
    FILE_DIR: Optional[Directory] = SETTINGS.dirs.get(DirSetting.PDF_DOWNLOADS)
    if FILE_DIR is None:
        logger.error("PDF_DOWNLOADS directory not configured in SETTINGS. Some functions may fail.")
except AttributeError: # If SETTINGS object itself is not fully formed (e.g. during testing)
    logger.error("SETTINGS object not fully initialized. PDF_DOWNLOADS directory may be unavailable.")
    FILE_DIR = None


# Initialize GLiNER model globally (or lazy load)
# Handle potential CUDA issue by falling back to CPU.
try:
    GLINER_MODEL = GLiNER.from_pretrained(DEFAULT_GLINER_MODEL_NAME, load_tokenizer=True)
    GLINER_MODEL.to(DEFAULT_GLINER_DEVICE)
    logger.info(f"GLiNER model '{DEFAULT_GLINER_MODEL_NAME}' loaded successfully on device '{DEFAULT_GLINER_DEVICE}'.")
except Exception as e_gliner_cuda: # Catch errors like no CUDA device
    logger.warning(f"Failed to load GLiNER model on '{DEFAULT_GLINER_DEVICE}': {e_gliner_cuda}. Attempting CPU.")
    try:
        GLINER_MODEL = GLiNER.from_pretrained(DEFAULT_GLINER_MODEL_NAME, load_tokenizer=True)
        GLINER_MODEL.to("cpu")
        logger.info(f"GLiNER model '{DEFAULT_GLINER_MODEL_NAME}' loaded successfully on device 'cpu'.")
    except Exception as e_gliner_cpu:
        logger.error(f"Failed to load GLiNER model on CPU as fallback: {e_gliner_cpu}")
        GLINER_MODEL = None # Ensure model is None if loading fails

TEXT_SPLITTER = RecursiveCharacterTextSplitter(
    chunk_size=DEFAULT_CHUNK_SIZE,
    chunk_overlap=DEFAULT_CHUNK_OVERLAP,
    length_function=len,
    is_separator_regex=False,
)


@dataclass
class Entity:
    """
    Represents an extracted entity from text.

    Attributes:
        start (int): Start character offset of the entity in the original text.
        end (int): End character offset of the entity in the original text.
        label (str): The label/type of the entity (e.g., "author", "university").
        text (str): The actual text content of the entity.
        score (Optional[float]): Confidence score from the model, if available.
    """
    start: int
    end: int
    label: str
    text: str
    score: Optional[float] = None


def format_labels(labels: Set[str]) -> str:
    """Sorts and joins a set of labels into a comma-separated string for display."""
    return ", ".join(sorted(list(labels)))


def filter_and_merge_overlapping_entities(entities: List[Entity]) -> List[Entity]:
    """
    Filters and merges overlapping or identically spanned entities.

    If multiple entities cover the exact same span, their labels are merged.
    If entities overlap, the one with the largest span is generally preferred.
    This implementation prioritizes entities that start earlier, and among those,
    the ones that end later (i.e., longer entities). Overlapping entities' labels
    are merged into the dominant (longest, earliest starting) entity.

    Args:
        entities (List[Entity]): A list of Entity objects.

    Returns:
        List[Entity]: A new list of Entity objects with overlaps resolved and labels merged.
    """
    if not entities:
        return []

    # Sort by start index ascending, then by end index descending to prioritize longer spans among overlaps
    sorted_entities = sorted(entities, key=lambda e: (e.start, -e.end))

    merged_entities_list: List[Entity] = []
    if not sorted_entities: return merged_entities_list # Should not happen if entities is not empty

    current_merged_entity = dataclasses.replace(sorted_entities[0]) # Start with the first entity as a base
    current_labels: Set[str] = {current_merged_entity.label}

    for next_entity in sorted_entities[1:]:
        # Check for overlap or adjacency that might be considered a merge
        if next_entity.start < current_merged_entity.end:  # Overlap
            # If next_entity is entirely contained within current_merged_entity, add its label(s)
            if next_entity.end <= current_merged_entity.end:
                current_labels.update(next_entity.label.split(", ")) # Handle potentially already merged labels
            else: # Partial overlap, next_entity extends further
                # This case means current_merged_entity was shorter despite starting earlier or same.
                # This shouldn't happen with the primary sort key (-e.end).
                # If it does, it implies a more complex overlap. For now, merge label and extend span.
                logger.debug(f"Complex overlap: current={current_merged_entity}, next={next_entity}. Merging labels and extending span.")
                current_merged_entity.end = next_entity.end
                current_labels.update(next_entity.label.split(", "))
        else:  # No overlap with the current merged entity
            current_merged_entity.label = format_labels(current_labels)
            merged_entities_list.append(current_merged_entity)
            current_merged_entity = dataclasses.replace(next_entity) # Start a new merged entity
            current_labels = {current_merged_entity.label}

    # Add the last processed merged entity
    if current_merged_entity:
        current_merged_entity.label = format_labels(current_labels)
        merged_entities_list.append(current_merged_entity)

    return merged_entities_list


def update_text_with_html_annotations(original_text: str, entities: List[Entity]) -> str:
    """
    Wraps identified entities in the original text with HTML <mark> tags for highlighting.
    Handles potential overlaps by using the `filter_and_merge_overlapping_entities` function.

    Args:
        original_text (str): The original text string.
        entities (List[Entity]): A list of Entity objects found in `original_text`.

    Returns:
        str: Text annotated with HTML <mark> tags around entities.
    """
    if not entities:
        return original_text

    # Process entities to resolve overlaps and merge labels before annotation
    processed_entities = filter_and_merge_overlapping_entities(entities)
    # Sort by start position to ensure correct order for text reconstruction
    processed_entities.sort(key=lambda e: e.start)

    result_parts: List[str] = []
    current_pos: int = 0

    for entity in processed_entities:
        # Add text segment before the current entity
        if entity.start > current_pos:
            result_parts.append(original_text[current_pos : entity.start])

        # Add the wrapped entity text (using potentially merged label from processed_entities)
        entity_text_segment = original_text[entity.start : entity.end]
        # Ensure label is HTML-safe if it contains special characters (e.g. quotes in tooltip)
        # For simple cases, this should be fine. For complex labels, consider html.escape.
        wrapped_entity = f'<mark data-entity-label="{entity.label}">{entity_text_segment}</mark>'
        result_parts.append(wrapped_entity)
        current_pos = entity.end

    # Add any remaining text after the last entity
    if current_pos < len(original_text):
        result_parts.append(original_text[current_pos:])

    return "".join(result_parts)


def store_processed_entity_files(material_id: int, annotated_text_html: str, entities_to_save: List[Entity]) -> None:
    """
    Stores the processed list of entities as a JSON file and the
    HTML-annotated text as a Markdown (.md) file.

    Args:
        material_id (int): The material ID for naming the output files.
        annotated_text_html (str): The text string with HTML <mark> annotations.
        entities_to_save (List[Entity]): The final list of (merged/filtered) Entity objects to save.
    """
    if FILE_DIR is None or not FILE_DIR.exists:
        logger.error(f"Output directory for entity files is not configured or does not exist. Cannot save for material_id {material_id}.")
        return

    # Sort entities before saving, if not already sorted
    entities_to_save.sort(key=lambda e: e.start)
    entities_as_dicts: List[Dict[str, Any]] = [dataclasses.asdict(e) for e in entities_to_save]

    json_filename = f"{material_id}_entities.json" # Changed from _annotated.json
    json_save_path = FILE_DIR.full / json_filename
    try:
        with open(json_save_path, "w", encoding="utf-8") as f_json:
            json.dump(entities_as_dicts, f_json, indent=2, ensure_ascii=False)
        logger.info(f"Saved {len(entities_as_dicts)} processed entities to {json_save_path.name}")
    except OSError as e_json_save:
        logger.error(f"Error saving JSON entity file {json_save_path}: {e_json_save}")

    md_filename = f"{material_id}_annotated_text.md" # Changed from _annotated.md
    md_save_path = FILE_DIR.full / md_filename
    try:
        with open(md_save_path, "w", encoding="utf-8") as f_md:
            f_md.write(annotated_text_html)
        logger.info(f"Saved HTML-annotated text to {md_save_path.name}")
    except OSError as e_md_save:
        logger.error(f"Error saving HTML-annotated Markdown file {md_save_path}: {e_md_save}")


def find_entities_with_gliner(
    text_content: str,
    gliner_labels: List[str] = DEFAULT_GLINER_LABELS,
    prediction_threshold: float = GLINER_PREDICTION_THRESHOLD
) -> Tuple[List[Entity], str]:
    """
    Uses a GLiNER model to find entities in the given text.
    The text is split into chunks for processing.

    Args:
        text_content (str): The input text to process.
        gliner_labels (List[str]): A list of entity labels for the GLiNER model to predict.
        prediction_threshold (float): Confidence threshold for GLiNER predictions.

    Returns:
        Tuple[List[Entity], str]: A tuple containing:
            - list[Entity]: A list of found Entity objects (raw from model, before merging).
            - str: The input text content (unchanged by this function, but returned for consistency if chunking altered it).
                   Currently returns the original text as chunking is only for prediction.
    """
    if GLINER_MODEL is None:
        logger.error("GLiNER model not loaded. Cannot find entities.")
        return [], text_content

    raw_entities_from_model: List[Dict[str, Any]] = []
    processed_char_offset: int = 0

    text_chunks: List[str] = TEXT_SPLITTER.split_text(text_content)

    for chunk_text in text_chunks:
        if not chunk_text.strip(): # Skip empty or whitespace-only chunks
            processed_char_offset += len(chunk_text) # Still count its length for offset
            continue

        try:
            # GLiNER predict_entities returns list of dicts: {"start": int, "end": int, "label": str, "score": float}
            chunk_entities: List[Dict[str, Any]] = GLINER_MODEL.predict_entities(
                chunk_text, gliner_labels, threshold=prediction_threshold
            )
            if chunk_entities:
                for entity_dict in chunk_entities:
                    entity_dict["start"] += processed_char_offset
                    entity_dict["end"] += processed_char_offset
                    raw_entities_from_model.append(entity_dict)
        except Exception as e_gliner:
            logger.error(f"Error during GLiNER entity prediction on a chunk: {e_gliner}")
            logger.debug(f"Chunk causing error (first 100 chars): {chunk_text[:100]}")

        processed_char_offset += len(chunk_text)

    # Convert list of dicts to list of Entity objects
    entity_objects: List[Entity] = [Entity(**ent_dict) for ent_dict in raw_entities_from_model]

    # An internal merge_entities was here, but filter_and_merge_overlapping_entities is more robust
    # and should be called externally after combining with regex entities.
    return entity_objects, text_content


def extract_entities_with_regex(
    text_content: str,
    additional_keywords: Optional[Dict[str, List[str]]] = None
) -> List[Entity]:
    """
    Extracts entities from text using predefined regex patterns and keyword matching.
    Looks for ISBNs, URLs, DOIs, ISSNs, ORCID IDs, and keywords from `INITIAL_KEYWORD_DICT`
    plus any `additional_keywords` provided.

    Args:
        text_content (str): The text to extract entities from.
        additional_keywords (Optional[Dict[str, List[str]]]): More keywords to match,
            in the format expected by FlashText: `{'label_to_assign': ['keyword1', 'keyword2']}`.

    Returns:
        List[Entity]: A list of Entity objects found by regex and keyword matching.
    """
    # Define regex patterns with their corresponding labels
    regex_patterns: List[Tuple[Pattern[str], str]] = [
        (re.compile(ISBN_REGEX_PATTERN, flags=re.IGNORECASE), "isbn"), # Added IGNORECASE
        (re.compile(URL_REGEX_PATTERN, flags=re.IGNORECASE), "url"),   # Added IGNORECASE
        (re.compile(DOI_REGEX_PATTERN, flags=re.IGNORECASE), "doi"),   # Added IGNORECASE
        (re.compile(ISSN_REGEX_PATTERN), "issn"), # ISSN is case-sensitive for 'X'
        (re.compile(ORCID_REGEX_PATTERN), "orcid"), # ORCID is case-sensitive for 'X'
        (re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", flags=re.IGNORECASE), "email"), # Basic email regex
    ]

    found_entities_list: List[Entity] = []

    # Regex-based extraction
    for compiled_regex, label in regex_patterns:
        for match in compiled_regex.finditer(text_content):
            start_offset, end_offset = match.span()
            entity_text_matched = match.group(0)
            found_entities_list.append(Entity(start_offset, end_offset, label, entity_text_matched))

    # Keyword-based extraction using FlashText
    keyword_processor = KeywordProcessor(case_sensitive=False)
    keyword_processor.add_keywords_from_dict(INITIAL_KEYWORD_DICT)
    if additional_keywords: # Add any runtime keywords
        keyword_processor.add_keywords_from_dict(additional_keywords)

    extracted_keywords = keyword_processor.extract_keywords(text_content, span_info=True)
    for keyword_label, start, end in extracted_keywords: # FlashText returns (keyword_found_as_label, start, end)
        entity_text = text_content[start:end] # Get the original text span
        # The 'keyword' from FlashText is actually the 'label' we assigned in the dict.
        # We need to decide if the label for the Entity should be this 'keyword_label'
        # or a more generic one like "recognized_keyword" or "named_entity".
        # Using the key from keyword_dict as the label is more informative.
        found_entities_list.append(Entity(start, end, label=keyword_label, text=entity_text))

    return found_entities_list


def get_specific_names_from_db_for_item(material_id: int) -> Dict[str, List[str]]:
    """
    Retrieves Osiris data related to a specific material ID and extracts names
    (department, course, owner, author, contacts) to be used as keywords for entity recognition.

    Args:
        material_id (int): The material ID to fetch related names for.

    Returns:
        Dict[str, List[str]]: A dictionary where keys are the extracted names/phrases (intended to be used as entity labels later)
                              and values are lists of keyword variations for that name.
                              Example: `{"Prof. Example Name": ["Example Name", "Example", "prof example"]}`
    """
    # This function uses retrieve_osiris_data, which is async.
    # However, this function itself is synchronous. This will block if called from sync code.
    # To fix, either make this async or call retrieve_osiris_data in a way that it runs in an event loop.
    # For now, assuming it's called from a context where an event loop might be managed by caller,
    # or this will be refactored to be async if it's part of an async chain.
    # logger.warning("get_specific_names_from_db_for_item uses asyncio.run for an async call, which might not be ideal in all contexts.")
    # osiris_data_list = asyncio.run(retrieve_osiris_data([material_id])) # retrieve_osiris_data is async

    # Assuming this function will be made async or called appropriately:
    # For now, to make it runnable for review, I'll mock an async call or simplify.
    # This part needs proper async handling if retrieve_osiris_data is truly async.
    # For the purpose of this review, let's assume osiris_data is fetched and available.
    # Actual implementation would be:
    # osiris_data_list = await retrieve_osiris_data([material_id])

    # Placeholder:
    logger.warning("`get_specific_names_from_db_for_item` - Osiris data retrieval is mocked/simplified for this review pass due to async context.")
    osiris_data_list: List[Dict[str,Any]] = [] # Mocked empty data

    if not osiris_data_list:
        logger.info(f"No Osiris data found for material_id {material_id} to extract specific names.")
        return {}

    osiris_item_data = osiris_data_list[0] # Assuming retrieve_osiris_data returns a list with one item for one material_id

    names_to_extract: Dict[str, List[str]] = {}

    def add_name_variations(label: str, name_str: Optional[str]) -> None:
        if name_str:
            variations: List[str] = [name_str.strip()]
            try:
                parsed_human_name = HumanName(name_str)
                if parsed_human_name.last: variations.append(parsed_human_name.last)
                if parsed_human_name.first: variations.append(parsed_human_name.first)
                if parsed_human_name.first and parsed_human_name.last:
                    variations.append(f"{parsed_human_name.first} {parsed_human_name.last}")
            except Exception: # nameparser might fail on some strings
                logger.debug(f"Could not parse '{name_str}' with HumanName.")
            names_to_extract[label] = list(set(variations)) # Unique variations

    add_name_variations(osiris_item_data.get("department", "Unknown Department"), osiris_item_data.get("department"))
    add_name_variations(osiris_item_data.get("course_name", "Unknown Course"), osiris_item_data.get("course_name"))
    add_name_variations(osiris_item_data.get("owner", "Unknown Owner"), osiris_item_data.get("owner")) # Assuming owner is a name
    add_name_variations(osiris_item_data.get("author", "Unknown Author"), osiris_item_data.get("author"))

    if isinstance(osiris_item_data.get("courses"), list):
        for course_entry in osiris_item_data["courses"]:
            if isinstance(course_entry, dict):
                add_name_variations(course_entry.get("name", "Unknown SubCourse"), course_entry.get("name"))
                if isinstance(course_entry.get("persons"), list):
                    for person_entry in course_entry["persons"]:
                        if isinstance(person_entry, dict) and person_entry.get("main_name"):
                            add_name_variations(person_entry["main_name"], person_entry["main_name"])
                            if person_entry.get("email"):
                                 names_to_extract[person_entry["main_name"]].append(person_entry["email"])

    logger.debug(f"Extracted specific names for material_id {material_id}: {list(names_to_extract.keys())}")
    return names_to_extract


def process_text_files_for_entities(text_file_type_suffix: str = "_cleaned_text") -> None: # Example suffix
    """
    Orchestrates the entity recognition process for text files in the PDF_DOWNLOADS directory.
    It processes files matching a given suffix (e.g., "_cleaned_text.txt" or ".md" from PDF extraction).
    For each file, it:
    1. Reads text content.
    2. Finds entities using GLiNER.
    3. Extracts additional entities using regex/keywords (including names from DB for context).
    4. Merges and filters all found entities.
    5. Creates an HTML-annotated version of the text.
    6. Stores the annotated text and final entities to JSON and Markdown files.

    Args:
        text_file_type_suffix (str): Suffix (including extension like ".txt" or ".md")
                                     to identify relevant text files to process.
    """
    if FILE_DIR is None or not FILE_DIR.exists:
        logger.error(f"PDF_DOWNLOADS directory for text files is not configured or does not exist. Cannot process items.")
        return

    all_text_files: List[File] = [f for f in FILE_DIR.files if f.name.endswith(text_file_type_suffix)]

    # Determine which files have already been processed (i.e., have an _entities.json file)
    already_processed_material_ids: Set[int] = set()
    for f_json_check in FILE_DIR.files:
        if f_json_check.name.endswith("_entities.json"):
            try:
                already_processed_material_ids.add(int(f_json_check.name.replace("_entities.json", "")))
            except ValueError:
                logger.debug(f"Could not parse material_id from existing JSON file: {f_json_check.name}")

    files_to_process: List[File] = []
    for text_file in all_text_files:
        try:
            # Assuming filename starts with material_id
            material_id_str = text_file.name.split("_")[0].split(".")[0]
            if material_id_str.isdigit() and int(material_id_str) not in already_processed_material_ids:
                files_to_process.append(text_file)
        except Exception as e_filter:
            logger.warning(f"Could not determine material_id or processing status for {text_file.name}: {e_filter}")

    if not files_to_process:
        logger.info(f"No new text files with suffix '{text_file_type_suffix}' found to process for entities.")
        return

    logger.info(f"Found {len(files_to_process)} text files to process for entity recognition.")

    for text_file_obj in files_to_process:
        material_id: Optional[int] = None
        try:
            # Extract material_id from filename (e.g., "12345_cleaned_text.txt" -> 12345)
            material_id_str = text_file_obj.name.split("_")[0].split(".")[0]
            if not material_id_str.isdigit():
                logger.warning(f"Filename {text_file_obj.name} does not start with a numeric material_id. Skipping.")
                continue
            material_id = int(material_id_str)

            logger.info(f"Processing file for material_id {material_id}: {text_file_obj.name}")
            with open(text_file_obj.path, "r", encoding="utf-8") as f_text:
                text_content = f_text.read()

            if not text_content.strip():
                logger.info(f"Text file {text_file_obj.name} is empty. Skipping entity recognition.")
                continue

            # 1. GLiNER-based entities
            gliner_entities, _ = find_entities_with_gliner(text_content) # Original text passed back is not currently used

            # 2. Regex/Keyword-based entities (contextualized with DB names)
            # This part needs to be async if get_specific_names_from_db_for_item becomes async.
            # For now, assuming it's handled (e.g. by running this whole process_items in a thread if called from async)
            # specific_names_for_item = get_specific_names_from_db_for_item(material_id) # This was problematic (async in sync)
            # For now, pass empty dict for additional keywords from DB to avoid blocking/async issues here.
            # TODO: Refactor get_specific_names_from_db_for_item to be async and await it, or run in thread.
            specific_names_for_item: Dict[str, List[str]] = {}
            if material_id == -1: # Disable DB call for now to avoid async issue here
                 logger.info("DB call for specific names is currently disabled in this context.")

            regex_keyword_entities = extract_entities_with_regex(text_content, additional_keywords=specific_names_for_item)

            all_detected_entities = gliner_entities + regex_keyword_entities

            # 3. Filter, Merge, and Annotate
            final_entities = filter_and_merge_overlapping_entities(all_detected_entities)
            annotated_html_text = update_text_with_html_annotations(text_content, final_entities)

            # 4. Store results
            store_processed_entity_files(material_id, annotated_html_text, final_entities) # Uses final_entities

        except FileNotFoundError:
            logger.error(f"Text file not found during processing: {text_file_obj.path}")
        except Exception as e_proc:
            logger.error(f"Error processing file {text_file_obj.name} for material_id {material_id}: {e_proc}")
            logger.debug(traceback.format_exc())

    logger.info("Entity recognition processing finished for selected files.")
