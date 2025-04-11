# get extracted text
# if .txt, turn into .md first
# use gliner to find entities in extracted text
# wrap each entity in html tags to highlight them and include the entity type
# then for each file create a json with the entities, types, and positions

import json

from gliner import GLiNER
from langchain_text_splitters import RecursiveCharacterTextSplitter
from rich import print

from easy_access.settings import SETTINGS, DirSetting

FILE_DIR = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]

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
) -> list[dict]:
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
    annotated_text = update_text(new_text, merged) if merged else new_text

    return merged, annotated_text


def update_text(text: str, entities: list[dict]) -> str:
    # input markdown text and a list of entities found in the text
    # wrap each entity in html tags to highlight them and include the entity type
    # return the updated markdown text

    # problem: when we add html tags, the start and end positions of the entities change
    # solution: start from the end of the text and move backwards, so that the start and end positions of the entities do not change

    # first sort the entities by their start position
    entities.sort(key=lambda x: x["start"])
    # then reverse the order of the entities, so that we can update the text from the end
    entities.reverse()
    for entity in entities:
        start = entity["start"]
        end = entity["end"]
        label = entity["label"]
        entity_text = text[start:end]
        # wrap the entity in html tags
        update = f'<span class="entity_{label} tooltip tooltip-primary" data-tip="{label}"><mark class="{label}">{entity_text}</mark></span>'
        text = text[:start] + update + text[end:]

    return text


def store_files(material_id: int, annotated_text: str, entities: list[dict]) -> None:
    entities_file = f"{material_id}_annotated.json"
    save_path = FILE_DIR.full / entities_file
    with open(save_path, "w", encoding="utf-8") as f:
        json.dump(entities, f)

    annotated_text_file = f"{material_id}_annotated.md"
    save_path = FILE_DIR.full / annotated_text_file
    with open(save_path, "w", encoding="utf-8") as f:
        f.write(annotated_text)

    print(f"Saved annotated text and entities for material ID {material_id}")


def regex_extraction(
    text: str, specific_entitites: list[str] | None = None
) -> list[dict]:
    """
    use regex to extract specific entities:
    - email + url
    - dois
    - isbns
    - issns
    - orcid ids
    - exact matches for names
    """
    isbn_regex = (
        r"(ISBN[-]*(1[03])*[ ]*(: ){0,1})*(([0-9Xx][- ]*){13}|([0-9Xx][- ]*){10})"
    )

    url_regex = r"^(?:(?:http|https|ftp|telnet|gopher|ms\-help|file|notes)://)?(?:(?:[a-z][\w~%!&amp;',;=\-\.$\(\)\*\+]*):.*@)?(?:(?:[a-z0-9][\w\-]*[a-z0-9]*\.)*(?:(?:(?:(?:[a-z0-9][\w\-]*[a-z0-9]*)(?:\.[a-z0-9]+)?)|(?:(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)))(?::[0-9]+)?))?(?:(?:(?:/(?:[\w`~!$=;\-\+\.\^\(\)\|\{\}\[\]]|(?:%\d\d))+)*/(?:[\w`~!$=;\-\+\.\^\(\)\|\{\}\[\]]|(?:%\d\d))*)(?:\?[^#]+)?(?:#[a-z0-9]\w*)?)?$"
    doi_regex = r"^(10\.\d{4,5}\/[\S]+[^;,.\s])$"

    issn_regex = r"\d{4}-\d{3}(\d|x|X)"
    orcid_regex = r"^(\d{4}-){3}\d{3}(\d|X)$"

    specific_names = []  # list of specific names to match -- publishers, employees, faculties, institutes, etc
    # join with specific_entities if provided
    if specific_entitites:
        if isinstance(specific_entitites, str):
            specific_entitites = [specific_entitites]
        specific_names.extend(specific_entitites)

    entities = []

    # perform search for each type
    # if match is found, add to entities list in the format:
    # {"start": start, "end": end, "label": label, "text": text}
    return entities


def process_items() -> None:
    all_extracted_text_files = [
        f
        for f in FILE_DIR.files
        if f.extension in [".md", ".txt"] and "annotated" not in f.name
    ]

    for file in all_extracted_text_files:
        material_id = int(file.name.split(".")[0])

        with open(file.path, encoding="utf-8") as f:
            text = f.read()

        json_file = f"{material_id}_annotated.json"
        json_path = FILE_DIR.full / json_file
        if json_path.exists():
            with open(json_path, encoding="utf-8") as f:
                merged = json.load(f)
        else:
            merged = None
        # remove excessive whitespace & newlines
        text = text.replace("\n\n", "\n")
        text = " ".join(text.split())
        if merged:
            entities, annotated_text = find_entities(
                text, skip_extraction=True, merged=merged
            )
        else:
            entities, annotated_text = find_entities(text, skip_extraction=False)
        if entities:
            store_files(material_id, annotated_text, entities)
