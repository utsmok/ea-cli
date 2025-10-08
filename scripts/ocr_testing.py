import json
import textwrap
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

import numpy
import pdf2image
from gliner import GLiNER
from kreuzberg import extract_file_sync
from paddleocr import PaddleOCR
from rich import print


def merge_entities(entities, full_text):
    if not entities:
        return []
    merged = []
    current = entities[0]
    for next_entity in entities[1:]:
        if next_entity["label"] == current["label"] and (
            next_entity["start"] == current["end"] + 1
            or next_entity["start"] == current["end"]
        ):
            current["text"] = full_text[current["start"] : next_entity["end"]].strip()
            current["end"] = next_entity["end"]
        else:
            merged.append(current)
            current = next_entity
    # Append the last entity
    merged.append(current)
    return merged


@dataclass
class OCRConfig:
    ocr: PaddleOCR
    gliner_model: GLiNER
    labels: list[str]
    MAX_PAGES: int
    result_dir: Path
    pdf_dir: Path


def init_ocr() -> OCRConfig:
    ocr = PaddleOCR(
        lang="en",
        device="gpu",
        use_doc_unwarping=False,
    )
    gliner_model = GLiNER.from_pretrained("numind/NuNerZero").to("cuda")
    labels = ["person", "author", "publisher", "copyright_holder", "organization"]
    MAX_PAGES = 5  # max pages to process from each pdf

    result_dir = Path("ocr_results")
    result_dir.mkdir(exist_ok=True)
    pdf_dir = Path("pdf_downloads")

    return OCRConfig(
        ocr=ocr,
        gliner_model=gliner_model,
        labels=labels,
        MAX_PAGES=MAX_PAGES,
        result_dir=result_dir,
        pdf_dir=pdf_dir,
    )


def run_ocr(config: OCRConfig):
    for pdf_path in config.pdf_dir.glob("*.pdf"):
        stored_text_path = str(config.result_dir / f"ocr_result_{pdf_path.name}.txt")
        json_path = str(config.result_dir / f"ocr_entities_{pdf_path.name}.json")

        texts: list[str] = []
        entities = []
        if Path(stored_text_path).exists():
            print(f"Loading stored text from {stored_text_path}")
            # read in the lines from the file
            # append to texts list if not empty
            with open(stored_text_path, encoding="utf-8") as f:
                text = f.read()
                for line in text.split("\n"):
                    line = line.strip()
                    if line:
                        texts.append(line)
            processed_pages = "loaded from disk"
        else:
            # first try direct text extraction w/ kreuzberg
            # if that fails move to ocr
            try:
                res = extract_file_sync(pdf_path)
                extracted_text = res.content
                if extracted_text and extracted_text.strip():
                    # limit to 5k characters
                    if len(extracted_text) > 5000:
                        extracted_text = extracted_text.strip()[:5000]
                    texts.append(extracted_text)
                    processed_pages = "extracted via kreuzberg, limited to 5k chars"
            except Exception as e:
                print(f"Error extracting text from {pdf_path} with kreuzberg: {e}")
                texts = []
        if not texts:
            print(f"Running OCR on {pdf_path}")
            # first we use pdf2image to convert the first 5 pages to images
            # store temp images on disk
            # pass the list of image paths to paddle ocr and run ocr
            # then delete tmp images
            images = pdf2image.convert_from_path(
                str(pdf_path.absolute()),
                first_page=1,
                last_page=config.MAX_PAGES,
            )
            print(f"got {len(images)} images from pdf2image for {pdf_path}.")
            if len(images) == 0:
                print(f"No images extracted from {pdf_path}, stopping.")
                continue
            processed_pages = len(images)
            images = [numpy.asarray(i) for i in images]
            results = config.ocr.predict(images)
            print(f"got {len(results)} results from OCR processing of {pdf_path}.")
            if not results:
                print(f"No OCR results for {pdf_path}, stopping.")
                break
            for res in results:
                text = " ".join(res.get("rec_texts", []))
                texts.append(text)
            with open(stored_text_path, "w", encoding="utf-8") as f:
                for text in texts:
                    f.write(text + "\n")
            print(f"Extracted {len(texts)} text blocks from OCR results. ")

        truncated_texts = []
        for text in texts:
            if not text.strip():
                continue
            text = text.strip()

            for text_batch in textwrap.wrap(text, width=380):  # split str into batches
                truncated_texts.append(text_batch)

        entities = config.gliner_model.run(texts=truncated_texts, labels=config.labels)
        entities_w_text = zip(entities, truncated_texts, strict=True)
        entities_w_text = [e for e in entities_w_text if e[0]]
        entities_merged = [merge_entities(e, text) for e, text in entities_w_text]
        print(f"Extracted {len(entities_merged)} entities from OCR text blocks.")
        # print counts per label (entity["label"])
        label_counts = Counter()
        entity_counts = Counter()
        scores_per_label = defaultdict(list)
        scores_per_entity = defaultdict(list)
        entity_to_label = {}
        label_to_entities = defaultdict(set)
        for entity_batch in entities_merged:
            for entity in entity_batch:
                print(entity)
                try:
                    label_counts[entity["label"]] += 1
                    entity_counts[entity["text"]] += 1
                    scores_per_label[entity["label"]].append(entity["score"])
                    scores_per_entity[entity["text"]].append(entity["score"])
                    entity_to_label[entity["text"]] = entity["label"]
                    label_to_entities[entity["label"]].add(entity["text"])
                except Exception:
                    print(f"Error counting label for entity: {entity}")
                    break

        avg_scores_per_label = {
            label: sum(scores) / len(scores) if scores else 0
            for label, scores in scores_per_label.items()
        }
        avg_scores_per_entity = {
            entity: sum(scores) / len(scores) if scores else 0
            for entity, scores in scores_per_entity.items()
        }

        # merge counts and average scores into single dict for labels and entities
        label_results = {
            label: {
                "count": count,
                "avg_score": avg_scores_per_label.get(label, 0),
                "entities": list(label_to_entities.get(label, [])),
            }
            for label, count in label_counts.items()
        }
        entity_results = {
            entity: {
                "count": count,
                "avg_score": avg_scores_per_entity.get(entity, 0),
                "label": entity_to_label.get(entity, ""),
            }
            for entity, count in entity_counts.items()
        }
        # sort both by count descending
        label_results = dict(
            sorted(
                label_results.items(), key=lambda item: item[1]["count"], reverse=True
            )
        )
        entity_results = dict(
            sorted(
                entity_results.items(), key=lambda item: item[1]["count"], reverse=True
            )
        )

        print("per label:")
        print(label_results)

        print("per entity:")
        print(entity_results)

        # store the raw and processed data in a json file
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "settings": {
                        "pdf_path": str(pdf_path),
                        "stored_text_path": stored_text_path,
                        "num_pages_processed": processed_pages,
                        "ocr_model": "PaddleOCR",
                        "entity_model": "numind/NuNerZero",
                    },
                    "results": {
                        "by_label": label_results,
                        "by_entity": entity_results,
                    },
                    "truncated_texts": truncated_texts,
                    "raw_entities": entities_merged,
                },
                f,
                indent=2,
            )

        # check number of files in result_dir
        num_files = len(list(config.result_dir.glob("*.json")))
        print(f"Processed {num_files / 2} pdf files.")


config = init_ocr()
run_ocr(config)
