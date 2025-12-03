import textwrap
from collections import Counter, defaultdict

from gliner import GLiNER

from easy_access.db.models import PDF


async def gliner_entity_recognition(model: GLiNER, labels: list[str], pdf: PDF):
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
                current["text"] = full_text[
                    current["start"] : next_entity["end"]
                ].strip()
                current["end"] = next_entity["end"]
            else:
                merged.append(current)
                current = next_entity
        # Append the last entity
        merged.append(current)
        return merged

    if pdf.extracted_text:
        text = pdf.extracted_text.extracted_text
    truncated_texts = []
    for text_batch in textwrap.wrap(text, width=380):  # split str into batches
        truncated_texts.append(text_batch)

    entities = model.run(texts=truncated_texts, labels=labels)
    entities_w_text = zip(entities, truncated_texts, strict=True)
    entities_w_text = [e for e in entities_w_text if e[0]]
    entities_merged = [merge_entities(e, text) for e, text in entities_w_text]
    print(f"Extracted {len(entities_merged)} entities from OCR text blocks.")

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
    {
        label: {
            "count": count,
            "avg_score": avg_scores_per_label.get(label, 0),
            "entities": list(label_to_entities.get(label, [])),
        }
        for label, count in label_counts.items()
    }
    {
        entity: {
            "count": count,
            "avg_score": avg_scores_per_entity.get(entity, 0),
            "label": entity_to_label.get(entity, ""),
        }
        for entity, count in entity_counts.items()
    }

    # match found entities with known entities

    # store data in db for entities
