# holds various models used for pdf processing etc

from dataclasses import dataclass


@dataclass
class Entity:
    """
    Dataclass to hold a single extracted entity.
    """

    text: str  # the actual text of the entity
    label: str  # the label/type of the entity
    entity_name: str  # the actual name of the entity that was matched with this text,
    start_char: int  # start char index (inclusive)
    end_char: int  # end char index (exclusive)
    confidence: float  # confidence score of the extraction (0-1)
    page: int  # page number where the entity was found in the PDF (1-based)


@dataclass
class ExtractedEntities:
    """
    Dataclass to hold extracted entities from a PDF,
    with various useful methods for grouping, filtering, displaying in the original text, etc.
    """

    entities: list[Entity]

    def __post_init__(self: "ExtractedEntities"):
        """
        Dynamically creates properties for each unique label in the entities list.
        E.g. if there are entities with labels `person` or `org`, this will create properties
        'person' and 'org' that return lists of entities with those labels.

        Example usage: pdf.entities.person --> returns list of entities with label `person`
        """
        unique_labels = set(entity.label for entity in self.entities)
        for label in unique_labels:
            setattr(
                self,
                label.lower(),
                property(
                    lambda self, lbl=label: [e for e in self.entities if e.label == lbl]
                ),
            )

    @classmethod
    def from_json(cls, data: list[dict]) -> "ExtractedEntities":
        """
        Creates an ExtractedEntities object from a list of dicts (as stored in the db).
        This is the canonical way to create an ExtractedEntities object from the db.
        """
        entities = [Entity(**item) for item in data]
        return cls(entities=entities)
