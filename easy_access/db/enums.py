from dataclasses import dataclass
from enum import Enum

"""
Mapping notes:

- Lange overname zou moeten mappen naar --> Ja, anders volgens SURF, denk niet dat dat klopt!
- V1 open access kan ook publiek domein zijn in v2
- studentwerk is in v1 niet apart, in v2 wel -- vaak als eigen werk gemarkeerd

"""


class Classification(Enum):
    OPEN_ACCESS = "open access"
    KORTE_OVERNAME = "korte overname"
    MIDDELLANGE_OVERNAME = "middellange overname"
    LANGE_OVERNAME = "lange overname"

    EIGEN_MATERIAAL_POWERPOINT = "eigen materiaal - powerpoint"
    EIGEN_MATERIAAL_TITELINDICATIE = "eigen materiaal - titelindicatie"
    EIGEN_MATERIAAL_OVERIG = "eigen materiaal - overig"
    EIGEN_MATERIAAL = "eigen materiaal"

    ONBEKEND = "onbekend"  # default for unclassified items
    NIET_GEANALYSEERD = "niet geanalyseerd"
    IN_ONDERZOEK = "in onderzoek"
    VERWIJDERVERZOEK_VERSTUURD = "verwijderverzoek verstuurd"
    LICENTIE_BESCHIKBAAR = "licentie beschikbaar"


# V2 classification system + mapping + notes


class ClassificationV2(Enum):
    """
    The new classification system for V2 of the copyright tool.
    """

    # Yes classifications
    JA_OPEN_LICENTIE = "Ja (open licentie)"
    JA_BIBLIOTHEEK_LICENTIE = "Ja (bibilotheek licentie)"
    JA_DIRECTE_TOESTEMMING = "Ja (directe toestemming)"
    JA_PUBLIEK_DOMEIN = "Ja (Publiek domein)"
    JA_EIGEN_WERK = "Ja (eigen werk)"
    JA_STUDENTWERK = "Ja (studentwerk)"
    JA_EASY_ACCESS = "Ja (easy access)"
    JA_ANDERS = "Ja (anders)"

    JA_DIRECTE_TOESTEMMING_TIJDELIJK = "Ja (directe toestemming) - tijdelijk"
    JA_BIBLIOTHEEK_LICENTIE_TIJDELIJK = "Ja (bibilotheek licentie)- tijdelijk"
    JA_ANDERS_TIJDELIJK = "Ja (anders) - tijdelijk"

    # No classifications
    NEE_LINK_BESCHIKBAAR = "Nee (Link beschikbaar)"
    NEE_STUDENTWERK = "Nee (studentwerk)"
    NEE = "Nee"

    # Other classifications // default to this when not classified
    ONBEKEND = "Onbekend"


class OvernameStatus(Enum):
    OVERNAME_INBREUKMAKENDE = "Overname (inbreukmakende)"
    OVERNAME_ANDERE = "Overname (andere)"
    GEEN_OVERNAME = "Geen overname"
    ONBEKEND = "Onbekend"


class Lengte(Enum):
    KORT = "Kort"
    MIDDELLANG = "Middellang"
    LANG = "Lang"
    ONBEKEND = "Onbekend"


@dataclass
class ClassificationMapping:
    classification: ClassificationV2
    overname_status: OvernameStatus
    length: Lengte


CLASSIFICATION_MAPPING_V1_TO_V2: dict[Classification, ClassificationMapping] = {
    Classification.OPEN_ACCESS: ClassificationMapping(
        classification=ClassificationV2.JA_OPEN_LICENTIE,
        overname_status=OvernameStatus.GEEN_OVERNAME,
        length=Lengte.ONBEKEND,
    ),
    Classification.EIGEN_MATERIAAL: ClassificationMapping(
        classification=ClassificationV2.JA_EIGEN_WERK,
        overname_status=OvernameStatus.GEEN_OVERNAME,
        length=Lengte.ONBEKEND,
    ),
    Classification.EIGEN_MATERIAAL_OVERIG: ClassificationMapping(
        classification=ClassificationV2.JA_EIGEN_WERK,
        overname_status=OvernameStatus.GEEN_OVERNAME,
        length=Lengte.ONBEKEND,
    ),
    Classification.EIGEN_MATERIAAL_POWERPOINT: ClassificationMapping(
        classification=ClassificationV2.JA_EIGEN_WERK,
        overname_status=OvernameStatus.GEEN_OVERNAME,
        length=Lengte.ONBEKEND,
    ),
    Classification.EIGEN_MATERIAAL_TITELINDICATIE: ClassificationMapping(
        classification=ClassificationV2.JA_EIGEN_WERK,
        overname_status=OvernameStatus.GEEN_OVERNAME,
        length=Lengte.ONBEKEND,
    ),
    Classification.KORTE_OVERNAME: ClassificationMapping(
        classification=ClassificationV2.JA_EASY_ACCESS,
        overname_status=OvernameStatus.OVERNAME_ANDERE,
        length=Lengte.KORT,
    ),
    Classification.MIDDELLANGE_OVERNAME: ClassificationMapping(
        classification=ClassificationV2.JA_EASY_ACCESS,
        overname_status=OvernameStatus.OVERNAME_ANDERE,
        length=Lengte.MIDDELLANG,
    ),
    Classification.LANGE_OVERNAME: ClassificationMapping(
        classification=ClassificationV2.NEE,
        overname_status=OvernameStatus.OVERNAME_INBREUKMAKENDE,
        length=Lengte.LANG,
    ),
    Classification.ONBEKEND: ClassificationMapping(
        classification=ClassificationV2.ONBEKEND,
        overname_status=OvernameStatus.ONBEKEND,
        length=Lengte.ONBEKEND,
    ),
    Classification.NIET_GEANALYSEERD: ClassificationMapping(
        classification=ClassificationV2.ONBEKEND,
        overname_status=OvernameStatus.ONBEKEND,
        length=Lengte.ONBEKEND,
    ),
    Classification.IN_ONDERZOEK: ClassificationMapping(
        classification=ClassificationV2.ONBEKEND,
        overname_status=OvernameStatus.ONBEKEND,
        length=Lengte.ONBEKEND,
    ),
    Classification.VERWIJDERVERZOEK_VERSTUURD: ClassificationMapping(
        classification=ClassificationV2.ONBEKEND,
        overname_status=OvernameStatus.OVERNAME_INBREUKMAKENDE,
        length=Lengte.ONBEKEND,
    ),
    Classification.LICENTIE_BESCHIKBAAR: ClassificationMapping(
        classification=ClassificationV2.NEE_LINK_BESCHIKBAAR,
        overname_status=OvernameStatus.OVERNAME_INBREUKMAKENDE,
        length=Lengte.ONBEKEND,
    ),
}


class Filetype(Enum):
    PDF = "pdf"
    PPT = "ppt"
    DOC = "doc"
    XLSX = "xlsx"
    MP4 = "mp4"
    JPG = "jpg"
    PNG = "png"
    UNKNOWN = "unknown"
    FILE = "file"


class Status(Enum):
    PUBLISHED = "Published"
    UNPUBLISHED = "Unpublished"
    DELETED = "Deleted"


class WorkflowStatus(Enum):
    ToDo = "ToDo"
    Done = "Done"
    InProgress = "InProgress"


class Infringement(Enum):
    YES = "yes"
    NO = "no"
    MAYBE = "maybe"
    UNDETERMINED = "undetermined"


"""
Programatically generate enums for years between 2020 and 2030 for valid periods using one of these formats:
YYYY-[12]{1}[AB]{1} (eg. 2022-1A or 2022-2B)
YYYY-3 (eg. 2022-3)
YYYY-SEM[12]{1} (eg. 2022-SEM1 or 2022-SEM2)
YYYY-JAAR (eg. 2022-JAAR)
"""
Period = Enum(
    "Period",
    {
        f"{year}_{period}": f"{year}-{period}"
        for year in range(2020, 2031)
        for period in ["1A", "1B", "2A", "2B", "3", "SEM1", "SEM2", "JAAR"]
    },
)
