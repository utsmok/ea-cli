"""
Merge rules for copyright items.

This module contains the field definitions and comparison logic for merging copyright items.
Extracted from legacy update_copyright_items logic.
"""

from easy_access.db.models import (
    Classification,
    ClassificationV2,
    Infringement,
    Lengte,
    OvernameStatus,
    WorkflowStatus,
)
from easy_access.settings import Settings

# Fields that can be added/updated with priorities (list for ordered preferences, None for no priority)
added_fields = {
    "workflow_status": [
        WorkflowStatus.Done.value,
        WorkflowStatus.InProgress.value,
        WorkflowStatus.ToDo.value,
    ],
    "retrieved_from_copyright_on": None,
    "possible_fine": None,
    "infringement": [
        Infringement.YES.value,
        Infringement.NO.value,
        Infringement.UNDETERMINED.value,
    ],
    "file_exists": [False, 0, True, 1],
}

# Fields changeable by checkers with priorities
changeable_fields = {
    "v2_manual_classification": [
        ClassificationV2.JA_OPEN_LICENTIE.value,
        ClassificationV2.JA_PUBLIEK_DOMEIN.value,
        ClassificationV2.JA_DIRECTE_TOESTEMMING.value,
        ClassificationV2.JA_EASY_ACCESS.value,
        ClassificationV2.JA_EIGEN_WERK.value,
        ClassificationV2.JA_STUDENTWERK.value,
        ClassificationV2.JA_BIBLIOTHEEK_LICENTIE.value,
        ClassificationV2.JA_ANDERS.value,
        ClassificationV2.JA_DIRECTE_TOESTEMMING_TIJDELIJK.value,
        ClassificationV2.JA_BIBLIOTHEEK_LICENTIE_TIJDELIJK.value,
        ClassificationV2.JA_ANDERS_TIJDELIJK.value,
        ClassificationV2.NEE_LINK_BESCHIKBAAR.value,
        ClassificationV2.NEE_STUDENTWERK.value,
        ClassificationV2.NEE.value,
        ClassificationV2.ONBEKEND.value,
    ],
    "v2_lengte": [
        Lengte.LANG.value,
        Lengte.MIDDELLANG.value,
        Lengte.KORT.value,
        Lengte.ONBEKEND.value,
    ],
    "v2_overnamestatus": [
        OvernameStatus.OVERNAME_INBREUKMAKENDE.value,
        OvernameStatus.OVERNAME_ANDERE.value,
        OvernameStatus.GEEN_OVERNAME.value,
        OvernameStatus.ONBEKEND.value,
    ],
    "manual_classification": [
        Classification.OPEN_ACCESS.value,
        Classification.EIGEN_MATERIAAL_POWERPOINT.value,
        Classification.EIGEN_MATERIAAL_TITELINDICATIE.value,
        Classification.EIGEN_MATERIAAL_OVERIG.value,
        Classification.EIGEN_MATERIAAL.value,
        Classification.LICENTIE_BESCHIKBAAR.value,
        Classification.LANGE_OVERNAME.value,
        Classification.MIDDELLANGE_OVERNAME.value,
        Classification.KORTE_OVERNAME.value,
        Classification.VERWIJDERVERZOEK_VERSTUURD.value,
        Classification.IN_ONDERZOEK.value,
        Classification.NIET_GEANALYSEERD.value,
        Classification.ONBEKEND.value,
    ],
    "manual_identifier": None,
    "remarks": None,
    "scope": None,
}


def build_merge_rules_from_settings(settings: Settings) -> tuple[dict, dict]:
    """
    Build merge rules dynamically from Settings.

    Args:
        settings: Settings instance with field definitions

    Returns:
        Tuple of (added_fields, changeable_fields) dictionaries
    """
    # Start with default added_fields
    dynamic_added_fields = added_fields.copy()

    # Start with default changeable_fields
    dynamic_changeable_fields = changeable_fields.copy()

    # Update manual_classification priorities from settings.classification_options
    if settings.classification_options:
        # Map settings classifications to Classification enum values
        classification_priorities = []
        for option in settings.classification_options:
            # Try to match with Classification enum values (case-insensitive)
            for cls in Classification:
                if cls.value.lower() == option.lower():
                    classification_priorities.append(cls.value)
                    break

        if classification_priorities:
            dynamic_changeable_fields["manual_classification"] = (
                classification_priorities
            )

    # Update workflow_status priorities from data_entry_cols dropdown
    if (
        settings.data_settings
        and hasattr(settings.data_settings, "data_entry_cols")
        and settings.data_settings.data_entry_cols
    ):
        for col_info in settings.data_settings.data_entry_cols:
            if col_info.name == "workflow_status" and col_info.dropdown_options:
                # Parse dropdown options (format: '"ToDo,Done,InProgress"')
                options_str = col_info.dropdown_options.strip()
                if (
                    options_str.startswith('"')
                    and options_str.endswith('"')
                    and "," in options_str
                ):
                    options_str = options_str.strip('"')
                    workflow_options = [
                        opt.strip() for opt in options_str.split(",") if opt.strip()
                    ]
                    if workflow_options:
                        # Always enforce canonical priority order: Done > InProgress > ToDo.
                        # Some spreadsheets list values in another order (e.g. ToDo first),
                        # which previously caused downgrades (Done -> ToDo) because the
                        # RankedFieldStrategy treats lower index as higher priority.
                        canonical_order = [
                            WorkflowStatus.Done.value,
                            WorkflowStatus.InProgress.value,
                            WorkflowStatus.ToDo.value,
                        ]
                        priority_index = {v: i for i, v in enumerate(canonical_order)}
                        # Keep only known values, preserve canonical priority
                        sorted_opts = sorted(
                            (opt for opt in workflow_options if opt in priority_index),
                            key=lambda v: priority_index[v],
                        )
                        # Append any unexpected / custom statuses at the end so they never outrank core ones
                        extras = [
                            opt for opt in workflow_options if opt not in priority_index
                        ]
                        dynamic_added_fields["workflow_status"] = [
                            *sorted_opts,
                            *extras,
                        ]

    return dynamic_added_fields, dynamic_changeable_fields


# Core fields that are not updated (identifiers, etc.)

# Core fields that are not updated (identifiers, etc.)
core_fields = {
    "material_id": None,
    "title": None,
    "author": None,
    "publisher": None,
    "isbn": None,
    "doi": None,
    "url": None,
    "filename": None,
    "filetype": None,
    "pages_x_students": None,
    "count_students_registered": None,
    "faculty": None,
    "department": None,
    "course_code": None,
    "course_name": None,
    "period": None,
    "classification": None,
    "ml_prediction": None,
    "auditor": None,
    "status": None,
    "reliability": None,
    "last_change": None,
    "wordcount": None,
    "picturecount": None,
    "pagecount": None,
    "in_collection": None,
}


def get_mergeable_fields() -> dict:
    """Get all fields that can be merged/updated."""
    return {**added_fields, **changeable_fields}


def is_trivial_update(field: str, new_value, old_value) -> bool:
    """
    Determine if an update is trivial (simple overwrite) or needs complex merge logic.
    For now, consider status and last_change as trivial if they differ.
    """
    return field in ["status", "last_change"]
