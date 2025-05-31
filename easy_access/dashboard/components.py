"""
This module contains functions that generate various UI components for the web dashboard.
It uses `fasthtml` and `monsterui` for HTML and component generation, and includes
logic for rendering data tables, filter elements, pagination, item detail cards,
and other interactive UI parts. HTMX is used for dynamic updates.
"""
import datetime
import json
import logging # Added
from typing import Any, Dict, List, Optional, Tuple, Union # For type hints

import fasthtml.common as fh # Standard alias for fasthtml
import polars as pl
# from fastcore.utils import * # Avoid wildcard imports if possible, specific imports preferred
from fastcore.utils import L # Example if L is used, otherwise remove
# from fasthtml.common import * # Avoid wildcard, fh is used
from fasthtml.components import Button, Details, Div, H2, H4, H5, H6, Img, Li, Nav, P, Span, Summary, Table, Tbody, Td, Template, Th, Thead, Tr, Ul, A # Specific imports
from monsterui.all import UkIcon, Label, LabelInput # Specific imports from monsterui
# from rich import print # Removed, using logging

import easy_access.dashboard.urls as urls # Module for URL references
from easy_access.dashboard.constants import (
    BADGE_STYLES,
    DEFAULT_PILL_STYLE,
    DISPLAY_COLUMNS,
    FILTERABLE_COLUMNS,
    LabelT, # Enum from constants
)
# from easy_access.dashboard.data import get_item_df # Not directly used in this file from what's shown
from easy_access.dashboard.urls import Url # Enum for URL names
from easy_access.dashboard.web import AppState # For type hinting app_state
from easy_access.db.models import ItemUpdate # For type hinting history list
from easy_access.settings import SETTINGS, DirSetting # For potential path/settings access

logger = logging.getLogger(__name__)


def ItemDetailCard(
    title: str,
    *body_content: Any, # Can be multiple FT objects or strings
    card_id: str,
    col_span: int = 1,
    start_collapsed: bool = False,
    color: str = "[var(--ring)]", # Theme color variable
    lazy_load_url: Optional[str] = None,
    tooltip: Optional[str] = None,
) -> Details: # Changed FT to more specific Details
    """
    Creates a collapsible card component using HTML <details> and <summary> tags.
    Supports optional lazy loading of content via HTMX.

    Args:
        title (str): The title displayed in the card header.
        *body_content (Any): Components or text to be rendered in the card body.
                             If `lazy_load_url` is provided, this is ignored initially.
        card_id (str): A unique HTML ID for the card (details element).
        col_span (int, optional): TailwindCSS column span (1, 2, or 3). Defaults to 1.
        start_collapsed (bool, optional): If True, the card is initially collapsed. Defaults to False.
        color (str, optional): Theme color prefix for styling (e.g., "primary", "secondary").
                               Defaults to a CSS variable string.
        lazy_load_url (Optional[str], optional): URL to fetch content from when the card is opened.
                                               If provided, content is loaded via HTMX. Defaults to None.
        tooltip (Optional[str], optional): Tooltip text to display next to the card title. Defaults to None.

    Returns:
        Details: A `fasthtml.components.Details` object representing the card.
    """
    col_span = max(1, min(col_span, 3))
    col_span_class = f"md:col-span-{col_span}"

    # Constructing CSS classes carefully
    # Ensure color string doesn't break class names if it's complex like "[var(--ring)]"
    # For direct color names like "primary", "bg-primary-100" works.
    # If color is a CSS variable, it might need to be applied via style attribute or specific CSS.
    # Assuming `color` is a simple prefix for Tailwind classes here.
    base_bg_cls = "bg-base-100" # Default background
    border_cls = f"border-{color}-500" if color not in ["[var(--ring)]", "base"] else "border-[var(--ring)]"
    header_bg_cls = f"bg-{color}-100" if color not in ["[var(--ring)]", "base"] else base_bg_cls

    card_classes = f"{base_bg_cls} border {border_cls} rounded-lg shadow-sm {col_span_class}"
    summary_classes = f"{header_bg_cls} p-3 cursor-pointer list-none flex items-center justify-between"
    content_classes = "p-4 border-t border-[var(--ring)]" # Content area border
    chevron_icon = Span("▼", cls="text-xs transition-transform duration-200 chevron-icon")

    content_target_id = f"{card_id}-content"
    summary_htmx_attrs: Dict[str, Any] = {}

    actual_card_body_content: Tuple[Any, ...]
    if lazy_load_url:
        actual_card_body_content = (
            Span("Loading...", cls="italic text-sm text-base-content/70"),
            Div(id=f"{card_id}-loading", cls="htmx-indicator text-center p-2")(
                Span(cls="loading loading-sm") # Loading spinner
            ),
        )
        summary_htmx_attrs = {
            "hx_get": lazy_load_url, "hx_target": f"#{content_target_id}",
            "hx_swap": "innerHTML", "hx_trigger": "click toggle", # "toggle" if summary click re-fetches
        }
    else:
        actual_card_body_content = body_content

    tooltip_icon: FT = Span("") # Empty span if no tooltip
    if tooltip:
        tooltip_icon = UkIcon( # Assuming UkIcon handles tooltips or is styled for it
            "circle-help", cls="ml-1 text-xs cursor-help", # Simpler classes
            height="12px", width="12px", title=tooltip # Standard title attribute for tooltip
        )

    content_div = Div(*actual_card_body_content, id=content_target_id, cls=content_classes)
    summary_element = Summary(
        H5(title, tooltip_icon, cls="font-semibold text-sm m-0 inline-flex items-center"), # Ensure items align
        chevron_icon,
        cls=summary_classes,
        **summary_htmx_attrs,
    )

    details_attrs = {"open": True} if not start_collapsed else {}
    return Details(summary_element, content_div, id=card_id, cls=card_classes, **details_attrs)


def create_checkbox_filter_group(
    filter_key: str,
    label_text: str,
    options: Dict[str, Union[str, LabelT]], # Value can be class string or LabelT enum
    current_values_str: Optional[str], # Pipe-separated string of selected values
    current_total_items: int,
    option_counts: Optional[Dict[str, Union[int, str]]] = None, # Counts for each option
) -> Div:
    """
    Creates a UI group of styled checkboxes for a given filter category.
    Checkboxes have counts and trigger HTMX requests on change.
    Options are sorted with enabled/active ones first, then by count/name.

    Args:
        filter_key (str): The internal key for this filter (e.g., "workflow_status").
        label_text (str): The user-visible label for the filter group.
        options (Dict[str, Union[str, LabelT]]): Dictionary of option values to their style info.
        current_values_str (Optional[str]): A pipe-separated string of currently selected filter values.
        current_total_items (int): Total items matching the *other* active filters (used for count display logic).
        option_counts (Optional[Dict[str, Union[int, str]]]): Pre-calculated counts for each option.
                                                             Value can be int or "-" (e.g. if option is disabled).
    Returns:
        Div: A Div component containing the styled checkbox group.
    """
    logger.debug(f"Creating checkbox group for: {filter_key}, current values: '{current_values_str}', counts: {option_counts}")
    selected_values_set: Set[str] = set(current_values_str.split("|")) if current_values_str else set()

    form_field_name = f"filter_{filter_key}" # Name used in form submission

    # Sort options: selected first, then by count (desc), then alphabetically
    def get_sort_key(option_tuple: Tuple[str, Union[str, LabelT]]) -> Tuple[int, int, str]:
        val, _ = option_tuple
        is_selected = val in selected_values_set
        count_val = option_counts.get(val) if option_counts else 0
        numeric_count = 0
        if isinstance(count_val, int): numeric_count = count_val
        elif isinstance(count_val, str) and count_val.isdigit(): numeric_count = int(count_val)

        return (
            0 if is_selected else 1,  # Selected items first
            -numeric_count,           # Then by count descending (higher counts first)
            val.lower()               # Then alphabetically
        )

    sorted_options_list = sorted(options.items(), key=get_sort_key)

    checkbox_items_list: List[FT] = []
    disabled_options_display_list: List[FT] = [] # For options with no items

    for option_value_str, style_or_class in sorted_options_list:
        safe_html_id_suffix = "".join(c if c.isalnum() else "-" for c in option_value_str) # For HTML IDs
        current_option_count = option_counts.get(option_value_str, "") if option_counts else ""
        is_checked_bool = option_value_str in selected_values_set

        # Determine if checkbox should be disabled (e.g., if count is 0 and not selected)
        is_disabled_bool = (isinstance(current_option_count, int) and current_option_count == 0 and not is_checked_bool) or \
                           (current_option_count == "-" and not is_checked_bool) # Explicitly no items

        # Determine checkbox and badge styling
        checkbox_color_class = "checkbox-primary" # Default
        badge_style_class = str(DEFAULT_PILL_STYLE) # Default

        if filter_key != "faculty": # Faculty uses predefined badge-color strings
            if isinstance(style_or_class, LabelT): # Map LabelT to specific badge classes
                if style_or_class == LabelT.primary: checkbox_color_class, badge_style_class = "checkbox-accent", "badge badge-neutral"
                elif style_or_class == LabelT.secondary: checkbox_color_class, badge_style_class = "checkbox-primary", "badge badge-primary"
                elif style_or_class == LabelT.destructive: checkbox_color_class, badge_style_class = "checkbox-error", "badge badge-error"
            elif isinstance(style_or_class, str): # Direct class string
                badge_style_class = style_or_class
        elif isinstance(style_or_class, str): # Faculty uses direct badge classes
             badge_style_class = style_or_class


        count_display_str: str
        count_tooltip_str: str = ""
        if is_checked_bool: # If currently selected
            count_display_str = f"= {current_total_items}" # Show total items for active filter
            count_tooltip_str = "Currently selected. Shows total items matching all active filters."
        elif isinstance(current_option_count, int):
            count_display_str = f"▼ {current_option_count}" if current_option_count < current_total_items else f"▲ {current_option_count}"
            if current_option_count == current_total_items and current_active_filters: # current_active_filters from closure
                 count_display_str = f"= {current_option_count}" # No change if added
            count_tooltip_str = f"Adds/removes this option. Resulting count: {current_option_count}."
        else: # count is "-" or other string
            count_display_str = "" # Don't show count for globally empty options unless selected
            count_tooltip_str = "No items match this option in the current dataset."


        if is_disabled_bool:
            disabled_options_display_list.append(
                Span(option_value_str if len(option_value_str) < 18 else option_value_str[:16] + "…",
                     cls="text-xs opacity-40 mr-2 cursor-not-allowed", title=f"{option_value_str} (no items)")
            )
        else:
            count_badge_span = Span(count_display_str, title=count_tooltip_str,
                                 cls=f"{badge_style_class} badge-outline ml-1 align-middle text-xs cursor-help") if count_display_str else Span()

            htmx_vals_for_cb = {form_field_name: option_value_str, "checked": str(not is_checked_bool).lower()}

            checkbox = fh.CheckboxX( # Assuming CheckboxX is a custom or fasthtml component
                name=form_field_name, value=option_value_str, id=f"cb-{filter_key}-{safe_html_id_suffix}",
                checked=is_checked_bool, cls=f"checkbox checkbox-xs align-middle mr-1 {checkbox_color_class}",
                hx_post=str(urls.URLS[Url.data_grid]), hx_target="#data-grid-component", hx_swap="outerHTML",
                hx_trigger="change", hx_vals=json.dumps(htmx_vals_for_cb), disabled=is_disabled_bool
            )
            badge_text_label = Span(option_value_str, cls=f"{badge_style_class} mr-1 cursor-pointer text-xs")

            wrapper_label_for_cb = fh.Label(checkbox, badge_text_label, count_badge_span,
                                         cls="flex items-center gap-x-1 py-0 px-1 mb-0.5 rounded min-h-0",
                                         html_for=f"cb-{filter_key}-{safe_html_id_suffix}") # Use html_for
            checkbox_items_list.append(wrapper_label_for_cb)

    # Construct the group card
    group_card_content: List[FT] = [
        fh.Label(label_text, cls="label-text pb-0 text-xs text-base-content/90 block mb-1 font-semibold tracking-wide"),
        Div(*checkbox_items_list, cls="grid grid-cols-1 gap-x-1 gap-y-0.5 mb-0.5")
    ]
    if disabled_options_display_list:
        group_card_content.append(
            Div(UkIcon("minus-circle", cls="text-xs opacity-40 cursor-help", title="Options with no items currently"),
                *disabled_options_display_list, cls="flex flex-wrap items-center gap-x-1 gap-y-0.5 mt-1 mb-0.5")
        )

    return Div(*group_card_content, cls="bg-base-100 border border-base-300 rounded-lg p-2 mb-2 shadow-xs", id=f"filter-group-{filter_key}")


def render_course_details(item_data: Dict[str, Any]) -> Tuple[FT, ...]:
    """
    Renders the content for the 'Course Details' card in the item modal.
    Displays information about associated courses, including names, codes, year,
    programme, faculty, and related organizations of contact persons.

    Args:
        item_data (Dict[str, Any]): The copyright item's data dictionary, expected
                                   to contain a 'courses' key with a list of course dicts.

    Returns:
        Tuple[FT, ...]: A tuple of `fasthtml` components to be rendered in the card.
    """
    course_items_list: List[FT] = []
    all_course_org_abbreviations: Set[str] = set()
    courses_data: List[Dict[str, Any]] = item_data.get("courses", [])

    if not courses_data:
        return (P("No course data associated with this item.", cls="text-sm text-base-content/70"),)

    for course in courses_data:
        course_items_list.append(
            Div(cls="mb-3 p-2 border rounded border-base-300")(
                H6(f"{course.get('name', 'Unknown Course')} ({course.get('cursuscode', '?')})", cls="font-semibold text-sm mb-1"),
                P(f"Year: {course.get('year', 'N/A')}, Programme: {course.get('programme', 'N/A')}, Faculty: {course.get('faculty_id', 'N/A')}",
                  cls="text-xs text-base-content/80"),
            )
        )
        persons_in_course: List[Dict[str, Any]] = course.get("persons", [])
        for person in persons_in_course:
            orgs_of_person: List[Dict[str, Any]] = person.get("organizations", [])
            for org in orgs_of_person:
                if org.get("full_abbreviation"): # Prefer full_abbreviation if available
                    all_course_org_abbreviations.add(str(org["full_abbreviation"]))

    if all_course_org_abbreviations:
        course_items_list.append(fh.Div(cls="my-2 border-t border-base-300")) # Using fh.Div for Divider
        course_items_list.append(fh.Strong("Related Organizations (from Course Contacts):", cls="text-xs font-medium block mb-1"))
        course_items_list.append(
            Div(cls="flex flex-wrap gap-1")(
                *[Label(org_abbr, cls=f"badge badge-sm {DEFAULT_PILL_STYLE}") for org_abbr in sorted(list(all_course_org_abbreviations))]
            )
        )
    return tuple(course_items_list)


def render_contact_info(item_data: Dict[str, Any]) -> Tuple[FT, ...]:
    """
    Renders the 'Contact Information' card contents for the item modal.
    Displays contact persons associated with the item's courses, their organization,
    and links for email and people pages. Includes a button to copy all contact emails.

    Args:
        item_data (Dict[str, Any]): The copyright item's data, expected to have a 'courses'
                                   list, each with a 'persons' list containing contact details.

    Returns:
        Tuple[FT, ...]: A tuple of `fasthtml` components.
    """
    contact_persons_list: List[Dict[str, Any]] = []
    all_contact_emails_set: Set[str] = set()
    courses_data: List[Dict[str, Any]] = item_data.get("courses", [])

    for course in courses_data:
        persons_in_course: List[Dict[str, Any]] = course.get("persons", [])
        for person in persons_in_course:
            if person.get("role") == "contact": # Filter for contacts
                contact_persons_list.append(person)
                if person.get("email") and isinstance(person["email"], str):
                    all_contact_emails_set.add(person["email"])

    rendered_items: List[FT] = []
    if all_contact_emails_set:
        emails_str = ";".join(sorted(list(all_contact_emails_set)))
        copy_js_onclick = f"navigator.clipboard.writeText('{emails_str}'); event.target.textContent='Copied!';" \
                          f"setTimeout(()=>event.target.textContent='Copy all emails', 2000);"
        rendered_items.append(
            Div(fh.Button(UkIcon("copy", cls="w-4 h-4 mr-1"), "Copy all emails",
                        cls="btn btn-primary btn-sm btn-block cursor-pointer", # Changed cursor-copy
                        onclick=copy_js_onclick), cls="mb-2")
        )

    if not contact_persons_list:
        rendered_items.append(P("No contact persons found for this item's courses.", cls="text-sm text-base-content/70"))
    else:
        # Avoid duplicate persons if they are contacts for multiple courses linked to the same item
        unique_person_ids_processed: Set[Any] = set()
        for person in sorted(contact_persons_list, key=lambda p: p.get("main_name", "zzzz")): # Sort by name
            person_id = person.get("id")
            if person_id in unique_person_ids_processed: continue
            unique_person_ids_processed.add(person_id)

            email_addr = person.get("email")
            people_page = person.get("people_page_url")
            orgs: List[Dict[str, Any]] = person.get("organizations", [])

            # Display top-level org abbreviation(s) for the person
            top_org_abbrs_set: Set[str] = set()
            if orgs: # Find highest level orgs (smallest hierarchy_level)
                min_level = min((o.get("hierarchy_level", float('inf')) for o in orgs if o.get("hierarchy_level") is not None), default=float('inf'))
                if min_level != float('inf'):
                    top_org_abbrs_set = {str(o["full_abbreviation"]) for o in orgs if o.get("hierarchy_level") == min_level and o.get("full_abbreviation")}

            rendered_items.append(
                Div(cls="flex items-center justify-between space-x-2 py-1.5 border-b border-base-200 last:border-b-0")(
                    Span(str(person.get("main_name", "Unknown Name")), cls="text-sm truncate", title=str(person.get("main_name", "Unknown Name"))),
                    Div(cls="flex items-center space-x-1 flex-shrink-0")(
                        *[Label(abbr, cls=f"badge badge-xs {DEFAULT_PILL_STYLE}") for abbr in sorted(list(top_org_abbrs_set))], # Badges for orgs
                        A(UkIcon("mail", cls="w-4 h-4"), href=f"mailto:{email_addr}", title=f"Email {email_addr}", cls="link text-primary") if email_addr else Span(UkIcon("mail-off", cls="w-4 h-4 text-base-content/50"), title="No email"), # mail-off icon
                        A(UkIcon("external-link", cls="w-4 h-4"), href=str(people_page), target="_blank", title="Open people page", cls="link text-primary") if people_page else Span(UkIcon("link-2-off", cls="w-4 h-4 text-base-content/50"), title="No people page URL"),
                    )
                )
            )
    return tuple(rendered_items)


def render_teacher_info(item_data: Dict[str, Any]) -> Tuple[FT, ...]:
    """
    Renders the 'Teachers/Persons Involved' card contents for the item modal.
    Displays unique persons (teachers, tutors, examiners, etc.) associated with the item's courses.

    Args:
        item_data (Dict[str, Any]): The copyright item's data.

    Returns:
        Tuple[FT, ...]: A tuple of `fasthtml` components.
    """
    # Collect all unique persons involved, regardless of role, from all courses
    all_persons_map: Dict[Any, Dict[str, Any]] = {} # Keyed by person ID
    courses_data: List[Dict[str, Any]] = item_data.get("courses", [])
    for course in courses_data:
        persons_in_course: List[Dict[str, Any]] = course.get("persons", [])
        for person in persons_in_course:
            person_id = person.get("id")
            if person_id and person_id not in all_persons_map:
                all_persons_map[person_id] = person

    rendered_items: List[FT] = []
    if not all_persons_map:
        return (P("No teachers or other persons found for this item's courses.", cls="text-sm text-base-content/70"),)

    sorted_persons_list = sorted(all_persons_map.values(), key=lambda p: str(p.get("main_name", "zzzz")))

    for person in sorted_persons_list:
        people_page = person.get("people_page_url")
        orgs: List[Dict[str, Any]] = person.get("organizations", [])
        top_org_abbrs_set: Set[str] = set()
        if orgs:
            min_level = min((o.get("hierarchy_level", float('inf')) for o in orgs if o.get("hierarchy_level") is not None), default=float('inf'))
            if min_level != float('inf'):
                top_org_abbrs_set = {str(o["full_abbreviation"]) for o in orgs if o.get("hierarchy_level") == min_level and o.get("full_abbreviation")}

        rendered_items.append(
            Div(cls="flex items-center justify-between space-x-2 py-1.5 border-b border-base-200 last:border-b-0")(
                Span(str(person.get("main_name", "Unknown Name")), cls="text-sm truncate", title=str(person.get("main_name", "Unknown Name"))),
                Div(cls="flex items-center space-x-1 flex-shrink-0")(
                    *[Label(abbr, cls=f"badge badge-xs {DEFAULT_PILL_STYLE}") for abbr in sorted(list(top_org_abbrs_set))],
                    A(UkIcon("external-link", cls="w-4 h-4"), href=str(people_page), target="_blank", title="Open people page", cls="link text-primary") if people_page else Span(UkIcon("link-2-off", cls="w-4 h-4 text-base-content/50"), title="No people page URL"),
                )
            )
        )
    return tuple(rendered_items)


def render_item_history(history_records: List[ItemUpdate]) -> FT: # Changed input type
    """
    Renders a card displaying the change history for a copyright item.

    Args:
        history_records (List[ItemUpdate]): A list of `ItemUpdate` ORM objects,
                                         ordered by recency (most recent first).
    Returns:
        FT: A `Details` component containing the formatted history, or a Div with "No history".
    """
    if not history_records:
        return Div(P("No history found for this item.", cls="text-sm p-4"))

    logger.debug(f"Rendering item history for {len(history_records)} records.")

    timeline_log_entries: List[FT] = []
    # Assuming history_records is already sorted (e.g., by -created_at from DB query)
    for record in history_records:
        change_details_dict: Dict[str, Any] = record.change_details # Already a dict

        timestamp_str = record.created_at.strftime("%Y-%m-%d %H:%M:%S") if record.created_at else "N/A"
        modified_by_str = str(change_details_dict.get("modified_by", "Unknown User")) # User who made change

        formatted_changes_list: List[str] = []
        if isinstance(change_details_dict, dict):
            for field_key, change_val_dict in change_details_dict.items():
                if field_key in ["material_id", "update_time", "modified_by", "modified_at"]: # Skip meta fields
                    continue
                if isinstance(change_val_dict, dict): # Expected structure: {"old": "val1", "new": "val2"}
                    old_val = str(change_val_dict.get('old', 'N/A')).strip() or "N/A"
                    new_val = str(change_val_dict.get('new', 'N/A')).strip() or "N/A"
                    field_display_name = field_key.replace('_', ' ').title()
                    formatted_changes_list.append(
                        f"<span class='font-semibold'>{field_display_name}:</span> "
                        f"<span class='text-error-content/80 line-through'>{old_val}</span> → "
                        f"<span class='text-success-content/90'>{new_val}</span>"
                    )
                else: # Fallback for unexpected change_details format
                    formatted_changes_list.append(f"{field_key}: {change_val_dict}")

        if not formatted_changes_list:
            formatted_changes_list.append("No specific field changes logged (or only meta fields updated).")

        # Constructing HTML for each log entry
        # Using NotStr for pre-formatted HTML strings to be rendered by fasthtml
        log_entry_html = f"""
            <div class="flex flex-col">
                <div class="text-xs text-base-content/70 mb-0.5">
                    <span class="font-semibold">{timestamp_str}</span> by <span class="italic">{modified_by_str}</span>
                </div>
                <div class="text-xs pl-2 border-l-2 border-base-300">
                    {'<br>'.join(formatted_changes_list)}
                </div>
            </div>
        """
        timeline_log_entries.append(Li(NotStr(log_entry_html))) # NotStr tells fasthtml it's raw HTML

    return ItemDetailCard( # Use the standard card component
        title="Change History",
        body_content=(Ul(*timeline_log_entries, cls="space-y-3 list-none p-0 m-0"),), # Pass content as tuple
        card_id="item-history-card",
        col_span=1, # Example, adjust as needed
        start_collapsed=True,
        tooltip="Chronological record of changes made to this item.",
    )


def render_table_rows(
    df_slice: pl.DataFrame,
    # app_state: AppState, # app_state might not be needed if all URLs are from constants
) -> List[FT]: # Return List[FT] instead of Tuple for easier appends
    """
    Renders table body rows (Tbody) for the main data grid.
    Each row corresponds to a copyright item and includes formatted cells for display
    and HTMX attributes for loading item details into a modal.

    Args:
        df_slice (pl.DataFrame): The slice of data (current page) to render.
        # app_state (AppState): The current application state (potentially useful for constructing URLs if dynamic).

    Returns:
        List[FT]: A list of `Tr` (table row) components.
    """
    table_rows: List[Tr] = [] # Changed to List[Tr]
    if df_slice is not None and not df_slice.is_empty(): # df_slice.height > 0 changed to not df_slice.is_empty()
        # Ensure columns to display actually exist in the DataFrame slice
        cols_to_render = [col for col in DISPLAY_COLUMNS if col in df_slice.columns]

        for row_data_dict in df_slice.select(cols_to_render).iter_rows(named=True):
            table_cells: List[Td] = [] # Changed to List[Td]
            material_id_val = row_data_dict.get("material_id")
            if material_id_val is None: continue # Skip row if no material_id

            for col_name in cols_to_render:
                cell_value = row_data_dict.get(col_name)
                td_cls_name = f"col-{col_name.replace('_', '-')}"

                cell_display_content: Any # FT or str
                # Default display for most values
                cell_display_text = str(cell_value) if cell_value is not None else "N/A"

                if col_name == "url" and cell_value:
                    cell_display_content = A(
                        UkIcon("external-link", cls="w-4 h-4 inline-block"), href=str(cell_value),
                        target="_blank", rel="noopener noreferrer", title=f"Open link: {cell_value}",
                        cls="link link-hover text-primary hover:text-primary-focus", onclick="event.stopPropagation();"
                    )
                elif col_name in BADGE_STYLES or "classification" in col_name or "prediction" in col_name:
                    style_key_map = BADGE_STYLES.get("classification") if "classification" in col_name or "prediction" in col_name else BADGE_STYLES.get(col_name, {})
                    style_class_enum = style_key_map.get(str(cell_value).lower() if cell_value else "none", DEFAULT_PILL_STYLE) # Use lower for matching
                    style_class_str = str(style_class_enum) # Convert LabelT to string class

                    display_text_for_badge = cell_display_text.split("-")[0].strip() if "-" in cell_display_text else cell_display_text.strip()
                    cell_display_content = Label(display_text_for_badge, title=cell_display_text, cls=f"{style_class_str} badge-sm cursor-help", onclick="event.stopPropagation();")
                else: # Default text rendering with truncation
                    cell_display_content = display_text_for_badge[:MAX_CELL_DISPLAY_LENGTH] + "..." if len(display_text_for_badge) > MAX_CELL_DISPLAY_LENGTH else display_text_for_badge

                table_cells.append(
                    Td(cell_display_content, title=cell_display_text, cls=f"py-2 px-3 border-b border-base-200/80 {td_cls_name}") # Adjusted padding
                )

            # Construct modal URL using the Url enum and replacing placeholder
            modal_url_str = str(urls.URLS[Url.show_item_details]).replace("{material_id:int}", str(material_id_val))
            row_htmx_attrs: Dict[str, Any] = {
                "id": f"row-{material_id_val}",
                "class": "hover:bg-primary/10 cursor-pointer transition-colors duration-150",
                "hx_get": modal_url_str, "hx_target": "#modal-placeholder", "hx_swap": "innerHTML",
                # "preload": "mouseover" # Consider preload on mouseover for faster modal
            }
            table_rows.append(Tr(*table_cells, **row_htmx_attrs))

    if not table_rows: # If loop finishes and table_rows is still empty
        table_rows.append(
            Tr(Td("No matching data found for the current filters.", colspan=str(len(DISPLAY_COLUMNS)), class_="text-center text-sm p-6"))
        )
    return table_rows


def page_header_component(
    user_details: Dict[str, Optional[str]], # User details from session/auth
    app_state: AppState, # Current application state
    filter_counts: Dict[str, Dict[str, Union[int, str]]], # Pre-calculated filter counts
    total_filtered_rows: int, # Total items matching current filters
) -> Div:
    """
    Renders the header section of the dashboard page.
    Includes the application title, last update time, user information display,
    logout button, and a collapsible sidebar for filters.

    Args:
        user_details (Dict[str, Optional[str]]): Dictionary containing user information
                                                 (e.g., 'name', 'faculty', 'role').
        app_state (AppState): The current application state.
        filter_counts (Dict[str, Dict[str, Union[int, str]]]): Pre-calculated counts for filter options.
        total_filtered_rows (int): Total number of items matching the current filters.

    Returns:
        Div: A `fasthtml.Div` component representing the page header.
    """
    current_active_filters = app_state.get_active_filters()
    is_admin_user = user_details.get("role") == "admin"
    user_faculty_display = user_details.get("faculty", "All items")
    if user_faculty_display == "all": user_faculty_display = "All items"


    # --- Filter Checkbox Groups ---
    filter_checkbox_groups_list: List[FT] = [
        create_checkbox_filter_group(
            filter_key="workflow_status", label_text="Workflow Status", options=BADGE_STYLES["workflow_status"],
            current_values_str=current_active_filters.get("workflow_status"), option_counts=filter_counts.get("workflow_status"),
            current_total_items=total_filtered_rows,
        ),
        create_checkbox_filter_group(
            filter_key="status", label_text="Status", options=BADGE_STYLES["status"],
            current_values_str=current_active_filters.get("status"), option_counts=filter_counts.get("status"),
            current_total_items=total_filtered_rows,
        ),
        create_checkbox_filter_group( # Using "classification" style for both actual and manual classification
            filter_key="classification", label_text="System Classification", options=BADGE_STYLES["classification"],
            current_values_str=current_active_filters.get("classification"), option_counts=filter_counts.get("classification"),
            current_total_items=total_filtered_rows,
        ),
        create_checkbox_filter_group(
            filter_key="manual_classification", label_text="Manual Classification", options=BADGE_STYLES["classification"],
            current_values_str=current_active_filters.get("manual_classification"), option_counts=filter_counts.get("manual_classification"),
            current_total_items=total_filtered_rows,
        ),
    ]
    if is_admin_user or not user_details.get("faculty") or user_details.get("faculty") == "all":
        filter_checkbox_groups_list.append(create_checkbox_filter_group(
            filter_key="faculty", label_text="Faculty", options=BADGE_STYLES["faculty"],
            current_values_str=current_active_filters.get("faculty"), option_counts=filter_counts.get("faculty"),
            current_total_items=total_filtered_rows,
        ))

    # --- Text Filter Inputs ---
    text_filter_inputs_list: List[FT] = []
    # Determine which columns are left for text filtering
    cols_for_checkboxes = {"workflow_status", "status", "classification", "manual_classification"}
    if is_admin_user or not user_details.get("faculty") or user_details.get("faculty") == "all":
        cols_for_checkboxes.add("faculty")

    # Ensure get_item_df() is called to get schema for columns; cache helps
    # item_df_cols = get_item_df().columns
    # This might be too slow if get_item_df() is not cached or involves heavy computation.
    # Assuming FILTERABLE_COLUMNS contains valid column names present in the DataFrame.

    for col_name in FILTERABLE_COLUMNS:
        if col_name not in cols_for_checkboxes: # Only add if not already a checkbox group
            # if col_name not in item_df_cols: continue # Skip if column doesn't exist in data

            field_label = f"{col_name.replace('_', ' ').title()}"
            input_id_str = f"filter-input-{col_name}"
            form_input_name = f"filter_{col_name}"
            # JS to get value for hx-vals; ensure proper quoting for string values
            js_hx_vals = f"{{'{form_input_name}': document.getElementById('{input_id_str}').value}}"

            text_filter_inputs_list.append(LabelInput( # Assuming LabelInput is a monsterui component
                label=field_label, name=form_input_name, value=current_active_filters.get(col_name, ""),
                placeholder="Filter...", id=input_id_str,
                input_cls="input input-bordered input-xs w-full focus:input-primary",
                label_cls="label-text pb-1 text-xs font-medium text-base-content/90",
                cls="form-control w-full mb-2", # Added mb-2 for spacing
                hx_post=str(urls.URLS[Url.data_grid]), hx_trigger="keyup changed delay:500ms, search",
                hx_target="#data-grid-component", hx_vals=js_hx_vals,
            ))

    # --- Filter Form ---
    filter_form_content = Form(
        Div( # Button container at top of sidebar
            fh.Button(UkIcon("rotate-ccw", cls="mr-1 w-4 h-4"), "Reset All Filters", type="button",
                        hx_post=str(urls.URLS[Url.data_grid]), hx_target="#data-grid-component",
                        hx_vals='{"action": "reset"}', cls="btn btn-xs btn-warning mx-1",
                        **{"@click": "slideOverOpen=false"} # Alpine.js to close sidebar
            ),
            cls="sticky top-0 z-20 flex justify-center bg-base-200 py-2 border-b border-base-300 mb-2"
        ),
        *filter_checkbox_groups_list,
        *text_filter_inputs_list,
        id="filter-form", cls="flex flex-col gap-y-1 p-2" # Added padding to form
    )

    # --- Sidebar Structure (using Alpine.js for toggle) ---
    sidebar_component = Div(x_data="{ slideOverOpen: false }", cls="relative z-50")( # Higher z-index
        # Button to open sidebar
        fh.Button(UkIcon("filter", cls="mr-1"), "Filters", **{"@click": "slideOverOpen=true"}, cls="btn btn-sm btn-primary"),
        # Sidebar panel (teleported to body)
        Template(x_teleport="body")(
            Div(x_show="slideOverOpen", **{"@keydown.window.escape": "slideOverOpen=false"},
                cls="fixed inset-0 overflow-hidden", x_cloak=True, hx_ext="preload")( # hx_ext for preloading if needed
                Div(x_show="slideOverOpen", x_transition_opacity_duration_600ms=True, # fasthtml specific transition
                    **{"@click": "slideOverOpen = false"}, cls="absolute inset-0 bg-black bg-opacity-25 backdrop-blur-sm"),
                Div(cls="fixed inset-y-0 left-0 flex max-w-full pr-10")( # Sidebar container
                    Div(x_show="slideOverOpen", **{"@click.away": "slideOverOpen = false"},
                        x_transition_enter="transform transition ease-in-out duration-300 sm:duration-500",
                        x_transition_enter_start="-translate-x-full", x_transition_enter_end="translate-x-0",
                        x_transition_leave="transform transition ease-in-out duration-300 sm:duration-500",
                        x_transition_leave_start="translate-x-0", x_transition_leave_end="-translate-x-full",
                        cls="w-screen max-w-xs md:max-w-sm")( # Responsive width
                        Div(cls="flex flex-col h-full overflow-y-auto bg-base-200 border-r border-base-300 shadow-xl")(
                            Div(cls="p-3 border-b border-base-300")( # Sidebar Header
                                Div(cls="flex items-center justify-between")(
                                    H2("Filters", cls="text-lg font-semibold text-base-content"),
                                    fh.Button("✕", **{"@click": "slideOverOpen=false"}, cls="btn btn-sm btn-circle btn-ghost")
                                )
                            ),
                            Div(cls="relative flex-1")(filter_form_content), # Scrollable filter form
                        )
                    )
                )
            )
        )
    )

    # Determine last update time for display
    last_update_display: str = "Unknown"
    raw_data_dir = SETTINGS.dirs.get(DirSetting.RAW_COPYRIGHT_DATA)
    if raw_data_dir and raw_data_dir.exists:
        # Get most recent file modification time in the directory
        # This might be slow if there are many files; consider storing last import time in DB or cache.
        try:
            latest_file_mod_time = max((f.modified for f in raw_data_dir.files if f.exists), default=None)
            if latest_file_mod_time: last_update_display = latest_file_mod_time.strftime("%Y-%m-%d %H:%M")
        except Exception as e_stat: logger.debug(f"Could not determine last update time from file system: {e_stat}")


    return Div( # Main header container
        Div(cls="flex flex-wrap justify-between items-center mb-4 gap-y-2")( # Responsive wrap
            Div( # Title and last update
                H4("Copyright Data Dashboard", cls="text-xl md:text-2xl font-bold text-primary"),
                P(f"Data last updated: {last_update_display}", cls="text-xs text-base-content/70 mt-1")
            ),
            Div(cls="flex items-center space-x-3 md:space-x-4")( # User info and actions
                sidebar_component, # Filter button and sidebar
                Div(cls="flex items-center space-x-2")( # Avatar and user name
                    Img(src=f"https://api.dicebear.com/9.x/initials/svg?seed={user_details.get('name', 'User')}&backgroundColor=00897b,00acc1,039be5,3949ab,5e35b1,8e24aa&backgroundType=gradientLinear&radius=50",
                        alt="avatar", cls="w-8 h-8 md:w-10 md:h-10 rounded-full border-2 border-base-300 shadow-sm"),
                    Div(
                        Span(f"Hi, {user_details.get('name', 'User')}!", cls="text-sm font-medium hidden md:block"),
                        Span(f"Faculty: {user_faculty_display}", cls="text-xs text-base-content/70 block")
                    )
                ),
                A("Logout", href=str(urls.URLS[Url.logout]), cls="btn btn-outline btn-error btn-sm") # Use Url enum
            )
        ),
        cls="mb-4 border-b border-base-300 pb-4 pt-2" # Added pt-2
    )


def render_data_grid_component(
    df_slice: pl.DataFrame,
    app_state: AppState,
    total_filtered_rows: int,
    total_pages: int,
) -> Div:
    """
    Renders the main data grid component, including the table, headers, and pagination.

    Args:
        df_slice (pl.DataFrame): The slice of data for the current page.
        app_state (AppState): The current application state.
        total_filtered_rows (int): Total number of rows matching current filters.
        total_pages (int): Total number of pages for the filtered data.

    Returns:
        Div: A Div component representing the entire data grid section.
    """
    header_cells_list: List[Th] = []
    # Use DISPLAY_COLUMNS to define order and presence of columns
    # Ensure get_item_df() is efficient if called here, or pass columns schema
    # item_df_schema_cols = get_item_df().columns # Potentially slow if get_item_df() is uncached and hits DB
    # For now, assume DISPLAY_COLUMNS are valid.

    for col_name_to_display in DISPLAY_COLUMNS:
        # if col_name_to_display not in item_df_schema_cols: continue # Skip if not in actual data

        header_display_text = "ID" if col_name_to_display == "material_id" else col_name_to_display.replace("_", " ").title()

        sort_icon_class = ""
        next_sort_direction_is_desc = False # Default to ASC for next click if not current sort col

        if col_name_to_display == app_state.sort_by: # This is the current sort column
            next_sort_direction_is_desc = not app_state.sort_desc # Toggle current direction
            sort_icon_class = "arrow-down" if app_state.sort_desc else "arrow-up"

        sort_htmx_vals = json.dumps({"sort_by": col_name_to_display, "sort_desc": str(next_sort_direction_is_desc).lower()})

        header_button_content: List[Any] = [header_display_text]
        if sort_icon_class:
            header_button_content.append(UkIcon(sort_icon_class, cls="inline-block ml-1.5 w-3 h-3"))

        header_cells_list.append(
            Th(
                Button(*header_button_content,
                    hx_get=str(urls.URLS[Url.data_grid]), hx_target="#data-grid-component", hx_vals=sort_htmx_vals,
                    cls=f"{ButtonT.ghost} text-xs uppercase tracking-wider p-1.5 h-auto min-h-0 font-bold text-base-content/70 hover:text-primary"
                ),
                cls=f"px-3 py-2 col-{col_name_to_display.replace('_', '-')}" # Adjusted padding
            )
        )

    table_header = Thead(Tr(*header_cells_list))
    table_body = Tbody(*render_table_rows(df_slice), id="data-table-body") # Pass app_state if render_table_rows needs it

    table_component = Div(
        Table(table_header, table_body, role="grid", id="data-table", cls="table table-sm table-zebra w-full"), # Added table-zebra
        id="table-wrapper", cls="overflow-x-auto border border-base-300 rounded-lg shadow-sm bg-base-100"
    )

    pagination_component = render_pagination(app_state, total_pages, total_filtered_rows)

    return Div(
        pagination_component,
        Div(table_component, cls="htmx-indicator-opacity opacity-100 transition-opacity duration-300"), # Simpler opacity indicator
        id="data-grid-component", cls="flex flex-col" # Removed justify-center
    )


def render_pagination(
    app_state: AppState,
    total_pages: int,
    total_items: int,
) -> Div: # Changed return to Div
    """
    Renders pagination controls for the data table.

    Includes 'First', 'Previous', 'Next', 'Last' buttons, and a series of page number links.
    Uses HTMX attributes to trigger data grid updates on page changes.

    Args:
        app_state (AppState): The current application state (used for current page, items per page).
        total_pages (int): Total number of pages available.
        total_items (int): Total number of items matching current filters.

    Returns:
        Div: A Div component containing the pagination controls.
    """
    if total_pages <= 1 and total_items <= app_state.per_page : # Hide if only one page and not many items
        return Div(cls="h-12 mb-1") # Placeholder for height consistency

    # CSS classes for styling pagination buttons
    btn_base_cls = "btn btn-xs btn-ghost" # Base for all buttons
    btn_enabled_cls = f"{btn_base_cls} hover:bg-primary hover:text-primary-content"
    btn_disabled_cls = f"{btn_base_cls} btn-disabled text-base-content/50"
    num_selected_cls = f"{btn_base_cls} btn-active btn-primary"
    num_other_cls = btn_enabled_cls
    ellipsis_cls = f"{btn_base_cls} btn-disabled"


    current_page: int = app_state.page
    start_item_idx: int = (app_state.per_page * (current_page - 1)) + 1
    end_item_idx: int = min(app_state.per_page * current_page, total_items) # Ensure it doesn't exceed total items

    page_buttons: List[FT] = [] # Buttons like "1", "2", "..."

    # Page numbers to display (e.g., current page +/- 2, plus first/last and ellipses)
    # This logic can be complex to make it look good for all page counts.
    # Simplified version: always show first, last, current, current-1, current+1 if valid
    pages_to_show_set: Set[Union[int, str]] = {1, total_pages, current_page}
    if current_page > 1: pages_to_show_set.add(current_page - 1)
    if current_page < total_pages: pages_to_show_set.add(current_page + 1)

    # Add more pages around current if space allows (e.g. up to 7-9 total buttons)
    # This is a more complex part of pagination UI logic.
    # For now, keeping it simpler. Add more pages if needed based on UX.
    # Example: Add page 2 and total_pages-1 if not already included by proximity to current_page
    if total_pages > 3 : pages_to_show_set.add(2)
    if total_pages > 4 : pages_to_show_set.add(total_pages-1)

    sorted_page_numbers = sorted([p for p in list(pages_to_show_set) if isinstance(p,int) and 1 <= p <= total_pages])

    # Add ellipses "..." where there are gaps
    final_page_display_elements: List[Union[int,str]] = []
    last_num_added = 0
    for num in sorted_page_numbers:
        if num > last_num_added + 1 and last_num_added != 0 : # Check if there's a gap after the first element
            final_page_display_elements.append("...")
        final_page_display_elements.append(num)
        last_num_added = num


    # Create First & Previous buttons
    page_buttons.append(A("« First", hx_get=str(urls.URLS[Url.data_grid]), hx_target="#data-grid-component", hx_vals=json.dumps({"page": 1}), role="button", cls=btn_disabled_cls if current_page <= 1 else btn_enabled_cls))
    page_buttons.append(A("‹ Prev", hx_get=str(urls.URLS[Url.data_grid]), hx_target="#data-grid-component", hx_vals=json.dumps({"page": max(1, current_page - 1)}), role="button", cls=btn_disabled_cls if current_page <= 1 else btn_enabled_cls))

    # Create numbered page buttons
    for page_num_display in final_page_display_elements:
        if page_num_display == "...":
            page_buttons.append(Span("...", cls=ellipsis_cls))
        else: # It's an integer page number
            page_num_int = int(page_num_display)
            page_buttons.append(A(str(page_num_int), hx_get=str(urls.URLS[Url.data_grid]), hx_target="#data-grid-component", hx_vals=json.dumps({"page": page_num_int}), role="button", cls=num_selected_cls if page_num_int == current_page else num_other_cls))

    # Create Next & Last buttons
    page_buttons.append(A("Next ›", hx_get=str(urls.URLS[Url.data_grid]), hx_target="#data-grid-component", hx_vals=json.dumps({"page": min(total_pages, current_page + 1)}), role="button", cls=btn_disabled_cls if current_page >= total_pages else btn_enabled_cls))
    page_buttons.append(A("Last »", hx_get=str(urls.URLS[Url.data_grid]), hx_target="#data-grid-component", hx_vals=json.dumps({"page": total_pages}), role="button", cls=btn_disabled_cls if current_page >= total_pages else btn_enabled_cls))

    # Info text: "Showing X to Y of Z results"
    info_text = P(f"Showing {start_item_idx} to {end_item_idx} of {total_items} results.", cls="text-sm text-base-content/80")

    return Div(
        Div(cls="flex-1")(info_text), # Info text on one side
        Div(cls="join")( # Join class for button group styling from DaisyUI/Tailwind
            *page_buttons
        ), # Buttons on the other side
        cls="flex items-center justify-between w-full h-14 px-3 border-t border-base-300 bg-base-200/50", # Pagination bar styling
    )


def render_modal_field(col_name: str, value: Any) -> Tuple[FT, str]:
    """
    Renders a field's value for display within the item detail modal.
    Applies specific formatting for URLs, badges (based on BADGE_STYLES),
    or default text rendering.

    Args:
        col_name (str): The name of the column/field.
        value (Any): The value of the field.

    Returns:
        Tuple[FT, str]: A tuple containing:
            - FT: The `fasthtml` component representing the formatted value.
            - str: A string indicating the primary HTML tag used (e.g., "label", "a", "span").
                   This can be used by callers for layout decisions.
    """
    val_str = str(value) if value is not None else ""
    display_text = val_str if val_str else "N/A" # Default display for None or empty

    # URL handling
    if col_name == "url" and value:
        return A(UkIcon("external-link", cls="w-4 h-4 inline-block mr-1"), "Open Link", # Clearer text
                 href=str(value), target="_blank", rel="noopener noreferrer",
                 title=f"Open link: {value}",
                 cls="link link-hover text-primary hover:text-primary-focus text-sm"), "a"
    elif col_name == "url" and not value:
        return Span(UkIcon("unlink", cls="w-4 h-4 inline-block text-base-content/50 mr-1"), "No URL",
                    title="No URL provided", cls="text-sm text-base-content/70"), "span"

    # Badge/Pill styling for specific columns
    # This covers classification, status, workflow_status, and faculty via BADGE_STYLES
    style_lookup_key = col_name
    if "classification" in col_name or "prediction" in col_name: # Use "classification" styles for these
        style_lookup_key = "classification"

    if style_lookup_key in BADGE_STYLES:
        style_info = BADGE_STYLES[style_lookup_key].get(val_str, DEFAULT_PILL_STYLE) # Use full val_str for lookup
        badge_class = str(style_info) if isinstance(style_info, LabelT) else style_info # style_info could be a class string
        return Label(display_text, cls=f"{badge_class} badge-sm"), "label" # Using monsterui Label

    # Default rendering for other types
    if isinstance(value, (int, float)):
        # For numbers, could apply specific formatting if needed, for now, just styled span
        return Span(str(value), cls="text-sm font-mono"), "span"
    elif value is None or val_str == "N/A":
        return Span("N/A", cls="text-base-content/70 text-sm italic"), "span"
    else: # Default for other strings
        return Span(display_text, cls="text-sm break-words"), "span"


def render_labelled_item(label_text: str, content_component: FT, content_tag_type: str = "span") -> Div:
    """
    Helper to create a standard layout for a labelled item in the modal:
    Label on one line, content component below or beside it.

    Args:
        label_text (str): The text for the label (e.g., field name).
        content_component (FT): The `fasthtml` component representing the field's content.
        content_tag_type (str): Hint about the primary tag of `content_component` (e.g. "label", "a", "span")
                                used for minor layout adjustments.

    Returns:
        Div: A Div component containing the label and content.
    """
    label_element = fh.Strong(label_text, cls="block text-xs font-semibold text-base-content/90 mb-0.5 tracking-wide uppercase")

    # Base classes for content area
    content_classes = "p-1.5 px-2 border rounded-md border-base-300/70 bg-base-100/50 text-sm break-words min-h-[2.25rem] flex items-center"

    # If content is a badge-like label, allow it to define its own padding/height via its classes
    if content_tag_type == "label": # Indicates it's a badge/pill from render_modal_field
        content_element = Div(content_component, cls="flex items-center") # Just wrap
    else:
        content_element = Div(content_component, cls=content_classes)

    return Div(label_element, content_element, cls="mb-2") # Overall container for one labelled item


def create_editable_pill_div(
    current_value: Optional[Union[int, str, float]],
    label_text: str,
    field_name: str, # The database/model field name
    options_map: Dict[str, Union[str, LabelT]], # {option_value: style_enum_or_class, ...}
    material_id: int,
) -> Div:
    """
    Creates a UI component for an editable "pill" (badge/label) with a dropdown for changing its value.
    Used in the item detail modal for fields like workflow_status, manual_classification.

    Args:
        current_value: The current value of the field for the item.
        label_text (str): User-visible label for this field (e.g., "Workflow Status").
        field_name (str): The programmatic name of the field (e.g., "workflow_status").
        options_map (Dict[str, Union[str, LabelT]]): Possible values and their display styles.
        material_id (int): The ID of the copyright item being edited.

    Returns:
        Div: A Div component representing the editable pill with its dropdown.
    """
    # Get the initial display component for the current value
    # The 'html_tag_type' is not strictly needed by render_labelled_item anymore
    display_component, _ = render_modal_field(field_name, current_value)

    # Ensure the display component has the correct ID for HTMX updates and JS interaction
    pill_display_id = f"pill-display-{field_name}-{material_id}"
    if hasattr(display_component, 'attrs') and isinstance(display_component.attrs, dict):
        display_component.attrs['id'] = pill_display_id
    else: # Wrap if it's a simple string or non-FT object without attrs
        display_component = Span(display_component, id=pill_display_id)

    # Create dropdown items for each option
    dropdown_list_items: List[Li] = []
    for option_key, style_info in options_map.items():
        style_class_str = str(style_info) if isinstance(style_info, LabelT) else style_info

        # Data for HTMX request when an option is clicked
        htmx_values_dict = json.dumps({
            "material_id": material_id, "field_name": field_name, "value": str(option_key)
        })
        # JavaScript to optimistically update the pill and close dropdown
        # Note: ensure `updatePill` JS function is available on the page
        js_onclick_str = f"updatePill('{field_name}-{material_id}', '{str(option_key)}', '{str(option_key)}', '{style_class_str}'); return false;"

        dropdown_list_items.append(
            Li(A(Label(str(option_key), cls=f"{style_class_str} badge-sm w-full text-left"), # Full width label
                   href="#", onclick=js_onclick_str,
                   hx_post=str(urls.URLS[Url.update_single_field]), hx_vals=htmx_values_dict, hx_swap="none",
                   # hx_indicator=f"#{pill_display_id}" # Show loading state on the pill itself
                   cls="block w-full hover:bg-base-200 rounded p-1" # Styling for dropdown link
            ))
        )

    # Dropdown container (using UIkit attributes as per original, adapt if UI framework changes)
    dropdown_container = Div(
        Ul(*dropdown_list_items, cls="uk-nav uk-dropdown-nav p-1"), # Simple list for dropdown items
        cls="uk-dropdown w-auto bg-base-100 p-1 shadow-lg rounded-md border border-base-300 z-50", # Ensure z-index
        uk_drop="mode: click; pos: bottom-right; boundary: !.modal-box; flip: false; delay-hide: 0",
    )

    # Pill display area that also triggers the dropdown
    pill_trigger_area = Div(display_component, cls="inline-block uk-inline cursor-pointer")

    # Main container for label and pill+dropdown
    return Div(
        fh.Strong(label_text, cls="block text-xs font-semibold text-base-content/90 mb-0.5 tracking-wide uppercase"),
        Div(pill_trigger_area, dropdown_container, cls="text-right"), # Align pill to the right
        cls="flex items-center justify-between space-x-2 mb-2",
    )


def create_readonly_item_div(
    current_value: Optional[Union[int, str, float, bool, datetime.date, datetime.datetime]],
    label_text: str,
    field_name: str, # For context if specific styling needed, not used currently for readonly
) -> Div:
    """
    Creates a display component for a read-only field in the item detail modal.
    Formats the value appropriately.

    Args:
        current_value: The value of the field.
        label_text (str): User-visible label for this field.
        field_name (str): Programmatic name of the field (for context or future styling).

    Returns:
        Div: A Div component containing the label and formatted read-only value.
    """
    # Use render_modal_field to get consistent display formatting
    content_component, _ = render_modal_field(field_name, current_value)

    # Use render_labelled_item for consistent layout
    return render_labelled_item(label_text, content_component, content_tag_type="readonly")
```
