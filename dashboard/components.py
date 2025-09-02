# functions to build the components displayed in the dashboard
import datetime
import json

import fasthtml.common as fh
import polars as pl
from fastcore.utils import *
from fasthtml.common import *
from fasthtml.components import Button
from loguru import logger
from monsterui.all import *

import dashboard.urls as urls
from dashboard.constants import (
    BADGE_STYLES,
    DEFAULT_PILL_STYLE,
    DISPLAY_COLUMNS,
    FILTERABLE_COLUMNS,
    LabelT,
)
from dashboard.data import get_item_df
from dashboard.urls import Url
from dashboard.web import AppState
from easy_access.db.models import ItemUpdate
from easy_access.settings import SETTINGS, DirSetting


def ItemDetailCard(
    title: str,
    *body_content: Any,
    card_id: str,
    col_span: int = 1,
    start_collapsed: bool = False,
    color: str = "[var(--ring)]",
    lazy_load_url: str | None = None,
    tooltip: str | None = None,
) -> FT:
    """
    Creates a collapsible card component for the modal using <details>.
    Includes optional lazy loading via HTMX.
    """
    col_span = max(1, min(col_span, 3))
    col_span_class = f"md:col-span-{col_span}"

    border_color_class = (
        f"border-{color}-500" if color != "base" else "border-[var(--ring)]"
    )
    card_classes = f"bg-{color}-100 border {border_color_class} rounded-lg shadow-sm  {col_span_class}"

    summary_classes = (
        f"p-3 bg-{color}-100 cursor-pointer list-none flex items-center justify-between"
    )
    content_classes = "p-4 border-t border-[var(--ring)]"
    chevron_icon = Span(
        "▼", cls="text-xs transition-transform duration-200 chevron-icon"
    )

    content_target_id = f"{card_id}-content"
    summary_htmx_attrs = {}
    if lazy_load_url:
        # If lazy loading, initial content is a placeholder + indicator
        actual_body_content = (
            Span("Loading...", cls="italic text-sm text-base-content/70"),
            Div(id=f"{card_id}-loading", cls="htmx-indicator text-center p-2")(
                Span(cls="loading loading-sm")
            ),
        )
        # --- Set HTMX attributes on the <summary> tag ---
        summary_htmx_attrs = {
            "hx_get": lazy_load_url,
            "hx_target": f"#{content_target_id}",  # Target the inner div
            "hx_swap": "innerHTML",
            "hx_trigger": "click once",  # Trigger on first click of the summary
        }
    else:
        actual_body_content = body_content
    if not tooltip:
        tooltip = Span("")
    else:
        tooltip = UkIcon(
            "circle-help",
            cls="ml-1 font-sm tooltip tooltip-top cursor-help",
            height="12px",
            width="12px",
            title=tooltip,
        )
    content_div = Div(*actual_body_content, id=content_target_id, cls=content_classes)
    summary_element = Summary(
        H5(title, tooltip, cls="font-semibold text-sm m-0"),
        chevron_icon,
        cls=summary_classes,
        **summary_htmx_attrs,  # Add HTMX attributes HERE
    )
    return Details(
        summary_element,
        content_div,
        id=card_id,
        cls=card_classes,
        **({} if start_collapsed else {"open": True}),
    )


# Add this new component function
def create_checkbox_filter_group(
    filter_key: str,
    label_text: str,
    options: dict[str, str | LabelT],
    current_values: Optional[str],
    current_total: int,
    counts: Optional[dict[str, int | str]] = None,
) -> FT:
    """Creates a compact group of styled checkboxes with counts and HTMX trigger. Disabled options are minimized."""

    selected_values = set(current_values.split("|") if current_values else [])
    outer_group_id = f"filter-group-{filter_key}"

    form_field_name = f"filter_{filter_key}"
    sorted_options = {}
    enabled_options = {k: v for k, v in options.items() if k in selected_values}
    sorted_options.update(
        sorted(enabled_options.items(), key=lambda item: item[0].lower())
    )
    int_options = {
        k: v
        for k, v in options.items()
        if k not in selected_values and isinstance(counts.get(k, 0), int)
    }
    enabled_options.update(
        sorted(
            {
                k: v
                for k, v in int_options.items()
                if counts.get(k, 0) > 0 and counts.get(k, 0) != current_total
            }.items(),
            key=lambda item: item[0].lower(),
        )
    )
    disabled_options = {k: v for k, v in options.items() if k not in enabled_options}

    sorted_options.update(enabled_options)
    sorted_options.update(
        sorted(disabled_options.items(), key=lambda item: item[0].lower())
    )
    full_items = []
    disabled_list = []
    for value, style_info in sorted_options.items():
        safe_value = "".join(c if c.isalnum() else "-" for c in value)
        count = counts.get(value, "") if counts else ""
        is_checked = value in selected_values

        is_disabled = not is_checked if isinstance(count, str) else False
        if isinstance(count, str) and not selected_values and not is_disabled:
            count = 0

        # --- Style class logic ---
        checkboxcolor = "checkbox-accent"
        style_class = "badge badge-secondary"

        if filter_key != "faculty":
            style_class = str(style_info) if isinstance(style_info, LabelT) else ""
            match style_class:
                case "uk-label-primary":
                    checkboxcolor = "checkbox-accent"
                    style_class = "badge badge-neutral"
                case "uk-label-secondary":
                    checkboxcolor = "checkbox-primary"
                    style_class = "badge badge-primary"
                case "uk-label-destructive":
                    checkboxcolor = "checkbox-error"
                    style_class = "badge badge-error"

        count_span_id = f"count-{filter_key}-{safe_value}"

        # --- Enabled option rendering ---

        if not is_disabled:
            badge_cls = (
                f"{style_class} badge-outline ml-1 align-middle text-xs cursor-help"
            )
            count_badge = Span(
                f"{'▲' if all([isinstance(count, int), selected_values, count != current_total]) else '▼' if isinstance(count, int) and count != current_total else '='} {count}",
                title=f"{'Will increase the current total itemcount to this amount when selected.' if all([isinstance(count, int), selected_values, count != current_total]) else 'Will decrease the current total itemcount to this amount when selected.' if isinstance(count, int) and count != current_total else 'The current total itemcount'} ",
                id=count_span_id,
                cls=badge_cls,
            )

            hx_vals_dict = {
                f"filter_{filter_key}": value,  # Use the correct key
                "checked": str(not is_checked),  # Send the intended state
            }

            badge_label = Span(value, cls=f"{style_class} mr-1 cursor-pointer text-xs")
            checkbox_input = fh.CheckboxX(
                name=form_field_name,
                value=value,
                id=f"cb-{filter_key}-{safe_value}",
                checked=is_checked,
                cls=f"checkbox checkbox-xs align-middle mr-1 {checkboxcolor}",
                hx_post=urls.URLS[Url.data_grid],  # Target the central data route
                hx_target="#data-grid-component",  # Update the grid
                hx_swap="outerHTML",  # Replace the grid component
                hx_trigger="change",  # Trigger on change with delay
                hx_vals=json.dumps(hx_vals_dict),  # Send only the specific change
                disabled=False,
            )
            wrapper_label = fh.Label(
                checkbox_input,
                badge_label,
                count_badge,
                cls="flex items-center gap-x-1 gap-y-0 py-0 px-1 mb-0.5 rounded min-h-0",
                id=f"label-cb-{filter_key}-{safe_value}",
            )
            full_items.append(wrapper_label)
        else:
            # --- Minimized disabled option ---
            disabled_list.append(
                Span(
                    value if len(value) < 18 else value[:16] + "…",
                    cls="text-xs opacity-40 mr-2 cursor-help",
                    title=f"{value} ",
                )
            )
    # --- Card-like block for the filter group ---
    group_card = Div(
        fh.Label(
            label_text,
            cls="label-text pb-0 text-xs text-base-content/90 block mb-1 font-semibold tracking-wide",
        ),
        Div(*full_items, cls="grid grid-cols-1 gap-x-1 gap-y-0.5 mb-0.5"),
        # Collapsed disabled row
        (
            Div(
                UkIcon(
                    "minus-circle",
                    cls="text-xs opacity-40 cursor-help",
                    title="no items match",
                ),
                *disabled_list,
                cls="flex flex-wrap gap-x-1 gap-y-0.5 mt-1 mb-0.5",
            )
            if disabled_list
            else None
        ),
        cls="bg-base-100 border border-base-300 rounded-lg p-2 mb-2 shadow-xs",
        id=outer_group_id,
    )
    return group_card


def render_course_details(item_data):
    """
    Renders the contents for the course details card.
    This includes course name, year, programme, faculty, and related organizations.
    """
    course_details_items = []
    all_course_orgs = set()
    courses = item_data.get("courses") or []

    if not courses:
        course_details_items.append(
            P("No course data found.", cls="text-sm text-base-content/70")
        )
    else:
        for course in courses:
            course_details_items.append(
                Div(cls="mb-3 p-2 border rounded border-base-300")(
                    H6(
                        f"{course.get('name', 'Unknown Course')} ({course.get('cursuscode', '?')})",
                        cls="font-semibold text-sm mb-1",
                    ),
                    P(
                        f"Year: {course.get('year', 'N/A')}, Programme: {course.get('programme', 'N/A')}, Faculty: {course.get('faculty_id', 'N/A')}",
                        cls="text-xs text-base-content/80",
                    ),
                )
            )
            persons = course.get("persons") or []
            for person in persons:
                orgs = person.get("organizations") or []
                for org in orgs:
                    if org.get("abbreviation"):
                        all_course_orgs.add(org.get("full_abbreviation"))

    if all_course_orgs:
        course_details_items.append(Divider(cls="my-2"))
        course_details_items.append(
            Strong("Related Organizations:", cls="text-xs font-medium block mb-1")
        )
        course_details_items.append(
            Div(cls="flex flex-wrap gap-1")(
                *[
                    Label(org_abbr, cls="badge badge-sm " + LabelT.secondary)
                    for org_abbr in sorted(list(all_course_orgs))
                ]
            )
        )
    return tuple(course_details_items)


def render_contact_info(item_data):
    """
    Renders the contact info card contents
    This includes contact persons and their email addresses,
    and user-friendly buttons to copy email addresses to the clipboard.
    """
    contact_persons = []
    all_contact_emails = set()
    courses = item_data.get("courses") or []
    for course in courses:
        persons = course.get("persons") or []
        for person in persons:
            if person.get("role") == "contact":
                contact_persons.append(person)
                if person.get("email"):
                    all_contact_emails.add(person["email"])

    contact_info_items = []
    emails_str = ";".join(all_contact_emails)

    copy_js = f"navigator.clipboard.writeText('{emails_str}');"
    contact_info_items.append(
        Div(
            fh.Button(
                UkIcon("copy", cls="w-4 h-4 mr-1"),
                "Copy all email addresses",
                cls="btn btn-primary btn-sm btn-block cursor-copy",
                onclick=copy_js,
            )
        )
    )

    if not contact_persons:
        contact_info_items.append(
            P("No contact persons found.", cls="text-sm text-base-content/70")
        )
    else:
        added_person_ids = set()
        for person in contact_persons:
            person_id = person.get("id")
            if person_id in added_person_ids:
                continue
            added_person_ids.add(person_id)

            email = person.get("email")
            people_page_url = person.get("people_page_url")
            orgs = person.get("organizations") or []
            highest_level = -1
            top_org_abbrs = set()
            for org in orgs:
                level = org.get("hierarchy_level")
                if level is not None:
                    if level > highest_level:
                        highest_level = level
                        top_org_abbrs = {org.get("full_abbreviation", "?")}
                    elif level == highest_level:
                        top_org_abbrs.add(org.get("full_abbreviation", "?"))

            contact_info_items.append(
                Div(
                    cls="flex items-center justify-between space-x-2 py-1 border-b border-base-200 last:border-b-0"
                )(
                    Span(person.get("main_name", "Unknown Name"), cls="text-sm"),
                    Div(cls="flex items-center space-x-1 flex-shrink-0")(
                        *[
                            Label(abbr, cls="badge badge-sm " + LabelT.secondary)
                            for abbr in sorted(list(top_org_abbrs))
                        ],
                        A(
                            UkIcon("mail", cls="w-4 h-4"),
                            href=f"mailto:{email}",
                            title=f"Email {email}",
                            cls="link text-primary",
                        )
                        if email
                        else Span(
                            UkIcon("mail-question", cls="w-4 h-4 text-base-content/50"),
                            title="No email",
                        ),
                        A(
                            UkIcon("external-link", cls="w-4 h-4"),
                            href=people_page_url,
                            target="_blank",
                            title="Open people page",
                            cls="link text-primary",
                        )
                        if people_page_url
                        else Span(
                            UkIcon("link-2-off", cls="w-4 h-4  text-base-content/50"),
                            title="No people page URL",
                        ),
                    ),
                )
            )
    return tuple(contact_info_items)


def render_teacher_info(item_data):
    teacher_persons = {}
    courses = item_data.get("courses") or []
    for course in courses:
        persons = course.get("persons") or []
        for person in persons:
            person_id = person.get("id")
            if person_id and person_id not in teacher_persons:
                teacher_persons[person_id] = person

    teacher_info_items = []
    if not teacher_persons:
        teacher_info_items.append(
            P("No teachers/persons found.", cls="text-sm text-base-content/70")
        )
    else:
        sorted_teachers = sorted(
            teacher_persons.values(), key=lambda p: p.get("main_name", "")
        )
        for person in sorted_teachers:
            people_page_url = person.get("people_page_url")
            orgs = person.get("organizations") or []
            highest_level = -1
            top_org_abbrs = set()
            for org in orgs:
                level = org.get("hierarchy_level")
                if level is not None:
                    if level > highest_level:
                        highest_level = level
                        top_org_abbrs = {org.get("full_abbreviation", "?")}
                    elif level == highest_level:
                        top_org_abbrs.add(org.get("full_abbreviation", "?"))

            teacher_info_items.append(
                Div(
                    cls="flex items-center justify-between space-x-2 py-1 border-b border-base-200 last:border-b-0"
                )(
                    Span(person.get("main_name", "Unknown Name"), cls="text-sm"),
                    Div(cls="flex items-center space-x-1 flex-shrink-0")(
                        *[
                            Label(abbr, cls="badge badge-sm " + LabelT.secondary)
                            for abbr in sorted(list(top_org_abbrs))
                        ],
                        A(
                            UkIcon("external-link", cls="w-4 h-4"),
                            href=people_page_url,
                            target="_blank",
                            title="Open people page",
                            cls="link text-primary",
                        )
                        if people_page_url
                        else Span(
                            UkIcon("link-2-off", cls="w-4 h-4 text-base-content/50"),
                            title="No people page URL",
                        ),
                    ),
                )
            )
    return tuple(teacher_info_items)


def render_item_history(history: list[ItemUpdate]):
    """
    For a given list of ItemUpdate objects, renders a card showing the changes made to the item.
    """

    timeline_elements = []
    if not history:
        return Div("No history found for this item.")

    logger.debug(f"Rendering item history for {len(history)} items")
    all_data: list[dict[str, str | dict[str, str]] | int | datetime.datetime] = [
        item.__dict__.get("change_details", {}) for item in history
    ]
    all_data.sort(
        key=lambda x: x.get("update_time", ""), reverse=True
    )  # start with most recent
    for data in all_data:
        if not data:
            continue
        date = data.get("update_time", "?")
        changes = {}
        for change in data:
            if change in [
                "material_id",
                "update_time",
                "modified_at",
                "possible_fine",
                "infringement",
                "retrieved_from_copyright_on",
            ]:
                continue
            if not isinstance(data[change], dict):
                new_value = data[change]
                old_value = "N/A"
                changes[change] = (old_value, new_value)
                continue
            old_value = data[change].get("old", "N/A")
            if old_value in [None, "None", ""]:
                old_value = "N/A"
            new_value = data[change].get("new", "N/A")
            if new_value in [None, "None", ""]:
                new_value = "N/A"

            changes[change] = (old_value, new_value)
        if not changes:
            continue
        max_key_len = max(len(key) for key in changes)
        try:
            change_date = datetime.datetime.strptime(date[:19], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            change_date = "?"
            logger.warning(f"Warning: Invalid date format '{date}'. Using placeholder.")
        edit_content = [
            f"<span class='badge badge-xs  badge-secondary text-[10px]'>{change_date}</span>"
        ]
        edit_content.extend(
            [
                f"<span class='flex flex-col ml-2'><span class='badge badge-xs badge-outline badge-secondary text-[10px]'>{key.replace('_', ' '):>{max_key_len}}</span><span class='ml-2'><span class='text-[10px] badge badge-xs badge-outline badge-accent'>{value[0]}</span>  →  <span class='text-[10px] badge badge-xs badge-primary'>{value[1]}</span></span></span>"
                for key, value in changes.items()
            ]
        )
        edit_content = "".join(edit_content)

        timeline_elements.append(
            NotStr(f"""
        <li>
            <div class="flex flex-col font-mono text-[10px]">{edit_content}</div>
        </li>
        """)
        )

    return ItemDetailCard(
        "Change history",
        Ul(
            *timeline_elements,
            cls="flex flex-col gap-2 list-none p-0 m-0",
        ),
        card_id="item-history-card",
        col_span=1,
        start_collapsed=True,
        tooltip="All changes made to this item, in reverse chronological order.",
    )


def render_table_rows(
    df_slice: pl.DataFrame,
    app_state: AppState,
) -> tuple[FT, ...]:
    """Renders Tbody rows with custom formatting and HTMX attributes."""
    rows: list[FT] = []
    if df_slice is not None and df_slice.height > 0:
        cols_to_display = [col for col in DISPLAY_COLUMNS if col in df_slice.columns]
        display_df_slice = df_slice.select(cols_to_display)

        for row_dict in display_df_slice.iter_rows(named=True):
            cells = []
            material_id = row_dict.get("material_id", None)
            if material_id is None:
                continue

            for col in cols_to_display:
                val = row_dict.get(col)
                td_class = f"col-{col.replace('_', '-')}"
                display_text = str(val) if val else "N/A"

                if col == "url":
                    if val:
                        cell_content = A(
                            UkIcon("external-link", cls="w-4 h-4 inline-block"),
                            href=val,
                            target="_blank",
                            rel="noopener noreferrer",
                            title=f"Open link: {val}",
                            cls="link link-hover text-primary hover:text-primary-focus",
                            onclick="event.stopPropagation()",
                        )
                    else:
                        cell_content = Span(
                            UkIcon(
                                "unlink",
                                cls="w-4 h-4 inline-block text-base-content/50",
                            ),
                            title="No URL provided",
                        )
                elif (
                    col in BADGE_STYLES
                    or "classification" in col
                    or "prediction" in col
                ):
                    if "classification" in col or "prediction" in col:
                        style = BADGE_STYLES["classification"].get(
                            val, DEFAULT_PILL_STYLE
                        )
                    else:
                        style = BADGE_STYLES.get(col, {}).get(val, DEFAULT_PILL_STYLE)

                    if "-" in display_text:
                        print_txt = display_text.split("-")[0].strip()
                    else:
                        print_txt = display_text.strip()

                    cell_content = Label(
                        print_txt,
                        title=display_text,
                        cls=style + "badge-sm cursor-help",
                        onclick="event.stopPropagation()",
                    )

                else:
                    if len(display_text) > 20:
                        cell_content = display_text[:20] + "..."
                    else:
                        cell_content = display_text

                cells.append(
                    Td(
                        cell_content,
                        title=display_text,
                        cls=f"py-2.5 px-4 border-b border-base-200/80 {td_class} cursor-help",
                    )
                )
            modal_url = urls.URLS[Url.show_item_details].replace(
                "{material_id:int}", str(material_id)
            )
            row_attrs: dict[str, Any] = {
                "id": f"row-{material_id}",
                "class": "hover:bg-primary/10 cursor-pointer transition-colors duration-150",
                "hx_get": modal_url,
                "hx_target": "#modal-placeholder",
                "hx_swap": "innerHTML",
                "preload": True,
            }

            rows.append(Tr(*cells, **row_attrs))
    if not rows:
        rows.append(
            Tr(
                Td(
                    "No matching data found.",
                    colspan=str(len(DISPLAY_COLUMNS)),
                    class_="text-center text-sm p-6",
                )
            )
        )
    return tuple(rows)


def page_header_component(
    user_details: dict[str, str],
    app_state: AppState,
    filter_counts: dict[str, dict[str, int]],
    total_filtered_rows: int,  # <-- ADDED parameter
) -> FT:
    """
    Renders the header area with Title, User Info, and a Sidebar for Filters.
    Now accepts total_filtered_rows directly.
    """
    current_filters = app_state.get_active_filters()
    auth_details = user_details
    is_admin = auth_details.get("role") == "admin"
    user_faculty = auth_details.get("faculty")

    # --- Filter Checkbox Groups ---
    filter_checkbox_groups = [
        create_checkbox_filter_group(
            filter_key="workflow_status",
            label_text="Workflow Status",
            options=BADGE_STYLES["workflow_status"],
            current_values=current_filters.get("workflow_status"),
            counts=filter_counts.get("workflow_status"),
            current_total=total_filtered_rows,  # <-- Use parameter here
        ),
        create_checkbox_filter_group(
            filter_key="status",
            label_text="Status",
            options=BADGE_STYLES["status"],
            current_values=current_filters.get("status"),
            counts=filter_counts.get("status"),
            current_total=total_filtered_rows,  # <-- Use parameter here
        ),
        create_checkbox_filter_group(
            filter_key="classification",
            label_text="Classification",
            options=BADGE_STYLES["classification"],
            current_values=current_filters.get("classification"),
            counts=filter_counts.get("classification"),
            current_total=total_filtered_rows,  # <-- Use parameter here
        ),
        create_checkbox_filter_group(
            filter_key="manual_classification",
            label_text="Manual Classification",
            options=BADGE_STYLES["classification"],
            current_values=current_filters.get("manual_classification"),
            counts=filter_counts.get("manual_classification"),
            current_total=total_filtered_rows,  # <-- Use parameter here
        ),
    ]
    # Conditionally add faculty filter group
    if is_admin or not user_faculty or user_faculty == "all":
        filter_checkbox_groups.append(
            create_checkbox_filter_group(
                filter_key="faculty",
                label_text="Faculty",
                options=BADGE_STYLES["faculty"],
                current_values=current_filters.get("faculty"),
                counts=filter_counts.get("faculty"),
                current_total=total_filtered_rows,  # <-- Use parameter here
            )
        )

    # --- Filter Inputs (Text - Remaining - unchanged HTMX setup) ---
    # ... (text_filter_inputs logic remains the same) ...
    text_filter_inputs = []
    remaining_filterable_cols = [
        col
        for col in FILTERABLE_COLUMNS
        if col
        not in [
            "workflow_status",
            "status",
            "classification",
            "manual_classification",
            "faculty",
        ]
    ]
    for col in remaining_filterable_cols:
        if col not in get_item_df().columns:
            continue
        label = f"{col.replace('_', ' ').title()}"
        input_id = f"filter-input-{col}"
        input_name = f"filter_{col}"
        hx_vals_js = f"{{'{input_name}': document.getElementById('{input_id}').value}}"
        text_filter_inputs.append(
            LabelInput(
                label,
                name=input_name,
                value=current_filters.get(col, ""),
                placeholder="Filter...",
                id=input_id,
                input_cls="input input-bordered input-xs w-full focus:input-primary",
                label_cls="label-text pb-1 text-xs font-medium text-base-content/90",
                cls="form-control w-full",
                hx_post=urls.URLS[Url.data_grid],
                hx_trigger="keyup changed delay:500ms, search",
                hx_target="#data-grid-component",
                hx_vals=hx_vals_js,
            )
        )

    # --- Filter Form (unchanged structure, just uses corrected groups/inputs) ---
    filter_form = Form(
        Div(  # Reset button container
            fh.Button(
                UkIcon("rotate-ccw", cls="mr-1"),
                "Reset",
                type="button",
                hx_post=urls.URLS[Url.data_grid],
                hx_target="#data-grid-component",
                hx_vals='{"action": "reset"}',
                cls="btn btn-xs btn-warning mx-1",
                **{"@click": "slideOverOpen=false"},
            ),
            cls="sticky top-0 z-20 flex justify-center gap-x-2 bg-base-200 py-2 border-b border-base-300 mb-1",
        ),
        *filter_checkbox_groups,
        *text_filter_inputs,
        id="filter-form",
        cls="flex flex-col gap-y-1",
    )

    # --- Sidebar Structure (Alpine.js - Unchanged) ---
    # ... (sidebar_component logic remains the same, containing filter_form) ...
    sidebar_component = Div(
        x_data="{ slideOverOpen: false }", cls="relative z-50 w-auto h-auto"
    )(
        Div(
            fh.Button(
                UkIcon("filter", cls="mr-1"),
                "Filter",
                **{"@click": "slideOverOpen=true"},
                cls="btn btn-xs btn-primary mx-1",
            ),
            fh.Button(
                UkIcon("rotate-ccw", cls="mr-1"),
                "Reset",
                type="button",
                hx_post=urls.URLS[Url.data_grid],
                hx_target="#data-grid-component",
                hx_vals='{"action": "reset"}',
                cls="btn btn-xs btn-warning mx-1",
            ),
            cls="flex flex-col items-center justify-center gap-y-2 me-2",
        ),
        Template(x_teleport="body")(
            Div(
                x_show="slideOverOpen",
                **{"@keydown.window.escape": "slideOverOpen=false"},
                cls="relative z-[99]",
                x_cloak=True,
                hx_ext="preload",
            )(
                Div(
                    x_show="slideOverOpen",
                    x_transition_opacity_duration_600ms=True,
                    **{"@click": "slideOverOpen = false"},
                    cls="fixed inset-0 bg-black bg-opacity-10",
                ),
                Div(cls="fixed inset-0 overflow-hidden")(
                    Div(cls="absolute inset-0 overflow-hidden")(
                        Div(cls="fixed inset-y-0 left-0 flex max-w-full pr-10")(
                            Div(
                                x_show="slideOverOpen",
                                **{"@click.away": "slideOverOpen = false"},
                                x_transition_enter="transform transition ease-in-out duration-500 sm:duration-700",
                                x_transition_enter_start="-translate-x-full",
                                x_transition_enter_end="translate-x-0",
                                x_transition_leave="transform transition ease-in-out duration-500 sm:duration-700",
                                x_transition_leave_start="translate-x-0",
                                x_transition_leave_end="-translate-x-full",
                                cls="w-screen max-w-md",
                            )(
                                Div(
                                    cls="flex flex-col h-full py-2 overflow-y-auto bg-base-200 border-r border-base-300 shadow-lg"
                                )(
                                    Div(cls="px-2 sm:px-2")(
                                        Div(
                                            cls="flex items-start justify-between pb-1"
                                        )(
                                            H2(
                                                "Filters",
                                                cls="text-lg font-semibold leading-6 text-base-content",
                                            ),
                                            Div(cls="flex items-center h-auto ml-3")(
                                                Button(
                                                    **{"@click": "slideOverOpen=false"},
                                                    cls="btn btn-sm btn-circle btn-ghost",
                                                )("✕")
                                            ),
                                        )
                                    ),
                                    Div(cls="relative flex-1 px-2 mt-2")(filter_form),
                                )
                            )
                        )
                    )
                ),
            )
        ),
    )

    update_times = [
        file.created
        for file in SETTINGS.dirs[DirSetting.RAW_COPYRIGHT_DATA].files
        if file.exists
    ]
    last_update_time = (
        max(update_times).strftime("%Y-%m-%d %H:%M") if update_times else "Unknown"
    )
    return Div(
        Div(cls="flex justify-between items-center mb-4")(
            Div(
                H4("Copyright Data Dashboard", cls="text-2xl font-bold text-primary"),
                H6(
                    f"Last update from Qlik: {last_update_time}",
                    cls="text-xs text-base-content/70 mt-1",
                ),
            ),
            Div(cls="flex items-center space-x-4")(
                sidebar_component,
                Div(cls="flex items-center space-x-2")(
                    Img(
                        src=f"https://api.dicebear.com/9.x/bottts/svg?seed={user_details.get('name', 'defaultuser')}&backgroundColor=547012",
                        alt="avatar",
                        cls="w-8 h-8 rounded-full border border-base-300",
                    ),
                    Div(
                        Span(
                            f"Hi, {user_details.get('name', 'User')}!",
                            cls="text-sm font-medium block",
                        ),
                        Span(
                            f"Viewing: {user_details.get('faculty', 'All items')}"
                            if user_details.get("faculty") != "all"
                            else "Viewing: All items",
                            cls="text-xs text-base-content/70 block",
                        ),
                    ),
                ),
                A("Logout", href="/logout", cls="btn btn-outline btn-primary btn-sm"),
            ),
        ),
        cls="mb-4 border-b border-base-300 pb-4",
    )


def render_data_grid_component(
    df_slice: pl.DataFrame,
    app_state: AppState,  # Accept AppState object
    total_filtered_rows: int,
    total_pages: int,
) -> FT:
    """Renders the main content area: Item count, table, pagination."""

    header_cells = []
    cols_in_header = [
        col for col in DISPLAY_COLUMNS if col in get_item_df().columns
    ]  # Use global to know columns

    for col in cols_in_header:
        header_text = "ID" if col == "material_id" else col.replace("_", " ").title()

        sort_indicator_icon = ""

        next_sort_desc_bool = False  # Default sort is ASC if not current col
        is_current_sort_col = col == app_state.sort_by

        if is_current_sort_col:
            # If current, toggle direction
            next_sort_desc_bool = not app_state.sort_desc
            sort_indicator_icon = "arrow-down" if app_state.sort_desc else "arrow-up"
        # else: next_sort_desc_bool remains False (sort ASC)

        # Prepare hx-vals for sorting
        sort_link_params = {
            "sort_by": col,
            "sort_desc": str(next_sort_desc_bool),
            # page will be reset server-side if needed, don't send it
        }

        header_content = Button(
            header_text,
            (
                UkIcon(
                    sort_indicator_icon,
                    cls="inline-block ml-1.5 w-3 h-3 stroke-current",
                )
                if sort_indicator_icon
                else ""
            ),
            # --- Updated HTMX Attributes for Sorting ---
            hx_get=urls.URLS[Url.data_grid],  # Target central data route
            hx_target="#data-grid-component",
            hx_vals=json.dumps(sort_link_params),  # Send only sort parameters
            # hx_include REMOVED
            cls=f"{ButtonT.ghost} text-xs uppercase tracking-wider p-1.5 h-auto min-h-0 font-bold text-base-content/70 hover:text-primary transition-colors duration-150",
        )
        th_class = f"col-{col.replace('_', '-')}"
        header_cells.append(Th(header_content, cls=f"px-4 py-2 {th_class}"))

    header = Thead(Tr(*header_cells))

    # Body rendering relies on render_table_rows, which itself uses show_item_details route.
    # Ensure show_item_details logic is updated.
    body = Tbody(
        *render_table_rows(
            df_slice, app_state
        ),  # Pass app_state if needed by row rendering
        id="data-table-body",
    )

    table_wrapper = Div(
        Table(header, body, role="grid", id="data-table", cls="table table-sm w-full"),
        id="table-wrapper",
        cls="overflow-x-auto border border-base-300 rounded-lg shadow-sm bg-base-100",
    )

    # Pagination rendering uses the updated render_pagination function
    pagination_html = render_pagination(app_state, total_pages, total_filtered_rows)

    # --- Final Component ---
    return Div(
        pagination_html,  # Pagination controls
        Div(  # Table Wrapper with loading effect
            table_wrapper,
            cls="htmx-request:opacity-50 htmx-request:transition-opacity htmx-request:duration-300 ease-in-out",  # Use htmx-request class
        ),
        id="data-grid-component",  # Target ID for updates
        cls="flex flex-col justify-center",  # Keep existing layout classes
    )


def render_pagination(
    app_state: AppState,  # Accept AppState object
    total_pages: int,
    total_items: int,
) -> FT:
    """
    Renders pagination controls using AppState and updated HTMX attributes.
    Uses tailwind classes from pines component library to display the page counts, first/prev/next/last buttons,
    and an array of directly selectable page numbers surrounding the current page.
    """

    if total_pages <= 1:
        return Div(cls="h-12 mb-1")

    classes = {
        "btn_enabled": "relative inline-flex items-center h-full px-3 rounded-l group hover:bg-blue-600 hover:text-white font-mono",
        "btn_disabled": "relative inline-flex items-center h-full px-3 rounded-l group bg-neutral-100 font-mono",
        "number_other": "relative inline-flex items-center h-full px-3 group hover:bg-blue-600 hover:text-white font-mono",
        "number_selected": "relative inline-flex items-center h-full px-3 text-white bg-blue-600 group font-mono",
        "disabled": "relative inline-flex items-center h-full px-2.5 bg-neutral-100 group font-mono",
    }
    current_page = app_state.page
    start_index = app_state.per_page * (current_page - 1) + 1
    end_index = app_state.per_page * current_page

    numbers_to_add: list[int] = []

    if total_pages <= 11:
        numbers_to_add = list(range(1, total_pages + 1))
    else:
        numbers_to_add.append(current_page)
        for i in range(1, 5):
            if current_page - i > 0:
                numbers_to_add.append(current_page - i)
        for i in range(1, 5):
            if current_page + i <= total_pages:
                numbers_to_add.append(current_page + i)

        numbers_to_add = sorted(set(numbers_to_add))

        if numbers_to_add[0] > 1:
            numbers_to_add.insert(0, "...")
        if numbers_to_add[-1] < total_pages:
            numbers_to_add.append("...")
        if len(numbers_to_add) < 11:
            missing_amount = 11 - len(numbers_to_add)

            if isinstance(numbers_to_add[-1], str) or (
                isinstance(numbers_to_add[0], int)
                and numbers_to_add[0] <= missing_amount
            ):
                if numbers_to_add[-1] == "...":
                    numbers_to_add.pop()
                start_num = numbers_to_add[-1]
                for i in range(1, missing_amount + 1):
                    if start_num + i <= total_pages:
                        numbers_to_add.append(start_num + i)
            elif isinstance(numbers_to_add[0], str) or (
                isinstance(numbers_to_add[-1], int)
                and numbers_to_add[-1] >= total_pages - missing_amount
            ):
                if numbers_to_add[0] == "...":
                    numbers_to_add.pop(0)
                start_num = numbers_to_add[0]
                for i in range(1, missing_amount + 1):
                    if start_num - i > 0:
                        numbers_to_add.insert(0, start_num - i)

        if isinstance(numbers_to_add[0], int) and numbers_to_add[0] > 1:
            numbers_to_add.insert(0, "...")
        if isinstance(numbers_to_add[-1], int) and numbers_to_add[-1] < total_pages:
            numbers_to_add.append("...")

    btn_block = []
    # 1. create first + prev buttons
    attr_dict = {}
    attr_dict["first"] = (
        "⏮",
        {
            "hx_get": urls.URLS[Url.data_grid],
            "hx_target": "#data-grid-component",
            "hx_vals": json.dumps({"page": 1}),
            "role": "button",
            "cls": classes["btn_disabled"]
            if current_page <= 1
            else classes["btn_enabled"],
        },
    )
    attr_dict["prev"] = (
        "⏴",
        {
            "hx_get": urls.URLS[Url.data_grid],
            "hx_target": "#data-grid-component",
            "hx_vals": json.dumps({"page": max(1, current_page - 1)}),
            "role": "button",
            "cls": classes["btn_disabled"]
            if current_page <= 1
            else classes["btn_enabled"],
        },
    )
    # add the buttons to the block
    for _, (text, attr) in attr_dict.items():
        btn_block.append(
            A(
                text,
                **attr,
                preload=True,
            )
        )
    # 2. add the numbered elements
    for number in numbers_to_add:
        if number == "...":
            cls = classes["disabled"]
        elif number == current_page:
            cls = classes["number_selected"]
        else:
            cls = classes["number_other"]

        btn_block.append(
            Li(
                A(
                    number,
                    hx_get=urls.URLS[Url.data_grid],
                    hx_target="#data-grid-component",
                    hx_vals=json.dumps({"page": number}),
                    role="button",
                    cls=cls,
                    preload=True,
                ),
                cls="hidden h-full md:block",
            )
        )
    # 3. create next + last buttons
    attr_dict = {}
    attr_dict["next"] = (
        "⏵",
        {
            "hx_get": urls.URLS[Url.data_grid],  # Target the central data route
            "hx_target": "#data-grid-component",
            "hx_vals": json.dumps(
                {"page": min(total_pages, current_page + 1)}
            ),  # Send only the page change
            "role": "button",
            "cls": classes["btn_disabled"]
            if current_page >= total_pages
            else classes["btn_enabled"],
        },
    )

    attr_dict["last"] = (
        "⏭",
        {
            "hx_get": urls.URLS[Url.data_grid],  # Target the central data route
            "hx_target": "#data-grid-component",
            "hx_vals": json.dumps({"page": total_pages}),  # Send only the page change
            "role": "button",
            "cls": classes["btn_disabled"]
            if current_page >= total_pages
            else classes["btn_enabled"],
        },
    )
    # add the buttons to the block
    for _, (text, attr) in attr_dict.items():
        btn_block.append(
            A(
                text,
                **attr,
            )
        )
    #  create the final elements and return
    nav_element = Nav(
        Div(
            Span(
                NotStr(
                    f"<span class='font-mono'>🗏 {start_index}-{end_index} (<span class='font-mono'>∑ {total_items}</span>)</span>"
                ),
                title="items shown / total items",
                cls="pl-1 pr-1 mr-2 cursor-help border border-blue-700/20 rounded bg-blue-700/10",
            ),
            Span(
                NotStr(f"<span class='font-mono'>🗐 {total_pages}</span>"),
                title=" total amount of  pages",
                cls="pl-1 pr-1 ml-2 cursor-help border border-green-700/20 rounded bg-emerald-700/10",
            ),
        ),
        Ul(
            *btn_block,
            cls="flex items-center text-sm leading-tight w-full bg-base-200 border border-neutral-200/70 rounded h-[34px] text-neutral-500 ",
        ),
        cls="flex flex-col items-center justify-cente",
    )
    pagination_text_element = ()

    return Div(
        pagination_text_element,
        nav_element,
        cls="flex items-center justify-end w-full h-16 px-3 border-t border-neutral-200",
    )


def render_modal_field(col_name: str, value: Any) -> tuple[FT, str]:
    """Renders a field for the modal, handling pills, icons, and text.
    Returns the core component (Label, A, Span, or raw value string/int)."""
    val_str = str(value) if value is not None else ""
    val_str_norm = val_str.lower().strip()
    display_text = val_str if val_str else "N/A"

    style = DEFAULT_PILL_STYLE

    if col_name == "url":
        if value:
            return A(
                UkIcon("external-link", cls="w-4 h-4 inline-block"),
                href=value,
                target="_blank",
                rel="noopener noreferrer",
                title=f"Open link: {value}",
                cls="link link-hover text-primary hover:text-primary-focus",
            ), "a"
        else:
            return Span(
                UkIcon("unlink", cls="w-4 h-4 inline-block text-base-content/50"),
                title="No URL provided",
            ), "span"

    elif isinstance(value, int | float):
        return Label(value, cls=style + " badge-sm"), "label"
    elif col_name in BADGE_STYLES or col_name in [
        "ml_classification",
        "classification",
        "manual_classification",
        "ml_prediction",
    ]:
        if "classification" in col_name or "prediction" in col_name:
            style = BADGE_STYLES["classification"].get(val_str_norm, DEFAULT_PILL_STYLE)
        else:
            style = BADGE_STYLES.get(col_name, {}).get(value, DEFAULT_PILL_STYLE)

        return Label(display_text, cls=style + " badge-sm"), "label"
    elif value is None:
        return Span("N/A", cls="text-base-content/70 text-sm"), "span"
    else:
        return Span(display_text, cls="text-sm break-words"), "span"


def render_labelled_item(
    label_text: str,
    content_component: FT,
    html_tag: str = "label",
) -> tuple[FT, FT]:
    """Helper to create a standard label Strong tag and return the content component.
    Layout is handled by the caller."""
    label_element = Strong(
        label_text, cls="block text-sm font-medium text-base-content/80 mb-1"
    )
    if html_tag == "number":
        content_element = Div(
            str(content_component),
            cls="p-1.5 px-2 border rounded border-blue-300 bg-base-200 text-sm break-words",
        )
    elif html_tag == "span":
        content_element = Div(
            content_component,
            cls="p-1.5 px-2 border rounded border-blue-300 bg-base-200 text-sm break-words",
        )
    else:
        content_element = content_component

    return label_element, content_element


def create_editable_pill_div(
    current_value: int | str | float | None,
    label_text: str,
    field_name: str,
    options_map: dict,
    material_id: int,  # Add material_id parameter
):
    content_component, html_tag = render_modal_field(field_name, current_value)
    label_el, _ = render_labelled_item(label_text, content_component, html_tag)

    # Logic to find the *actual* style class applied by render_modal_field
    if hasattr(content_component, "attrs") and "cls" in content_component.attrs:
        current_classes = content_component.attrs["cls"].split()
        labelt_values = (
            {str(lt) for lt in LabelT}
            | {str(st) for st in BADGE_STYLES["status"].values()}
            | {str(wt) for wt in BADGE_STYLES["workflow_status"].values()}
        )
        found_style = next(
            (cls for cls in current_classes if cls in labelt_values), None
        )
        if (
            not found_style
            and "badge-sm" not in current_classes
            and isinstance(content_component, FT)
            and (
                content_component.tag == "span"
                and "uk-label" in content_component.attrs.get("cls", "")
            )
        ):
            content_component.attrs["cls"] += " badge-sm"

    component_with_id = content_component
    target_id = f"pill-display-{field_name}-{material_id}"

    if isinstance(component_with_id, FT) and hasattr(component_with_id, "attrs"):
        if not isinstance(component_with_id.attrs, dict):
            component_with_id.attrs = {}
        component_with_id.attrs["id"] = target_id
        logger.debug(
            f"Assigned ID '{target_id}' to existing FT component: {component_with_id.tag}"
        )
    elif isinstance(component_with_id, str | int | float) or component_with_id is None:
        component_with_id = Label(
            str(component_with_id) if component_with_id is not None else "",
            id=target_id,
        )
        logger.debug(f"Wrapped simple content in Label with ID '{target_id}'")
    else:
        logger.warning(
            f"Warning: Wrapping unknown component type for '{field_name}' in Label with ID '{target_id}'"
        )
        component_with_id = Label(component_with_id, id=target_id)

    dropdown_items = []
    for opt_val, opt_style_enum in options_map.items():
        opt_style_class = (
            str(opt_style_enum) if opt_style_enum else str(DEFAULT_PILL_STYLE)
        )
        js_opt_val = json.dumps(str(opt_val))
        js_opt_text = json.dumps(str(opt_val))  # Text is usually same as value here
        onclick_js = f"updatePill('{field_name}', {js_opt_val}, {js_opt_text}, '{opt_style_class}'); return false;"
        htmx_vals = json.dumps(
            {
                "material_id": material_id,
                "field_name": field_name,
                "value": str(opt_val),
            }
        )
        dropdown_items.append(
            Li(
                A(
                    Label(opt_val, cls=f"{opt_style_class} badge-sm"),
                    href="#",
                    onclick=onclick_js,
                    hx_post=urls.URLS[Url.update_single_field],
                    hx_vals=htmx_vals,
                    hx_swap="none",
                    # Optional: Add indicator during request
                    hx_indicator=f"#pill-display-{field_name}",
                )
            )
        )

    pill_container = Div(
        component_with_id,  # The component with the ID set
        Div(
            Ul(*dropdown_items, cls="uk-nav uk-dropdown-nav"),
            cls="uk-dropdown w-auto bg-base-100 p-2 shadow-lg rounded-md border border-base-300",
            uk_drop="mode: click; pos: bottom-right; boundary: !.modal-box; flip: false",
        ),
        cls="inline-block uk-inline",
    )

    return Div(
        label_el,
        Div(pill_container, cls="text-right"),
        cls="flex items-center justify-between space-x-2 mb-2",
    )


def create_readonly_item_div(
    current_value: int | str | float | None,
    label_text: str,
    field_name: str,
):
    content_component, html_tag = render_modal_field(field_name, current_value)
    label_el, content_el = render_labelled_item(label_text, content_component, html_tag)
    is_inline = field_name in [
        "classification",
        "manual_classification",
        "ml_prediction",
        "pagecount",
        "wordcount",
        "picturecount",
    ]
    container_cls = (
        "flex items-center justify-between space-x-2 mb-2" if is_inline else "mb-3"
    )
    content_wrapper_cls = "text-right" if is_inline else ""
    return Div(label_el, Div(content_el, cls=content_wrapper_cls), cls=container_cls)
