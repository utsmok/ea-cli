# dash.py

import contextlib
import json
import math
import traceback
from enum import auto
from typing import Any

import fasthtml.common as fh
import polars as pl
from fastcore.xml import FT
from fasthtml.common import *
from fasthtml.components import Htmx_toasts
from monsterui.all import *
from monsterui.foundations import VEnum, str2ukcls
from rich import print

from easy_access.db.retrieve import retrieve_copyright_items
from easy_access.db.update import update_copyright_items

# --- monsterui fixes ---


class LabelT(VEnum):
    """FIX: replaces danger with destructive"""

    def _generate_next_value_(name, start, count, last_values):
        return str2ukcls("label", name)

    primary = auto()
    secondary = auto()
    destructive = auto()


# NOTE: do a manual replace of this function to fix theme switching
def _headers_theme(
    color, mode="auto", radii=ThemeRadii.sm, shadows=ThemeShadows.sm, font=ThemeFont.sm
):
    """
    replace existing function in monsterui core.py with this to enable proper theme switching
    """
    franken_init = (
        'const __FRANKEN__ = JSON.parse(localStorage.getItem("__FRANKEN__") || "{}");'
    )
    mode_script = {
        "auto": f"""
            {franken_init}
        if (
            __FRANKEN__.mode === "dark" ||
            (!__FRANKEN__.mode &&
            window.matchMedia("(prefers-color-scheme: dark)").matches)
        ) {{
            htmlElement.classList.add("dark");
        }} else {{
            htmlElement.classList.remove("dark");
        }}

        """,
        "light": f'{franken_init} htmlElement.classList.remove("dark");',
        "dark": f'{franken_init} htmlElement.classList.add("dark");',
    }
    return_val = "const htmlElement = document.documentElement;"
    return_val += f"\n{mode_script[mode]}"
    return_val += (
        f'\nhtmlElement.classList.add("uk-theme-{color}");'
        if color
        else "htmlElement.classList.add(__FRANKEN__.theme);"
    )
    return_val += (
        f'\nhtmlElement.classList.add("{radii}")'
        if radii
        else "htmlElement.classList.add(__FRANKEN__.radii);"
    )
    return_val += (
        f'\nhtmlElement.classList.add("{font}")'
        if font
        else "htmlElement.classList.add(__FRANKEN__.font);"
    )
    return_val += (
        f'\nhtmlElement.classList.add("{shadows}")'
        if shadows
        else "htmlElement.classList.add(__FRANKEN__.shadows);"
    )

    return fh.Script(return_val)


# --- app setup ---
GLOBAL_STYLES = Style("""
    * {
        font-family: "Inter", sans-serif;
        font-optical-sizing: auto;
        font-weight: 400;
        font-style: normal;
    }
    html, body { scrollbar-gutter: auto !important; height: 100%; overflow: hidden; margin: 0; padding: 0; background-color: hsl(var(--b2)); }
    #page-container { display: flex; flex-direction: column; height: 100vh; background-color: hsl(var(--b1)); }
    #content-area {  padding: 1rem 1.5rem; flex-grow: 1; overflow-y: auto; overflow-x: hidden; }
    #table-wrapper { overflow-x: auto; overflow-y: hidden; }
    #data-table thead th { position: sticky; top: 0; z-index: 10; background-color: hsl(var(--b2)/0.95); backdrop-filter: blur(4px);  }
    #data-table th, #data-table td {
    white-space: nowrap;
    padding: 0.5rem 0.75rem;
    }

    /* Data table */

    /* column widths */
    #data-table .col-manual-classification { width: 180px; text-align: center; }
    #data-table .col-ml-classification { width: 180px; text-align: center; }
    #data-table .col-id { width: 90px; }
    #data-table .col-url { width: 60px; text-align: center; }
    #data-table .col-workflow-status { width: 130px; text-align: center; }
    #data-table .col-status { width: 130px; text-align: center; }
    #data-table .col-remarks {
        white-space: normal;
        min-width: 150px;
    }
    #data-table .col-title {
        white-space: normal;
        min-width: 150px;
    }
    #data-table .col-filename {
        white-space: normal;
        min-width: 150px;
    }
    /* Pill styling  */
    #data-table .badge { /* If using daisyUI badges instead of Label */
        min-width: 80px;
        display: inline-block;
    }

    #modal-loading-indicator {
        opacity: 0;
        transition: opacity 200ms ease-in;
        pointer-events: none;
    }

    .htmx-request #modal-loading-indicator,
    .htmx-request#modal-loading-indicator {
        opacity: 1;
        pointer-events: auto;
    }




    /* Custom color style */

    .uk-theme-slate  {
        --background: 78 0% 94.46%;
        --foreground: 78 70% 1%;
        --muted: 198 20% 85%;
        --muted-foreground: 198 0% 40%;
        --popover: 78 0% 97.08%;
        --popover-foreground: 78 70% 1%;
        --card: 0 0% 96%;
        --card-foreground: 78 70% 0%;
        --border: 78 0% 74.31%;
        --input: 308.06 63.49% 43.57%;
        --primary: 78 70% 30%;
        --primary-foreground: 0 0% 100%;
        --secondary: 198 70% 30%;
        --secondary-foreground: 0 0% 100%;
        --accent: 318 70% 30%;
        --accent-foreground: 318 70% 90%;
        --destructive: 3 96.56% 36.63%;
        --destructive-foreground: 3 0% 100%;
        --ring: 304.55 61.68% 41.96%;
        --chart-1: 78 70% 30%;
        --chart-2: 198 70% 30%;
        --chart-3: 318 70% 30%;
        --chart-4: 198 70% 33%;
        --chart-5: 78 73% 30%;
        --radius: 0.5rem;



    }
    .dark.uk-theme-slate  {
        --background: 78 59% 4%;
        --foreground: 78 32% 99%;
        --muted: 198 20% 15%;
        --muted-foreground: 198 0% 60%;
        --popover: 78 59% 4%;
        --popover-foreground: 78 32% 99%;
        --card: 78 59% 5%;
        --card-foreground: 0 0% 100%;
        --border: 78 14% 13%;
        --input: 78 14% 13%;
        --primary: 98.02 55.69% 34.09%;
        --primary-foreground: 0 0% 100%;
        --secondary: 198 70% 30%;
        --secondary-foreground: 0 0% 100%;
        --accent: 318 70% 30%;
        --accent-foreground: 318 70% 90%;
        --destructive: 3 100% 63.57%;
        --destructive-foreground: 0 0% 100%;
        --ring: 78 70% 30%;
        --chart-1: 78 70% 30%;
        --chart-2: 198 70% 30%;
        --chart-3: 318 70% 30%;
        --chart-4: 198 70% 33%;
        --chart-5: 78 73% 30%;
    }
""")


app, rt = fast_app(
    hdrs=(
        Theme.slate.headers(
            mode="light",
            daisy=True,
            radii=ThemeRadii.lg,
            shadows=ThemeShadows.lg,
        ),
        GLOBAL_STYLES,
        Link(
            rel="stylesheet",
            href="https://fonts.googleapis.com/css2?family=Inter:ital,opsz,wght@0,14..32,100..900;1,14..32,100..900&display=swap",
            type="text/css",
        ),
        Script(
            src="https://unpkg.com/@htmx/htmx-toasts@latest/dist/index.js",
            type="module",
        ),
    ),
    exts="loading-states",
)

DEFAULT_PER_PAGE = 15
MAX_CELL_LENGTH = 35

WORKFLOW_STYLES = {
    "ToDo": LabelT.destructive,
    "InProgress": LabelT.secondary,
    "Done": LabelT.primary,
}
STATUS_STYLES = {
    "Published": LabelT.primary,
    "Unpublished": LabelT.secondary,
    "Deleted": LabelT.destructive,
}
PRIMARY_CLASSIFICATIONS = {
    "open access",
    "eigen materiaal - powerpoint",
    "eigen materiaal - titelindicatie",
    "eigen materiaal - overig",
    "eigen materiaal",
}
SECONDARY_CLASSIFICATIONS = {
    "onbekend",
    "niet geanalyseerd",
    "in onderzoek",
    "licentie beschikbaar",
    "verwijderverzoek verstuurd",
}
DESTRUCTIVE_CLASSIFICATIONS = {
    "korte overname",
    "middellange overname",
    "lange overname",
}
DEFAULT_PILL_STYLE = LabelT.secondary  # Use secondary as default for unknown pills


# --- data setup ---
copyright_df_global: pl.DataFrame = retrieve_copyright_items()

FILTERABLE_COLUMNS: list[str] = [
    "department",
    "course_name",
    "manual_classification",
    "ml_prediction",
    "status",
    "faculty",
    "author",
    "publisher",
    "workflow_status",
]
DISPLAY_COLUMNS: list[str] = [
    "material_id",
    "url",
    "workflow_status",
    "status",
    "ml_prediction",
    "manual_classification",
    "remarks",
    "filename",
    "title",
    "faculty",
]


def get_filtered_sorted_df(
    sort_by: str | None = None,
    sort_desc: bool = False,
    filters: dict[str, str] | None = None,
) -> pl.DataFrame:
    """Applies filtering and sorting."""
    df = copyright_df_global
    if filters:
        filter_expressions = []
        for col, value in filters.items():
            actual_col = col
            if value and actual_col in df.columns:
                try:
                    if df[actual_col].dtype == pl.Utf8:
                        if col == "status":
                            filter_expressions.append(
                                pl.col(actual_col).str.starts_with(value.capitalize())
                            )
                        else:
                            filter_expressions.append(
                                pl.col(actual_col).str.contains(f"(?i){value}")
                            )
                    elif df[actual_col].dtype in (
                        pl.Int64,
                        pl.Int32,
                        pl.Float64,
                        pl.Float32,
                    ):
                        with contextlib.suppress(ValueError):
                            filter_expressions.append(
                                pl.col(actual_col) == float(value)
                            )
                except Exception as e:
                    print(f"Filter warning on '{actual_col}': {e}")
        if filter_expressions:
            try:
                df = df.filter(pl.all_horizontal(filter_expressions))
            except Exception as e:
                print(f"Filter error: {e}")
                df = copyright_df_global
    if sort_by and sort_by in df.columns:
        try:
            df = df.sort(by=sort_by, descending=sort_desc, nulls_last=True)
        except Exception as e:
            print(f"Sort warning on '{sort_by}': {e}")
    return df


async def store_item_changes(
    input_data: list[dict[str, str | int]] | dict[str, str | int],
) -> None:
    """
    For a list of dicts containing at least a material_id and updated fields,
    for each item, retrieve the full item data for the other fields from the copyright_df_global
    make sure the updated fields in the dict are not overwritten with the original values

    then send the new list of dicts to the database for update
    """
    full_data_list = []
    if not isinstance(input_data, list):
        input_data = [input_data]
    for item in input_data:
        material_id = item.get("material_id")
        if material_id is not None:
            full_item_data = copyright_df_global.filter(
                pl.col("material_id") == material_id
            ).to_dicts()
            if full_item_data:
                print(f"Updating material_id {material_id}")
                full_item_data = full_item_data[0]
                # update fields that are in the input item
                for key, value in item.items():
                    if key == "material_id":
                        continue
                    if key == "faculty":
                        # special case for faculty, we need to update the faculty_id instead
                        full_item_data["faculty_id"] = value
                        del full_item_data["faculty"]
                        continue
                    if key in full_item_data:
                        if value != full_item_data[key]:
                            print(f"Updating {key}: {full_item_data[key]} --> {value}")
                        full_item_data[key] = value
                full_data_list.append(full_item_data)

    await update_copyright_items(full_data_list, update_relations=False, overwrite=True)


# --- Component rendering ---


def render_table_rows(
    df_slice: pl.DataFrame,
    current_page_for_modal: int,
    per_page_for_modal: int,
    sort_by_for_modal: str | None,
    sort_desc_for_modal: bool,
    filters_for_modal: dict[str, str] | None,
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
                val_str_norm = str(val).lower().strip() if val is not None else ""
                cell_content: Any = str(val) if val is not None else ""
                td_class = f"col-{col.replace('_', '-')}"

                if col == "url":
                    if val:
                        cell_content = A(
                            UkIcon("external-link", cls="w-4 h-4 inline-block"),
                            href=val,
                            target="_blank",
                            rel="noopener noreferrer",
                            title=f"Open link: {val}",
                            cls="link link-hover text-primary hover:text-primary-focus",
                        )
                    else:
                        cell_content = Span(
                            UkIcon(
                                "unlink",
                                cls="w-4 h-4 inline-block text-base-content/50",
                            ),
                            title="No URL provided",
                        )
                elif col == "workflow_status":
                    style = WORKFLOW_STYLES.get(str(val))
                    cell_content = Label(
                        str(val) if val else "N/A", cls=style + " badge-sm"
                    )
                elif col == "status":
                    style = STATUS_STYLES.get(str(val))
                    cell_content = Label(
                        str(val) if val else "N/A", cls=style + " badge-sm"
                    )

                elif col in ("manual_classification", "ml_prediction"):
                    if not val_str_norm:
                        display_text = "-"
                        style = ""
                    elif val_str_norm in PRIMARY_CLASSIFICATIONS:
                        style = LabelT.primary
                        display_text = str(val)
                    elif val_str_norm in SECONDARY_CLASSIFICATIONS:
                        style = LabelT.secondary
                        display_text = str(val)
                    elif val_str_norm in DESTRUCTIVE_CLASSIFICATIONS:
                        style = LabelT.destructive
                        display_text = str(val)
                    else:
                        display_text = str(val)
                        style = ""

                    cell_content = Label(display_text, cls=style + " badge-sm")

                cells.append(
                    Td(
                        cell_content,
                        cls=f"py-2.5 px-4 border-b border-base-200/80 {td_class}",
                    )
                )

            modal_params = {
                "material_id": material_id,
                "page": current_page_for_modal,
                "per_page": per_page_for_modal,
                "sort_by": sort_by_for_modal or "",
                "sort_desc": str(sort_desc_for_modal),
                **{f"filter_{k}": v for k, v in filters_for_modal.items()},
            }
            row_attrs: dict[str, Any] = {
                "id": f"row-{material_id}",
                "class": "hover:bg-primary/10 cursor-pointer transition-colors duration-150",
                "hx_get": show_item_details.to(**modal_params),
                "hx_target": "#modal-placeholder",
                "hx_swap": "innerHTML",
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
    current_filters: dict[str, str],
    per_page: int,
    sort_by: str | None,
    sort_desc: bool,
) -> FT:
    """Renders the header area with Title and Filters."""
    current_filters = current_filters or {}
    filter_inputs = []
    for col in FILTERABLE_COLUMNS:
        actual_col_name = col
        if actual_col_name not in copyright_df_global.columns:
            continue
        filter_inputs.append(
            LabelInput(
                f"{col.replace('_', ' ').title()}",
                name=f"filter_{col}",
                value=current_filters.get(col, ""),
                placeholder="Filter...",
                id=f"filter-input-{col}",
                hx_get=data_grid.to(),
                hx_trigger="keyup changed delay:500ms, search",
                hx_target="#data-grid-component",
                hx_include="closest form",
                hx_indicator="#grid-loading-indicator",
                input_cls="input input-bordered input-xs w-full focus:input-primary",
                label_cls="label-text pb-1 text-xs font-medium text-base-content/90",
                cls="form-control w-full",
            )
        )

    filter_form = Form(
        Div(
            *filter_inputs,
            cls="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-x-4 gap-y-3",
        ),
        Input(type="hidden", name="page", value="1"),
        Input(type="hidden", name="per_page", value=str(per_page)),
        Input(type="hidden", name="sort_by", value=sort_by or ""),
        Input(type="hidden", name="sort_desc", value=str(sort_desc)),
        id="filter-form",
        cls="",
    )

    return Div(
        Div(
            H4(
                "Copyright Data Dashboard",
                cls="text-4xl font-bold underline decoration-pink-500 text-primary",
            ),
            cls="flex-shrink-0 mr-8 pt-1",
        ),
        Div(filter_form, cls="flex-grow justify-center"),
        cls="flex items-start mb-4 border-b border-base-300 pb-3",
    )


def render_data_grid_component(
    df_slice: pl.DataFrame,
    current_page: int,
    total_pages: int,
    per_page: int,
    total_filtered_rows: int,
    offset: int,
    sort_by: str | None = None,
    sort_desc: bool = False,
    current_filters: dict[str, str] | None = None,
) -> FT:
    """Renders the main content area: Item count, table, pagination."""
    current_filters = current_filters or {}
    print(f"filters in render_data_grid: {current_filters}")

    header_cells = []
    cols_in_header = [
        col for col in DISPLAY_COLUMNS if col in copyright_df_global.columns
    ]
    for col in cols_in_header:
        header_text = "ID" if col == "material_id" else col.replace("_", " ").title()

        sort_indicator_icon, next_sort_desc = "", "False"
        if col == sort_by:
            sort_indicator_icon, next_sort_desc = (
                ("arrow-down", "False") if sort_desc else ("arrow-up", "True")
            )
        filter_params = {f"filter_{k}": v for k, v in current_filters.items() if v}
        sort_link_params = {
            "page": current_page,
            "per_page": per_page,
            "sort_by": col,
            "sort_desc": next_sort_desc,
            **filter_params,
        }
        header_content = Button(
            header_text,
            UkIcon(
                sort_indicator_icon, cls="inline-block ml-1.5 w-3 h-3 stroke-current"
            )
            if sort_indicator_icon
            else "",
            hx_get=data_grid.to(**sort_link_params),
            hx_target="#data-grid-component",
            hx_indicator="#grid-loading-indicator",
            cls=f"{ButtonT.ghost} text-xs uppercase tracking-wider p-1.5 h-auto min-h-0 font-bold text-base-content/70 hover:text-primary transition-colors duration-150",
        )
        th_class = f"col-{col.replace('_', '-')}"
        header_cells.append(Th(header_content, cls=f"px-4 py-2 {th_class}"))
    print(f"filters in render_data_grid: {current_filters}")

    header = Thead(Tr(*header_cells))
    body = Tbody(
        *render_table_rows(
            df_slice,
            current_page_for_modal=current_page,
            per_page_for_modal=per_page,
            sort_by_for_modal=sort_by,
            sort_desc_for_modal=sort_desc,
            filters_for_modal=current_filters,
        ),
        id="data-table-body",
    )

    table_wrapper = Div(
        Table(header, body, role="grid", id="data-table", cls="table table-sm w-full"),
        id="table-wrapper",
        cls="overflow-x-auto border border-base-300 rounded-lg shadow-sm bg-base-100",
    )

    pagination_html = render_pagination(
        current_page, total_pages, per_page, sort_by, sort_desc, current_filters
    )

    return Div(
        P(
            Strong(str(total_filtered_rows)),
            Span(" items found", cls="text-base-content/70"),
            cls=f"{TextPresets.muted_sm} mb-2 text-sm",
        ),
        pagination_html,
        Div(
            table_wrapper,
            cls="htmx-loading:opacity-50 htmx-loading:transition-opacity htmx-loading:duration-300 ease-in-out",
        ),
        Div(id="grid-loading-indicator", cls="htmx-indicator text-center p-4")(
            "Loading..."
        ),
        id="data-grid-component",
        cls="flex flex-col text-center justify-center",
    )


def render_pagination(
    current_page: int,
    total_pages: int,
    per_page: int,
    sort_by: str | None = None,
    sort_desc: bool = False,
    current_filters: dict[str, str] | None = None,
) -> FT:
    """Renders pagination controls using flexbox. (Unchanged)"""
    if total_pages <= 1:
        return Div(cls="h-12")
    current_filters = current_filters or {}
    filter_params = {f"filter_{k}": v for k, v in current_filters.items() if v}
    base_params = {
        "per_page": per_page,
        "sort_by": sort_by or "",
        "sort_desc": str(sort_desc),
        **filter_params,
    }
    print(f"filters in render_pagination: {current_filters}")

    pagination_items = []

    prev_disabled, prev_page = current_page <= 1, max(1, current_page - 1)
    prev_attrs = {
        "hx_get": data_grid.to(page=prev_page, **base_params),
        "hx_target": "#data-grid-component",
        "hx_indicator": "#grid-loading-indicator",
        "role": "button",
    }
    pagination_items.append(
        A(
            "« Previous",
            **(
                {"aria-disabled": "true", "cls": "btn btn-sm btn-disabled"}
                if prev_disabled
                else {"cls": f"btn btn-sm {ButtonT.primary}", **prev_attrs}
            ),
        )
    )

    pagination_items.append(
        Span(
            f"Page {current_page} of {total_pages} ",
            cls="underline decoration-pink-500 bg-base-200 text-base-content font-semibold mr-5 ml-5 mb-1",
        )
    )

    next_disabled, next_page = (
        current_page >= total_pages,
        min(total_pages, current_page + 1),
    )
    print("base_params in render_pagination:", base_params)
    next_attrs = {
        "hx_get": data_grid.to(page=next_page, **base_params),
        "hx_target": "#data-grid-component",
        "hx_indicator": "#grid-loading-indicator",
        "role": "button",
    }
    pagination_items.append(
        A(
            "Next »",
            **(
                {"aria-disabled": "true", "cls": "btn btn-sm btn-disabled mb-1"}
                if next_disabled
                else {"cls": f"btn btn-sm {ButtonT.primary} mb-1", **next_attrs}
            ),
        )
    )

    return Div(
        *pagination_items,
        aria_label="pagination",
        class_="flex justify-center items-center space-x-5 gap-4 pt-4 pb-2",
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
    elif col_name == "workflow_status":
        style = WORKFLOW_STYLES.get(val_str, DEFAULT_PILL_STYLE)
        return Label(display_text, cls=style + " badge-sm"), "label"
    elif col_name == "status":
        style = STATUS_STYLES.get(val_str, DEFAULT_PILL_STYLE)
        return Label(display_text, cls=style + " badge-sm"), "label"
    elif col_name in ("classification", "manual_classification", "ml_prediction"):
        if not val_str_norm:
            return Span(display_text, cls="text-base-content/70 text-sm"), "span"
        elif val_str_norm in PRIMARY_CLASSIFICATIONS:
            style = LabelT.primary
        elif val_str_norm in SECONDARY_CLASSIFICATIONS:
            style = LabelT.secondary
        elif val_str_norm in DESTRUCTIVE_CLASSIFICATIONS:
            style = LabelT.destructive
        return Label(display_text, cls=style + " badge-sm"), "label"

    elif value is None:
        return Span("N/A", cls="text-base-content/70 text-sm"), "span"
    elif isinstance(value, int | float):
        return value, "number"
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
        label_text, cls="block text-xs font-medium text-base-content/80 mb-1"
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


# --- Routes ---


@rt("/modal/{material_id:int}")
async def show_item_details(
    material_id: int,
    page: int = 1,
    per_page: int = DEFAULT_PER_PAGE,
    sort_by: str | None = None,
    sort_desc: str = "False",
    filter_department: str = "",
    filter_course_name: str = "",
    filter_manual_classification: str = "",
    filter_ml_prediction: str = "",
    filter_status: str = "",
    filter_faculty: str = "",
    filter_author: str = "",
    filter_publisher: str = "",
    filter_workflow_status: str = "",
):
    """Fetches data and returns structured INNER content for the modal dialog
    using direct Tailwind classes."""

    def get_val(key, default=None):
        return item_data.get(key, default)

    def create_editable_pill_div(label_text: str, field_name: str, options_map: dict):
        current_value = get_val(field_name)

        content_component, html_tag = render_modal_field(field_name, current_value)
        label_el, _ = render_labelled_item(label_text, content_component, html_tag)

        original_value_str = str(current_value) if current_value is not None else ""
        original_style_class = DEFAULT_PILL_STYLE + " badge-sm"
        component_with_id = content_component

        if "attrs" in content_component.__dict__:
            attrs_orig = content_component.__dict__["attrs"]

            attrs_new = attrs_orig.copy() if isinstance(attrs_orig, dict) else {}
            attrs_new["id"] = f"pill-display-{field_name}"

            component_with_id.__setattr__("attrs", attrs_new)

            current_classes = attrs_new.get("class", "").split()
            labelt_values = {str(lt) for lt in LabelT}
            found_style = next(
                (cls for cls in current_classes if cls in labelt_values), None
            )
            if found_style:
                original_style_class = found_style

        dropdown_items = []
        for opt_val, opt_style_enum in options_map.items():
            opt_style_class = str(opt_style_enum)
            onclick_js = f"updatePill('{field_name}', {repr(str(opt_val))}, {repr(str(opt_val))}, '{opt_style_class}'); return false;"
            dropdown_items.append(
                Li(
                    A(
                        Label(opt_val, cls=opt_style_enum + " badge-sm"),
                        href="#",
                        onclick=onclick_js,
                    )
                )
            )

        hidden_input = Input(
            type="hidden",
            id=f"input-{field_name}",
            name=field_name,
            value=original_value_str,
            data_original_value=original_value_str,
            data_original_text=original_value_str,
            data_original_style=original_style_class,
        )

        pill_wrapper = Div(cls="inline-block uk-inline")(
            Button(
                component_with_id,  #
                type="button",
                cls="p-0 m-0 bg-transparent border-none hover:opacity-80 focus:outline-none ring-0",
            ),
            Div(
                Ul(*dropdown_items, cls="uk-nav uk-dropdown-nav"),
                cls="uk-dropdown w-auto z-10",
                uk_drop="mode: click; pos: bottom-right",
            ),
            hidden_input,
        )

        return Div(
            label_el,
            Div(pill_wrapper, cls="text-right"),
            cls="flex items-center justify-between space-x-2 mb-2",
        )

    def create_readonly_item_div(
        label_text: str, field_name: str, is_inline: bool = False
    ):
        content_component, html_tag = render_modal_field(
            field_name, get_val(field_name)
        )
        label_el, content_el = render_labelled_item(
            label_text, content_component, html_tag
        )
        container_cls = (
            "flex items-center justify-between space-x-2 mb-2" if is_inline else "mb-3"
        )
        content_wrapper_cls = "text-right" if is_inline else ""
        return Div(
            label_el, Div(content_el, cls=content_wrapper_cls), cls=container_cls
        )

    try:
        current_modal_filters: dict[str, str] = {
            "department": filter_department.strip(),
            "course_name": filter_course_name.strip(),
            "manual_classification": filter_manual_classification.strip(),
            "ml_prediction": filter_ml_prediction.strip(),
            "status": filter_status.strip(),
            "faculty": filter_faculty.strip(),
            "author": filter_author.strip(),
            "publisher": filter_publisher.strip(),
            "workflow_status": filter_workflow_status.strip(),
        }
        nav_params = {
            "page": page,
            "per_page": per_page,
            "sort_by": sort_by or "",
            "sort_desc": sort_desc,
            **{f"filter_{k}": v for k, v in current_modal_filters.items()},
        }
        print(f"nav_params in show_item_details: {nav_params}")
        print(f"current_modal_filters in show_item_details: {current_modal_filters}")
        current_modal_filters = {k: v for k, v in current_modal_filters.items() if v}
        modal_sort_desc_bool = sort_desc.lower() == "true"
        ordered_df = get_filtered_sorted_df(
            sort_by, modal_sort_desc_bool, current_modal_filters
        )
        ordered_ids = ordered_df.get_column("material_id").to_list()
        try:
            current_index = ordered_ids.index(material_id)
        except ValueError:
            current_index = -1
        prev_id = ordered_ids[current_index - 1] if current_index > 0 else None
        next_id = (
            ordered_ids[current_index + 1]
            if current_index != -1 and current_index < len(ordered_ids) - 1
            else None
        )

        item_df = ordered_df.filter(pl.col("material_id") == material_id)
        if item_df.height == 0:
            return (modal_box_content, modal_backdrop), HtmxResponseHeaders(
                trigger="openModalEvent"
            )

        item_data = item_df.to_dicts()[0]

<<<<<<< HEAD
=======
        def get_val(key, default=None):
            return item_data.get(key, default)

        # --- Helper Function for Dropdown Pills ---
        def create_editable_pill_div(
            label_text: str, field_name: str, options_map: dict
        ):
            current_value = get_val(field_name)
            # Get the original component (e.g., a Label tuple) and its type hint
            content_component, _ = render_modal_field(field_name, current_value)
            label_el, _ = render_labelled_item(
                label_text, content_component
            )  # Get label

            # Store original value and style for reset
            original_value_str = str(current_value) if current_value is not None else ""
            original_style_class = DEFAULT_PILL_STYLE + " badge-sm"  # Default
            # Safer check for FT tuple and attributes dictionary
            if (
                isinstance(content_component, tuple)
                and len(content_component) == 3
                and isinstance(content_component[2], dict)
            ):
                current_classes = content_component[2].get("cls", "").split()
                labelt_values = {str(lt) for lt in LabelT}
                found_style = next(
                    (cls for cls in current_classes if cls in labelt_values), None
                )
                if found_style:
                    original_style_class = (
                        found_style  # Store only the base style class
                    )

            # Dropdown Items
            dropdown_items = []
            for opt_val, opt_style_enum in options_map.items():
                opt_style_class = str(opt_style_enum)
                # Use repr for safe JS string quoting
                onclick_js = f"updatePill('{field_name}', {repr(str(opt_val))}, {repr(str(opt_val))}, '{opt_style_class}'); return false;"
                dropdown_items.append(
                    Li(
                        A(
                            Label(opt_val, cls=opt_style_enum + " badge-sm"),
                            href="#",
                            onclick=onclick_js,
                        )
                    )
                )

            # Hidden input to store current value
            hidden_input = Input(
                type="hidden",
                id=f"input-{field_name}",
                name=field_name,
                value=original_value_str,
                data_original_value=original_value_str,
                data_original_text=original_value_str,  # Store original text for reset
                data_original_style=original_style_class,  # Store original style for reset
            )

            # --- FIX: Reconstruct the component with the ID ---
            component_with_id = (
                content_component  # Default to original if not a modifiable FT tuple
            )
            if isinstance(content_component, tuple) and len(content_component) == 3:
                tag, children, attrs_orig = content_component
                # Ensure attrs_orig is a dict before copying, handle None case
                attrs_new = attrs_orig.copy() if isinstance(attrs_orig, dict) else {}
                attrs_new["id"] = f"pill-display-{field_name}"  # Add the ID
                # Create the new tuple
                component_with_id = (tag, children, attrs_new)
            # --- END FIX ---

            # Wrap the pill in a Button for dropdown trigger
            pill_wrapper = Div(
                cls="inline-block uk-inline"
            )(  # Added uk-inline for dropdown positioning
                Button(
                    component_with_id,  # Use the reconstructed component with the ID
                    type="button",
                    cls="p-0 m-0 bg-transparent border-none hover:opacity-80",  # Make button invisible
                ),
                # Use Div for dropdown content as DropDownNavContainer might add unwanted styles
                Div(
                    Ul(
                        *dropdown_items, cls="uk-nav uk-dropdown-nav"
                    ),  # Standard UIKit nav list
                    cls="uk-dropdown",  # Basic dropdown class
                    uk_drop="mode: click; pos: bottom-right",  # Ensure uk-drop attribute is set
                ),
                hidden_input,
            )

            # Apply inline layout classes
            return Div(
                label_el,
                Div(pill_wrapper, cls="text-right"),
                cls="flex items-center justify-between space-x-2 mb-2",
            )

        # --- End Dropdown Helper ---

        # --- Create Combined Options for Manual Classification ---
>>>>>>> 3bebb0213a1f6c944bb768dbb8f650e564cfe9d4
        manual_classification_options = {}
        for val in PRIMARY_CLASSIFICATIONS:
            manual_classification_options[val] = LabelT.primary
        for val in SECONDARY_CLASSIFICATIONS:
            manual_classification_options[val] = LabelT.secondary
        for val in DESTRUCTIVE_CLASSIFICATIONS:
            manual_classification_options[val] = LabelT.destructive
        if "onbekend" not in manual_classification_options:
            manual_classification_options["onbekend"] = LabelT.secondary

        # --- 1. Header Row ---
        filename = get_val("filename", "N/A") or "(file deleted or not found)"
        filename_content = H4(filename, cls="font-semibold text-lg break-all")
        file_url = get_val("url")
        if not file_url or file_url == "":
            file_url = None
        if file_url:
            filename_content = A(
                filename_content, href=file_url, target="_blank", cls="link-hover"
            )
        status_pill, _ = render_modal_field("status", get_val("status"))
        url_element, _ = render_modal_field("url", file_url)

        header_content = Div(cls="flex items-center justify-between space-x-4")(
            Div(cls="flex items-center space-x-3 flex-grow min-w-0")(
                status_pill,
                filename_content,
                url_element,
            ),
            Form(method="dialog")(
                Button("✕", cls="btn btn-sm btn-circle btn-ghost flex-shrink-0")
            ),
        )

        # --- 2/3. Main Content & Details (Wrap in Form) ---
        block_classes = "border border-[hsl(var(--ring))] rounded-md p-4 shadow-sm"

        # Form wrapper for editable fields + hidden ID + save/reset
        modal_form_content = Form(
            # Hidden input for material_id MUST be inside the form
            Input(type="hidden", name="material_id", value=material_id),
            Input(type="hidden", name="page", value=nav_params["page"]),
            Input(type="hidden", name="per_page", value=nav_params["per_page"]),
            Input(type="hidden", name="sort_by", value=nav_params["sort_by"]),
            Input(type="hidden", name="sort_desc", value=nav_params["sort_desc"]),
            *[
                Input(type="hidden", name=f"filter_{k}", value=v)
                for k, v in current_modal_filters.items()
            ],
            # Main Content Grid (Item Info + Data Entry)
            Div(cls="grid grid-cols-1 md:grid-cols-2 gap-4")(
                # Left Column (Item Info Block - Readonly within form)
                Div(cls=f"{block_classes}")(
                    create_readonly_item_div(
                        "Classification", "classification", is_inline=True
                    ),
                    create_readonly_item_div(
                        "ML Prediction", "ml_prediction", is_inline=True
                    ),
                    create_readonly_item_div("Period", "period"),
                    create_readonly_item_div("Faculty", "faculty"),
                    create_readonly_item_div("Owner", "owner"),
                    create_readonly_item_div("Department", "department"),
                    create_readonly_item_div("Course Name", "course_name"),
                    create_readonly_item_div("Course Code", "course_code"),
                ),
                # Right Column (Data Entry Block - Interactive within form)
                Div(cls=f"{block_classes} ")(
                    create_editable_pill_div(
                        "Workflow Status", "workflow_status", WORKFLOW_STYLES
                    ),
                    create_editable_pill_div(
                        "Manual Classification",
                        "manual_classification",
                        manual_classification_options,
                    ),
                    Div(  # Remarks block layout
                        Strong(
                            "Remarks",
                            cls="block text-xs font-medium text-base-content/80 mb-1 underline decoration-pink-300",
                        ),
                        TextArea(
                            get_val("remarks", ""),
                            id="modal_remarks",
                            name="remarks",
                            rows="5",
                            cls="textarea textarea-bordered w-full text-sm bg-base-100",
                            data_original_value=get_val("remarks", ""),
                            oninput="markDirty()",
                        ),
                    ),
                    # Save/Reset Buttons are part of this form now
                    Div(cls="flex justify-end space-x-2 mt-4")(
                        Button(
                            "Reset",
                            type="button",
                            cls=ButtonT.secondary + " btn-sm",
                            onclick="resetModalForm(); return false;",
                        ),  # type=button prevents form submission
                        Button(  # Save Button
                            "Save",
                            id="modal-save-btn",
                            type="submit",
                            cls=ButtonT.primary + " btn-sm",
                            disabled=True,
                        ),
                        Span(  # Indicator Span
                            Span(
                                cls="animate-ping absolute inline-flex h-full w-full rounded-full bg-pink-500 opacity-75"
                            ),
                            Span(
                                cls="relative inline-flex rounded-full h-3 w-3 bg-pink-500"
                            ),
                            id="save-indicator",
                            cls="hidden flex relative h-3 w-3 top-0 right-4 -mt-2 -mr-1",
                        ),
                    ),
                ),
            ),  # End Main Content Grid
            # Item Details Grid (Readonly info outside form scope, but within modal body)
            Div(cls="grid grid-cols-1 md:grid-cols-5 gap-4 mt-4")(  # Added mt-4
                # Left Block (Text Details)
                Div(cls=f"{block_classes} md:col-span-4")(
                    create_readonly_item_div("Title", "title"),
                    create_readonly_item_div("Author", "author"),
                    create_readonly_item_div("Publisher", "publisher"),
                    create_readonly_item_div("DOI", "doi"),
                    create_readonly_item_div("ISBN", "isbn"),
                ),
                # Right Block (Numeric Details)
                Div(cls=f"{block_classes} md:col-span-1")(  # Apply block styles
                    Div(
                        Strong("Pages"),
                        Div(get_val("pagecount", "N/A"), cls="num-detail-val"),
                        cls="num-detail-item",
                    ),
                    Div(
                        Strong("Words"),
                        Div(get_val("wordcount", "N/A"), cls="num-detail-val"),
                        cls="num-detail-item",
                    ),
                    Div(
                        Strong("Pictures"),
                        Div(get_val("picturecount", "N/A"), cls="num-detail-val"),
                        cls="num-detail-item",
                    ),
                ),
            ),
            id="modal-details-form",
            hx_post=save_item_details.to(),
            hx_indicator="#modal-loading-indicator",
        )

        # --- 4. Footer Row (with HTMX for Next/Prev) ---

        indicator_attrs = {"hx_indicator": "#modal-loading-indicator"}

        prev_button_attrs = {
            "id": "modal-prev-btn",
            "cls": ButtonT.secondary + " btn-sm",
            "disabled": prev_id is None,
        }
        if prev_id is not None:
            prev_button_attrs.update(
                {
                    "hx_get": show_item_details.to(material_id=prev_id, **nav_params),
                    "hx_target": "#modal-placeholder",
                    "hx_swap": "innerHTML",
                    **indicator_attrs,
                }
            )
        prev_button = Button("< Prev", **prev_button_attrs)

        next_button_attrs = {
            "id": "modal-next-btn",
            "cls": ButtonT.secondary + " btn-sm",
            "disabled": next_id is None,
        }
        if next_id is not None:
            next_button_attrs.update(
                {
                    "hx_get": show_item_details.to(material_id=next_id, **nav_params),
                    "hx_target": "#modal-placeholder",
                    "hx_swap": "innerHTML",
                    **indicator_attrs,
                }
            )
        next_button = Button("Next >", **next_button_attrs)

        footer_content = Div(cls="modal-action mt-4 pt-4 border-t")(
            Div(cls="flex justify-between w-full")(
                prev_button,
                Form(method="dialog")(Button("Close", cls=ButtonT.primary + " btn-sm")),
                next_button,
            )
        )
        # --- Assemble Modal Box Content ---
        modal_box_content = Div(
            cls="modal-box w-[85vw] max-w-none h-[calc(100vh-5rem)] max-h-none flex flex-col"
        )(
            Div(header_content, cls="border-b pb-2 flex-shrink-0"),
            Div(cls="relative py-4 flex-grow overflow-y-auto ")(
                Div(
                    id="modal-loading-indicator",
                    cls="htmx-indicator absolute inset-0 bg-base-100/50 flex items-center justify-center z-50",
                )(Span("Loading...", cls="loading loading-lg")),
                modal_form_content,
            ),
            footer_content(cls="flex-shrink-0"),
        )

        # --- Backdrop ---
        modal_backdrop = Form(method="dialog", cls="modal-backdrop")(
            NotStr(
                '<button class="absolute inset-0 w-full h-full cursor-default outline-none" aria-label="close modal"></button>'
            )
        )

        response_content = (modal_box_content, modal_backdrop)
        return response_content, HtmxResponseHeaders(
            trigger="openModalEvent"
        )  # Trigger open on initial load/next/prev

    except Exception as e:
        print(f"Error generating modal content for ID {material_id}: {e}")
        print(traceback.format_exc())
        modal_box_content = Div(cls="modal-box")(
            H3("Error"),
            P(f"An error occurred: {e}"),
        )
        modal_backdrop = Form(method="dialog", cls="modal-backdrop")(
            NotStr(
                '<button class="absolute inset-0 w-full h-full cursor-default outline-none" aria-label="close"></button>'
            )
        )
        return (modal_box_content, modal_backdrop), HtmxResponseHeaders(
            trigger="openModalEvent"
        )


@rt("/data")
async def data_grid(
    request: Request,
    page: int = 1,
    per_page: int = DEFAULT_PER_PAGE,
    sort_by: str | None = None,
    sort_desc: str = "False",
    filter_department: str = "",
    filter_course_name: str = "",
    filter_manual_classification: str = "",
    filter_ml_classification: str = "",
    filter_status: str = "",
    filter_faculty: str = "",
    filter_author: str = "",
    filter_publisher: str = "",
    filter_workflow_status: str = "",
):
    """Endpoint renders the main page structure OR just the data grid component."""
    current_filters: dict[str, str] = {
        "department": filter_department.strip(),
        "course_name": filter_course_name.strip(),
        "manual_classification": filter_manual_classification.strip(),
        "ml_classification": filter_ml_classification.strip(),
        "status": filter_status.strip(),
        "faculty": filter_faculty.strip(),
        "author": filter_author.strip(),
        "publisher": filter_publisher.strip(),
        "workflow_status": filter_workflow_status.strip(),
    }
    current_filters = {k: v for k, v in current_filters.items() if v}
    print(f"filters in data_grid route: {current_filters}")

    sort_desc_bool = sort_desc.lower() == "true"

    filtered_sorted_df = get_filtered_sorted_df(
        sort_by, sort_desc_bool, current_filters
    )
    total_filtered_rows = filtered_sorted_df.height
    total_pages = math.ceil(total_filtered_rows / per_page) if per_page > 0 else 1
    page = max(1, min(page, total_pages if total_pages > 0 else 1))
    offset = (page - 1) * per_page
    df_slice = filtered_sorted_df.slice(offset, per_page)
    print(f"filters in data_grid route: {current_filters}")
    grid_component = render_data_grid_component(
        df_slice=df_slice,
        current_page=page,
        total_pages=total_pages,
        per_page=per_page,
        total_filtered_rows=total_filtered_rows,
        offset=offset,
        sort_by=sort_by,
        sort_desc=sort_desc_bool,
        current_filters=current_filters,
    )

    header_component = page_header_component(
        current_filters=current_filters,
        per_page=per_page,
        sort_by=sort_by,
        sort_desc=sort_desc_bool,
    )

    is_htmx = request.headers.get("hx-request", "false").lower() == "true"

    if not is_htmx:
        modal_placeholder = Dialog(
            id="modal-placeholder", cls="modal modal-bottom sm:modal-middle"
        )
        modal_trigger_script = Script("""
            document.body.addEventListener('openModalEvent', function(evt) {
                const modalDialog = document.getElementById('modal-placeholder');
                if (modalDialog && typeof modalDialog.showModal === 'function') {
                    console.log('Opening modal via openModalEvent (Target: #modal-placeholder)');
                    modalDialog.showModal();
                } else {
                    console.error('Modal dialog (#modal-placeholder) not found or showModal not supported.');
                }
            });
        """)

        modal_interaction_script = Script(f"""
            // Function to update a pill's appearance and hidden input value
            function updatePill(fieldName, newValue, newText, newStyleClass) {{
                const pillElement = document.getElementById(`pill-display-${{fieldName}}`);
                const inputElement = document.getElementById(`input-${{fieldName}}`);
                if (pillElement && inputElement) {{
                    // Update hidden input
                    inputElement.value = newValue;
                    // Update visible pill text
                    pillElement.textContent = newText;
                    // Update visible pill style (remove old, add new)
                    // Assumes style classes are like 'uk-label-primary', 'uk-label-secondary', etc.
                    pillElement.classList.remove('uk-label-primary', 'uk-label-secondary', 'uk-label-destructive');
                    if (newStyleClass) {{ // Add new style if provided
                    pillElement.classList.add(newStyleClass);
                    }} else {{ // Fallback if no specific style maps (e.g., for 'N/A')
                        pillElement.classList.add('uk-label-secondary'); // Or your default
                    }}
                    // Close the dropdown (assuming uk-drop is used)
                    const drop = UIkit.drop(pillElement.closest('[uk-drop]'));
                    if (drop) {{ drop.hide(false); }}
                    markDirty(); // Mark form as dirty
                }} else {{
                    console.error(`Cannot find pill or input elements for ${{fieldName}}`);
                }}
            }}
            function markDirty() {{
                const indicator = document.getElementById('save-indicator');
                const saveButton = document.getElementById('modal-save-btn');
                if (indicator) indicator.classList.remove('hidden');
                if (saveButton) saveButton.disabled = false; // Make sure it's enabled
            }}

            // Function to reset the modal form fields to original values
            function resetModalForm() {{
                console.log('Resetting modal form'); // Debug log
                const form = document.getElementById('modal-details-form');
                if (!form) return;

                const pillInputs = form.querySelectorAll('input[data-original-value][id^="input-"]'); // Target only pill inputs
                pillInputs.forEach(input => {{
                    const originalValue = input.dataset.originalValue;
                    const originalText = input.dataset.originalText || originalValue; // Fallback text
                    const originalStyle = input.dataset.originalStyle || '{str(DEFAULT_PILL_STYLE)}'; // Fallback style
                    const fieldName = input.id.replace('input-', '');

                    // Reset hidden input
                    input.value = originalValue;

                    // Reset visible pill
                    const pillElement = document.getElementById(`pill-display-${{fieldName}}`);
                    if (pillElement) {{
                        pillElement.textContent = originalText;
                        pillElement.classList.remove('uk-label-primary', 'uk-label-secondary', 'uk-label-destructive');
                        if (originalStyle && originalStyle.startsWith('uk-label-')) {{ // Ensure it's a valid style class
                            pillElement.classList.add(originalStyle);
                        }} else {{
                             pillElement.classList.add('{str(DEFAULT_PILL_STYLE)}'); // Default
                        }}
                    }}
                }});

                const remarksTextarea = form.querySelector('#modal_remarks');
                if (remarksTextarea && typeof remarksTextarea.dataset.originalValue !== 'undefined') {{ // Check if attribute exists
                    remarksTextarea.value = remarksTextarea.dataset.originalValue;
                }}

                // Hide save indicator and disable save button
                const indicator = document.getElementById('save-indicator');
                const saveButton = document.getElementById('modal-save-btn');
                if (indicator) {{
                    indicator.classList.add('hidden');
                }}
                if (saveButton) {{
                    saveButton.disabled = true;
                }}
            }}
        """)

        htmx_toast_template = Template(
            Div(
                Span(slot="message"),
                Button(
                    "✕",
                    type="button",
                    cls="btn btn-sm btn-outline",
                    aria_label="Close",
                    slot="close",
                ),
                cls="alert",
                slot="alert",
            ),
            id="htmx-toasts-template",
        )

        htmx_toast_settings = Htmx_toasts(
            timeout="30000",
            cls="toast",
            role="status",
            aria_live="polite",
            error_class="alert-error",
            info_class="alert-info",
            warn_class="alert-warning",
            success_class="alert-success",
        )

        htmx_toast_script = Script("""
        document.addEventListener("save_success", function(evt) {
                window.dispatchEvent(new CustomEvent('htmx-toasts:notify', {
                    detail: {
                        message: 'Successfully saved data for material ID: ' + evt.detail.material_id,
                        level: 'success'
                    }
                }));

                // 2. Reset modal dirty state
                // Hide save indicator
                const indicator = document.getElementById('save-indicator');
                if (indicator) indicator.classList.add('hidden');

                const saveButton = document.getElementById('modal-save-btn');
                if (saveButton) saveButton.disabled = true;

                const form = document.getElementById('modal-details-form');
                if (form) {
                    const fieldsToUpdate = ['input-workflow_status', 'input-manual_classification', 'modal_remarks'];
                    fieldsToUpdate.forEach(id => {
                        const element = form.querySelector('#' + id);
                        if (element) {
                            element.dataset.originalValue = element.value;
                            if(id.startsWith('input-')) {
                                const fieldName = id.replace('input-', '');
                                const pillElement = document.getElementById(`pill-display-${fieldName}`);
                                if(pillElement) {
                                    element.dataset.originalText = pillElement.textContent;
                                    const styleClass = Array.from(pillElement.classList).find(cls => cls.startsWith('uk-label-')) || '';
                                    element.dataset.originalStyle = styleClass;
                                }
                            }
                        }
                    });
                }
            });


        document.addEventListener("save_error", function(evt) {
                window.dispatchEvent(new CustomEvent('htmx-toasts:notify', {
                    detail: {
                        message: 'Error saving data for material ID: ' + evt.detail.material_id,
                        level: 'error'
                    }
                }));
            });
        """)

        # Update the return tuple
        return (
            Title("Copyright Data Dashboard"),
            Div(  # page-container
                Div(header_component, grid_component, id="content-area"),
                id="page-container",
            ),
            modal_placeholder,
            modal_trigger_script,
            modal_interaction_script,
            htmx_toast_template,
            htmx_toast_settings,
            htmx_toast_script,
        )

    else:
        return grid_component


@rt("/save_details", methods=["POST"])
async def save_item_details(
    request: Request,
    # Editable fields
    material_id: int,
    workflow_status: str = "",
    manual_classification: str = "",
    remarks: str = "",
    # State fields from hidden inputs
    page: int = 1,
    per_page: int = DEFAULT_PER_PAGE,
    sort_by: Optional[str] = None,
    sort_desc: str = "False",
    filter_department: str = "",
    filter_course_name: str = "",
    filter_manual_classification: str = "",  # Note: might conflict with editable field?
    filter_ml_prediction: str = "",
    filter_status: str = "",
    filter_faculty: str = "",
    filter_author: str = "",
    filter_publisher: str = "",
    filter_workflow_status: str = "",  # Note: might conflict with editable field?
):
    """Handles saving changes, updates global df, returns OOB grid refresh + JS trigger."""
    global copyright_df_global

    print(f"Saving changes for material_id: {material_id}")
    update_data = {
        "material_id": material_id,
        "workflow_status": workflow_status,
        "manual_classification": manual_classification,
        "remarks": remarks,
    }

    oob_grid_swap = Div()  # Default empty div
    trigger_name = ""

    try:
        await store_item_changes(update_data)

        print("Reloading global DataFrame after save...")
        copyright_df_global = retrieve_copyright_items()
        print("Global DataFrame reloaded.")

        current_grid_filters: dict[str, str] = {
            "department": filter_department.strip(),
            "course_name": filter_course_name.strip(),
            "manual_classification": filter_manual_classification.strip(),
            "ml_prediction": filter_ml_prediction.strip(),
            "status": filter_status.strip(),
            "faculty": filter_faculty.strip(),
            "author": filter_author.strip(),
            "publisher": filter_publisher.strip(),
            "workflow_status": filter_workflow_status.strip(),
        }
        current_grid_filters = {k: v for k, v in current_grid_filters.items() if v}
        grid_sort_desc_bool = sort_desc.lower() == "true"

        # Re-filter and slice for the current page view
        filtered_sorted_df = get_filtered_sorted_df(
            sort_by, grid_sort_desc_bool, current_grid_filters
        )
        total_filtered_rows = filtered_sorted_df.height
        total_pages = math.ceil(total_filtered_rows / per_page) if per_page > 0 else 1
        page = max(1, min(page, total_pages if total_pages > 0 else 1))
        offset = (page - 1) * per_page
        df_slice = filtered_sorted_df.slice(offset, per_page)

        # Render the grid component HTML
        updated_grid_component = render_data_grid_component(
            df_slice=df_slice,
            current_page=page,
            total_pages=total_pages,
            per_page=per_page,
            total_filtered_rows=total_filtered_rows,
            offset=offset,
            sort_by=sort_by,
            sort_desc=grid_sort_desc_bool,
            current_filters=current_grid_filters,
        )

        # --- Create OOB Swap Div for the Grid ---
        oob_grid_swap = Div(
            to_xml(updated_grid_component),
            hx_swap_oob="outerHTML:#data-grid-component",
        )

        trigger_name = {"save_success": {"material_id": material_id}}

    except Exception as e:
        print(f"Error saving changes for {material_id}: {e}")
        trigger_name: dict[str, dict[str, int]] = {
            "save_error": {"material_id": material_id}
        }

    return oob_grid_swap, HtmxResponseHeaders(
        reswap="none",
        trigger=json.dumps(trigger_name),  # Directly trigger the JS function name
    )


@rt("/")
async def root_redirect():
    """Redirect root to data grid."""
    return RedirectResponse(url=data_grid.to(), status_code=302)


if __name__ == "__main__":
    print("Starting FastHTML server...")
    serve(port=8000, reload=True)
