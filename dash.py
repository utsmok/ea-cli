# dash.py

import contextlib
import math
import traceback
from enum import auto
from typing import Any

import fasthtml.common as fh
import polars as pl
from fastcore.xml import FT
from fasthtml.common import *
from monsterui.all import *
from monsterui.foundations import VEnum, str2ukcls
from rich import print

from easy_access.db.retrieve import retrieve_copyright_items

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
    """Applies filtering and sorting. (Logic unchanged)"""
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


# --- Component rendering ---
def render_table_rows(
    df_slice: pl.DataFrame,
    current_page_for_modal: int,
    per_page_for_modal: int,
    sort_by_for_modal: Optional[str],
    sort_desc_for_modal: bool,
    filters_for_modal: Optional[Dict[str, str]],
) -> tuple[FT, ...]:
    """Renders Tbody rows with custom formatting and HTMX attributes."""
    rows: list[FT] = []
    if df_slice is not None and df_slice.height > 0:
        cols_to_display = [col for col in DISPLAY_COLUMNS if col in df_slice.columns]
        display_df_slice = df_slice.select(cols_to_display)

        # --- Define Pill Styles ---
        workflow_styles = {
            "ToDo": LabelT.destructive,
            "InProgress": LabelT.secondary,
            "Done": LabelT.primary,
        }
        status_styles = {
            "Published": LabelT.primary,
            "Unpublished": LabelT.secondary,
            "Deleted": LabelT.destructive,
        }
        # --- Define Classification Styles ---
        primary_classifications = {
            "open access",
            "eigen materiaal - powerpoint",
            "eigen materiaal - titelindicatie",
            "eigen materiaal - overig",
            "eigen materiaal",
        }
        secondary_classifications = {
            "onbekend",
            "niet geanalyseerd",
            "in onderzoek",
            "licentie beschikbaar",
            "verwijderverzoek verstuurd",
        }
        destructive_classifications = {
            "korte overname",
            "middellange overname",
            "lange overname",
        }
        # --- End Define Styles ---

        for row_dict in display_df_slice.iter_rows(named=True):
            cells = []
            material_id = row_dict.get("material_id", None)
            if material_id is None:
                continue

            for col in cols_to_display:
                val = row_dict.get(col)
                # Normalize value for comparisons where applicable
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
                    # Use original value for display, normalized for lookup
                    style = workflow_styles.get(str(val))
                    cell_content = Label(
                        str(val) if val else "N/A", cls=style + " badge-sm"
                    )
                elif col == "status":
                    # Use original value for display, normalized for lookup
                    style = status_styles.get(str(val))
                    cell_content = Label(
                        str(val) if val else "N/A", cls=style + " badge-sm"
                    )

                elif col in ("manual_classification", "ml_prediction"):
                    if not val_str_norm:  # Handle empty/None case
                        display_text = "-"
                        style = ""
                    elif val_str_norm in primary_classifications:
                        style = LabelT.primary
                        display_text = str(val)  # Use original casing for display
                    elif val_str_norm in secondary_classifications:
                        style = LabelT.secondary
                        display_text = str(val)
                    elif val_str_norm in destructive_classifications:
                        style = LabelT.destructive
                        display_text = str(val)
                    else:  # Fallback for any unexpected values
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
            row_attrs: dict[str, Any] = {
                "id": f"row-{material_id}",
                "class": "hover:bg-primary/10 cursor-pointer transition-colors duration-150",
                "hx_get": show_item_details.to(material_id=material_id),
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
    current_filters: Dict[str, str],
    per_page: int,
    sort_by: Optional[str],
    sort_desc: bool,
) -> FT:
    """Renders the header area with Title and Filters. (Unchanged)"""
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
    sort_by: Optional[str] = None,
    sort_desc: bool = False,
    current_filters: Optional[Dict[str, str]] = None,
) -> FT:
    """Renders the main content area: Item count, table, pagination."""
    current_filters = current_filters or {}

    # --- Table Header ---
    header_cells = []
    cols_in_header = [
        col for col in DISPLAY_COLUMNS if col in copyright_df_global.columns
    ]
    for col in cols_in_header:
        # --- RENAME Header ---
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
            header_text,  # Use renamed text
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
        # --- ADD CSS Class to TH ---
        th_class = f"col-{col.replace('_', '-')}"  # Generate class name
        header_cells.append(
            Th(header_content, cls=f"px-4 py-2 {th_class}")
        )  # Add class

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
    sort_by: Optional[str] = None,
    sort_desc: bool = False,
    current_filters: Optional[Dict[str, str]] = None,
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

    # --- Pills & Icons ---
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
    # Add params matching data_grid
    page: int = 1,
    per_page: int = DEFAULT_PER_PAGE,
    sort_by: Optional[str] = None,
    sort_desc: str = "False",
    filter_department: str = "",
    filter_course_name: str = "",
    filter_manual_classification: str = "",
    filter_ml_prediction: str = "",  # Assuming ml_prediction filter exists
    filter_status: str = "",
    filter_faculty: str = "",  # Assuming faculty filter exists
    filter_author: str = "",
    filter_publisher: str = "",
    filter_workflow_status: str = "",
):
    """Fetches data and returns structured INNER content for the modal dialog
    using direct Tailwind classes."""
    try:
        # --- Determine Current List Order ---
        # Reconstruct filters from incoming parameters
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
        # --- End Re-filtering ---

        item_df = ordered_df.filter(pl.col("material_id") == material_id)
        if item_df.height == 0:  # ... (error handling) ...
            # ...
            return (modal_box_content, modal_backdrop), HtmxResponseHeaders(
                trigger="openModalEvent"
            )

        item_data = item_df.to_dicts()[0]

        def get_val(key, default=None):
            return item_data.get(key, default)

        # --- Helper Function for Dropdown Pills ---
        def create_editable_pill_div(
            label_text: str, field_name: str, options_map: dict
        ):
            current_value = get_val(field_name)
            content_component, _ = render_modal_field(
                field_name, current_value
            )  # Get current pill
            label_el, _ = render_labelled_item(
                label_text, content_component
            )  # Get label

            # Store original value and style for reset
            original_value_str = str(current_value) if current_value is not None else ""
            original_style_class = getattr(
                content_component, "class", [DEFAULT_PILL_STYLE + " badge-sm"]
            )[0].split()[0]  # Heuristic to get base style class

            # Dropdown Items
            dropdown_items = []
            for opt_val, opt_style_enum in options_map.items():
                # Get the actual style class string from the enum
                opt_style_class = str(opt_style_enum)
                # Use repr to safely quote strings in JS call
                onclick_js = f"updatePill('{field_name}', {repr(opt_val)}, {repr(opt_val)}, '{opt_style_class}'); return false;"
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
                name=field_name,  # Important for potential future form submission
                value=original_value_str,
                data_original_value=original_value_str,  # Store original value
                data_original_text=original_value_str,  # Store original text
                data_original_style=original_style_class,  # Store original style
            )

            # Wrap the pill in a Button for dropdown trigger, add ID to the pill itself
            pill_wrapper = Div(
                cls="inline-block"
            )(  # Dropdown container
                Button(
                    # Add ID to the visual pill element itself for easy JS targeting
                    content_component(id=f"pill-display-{field_name}"),
                    type="button",
                    cls="p-0 m-0 bg-transparent border-none hover:opacity-80",  # Make button invisible
                ),
                DropDownNavContainer(
                    *dropdown_items, cls="uk-dropdown-nav", uk_drop="mode: click"
                ),
                hidden_input,  # Include hidden input
            )

            # Apply inline layout classes
            return Div(
                label_el,
                Div(pill_wrapper, cls="text-right"),
                cls="flex items-center justify-between space-x-2 mb-2",  # label-inline-item styling
            )

        # --- End Dropdown Helper ---

        # --- Create Combined Options for Manual Classification ---
        manual_classification_options = {}
        for val in PRIMARY_CLASSIFICATIONS:
            manual_classification_options[val] = LabelT.primary
        for val in SECONDARY_CLASSIFICATIONS:
            manual_classification_options[val] = LabelT.secondary
        for val in DESTRUCTIVE_CLASSIFICATIONS:
            manual_classification_options[val] = LabelT.destructive
        # Add 'onbekend' if not already present
        if "onbekend" not in manual_classification_options:
            manual_classification_options["onbekend"] = LabelT.secondary

        # --- 1. Header Row ---
        # ... (header_content generation remains the same) ...
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

        # --- 2. Main Content Row (2 Columns) ---
        block_classes = "border border-[hsl(var(--ring))] rounded-md p-4 shadow-sm"

        # Using the simplified create_labelled_item_div for non-editable fields
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
                "flex items-center justify-between space-x-2 mb-2"
                if is_inline
                else "mb-3"
            )
            content_wrapper_cls = "text-right" if is_inline else ""
            return Div(
                label_el, Div(content_el, cls=content_wrapper_cls), cls=container_cls
            )

        item_info_block = Div(cls=f"{block_classes} space-y-1")(
            create_readonly_item_div(
                "Classification", "classification", is_inline=True
            ),
            create_readonly_item_div(
                "ML Prediction", "ml_prediction", is_inline=True
            ),  # Keep ML readonly for now
            create_readonly_item_div("Period", "period"),
            create_readonly_item_div("Faculty", "faculty"),
            create_readonly_item_div("Owner", "owner"),
            create_readonly_item_div("Department", "department"),
            create_readonly_item_div("Course Name", "course_name"),
            create_readonly_item_div("Course Code", "course_code"),
        )

        data_entry_block = Div(
            cls=f"{block_classes} space-y-1"
        )(
            create_editable_pill_div(
                "Workflow Status", "workflow_status", WORKFLOW_STYLES
            ),  # Editable
            create_editable_pill_div(
                "Manual Classification",
                "manual_classification",
                manual_classification_options,
            ),  # Editable
            # Remarks (Editable)
            Div(
                Strong(
                    "Remarks", cls="block text-xs font-medium text-base-content/80 mb-1"
                ),
                TextArea(
                    get_val("remarks", ""),
                    id="modal_remarks",
                    name="remarks",
                    rows="5",
                    cls="textarea textarea-bordered w-full text-sm bg-base-100",  # Use base-100 for editable
                    # Store original value and add JS trigger
                    data_original_value=get_val("remarks", ""),
                    oninput="markDirty()",
                ),
                cls="mb-3",
            ),
            Div(cls="flex justify-end space-x-2 mt-4")(
                # Add onclick for Reset
                Button(
                    "Reset",
                    cls=ButtonT.secondary + " btn-sm",
                    onclick="resetModalForm(); return false;",
                ),
                # Add Save indicator and onclick alert, disable initially
                Button(
                    # Save indicator structure (initially hidden)
                    Span(  # Relative container for positioning the ping
                        Span(
                            cls="relative flex size-3 mr-2"
                        )(  # Use relative positioning
                            Span(
                                cls="absolute inline-flex h-full w-full animate-ping rounded-full bg-red-400 opacity-75"
                            ),
                            UkIcon(
                                "alert-triangle",
                                cls="relative inline-flex size-3 text-red-500",
                            ),  # Icon instead of solid circle
                        ),
                        id="save-indicator",
                        cls="hidden",  # Hide initially
                    ),
                    "Save",  # Button text
                    id="modal-save-btn",  # Button ID
                    cls=ButtonT.primary + " btn-sm",
                    onclick="alert('Save functionality is not yet implemented.'); return false;",
                    disabled=True,  # Disabled initially
                ),
            ),
        )

        main_content_grid = Div(
            item_info_block,
            data_entry_block,
            cls="grid grid-cols-1 md:grid-cols-2 gap-4",
        )

        # --- 3. Item Details Row (Split Blocks) ---
        text_details_block = Div(
            cls=f"{block_classes} space-y-1"
        )(  # Apply block styles
            create_readonly_item_div("Title", "title"),
            create_readonly_item_div("Author", "author"),
            create_readonly_item_div("Publisher", "publisher"),
            create_readonly_item_div("DOI", "doi"),
            create_readonly_item_div("ISBN", "isbn"),
        )
        numeric_details_block = Div(
            cls=f"{block_classes} space-y-0.5"
        )(  # Apply block styles
            Div(
                Strong("Pages"),
                Div(
                    get_val("pagecount", "N/A"),
                    cls="p-1 px-1.5 border rounded bg-base-200 text-sm text-red-600 font-mono text-right min-w-[40px]",
                ),
                cls="flex justify-between items-baseline text-sm mb-1",
            ),
            Div(
                Strong("Words"),
                Div(
                    get_val("wordcount", "N/A"),
                    cls="p-1 px-1.5 border rounded bg-base-200 text-sm text-red-600 font-mono text-right min-w-[40px]",
                ),
                cls="flex justify-between items-baseline text-sm mb-1",
            ),
            Div(
                Strong("Pictures"),
                Div(
                    get_val("picturecount", "N/A"),
                    cls="p-1 px-1.5 border rounded bg-base-200 text-sm text-red-600 font-mono text-right min-w-[40px]",
                ),
                cls="flex justify-between items-baseline text-sm mb-1",
            ),
        )

        # Use 5 columns for 80/20 split, apply gap
        item_details_grid = Div(
            text_details_block(cls="md:col-span-4"),  # Takes 4 columns
            numeric_details_block(cls="md:col-span-1"),  # Takes 1 column
            cls="grid grid-cols-1 md:grid-cols-5 gap-4",
        )

        # --- 4. Footer Row (with HTMX for Next/Prev) ---
        nav_params = {  # ... (nav_params setup remains the same) ...
            "page": page,
            "per_page": per_page,
            "sort_by": sort_by or "",
            "sort_desc": sort_desc,
            **{f"filter_{k}": v for k, v in current_modal_filters.items()},
        }

        # Add hx-indicator to buttons
        indicator_attrs = {"hx_indicator": "#modal-loading-indicator"}  # Define once

        prev_button_attrs = {  # ... (existing attrs) ...
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

        next_button_attrs = {  # ... (existing attrs) ...
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
            # Add ID for form reset targeting
            id="modal-details-form",
            cls="modal-box w-[85vw] max-w-none h-[calc(100vh-5rem)] max-h-none flex flex-col",
        )(
            Div(header_content, cls="border-b pb-2 flex-shrink-0"),
            # Add loading indicator inside scroll area but outside specific blocks
            Div(
                # Loading indicator element (initially hidden by htmx-indicator rules)
                Div(
                    id="modal-loading-indicator",
                    cls="htmx-indicator absolute inset-0 bg-base-100/50 flex items-center justify-center z-50",
                )(
                    Span(
                        "Loading...", cls="loading loading-lg"
                    )  # Or use monsterui Loading
                ),
                main_content_grid,
                item_details_grid,
                cls="relative py-4 flex-grow overflow-y-auto space-y-4",  # Add relative for indicator positioning
            ),
            footer_content(cls="flex-shrink-0"),  # Ensure footer doesn't grow
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
        )  # Still trigger open

    except Exception as e:
        # ... (Error handling remains similar) ...
        print(f"Error generating modal content for ID {material_id}: {e}")
        print(traceback.format_exc())
        modal_box_content = Div(cls="modal-box")(
            H3("Error"),
            P(f"An error occurred: {e}"),  # ...
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
    sort_by: Optional[str] = None,
    sort_desc: str = "False",
    # Explicit filter args (unchanged)
    filter_department: str = "",
    filter_course_name: str = "",
    filter_manual_classification: str = "",
    filter_ml_classification: str = "",
    filter_status: str = "",
    filter_faculty_id: str = "",
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
        "faculty_id": filter_faculty_id.strip(),
        "author": filter_author.strip(),
        "publisher": filter_publisher.strip(),
        "workflow_status": filter_workflow_status.strip(),
    }
    current_filters = {k: v for k, v in current_filters.items() if v}
    sort_desc_bool = sort_desc.lower() == "true"

    filtered_sorted_df = get_filtered_sorted_df(
        sort_by, sort_desc_bool, current_filters
    )
    total_filtered_rows = filtered_sorted_df.height
    total_pages = math.ceil(total_filtered_rows / per_page) if per_page > 0 else 1
    page = max(1, min(page, total_pages if total_pages > 0 else 1))
    offset = (page - 1) * per_page
    df_slice = filtered_sorted_df.slice(offset, per_page)

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
        modal_interaction_script = Script("""
            // Function to update a pill's appearance and hidden input value
            function updatePill(fieldName, newValue, newText, newStyleClass) {
                const pillElement = document.getElementById(`pill-display-${fieldName}`);
                const inputElement = document.getElementById(`input-${fieldName}`);
                if (pillElement && inputElement) {
                    // Update hidden input
                    inputElement.value = newValue;
                    // Update visible pill text
                    pillElement.textContent = newText;
                    // Update visible pill style (remove old, add new)
                    // Assumes style classes are like 'uk-label-primary', 'uk-label-secondary', etc.
                    pillElement.classList.remove('uk-label-primary', 'uk-label-secondary', 'uk-label-destructive');
                    if (newStyleClass) { // Add new style if provided
                       pillElement.classList.add(newStyleClass);
                    } else { // Fallback if no specific style maps (e.g., for 'N/A')
                        pillElement.classList.add('uk-label-secondary'); // Or your default
                    }
                    // Close the dropdown (assuming uk-drop is used)
                    const drop = UIkit.drop(pillElement.closest('[uk-drop]'));
                    if (drop) { drop.hide(false); }

                    markDirty(); // Mark form as dirty
                } else {
                    console.error(`Cannot find pill or input elements for ${fieldName}`);
                }
            }

            // Function to mark the form as dirty (show save indicator)
            function markDirty() {
                const indicator = document.getElementById('save-indicator');
                const saveButton = document.getElementById('modal-save-btn');
                if (indicator) {
                    indicator.classList.remove('hidden');
                }
                 if (saveButton) {
                    saveButton.disabled = false; // Enable save button
                }
            }

            // Function to reset the modal form fields to original values
            function resetModalForm() {
                const form = document.getElementById('modal-details-form'); // Need to add this ID to the modal-box div or a form wrapper
                if (!form) return;

                // Reset Pills (Workflow Status, Manual Classification)
                const pillInputs = form.querySelectorAll('input[data-original-value]');
                pillInputs.forEach(input => {
                    const originalValue = input.dataset.originalValue;
                    const originalText = input.dataset.originalText; // Need to store this
                    const originalStyle = input.dataset.originalStyle; // Need to store this
                    const fieldName = input.id.replace('input-', ''); // Extract field name

                    // Reset hidden input
                    input.value = originalValue;

                    // Reset visible pill
                    const pillElement = document.getElementById(`pill-display-${fieldName}`);
                    if (pillElement) {
                        pillElement.textContent = originalText;
                        pillElement.classList.remove('uk-label-primary', 'uk-label-secondary', 'uk-label-destructive');
                         if (originalStyle) {
                             pillElement.classList.add(originalStyle);
                         } else {
                             pillElement.classList.add('uk-label-secondary'); // Default
                         }
                    }
                });

                // Reset Remarks Textarea
                const remarksTextarea = form.querySelector('#modal_remarks');
                if (remarksTextarea && remarksTextarea.dataset.originalValue) {
                    remarksTextarea.value = remarksTextarea.dataset.originalValue;
                }

                // Hide save indicator and disable save button
                const indicator = document.getElementById('save-indicator');
                const saveButton = document.getElementById('modal-save-btn');
                 if (indicator) {
                    indicator.classList.add('hidden');
                }
                 if (saveButton) {
                    saveButton.disabled = true;
                }
            }
        """)
        modal_trigger_script = Script("""
            document.body.addEventListener('openModalEvent', function(evt) {
                // Target the persistent dialog element directly by its ID
                const modalDialog = document.getElementById('modal-placeholder');
                if (modalDialog && typeof modalDialog.showModal === 'function') {
                    console.log('Opening modal via openModalEvent (Target: #modal-placeholder)');
                    modalDialog.showModal();
                } else {
                    console.error('Modal dialog (#modal-placeholder) not found or showModal not supported.');
                }
            });
        """)
        # Update the return tuple for the full page load
        return (
            Title("Copyright Data Dashboard"),
            Div(  # page-container
                Div(header_component, grid_component, id="content-area"),
                id="page-container",
            ),
            modal_placeholder,
            modal_trigger_script,
            modal_interaction_script,  # Add the new script
        )

    else:
        # HTMX swap: Replace just the grid component
        return grid_component


@rt("/")
async def root_redirect():
    """Redirect root to data grid."""
    return RedirectResponse(url=data_grid.to(), status_code=302)


if __name__ == "__main__":
    print("Starting FastHTML server...")
    serve(port=8000, reload=True)
