# dash.py

import contextlib
import math
from datetime import datetime
from typing import Any

import polars as pl
from fastcore.xml import FT
from fasthtml.common import *
from monsterui.all import *

from easy_access.db.retrieve import retrieve_copyright_items

GLOBAL_STYLES = Style("""
    * {
        font-family: "Inter", sans-serif;
        font-optical-sizing: auto;
        font-weight: 400;
        font-style: normal;
    }
    html, body { height: 100%; overflow: hidden; margin: 0; padding: 0; background-color: hsl(var(--b2)); /* Use theme base-200 for body */ }
    #page-container { display: flex; flex-direction: column; height: 100vh; background-color: hsl(var(--b1)); /* base-100 for main page area */ }
    #content-area { flex-grow: 1; overflow-y: auto; overflow-x: hidden; padding: 1rem 1.5rem; /* Add padding to scroll area */ }
    /* Table wrapper handles horizontal scroll */
    #table-wrapper { overflow-x: auto; overflow-y: hidden; }
    /* Sticky table header within the scrollable content area */
    #data-table thead th { position: sticky; top: 0; z-index: 10; background-color: hsl(var(--b2)/0.95); backdrop-filter: blur(4px); /* semi-transparent sticky header */ }
""")

# Using Orange theme, adding a global style for body overflow
MONSTER_HEADERS: tuple[FT, ...] = (
    *Theme.orange.headers(  # Unpack the original headers
        mode="light",
        daisy=True,
        radii=ThemeRadii.lg,
        shadows=ThemeShadows.lg,
    ),
    # Add global style to prevent body scrolling and ensure full height
    GLOBAL_STYLES,
)


copyright_df_global: pl.DataFrame = retrieve_copyright_items()

app, rt = fast_app(
    hdrs=(
        MONSTER_HEADERS,
        Link(
            rel="stylesheet",
            href="https://fonts.googleapis.com/css2?family=Inter:ital,opsz,wght@0,14..32,100..900;1,14..32,100..900&display=swap",
            type="text/css",
        ),
    ),
    exts="loading-states",
)

# --- Configuration ---
FILTERABLE_COLUMNS: list[str] = [
    "department",
    "course_name",
    "manual_classification",
    "ml_classification",
    "status",
    "faculty_id",
    "author",
    "publisher",
    "workflow_status",
]
DISPLAY_COLUMNS: list[str] = [
    "material_id",
    "url",
    "workflow_status",
    "manual_classification",
    "status",
    "remarks",
    "title",
    "filename",
    "faculty_id",
    "period",
    "department",
    "author",
    "publisher",
]
DEFAULT_PER_PAGE = 15
MAX_CELL_LENGTH = 35

# --- Helper Functions ---


def get_filtered_sorted_df(
    sort_by: Optional[str] = None,
    sort_desc: bool = False,
    filters: Optional[Dict[str, str]] = None,
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
                            print(f"filter_expressions: {filter_expressions}")
                            print(f"actual_col: {actual_col}")
                            print(f"col: {col}")
                            print(f"value: {value.capitalize()}")
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


def render_table_rows(df_slice: pl.DataFrame) -> tuple[FT, ...]:
    """Renders Tbody rows with consistent hover."""
    rows: list[FT] = []
    if df_slice is not None and df_slice.height > 0:
        cols_to_display = [col for col in DISPLAY_COLUMNS if col in df_slice.columns]
        display_df_slice = df_slice.select(cols_to_display)
        for row_dict in display_df_slice.iter_rows(named=True):
            cells = []
            material_id = row_dict.get("material_id", None)
            for col in cols_to_display:
                val = row_dict.get(col)
                cell_content: Any = str(val) if val is not None else ""
                if col == "url" and val:
                    display_url = (
                        val[: MAX_CELL_LENGTH - 3] + "..."
                        if len(val) > MAX_CELL_LENGTH
                        else val
                    )
                    cell_content = A(
                        display_url,
                        href=val,
                        target="_blank",
                        rel="noopener noreferrer",
                        cls="link link-hover text-primary hover:text-primary-focus text-sm underline",
                    )
                elif isinstance(val, str) and len(val) > MAX_CELL_LENGTH:
                    cell_content = val[: MAX_CELL_LENGTH - 3] + "..."
                cells.append(
                    Td(cell_content, cls="py-2.5 px-4 border-b border-base-200/80")
                )  # Lighter border
            row_attrs: dict[str, Any] = (
                {"id": f"row-{material_id}"} if material_id else {}
            )
            row_attrs["class"] = (
                "hover:bg-primary/10 cursor-pointer transition-colors duration-150"
            )
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
        ),  # Adjusted grid/gap
        Input(type="hidden", name="page", value="1"),
        Input(type="hidden", name="per_page", value=str(per_page)),
        Input(type="hidden", name="sort_by", value=sort_by or ""),
        Input(type="hidden", name="sort_desc", value=str(sort_desc)),
        id="filter-form",
        cls="",
    )  # No extra classes needed on form itself

    # Layout using flex: Title on left, Filters take remaining space
    return Div(
        # Title Area
        Div(
            H4(
                "Copyright Data Dashboard",
                cls="text-4xl font-bold underline decoration-pink-500 text-primary",
            ),
            cls="flex-shrink-0 mr-8 pt-1",
        ),  # Added padding-top
        # Filter Area
        Div(
            filter_form, cls="flex-grow justify-center"
        ),  # Filters take up available space
        cls="flex items-start mb-4 border-b border-base-300 pb-3",  # Flex layout for header elements
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
            col.replace("_", " ").title(),
            UkIcon(
                sort_indicator_icon, cls="inline-block ml-1.5 w-3 h-3 stroke-current"
            )
            if sort_indicator_icon
            else "",
            hx_get=data_grid.to(**sort_link_params),
            hx_target="#data-grid-component",  # Target self
            hx_indicator="#grid-loading-indicator",
            # Subtle hover, no background change that obscures text
            cls=f"{ButtonT.ghost} text-xs uppercase tracking-wider p-1.5 h-auto min-h-0 font-bold text-base-content/70 hover:text-primary transition-colors duration-150",
        )
        header_cells.append(Th(header_content, cls="px-4 py-2"))  # Keep sticky via CSS

    header = Thead(Tr(*header_cells))
    body = Tbody(*render_table_rows(df_slice), id="data-table-body")

    # --- Table Wrapper (Horizontal Scroll Only) ---
    table_wrapper = Div(
        Table(header, body, role="grid", id="data-table", cls="table table-sm w-full"),
        id="table-wrapper",
        cls="overflow-x-auto border border-base-300 rounded-lg shadow-sm bg-base-100",  # Added border/shadow
    )

    # --- Pagination --- Moved Here ---
    pagination_html = render_pagination(  # Call the pagination renderer
        current_page, total_pages, per_page, sort_by, sort_desc, current_filters
    )

    # --- Combine Data Grid Components ---
    # This div is the target for HTMX swaps from filters/sorting/pagination
    return Div(
        P(
            Strong(str(total_filtered_rows)),
            Span(" items found", cls="text-base-content/70"),
            cls=f"{TextPresets.muted_sm} mb-2 text-sm",
        ),
        pagination_html,  # Pagination now part of this component
        Div(
            table_wrapper,
            cls="htmx-loading:opacity-50 htmx-loading:transition-opacity htmx-loading:duration-300 ease-in-out",
        ),
        id="data-grid-component",
        cls="flex flex-col text-center justify-center",  # Allow table wrapper to grow if needed (though parent handles scroll)
    )


def render_pagination(
    current_page: int,
    total_pages: int,
    per_page: int,
    sort_by: Optional[str] = None,
    sort_desc: bool = False,
    current_filters: Optional[Dict[str, str]] = None,
) -> FT:
    """Renders pagination controls using flexbox for horizontal layout."""
    if total_pages <= 1:
        return Div(cls="h-12")  # Placeholder height
    current_filters = current_filters or {}
    filter_params = {f"filter_{k}": v for k, v in current_filters.items() if v}
    base_params = {
        "per_page": per_page,
        "sort_by": sort_by or "",
        "sort_desc": str(sort_desc),
        **filter_params,
    }

    # --- Use Div with flex utilities instead of Ul/Li ---
    pagination_items = []

    # Previous Button
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

    # Page Info
    pagination_items.append(
        Span(
            f"Page {current_page} of {total_pages} ",
            cls="underline decoration-pink-500 bg-base-200 text-base-content font-semibold mr-5 ml-5 mb-1",
        )
    )

    # Next Button
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

    # Return Div with flex classes
    return Div(
        *pagination_items,
        aria_label="pagination",
        class_="flex-auto space-x-5 gap-4 pt-4 pb-2",
    )  # Use space-x for horizontal spacing


# --- Routes ---


@rt("/data")
async def data_grid(
    request: Request,
    page: int = 1,
    per_page: int = DEFAULT_PER_PAGE,
    sort_by: Optional[str] = None,
    sort_desc: str = "False",
    # Explicit filter args
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
    start_time = datetime.now()
    current_filters: Dict[str, str] = {
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

    # Perform data operations
    filtered_sorted_df = get_filtered_sorted_df(
        sort_by, sort_desc_bool, current_filters
    )
    total_filtered_rows = filtered_sorted_df.height
    total_pages = math.ceil(total_filtered_rows / per_page) if per_page > 0 else 1
    page = max(1, min(page, total_pages if total_pages > 0 else 1))
    offset = (page - 1) * per_page
    df_slice = filtered_sorted_df.slice(offset, per_page)

    # RENDER THE MAIN CONTENT GRID (filters moved to header, pagination moved inside)
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

    # RENDER HEADER SEPARATELY (Title + Filters)
    header_component = page_header_component(
        current_filters=current_filters,
        per_page=per_page,
        sort_by=sort_by,
        sort_desc=sort_desc_bool,
    )

    # Logging
    end_time = datetime.now()
    processing_time = (end_time - start_time).total_seconds()
    print(
        f"Request: {processing_time:.4f}s | Filters: {current_filters} | Sort: {sort_by} ({'DESC' if sort_desc_bool else 'ASC'}) | Page: {page}"
    )

    is_htmx = request.headers.get("hx-request", "false").lower() == "true"

    if not is_htmx:
        # --- Full Page Response: Render structure with header and content area ---
        return (
            Title("Copyright Data Dashboard"),  # HTML title
            # Main page container (uses global CSS for flex column, full height)
            Div(
                # Content Area (scrollable, includes header + grid)
                Div(
                    header_component,  # Title and Filters
                    grid_component,  # Item Count, Table, Pagination
                    id="content-area",
                ),
                id="page-container",
            ),
        )
    else:
        # HTMX swap: Replace the grid component (count, table, pagination)
        # Triggered by filters, sorting, or pagination clicks
        return grid_component


@rt("/")
async def root_redirect():
    """Redirect root to data grid."""
    return RedirectResponse(url=data_grid.to(), status_code=302)


# --- Run Server ---
if __name__ == "__main__":
    print("Starting FastHTML server...")
    serve(port=8000, reload=True)
