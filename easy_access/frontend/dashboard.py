import math

import polars as pl
from fasthtml.common import *

from easy_access.db.retrieve import retrieve_copyright_items

app, rt = fast_app()

copyright_df = retrieve_copyright_items()


def render_table(df_slice: pl.DataFrame) -> FT:
    """Renders a Polars DataFrame slice into an HTML table using FastTags."""
    if df_slice is None or df_slice.height == 0:
        return P("No data to display for this page.")

    cols = df_slice.columns
    header = Thead(Tr(*[Th(col) for col in cols]))

    rows = []
    for row_dict in df_slice.iter_rows(named=True):
        cells = []
        for col in cols:
            val = row_dict[col]
            cell_content = str(val) if val is not None else ""
            cells.append(Td(cell_content))
        rows.append(Tr(*cells))

    body = Tbody(*rows)
    return Table(header, body, role="grid")


def render_pagination(current_page: int, total_pages: int, per_page: int) -> FT:
    """Renders pagination controls."""
    if total_pages <= 1:
        return Div()

    items = []
    prev_disabled = current_page <= 1
    prev_page = max(1, current_page - 1)
    items.append(
        Li(
            A(
                "« Previous",
                **(
                    {"aria-disabled": "true"}
                    if prev_disabled
                    else {"href": index.to(page=prev_page, per_page=per_page)}
                ),
            )
        )
    )

    items.append(Li(Span(f" Page {current_page} of {total_pages} ")))

    next_disabled = current_page >= total_pages
    next_page = min(total_pages, current_page + 1)
    items.append(
        Li(
            A(
                "Next »",
                **(
                    {"aria-disabled": "true"}
                    if next_disabled
                    else {"href": index.to(page=next_page, per_page=per_page)}
                ),
            )
        )
    )

    return Nav(Ul(*items), aria_label="pagination", style="margin-top: 1rem;")


@rt("/")
def index(page: int = 1, per_page: int = 10):
    """Main dashboard route displaying the paginated table."""
    global copyright_df

    total_rows = copyright_df.height
    total_pages = math.ceil(total_rows / per_page)

    page = max(1, min(page, total_pages))

    offset = (page - 1) * per_page

    df_slice = copyright_df.slice(offset, per_page)

    table_html = render_table(df_slice)
    pagination_html = render_pagination(page, total_pages, per_page)

    return Titled(
        "Copyright Data Dashboard",
        P(
            f"Displaying rows {offset + 1} to {min(offset + per_page, total_rows)} of {total_rows}."
        ),
        table_html,
        pagination_html,
    )
