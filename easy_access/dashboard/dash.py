# dash.py

import asyncio
import json
import traceback
from collections import defaultdict
from pathlib import Path

import fasthtml.common as fh
import polars as pl
from fastcore.utils import *
from fasthtml.common import *
from fasthtml.components import Button
from monsterui.all import *
from rich import print
from starlette.responses import Response
from starlette.staticfiles import StaticFiles

from easy_access.dashboard.components import (
    ItemDetailCard,
    create_checkbox_filter_group,
    create_editable_pill_div,
    create_readonly_item_div,
    page_header_component,
    render_contact_info,
    render_course_details,
    render_data_grid_component,
    render_item_history,
    render_modal_field,
    render_teacher_info,
)
from easy_access.dashboard.constants import (
    BADGE_STYLES,
    INIT_HEADERS,
    MODAL_INTERACTION,
    MODAL_TRIGGER,
    PORT,
    ROOT_URL,
)
from easy_access.dashboard.data import (
    fetch_data,
    get_filtered_sorted_df,
    get_item_history,
    process_state,
    store_item_changes,
)
from easy_access.dashboard.files import Entities, get_entities, get_extracted_text
from easy_access.dashboard.urls import URLS as global_urls
from easy_access.dashboard.urls import Url
from easy_access.dashboard.web import Login, bware, load_app_state, login_redir, users
from easy_access.db.retrieve import (
    retrieve_osiris_data,
)
from easy_access.settings import SETTINGS, DirSetting

# ------------------------------
#           app setup
# ------------------------------

app, rt = fast_app(
    before=bware,
    hdrs=INIT_HEADERS,
    exts="loading-states",
    debug=True,
)
dashboard_root = Path(__file__).parent
static_dir = dashboard_root / "static"

# Mount static files directory AFTER creating the app instance
app.mount(
    "/static", StaticFiles(directory=static_dir), name="static_assets"
)  # Renamed for clarity

# use this to enable toast notifications using fasthtml
setup_toasts(app)


# ------------------------------
#         routes
# ------------------------------


@rt("/save_details", methods=["POST"])
async def save_item_details(
    request: Request,
    session: dict,
    material_id: int,
    remarks: str = "",
):
    """
    Route to store edited remarks back to the database
    """
    print(f"Saving remarks for material_id: {material_id}")
    update_data = {
        "material_id": material_id,
        "remarks": remarks,
    }
    try:
        await store_item_changes(update_data, session.get("auth", {}))
        add_toast(session, f"Remarks for item {material_id} saved.", "success")
        return Response(
            status_code=200, headers=HtmxResponseHeaders(trigger="remarksSaveSuccess")
        )

    except Exception as e:
        print(f"Error saving remarks for {material_id}: {e}")
        traceback.print_exc()
        add_toast(session, f"Error saving remarks for {material_id}: {e}", "error")
        return Response(status_code=500)


@rt("/update_field", methods=["POST"])
async def update_single_field(
    session: dict,
    material_id: int,
    field_name: str,
    value: str,
):
    """
    Route to update a single field for an item.
    Used by editable pills for immediate updates.
    """
    print(f"Updating single field for {material_id}: {field_name} = {value}")
    update_data = {
        "material_id": material_id,
        field_name: value if value.lower() != "none" else None,
    }
    try:
        await store_item_changes(update_data, session.get("auth", {}))
        # Return 200 OK, no content needed, HTMX swap='none' will be used client-side
        return Response(status_code=200)
    except Exception as e:
        print(f"Error updating single field for {material_id}: {e}")
        traceback.print_exc()
        # Optionally add a toast message for the error
        add_toast(session, f"Error updating {field_name}", "error")
        # Return an error status code
        return Response(status_code=500)


@rt("/osiris/{material_id:int}")
async def get_osiris_data(material_id: int):
    """
    Returns enriched data for the given material_id.
    should always return at least base item data, even if no enriched data is found.
    Logic is almost completely found in the db.retrieve module.

    """
    data = await asyncio.to_thread(retrieve_osiris_data, [material_id], SETTINGS)
    if not data or len(data) == 0 or not isinstance(data, list):
        return []
    return data[0]


@rt("/pdf/{material_id:int}")
async def get_pdf_element(material_id: int):
    """
    Returns element displaying the PDF file for the given material_id as an embedded PDF viewer.
    If the PDF file is not found, a message will be displayed instead.
    """

    pdf_file_path = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full / f"{material_id}.pdf"
    print(f"PDF file path: {pdf_file_path}")
    pdf_element = None
    if not pdf_file_path.exists():
        pdf_element = Div("PDF file not found", cls="text-red-500")
    else:
        pdf_element = Embed(
            src=ROOT_URL + f"/file/{material_id}",
            type="application/pdf",
            width="100%",
            height="800px",
        )

    return pdf_element


@rt(r"/file/{material_id:int}")
async def get_file(material_id: int):
    """
    Directly returns the PDF file for the given material_id.
    """
    file_path = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full / f"{material_id}.pdf"
    print(f"requested file: {file_path}")
    if not file_path.exists():
        return HTMLResponse("File not found", status_code=404)
    return FileResponse(file_path)


@rt("/text/{material_id:int}", methods=["GET"])
async def get_extracted_text_element(material_id: int) -> FT:
    text = get_extracted_text(material_id)
    return Div(NotStr(text))


@rt("/entities/{material_id:int}")
async def get_entities_element(material_id: int):
    """
    Returns element displaying entities extracted from the annotated text for the given material_id.
    If no entities are found, a message will be displayed.
    """
    entities_element = Div(
        H4("No entities found (yet?) in the text"),
        cls="mt-4",
    )

    entities: Entities | None = get_entities(material_id)
    if not entities:
        return entities_element

    entities_by_label, sorted_labels = entities.group_and_sort()
    entities_elements: list[tuple[str, list[FT]]] = []
    for label in sorted_labels:
        entity_list = entities_by_label[label]
        cur_list = []
        seen_items = defaultdict(int)
        for entity in entity_list:
            seen_items[entity.text] += 1
            if entity.text in seen_items and seen_items[entity.text] > 1:
                continue
        added_items = set()
        for entity in entity_list:
            entity_element = []
            if entity.text in added_items:
                continue
            if seen_items[entity.text] > 1:
                entity_element.append(
                    Span(
                        Span(
                            f"{seen_items[entity.text]}x",
                            cls="badge badge-info mr-2 ml-1",
                        ),
                        Span(f"{entity.text}"),
                    )
                )
            else:
                entity_element.append(Span(entity.text, cls="ml-1"))
            added_items.add(entity.text)
            if "recognized" in label.lower():
                color_cls = "badge badge-accent badge-outline"
                entity.score = 1.0
            elif entity.score is not None:
                if entity.score < 0.9:
                    color_cls = "badge badge-warning"
                else:
                    color_cls = "badge badge-success"
            else:
                color_cls = "badge badge-outline"
                entity.score = 1
            entity_element.append(
                Span(f"{entity.score:.0%}", cls=f"ml-4 gap-2 {color_cls}")
            )
            cur_list.append(
                Li(
                    *entity_element,
                )
            )

        entities_elements.append((label, cur_list))

    entities_element = Div(
        *[
            Div(
                H5(label),
                Ul(
                    *entity_list_objs,
                    cls="list-disc",
                ),
            )
            for label, entity_list_objs in entities_elements
        ]
    )

    return entities_element


# --- main page routes ---


@rt("/data")
async def data_grid(session: dict, request: Request):
    """
    Main route for displaying/updating the data grid.
    Handles pagination, sorting, filtering requests via HTMX or initial page load.
    Orchestrates state processing and data fetching.
    """
    # 1. Parse request parameters (combining query and form)
    form_data_dict = {}
    content_type = request.headers.get("content-type", "").lower()
    # Check method or content type to decide if form parsing is needed/possible
    if (
        request.method == "POST"
        or "application/x-www-form-urlencoded" in content_type
        or "multipart/form-data" in content_type
    ):
        try:
            form = await request.form()
            form_data_dict = dict(form)
        except Exception as e:
            # Gracefully handle cases like GET requests with unexpected content-type,
            # or malformed form data. Log the error.
            print(f"Warning: Could not parse form data (maybe GET request?): {e}")
            form_data_dict = {}  # Ensure it's an empty dict

    request_params = {**dict(request.query_params), **form_data_dict}
    print("--- Request to /data ---")
    print(f"Method: {request.method}, Params: {request_params}")

    # 2. Process state based on request (updates session['app_state'])
    # process_state handles loading, applying changes, and saving back to session
    session, requested_app_state = process_state(session, request_params)

    # 3. Fetch data and get validated state
    # fetch_data handles auth constraints, filtering, validation (page#), counts
    processed_data = await fetch_data(requested_app_state, session)

    # 4. Update session *again* only if page validation changed the state
    # This ensures the session reflects the *actual* page being displayed
    if processed_data.app_state.page != requested_app_state.page:
        print("Updating session state again due to page validation.")
        session["app_state"] = asdict(processed_data.app_state)

    # 5. Render the main grid component using processed data
    grid_component = render_data_grid_component(
        df_slice=processed_data.df_slice,
        app_state=processed_data.app_state,  # Use the final, validated state
        total_filtered_rows=processed_data.total_filtered_rows,
        total_pages=processed_data.total_pages,
    )

    # 6. Render OOB filter fragments
    oob_filter_fragments = []
    auth_details = session.get("auth", {})  # Get auth details for rendering logic
    is_admin = auth_details.get("role") == "admin"
    user_faculty = auth_details.get("faculty")

    checkbox_groups_to_render = [
        ("workflow_status", "Workflow Status", BADGE_STYLES["workflow_status"]),
        ("status", "Status", BADGE_STYLES["status"]),
        ("classification", "Classification", BADGE_STYLES["classification"]),
        (
            "manual_classification",
            "Manual Classification",
            BADGE_STYLES["classification"],
        ),
    ]
    # Conditionally add faculty group if user is admin or has no specific faculty assigned
    if is_admin or not user_faculty or user_faculty == "all":
        checkbox_groups_to_render.append(
            ("faculty", "Faculty", BADGE_STYLES["faculty"])
        )

    current_total = processed_data.total_filtered_rows
    print(
        f"Current total in data_grid when starting to render checkbox groups: {current_total}"
    )
    for key, label, options_map in checkbox_groups_to_render:
        group_id = f"filter-group-{key}"
        counts_for_group = processed_data.filter_counts.get(key, {})

        rendered_group = create_checkbox_filter_group(
            filter_key=key,
            label_text=label,
            options=options_map,
            current_values=processed_data.app_state.filters.get(key),
            counts=counts_for_group,
            current_total=current_total,
        )
        if hasattr(rendered_group, "attrs"):
            rendered_group.attrs["hx-swap-oob"] = f"outerHTML:#{group_id}"
        else:
            print(f"Warning: Cannot add OOB swap to non-FT object for key {key}")
        oob_filter_fragments.append(rendered_group)

    # 7. Determine Return Value (HTMX vs Full Page)
    is_htmx = request.headers.get("hx-request", "false").lower() == "true"

    if is_htmx:
        print("--- HTMX Request: Returning Grid + OOB Filters ---")
        return (grid_component, *oob_filter_fragments)
    else:
        print("--- Full Page Request: Rendering Header + Grid + Modals + Scripts ---")

        header_component = page_header_component(
            auth_details,
            processed_data.app_state,
            processed_data.filter_counts,
            processed_data.total_filtered_rows,
        )

        modal_placeholder = Dialog(
            id="modal-placeholder",
            cls="modal modal-bottom sm:modal-middle",
            hx_trigger="close",
            hx_get=data_grid.to(),
            hx_target="#data-grid-component",
            hx_swap="outerHTML",
        )
        return (
            Title("CDD//UT Dashboard"),
            MODAL_TRIGGER,
            MODAL_INTERACTION,
            Div(
                Div(header_component, grid_component, id="content-area"),
                id="page-container",
                hx_ext="preload",
            ),
            modal_placeholder,
        )


@rt("/modal/{material_id:int}")
async def show_item_details(session: dict, material_id: int):
    """Fetches data and returns structured INNER content for the modal dialog."""

    # Load state primarily to determine prev/next IDs based on current view
    app_state = load_app_state(session)
    auth_details = session.get("auth", {})

    # Determine faculty constraint for filtering to get correct prev/next
    faculty_constraint = None
    user_faculty = auth_details.get("faculty")
    user_role = auth_details.get("role")
    if user_role != "admin" and user_faculty and user_faculty != "all":
        faculty_constraint = {"faculty": user_faculty}

    # Get the ordered list of IDs based on current view state + auth
    try:
        filtered_sorted_df = get_filtered_sorted_df(
            app_state, extra_constraints=faculty_constraint
        )
        ordered_ids = filtered_sorted_df.get_column("material_id").to_list()
        current_index = ordered_ids.index(material_id)
    except (ValueError, pl.exceptions.ColumnNotFoundError):
        # Handle case where item not found in current view or ID column missing
        print(
            f"Warning: Material ID {material_id} not found in current filtered/sorted view."
        )
        current_index = -1
        ordered_ids = []  # Ensure list is empty
    except Exception as e:
        print(f"Error getting ordered IDs for modal prev/next: {e}")
        current_index = -1
        ordered_ids = []

    prev_id = ordered_ids[current_index - 1] if current_index > 0 else None
    next_id = (
        ordered_ids[current_index + 1]
        if current_index != -1 and current_index < len(ordered_ids) - 1
        else None
    )

    def get_val(key, default=None):
        return item_data.get(key, default)

    try:
        nested_data_list = await asyncio.to_thread(
            retrieve_osiris_data, [material_id], SETTINGS
        )
        if not nested_data_list:
            raise ValueError(
                f"Material ID {material_id} not found in detailed data source."
            )
        item_data = nested_data_list[0]

        if "faculty_id" in item_data and "faculty" not in item_data:
            item_data["faculty"] = item_data.pop("faculty_id")

        # --- Header Row  ---
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
        # add material_id pill right below the status pill
        status_pill = Div(
            status_pill,
            cls="flex items-center space-x-2",
        )(
            Label(
                "🆔 " + str(material_id),
                cls="text-sm",
            ),
        )

        url_element, _ = render_modal_field("url", file_url)

        header_content = Div(cls="flex items-center justify-between space-x-4")(
            Div(cls="flex items-center space-x-3 flex-grow min-w-0")(
                status_pill, filename_content, url_element
            ),
            Form(method="dialog")(
                Button("✕", cls="btn btn-sm btn-circle btn-ghost flex-shrink-0")
            ),
        )

        # --- Cards ---
        # --- Data Entry Card ---
        data_entry_content = (
            # --- Pass material_id here ---
            create_editable_pill_div(
                get_val("workflow_status"),
                "Workflow Status",
                "workflow_status",
                BADGE_STYLES["workflow_status"],
                material_id=material_id,
            ),
            # --- Pass material_id here ---
            create_editable_pill_div(
                get_val("manual_classification"),
                "Manual Classification",
                "manual_classification",
                BADGE_STYLES["classification"],
                material_id=material_id,
            ),
            # --- Remarks Text Area (AlpineJS data/model) ---
            Div(
                Strong(
                    "Remarks", cls="block text-xs font-medium text-base-content/80 mb-1"
                ),
                Div()(
                    TextArea(
                        x_model="remarks",
                        id="modal_remarks",
                        name="remarks",
                        rows="5",
                        cls="textarea textarea-bordered w-full text-sm bg-base-100",
                    )
                ),
                cls="mb-3",
            ),
            # --- Buttons (Modify AlpineJS logic) ---
            Div(cls="flex justify-end space-x-2 mt-4")(
                Button(
                    "Reset",
                    type="button",
                    cls=ButtonT.secondary + " btn-sm",
                    **{
                        "@click": "remarks = originalRemarks; document.getElementById('modal-save-btn').disabled = true; document.getElementById('save-indicator').classList.add('hidden');"
                    },
                ),
                Button(
                    "Save",
                    Span(
                        cls="absolute top-0 right-0 -mr-1 -mt-1 flex size-3 hidden",
                        id="save-indicator",
                    )(
                        # Ping animation span
                        Span(
                            cls="absolute inline-flex h-full w-full animate-ping rounded-full bg-info opacity-75"
                        ),
                        # Visible dot span
                        Span(cls="relative inline-flex size-3 rounded-full bg-info"),
                    ),
                    id="modal-save-btn",
                    type="submit",
                    cls=ButtonT.primary + " btn-sm relative",
                    **{
                        ":disabled": "remarks == originalRemarks",
                        ":class": "{ 'opacity-50 cursor-not-allowed': remarks == originalRemarks }",
                        "x-init": "$watch('remarks', value => { document.getElementById('save-indicator').classList.toggle('hidden', value == originalRemarks) })",
                    },
                ),
            ),
        )

        # --- Item Info Card ---
        item_info_fields = [
            "classification",
            "ml_prediction",
            "period",
            "faculty",
            "owner",
            "department",
            "course_name",
            "course_code",
        ]
        item_info_details = [
            (get_val(k), k.replace("_", " ").title(), k) for k in item_info_fields
        ]
        item_info_content = [
            create_readonly_item_div(*item_info) for item_info in item_info_details
        ]

        # --- Text Details Card ---
        text_details_fields = ["title", "author", "publisher", "doi", "isbn"]
        text_details = [
            (get_val(k), k.replace("_", " ").title(), k) for k in text_details_fields
        ]
        text_details_content = [
            create_readonly_item_div(*text_detail) for text_detail in text_details
        ]
        # --- Numeric Details Card ---
        numeric_details_fields = ["pagecount", "wordcount", "picturecount"]
        numeric_details = [
            (get_val(k), k.replace("_", " ").replace("count", "").title(), k)
            for k in numeric_details_fields
        ]
        numeric_details_content = [
            create_readonly_item_div(*numeric_detail)
            for numeric_detail in numeric_details
        ]

        # --- Osiris Data Cards ---
        contact_info_content = render_contact_info(item_data)
        course_details_content = render_course_details(item_data)
        teachers_content = render_teacher_info(item_data)

        # --- Item History Card ---
        item_history_card = render_item_history(await get_item_history(material_id))

        # --- Assemble Cards into Grid ---
        modal_cards_grid = Div(cls="grid grid-cols-1 md:grid-cols-3 gap-4")(
            # Column 1
            Div(
                ItemDetailCard(
                    "Data Entry",
                    *data_entry_content,
                    card_id="data-entry-card",
                    col_span=1,
                ),
                ItemDetailCard(
                    "Entities",
                    Span("Click to load...", cls="italic text-sm"),
                    card_id="entities-card",
                    col_span=1,
                    start_collapsed=True,
                    lazy_load_url=get_entities_element.to(material_id=material_id),
                ),
                cls="flex flex-col space-y-4",
            ),
            # Column 2
            Div(
                ItemDetailCard(
                    "Item Info",
                    *item_info_content,
                    card_id="item-info-card",
                    start_collapsed=True,
                    col_span=1,
                ),
                ItemDetailCard(
                    "Text Details",
                    *text_details_content,
                    card_id="text-details-card",
                    start_collapsed=True,
                    col_span=1,
                ),
                ItemDetailCard(
                    "Counts",
                    *numeric_details_content,
                    card_id="counts-card",
                    start_collapsed=True,
                    col_span=1,
                ),
                item_history_card,
                cls="flex flex-col space-y-4",
            ),
            # Column 3
            Div(
                ItemDetailCard(
                    "Contact Info",
                    *contact_info_content,
                    card_id="contact-info-card",
                    col_span=1,
                ),
                ItemDetailCard(
                    "Course Details",
                    *course_details_content,
                    card_id="course-details-card",
                    start_collapsed=True,
                    col_span=1,
                ),
                ItemDetailCard(
                    "Teachers",
                    *teachers_content,
                    card_id="teachers-card",
                    start_collapsed=True,
                    col_span=1,
                ),
                cls="flex flex-col space-y-4",
            ),
            # Full Width Lazy Loaded Cards
            ItemDetailCard(
                "PDF Viewer",
                Span("Click to load...", cls="italic text-sm"),
                card_id="pdf-card",
                col_span=3,
                start_collapsed=True,
                lazy_load_url=get_pdf_element.to(material_id=material_id),
            ),
            ItemDetailCard(
                "Extracted Text",
                Span("Click to load...", cls="italic text-sm"),
                card_id="text-card",
                col_span=3,
                start_collapsed=True,
                lazy_load_url=get_extracted_text_element.to(material_id=material_id),
            ),
        )

        # --- Footer Row  ---

        # Prev Button
        prev_button_attrs = {
            "id": "modal-prev-btn",
            "cls": ButtonT.secondary + " btn-sm",
            "disabled": prev_id is None,
        }
        if prev_id is not None:
            prev_button_attrs.update(
                {
                    "hx_get": show_item_details.to(material_id=prev_id),  # Only need ID
                    "hx_target": "#modal-placeholder",
                    "hx_swap": "innerHTML",
                }
            )
        prev_button = Button("< Prev", **prev_button_attrs)

        # Next Button
        next_button_attrs = {
            "id": "modal-next-btn",
            "cls": ButtonT.secondary + " btn-sm",
            "disabled": next_id is None,
        }
        if next_id is not None:
            next_button_attrs.update(
                {
                    "hx_get": show_item_details.to(material_id=next_id),  # Only need ID
                    "hx_target": "#modal-placeholder",
                    "hx_swap": "innerHTML",
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

        # --- Assemble Modal Box ---
        modal_box_content = Div(
            Div(
                Div(header_content, cls="border-b pb-2 flex-shrink-0"),
                Div(
                    Form(
                        Input(type="hidden", name="material_id", value=material_id),
                        modal_cards_grid,
                        id="modal-details-form",
                        hx_post=save_item_details.to(),
                        hx_include="[name='material_id'], [name='remarks']",
                        hx_target="body",
                        hx_swap="none",
                        x_data=f"{{ remarks: {json.dumps(get_val('remarks', ''))}, originalRemarks: {json.dumps(get_val('remarks', ''))} }}",
                        **{
                            "@remarks-save-success.window": "originalRemarks = remarks; console.log('Remarks saved, updated originalRemarks.')"
                        },
                    ),
                    cls="relative py-4 flex-grow overflow-y-auto",
                ),
                footer_content,
                cls="flex-shrink-0",
            ),
            cls="modal-box w-[85vw] max-w-none h-[calc(100vh-5rem)] max-h-none flex flex-col",
        )

        modal_backdrop = Form(method="dialog", cls="modal-backdrop")(
            NotStr(
                '<button class="absolute inset-0 w-full h-full cursor-default outline-none" aria-label="close modal"></button>'
            )
        )

        response_content = (modal_box_content, modal_backdrop)
        return response_content, HtmxResponseHeaders(trigger="openModalEvent")

    except Exception as e:
        print(f"Error generating modal content for ID {material_id}: {e}")
        print(traceback.format_exc())
        modal_box_content = Div(cls="modal-box")(
            H3("Error"),
            P(f"Could not load details for item {material_id}."),
            P(f"Details: {e}", cls="text-xs text-error"),
            Form(method="dialog")(Button("Close", cls="btn btn-sm mt-4")),
        )
        modal_backdrop = Form(method="dialog", cls="modal-backdrop")(
            NotStr(
                '<button class="absolute inset-0 w-full h-full cursor-default outline-none" aria-label="close"></button>'
            )
        )
        return (modal_box_content, modal_backdrop), HtmxResponseHeaders(
            trigger="openModalEvent"
        )


@rt("/login")
def get():
    frm = Form(
        Div(
            NotStr(
                r'<label class="block text-gray-700 text-sm font-bold mb-2" for="email">Email</label>'
            ),
            Input(
                id="email",
                type="text",
                placeholder="Email",
                name="email",
                cls="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline",
            ),
            cls="mb-4",
        ),
        Div(
            NotStr(
                r'<label class="block text-gray-700 text-sm font-bold mb-2" for="password">Password</label>'
            ),
            Input(
                id="password",
                name="pwd",
                type="password",
                placeholder="******************",
                cls="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline",
            ),
            cls="mb-6",
        ),
        fh.Button(
            "login",
            cls="btn btn-secondary bg-indigo-500 font-bold py-2 px-4 rounded focus:outline-none focus:shadow-outline",
            type="submit",
        ),
        action="/login",
        method="post",
        cls="bg-white shadow-md rounded px-8 pt-6 pb-8 mb-4",
    )
    return Title("Login to CDD//UT "), Div(
        Card(
            CardTitle(
                NotStr(
                    "<p>Login to the <strong class='underline decoration-pink-500'>C</strong>opyright <strong class='underline decoration-pink-500'>D</strong>ata <strong class='underline decoration-pink-500'>D</strong>ashboard</p>"
                ),
                Img(src="static/ut_logo.png"),
                cls="text-center text-xl",
            ),
            CardBody(frm),
            cls="uk-card-secondary bg-indigo-400",
        ),
        cls="flex items-center justify-center h-screen",
    )


@rt("/login")
def post(login: Login, sess):  # login uses dataclass binding
    if not login.email or not login.pwd:
        return login_redir
    try:
        u = users[login.email]  # Query user by email (primary key)
        if not u:
            raise NotFoundError  # If user dict is empty/None
    except NotFoundError:
        print(f"Login attempt failed: User '{login.email}' not found.")
        # Optional: Add toast message for user feedback
        add_toast(sess, "Invalid email or password", "error")  # Requires setup_toasts
        return login_redir  # Redirect back to login on failure

    # Verify password
    if not compare_digest(u.get("pwd", "").encode("utf-8"), login.pwd.encode("utf-8")):
        print(f"Login attempt failed: Incorrect password for '{login.email}'.")
        add_toast(sess, "Invalid email or password", "error")
        return login_redir  # Redirect back to login on failure

    # --- Store structured auth info in session ---
    sess["auth"] = {
        "email": login.email,
        "name": u.get("name"),
        "faculty": u.get("faculty"),
        "role": u.get("role"),
    }
    if "app_state" in sess:
        del sess["app_state"]
    print(f"Login successful for '{login.email}'. Redirecting to data grid.")
    add_toast(sess, f"Login successful! Welcome, {u.get('name')}!", "success")
    return RedirectResponse(data_grid.to(), status_code=303)


@rt("/logout")
def logout(sess):
    if "auth" in sess:
        del sess["auth"]
    if "app_state" in sess:
        del sess["app_state"]
    return login_redir


# --- root route ---
@rt("/")
async def root_redirect():
    """Redirect root to data grid."""
    return RedirectResponse(url=data_grid.to(), status_code=302)


URLS = {
    Url.root: root_redirect.to(),
    Url.data_grid: data_grid.to(),
    Url.logout: logout.to(),
    Url.show_item_details: show_item_details.to(),
    Url.get_entities_element: get_entities_element.to(),
    Url.get_pdf_element: get_pdf_element.to(),
    Url.get_extracted_text_element: get_extracted_text_element.to(),
    Url.save_item_details: save_item_details.to(),
    Url.get_file: get_file.to(),
    Url.get_osiris_data: get_osiris_data.to(),
    Url.update_single_field: update_single_field.to(),
}
global_urls.update(URLS)


def start():
    """
    Start the FastHTML server with the defined routes.
    """

    print("Starting the FastHTML server...")
    serve(
        app=app,
        port=PORT,
        reload=True,
    )


# if __name__ == "__main__":
#    start()
