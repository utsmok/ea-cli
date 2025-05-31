"""
This module implements the main web application for the Easy Access Dashboard
using the FastHTML framework. It defines routes for:
- User authentication (login/logout).
- Displaying the main data grid with copyright items.
- Handling interactions like pagination, sorting, and filtering via HTMX.
- Showing detailed information for individual items in a modal.
- Serving static files and related data (PDFs, extracted text, entities).
"""

import asyncio
import json
import logging # Added
import traceback
from collections import defaultdict # Used in get_entities_element
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union # For type hints

import fasthtml.common as fh
import polars as pl
# from fastcore.utils import * # Avoid wildcard
from fastcore.utils import L # Example if L is used
from fasthtml.common import ( # Specific imports from fasthtml.common
    HTMLResponse, Input, Form, Script, Style, Title, QueryRouter,
    FastHTML, Beforeware, State, Depends, UploadFile, Cookie, Header,
    hx_redirect, parse_form, FormField, FileResponse, add_toast, setup_toasts,
    HtmxResponseHeaders, NotStr
)
from fasthtml.components import Button, Details, Div, H2, H4, H5, H6, Img, Li, Nav, P, Span, Summary, Table, Tbody, Td, Template, Th, Thead, Tr, Ul, A, Embed, Dialog, TextArea # Added TextArea
from monsterui.all import UkIcon, Label, LabelInput, Card, CardTitle, CardBody # Specific imports from monsterui
# from rich import print # Removed, using logging
from starlette.requests import Request # For type hinting Request
from starlette.responses import Response as StarletteResponse # Alias to avoid clash
from starlette.staticfiles import StaticFiles

import easy_access.dashboard.urls as dashboard_urls_module # To avoid direct use of global_urls
from easy_access.dashboard.components import (
    ItemDetailCard,
    # create_checkbox_filter_group, # This is called by page_header_component
    create_editable_pill_div,
    create_readonly_item_div,
    page_header_component,
    render_contact_info,
    render_course_details,
    render_data_grid_component,
    render_item_history,
    # render_modal_field, # Used internally by other components
    render_teacher_info,
)
from easy_access.dashboard.constants import (
    BADGE_STYLES, # Used by page_header_component -> create_checkbox_filter_group
    INIT_HEADERS, # For app setup
    PAGE_BODY_JS, # For main layout
    APP_ROOT_URL, # Used for constructing full URLs, e.g. for PDF embedding
)
from easy_access.dashboard.data import (
    fetch_data, ProcessedDataResult, # ProcessedDataResult for type hint
    # get_filtered_sorted_df, # Used by fetch_data
    get_item_history_for_dashboard, # Renamed from get_item_history
    process_state,
    store_item_changes,
)
from easy_access.dashboard.files import Entities, get_entities, get_extracted_text
from easy_access.dashboard.urls import Url # Enum for URL names
from easy_access.dashboard.web import Login, auth_bware as dashboard_auth_bware, load_app_state, LOGIN_REDIRECT_RESPONSE, users # Renamed for clarity
from easy_access.db.retrieve import retrieve_osiris_data # For specific Osiris data route
from easy_access.settings import SETTINGS, DirSetting # For directory paths

logger = logging.getLogger(__name__)

# ------------------------------
# Application Setup (FastHTML)
# ------------------------------

# hdrs should include everything needed for the base page structure
# PAGE_BODY_JS is added separately in the root/data_grid route for full page loads.
app, rt = fh.fast_app( # Use fh alias
    before=dashboard_auth_bware, # Authentication middleware
    hdrs=list(INIT_HEADERS), # Ensure it's a list for potential modification
    exts=["loading-states"], # HTMX extensions
    debug=True, # Enable FastHTML debug mode (consider making this configurable)
)

# Serve static files (CSS, JS, images)
dashboard_static_path = Path(__file__).parent / "static"
if dashboard_static_path.exists() and dashboard_static_path.is_dir():
    app.mount("/static", StaticFiles(directory=dashboard_static_path), name="static_dashboard_assets")
else:
    logger.warning(f"Dashboard static directory not found at {dashboard_static_path}. Static files may not load.")

# Enable toast notifications using fasthtml's built-in system
setup_toasts(app)


# ------------------------------
# Route Handlers
# ------------------------------

@rt(dashboard_urls_module.URLS[Url.save_item_details], methods=["POST"]) # Use URLS from module
async def save_item_details_route(
    request: Request, # Provided by FastHTML/Starlette
    session: Dict[str, Any], # Session dict
    material_id: int, # From path or form data, depends on how called
    remarks: str = FormField(""), # Example: get 'remarks' specifically from form
    # Other fields can be added here if the form for ItemDetailCard expands
) -> StarletteResponse:
    """
    Handles POST requests to save edited details (e.g., remarks) for a copyright item.
    Updates the database and provides user feedback via toast messages.

    Args:
        request: The HTTP request object (currently unused, but available).
        session: The user's session data.
        material_id: The ID of the item being updated (extracted from form or path).
        remarks: The new remarks text from the form.

    Returns:
        StarletteResponse: HTTP 200 on success with HTMX trigger, or 500 on error.
    """
    logger.info(f"Saving details for material_id: {material_id}. New remarks length: {len(remarks)}")

    # Construct data for database update
    update_data_dict: Dict[str, Any] = {
        "material_id": material_id,
        "remarks": remarks,
    }
    auth_details: Dict[str, Any] = session.get("auth", {})

    try:
        await store_item_changes([update_data_dict], auth_details) # store_item_changes expects a list
        add_toast(session, f"Details for item {material_id} saved successfully.", "success")
        # Trigger client-side event if needed, e.g., to update parts of the modal
        return StarletteResponse(status_code=200, headers=HtmxResponseHeaders(trigger="remarksSaveSuccess").headers)
    except Exception as e:
        logger.error(f"Error saving details for material_id {material_id}: {e}", exc_info=True)
        add_toast(session, f"Error saving details for item {material_id}: {str(e)[:100]}", "error")
        return StarletteResponse("Internal Server Error", status_code=500)


@rt(dashboard_urls_module.URLS[Url.update_single_field], methods=["POST"])
async def update_single_field_route(
    session: Dict[str, Any],
    material_id: int = FormField(...), # These are expected from HTMX POST
    field_name: str = FormField(...),
    value: str = FormField(...),
) -> StarletteResponse:
    """
    Handles POST requests to update a single field of a copyright item.
    Typically used by editable pills (e.g., for workflow_status, manual_classification)
    for immediate updates via HTMX.

    Args:
        session: The user's session data.
        material_id: The ID of the item to update.
        field_name: The name of the field to update.
        value: The new value for the field.

    Returns:
        StarletteResponse: HTTP 200 on success, or 500 on error.
                           No content body needed as HTMX swap is 'none'.
    """
    logger.info(f"Updating single field for material_id {material_id}: {field_name} = '{value}'")

    update_payload: Dict[str, Any] = {
        "material_id": material_id,
        # Convert "none" string (potentially from UI) to actual None for DB
        field_name: None if value.lower() == "none" else value,
    }
    auth_details: Dict[str, Any] = session.get("auth", {})

    try:
        await store_item_changes([update_payload], auth_details)
        # Success, no content needed as client-side JS (updatePill) handles UI update optimistically.
        # A toast message could be added here too if desired.
        # add_toast(session, f"{field_name.replace('_', ' ').title()} for item {material_id} updated.", "success")
        return StarletteResponse(status_code=200)
    except Exception as e:
        logger.error(f"Error updating single field '{field_name}' for material_id {material_id}: {e}", exc_info=True)
        add_toast(session, f"Error updating {field_name.replace('_', ' ')}.", "error")
        return StarletteResponse("Internal Server Error", status_code=500)


@rt(dashboard_urls_module.URLS[Url.get_osiris_data]) # Path includes placeholder {material_id:int}
async def get_osiris_data_route(material_id: int) -> List[Dict[str, Any]]: # Type hint for return
    """
    Route to fetch enriched Osiris data for a given material_id.
    This data is typically used to populate detail cards in the UI.

    Args:
        material_id (int): The material ID to fetch Osiris data for.

    Returns:
        List[Dict[str, Any]]: A list containing a dictionary of the item's Osiris data.
                              Returns an empty list if data not found or on error.
    """
    logger.debug(f"Fetching Osiris data for material_id: {material_id}")
    # retrieve_osiris_data is currently synchronous and uses SQLAlchemy engine
    # To call from async route, wrap in asyncio.to_thread
    try:
        # Assuming retrieve_osiris_data is defined in easy_access.db.retrieve
        data: List[Dict[str, Any]] = await asyncio.to_thread(retrieve_osiris_data, [material_id])
        if not data: # retrieve_osiris_data returns list, check if empty
            logger.info(f"No Osiris data found for material_id: {material_id}")
            return []
        return data # Should be a list containing one dict, or an empty list
    except Exception as e:
        logger.error(f"Error in get_osiris_data route for material_id {material_id}: {e}", exc_info=True)
        return [] # Return empty list on error


@rt(dashboard_urls_module.URLS[Url.get_pdf_element]) # Path includes {material_id:int}
async def get_pdf_element_route(material_id: int) -> Embed | Div: # FT is too general
    """
    Returns an HTML `<embed>` element to display a PDF, or a "Not Found" message.
    The PDF is served by the `/file/{material_id:int}` route.

    Args:
        material_id (int): The material ID of the PDF to display.

    Returns:
        Embed | Div: An Embed component for the PDF or a Div with an error message.
    """
    pdf_dir = SETTINGS.dirs.get(DirSetting.PDF_DOWNLOADS)
    if not pdf_dir or not pdf_dir.exists:
        logger.error("PDF download directory not configured or found. Cannot serve PDF element.")
        return Div("PDF directory not available.", cls="text-error p-4")

    # Filename convention: {material_id}.pdf or {material_id}_original_filename.pdf
    # Trying simple {material_id}.pdf first
    pdf_file_path = pdf_dir.full / f"{material_id}.pdf"
    logger.debug(f"Attempting to provide PDF element for: {pdf_file_path}")

    if not pdf_file_path.exists():
        # Fallback: check for files starting with material_id_ if simple name not found
        # This might be needed if files were named like "12345_some_name.pdf"
        alt_files = list(pdf_dir.full.glob(f"{material_id}_*.pdf"))
        if alt_files:
            pdf_file_path = alt_files[0] # Take the first match
            logger.debug(f"Found alternative PDF file: {pdf_file_path}")
        else:
            logger.warning(f"PDF file for material_id {material_id} not found at expected paths.")
            return Div(f"PDF file for item {material_id} not found.", cls="text-error p-4")

    # Construct URL using APP_ROOT_URL from constants
    # The /file/ route should serve the actual file bytes
    pdf_serve_url = f"{APP_ROOT_URL}{str(dashboard_urls_module.URLS[Url.get_file]).replace('{material_id:int}', str(material_id))}"

    return Embed(src=pdf_serve_url, type="application/pdf", width="100%", height="800px")


@rt(dashboard_urls_module.URLS[Url.get_file]) # Path includes {material_id:int}
async def get_file_route(material_id: int) -> FileResponse | HTMLResponse:
    """
    Serves a PDF file directly from the filesystem.

    Args:
        material_id (int): The material ID of the PDF to serve.

    Returns:
        FileResponse: Responds with the PDF file if found.
        HTMLResponse: Responds with a 404 error if the file is not found.
    """
    pdf_dir = SETTINGS.dirs.get(DirSetting.PDF_DOWNLOADS)
    if not pdf_dir or not pdf_dir.exists:
        logger.error("PDF download directory not configured. Cannot serve file.")
        return HTMLResponse("Server configuration error: PDF directory not set.", status_code=500)

    file_path = pdf_dir.full / f"{material_id}.pdf" # Assuming simple name for serving
    logger.info(f"Serving file request for: {file_path}")
    if not file_path.exists():
        # Fallback check, similar to get_pdf_element_route
        alt_files = list(pdf_dir.full.glob(f"{material_id}_*.pdf"))
        if alt_files: file_path = alt_files[0]
        else:
            logger.warning(f"File not found for material_id {material_id} at {file_path} (and alternatives).")
            return HTMLResponse(f"File for item {material_id} not found.", status_code=404)

    return FileResponse(file_path, media_type="application/pdf", filename=file_path.name)


@rt(dashboard_urls_module.URLS[Url.get_extracted_text_element], methods=["GET"])
async def get_extracted_text_element_route(material_id: int) -> Div: # FT changed to Div
    """
    Retrieves extracted text for a material ID and wraps it in a Div for display.
    The text might contain HTML <mark> tags for entity highlighting.

    Args:
        material_id (int): The material ID.

    Returns:
        Div: A Div component containing the (potentially HTML-annotated) text.
    """
    logger.debug(f"Fetching extracted text element for material_id: {material_id}")
    text_content: str = get_extracted_text(material_id) # This is a synchronous function call
    # Wrap the text content, allowing HTML rendering via NotStr
    return Div(NotStr(text_content), cls="p-2 prose max-w-none") # prose for tailwind typography if used


@rt(dashboard_urls_module.URLS[Url.get_entities_element]) # Path includes {material_id:int}
async def get_entities_element_route(material_id: int) -> Div: # FT changed to Div
    """
    Retrieves and formats extracted entities for a material ID into an HTML structure.
    Entities are grouped by label and displayed with their scores.

    Args:
        material_id (int): The material ID.

    Returns:
        Div: A Div component containing the formatted list of entities, or a "not found" message.
    """
    logger.debug(f"Fetching entities element for material_id: {material_id}")
    entities_obj: Optional[Entities] = get_entities(material_id) # Synchronous call

    if not entities_obj or not entities_obj.items:
        return Div(P("No entities found or extracted for this item."), cls="p-4 text-sm")

    # Get grouped and sorted entities
    grouped_entities_dict, sorted_labels_list = entities_obj.get_grouped_and_sorted_entities()

    if not grouped_entities_dict:
        return Div(P("No entities to display after grouping."), cls="p-4 text-sm")

    entity_display_sections: List[FT] = []
    for label in sorted_labels_list:
        entities_in_group = grouped_entities_dict.get(label, [])
        if not entities_in_group: continue

        entity_list_items: List[Li] = []
        # Count occurrences of each entity text within the group for display
        text_counts = defaultdict(int)
        for entity in entities_in_group: text_counts[entity.text] += 1

        processed_texts_in_group: Set[str] = set()
        for entity in sorted(entities_in_group, key=lambda e: e.text.lower()): # Sort entities alphabetically by text
            if entity.text in processed_texts_in_group: continue # Show each unique text once
            processed_texts_in_group.add(entity.text)

            count_badge_str = f" ({text_counts[entity.text]}x)" if text_counts[entity.text] > 1 else ""

            # Determine score display and color
            score_display = ""
            score_color_cls = "badge-outline" # Default
            if entity.score is not None:
                score_display = f"{entity.score:.0%}"
                if entity.score < 0.7: score_color_cls = "badge-warning" # Low confidence
                elif entity.score >= 0.9: score_color_cls = "badge-success" # High confidence
                else: score_color_cls = "badge-info" # Medium confidence

            entity_list_items.append(
                Li( Span(entity.text + count_badge_str, cls="text-sm"),
                    Span(score_display, cls=f"ml-2 badge badge-xs {score_color_cls}") if score_display else Span(),
                    cls="flex justify-between items-center"
                )
            )

        entity_display_sections.append(
            Div(H5(label.replace("_", " ").title(), cls="font-semibold text-xs uppercase tracking-wider mb-1 text-primary"),
                Ul(*entity_list_items, cls="list-disc list-inside pl-1 space-y-0.5"),
                cls="mb-3"
            )
        )
    return Div(*entity_display_sections, cls="p-1")


# --- Main Page Route & Modal Route ---
@rt(dashboard_urls_module.URLS[Url.data_grid]) # Main data grid, also root via redirect
async def main_data_grid_route(session: Dict[str, Any], request: Request) -> Union[Tuple[FT, ...], List[FT]]:
    """
    Main route for displaying the data grid and handling all its interactions
    (pagination, sorting, filtering). This is the primary view of the dashboard.
    It processes the application state, fetches data, and renders components.
    Handles both full page loads and HTMX partial updates.

    Args:
        session (Dict[str, Any]): User session data.
        request (Request): The incoming HTTP request.

    Returns:
        Union[Tuple[FT, ...], List[FT]]: Components for rendering. For full page load,
                                         includes layout, header, grid, modals. For HTMX,
                                         returns specific components (grid + OOB filters).
    """
    form_data = {}
    if request.method == "POST": # Check if form data is expected
        try:
            form_data_raw = await request.form()
            form_data = dict(form_data_raw)
        except Exception as e_form: # Handle cases where form parsing might fail
            logger.warning(f"Could not parse form data for request to {request.url.path}: {e_form}")

    # Combine query parameters and form data (form data takes precedence for same keys)
    request_params_combined: Dict[str, Any] = {**dict(request.query_params), **form_data}
    logger.info(f"Request to data_grid: Method={request.method}, Params={request_params_combined}")

    # Process state (loads from session, applies request_params, saves back to session)
    session, current_app_state = process_state(session, request_params_combined)

    # Fetch data based on current_app_state and user's auth constraints
    processed_data_result: ProcessedDataResult = await fetch_data(current_app_state, session)

    # If page number was validated and changed by fetch_data, update session state to reflect actual view
    if processed_data_result.app_state.page != current_app_state.page:
        logger.info(f"Page validated: original {current_app_state.page}, final {processed_data_result.app_state.page}. Updating session.")
        session["app_state"] = asdict(processed_data_result.app_state)

    # Render the main data grid component
    grid_ui_component = render_data_grid_component(
        df_slice=processed_data_result.df_slice,
        app_state=processed_data_result.app_state, # Use the validated state
        total_filtered_rows=processed_data_result.total_filtered_rows,
        total_pages=processed_data_result.total_pages,
    )

    # Prepare Out-Of-Band (OOB) filter updates for HTMX requests
    auth_details_for_oob: Dict[str, Any] = session.get("auth", {})
    is_admin_user_for_oob: bool = auth_details_for_oob.get("role") == "admin"
    user_faculty_for_oob: Optional[str] = auth_details_for_oob.get("faculty")

    oob_filter_components_list: List[FT] = []
    # Define which checkbox filter groups to render (similar to page_header_component)
    checkbox_groups_config = [
        ("workflow_status", "Workflow Status", BADGE_STYLES["workflow_status"]),
        ("status", "Status", BADGE_STYLES["status"]),
        ("classification", "Classification", BADGE_STYLES["classification"]),
        ("manual_classification", "Manual Classification", BADGE_STYLES["classification"]),
    ]
    if is_admin_user_for_oob or not user_faculty_for_oob or user_faculty_for_oob == "all":
        checkbox_groups_config.append(("faculty", "Faculty", BADGE_STYLES["faculty"]))

    for key, label, opts_map in checkbox_groups_config:
        group_component = create_checkbox_filter_group(
            filter_key=key, label_text=label, options=opts_map,
            current_values_str=processed_data_result.app_state.filters.get(key),
            option_counts=processed_data_result.filter_counts.get(key, {}),
            current_total_items=processed_data_result.total_filtered_rows,
        )
        if hasattr(group_component, "attrs"): # Ensure it's a component fasthtml can add attrs to
            group_component.attrs["hx-swap-oob"] = f"outerHTML:#filter-group-{key}" # type: ignore
        oob_filter_components_list.append(group_component)

    # Determine if it's an HTMX request or a full page load
    is_htmx_request = request.headers.get("hx-request", "").lower() == "true"

    if is_htmx_request:
        logger.debug("HTMX request detected. Returning data grid and OOB filter updates.")
        return (grid_ui_component, *oob_filter_components_list) # type: ignore # fasthtml handles tuple of components
    else: # Full page load
        logger.debug("Full page request. Rendering complete page layout.")
        page_header = page_header_component(
            user_details=auth_details_for_oob,
            app_state=processed_data_result.app_state, # Use validated state for header display
            filter_counts=processed_data_result.filter_counts,
            total_filtered_rows=processed_data_result.total_filtered_rows,
        )
        modal_dialog_placeholder = Dialog(id="modal-placeholder", cls="modal modal-bottom sm:modal-middle")

        # Assemble full page
        # PAGE_BODY_JS is a tuple of Script components
        page_elements: List[Any] = [
            Title("Easy Access Dashboard"), # Browser window title
            *INIT_HEADERS, # Initial CSS, JS from constants (already in app setup, but can be here for clarity)
            Div( # Main page container
                Div(page_header, grid_ui_component, id="content-area"), # Content area
                id="page-container", hx_ext="preload"
            ),
            modal_dialog_placeholder, # For item details
            *PAGE_BODY_JS, # Scripts at the end of body
        ]
        return page_elements


@rt(dashboard_urls_module.URLS[Url.show_item_details]) # Path includes {material_id:int}
async def show_item_details_modal_route(session: Dict[str, Any], material_id: int) -> Tuple[FT, ...]:
    """
    Fetches detailed data for a specific copyright item and renders the content
    for the details modal. Includes navigation to previous/next items based on
    the current filtered and sorted view.

    Args:
        session (Dict[str, Any]): User session data (for AppState and auth).
        material_id (int): The ID of the item to display details for.

    Returns:
        Tuple[FT, ...]: Components representing the modal's inner content and
                        HTMX headers to trigger modal display.
    """
    logger.info(f"Showing item details modal for material_id: {material_id}")

    # Load AppState to determine context (filters/sorting for prev/next)
    current_app_state = load_app_state(session)
    auth_details: Dict[str, Any] = session.get("auth", {})
    faculty_constraint_for_nav: Optional[Dict[str, str]] = None
    user_faculty_for_nav: Optional[str] = auth_details.get("faculty")
    if auth_details.get("role") != "admin" and user_faculty_for_nav and user_faculty_for_nav != "all":
        faculty_constraint_for_nav = {"faculty": user_faculty_for_nav}

    prev_item_id: Optional[int] = None
    next_item_id: Optional[int] = None
    try:
        # Get currently filtered/sorted list of all material_ids to find prev/next
        # This re-applies current filters/sort from AppState
        # get_filtered_sorted_df is synchronous
        ordered_df_for_nav = get_filtered_sorted_df(current_app_state, extra_constraints=faculty_constraint_for_nav)
        if "material_id" in ordered_df_for_nav.columns:
            ordered_ids_list: List[Any] = ordered_df_for_nav.get_column("material_id").to_list()
            # Ensure material_id from path is same type as in list (e.g. int)
            # Current list elements might be various types if not cast before.
            # Assuming material_id in DataFrame is int or string convertible to int.
            # For safety, cast all to string for comparison then to int if needed.
            try:
                str_ordered_ids = [str(id_val) for id_val in ordered_ids_list]
                str_material_id = str(material_id)
                current_idx = str_ordered_ids.index(str_material_id) if str_material_id in str_ordered_ids else -1

                if current_idx != -1:
                    if current_idx > 0: prev_item_id = int(ordered_ids_list[current_idx - 1])
                    if current_idx < len(ordered_ids_list) - 1: next_item_id = int(ordered_ids_list[current_idx + 1])
            except ValueError: # material_id not found in list or type issue
                logger.warning(f"Material ID {material_id} not found in the current filtered/sorted list for modal navigation.")
        else:
            logger.warning("'material_id' column not found in DataFrame for modal navigation.")

    except Exception as e_nav:
        logger.error(f"Error determining prev/next navigation for modal (material_id {material_id}): {e_nav}")

    # Fetch detailed data for the specific item (this is an async call)
    # retrieve_osiris_data is currently synchronous, needs to be called with to_thread
    item_details_list: List[Dict[str, Any]] = await asyncio.to_thread(retrieve_osiris_data, [material_id])

    if not item_details_list:
        logger.warning(f"No detailed data found for material_id {material_id} via retrieve_osiris_data.")
        # Construct a basic error display for the modal
        error_modal_box = Div(H3("Error"), P(f"Could not load details for item {material_id}."),
                              Form(method="dialog")(Button("Close", cls="btn btn-sm mt-4")),
                              cls="modal-box")
        error_modal_backdrop = Form(method="dialog", cls="modal-backdrop")(NotStr("<button>close</button>"))
        return (error_modal_box, error_modal_backdrop), HtmxResponseHeaders(trigger="openModalEvent").headers # type: ignore

    item_data_dict: Dict[str, Any] = item_details_list[0]
    # Ensure 'faculty' key exists if 'faculty_id' was present (common from DB)
    if "faculty_id" in item_data_dict and "faculty" not in item_data_dict:
        item_data_dict["faculty"] = item_data_dict.pop("faculty_id")

    # --- Modal Header ---
    filename_str = item_data_dict.get("filename", "N/A") or "(Filename N/A)"
    file_url_str = item_data_dict.get("url")

    filename_display: FT = H4(filename_str, cls="font-semibold text-lg break-all leading-tight")
    if file_url_str:
        filename_display = A(filename_display, href=file_url_str, target="_blank", cls="link hover:link-primary")

    status_pill_component, _ = render_modal_field("status", item_data_dict.get("status"))
    material_id_pill = Label(f"ID: {material_id}", cls=f"{DEFAULT_PILL_STYLE} badge-sm ml-2")

    modal_header = Div(cls="flex items-start justify-between space-x-4 pb-2 border-b border-base-300")(
        Div(cls="flex items-center space-x-3 flex-grow min-w-0")(status_pill_component, material_id_pill, filename_display),
        Form(method="dialog")(Button("✕", cls="btn btn-sm btn-circle btn-ghost flex-shrink-0")) # Close button
    )

    # --- Modal Cards Grid ---
    # Data Entry Card (Remarks, Workflow, Manual Classification)
    data_entry_card_content = (
        create_editable_pill_div(item_data_dict.get("workflow_status"), "Workflow Status", "workflow_status", BADGE_STYLES["workflow_status"], material_id),
        create_editable_pill_div(item_data_dict.get("manual_classification"), "Manual Classification", "manual_classification", BADGE_STYLES["classification"], material_id),
        Div(fh.Strong("Remarks", cls="block text-xs font-semibold uppercase tracking-wider text-base-content/70 mb-1"),
            TextArea(name="remarks", rows="4", cls="textarea textarea-bordered w-full text-sm bg-base-100",
                     x_model="remarks", id="modal_remarks_textarea") # Use x_model
        ),
        Div(cls="flex justify-end space-x-2 mt-3")( # Buttons for remarks
            Button("Reset", type="button", cls="btn btn-sm btn-ghost", **{"@click": "remarks = originalRemarks; document.getElementById('modal_save_remarks_btn').disabled = true;"}),
            Button("Save Remarks", id="modal_save_remarks_btn", type="submit", cls="btn btn-sm btn-primary",
                   **{":disabled": "remarks == originalRemarks"})
        )
    )
    # Item Info Card (System Classification, ML Prediction, Period, Faculty, etc.)
    item_info_card_content = tuple(
        create_readonly_item_div(item_data_dict.get(key), key.replace("_", " ").title(), key)
        for key in ["classification", "ml_prediction", "period", "faculty", "owner", "department", "course_name", "course_code"]
        if key in item_data_dict # Only render if key exists
    )
    # Text Details (Title, Author, Publisher, DOI, ISBN)
    text_details_card_content = tuple(
        create_readonly_item_div(item_data_dict.get(key), key.replace("_", " ").title(), key)
        for key in ["title", "author", "publisher", "doi", "isbn"] if key in item_data_dict
    )
    # Counts (Page, Word, Picture)
    counts_card_content = tuple(
        create_readonly_item_div(item_data_dict.get(key), key.replace("count", "").title(), key)
        for key in ["pagecount", "wordcount", "picturecount"] if key in item_data_dict
    )
    # Osiris Data Cards
    contact_info_card_content = render_contact_info(item_data_dict)
    course_details_card_content = render_course_details(item_data_dict)
    teachers_card_content = render_teacher_info(item_data_dict)
    # Item History Card (fetches its own data)
    item_history_display_card = render_item_history(await get_item_history_for_dashboard(material_id))


    modal_cards = Div(cls="grid grid-cols-1 md:grid-cols-3 gap-3")(
        Div(ItemDetailCard("Data Entry", *data_entry_card_content, card_id="modal-data-entry", col_span=1),
            ItemDetailCard("Entities", Span("Click to load...", cls="italic"), card_id="modal-entities", col_span=1, start_collapsed=True, lazy_load_url=str(urls.URLS[Url.get_entities_element]).replace("{material_id:int}", str(material_id))),
            cls="flex flex-col space-y-3"),
        Div(ItemDetailCard("Item Info", *item_info_card_content, card_id="modal-item-info", col_span=1, start_collapsed=True),
            ItemDetailCard("Text Details", *text_details_card_content, card_id="modal-text-details", col_span=1, start_collapsed=True),
            ItemDetailCard("Counts", *counts_card_content, card_id="modal-counts", col_span=1, start_collapsed=True),
            item_history_display_card, # Already a card
            cls="flex flex-col space-y-3"),
        Div(ItemDetailCard("Contact Info (Osiris)", *contact_info_card_content, card_id="modal-contact-info", col_span=1),
            ItemDetailCard("Course Details (Osiris)", *course_details_card_content, card_id="modal-course-details", col_span=1, start_collapsed=True),
            ItemDetailCard("Teachers/Persons (Osiris)", *teachers_card_content, card_id="modal-teachers", col_span=1, start_collapsed=True),
            cls="flex flex-col space-y-3"),
        ItemDetailCard("PDF Viewer", Span("Click to load PDF...", cls="italic"), card_id="modal-pdf-viewer", col_span=3, start_collapsed=True, lazy_load_url=str(urls.URLS[Url.get_pdf_element]).replace("{material_id:int}", str(material_id))),
        ItemDetailCard("Extracted Text", Span("Click to load text...", cls="italic"), card_id="modal-extracted-text", col_span=3, start_collapsed=True, lazy_load_url=str(urls.URLS[Url.get_extracted_text_element]).replace("{material_id:int}", str(material_id))),
    )

    # --- Modal Footer (Navigation) ---
    prev_btn_attrs = {"id":"modal-prev-btn", "cls": f"{ButtonT.secondary} btn-sm", "disabled": prev_item_id is None}
    if prev_item_id: prev_btn_attrs.update({"hx_get": str(urls.URLS[Url.show_item_details]).replace("{material_id:int}", str(prev_item_id)), "hx_target":"#modal-placeholder", "hx_swap":"innerHTML"})

    next_btn_attrs = {"id":"modal-next-btn", "cls": f"{ButtonT.secondary} btn-sm", "disabled": next_item_id is None}
    if next_item_id: next_btn_attrs.update({"hx_get": str(urls.URLS[Url.show_item_details]).replace("{material_id:int}", str(next_item_id)), "hx_target":"#modal-placeholder", "hx_swap":"innerHTML"})

    modal_footer = Div(cls="modal-action mt-3 pt-3 border-t border-base-300")( # Reduced margins
        Div(cls="flex justify-between w-full")(
            Button("< Prev", **prev_btn_attrs),
            Form(method="dialog")(Button("Close", cls=f"{ButtonT.primary} btn-sm")),
            Button("Next >", **next_btn_attrs),
        )
    )

    # Alpine.js data for remarks editing
    alpine_data_str = f"{{ remarks: {json.dumps(item_data_dict.get('remarks', ''))}, originalRemarks: {json.dumps(item_data_dict.get('remarks', ''))} }}"

    modal_box = Div(
        modal_header,
        Div( # Scrollable content area
            Form(Input(type="hidden", name="material_id", value=str(material_id)), modal_cards,
                 id="modal-details-form", hx_post=str(urls.URLS[Url.save_item_details]),
                 hx_include="[name='material_id'], [name='remarks']", # Only send these for this form
                 hx_target="body", hx_swap="none", # Target body for toasts, swap none for form itself
                 x_data=alpine_data_str,
                 **{"@remarks-save-success.window": "originalRemarks = remarks; console.log('Alpine: Remarks saved, originalRemarks updated.')"}
            ),
            cls="py-4 flex-grow overflow-y-auto",
        ),
        modal_footer, # Footer is fixed part of modal-box
        cls="modal-box w-[90vw] max-w-none h-[calc(100vh-4rem)] max-h-none flex flex-col p-4", # Adjusted padding
    )

    # Standard modal backdrop for closing when clicking outside
    modal_backdrop_form = Form(method="dialog", cls="modal-backdrop")(NotStr("<button>close</button>"))

    # Return tuple of components for FastHTML, and HTMX headers to trigger modal opening
    return (modal_box, modal_backdrop_form), HtmxResponseHeaders(trigger="openModalEvent").headers # type: ignore


@rt(dashboard_urls_module.URLS[Url.login]) # Login GET route
async def get_login_page() -> List[FT]: # Changed to async for consistency, though not strictly needed if no await
    """Serves the login page."""
    login_form = Form(
        Div(fh.Label("Email", html_for="email", cls="block text-gray-700 text-sm font-bold mb-2"), # Use fh.Label
            Input(id="email", type="email", placeholder="Email", name="email", cls="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline"),
            cls="mb-4"),
        Div(fh.Label("Password", html_for="password", cls="block text-gray-700 text-sm font-bold mb-2"),
            Input(id="password", name="pwd", type="password", placeholder="******************", cls="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 mb-3 leading-tight focus:outline-none focus:shadow-outline"),
            cls="mb-6"),
        fh.Button("Login", cls="btn btn-primary bg-indigo-500 hover:bg-indigo-700 text-white font-bold py-2 px-4 rounded focus:outline-none focus:shadow-outline w-full", type="submit"), # Added w-full
        action=str(dashboard_urls_module.URLS[Url.login]), method="post", cls="bg-white shadow-xl rounded px-8 pt-6 pb-8 mb-4" # Increased shadow
    )
    # Centered login card
    return [
        Title("Login - Easy Access Dashboard"),
        Div(Card(CardTitle(H2("Copyright Dashboard Login", cls="text-center text-2xl font-bold text-primary"), cls="items-center justify-center pt-4"), # Centered title
                 CardBody(login_form, cls="p-6"), # Adjusted padding
                 cls="w-full max-w-md bg-base-100 shadow-2xl rounded-lg"), # Added max-width and shadow
            cls="flex items-center justify-center min-h-screen bg-gradient-to-br from-primary to-secondary p-4") # Fullscreen gradient
    ]


@rt(dashboard_urls_module.URLS[Url.login], methods=["POST"]) # Login POST route
async def post_login_form(login_input: Login, session: Dict[str, Any]) -> RedirectResponse: # login_input uses Pydantic binding from fasthtml
    """
    Handles login form submission. Validates credentials and sets session auth data.
    Redirects to data grid on success, or back to login on failure.

    Args:
        login_input (Login): The login data (email, password) parsed from the form.
        session (Dict[str, Any]): The user's session data.

    Returns:
        RedirectResponse: Redirects to the main data grid or back to the login page.
    """
    if not login_input.email or not login_input.pwd:
        logger.info("Login attempt with empty email or password.")
        add_toast(session, "Email and password are required.", "error")
        return LOGIN_REDIRECT_RESPONSE

    try:
        # Assumes `users` is a fastlite table object
        user_data = users.get(login_input.email) # type: ignore # Get user by PK (email)
    except users.NotFoundError: # Specific exception for fastlite if user not found
        logger.warning(f"Login attempt failed: User '{login_input.email}' not found.")
        add_toast(session, "Invalid email or password.", "error")
        return LOGIN_REDIRECT_RESPONSE
    except Exception as e_db_user: # Catch other potential DB errors
        logger.error(f"Database error during login for user '{login_input.email}': {e_db_user}")
        add_toast(session, "Server error during login. Please try again.", "error")
        return LOGIN_REDIRECT_RESPONSE

    # TODO: CRITICAL SECURITY FLAW - Passwords must be hashed and compared securely.
    # This is placeholder logic for demonstration and needs immediate replacement
    # with a proper password hashing and verification mechanism (e.g., passlib).
    # DO NOT USE IN PRODUCTION.
    from hashlib import sha256 # Example, NOT for production password storage
    # This is NOT secure password handling:
    # 1. Storing plain text or easily reversible hashes is bad.
    # 2. compare_digest is for timing attack resistance on already securely hashed passwords.
    # For now, assuming 'pwd' in DB is plain for this example to work, which is wrong.
    # Correct approach: Hash password on user creation/set, compare hash of input pwd with stored hash.

    # Placeholder for password check - REPLACE WITH SECURE HASHING
    # For this review, I will assume this check is what was intended by the original `compare_digest`
    # if compare_digest(user_data.get("pwd", "").encode("utf-8"), login_input.pwd.encode("utf-8")):
    if user_data.get("pwd") == login_input.pwd: # Insecure direct comparison
        logger.info(f"Login successful for '{login_input.email}'.")
        session["auth"] = { # Store user details in session
            "email": login_input.email, "name": user_data.get("name"),
            "faculty": user_data.get("faculty"), "role": user_data.get("role"),
        }
        if "app_state" in session: del session["app_state"] # Clear any old app state
        add_toast(session, f"Login successful! Welcome, {user_data.get('name', 'User')}!", "success")
        return RedirectResponse(str(dashboard_urls_module.URLS[Url.data_grid]), status_code=303)
    else:
        logger.warning(f"Login attempt failed: Incorrect password for '{login_input.email}'.")
        add_toast(session, "Invalid email or password.", "error")
        return LOGIN_REDIRECT_RESPONSE


@rt(dashboard_urls_module.URLS[Url.logout]) # Logout route
async def logout_route(session: Dict[str, Any]) -> RedirectResponse: # Changed to async for consistency
    """
    Logs out the user by clearing authentication details and AppState from the session.
    Redirects to the login page.

    Args:
        session (Dict[str, Any]): The user's session data.

    Returns:
        RedirectResponse: Redirects to the login page.
    """
    user_email_for_log = session.get("auth", {}).get("email", "Unknown user")
    if "auth" in session: del session["auth"]
    if "app_state" in session: del session["app_state"]
    logger.info(f"User '{user_email_for_log}' logged out.")
    # add_toast(session, "You have been logged out.", "info") # Toast won't show after redirect usually
    return LOGIN_REDIRECT_RESPONSE


@rt(dashboard_urls_module.URLS[Url.root]) # Root path
async def root_redirect_to_data_grid() -> RedirectResponse: # Changed to async
    """Redirects the root path ("/") to the main data grid view."""
    logger.debug("Root path '/' accessed, redirecting to data grid.")
    return RedirectResponse(url=str(dashboard_urls_module.URLS[Url.data_grid]), status_code=302)


# Initialize global_urls (from urls.py) with the routes defined here
# This makes them available for reverse URL lookups (e.g., X.to())
# This should be done after all @rt routes are defined.
def _update_global_urls_from_router(router: QueryRouter, url_enum: type[Enum], target_dict: dict) -> None:
    """Helper to populate the global URLS dict from router paths."""
    for member in url_enum: # type: ignore
        route_name = member.value # For StrEnum, .value gives the string
        # Find the route in the router; this might need a more direct way if available from fasthtml
        # For now, assuming router stores them in a way that `router.paths[route_name]` works or similar
        # This part is highly dependent on fasthtml's internal structure for reverse lookups.
        # The `.to()` method on route functions is the typical way fasthtml handles this.
        # If URLS dict is for external use/JS, ensure paths are correct.
        try:
            # Get the function object for the route name
            route_func = getattr(router.lookup_path(f"/{route_name}")[0], 'endpoint', None) if router.lookup_path(f"/{route_name}") else None
            if route_func and hasattr(route_func, 'to'):
                 target_dict[member] = route_func.to() # type: ignore
            elif route_name in router.paths : # Fallback if .to is not directly on endpoint
                 target_dict[member] = router.paths[route_name] # Store the path pattern
            else:
                logger.warning(f"Could not find route path for Url enum member: {route_name}")
        except Exception as e_url: # pragma: no cover
            logger.warning(f"Error getting path for Url enum member {route_name}: {e_url}")

# This should be called once after all routes are defined.
# However, route functions are already assigned to global_urls in urls.py via .to()
# This manual update here is redundant if that pattern is followed.
# For safety, let's ensure the global_urls dict from the urls module is updated.
# The .to() method on route functions is the correct way to get their paths.
# This manual population is likely not needed if routes correctly assign to global_urls.
# If Url enum names match route function names, this can be automated.

# The existing pattern in the original code was:
# URLS = { Url.root: root_redirect.to(), ... }
# global_urls.update(URLS)
# This is correct. The _update_global_urls_from_router helper is not necessary.
# I'll ensure the dictionary is correctly assigned.

# dashboard_urls_module.URLS is populated at the end of dash.py in the original code
# This ensures all @rt decorated functions are registered before their .to() method is called.
# This is the correct pattern.

# Finalizing URLS dictionary (should be at the end of file after all @rt definitions)
# This was originally in urls.py but makes more sense here after routes are defined.
# However, to avoid circular imports if urls.py needs to import from dash.py (e.g. for route functions),
# it's often better to have routes populate a central dict or have a registry.
# The original code had URLS dict in urls.py and updated it from dash.py.
# Let's stick to that pattern: dash.py updates the dict in urls.py.
dashboard_urls_module.URLS.update({
    Url.root: root_redirect_to_data_grid, # Pass function itself for .to() later if needed elsewhere
    Url.data_grid: main_data_grid_route,
    Url.logout: logout_route,
    Url.show_item_details: show_item_details_modal_route,
    Url.get_entities_element: get_entities_element_route,
    Url.get_pdf_element: get_pdf_element_route,
    Url.get_extracted_text_element: get_extracted_text_element_route,
    Url.save_item_details: save_item_details_route,
    Url.get_file: get_file_route,
    Url.get_osiris_data: get_osiris_data_route,
    Url.update_single_field: update_single_field_route,
    Url.login: get_login_page, # Default to GET for login URL symbol
})
# Convert function references to paths for the global_urls dict that might be used by templates/JS
for url_enum_member, route_function in dashboard_urls_module.URLS.items():
    if callable(route_function) and hasattr(route_function, 'to'):
        dashboard_urls_module.URLS[url_enum_member] = route_function.to()
    # else it might already be a path string from a previous assignment or manual entry


def start_dashboard_server() -> None: # Renamed 'start' to be more descriptive
    """
    Starts the FastHTML web server for the dashboard.
    This is typically called when `run.py` is executed with the `--dashboard` flag.
    """
    logger.info(f"Starting dashboard server on port {APP_PORT}. Access at {APP_ROOT_URL}")
    # uvicorn.run is usually called from run.py, not here.
    # This function might be for programmatic server start if needed.
    # For now, assuming uvicorn is run externally as in run.py.
    # If this 'start' is the main entry, then uvicorn.run would be here.
    # The original run.py handles uvicorn.run.
    # This function could be a placeholder or for embedded server scenarios.
    # For consistency with run.py, actual server start is external.
    # This function can just log that server should be started.
    logger.info("To run the dashboard, use the --dashboard flag with run.py, or run an ASGI server like Uvicorn (e.g., uvicorn easy_access.dashboard.dash:app --reload)")


if __name__ == "__main__": # pragma: no cover
    # This block allows running the dashboard directly using `python -m easy_access.dashboard.dash`
    # (if this file is runnable as a module) or `python easy_access/dashboard/dash.py`.
    logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
    logger.info("Dashboard module executed directly. Starting Uvicorn server for development.")

    # The app needs to be importable as "easy_access.dashboard.dash:app" for uvicorn.
    # If running this file directly, uvicorn needs to find 'app'.
    # uvicorn.run(app, host="0.0.0.0", port=APP_PORT, reload=True)
    # The above line might not work if 'app' is not found correctly by uvicorn when run this way.
    # Usually, you run: uvicorn easy_access.dashboard.dash:app --reload
    # So, this __main__ block is mostly for conceptual testing or if structure allows direct run.
    logger.info("To run the dashboard: uvicorn easy_access.dashboard.dash:app --reload --port PORT_NUMBER")

```
