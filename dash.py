# dash.py

import contextlib
import datetime
import json
import math
import traceback
from collections import defaultdict
from copy import deepcopy
from dataclasses import dataclass, field
from enum import auto
from pathlib import Path
from typing import Any

import fasthtml.common as fh
import polars as pl
from fastcore.utils import *
from fasthtml.common import *
from fasthtml.components import Button
from fastlite import database
from monsterui.all import *
from monsterui.foundations import VEnum, str2ukcls
from rich import print
from starlette.staticfiles import StaticFiles

from easy_access.db.models import ItemUpdate
from easy_access.db.retrieve import (
    retrieve_copyright_items,
    retrieve_item_history,
    retrieve_osiris_data,
)
from easy_access.db.update import update_copyright_items
from easy_access.settings import SETTINGS, DirSetting

# --- monsterui fixes ---


# NOTE: do a manual replace of this function in the monsterui library in order to fix theme switching
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


class LabelT(VEnum):
    """FIX: replaces danger with destructive"""

    def _generate_next_value_(name, start, count, last_values):
        return str2ukcls("label", name)

    primary = auto()
    secondary = auto()
    destructive = auto()


# --- globals / constants ---

WORKFLOW_STYLES: dict[str, LabelT] = {
    "ToDo": LabelT.destructive,
    "InProgress": LabelT.secondary,
    "Done": LabelT.primary,
}
STATUS_STYLES: dict[str, LabelT] = {
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
DEFAULT_PILL_STYLE = LabelT.secondary
FILTERABLE_COLUMNS: list[str] = [
    "workflow_status",
    "classification",
    "status",
    "manual_classification",
    "faculty",
    "department",
    "course_name",
]
DISPLAY_COLUMNS: list[str] = [
    "material_id",
    "url",
    "workflow_status",
    "status",
    "classification",
    "ml_prediction",
    "manual_classification",
    "remarks",
    "filename",
    "title",
    "faculty",
]
# --- Filter Toggle Definitions ---
WORKFLOW_FILTER_STATES = [
    None,
    "InProgress|ToDo",
    "InProgress",
    "ToDo",
]
WORKFLOW_FILTER_LABELS = {
    None: "Workflow: All",
    "InProgress|ToDo": "Workflow: Not Done",
    "InProgress": "Workflow: In Progress",
    "ToDo": "Workflow: ToDo",
}

STATUS_FILTER_STATES = [None, "Published|Unpublished", "Published"]
STATUS_FILTER_LABELS = {
    None: "Status: All",
    "Published|Unpublished": "Status: Not Deleted",
    "Published": "Status: Published",
}

# Using specific values from the DESTRUCTIVE/SECONDARY sets
CLASSIFICATION_FILTER_STATES = [None, "lange overname", "onbekend"]
CLASSIFICATION_FILTER_LABELS = {
    None: "Class: All",
    "lange overname": "Class: Lange Overname",
    "onbekend": "Class: Onbekend",
}


copyright_df_global: pl.DataFrame = retrieve_copyright_items()


db = database("db.sqlite3")

items = db.t.items

PORT = 8000
ROOT_URL = f"http://localhost:{PORT}"
DEFAULT_PER_PAGE = 15
MAX_CELL_LENGTH = 35


@dataclass
class AppState:
    """
    Data structure to hold the UI state of the application (pagination, sorting, filters).
    State is loaded from the session, modified based on user requests, and saved back.
    """

    page: int = 1
    per_page: int = DEFAULT_PER_PAGE
    sort_by: str | None = None
    sort_desc: bool = False
    filters: dict[str, str] = field(default_factory=dict)
    # current_total removed - calculate dynamically

    def set_page(self, page: int):
        """Sets the current page number."""
        self.page = max(1, page)  # Basic validation

    def set_per_page(self, per_page: int):
        """Sets the number of items per page."""
        self.per_page = max(1, per_page)  # Basic validation

    def set_sort(self, sort_by: str | None, sort_desc: bool):
        """Sets the sorting column and direction."""
        # Optional: Validate sort_by against available columns?
        self.sort_by = sort_by
        self.sort_desc = sort_desc

    def apply_filter_change(self, key: str, value: str, is_checked: bool) -> bool:
        """
        Applies a change from a checkbox filter.

        Updates the pipe-separated string for the given filter key.

        Args:
            key: The filter key (e.g., 'workflow_status').
            value: The specific value being toggled (e.g., 'ToDo').
            is_checked: The *intended* state of the checkbox (True if checked).

        Returns:
            True if the filters for this key actually changed, False otherwise.
        """
        current_values = (
            set(self.filters.get(key, "").split("|"))
            if self.filters.get(key)
            else set()
        )
        original_filter_str = self.filters.get(key)

        if is_checked:
            current_values.add(value)
        else:
            current_values.discard(value)  # Use discard to avoid error if not present

        # Filter out empty strings that might result from split('')
        current_values.discard("")

        if not current_values:
            if key in self.filters:
                del self.filters[key]
                return True  # Changed from something to nothing
            else:
                return False  # Was already nothing
        else:
            new_filter_str = "|".join(sorted(list(current_values)))
            if new_filter_str != original_filter_str:
                self.filters[key] = new_filter_str
                return True  # Changed
            else:
                return False  # No effective change

    def apply_text_filter(self, key: str, value: str) -> bool:
        """
        Sets or removes a text-based filter value.

        Args:
            key: The filter key (e.g., 'department').
            value: The value from the text input. Empty string removes the filter.

        Returns:
            True if the filter value changed, False otherwise.
        """
        original_value = self.filters.get(key)
        if value:
            if value != original_value:
                self.filters[key] = value
                return True
            else:
                return False
        else:  # Empty value means remove the filter
            if key in self.filters:
                del self.filters[key]
                return True
            else:
                return False

    def clear_filters(self):
        """Clears all user-defined filters."""
        self.filters = {}
        # Note: Does not handle auth-based filters, they are applied dynamically later.

    # --- Read Methods (mostly unchanged) ---

    def nav_params_dict(self) -> dict[str, Any]:
        """Returns dictionary of navigation parameters (page, sort, per_page)."""
        return {
            "page": self.page,
            "per_page": self.per_page,
            "sort_by": self.sort_by or "",  # Ensure string for URL gen
            "sort_desc": str(self.sort_desc),
        }

    def filter_params_dict(self) -> dict[str, str]:
        """Returns dictionary of currently active filters, prefixed for forms/URLs."""
        if not self.filters:
            return {}
        return {f"filter_{k}": v for k, v in self.filters.items()}

    def all_params_dict(self) -> dict[str, Any]:
        """Returns combined navigation and filter parameters."""
        return {**self.nav_params_dict(), **self.filter_params_dict()}

    def get_active_filters(self) -> dict[str, str]:
        """Returns a copy of the currently active user filters."""
        # Return a copy to prevent external modification
        return self.filters.copy()


def load_app_state(session: dict) -> AppState:
    """
    Loads the AppState object from the session dictionary.

    If 'app_state' is not found in the session or is invalid,
    it returns a default AppState instance.

    Args:
        session: The session dictionary.

    Returns:
        An AppState instance.
    """
    app_state_dict = session.get("app_state", {})
    try:
        # Attempt to create AppState; might fail if dict structure is wrong
        state = AppState(**app_state_dict)
    except (TypeError, ValueError):
        print("Warning: Invalid or missing app_state in session. Using default.")
        state = AppState()  # Return default state
    return state


@dataclass
class ProcessedDataResult:
    """
    Holds the results after processing state and fetching data.
    """

    app_state: AppState  # The validated state reflecting the data fetched
    df_slice: pl.DataFrame
    total_filtered_rows: int
    total_pages: int
    filter_counts: dict[str, dict[str, int]]  # Counts for UI elements


def process_state(session: dict, request_params: dict) -> tuple[dict, AppState]:
    """
    Loads state, applies the single requested change based on request_params,
    saves the potentially modified state back to the session immediately,
    and returns the updated session dict and the AppState object reflecting the user's request intent.

    Does not perform data-dependent validation (like page number).

    Args:
        session: The current session dictionary.
        request_params: Dictionary containing parameters from the request (query + form).

    Returns:
        A tuple containing the (potentially updated) session dictionary and the AppState instance.
    """
    print(f"Processing state with params: {request_params}")  # Debug incoming params

    # 1. Load initial state
    app_state = load_app_state(session)
    initial_state_dict = asdict(app_state)  # For change detection

    # 2. Determine action and modify state based on request_params
    filters_changed = False
    if request_params.get("action") == "reset":
        print("Processing state: Reset action")
        app_state.clear_filters()
        app_state.set_page(1)
        # Check if filters *were* present before clearing
        filters_changed = bool(initial_state_dict.get("filters"))
        if not filters_changed and app_state.page == initial_state_dict.get("page", 1):
            print("No change made -- page already default")
            pass  # No effective change from reset if already default
        else:
            filters_changed = True  # Consider page reset or filter clear a change

    elif "page" in request_params:
        try:
            page = int(request_params["page"])
            if page != app_state.page:
                print(f"Processing state: Set page to {page}")
                app_state.set_page(page)
        except (ValueError, TypeError):
            print(
                f"Warning: Invalid page parameter '{request_params['page']}'. Ignoring."
            )
    elif "sort_by" in request_params:
        sort_by = request_params.get("sort_by")
        sort_desc = request_params.get("sort_desc", "False").lower() == "true"
        if sort_by != app_state.sort_by or sort_desc != app_state.sort_desc:
            print(
                f"Processing state: Set sort to {sort_by} ({'DESC' if sort_desc else 'ASC'})"
            )
            app_state.set_sort(sort_by if sort_by else None, sort_desc)
    elif "per_page" in request_params:
        try:
            per_page = int(request_params["per_page"])
            if per_page != app_state.per_page:
                print(f"Processing state: Set per_page to {per_page}")
                app_state.set_per_page(per_page)
                app_state.set_page(1)
        except (ValueError, TypeError):
            print(
                f"Warning: Invalid per_page parameter '{request_params['per_page']}'. Ignoring."
            )
    else:  # Check for filter changes (expects only one per request now)
        is_checkbox_change = "checked" in request_params
        filter_key_from_param = None
        filter_value_from_param = None

        if is_checkbox_change:
            # For checkbox changes, find the specific 'filter_*' key that has a NON-EMPTY value.
            # This assumes hx-vals correctly sends {"filter_thekey": "thevalue", "checked": "True/False"}
            # and the form submission noise might add other empty filter_* keys.
            found_key = None
            for key, value in request_params.items():
                if key.startswith("filter_") and value:  # Find the key *with* a value
                    found_key = key
                    break

            if found_key:
                filter_key_from_param = found_key[len("filter_") :]
                filter_value_from_param = request_params[found_key]

                if filter_key_from_param in FILTERABLE_COLUMNS:
                    print(
                        f"Processing state: Checkbox change for '{filter_key_from_param}' = '{filter_value_from_param}'"
                    )
                    is_checked = request_params["checked"].lower() == "true"
                    if app_state.apply_filter_change(
                        filter_key_from_param, filter_value_from_param, is_checked
                    ):
                        filters_changed = True
                else:
                    print(
                        f"Warning: Received unknown filter key '{filter_key_from_param}' from checkbox. Ignoring."
                    )
            else:
                print(
                    "Warning: Checkbox change detected but no filter key with value found in params."
                )

        else:  # Handle text input change (assume only one relevant filter_*)
            # Find the first filter_* key (could be empty value for text clear)
            found_key = None
            for key in request_params:
                if key.startswith("filter_"):
                    found_key = key
                    break

            if found_key:
                filter_key_from_param = found_key[len("filter_") :]
                filter_value_from_param = request_params[found_key]  # Can be empty

                if filter_key_from_param in FILTERABLE_COLUMNS:
                    print(
                        f"Processing state: Text filter change for '{filter_key_from_param}' = '{filter_value_from_param}'"
                    )
                    if app_state.apply_text_filter(
                        filter_key_from_param, filter_value_from_param
                    ):
                        filters_changed = True
                else:
                    print(
                        f"Warning: Received unknown text filter key '{filter_key_from_param}'. Ignoring."
                    )
            # else: No filter_* keys found, likely not a filter action

    if filters_changed:
        print("Processing state: Filters changed, resetting page to 1")
        app_state.set_page(1)

    # Save state back to session *only if* it actually changed from the initial load
    final_state_dict = asdict(app_state)
    state_has_changed = final_state_dict != initial_state_dict
    if state_has_changed:
        print("Processing state: State changed, updating session.")
        session["app_state"] = final_state_dict
    else:
        print("Processing state: No effective state change detected.")

    return session, app_state


async def fetch_data(app_state: AppState, session: dict) -> ProcessedDataResult:
    """
    Takes an AppState reflecting user request, applies auth constraints,
    fetches/filters data, validates state (page number), calculates derived values,
    and returns results including the validated state.

    Args:
        app_state: The AppState reflecting the user's requested filters/sort/page.
        session: The session dictionary (needed for auth details).

    Returns:
        A ProcessedDataResult object.
    """
    print(f"Fetching data for state: {app_state}")
    # 1. Get Auth Details from Session
    auth_details = session.get("auth", {})

    # 2. Determine Auth Constraint
    faculty_constraint = None
    user_faculty = auth_details.get("faculty")
    user_role = auth_details.get("role")
    if user_role != "admin" and user_faculty and user_faculty != "all":
        faculty_constraint = {"faculty": user_faculty}
        print(f"Applying faculty constraint: {faculty_constraint}")

    # 3. Filter Data (using adapted get_filtered_sorted_df)
    filtered_df = get_filtered_sorted_df(
        app_state, extra_constraints=faculty_constraint
    )
    total_filtered_rows = filtered_df.height
    print(f"Filtered data rows: {total_filtered_rows}")

    # 4. Calculate Total Pages & *Validate* Page Number in a *copy* of the state
    validated_app_state = deepcopy(app_state)
    total_pages = (
        math.ceil(total_filtered_rows / validated_app_state.per_page)
        if validated_app_state.per_page > 0
        else 1
    )
    original_page_request = validated_app_state.page
    validated_app_state.page = max(
        1, min(validated_app_state.page, total_pages if total_pages > 0 else 1)
    )
    if validated_app_state.page != original_page_request:
        print(
            f"Page number validated: Requested {original_page_request}, Validated {validated_app_state.page} (Total Pages: {total_pages})"
        )

    # 5. Calculate Filter Counts (using adapted calculate_counts_for_ui)
    # Pass the *requested* state (app_state) because counts show potential changes
    # Pass global_df (copyright_df_global assumed available) and constraint
    print("Calculating filter counts...")
    filter_counts = calculate_counts_for_ui(
        app_state, copyright_df_global, faculty_constraint
    )
    print("Filter counts calculated.")

    # 6. Prepare Slice using *Validated* State
    offset = (validated_app_state.page - 1) * validated_app_state.per_page
    df_slice = filtered_df.slice(offset, validated_app_state.per_page)
    print(f"Data slice prepared: Offset {offset}, Length {df_slice.height}")

    # 7. Return results with the *Validated* State
    return ProcessedDataResult(
        app_state=validated_app_state,  # Return the state with corrected page number
        df_slice=df_slice,
        total_filtered_rows=total_filtered_rows,
        total_pages=total_pages,
        filter_counts=filter_counts,
    )


def calculate_counts_for_ui(
    app_state: AppState, global_df: pl.DataFrame, faculty_constraint: dict | None
) -> dict[str, dict[str, int]]:
    """
    Calculates counts for UI filter display based on current filters.

    Simulates toggling each option relative to the current selection,
    applied against the global dataset, while respecting faculty constraints.

    Args:
        app_state: The current AppState reflecting user's choices.
        global_df: The complete, unfiltered DataFrame.
        faculty_constraint: An optional dictionary specifying the faculty constraint (e.g., {"faculty": "EEMCS"}).

    Returns:
        A dictionary where keys are filter columns and values are dicts
        mapping filter options to their calculated counts if toggled.
    """
    all_counts = {}
    current_filters = app_state.get_active_filters()

    # Define the filter options and their styles (assuming these maps are available globally)
    filter_options_map = {
        "workflow_status": WORKFLOW_STYLES,
        "status": STATUS_STYLES,
        "classification": {  # Combine classifications for options
            **{v: LabelT.primary for v in PRIMARY_CLASSIFICATIONS},
            **{v: LabelT.secondary for v in SECONDARY_CLASSIFICATIONS},
            **{v: LabelT.destructive for v in DESTRUCTIVE_CLASSIFICATIONS},
            "None": LabelT.destructive,  # Add 'None' option
        },
        "manual_classification": {  # Same options as classification
            **{v: LabelT.primary for v in PRIMARY_CLASSIFICATIONS},
            **{v: LabelT.secondary for v in SECONDARY_CLASSIFICATIONS},
            **{v: LabelT.destructive for v in DESTRUCTIVE_CLASSIFICATIONS},
            "None": LabelT.destructive,
        },
        "faculty": FACULTY_BADGE_STYLES,
        # Add other filterable columns if they have predefined options
    }

    for filter_key, options in filter_options_map.items():
        # Skip faculty count calculation entirely if user is constrained to a faculty
        # *not* present in the available options (edge case, but possible).
        # Or, more likely, only show the count for *their* faculty.
        if (
            faculty_constraint
            and filter_key == "faculty"
            and faculty_constraint["faculty"] not in options
        ):
            print(
                f"Skipping count for faculty '{faculty_constraint['faculty']}' as it's not in options map for key '{filter_key}'"
            )
            continue

        empty_values = EMPTY_FILTERS.get(filter_key, [])
        all_counts[filter_key] = {}
        # Get the currently selected values for *this specific filter key* from user's state
        current_selections_for_key = (
            set(current_filters.get(filter_key, "").split("|"))
            if current_filters.get(filter_key)
            else set()
        )

        for option_value in options:
            # If constrained and this is the faculty filter, only calculate for the constrained faculty
            if (
                faculty_constraint
                and filter_key == "faculty"
                and option_value != faculty_constraint["faculty"]
            ):
                all_counts[filter_key][option_value] = (
                    0  # Or some indicator it's unavailable
                )
                continue

            if option_value in empty_values:
                all_counts[filter_key][option_value] = 0
                continue

            # --- Calculate count IF this option were toggled ---
            # Start with a fresh copy of the user's current filters for simulation
            temp_filters = current_filters.copy()
            # Apply faculty constraint on top for base filtering during simulation
            if faculty_constraint:
                temp_filters.update(
                    faculty_constraint
                )  # Ensures constraint is always part of base

            temp_selections = current_selections_for_key.copy()

            # Simulate toggling this specific option_value
            is_currently_selected = option_value in temp_selections
            if is_currently_selected:
                temp_selections.remove(option_value)  # Simulate Toggle OFF
            else:
                temp_selections.add(option_value)  # Simulate Toggle ON

            # Update the temporary filter dictionary based on the simulated selections
            if temp_selections:
                temp_filters[filter_key] = "|".join(sorted(list(temp_selections)))
            elif filter_key in temp_filters and not (
                faculty_constraint and filter_key == "faculty"
            ):
                del temp_filters[filter_key]

            # Crucially, if faculty is constrained, ensure the constraint remains
            # even if the user simulation temporarily removed it.
            if (
                faculty_constraint
                and filter_key == "faculty"
                and "faculty" not in temp_filters
            ):
                temp_filters["faculty"] = faculty_constraint["faculty"]

            # Calculate count with these temporary filters using the GLOBAL df
            # Pass the unmodified global_df to the counting helper
            count = _apply_filters_for_count(temp_filters, global_df=global_df)
            all_counts[filter_key][option_value] = count
            # print(f"  Count for {filter_key}={option_value} (toggled {'OFF' if is_currently_selected else 'ON'}) -> {count} (using temp filters: {temp_filters})")

    # print(f"Final Counts: {all_counts}")
    return all_counts


# --- auth ---


@dataclass
class Login:
    email: str
    pwd: str


def create_users_from_secrets():
    # read in users.json.secret from root
    with open("users.json.secret", encoding="utf-8") as f:
        users_data = json.load(f)
    existing_users: list[dict[str, str]] = db.q(f"select email from {users}")
    existing_users = [user.get("email") for user in existing_users]
    for email, data in users_data.items():
        if email not in existing_users:
            users.insert(
                name=data["name"],
                email=email,
                pwd=data["password"],
                faculty=data["faculty"],
                role=data["role"],
            )


users = db.t.users

if users not in db.t:
    users.create(dict(email=str, name=str, pwd=str, faculty=str, role=str), pk="email")

create_users_from_secrets()

login_redir = RedirectResponse("/login", status_code=303)


def before(req, sess):
    auth = req.scope["auth"] = sess.get("auth", None)
    if not auth:
        return login_redir


# --- app setup including CSS + JS ---

CSS = Style("""
    * {
        font-family: "Inter", sans-serif;
        font-optical-sizing: auto;
        font-weight: 400;
        font-style: normal;
    }
    @font-face {
    font-family: 'linotype_univers330_light';
    src: url('linotype_univers_330_light-229790e6f7c56ae5de7a50a0e9404f0a2abb57b023bed877a08ddf385b6244f1-webfont.woff2') format('woff2'),
        url('linotype_univers_330_light-229790e6f7c56ae5de7a50a0e9404f0a2abb57b023bed877a08ddf385b6244f1-webfont.woff') format('woff');
    font-weight: normal; /* Or adjust if needed, e.g., 300 for light */
    font-style: normal;
    }

    /* Base styles for all logo variants */
    .ut-logo-base {
    /* Font applied here mainly sets context, but the children rule is key */
    font-family: 'linotype_univers330_light', sans-serif !important;
    text-transform: uppercase;
    font-weight: bold; /* Explicitly set bold */
    font-size: 2em;
    line-height: 1.1;
    display: inline-block;
    white-space: nowrap;
    /* Resetting potentially interfering properties inherited from * or libraries */
    letter-spacing: normal; /* Reset letter spacing before applying specific values */
    font-optical-sizing: auto; /* Reset */
    font-style: normal; /* Reset */
    }

    .ut-logo-base p {

    font-family: 'linotype_univers330_light', sans-serif !important; /* Apply directly */
    font-weight: bold !important; /* Ensure bold weight is also applied here */
    /* Reset properties that might be inherited incorrectly or set by '*' */
    font-optical-sizing: auto !important;
    font-style: normal !important;
    /* Keep original margin/padding resets */
    margin: 0;
    padding: 0;
    }
    .ut-logo-base p span{

    font-family: 'linotype_univers330_light', sans-serif !important; /* Apply directly */
    font-weight: bold !important; /* Ensure bold weight is also applied here */
    /* Reset properties that might be inherited incorrectly or set by '*' */
    font-optical-sizing: auto !important;
    font-style: normal !important;
    /* Keep original margin/padding resets */
    margin: 0;
    padding: 0;
    }


    /* --- Letter Spacing Adjustments --- */
    .ut-logo-base p:nth-child(1) { /* UNIVERSITY */
    letter-spacing: 0em !important;
    }
    .ut-logo-base p:nth-child(2) {
    letter-spacing: 0em !important;
    }
    .ut-logo-base p:nth-child(3) {
    letter-spacing: 0.1em !important;
    }

    /* --- Color Variants (keep as they were) --- */

    /* Variant 1: Black / Green */
    .logo-variant-1 p:nth-child(1),
    .logo-variant-1 p:nth-child(2) {
    color: #000000; /* Black */
    }
    .logo-variant-1 p:nth-child(3) {
    color: #00675A; /* Green */
    }

    /* Variant 2: Black / Red */
    .logo-variant-2 p:nth-child(1),
    .logo-variant-2 p:nth-child(2) {
    color: #000000; /* Black */
    }
    .logo-variant-2 p:nth-child(3) {
    color: #822433; /* Red */
    }

    /* Variant 3: Blue / Blue-Orange Split */
    .logo-variant-3 p:nth-child(1),
    .logo-variant-3 p:nth-child(2),
    .logo-variant-3 p:nth-child(3) {
    color: #002C5F; /* Blue */
    }
    .logo-variant-3 p:nth-child(3) .split {
    color: #EC7A08; /* Orange */
    }

    /* Variant 4: Blue / Blue-Purple-Red Split */
    .logo-variant-4 p:nth-child(1),
    .logo-variant-4 p:nth-child(2) {
    color: #002C5F; /* Blue */
    }
    .logo-variant-4 p:nth-child(3) {
        color: #4F2D7F; /* Purple */
    }
    .logo-variant-4 p:nth-child(3) .split {
        color: #822433; /* Red */
    }
    /* Note: The original image 4 had Blue/Purple/Red distribution.
    Another interpretation could be:
    .logo-variant-4 p:nth-child(1) { color: #002C5F; } // Blue
    .logo-variant-4 p:nth-child(2) { color: #4F2D7F; } // Purple
    .logo-variant-4 p:nth-child(3) { color: #822433; } // Red
    Choose the interpretation you prefer. The code above uses the split similar to Variant 3.
    */
    html, body { scrollbar-gutter: auto !important; height: 100%; margin: 0; padding: 0; background-color: hsl(var(--b2)); }
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

    /* text highlight color */
    mark {
        background-color: hsl(78 70% 50%) !important;
    }

    /* styles for the card element */
    details > summary { transition: margin 150ms ease-out; }
    details[open] > summary { margin-bottom: 0; }
    details summary .chevron-icon { transition: transform 0.2s ease-in-out; }
    details[open] summary .chevron-icon { transform: rotate(-180deg); }

""")

FONT = Link(
    rel="stylesheet",
    href="https://fonts.googleapis.com/css2?family=Inter:ital,opsz,wght@0,14..32,100..900;1,14..32,100..900&display=swap",
    type="text/css",
)

ALPINE_TOOLTIP_JS = r"""
    document.addEventListener('alpine:init', () => {

        Alpine.directive('tooltip', (el, { modifiers, expression }, { cleanup }) => {
            let tooltipText = expression;
            let tooltipArrow = modifiers.includes('noarrow') ? false : true;
            let tooltipPosition = 'top';
            let tooltipId = 'tooltip-' + Date.now().toString(36) + Math.random().toString(36).substring(2, 7);
            let positions = ['top', 'bottom', 'left', 'right'];
            let elementPosition = getComputedStyle(el).position;

            for (let position of positions) {
                if (modifiers.includes(position)) {
                    tooltipPosition = position;
                    break;
                }
            }

            if(!['relative', 'absolute', 'fixed'].includes(elementPosition)){
                el.style.position='relative';
            }

            let tooltipHTML = `
                <div id="${tooltipId}" x-data="{ tooltipVisible: false, tooltipText: '${tooltipText}', tooltipArrow: ${tooltipArrow}, tooltipPosition: '${tooltipPosition}' }" x-ref="tooltip" x-init="setTimeout(function(){ tooltipVisible = true; }, 1);" x-show="tooltipVisible" :class="{ 'top-0 left-1/2 -translate-x-1/2 -mt-0.5 -translate-y-full' : tooltipPosition == 'top', 'top-1/2 -translate-y-1/2 -ml-1.5 left-0 -translate-x-full' : tooltipPosition == 'left', 'bottom-0 left-1/2 -translate-x-1/2 -mb-0.5 translate-y-full' : tooltipPosition == 'bottom', 'top-1/2 -translate-y-1/2 -mr-1.5 right-0 translate-x-full' : tooltipPosition == 'right' }" class="absolute w-auto text-sm" x-cloak>
                    <div x-show="tooltipVisible"
                        x-transition:enter="transition ease-out duration-200"
                        x-transition:enter-start="opacity-0 scale-90 -translate-x-2"
                        x-transition:enter-end="opacity-100 scale-100 translate-x-0"
                        x-transition:leave="transition ease-in duration-200"
                        x-transition:leave-start="opacity-100 scale-100 translate-x-0"
                        x-transition:leave-end="opacity-0 scale-90 -translate-x-2"
                    class="relative px-2 py-1 text-white bg-emerald-500 rounded bg-opacity-90">
                        <p x-text="tooltipText" class="flex-shrink-0 block text-xs whitespace-nowrap"></p>
                        <div x-ref="tooltipArrow" x-show="tooltipArrow" :class="{ 'bottom-0 -translate-x-1/2 left-1/2 w-2.5 translate-y-full' : tooltipPosition == 'top', 'right-0 -translate-y-1/2 top-1/2 h-2.5 -mt-px translate-x-full' : tooltipPosition == 'left', 'top-0 -translate-x-1/2 left-1/2 w-2.5 -translate-y-full' : tooltipPosition == 'bottom', 'left-0 -translate-y-1/2 top-1/2 h-2.5 -mt-px -translate-x-full' : tooltipPosition == 'right' }" class="absolute inline-flex items-center justify-center overflow-hidden">
                            <div :class="{ 'origin-top-left -rotate-45' : tooltipPosition == 'top', 'origin-top-left rotate-45' : tooltipPosition == 'left', 'origin-bottom-left rotate-45' : tooltipPosition == 'bottom', 'origin-top-right -rotate-45' : tooltipPosition == 'right' }" class="w-1.5 h-1.5 transform bg-black bg-opacity-90"></div>
                        </div>
                    </div>
                </div>
            `;

            el.dataset.tooltip = tooltipId;

            let mouseEnter = function(event){
                el.innerHTML += tooltipHTML;
            };

            let mouseLeave = function(event){
                document.getElementById(event.target.dataset.tooltip).remove();
            };

            el.addEventListener('mouseenter', mouseEnter);
            el.addEventListener('mouseleave', mouseLeave);

            cleanup(() => {
                el.removeEventListener('mouseenter', mouseEnter);
                el.removeEventListener('mouseleave', mouseLeave);
            })
        })

    })


"""
FILTER_JS_HELPER = Script("""
function getFilterValues() {
    const form = document.getElementById('filter-form');
    if (!form) return '{}'; // Return empty JSON object string if form not found

    const formData = new FormData(form);
    const filters = {};

    // Group checkbox values by name
    const checkboxGroups = {};
    form.querySelectorAll('input[type="checkbox"][name^="filter_"]').forEach(cb => {
        if (!checkboxGroups[cb.name]) {
            checkboxGroups[cb.name] = [];
        }
        if (cb.checked) {
            checkboxGroups[cb.name].push(cb.value);
        }
    });

    // Add grouped checkbox values (joined by '|')
    for (const name in checkboxGroups) {
        if (checkboxGroups[name].length > 0) {
            // Extract filter key like 'workflow_status' from 'filter_workflow_status'
            const filterKey = name.substring('filter_'.length);
            filters[filterKey] = checkboxGroups[name].sort().join('|');
        }
    }

    // Add text input values
    form.querySelectorAll('input[type="text"][name^="filter_"], input[type="search"][name^="filter_"]').forEach(input => {
        if (input.value) {
            const filterKey = input.name.substring('filter_'.length);
             // Only add if not already handled by checkboxes (shouldn't happen with current setup, but safe)
            if (!(filterKey in filters)) {
                 filters[filterKey] = input.value;
            }
        }
    });

    // Add hidden nav state (page will be reset, others kept)
    const navState = {};
    form.querySelectorAll('input[type="hidden"]').forEach(input => {
        if (input.name !== 'page') { // Exclude page, reset below
            navState[input.name] = input.value;
        }
    });

    // Prepare final hx-vals object structure expected by AppState/backend
    const hxVals = {
        ...navState, // Include per_page, sort_by, sort_desc
        page: '1',   // Always reset page to 1 on apply
    };
    // Add the filter_* prefixed keys
    for (const key in filters) {
        hxVals[`filter_${key}`] = filters[key];
    }

    console.log("getFilterValues returning:", JSON.stringify(hxVals)); // Debugging
    return JSON.stringify(hxVals);
}
""")

FILTER_JS_LISTENER = Script(
    """
document.body.addEventListener('htmx:configRequest', function(event) {
    // Check if the event was triggered by our Apply Filters button
    if (event.detail.elt.id === 'apply-filters-btn') {
        console.log("htmx:configRequest triggered for Apply button");
        try {
            // Call getFilterValues, which returns a JSON string
            const filterValsJson = getFilterValues();
            // Parse the JSON string into an object
            const filterParams = JSON.parse(filterValsJson);

            // event.detail.parameters is the object HTMX will use for the request
            // Merge our dynamic filter parameters into it
            Object.assign(event.detail.parameters, filterParams);

            console.log("Parameters after config:", event.detail.parameters);
        } catch (e) {
            console.error("Error processing filter values in htmx:configRequest:", e);
            // Optionally prevent the request if parameters are bad
            // event.preventDefault();
        }
    }
    // You could add similar logic for the reset button if needed,
    // but its parameters are simpler and set directly in hx-get.
});
""",
    defer=True,
)

MODAL_TRIGGER = Script("""
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
MODAL_INTERACTION = Script(f"""
    // Function to update a pill's appearance and hidden input value
    // ... (keep existing implementation of updatePill, markDirty, resetModalForm) ...
        function updatePill(fieldName, newValue, newText, newStyleClass) {{
        const pillElement = document.getElementById(`pill-display-${{fieldName}}`);
        const inputElement = document.getElementById(`input-${{fieldName}}`);
        if (pillElement && inputElement) {{
            inputElement.value = newValue;
            pillElement.textContent = newText;
            pillElement.classList.remove('uk-label-primary', 'uk-label-secondary', 'uk-label-destructive', 'badge-neutral', 'badge-primary', 'badge-error'); // Remove old styles
            if (newStyleClass) {{
                pillElement.classList.add(newStyleClass); // Add new style
            }} else {{
                    pillElement.classList.add('{str(DEFAULT_PILL_STYLE)}'); // Default fallback
            }}
            const drop = UIkit.drop(pillElement.closest('[uk-drop]'));
            if (drop) {{ drop.hide(false); }}
            markDirty();
        }} else {{ console.error(`Cannot find pill or input elements for ${{fieldName}}`); }}
    }}
    function markDirty() {{
        const indicator = document.getElementById('save-indicator');
        const saveButton = document.getElementById('modal-save-btn');
        if (indicator) indicator.classList.remove('hidden');
        if (saveButton) saveButton.disabled = false;
    }}
    function resetModalForm() {{
        console.log('Resetting modal form');
        const form = document.getElementById('modal-details-form');
        if (!form) return;
        // Reset editable pills
        form.querySelectorAll('input[data-original-value][id^="input-"]').forEach(input => {{
            const originalValue = input.dataset.originalValue;
            const originalText = input.dataset.originalText || originalValue;
            const originalStyle = input.dataset.originalStyle || '{str(DEFAULT_PILL_STYLE)}';
            const fieldName = input.id.replace('input-', '');
            input.value = originalValue;
            const pillElement = document.getElementById(`pill-display-${{fieldName}}`);
            if (pillElement) {{
                pillElement.textContent = originalText;
                pillElement.className = ''; // Clear all classes first
                pillElement.classList.add('uk-label', 'badge-sm'); // Re-add base classes
                if (originalStyle) {{ pillElement.classList.add(originalStyle); }}
                else {{ pillElement.classList.add('{str(DEFAULT_PILL_STYLE)}'); }}
            }}
        }});
        // Reset remarks textarea
        const remarksTextarea = form.querySelector('#modal_remarks');
        if (remarksTextarea && typeof remarksTextarea.dataset.originalValue !== 'undefined') {{
            remarksTextarea.value = remarksTextarea.dataset.originalValue;
        }}
        // Reset save indicator/button
        const indicator = document.getElementById('save-indicator');
        const saveButton = document.getElementById('modal-save-btn');
        if (indicator) indicator.classList.add('hidden');
        if (saveButton) saveButton.disabled = true;
    }}
""")

JS = (
    Script(src="https://cdn.jsdelivr.net/npm/uikit@3.latest/dist/js/uikit.min.js"),
    Script(
        src="https://cdn.jsdelivr.net/npm/uikit@3.latest/dist/js/uikit-icons.min.js"
    ),
    Script(src="https://unpkg.com/alpinejs", defer="defer"),
    Script(
        ALPINE_TOOLTIP_JS,
    ),
    Script(src="https://cdn.tailwindcss.com"),
)

BODY_JS = (MODAL_TRIGGER, MODAL_INTERACTION)


# for static files, use the following regex to match the file extensions
reg_re_param(
    "static",
    "ico|gif|jpg|jpeg|webm|css|js|woff|png|svg|mp4|webp|ttf|otf|eot|woff2|txt|xml|html|pdf|md",
)

bware = Beforeware(
    before, skip=[r"/favicon\.ico", r"/static/.*", r"/imgs/.*", r".*\.css", "/login"]
)
init_headers = (
    Theme.slate.headers(
        mode="light",
        daisy=True,
        katex=False,
        radii=ThemeRadii.lg,
        shadows=ThemeShadows.lg,
    ),
    CSS,
    FONT,
    *JS,
)
app, rt = fast_app(
    before=bware,
    hdrs=init_headers,
    exts="loading-states",
    debug=True,
)

# Mount static files directory AFTER creating the app instance
app.mount("/imgs", StaticFiles(directory="."), name="images")

setup_toasts(app)

# --- constants / mappings ---


# a global dict that stores values that are missing from the global df for each filterable column that has a list of expected values


EMPTY_FILTERS: dict[str, list[str]] = {
    "classification": [],
    "status": [],
    "workflow_status": [],
    "manual_classification": [],
    "faculty": [],
    "department": [],
    "course_name": [],
}


def _update_empty_filters(filter_counts: dict[str, dict[str, int]]):
    """
    input filter_counts is the result of get_filter_counts() run on a non-filtered global df -- so on init and on refresh.
    Stores the values that are missing from the global df for each filterable column in EMPTY_FILTERS.
    """
    global EMPTY_FILTERS

    # Update EMPTY_FILTERS based on filter_counts
    for col, counts in filter_counts.items():
        if col in EMPTY_FILTERS:
            EMPTY_FILTERS[col] = [key for key, count in counts.items() if count == 0]


# --- component templates ---
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
    # print(f"--- checkbox filter group: {filter_key} ---")
    # print(f"  current_values: {current_values}")
    # print(f"  current_total: {current_total}")
    # print(f"  counts: {counts}")
    # print(f"  options: {options}")
    # print(f"  label_text: {label_text}")
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
        is_disabled = (count == 0) and not is_checked
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

        if isinstance(count, int) and count > current_total:
            count -= current_total

        # --- Enabled option rendering ---

        if not is_disabled:
            badge_cls = (
                f"{style_class} badge-outline ml-1 align-middle text-xs cursor-help"
            )
            count_badge = Span(
                f"{'▲' if all([isinstance(count, int), selected_values, count != current_total]) else '▼' if isinstance(count, int) and count != current_total else 'Σ'} {count}",
                title=f"{'Will increase the current total itemcount by this amount when selected.' if all([isinstance(count, int), selected_values, count != current_total]) else 'Will decrease the current total itemcount to this amount when selected.' if isinstance(count, int) and count != current_total else 'The current total itemcount'} ",
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
                hx_post=data_grid.to(),  # Target the central data route
                hx_target="#data-grid-component",  # Update the grid
                hx_swap="outerHTML",  # Replace the grid component
                hx_trigger="change delay:300ms",  # Trigger on change with delay
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


# --- Faculty color mapping for unique badge styles ---
FACULTY_BADGE_STYLES = {
    "BMS": "badge-success",
    "EEMCS": "badge-info",
    "ET": "badge-warning",
    "ITC": "badge-error",
    "TNW": "badge-secondary",
}
# --- filter dataframe ---


def get_filtered_sorted_df(
    app_state: AppState, extra_constraints: dict = None
) -> pl.DataFrame:
    """
    Applies filtering and sorting based on AppState and optional extra_constraints.

    Args:
        app_state: The AppState object containing user filters and sort settings.
        extra_constraints: Optional dictionary of additional constraints to apply (e.g., {"faculty": "EEMCS"}).

    Returns:
        A Polars DataFrame containing the filtered and sorted data.
    """
    df = copyright_df_global  # Start with the global DataFrame
    user_filters = app_state.get_active_filters()
    sort_by = app_state.sort_by
    sort_desc = app_state.sort_desc

    filter_expressions = []

    # --- Process user filters from app_state ---
    if user_filters:
        for col, value in user_filters.items():
            actual_col = col  # Assume col name matches DataFrame column directly
            if actual_col not in df.columns:
                print(
                    f"Warning: Filter column '{actual_col}' not in DataFrame. Skipping."
                )
                continue
            if value is None or value == "":  # Skip empty filter values explicitly
                continue

            try:
                or_values = (
                    value.split("|")
                    if isinstance(value, str) and "|" in value
                    else [value]
                )
                or_expressions = []
                for or_value in or_values:
                    or_value = or_value.strip()
                    if not or_value:
                        continue

                    # Handle 'None' string specifically for classification fields
                    if (
                        col in ["manual_classification", "classification"]
                        and or_value.lower() == "none"
                    ):
                        or_expressions.append(
                            (pl.col(actual_col).is_null())
                            | (pl.col(actual_col) == "")
                            | (
                                pl.col(actual_col) == "-"
                            )  # Include hyphen check if needed
                        )
                        continue  # Move to next or_value

                    # Handle other string comparisons
                    if df[actual_col].dtype == pl.Utf8:
                        # Exact match for categorical-like fields
                        if col in [
                            "status",
                            "workflow_status",
                            "classification",
                            "manual_classification",
                            "faculty",
                        ]:
                            or_expressions.append(
                                pl.col(actual_col).str.to_lowercase()
                                == or_value.lower()
                            )
                        # Contains match for text search fields
                        else:
                            # Use regex for case-insensitive contains
                            or_expressions.append(
                                pl.col(actual_col).str.contains(
                                    f"(?i){re.escape(or_value)}"
                                )
                            )
                    # Handle numeric comparisons (example, adjust as needed)
                    elif df[actual_col].dtype in (
                        pl.Int64,
                        pl.Int32,
                        pl.Float64,
                        pl.Float32,
                    ):
                        with contextlib.suppress(ValueError):
                            or_expressions.append(pl.col(actual_col) == float(or_value))
                    # Add other type handling if necessary

                # Combine OR expressions for a single filter key
                if len(or_expressions) > 1:
                    filter_expressions.append(pl.any_horizontal(or_expressions))
                elif len(or_expressions) == 1:
                    filter_expressions.append(or_expressions[0])

            except Exception as e:
                print(f"Filter warning processing '{actual_col}'='{value}': {e}")

    # --- Process extra constraints (e.g., faculty auth) ---
    if extra_constraints:
        print(f"Applying extra constraints: {extra_constraints}")
        for col, value in extra_constraints.items():
            if col not in df.columns:
                print(f"Warning: Constraint column '{col}' not in DataFrame. Skipping.")
                continue
            if value:  # Only apply if constraint value is non-empty
                # Assuming simple equality constraint for now (e.g., faculty)
                # Adapt if more complex constraints are needed
                filter_expressions.append(
                    pl.col(col).str.to_lowercase() == str(value).lower()
                )  # Ensure comparison is case-insensitive

    # --- Apply combined filters ---
    if filter_expressions:
        try:
            # Combine all expressions with AND logic
            combined_expression = pl.all_horizontal(filter_expressions)
            df = df.filter(combined_expression)
        except Exception as e:
            print(f"Error applying combined filters: {e}")
            # Fallback: return empty DataFrame on error? Or original?
            return df.clear()  # Example: return empty

    # --- Apply sorting ---
    if sort_by and sort_by in df.columns:
        try:
            df = df.sort(by=sort_by, descending=sort_desc, nulls_last=True)
        except Exception as e:
            print(f"Sort warning on '{sort_by}': {e}")
    elif sort_by:
        print(
            f"Warning: Sort column '{sort_by}' not found in DataFrame. Skipping sort."
        )

    return df


def _apply_filters_for_count(filters: dict[str, str], global_df: pl.DataFrame) -> int:
    """
    Helper to apply filters to the provided global DataFrame and return the count.
    Optimized for counting.

    Args:
        filters: Dictionary of filters to apply.
        global_df: The complete, unfiltered DataFrame to filter on.

    Returns:
        The count of rows matching the filters.
    """
    df = global_df  # Use the passed global DataFrame

    if not filters:
        return df.height  # Return total count if no filters

    filter_expressions = []
    # Re-use the filter expression building logic from get_filtered_sorted_df
    # This ensures consistency in how filters are interpreted for counting vs. display
    for col, value in filters.items():
        actual_col = col
        if actual_col not in df.columns:
            continue  # Skip if column doesn't exist
        if value is None or value == "":
            continue

        # --- Start copied/adapted logic ---
        try:
            or_values = (
                value.split("|") if isinstance(value, str) and "|" in value else [value]
            )
            or_expressions = []
            for or_value in or_values:
                or_value = or_value.strip()
                if not or_value:
                    continue

                if (
                    col in ["manual_classification", "classification"]
                    and or_value.lower() == "none"
                ):
                    or_expressions.append(
                        (pl.col(actual_col).is_null())
                        | (pl.col(actual_col) == "")
                        | (pl.col(actual_col) == "-")
                    )
                    continue

                if df[actual_col].dtype == pl.Utf8:
                    if col in [
                        "status",
                        "workflow_status",
                        "classification",
                        "manual_classification",
                        "faculty",
                    ]:
                        or_expressions.append(
                            pl.col(actual_col).str.to_lowercase() == or_value.lower()
                        )
                    else:
                        or_expressions.append(
                            pl.col(actual_col).str.contains(
                                f"(?i){re.escape(or_value)}"
                            )
                        )
                elif df[actual_col].dtype in (
                    pl.Int64,
                    pl.Int32,
                    pl.Float64,
                    pl.Float32,
                ):
                    with contextlib.suppress(ValueError):
                        or_expressions.append(pl.col(actual_col) == float(or_value))

            if len(or_expressions) > 1:
                filter_expressions.append(pl.any_horizontal(or_expressions))
            elif len(or_expressions) == 1:
                filter_expressions.append(or_expressions[0])
        except Exception as e:
            print(f"Count Filter warning on '{actual_col}'='{value}': {e}")
        # --- End copied/adapted logic ---

    if filter_expressions:
        try:
            # Apply filters and get count directly using lazy evaluation
            count = (
                df.lazy()
                .filter(pl.all_horizontal(filter_expressions))
                .select(pl.count())
                .collect()
                .item()
            )
            return count
        except Exception as e:
            print(f"Count Filter error applying filters {filters}: {e}")
            return 0  # Return 0 on error
    else:
        # No valid filters were generated, or filters were empty
        return df.height


'''
def _calc_total_for_filters(filters: dict[str, str | int]) -> int:
    """
    For a given dict with filter keys and values, calculate the total count of items that match the filters.
    """
    df = copyright_df_global
    if not filters:
        return df.height  # Return total count if no filters

    filter_expressions = []
    for col, value in filters.items():
        actual_col = col
        if value and actual_col in df.columns:
            try:
                # --- Re-use the same OR logic as get_filtered_sorted_df ---
                or_values = (
                    value.split("|")
                    if isinstance(value, str) and "|" in value
                    else [value]
                )
                or_expressions = []
                for or_value in or_values:
                    or_value = or_value.strip()
                    if not or_value:
                        continue
                    # --- Special handling for manual_classification empty filter ---
                    if col == "manual_classification" and or_value == "None":
                        or_expressions.append(
                            (pl.col(actual_col).is_null())
                            | (pl.col(actual_col) == "")
                            | (pl.col(actual_col) == "-")
                        )
                        continue
                    if df[actual_col].dtype == pl.Utf8:
                        if col in [
                            "status",
                            "workflow_status",
                            "classification",
                            "manual_classification",
                            "faculty",
                        ]:
                            or_expressions.append(
                                pl.col(actual_col).str.to_lowercase()
                                == or_value.lower()
                            )
                        else:
                            or_expressions.append(
                                pl.col(actual_col).str.contains(f"(?i){or_value}")
                            )
                    elif df[actual_col].dtype in (
                        pl.Int64,
                        pl.Int32,
                        pl.Float64,
                        pl.Float32,
                    ):
                        with contextlib.suppress(ValueError):
                            or_expressions.append(pl.col(actual_col) == float(or_value))
                if len(or_expressions) > 1:
                    filter_expressions.append(pl.any_horizontal(or_expressions))
                elif len(or_expressions) == 1:
                    filter_expressions.append(or_expressions[0])
            except Exception as e:
                print(
                    f"Count Filter warning on '{actual_col}' for value '{value}': {e}"
                )  # Less critical here

    if filter_expressions:
        try:
            # Apply filters, collect, and return height
            # Use lazy frame for potentially better optimization
            total = (
                df.lazy().filter(pl.all_horizontal(filter_expressions)).collect().height
            )
            return total
        except Exception as e:
            print(f"calc_total_for_filters error applying filters {filters}: {e}")
            return 0  # Return 0 on error
    else:
        # No valid filters were generated, return total count
        return df.height
'''

"""
@rt("/update_filter_counts", methods=["POST"])
async def update_filter_counts_endpoint(session: dict, request: Request):
    '''
    Updates session state based on checkbox change, calculates all filter counts,
    and returns OOB swaps for all checkbox filter groups.
    '''
    form_data = await request.form()
    request_params = dict(form_data)

    # 1. Get current AppState from session
    app_state_dict = session.get("app_state", {})
    app_state = AppState(**app_state_dict)

    # 2. Update AppState filters based *only* on the received form data
    #    (Use the simplified update logic relying on presence/absence)
    # --- Process Filters: Start fresh, only add if present in request_params ---
    new_filters = {}
    all_filter_keys = [
        "workflow_status",
        "status",
        "classification",
        "manual_classification",
        "faculty",
        "department",
        "course_name",  # Include all potential filter keys
    ]
    for key in all_filter_keys:
        filter_param_key = f"filter_{key}"
        # Handle multi-value checkboxes correctly using getlist from the ORIGINAL form_data
        values = form_data.getlist(filter_param_key)
        if values:  # If any checkbox for this key was checked
            # Filter out potential empty strings if using hidden input method (though we removed it)
            checked_values = [v for v in values if v != ""]
            if checked_values:
                new_filters[key] = "|".join(sorted(checked_values))
        elif filter_param_key in request_params:  # Handle text inputs etc. from dict
            value = request_params[filter_param_key]
            if value:  # Add if non-empty
                new_filters[key] = str(value)

    # Update the state's filters object
    app_state.update_from_req(
        request_params=request_params,
        new_filters=new_filters,
        auth_details=session.get("auth", {}),
    )
    # Note: Don't update page/sort state here, only filters change

    # 3. Calculate counts based on the NEW state
    counts = get_filter_counts(app_state)
    print(counts)
    # 4. Prepare OOB fragments by re-rendering ALL checkbox groups
    oob_fragments = []
    auth_details = session.get("auth", {})  # Get auth details for permission check
    is_admin = auth_details.get("role") == "admin"
    user_faculty = auth_details.get("faculty")

    # Define options map (same as in page_header_component)
    # Define the options we need counts for (reuse from page_header_component)
    # Combine all options into one structure for iteration
    # --- Define Options for Checkbox Groups (Keep these as they are) ---
    workflow_options = {
        "ToDo": WORKFLOW_STYLES.get("ToDo"),
        "InProgress": WORKFLOW_STYLES.get("InProgress"),
        "Done": WORKFLOW_STYLES.get("Done"),
    }
    status_options = {
        "Published": STATUS_STYLES.get("Published"),
        "Unpublished": STATUS_STYLES.get("Unpublished"),
        "Deleted": STATUS_STYLES.get("Deleted"),
    }
    classification_options = {
        **{v: LabelT.primary for v in PRIMARY_CLASSIFICATIONS},
        **{v: LabelT.secondary for v in SECONDARY_CLASSIFICATIONS},
        **{v: LabelT.destructive for v in DESTRUCTIVE_CLASSIFICATIONS},
        "None": LabelT.destructive,
    }
    manual_classification_options = classification_options
    faculty_options = {
        f: FACULTY_BADGE_STYLES.get(f, "badge-secondary")
        for f in ["BMS", "EEMCS", "ET", "ITC", "TNW"]
    }
    checkbox_groups_to_render = [
        ("workflow_status", "Workflow Status", workflow_options),
        ("status", "Status", status_options),
        (
            "manual_classification",
            "Manual Classification",
            manual_classification_options,
        ),
        ("classification", "Classification", classification_options),
    ]
    # Conditionally add faculty
    if is_admin or not user_faculty or user_faculty == "all":
        checkbox_groups_to_render.append(("faculty", "Faculty", faculty_options))

    for key, label, options_map in checkbox_groups_to_render:
        group_id = f"filter-group-{key}"
        # Re-render the group using the NEW app_state and NEW counts
        rendered_group = create_checkbox_filter_group(
            filter_key=key,
            label_text=label,
            options=options_map,
            current_values=app_state.filters.get(key),
            counts=counts.get(key),
            current_total=app_state.current_total,
        )
        # Add the OOB swap attribute to the outer Div
        if hasattr(rendered_group, "attrs"):
            rendered_group.attrs["hx-swap-oob"] = f"outerHTML:#{group_id}"
        else:
            # Should not happen if create_... returns a Div, but handle defensively
            print(f"Warning: Cannot add OOB swap to non-FT object for key {key}")

        oob_fragments.append(rendered_group)

    return oob_fragments, HtmxResponseHeaders(reswap="none")
"""


# Function to get counts for all options based on current selections
def get_filter_counts(app_state: AppState) -> dict[str, dict[str, int]]:
    """Calculates the counts for each filter option based on current state."""
    all_counts = {}
    current_filters = app_state.get_active_filters()
    # Define the options we need counts for (reuse from page_header_component)
    # Combine all options into one structure for iteration
    filter_options_map = {
        "workflow_status": {k: v for k, v in WORKFLOW_STYLES.items()},
        "status": {k: v for k, v in STATUS_STYLES.items()},
        "classification": {
            **{v: LabelT.primary for v in PRIMARY_CLASSIFICATIONS},
            **{v: LabelT.secondary for v in SECONDARY_CLASSIFICATIONS},
            **{v: LabelT.destructive for v in DESTRUCTIVE_CLASSIFICATIONS},
            "None": LabelT.destructive,
        },
        "manual_classification": {
            **{v: LabelT.primary for v in PRIMARY_CLASSIFICATIONS},
            **{v: LabelT.secondary for v in SECONDARY_CLASSIFICATIONS},
            **{v: LabelT.destructive for v in DESTRUCTIVE_CLASSIFICATIONS},
            "None": LabelT.destructive,
        },
        "faculty": {
            f: FACULTY_BADGE_STYLES.get(f, "badge-secondary")
            for f in ["BMS", "EEMCS", "ET", "ITC", "TNW"]
        },
    }

    for filter_key, options in filter_options_map.items():
        empty_values = []
        if filter_key in EMPTY_FILTERS:
            empty_values = EMPTY_FILTERS[filter_key]
        all_counts[filter_key] = {}
        current_selections_for_key = (
            set(current_filters.get(filter_key, "").split("|"))
            if current_filters.get(filter_key)
            else set()
        )

        for option_value in options:
            if option_value in empty_values:
                all_counts[filter_key][option_value] = 0
                continue

            # --- Calculate count IF this option were toggled ---
            temp_filters = current_filters.copy()
            temp_selections = current_selections_for_key.copy()

            # Simulate toggling this option
            if option_value in temp_selections:
                all_counts[filter_key][option_value] = "-"
            else:
                temp_selections.add(option_value)  # Toggle ON

            # Update the temporary filter value for this key
            if temp_selections:
                temp_filters[filter_key] = "|".join(sorted(list(temp_selections)))
            else:
                temp_filters.pop(filter_key, None)  # Remove key if no selections left

            # Calculate count with these temporary filters
            count = _apply_filters_for_count(temp_filters)
            all_counts[filter_key][option_value] = count

    return all_counts


# --- save changes ---
async def store_item_changes(
    input_data: list[dict[str, str | int]] | dict[str, str | int], auth_details: dict
) -> None:
    """
    For a list of dicts containing at least a material_id and updated fields,
    for each item, retrieve the full item data for the other fields from the copyright_df_global
    make sure the updated fields in the dict are not overwritten with the original values

    then send the new list of dicts to the database for update
    """
    user_info = {
        "email": auth_details.get("email"),
        "name": auth_details.get("name"),
        "faculty": auth_details.get("faculty"),
        "role": auth_details.get("role"),
    }
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

    await update_copyright_items(
        full_data_list, update_relations=False, overwrite=True, user_info=user_info
    )


@rt("/save_details", methods=["POST"])
async def save_item_details(
    request: Request,
    session: dict,
    # Only expect editable fields + ID from the form
    material_id: int,
    workflow_status: str = "",
    manual_classification: str = "",
    remarks: str = "",
):
    """
    Route to store edited item details from the modal back to the database
    and return an updated grid view via OOB swap.
    """
    print(f"Saving changes for material_id: {material_id}")
    update_data = {
        "material_id": material_id,
        "workflow_status": workflow_status,
        "manual_classification": manual_classification
        if manual_classification.lower() != "none"
        else None,  # Handle 'None' string
        "remarks": remarks,
    }

    oob_grid_swap = Div(hx_swap_oob="true")  # Default empty OOB swap

    try:
        # 1. Save changes to DB
        await store_item_changes(update_data, session.get("auth", {}))

        # 2. Reload global data (consider efficiency later)
        global copyright_df_global
        print("Reloading global DataFrame after save...")
        copyright_df_global = retrieve_copyright_items()  # Assumes this function exists
        print("Global DataFrame reloaded.")

        # 3. Load current state (reflecting filters/sort *before* save)
        # We need the state to know *which* page/sort order to display after save
        app_state_before_save = load_app_state(session)

        # 4. Fetch updated data based on the *existing* state
        # This will apply current filters/sort and validate the page number
        # against the *newly reloaded* global data.
        processed_data = await fetch_data(app_state_before_save, session)

        # 5. Update session if page validation occurred in fetch_data
        if processed_data.app_state.page != app_state_before_save.page:
            session["app_state"] = asdict(processed_data.app_state)

        # 6. Render the updated grid component
        updated_grid_component = render_data_grid_component(
            df_slice=processed_data.df_slice,
            app_state=processed_data.app_state,  # Use validated state
            total_filtered_rows=processed_data.total_filtered_rows,
            total_pages=processed_data.total_pages,
        )

        # 7. Create OOB Swap Div for the Grid
        oob_grid_swap = Div(
            to_xml(updated_grid_component),  # Render the component to HTML string
            hx_swap_oob="outerHTML:#data-grid-component",  # Target the grid for replacement
        )

        add_toast(session, f"Successfully updated item {material_id}", "success")

    except Exception as e:
        print(f"Error saving changes for {material_id}: {e}")
        traceback.print_exc()
        add_toast(
            session, f"Error saving changes for {material_id}: {str(e)[:100]}", "error"
        )  # Show truncated error

    # Return only the OOB swap for the grid.
    # The modal should be closed client-side after successful save (or on explicit close).
    # Use HTMX headers to prevent default swap.
    return oob_grid_swap, HtmxResponseHeaders(reswap="none")


# --- retrieve detailed data  ---
@rt("/osiris/{material_id:int}")
async def get_osiris_data(material_id: int):
    """
    Returns enriched data for the given material_id.
    should always return at least base item data, even if no enriched data is found.
    Logic is almost completely found in the db.retrieve module.

    """
    data = retrieve_osiris_data([material_id])
    if not data or len(data) == 0 or not isinstance(data, list):
        return []
    return data[0]


# --- retrieve pdfs ---
@rt(r"/file/{material_id:int}")
async def get_file(material_id: int):
    """
    Directly returns the PDF file for the given material_id.
    """
    pdf_root_dir = Path("pdf_downloads")
    ext = "pdf"
    file_path = pdf_root_dir / f"{material_id}.{ext}"
    print(f"requested file: {file_path}")
    if not file_path.exists():
        return HTMLResponse("File not found", status_code=404)
    return FileResponse(file_path)


@rt("/pdf/{material_id:int}")
async def get_pdf_element(material_id: int):
    """
    Returns element displaying the PDF file for the given material_id as an embedded PDF viewer.
    If the PDF file is not found, a message will be displayed instead.
    """
    root_folder = Path("pdf_downloads")

    pdf_file_path = root_folder / f"{material_id}.pdf"
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


# --- retrieve extracted text ---


def get_extracted_text(material_id: int) -> str:
    """
    Returns extracted text from the PDF file for the given material_id.
    If annotated text is available, it will be used; otherwise, the plain text will be returned.
    If neither are available, an str with an error message will be returned.
    """
    root_folder = Path("pdf_downloads")
    # suffixes to check in order of priority
    suffixes = ["_annotated.md", "_paddle.txt", ".md", ".txt"]
    for suff in suffixes:
        extracted_text_path = root_folder / f"{material_id}{suff}"
        if extracted_text_path.exists():
            break
    if not extracted_text_path.exists():
        return "No extracted text found."

    print(f"retrieving extracted text for {material_id} from {extracted_text_path}")

    text_element = f"Error loading text from {extracted_text_path.name}."
    if extracted_text_path.exists():
        try:
            with open(extracted_text_path, encoding="utf-8") as f:
                text_element = f.read()

        except Exception as e:
            print(f"Error processing text file {extracted_text_path}: {e}")

    return text_element


@rt("/text/{material_id:int}", methods=["GET"])
async def get_extracted_text_element(material_id: int) -> FT:
    text = get_extracted_text(material_id)
    return Div(NotStr(text))


# --- retrieve found entities ---


# dataclasses for entities
@dataclass
class Entity:
    """
    Represents an entity extracted from the text.
    Attributes:
        start (int): The starting index of the entity in the text.
        end (int): The ending index of the entity in the text.
        label (str): The label of the entity.
        text (str): The text of the entity.
        score (float | None): The confidence score of the entity (optional).
    """

    start: int
    end: int
    label: str
    text: str
    score: float | None = None


@dataclass
class Entities:
    """
    Represents a collection of entities extracted from the text.
    Attributes:
        items (list[Entity]): A list of Entity objects.

    Functions:
        group_by_label: Groups entities by their labels.
        group_and_sort: Sorts the grouped entities by label and alphabetically within each group.
    """

    items: list[Entity] = field(init=False, default_factory=list)
    grouped_by_label: dict[str, list[Entity]] = field(init=False, default=None)

    def __init__(self, entities: list[dict[str, int | str | float]]):
        self.items = [Entity(**ent) for ent in entities]

    def group_by_label(self):
        self.grouped_by_label: dict[str, list[Entity]] = defaultdict(list)

        for ent in self.items:
            self.grouped_by_label[ent.label].append(ent)

    def group_and_sort(self):
        if not self.grouped_by_label:
            self.group_by_label()

        sorted_labels = sorted(
            self.grouped_by_label.keys(),
            key=lambda x: (
                0
                if "recognized" in x.lower()
                # recognized first
                else 1
                if "copyright" in x.lower() or "license" in x.lower()
                # copyright or license next
                else 2
                if "university" in x.lower()
                # university next
                else 3  # everything else last
            ),
        )

        for label in sorted_labels:
            # sort alphabetically within each label group
            self.grouped_by_label[label] = sorted(
                self.grouped_by_label[label], key=lambda x: x.label.lower()
            )

        return self.grouped_by_label, sorted_labels


def get_entities(material_id: int) -> Entities | None:
    """
    This function retrieves the entities extracted from the annotated text for the given material_id.
    Entities are stored in a JSON file named "{material_id}_annotated.json" in the "pdf_downloads" folder.
    """
    root_folder = Path("pdf_downloads")
    entities_path = root_folder / f"{material_id}_annotated.json"
    if not entities_path.exists():
        return None
    try:
        with open(entities_path, encoding="utf-8") as f:
            entities = Entities(json.load(f))
    except Exception as e:
        print(f"Error processing entities file {entities_path}: {e}")
        return None

    return entities


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


# --- component rendering ---


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

    return (
        render_contact_info(item_data),
        render_course_details(item_data),
        render_teacher_info(item_data),
    )


async def get_item_history(material_id: int):
    """
    retrieves the items edit history and parses it
    """
    history = await retrieve_item_history([material_id])

    return history


def render_item_history(history: list[ItemUpdate]):
    """
    For a given list of ItemUpdate objects, renders a card showing the changes made to the item.
    """

    edit_icon_svg = """
    <svg fill="#000000" height="15px" width="15px" version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="-117.76 -117.76 747.52 747.52" xml:space="preserve" stroke="#000000" stroke-width="0.00512"><g id="SVGRepo_bgCarrier" stroke-width="0" transform="translate(0,0), scale(1)"><rect x="-117.76" y="-117.76" width="747.52" height="747.52" rx="373.76" fill="#9bc5ad" strokewidth="0"></rect></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round" stroke="#CCCCCC" stroke-width="1.024"></g><g id="SVGRepo_iconCarrier"> <g> <g> <g> <path d="M414.682,218.274c0-4.465-3.618-8.084-8.084-8.084H212.727c-4.465,0-8.084,3.62-8.084,8.084 c0,4.465,3.62,8.084,8.084,8.084h193.871C411.063,226.358,414.682,222.738,414.682,218.274z"></path> <path d="M245.038,285.642c-4.465,0-8.084,3.62-8.084,8.084s3.62,8.084,8.084,8.084h1.078c4.465,0,8.084-3.62,8.084-8.084 s-3.62-8.084-8.084-8.084H245.038z"></path> <path d="M457.764,134.783V51.2c0-28.231-22.969-51.2-51.2-51.2H105.053c-28.231,0-51.2,22.969-51.2,51.2v212.948 C23.917,265.561,0,290.375,0,320.674c0,31.204,25.369,56.589,56.552,56.589h147.707v28.003c0,3.609,1.694,7.017,4.65,9.352 c0.001,0.001,0.002,0.001,0.002,0.001c5.026,3.97,12.608,3.97,17.637,0l41.887-33.081c3.958-3.126,6.228-7.631,6.228-12.359 s-2.269-9.233-6.228-12.359l-41.887-33.081c-5.029-3.972-12.61-3.971-17.637-0.001c-2.957,2.335-4.652,5.743-4.652,9.353v28.004 H70.021V51.2c0-19.316,15.716-35.032,35.032-35.032h301.512c19.317,0,35.032,15.716,35.032,35.032v83.537h-134.62v-28.003 c0-3.61-1.695-7.018-4.652-9.353c-5.027-3.971-12.608-3.969-17.635,0l-41.887,33.081c-3.959,3.126-6.229,7.631-6.229,12.359 s2.27,9.233,6.228,12.359l41.888,33.081c5.028,3.971,12.609,3.97,17.633,0.002c2.958-2.334,4.654-5.743,4.654-9.354v-28.004 H449.59c0.059,0.001,0.12,0.001,0.179,0h5.679c22.267,0,40.383,18.132,40.383,40.421c0,21.51-16.873,39.15-38.068,40.355v-52.751 c0-4.465-3.618-8.084-8.084-8.084c-4.466,0-8.084,3.62-8.084,8.084V460.8c0,19.316-15.715,35.032-35.032,35.032H105.053 c-19.316,0-35.032-15.716-35.032-35.032v-54.434c0-4.465-3.62-8.084-8.084-8.084c-4.465,0-8.084,3.62-8.084,8.084V460.8 c0,28.231,22.969,51.2,51.2,51.2h301.512c28.231,0,51.2-22.969,51.2-51.2V247.869C487.876,246.648,512,221.753,512,191.326 C512,160.9,487.876,136.004,457.764,134.783z M53.853,361.005c-21.015-1.394-37.684-18.95-37.684-40.332 c0-21.381,16.67-38.938,37.684-40.332V361.005z M220.427,339.509l37.569,29.67l-37.569,29.671V339.509z M290.807,172.491 l-37.569-29.67l37.569-29.671V172.491z"></path> <path d="M406.598,361.095h-96.936c-4.466,0-8.084,3.62-8.084,8.084c0,4.465,3.618,8.084,8.084,8.084h96.936 c4.466,0,8.084-3.62,8.084-8.084C414.682,364.714,411.063,361.095,406.598,361.095z"></path> <path d="M105.02,301.811h107.706c4.465,0,8.084-3.62,8.084-8.084s-3.62-8.084-8.084-8.084H105.02 c-4.465,0-8.084,3.62-8.084,8.084S100.556,301.811,105.02,301.811z"></path> <path d="M406.598,285.642h-96.936c-4.466,0-8.084,3.62-8.084,8.084s3.618,8.084,8.084,8.084h96.936 c4.466,0,8.084-3.62,8.084-8.084S411.063,285.642,406.598,285.642z"></path> <path d="M210.04,142.821c0-4.465-3.62-8.084-8.084-8.084H105.02c-4.465,0-8.084,3.62-8.084,8.084c0,4.465,3.62,8.084,8.084,8.084 h96.935C206.42,150.905,210.04,147.286,210.04,142.821z"></path> <path d="M276.273,301.811h1.077c4.466,0,8.084-3.62,8.084-8.084s-3.618-8.084-8.084-8.084h-1.077 c-4.466,0-8.084,3.62-8.084,8.084S271.807,301.811,276.273,301.811z"></path> <path d="M179.337,210.189c-4.465,0-8.084,3.62-8.084,8.084c0,4.465,3.62,8.084,8.084,8.084h1.077c4.465,0,8.084-3.62,8.084-8.084 c0-4.465-3.62-8.084-8.084-8.084H179.337z"></path> <path d="M105.02,226.358h43.082c4.465,0,8.084-3.62,8.084-8.084c0-4.465-3.62-8.084-8.084-8.084H105.02 c-4.465,0-8.084,3.62-8.084,8.084C96.936,222.738,100.556,226.358,105.02,226.358z"></path> </g> </g> </g> </g></svg>
    """
    timeline_elements = []
    if not history:
        return Div("No history found for this item.")

    print(f"Rendering item history for {len(history)} items")
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
        max_old_new_len = max(
            len(str(old)) + len(" → ") + len(str(new)) for old, new in changes.values()
        )
        edit_content = [
            f"<span class='badge badge-xs  badge-secondary text-[10px]'>{datetime.datetime.strptime(date[:19], '%Y-%m-%d %H:%M:%S')}</span>"
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

                elif col in (
                    "manual_classification",
                    "ml_prediction",
                    "classification",
                ):
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
            modal_params = {"material_id": material_id}
            row_attrs: dict[str, Any] = {
                "id": f"row-{material_id}",
                "class": "hover:bg-primary/10 cursor-pointer transition-colors duration-150",
                "hx_get": show_item_details.to(**modal_params),  # Pass only material_id
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

    # --- Define Options for Checkbox Groups (keep existing) ---
    # ... (workflow_options, status_options, etc. definitions) ...
    workflow_options = {
        "ToDo": WORKFLOW_STYLES.get("ToDo"),
        "InProgress": WORKFLOW_STYLES.get("InProgress"),
        "Done": WORKFLOW_STYLES.get("Done"),
    }
    status_options = {
        "Published": STATUS_STYLES.get("Published"),
        "Unpublished": STATUS_STYLES.get("Unpublished"),
        "Deleted": STATUS_STYLES.get("Deleted"),
    }
    classification_options = {
        **{v: LabelT.primary for v in PRIMARY_CLASSIFICATIONS},
        **{v: LabelT.secondary for v in SECONDARY_CLASSIFICATIONS},
        **{v: LabelT.destructive for v in DESTRUCTIVE_CLASSIFICATIONS},
        "None": LabelT.destructive,
    }
    manual_classification_options = classification_options
    faculty_options = {
        f: FACULTY_BADGE_STYLES.get(f, "badge-secondary")
        for f in ["BMS", "EEMCS", "ET", "ITC", "TNW"]
    }

    # --- Filter Checkbox Groups ---
    filter_checkbox_groups = [
        create_checkbox_filter_group(
            filter_key="workflow_status",
            label_text="Workflow Status",
            options=workflow_options,
            current_values=current_filters.get("workflow_status"),
            counts=filter_counts.get("workflow_status"),
            current_total=total_filtered_rows,  # <-- Use parameter here
        ),
        create_checkbox_filter_group(
            filter_key="status",
            label_text="Status",
            options=status_options,
            current_values=current_filters.get("status"),
            counts=filter_counts.get("status"),
            current_total=total_filtered_rows,  # <-- Use parameter here
        ),
        create_checkbox_filter_group(
            filter_key="classification",
            label_text="Classification",
            options=classification_options,
            current_values=current_filters.get("classification"),
            counts=filter_counts.get("classification"),
            current_total=total_filtered_rows,  # <-- Use parameter here
        ),
        create_checkbox_filter_group(
            filter_key="manual_classification",
            label_text="Manual Classification",
            options=manual_classification_options,
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
                options=faculty_options,
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
        if col not in copyright_df_global.columns:
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
                hx_post=data_grid.to(),
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
                hx_post=data_grid.to(),
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
        Button(
            UkIcon("filter"),
            "Filters",
            **{"@click": "slideOverOpen=true"},
            cls="btn btn-sm btn-outline btn-primary align-center",
        ),
        Template(x_teleport="body")(
            Div(
                x_show="slideOverOpen",
                **{"@keydown.window.escape": "slideOverOpen=false"},
                cls="relative z-[99]",
                x_cloak=True,
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

    # --- Header Structure (Unchanged) ---
    # ... (last_update_time, user info, logout button logic remains the same) ...
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
        col for col in DISPLAY_COLUMNS if col in copyright_df_global.columns
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
            hx_get=data_grid.to(),  # Target central data route
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
    and an array of directly selectable page numbers.
    first we show the |< button (skip to first page)
    then the < button (previous page)

    then numbered elements:
        - current page [highlighted]
        - max 3 elements before and after, less if close to the edges [regular]
        - if not already represented, the first 3 and last 3 pages [regular].
        - ellipsis elements if there is a gap between the current pages set and the first/last pages sets

    then the > button (next page)
    finally the >| button (skip to last page)

    example:
    |< < 1 2 3 ...  18 19 20 [!21!] 22 23 24 ... 88 89 90 > >|

    or:
    |< < 1 2 3 [!4!] 5 6 7 ... 88 89 90 > >|

    """

    if total_pages <= 1:
        return Div(cls="h-12 mb-1")  # Reserve space even if no pagination

    # init all the values we'll need

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

    # Create the list of numbers to show in the pagination element.
    # always show the current page, and 8 other pages (4 before and 4 after), and ellipses if not at the edges
    # this means max amount of these elements is 9 numbers + 2 ellipses = 11 elements
    # to prevent ux jumping around, we always show 11 elements
    # if we do not need 2 ellipses, expand the numbers on the other side to compensate and get to 11 elements
    # if we cannot get 11 elements because the total num of pages is < 9, show all pages always
    # if a symmetric set of pages is not possible because we're at an edge, expand the other side so the total is always 11

    numbers_to_add: list[int] = []

    if total_pages <= 11:
        # if we have less than 11 pages, show them all
        numbers_to_add = list(range(1, total_pages + 1))
    else:
        # if we have more than 11 pages, create the pagination elements
        # start with the current page
        numbers_to_add.append(current_page)
        # add the previous pages
        for i in range(1, 5):
            if current_page - i > 0:
                numbers_to_add.append(current_page - i)
        # add the next pages
        for i in range(1, 5):
            if current_page + i <= total_pages:
                numbers_to_add.append(current_page + i)

        # now sort the list and remove duplicates
        numbers_to_add = sorted(set(numbers_to_add))

        # check if we need to add ellipses
        if numbers_to_add[0] > 1:
            numbers_to_add.insert(0, "...")
        if numbers_to_add[-1] < total_pages:
            numbers_to_add.append("...")
        # now we need to check if we have 11 elements, if not, add the missing elements
        if len(numbers_to_add) < 11:
            missing_amount = 11 - len(numbers_to_add)
            print(f"count is {len(numbers_to_add)}, missing {missing_amount}")
            # we need to add elements to the left or right side
            # if we have ellipses, add to the other side
            print(numbers_to_add)
            if isinstance(numbers_to_add[-1], str) or (
                isinstance(numbers_to_add[0], int)
                and numbers_to_add[0] <= missing_amount
            ):
                print("adding to right side")
                # add to the right side
                if numbers_to_add[-1] == "...":
                    numbers_to_add.pop()
                start_num = numbers_to_add[-1]
                for i in range(1, missing_amount + 1):
                    if start_num + i <= total_pages:
                        print(f"adding {start_num + i} to right side")
                        numbers_to_add.append(start_num + i)
            elif isinstance(numbers_to_add[0], str) or (
                isinstance(numbers_to_add[-1], int)
                and numbers_to_add[-1] >= total_pages - missing_amount
            ):
                print("adding to left side")
                # add to the left side
                if numbers_to_add[0] == "...":
                    numbers_to_add.pop(0)
                start_num = numbers_to_add[0]
                for i in range(1, missing_amount + 1):
                    if start_num - i > 0:
                        print(f"adding {start_num - i} to left side")
                        numbers_to_add.insert(0, start_num - i)

        if isinstance(numbers_to_add[0], int) and numbers_to_add[0] > 1:
            numbers_to_add.insert(0, "...")
        if isinstance(numbers_to_add[-1], int) and numbers_to_add[-1] < total_pages:
            numbers_to_add.append("...")

    print(
        f"{len(numbers_to_add)} numbers that will be used for pagination display:{numbers_to_add}"
    )

    # now create the elements for the button block
    btn_block = []
    # 1. create first + prev buttons
    attr_dict = {}
    attr_dict["first"] = (
        "⏮",
        {
            "hx_get": data_grid.to(),  # Target the central data route
            "hx_target": "#data-grid-component",
            "hx_vals": json.dumps({"page": 1}),  # Send only the page change
            "role": "button",
            "cls": classes["btn_disabled"]
            if current_page <= 1
            else classes["btn_enabled"],
        },
    )
    attr_dict["prev"] = (
        "⏴",
        {
            "hx_get": data_grid.to(),  # Target the central data route
            "hx_target": "#data-grid-component",
            "hx_vals": json.dumps(
                {"page": max(1, current_page - 1)}
            ),  # Send only the page change
            "role": "button",
            "cls": classes["btn_disabled"]
            if current_page <= 1
            else classes["btn_enabled"],
        },
    )
    # add the buttons to the block
    for key, (text, attr) in attr_dict.items():
        btn_block.append(
            A(
                text,
                **attr,
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
                    hx_get=data_grid.to(),
                    hx_target="#data-grid-component",
                    hx_vals=json.dumps({"page": number}),
                    role="button",
                    cls=cls,
                ),
                cls="hidden h-full md:block",
            )
        )
    # 3. create next + last buttons
    attr_dict = {}
    attr_dict["next"] = (
        "⏵",
        {
            "hx_get": data_grid.to(),  # Target the central data route
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
            "hx_get": data_grid.to(),  # Target the central data route
            "hx_target": "#data-grid-component",
            "hx_vals": json.dumps({"page": total_pages}),  # Send only the page change
            "role": "button",
            "cls": classes["btn_disabled"]
            if current_page >= total_pages
            else classes["btn_enabled"],
        },
    )
    # add the buttons to the block
    for key, (text, attr) in attr_dict.items():
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
    elif col_name == "workflow_status":
        style = WORKFLOW_STYLES.get(val_str, DEFAULT_PILL_STYLE)
        return Label(display_text, cls=style + " badge-sm"), "label"
    elif col_name == "status":
        style = STATUS_STYLES.get(val_str, DEFAULT_PILL_STYLE)
        return Label(display_text, cls=style + " badge-sm"), "label"
    elif col_name in ("classification", "manual_classification", "ml_prediction"):
        if not val_str_norm:
            display_text = "-"
            style = ""
        elif val_str_norm in PRIMARY_CLASSIFICATIONS:
            style = LabelT.primary
            display_text = str(val_str_norm)
        elif val_str_norm in SECONDARY_CLASSIFICATIONS:
            style = LabelT.secondary
            display_text = str(val_str_norm)
        elif val_str_norm in DESTRUCTIVE_CLASSIFICATIONS:
            style = LabelT.destructive
            display_text = str(val_str_norm)
        else:
            display_text = str(val_str_norm)
            style = ""

        return Label(display_text, cls=style + " badge-sm"), "label"

    elif value is None:
        return Span("N/A", cls="text-base-content/70 text-sm"), "span"
    elif isinstance(value, int | float):
        return Label(value, cls=style + " badge-sm"), "label"
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


# --- main page routes ---
def _apply_filters_for_count(filters: dict[str, str], global_df: pl.DataFrame) -> int:
    """
    Helper to apply filters to the provided global DataFrame and return the count.
    Optimized for counting.

    Args:
        filters: Dictionary of filters to apply.
        global_df: The complete, unfiltered DataFrame to filter on.

    Returns:
        The count of rows matching the filters.
    """
    df = global_df  # Use the passed global DataFrame

    if not filters:
        return df.height  # Return total count if no filters

    filter_expressions = []
    # Re-use the filter expression building logic from get_filtered_sorted_df
    # This ensures consistency in how filters are interpreted for counting vs. display
    for col, value in filters.items():
        actual_col = col
        if actual_col not in df.columns:
            continue  # Skip if column doesn't exist
        if value is None or value == "":
            continue

        # --- Start copied/adapted logic ---
        try:
            or_values = (
                value.split("|") if isinstance(value, str) and "|" in value else [value]
            )
            or_expressions = []
            for or_value in or_values:
                or_value = or_value.strip()
                if not or_value:
                    continue

                if (
                    col in ["manual_classification", "classification"]
                    and or_value.lower() == "none"
                ):
                    or_expressions.append(
                        (pl.col(actual_col).is_null())
                        | (pl.col(actual_col) == "")
                        | (pl.col(actual_col) == "-")
                    )
                    continue

                if df[actual_col].dtype == pl.Utf8:
                    if col in [
                        "status",
                        "workflow_status",
                        "classification",
                        "manual_classification",
                        "faculty",
                    ]:
                        or_expressions.append(
                            pl.col(actual_col).str.to_lowercase() == or_value.lower()
                        )
                    else:
                        or_expressions.append(
                            pl.col(actual_col).str.contains(
                                f"(?i){re.escape(or_value)}"
                            )
                        )
                elif df[actual_col].dtype in (
                    pl.Int64,
                    pl.Int32,
                    pl.Float64,
                    pl.Float32,
                ):
                    with contextlib.suppress(ValueError):
                        or_expressions.append(pl.col(actual_col) == float(or_value))

            if len(or_expressions) > 1:
                filter_expressions.append(pl.any_horizontal(or_expressions))
            elif len(or_expressions) == 1:
                filter_expressions.append(or_expressions[0])
        except Exception as e:
            print(f"Count Filter warning on '{actual_col}'='{value}': {e}")
        # --- End copied/adapted logic ---

    if filter_expressions:
        try:
            # Apply filters and get count directly using lazy evaluation
            count = (
                df.lazy()
                .filter(pl.all_horizontal(filter_expressions))
                .select(pl.count())
                .collect()
                .item()
            )
            return count
        except Exception as e:
            print(f"Count Filter error applying filters {filters}: {e}")
            return 0  # Return 0 on error
    else:
        # No valid filters were generated, or filters were empty
        return df.height


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

    # Define the structure for rendering checkbox groups
    # Assumes options maps like workflow_options are defined globally
    workflow_options = {
        "ToDo": WORKFLOW_STYLES.get("ToDo"),
        "InProgress": WORKFLOW_STYLES.get("InProgress"),
        "Done": WORKFLOW_STYLES.get("Done"),
    }
    status_options = {
        "Published": STATUS_STYLES.get("Published"),
        "Unpublished": STATUS_STYLES.get("Unpublished"),
        "Deleted": STATUS_STYLES.get("Deleted"),
    }
    classification_options = {
        **{v: LabelT.primary for v in PRIMARY_CLASSIFICATIONS},
        **{v: LabelT.secondary for v in SECONDARY_CLASSIFICATIONS},
        **{v: LabelT.destructive for v in DESTRUCTIVE_CLASSIFICATIONS},
        "None": LabelT.destructive,
    }
    manual_classification_options = classification_options  # Reuse same options
    faculty_options = {
        f: FACULTY_BADGE_STYLES.get(f, "badge-secondary")
        for f in ["BMS", "EEMCS", "ET", "ITC", "TNW"]
    }

    checkbox_groups_to_render = [
        ("workflow_status", "Workflow Status", workflow_options),
        ("status", "Status", status_options),
        ("classification", "Classification", classification_options),
        (
            "manual_classification",
            "Manual Classification",
            manual_classification_options,
        ),
    ]
    # Conditionally add faculty group if user is admin or has no specific faculty assigned
    if is_admin or not user_faculty or user_faculty == "all":
        checkbox_groups_to_render.append(("faculty", "Faculty", faculty_options))

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

        initial_counts_for_empty = calculate_counts_for_ui(
            processed_data.app_state, copyright_df_global, None
        )
        _update_empty_filters(initial_counts_for_empty)

        header_component = page_header_component(
            auth_details,
            processed_data.app_state,
            processed_data.filter_counts,
            processed_data.total_filtered_rows,
        )

        modal_placeholder = Dialog(
            id="modal-placeholder", cls="modal modal-bottom sm:modal-middle"
        )
        return (
            Title("CDD//UT Dashboard"),
            MODAL_TRIGGER,
            MODAL_INTERACTION,
            Div(
                Div(header_component, grid_component, id="content-area"),
                id="page-container",
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

    def create_editable_pill_div(label_text: str, field_name: str, options_map: dict):
        # ... (Keep existing implementation)
        current_value = get_val(field_name)
        content_component, html_tag = render_modal_field(field_name, current_value)
        label_el, _ = render_labelled_item(label_text, content_component, html_tag)
        original_value_str = str(current_value) if current_value is not None else ""
        original_style_class = str(DEFAULT_PILL_STYLE)  # Default style

        # Logic to find the *actual* style class applied by render_modal_field
        if hasattr(content_component, "attrs") and "cls" in content_component.attrs:
            current_classes = content_component.attrs["cls"].split()
            labelt_values = (
                {str(lt) for lt in LabelT}
                | {str(st) for st in STATUS_STYLES.values()}
                | {str(wt) for wt in WORKFLOW_STYLES.values()}
            )
            found_style = next(
                (cls for cls in current_classes if cls in labelt_values), None
            )
            if found_style:
                original_style_class = found_style
            elif (
                "badge-sm" not in current_classes
                and isinstance(content_component, FT)
                and (
                    content_component.tag == "span"
                    and "uk-label" in content_component.attrs.get("cls", "")
                )
            ):
                content_component.attrs["cls"] += " badge-sm"

        # Add ID to the display component for JS targeting
        component_with_id = content_component
        # Safely add/update the ID
        if hasattr(component_with_id, "attrs"):
            component_with_id.attrs = (
                component_with_id.attrs.copy()
                if isinstance(component_with_id.attrs, dict)
                else {}
            )
            component_with_id.attrs["id"] = f"pill-display-{field_name}"
        # else: # Handle cases where it might be a simple string or needs wrapping
        # component_with_id = Span(content_component, id=f"pill-display-{field_name}") # Example wrap

        dropdown_items = []
        for opt_val, opt_style_enum in options_map.items():
            opt_style_class = (
                str(opt_style_enum) if opt_style_enum else str(DEFAULT_PILL_STYLE)
            )
            # Ensure opt_val is treated as a string for JS
            js_opt_val = json.dumps(
                str(opt_val)
            )  # Use json.dumps for safe JS string representation
            js_opt_text = json.dumps(str(opt_val))  # Text is usually same as value here
            onclick_js = f"updatePill('{field_name}', {js_opt_val}, {js_opt_text}, '{opt_style_class}'); return false;"
            # Ensure the label inside the dropdown also has badge-sm
            dropdown_items.append(
                Li(
                    A(
                        Label(opt_val, cls=f"{opt_style_class} badge-sm"),
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
            data_original_text=original_value_str,  # Use same as value for simplicity here
            data_original_style=original_style_class,
        )

        pill_container = Div(
            component_with_id,  # The component with the ID set
            Div(
                Ul(*dropdown_items, cls="uk-nav uk-dropdown-nav"),
                cls="uk-dropdown w-auto bg-base-100 p-2 shadow-lg rounded-md border border-base-300",
                uk_drop="mode: click; pos: bottom-right; boundary: !.modal-box; flip: false",
            ),
            hidden_input,
            cls="inline-block uk-inline",
        )

        return Div(
            label_el,
            Div(pill_container, cls="text-right"),
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
        nested_data_list = retrieve_osiris_data([material_id])
        if not nested_data_list:
            raise ValueError(
                f"Material ID {material_id} not found in detailed data source."
            )
        item_data = nested_data_list[0]

        if "faculty_id" in item_data and "faculty" not in item_data:
            item_data["faculty"] = item_data.pop("faculty_id")

        manual_classification_options = {}
        for val in PRIMARY_CLASSIFICATIONS:
            manual_classification_options[val] = LabelT.primary
        for val in SECONDARY_CLASSIFICATIONS:
            manual_classification_options[val] = LabelT.secondary
        for val in DESTRUCTIVE_CLASSIFICATIONS:
            manual_classification_options[val] = LabelT.destructive
        if "onbekend" not in manual_classification_options:
            manual_classification_options["onbekend"] = LabelT.secondary
        manual_classification_options["None"] = LabelT.destructive

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
            create_editable_pill_div(
                "Workflow Status", "workflow_status", WORKFLOW_STYLES
            ),
            create_editable_pill_div(
                "Manual Classification",
                "manual_classification",
                manual_classification_options,
            ),
            Div(
                Strong(
                    "Remarks", cls="block text-xs font-medium text-base-content/80 mb-1"
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
                cls="mb-3",
            ),
            Div(cls="flex justify-end space-x-2 mt-4")(
                Button(
                    "Reset",
                    type="button",
                    cls=ButtonT.secondary + " btn-sm",
                    onclick="resetModalForm(); return false;",
                ),
                Button(
                    Span(
                        Span(cls="relative flex size-3 mr-2")(
                            Span(
                                cls="absolute inline-flex h-full w-full animate-ping rounded-full bg-red-400 opacity-75"
                            ),
                            UkIcon(
                                "alert-triangle",
                                cls="relative inline-flex size-3 text-red-500",
                            ),
                        ),
                        id="save-indicator",
                        cls="hidden",
                    ),
                    "Save",
                    id="modal-save-btn",
                    type="submit",
                    cls=ButtonT.primary + " btn-sm",
                    disabled=True,
                ),
            ),
        )
        # --- Item Info Card ---
        item_info_content = (
            create_readonly_item_div(
                "Classification", "classification", is_inline=True
            ),
            create_readonly_item_div("ML Prediction", "ml_prediction", is_inline=True),
            create_readonly_item_div("Period", "period"),
            create_readonly_item_div("Faculty", "faculty"),
            create_readonly_item_div("Owner", "owner"),
            create_readonly_item_div("Department", "department"),
            create_readonly_item_div("Course Name", "course_name"),
            create_readonly_item_div("Course Code", "course_code"),
        )
        # --- Text Details Card ---
        text_details_content = (
            create_readonly_item_div("Title", "title"),
            create_readonly_item_div("Author", "author"),
            create_readonly_item_div("Publisher", "publisher"),
            create_readonly_item_div("DOI", "doi"),
            create_readonly_item_div("ISBN", "isbn"),
        )
        # --- Numeric Details Card ---
        numeric_details_content = (
            create_readonly_item_div("Pages", "pagecount", is_inline=True),
            create_readonly_item_div("Words", "wordcount", is_inline=True),
            create_readonly_item_div("Pictures", "picturecount", is_inline=True),
        )
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
                Div(header_content, cls="border-b pb-2 flex-shrink-0"),  # Header
                Div(  # Scrollable Body
                    # Form for saving details
                    Form(
                        Input(
                            type="hidden", name="material_id", value=material_id
                        ),  # Essential ID
                        # REMOVED hidden state inputs
                        modal_cards_grid,  # The main content grid (includes editable fields)
                        id="modal-details-form",
                        hx_post=save_item_details.to(),  # Target the save endpoint
                        hx_target="body",  # Target body to receive OOB swaps from save endpoint
                        hx_swap="none",  # Let OOB swaps handle UI updates
                    ),
                    cls="relative py-4 flex-grow overflow-y-auto",
                ),
                footer_content,  # Footer with prev/next/close
                cls="flex-shrink-0",  # Ensure footer doesn't scroll
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
                Img(src="imgs/ut_logo.png"),
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


# --- main entry point ---
if __name__ == "__main__":
    print("Starting FastHTML server...")
    serve(port=PORT, reload=True)
