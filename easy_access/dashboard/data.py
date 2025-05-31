"""
This module handles data processing, state management, and data retrieval
specifically for the web dashboard. It includes functions for:
- Managing the application state (filters, sorting, pagination) based on user interactions.
- Fetching and filtering copyright item data from the database using Polars.
- Calculating counts for UI filter elements.
- Storing changes made by users via the dashboard back to the database.
"""
import contextlib
import logging
import math
import re
from copy import copy, deepcopy
from dataclasses import asdict, dataclass
from functools import lru_cache
from typing import Any, Dict, List, Tuple, Optional # Use specific Dict, List, Tuple

import polars as pl
# from rich import print # Removed, using logging

from easy_access.dashboard.constants import (
    BADGE_STYLES,
    FILTERABLE_COLUMNS, # List of column names that can be filtered
)
from easy_access.dashboard.web import AppState, load_app_state # AppState for state management
from easy_access.db.retrieve import retrieve_copyright_items, retrieve_item_history
from easy_access.db.update import update_copyright_items

logger = logging.getLogger(__name__)

# Global variable to cache columns that have no filterable values in the current dataset.
# This is initialized/updated by _update_empty_filters, called from get_filter_counts.
EMPTY_FILTERS: Dict[str, List[str]] = {}


@dataclass
class ProcessedDataResult:
    """
    Holds the results after processing application state and fetching data for the dashboard.

    Attributes:
        app_state (AppState): The validated AppState reflecting the current view (filters, page, etc.).
        df_slice (pl.DataFrame): The slice of data to be displayed on the current page.
        total_filtered_rows (int): Total number of rows matching the current filters (before pagination).
        total_pages (int): Total number of pages available for the filtered data.
        filter_counts (Dict[str, Dict[str, int | str]]): Counts for each option within filterable columns,
                                                       used to populate UI filter elements. Values can be int or "-"
    """
    app_state: AppState
    df_slice: pl.DataFrame
    total_filtered_rows: int
    total_pages: int
    filter_counts: Dict[str, Dict[str, Union[int, str]]] # type: ignore


def process_state(
    session: Dict[str, Any], request_params: Dict[str, Any]
) -> Tuple[Dict[str, Any], AppState]:
    """
    Loads the current application state from the session, applies changes based on
    request parameters (e.g., user interactions like sorting, filtering, pagination),
    saves the modified state back to the session if changed, and returns both the
    updated session and the new AppState.

    This function does not perform data-dependent validation (like page number bounds).

    Args:
        session (Dict[str, Any]): The current user session dictionary.
        request_params (Dict[str, Any]): Dictionary of parameters from the HTTP request.

    Returns:
        Tuple[Dict[str, Any], AppState]: The (potentially updated) session dictionary
                                         and the AppState instance reflecting the requested changes.
    """
    logger.debug(f"Processing state with request_params: {request_params}")

    app_state: AppState = load_app_state(session)
    initial_state_as_dict: Dict[str, Any] = asdict(app_state) # For detecting if state changed

    filter_action_occurred: bool = False

    action_param = request_params.get("action")
    if action_param == "reset":
        if app_state.filters or app_state.page != 1: # Only reset if there's something to reset
            app_state.clear_filters()
            app_state.set_page(1) # Reset to first page
            filter_action_occurred = True
            logger.debug("State reset: Filters cleared and page set to 1.")
        else:
            logger.debug("Reset action requested, but no active filters or non-first page. No state change.")

    # Handle specific parameter changes if no 'reset' action
    elif "page" in request_params:
        try:
            new_page = int(request_params["page"])
            if new_page != app_state.page:
                app_state.set_page(new_page)
                logger.debug(f"Page changed to: {new_page}")
        except (ValueError, TypeError):
            logger.warning(f"Invalid page parameter '{request_params['page']}'. Ignoring.")

    elif "sort_by" in request_params:
        new_sort_by = request_params.get("sort_by")
        # sort_desc can come as 'true'/'false' string from JS/HTMX
        new_sort_desc_str = str(request_params.get("sort_desc", "False")).lower()
        new_sort_desc = new_sort_desc_str == "true"

        if new_sort_by != app_state.sort_by or new_sort_desc != app_state.sort_desc:
            app_state.set_sort(new_sort_by if new_sort_by else None, new_sort_desc)
            app_state.set_page(1) # Reset to first page on sort change
            logger.debug(f"Sort changed to: by '{new_sort_by}', desc: {new_sort_desc}. Page reset.")

    elif "per_page" in request_params:
        try:
            new_per_page = int(request_params["per_page"])
            if new_per_page != app_state.per_page:
                app_state.set_per_page(new_per_page)
                app_state.set_page(1)  # Reset page when items per page changes
                logger.debug(f"Items per page changed to: {new_per_page}. Page reset.")
        except (ValueError, TypeError):
            logger.warning(f"Invalid per_page parameter '{request_params['per_page']}'. Ignoring.")

    else: # Check for filter changes (checkboxes or text inputs)
        # This block assumes filter changes are the only other type of state modification
        is_checkbox_filter_change = "checked" in request_params # Specific param from checkbox toggle
        current_active_filters = app_state.get_active_filters()

        for param_key, param_value in request_params.items():
            if param_key.startswith("filter_"):
                filter_field_name = param_key[len("filter_"):] # e.g., "workflow_status"
                if filter_field_name not in FILTERABLE_COLUMNS:
                    logger.debug(f"Skipping unknown filter key '{filter_field_name}' from request parameters.")
                    continue

                current_filter_value_str = current_active_filters.get(filter_field_name, "")

                if is_checkbox_filter_change:
                    is_checked_bool = str(request_params["checked"]).lower() == "true"
                    # param_value here is the value of the specific checkbox that was clicked
                    clicked_checkbox_value = str(param_value)

                    # Simulate the change to see if it's different from current state for this filter_field_name
                    temp_current_values_set = set(current_filter_value_str.split("|")) if current_filter_value_str else set()
                    if is_checked_bool: temp_current_values_set.add(clicked_checkbox_value)
                    else: temp_current_values_set.discard(clicked_checkbox_value)
                    temp_current_values_set.discard("") # Clean up

                    simulated_new_value_str = "|".join(sorted(list(temp_current_values_set))) if temp_current_values_set else ""

                    if simulated_new_value_str != current_filter_value_str:
                        if app_state.apply_filter_change(filter_field_name, clicked_checkbox_value, is_checked_bool):
                            filter_action_occurred = True
                        break # Process only one checkbox change at a time

                else: # Text filter change
                    new_text_filter_value = str(param_value)
                    if new_text_filter_value != current_filter_value_str:
                        if app_state.apply_text_filter(filter_field_name, new_text_filter_value):
                            filter_action_occurred = True
                        break # Process only one text filter change at a time

        if filter_action_occurred and "per_page" not in request_params: # Avoid double page reset
            app_state.set_page(1)
            logger.debug("Filter changed. Page reset to 1.")

    # Save state back to session if it has changed
    final_state_as_dict = asdict(app_state)
    if final_state_as_dict != initial_state_as_dict:
        logger.info("Application state changed. Saving updated state to session.")
        session["app_state"] = final_state_as_dict
    else:
        logger.debug("No state change detected after processing request parameters.")

    return session, app_state


async def fetch_data(app_state: AppState, session: Dict[str, Any]) -> ProcessedDataResult:
    """
    Fetches and processes data based on the provided application state and user authentication.

    This involves:
    1. Applying faculty constraints based on user role from the session.
    2. Filtering and sorting the main dataset (`get_item_df()`) via `get_filtered_sorted_df`.
    3. Calculating total pages and validating the current page number in `app_state`.
       A *copy* of app_state is validated and returned in ProcessedDataResult.
    4. Calculating counts for UI filter options via `calculate_counts_for_ui`.
    5. Slicing the filtered data for the current page.

    Args:
        app_state (AppState): The current application state (filters, sort, pagination).
        session (Dict[str, Any]): The user session dictionary, used for auth details.

    Returns:
        ProcessedDataResult: An object containing the data slice for the current page,
                             pagination details, filter counts, and the *validated* AppState.
    """
    logger.debug(f"Fetching data for AppState: Page {app_state.page}, Filters {app_state.filters}, Sort {app_state.sort_by} {app_state.sort_desc}")

    auth_details: Dict[str, Any] = session.get("auth", {})
    faculty_constraint: Optional[Dict[str, str]] = None
    user_faculty: Optional[str] = auth_details.get("faculty")
    user_role: Optional[str] = auth_details.get("role")

    if user_role != "admin" and user_faculty and user_faculty != "all":
        faculty_constraint = {"faculty": user_faculty}
        logger.info(f"Applying faculty constraint for user: {user_faculty}")

    # Get all data, filtered and sorted according to current app_state and constraints
    # This uses the AppState *before* page validation.
    all_filtered_df: pl.DataFrame = get_filtered_sorted_df(app_state, extra_constraints=faculty_constraint)
    total_filtered_rows: int = all_filtered_df.height
    logger.debug(f"Total rows after filtering: {total_filtered_rows}")

    # Validate AppState (especially page number) against the filtered data.
    # Use a deepcopy to avoid modifying the original app_state that might be used for filter counts.
    validated_app_state = deepcopy(app_state)

    total_pages: int = 1
    if validated_app_state.per_page > 0:
        total_pages = math.ceil(total_filtered_rows / validated_app_state.per_page)
    total_pages = max(1, total_pages) # Ensure at least 1 page

    # Correct current page if it's out of bounds
    validated_app_state.page = max(1, min(validated_app_state.page, total_pages))
    if validated_app_state.page != app_state.page:
        logger.info(f"Page number validated and changed from {app_state.page} to {validated_app_state.page} (Total pages: {total_pages})")

    # Calculate counts for UI filter display elements.
    # This should use the *original* app_state filters to show counts relative to *current* filter selection.
    # And it needs the faculty constraint.
    ui_filter_counts: Dict[str, Dict[str, Union[int, str]]] = calculate_counts_for_ui(
        app_state, faculty_constraint=faculty_constraint
    )

    # Slice the data for the current (validated) page
    offset: int = (validated_app_state.page - 1) * validated_app_state.per_page
    current_page_df_slice: pl.DataFrame = all_filtered_df.slice(offset, validated_app_state.per_page)
    logger.debug(f"Data slice for page {validated_app_state.page}: {current_page_df_slice.height} rows.")

    return ProcessedDataResult(
        app_state=validated_app_state, # Return the validated state
        df_slice=current_page_df_slice,
        total_filtered_rows=total_filtered_rows,
        total_pages=total_pages,
        filter_counts=ui_filter_counts,
    )


def calculate_counts_for_ui(
    app_state: AppState,
    faculty_constraint: Optional[Dict[str, str]]
) -> Dict[str, Dict[str, Union[int, str]]]:
    """
    Calculates counts for each option of each filterable column to display in the UI.
    The counts reflect how many items would match if that *specific* option were
    added to (or is part of) the current set of filters.

    Args:
        app_state (AppState): The current application state reflecting user's filter choices.
        faculty_constraint (Optional[Dict[str, str]]): Faculty constraint to apply to all counts.

    Returns:
        Dict[str, Dict[str, Union[int, str]]]:
            A nested dictionary: {filter_key: {option_value: count_or_dash, ...}, ...}.
            Count is an integer, or "-" if the option is part of EMPTY_FILTERS (globally no items)
            or if the option is already selected (count implicitly shown by total items).
    """
    # Get the DataFrame reflecting current filters + faculty constraint
    # This df is used to determine the 'base' count when an option is already selected.
    base_filtered_df_lazy = get_filtered_sorted_df(app_state, extra_constraints=faculty_constraint).lazy()
    count_if_option_already_selected = base_filtered_df_lazy.select(pl.len()).collect().item()


    # Initialize EMPTY_FILTERS if it's the first run or needs refresh.
    # This should ideally be done once or when global data changes significantly.
    # It finds filter options that have zero items in the *unfiltered* (but faculty-constrained) dataset.
    if not EMPTY_FILTERS:
        temp_empty_state = AppState() # Fresh state with no filters
        _update_empty_filters(temp_empty_state, faculty_constraint) # Pass constraint here too

    current_active_filters: Dict[str, str] = app_state.get_active_filters()
    calculated_counts: Dict[str, Dict[str, Union[int, str]]] = {}

    for filter_col_name in FILTERABLE_COLUMNS:
        if filter_col_name not in get_item_df().columns: # Ensure column exists in base data
            continue

        options_for_column: Dict[str, Any] # From BADGE_STYLES
        if filter_col_name == "manual_classification" and "classification" in BADGE_STYLES: # Special case
            options_for_column = BADGE_STYLES["classification"]
        elif filter_col_name in BADGE_STYLES:
            options_for_column = BADGE_STYLES[filter_col_name]
        else: # No predefined badge styles/options for this filterable column, attempt to get unique values
            # This path might be slow if not cached. For now, assume BADGE_STYLES cover options.
            logger.debug(f"No predefined badge styles for {filter_col_name}, counts might be incomplete for UI options.")
            options_for_column = {} # Or derive from data if necessary: {val:val for val in get_item_df().get_column(filter_col_name).unique().drop_nulls().to_list()}


        calculated_counts[filter_col_name] = {}
        current_selections_for_this_filter_key: Set[str] = set(
            current_active_filters.get(filter_col_name, "").split("|")
        ) if current_active_filters.get(filter_col_name) else set()

        for option_val_str in options_for_column.keys():
            if option_val_str in EMPTY_FILTERS.get(filter_col_name, []):
                calculated_counts[filter_col_name][option_val_str] = "-" # Globally no items for this option
                continue

            if option_val_str in current_selections_for_this_filter_key:
                # If option is already selected, its effect is part of the main count.
                # UI might show total, or use "-" to indicate "active".
                calculated_counts[filter_col_name][option_val_str] = count_if_option_already_selected # Or "-"
            else:
                # Simulate adding this option to current filters
                temp_simulated_filters = current_active_filters.copy()
                new_filter_val_parts = current_selections_for_this_filter_key.copy()
                new_filter_val_parts.add(option_val_str)
                temp_simulated_filters[filter_col_name] = "|".join(sorted(list(new_filter_val_parts)))

                # Calculate count with this simulated filter state
                count = _apply_filters_for_count(temp_simulated_filters, faculty_constraint)
                calculated_counts[filter_col_name][option_val_str] = count

    return calculated_counts


def _update_empty_filters(app_state_for_empty_calc: AppState, faculty_constraint: Optional[Dict[str,str]]) -> None:
    """
    Identifies filter options that have zero matching items in the (potentially faculty-constrained) global dataset.
    Updates the global `EMPTY_FILTERS` dictionary. This is used by `calculate_counts_for_ui`
    to mark such options as "-" (not available).

    Args:
        app_state_for_empty_calc (AppState): An AppState, typically with no user filters applied,
                                            to calculate baseline counts against.
        faculty_constraint (Optional[Dict[str,str]]): Faculty constraint to apply.
    """
    global EMPTY_FILTERS
    logger.debug("Updating EMPTY_FILTERS cache...")
    # Calculate counts with NO user filters, but WITH faculty constraint if any
    base_df_for_empty_calc = get_filtered_sorted_df(app_state_for_empty_calc, extra_constraints=faculty_constraint)

    temp_empty_filters: Dict[str, List[str]] = {}
    for filter_col_name in FILTERABLE_COLUMNS:
        if filter_col_name not in base_df_for_empty_calc.columns:
            continue

        options_for_column_keys: List[str]
        if filter_col_name == "manual_classification" and "classification" in BADGE_STYLES:
            options_for_column_keys = list(BADGE_STYLES["classification"].keys())
        elif filter_col_name in BADGE_STYLES:
            options_for_column_keys = list(BADGE_STYLES[filter_col_name].keys())
        else:
            options_for_column_keys = base_df_for_empty_calc.get_column(filter_col_name).unique().drop_nulls().cast(pl.Utf8).to_list()

        current_col_empty_options: List[str] = []
        for option_val in options_for_column_keys:
            # Count how many items have this option_val for this filter_col_name in the base_df
            # This is complex as _create_filter_expression handles "None" strings etc.
            # A simpler way: group by the column and get counts.
            if base_df_for_empty_calc.is_empty(): # If base df is empty, all options are "empty"
                 current_col_empty_options.append(option_val)
                 continue

            try:
                # Create a temporary filter for just this option to count against base_df
                temp_filter_for_option = {filter_col_name: str(option_val)}
                count_expr = _create_filter_expression(base_df_for_empty_calc, temp_filter_for_option)
                option_count = base_df_for_empty_calc.lazy().filter(count_expr).select(pl.len()).collect().item()
                if option_count == 0:
                    current_col_empty_options.append(option_val)
            except Exception as e:
                logger.warning(f"Error calculating empty filter status for {filter_col_name}={option_val}: {e}")

        if current_col_empty_options:
            temp_empty_filters[filter_col_name] = current_col_empty_options

    EMPTY_FILTERS = temp_empty_filters
    logger.debug(f"EMPTY_FILTERS updated: {EMPTY_FILTERS}")


@lru_cache(maxsize=1) # Cache the main DataFrame to avoid repeated DB calls
def get_item_df() -> pl.DataFrame:
    """
    Retrieves and caches the main copyright items DataFrame from the database.
    Uses `lru_cache(maxsize=1)` to cache the result. Call `refresh_copyright_df()`
    to clear the cache and re-fetch.

    Returns:
        pl.DataFrame: The cached DataFrame of all copyright items.
    """
    logger.info("Fetching main copyright items DataFrame from database (or cache)...")
    return retrieve_copyright_items()


def refresh_copyright_df() -> pl.DataFrame:
    """
    Clears the cache for `get_item_df()` and re-fetches the main copyright items DataFrame.

    Returns:
        pl.DataFrame: The newly fetched DataFrame.
    """
    logger.info("Refreshing main copyright items DataFrame cache...")
    get_item_df.cache_clear()
    # After clearing cache, update EMPTY_FILTERS based on the new global data
    # This requires an AppState, use a default one. Also needs faculty constraint if applicable.
    # This is tricky because faculty_constraint isn't globally available here.
    # EMPTY_FILTERS might need to be less global or updated more contextually.
    # For now, let's assume it's cleared and will be repopulated on next calculate_counts_for_ui.
    global EMPTY_FILTERS
    EMPTY_FILTERS = {}
    logger.info("EMPTY_FILTERS cache cleared due to data refresh.")
    return get_item_df()


def _create_filter_expression(
    df_schema_provider: pl.DataFrame, # Used only for schema, not data
    user_filters: Dict[str, str],
    extra_constraints: Optional[Dict[str, str]] = None,
) -> pl.Expr:
    """
    Creates a Polars filter expression from user-supplied filters and extra constraints.

    Args:
        df_schema_provider (pl.DataFrame): A DataFrame instance used to check column existence and types.
                                          Does not filter the data from this DataFrame directly.
        user_filters (Dict[str, str]): Filters from user input (AppState). Keys are column names,
                                     values are filter strings (pipe-separated for OR).
        extra_constraints (Optional[Dict[str, str]]): Additional fixed constraints (e.g., faculty).

    Returns:
        pl.Expr: A Polars expression that can be used in a `.filter()` call.
                 Returns `pl.lit(True)` if no valid filters/constraints are generated.
    """
    filter_expressions_list: List[pl.Expr] = []

    # Process user_filters
    for col_name, filter_value_str in user_filters.items():
        if col_name not in df_schema_provider.columns:
            logger.debug(f"Filter column '{col_name}' not in DataFrame schema. Skipping.")
            continue
        if not filter_value_str:  # Skip empty filter strings
            continue

        or_value_parts = filter_value_str.split("|")
        per_column_or_expressions: List[pl.Expr] = []
        for or_val_part in or_value_parts:
            or_val_cleaned = or_val_part.strip()
            if not or_val_cleaned: continue

            # Specific handling for 'None' string in classification-like columns
            if col_name in ["manual_classification", "classification"] and or_val_cleaned.lower() == "none":
                per_column_or_expressions.append(
                    (pl.col(col_name).is_null()) | (pl.col(col_name) == "") | (pl.col(col_name) == "-")
                )
            # General string column handling
            elif df_schema_provider[col_name].dtype == pl.Utf8:
                if col_name in ["status", "workflow_status", "classification", "manual_classification", "faculty"]:
                    # Exact match for categorical-like string fields (case-insensitive)
                    per_column_or_expressions.append(pl.col(col_name).str.to_lowercase() == or_val_cleaned.lower())
                else:
                    # Contains match for other text search fields (case-insensitive regex)
                    per_column_or_expressions.append(pl.col(col_name).str.contains(f"(?i){re.escape(or_val_cleaned)}"))
            # Example for numeric (adapt if other numeric types or precise comparison needed)
            elif df_schema_provider[col_name].dtype in (pl.Int64, pl.Int32, pl.Float64, pl.Float32):
                with contextlib.suppress(ValueError): # Ignore if or_val_cleaned is not a valid number
                    per_column_or_expressions.append(pl.col(col_name) == float(or_val_cleaned))
            # Add more type handling here if necessary

        if per_column_or_expressions:
            if len(per_column_or_expressions) > 1:
                filter_expressions_list.append(pl.any_horizontal(per_column_or_expressions))
            else:
                filter_expressions_list.append(per_column_or_expressions[0])

    # Process extra_constraints (typically direct equality)
    if extra_constraints:
        for col_name, constraint_val_str in extra_constraints.items():
            if col_name not in df_schema_provider.columns:
                logger.debug(f"Constraint column '{col_name}' not in DataFrame schema. Skipping.")
                continue
            if constraint_val_str: # Ensure constraint value is not empty
                 # Assuming simple equality for constraints (e.g., faculty='BMS')
                filter_expressions_list.append(pl.col(col_name).str.to_lowercase() == str(constraint_val_str).lower())

    if not filter_expressions_list:
        return pl.lit(True) # No filters, return expression that evaluates to True for all rows
    return pl.all_horizontal(filter_expressions_list) # Combine all filter parts with AND


def get_filtered_sorted_df(
    app_state: AppState, extra_constraints: Optional[Dict[str, str]] = None
) -> pl.DataFrame:
    """
    Applies filtering and sorting to the main items DataFrame based on AppState
    and any additional constraints.

    Args:
        app_state (AppState): The current application state (filters, sort settings).
        extra_constraints (Optional[Dict[str, str]]): Additional fixed constraints to apply (e.g., faculty).

    Returns:
        pl.DataFrame: The filtered and sorted DataFrame.
    """
    base_df = get_item_df() # Get cached or fresh main DataFrame
    if base_df.is_empty():
        return base_df # Return empty if base is empty

    user_active_filters = app_state.get_active_filters()

    final_filter_expression = _create_filter_expression(base_df, user_active_filters, extra_constraints)

    filtered_df: pl.DataFrame
    try:
        filtered_df = base_df.filter(final_filter_expression)
    except Exception as e_filter:
        logger.error(f"Error applying filters: {e_filter}. Returning empty DataFrame for this query.", exc_info=True)
        return base_df.clear() # Return empty DataFrame of same schema

    # Apply sorting
    sort_by_col = app_state.sort_by
    is_sort_desc = app_state.sort_desc
    if sort_by_col and sort_by_col in filtered_df.columns:
        try:
            filtered_df = filtered_df.sort(by=sort_by_col, descending=is_sort_desc, nulls_last=True)
        except Exception as e_sort: # Catch potential Polars sort errors
            logger.warning(f"Sort warning on column '{sort_by_col}': {e_sort}. Proceeding without this sort.")
    elif sort_by_col:
        logger.warning(f"Sort column '{sort_by_col}' not found in filtered DataFrame. Skipping sort.")

    return filtered_df


def _apply_filters_for_count(
    filters_to_apply: Dict[str, str],
    faculty_constraint: Optional[Dict[str, str]] = None
) -> int:
    """
    Helper to apply a given set of filters (and optional faculty constraint)
    to the global DataFrame and return the count of matching rows.
    Optimized for counting using lazy evaluation.

    Args:
        filters_to_apply (Dict[str, str]): Dictionary of filters to apply.
        faculty_constraint (Optional[Dict[str, str]]): Faculty constraint.

    Returns:
        int: The count of rows matching the filters. Returns 0 on error.
    """
    base_df_for_count = get_item_df()
    if base_df_for_count.is_empty(): return 0

    if not filters_to_apply and not faculty_constraint:
        return base_df_for_count.height

    # Combine user filters and faculty constraint for expression generation
    combined_filters = filters_to_apply.copy()
    if faculty_constraint:
        combined_filters.update(faculty_constraint) # faculty_constraint will overwrite if key exists

    filter_expr = _create_filter_expression(base_df_for_count, combined_filters) # Pass combined
    try:
        count = base_df_for_count.lazy().filter(filter_expr).select(pl.len()).collect().item(0,0)
        return count
    except Exception as e:
        logger.error(f"Error applying filters for count calculation ({filters_to_apply}, {faculty_constraint}): {e}")
        return 0


async def store_item_changes(
    items_to_update_data: List[Dict[str, Any]], # Changed from single dict or list to always list
    auth_details: Dict[str, Any]
) -> None:
    """
    Updates copyright items in the database based on changes from the dashboard.

    For each item in `items_to_update_data` (which contains partial updates):
    1. Fetches the full existing item data from the database.
    2. Overwrites fields in the fetched data with values from `items_to_update_data`.
    3. Calls `update_copyright_items` to save these changes to the database,
       attributing changes to the user from `auth_details`.

    Args:
        items_to_update_data (List[Dict[str, Any]]): A list of dictionaries, where each dict
            represents an item to update. It must contain 'material_id' and other fields
            that have been changed by the user.
        auth_details (Dict[str, Any]): User authentication details (e.g., email, name)
                                     for audit logging.
    """
    user_info_for_log: Dict[str, Optional[str]] = {
        "email": auth_details.get("email"),
        "name": auth_details.get("name"), # For potential future use in log
    }

    material_ids_to_update: List[str] = [
        str(item_data.get("material_id")) for item_data in items_to_update_data if item_data.get("material_id")
    ]
    if not material_ids_to_update:
        logger.info("No material_ids provided in input data for store_item_changes.")
        return

    # Fetch current state of these items from DB (or cache)
    # Ensure material_id is string for filtering if it's string in DataFrame
    current_items_df = get_item_df().filter(pl.col("material_id").cast(pl.Utf8).is_in(material_ids_to_update))

    updated_item_data_list: List[Dict[str, Any]] = []

    for item_changes_from_ui in items_to_update_data:
        mat_id_ui = str(item_changes_from_ui.get("material_id"))
        if not mat_id_ui: continue

        current_item_rows = current_items_df.filter(pl.col("material_id").cast(pl.Utf8) == mat_id_ui)
        if current_item_rows.is_empty():
            logger.warning(f"Material ID {mat_id_ui} not found in current data for update. Skipping.")
            continue

        full_item_data_dict: Dict[str, Any] = current_item_rows.to_dicts()[0] # Get first (and only) row as dict

        # Apply changes from UI to this full data dict
        for key, new_value in item_changes_from_ui.items():
            if key == "material_id": continue # Don't change ID

            # Special handling for faculty if it comes as abbreviation but DB expects faculty_id (or ORM handles it)
            # Assuming update_copyright_items and ORM can handle 'faculty' field with abbreviation.
            # If 'faculty_id' is directly manipulated, that needs care.
            # The current 'update_copyright_items' takes a DataFrame and its logic for 'faculty' vs 'faculty_id'
            # would apply. If item_changes_from_ui sends 'faculty' as abbreviation, it should be fine.

            if key in full_item_data_dict:
                if str(new_value) != str(full_item_data_dict[key]): # Compare as strings for simplicity
                    logger.debug(f"Updating {key} for {mat_id_ui}: '{full_item_data_dict[key]}' -> '{new_value}'")
                full_item_data_dict[key] = new_value
            else: # Key from UI not in current DB data dict (new field or schema mismatch?)
                logger.debug(f"Adding new key '{key}' with value '{new_value}' for {mat_id_ui} during update.")
                full_item_data_dict[key] = new_value

        updated_item_data_list.append(full_item_data_dict)

    if updated_item_data_list:
        logger.info(f"Storing changes for {len(updated_item_data_list)} items.")
        # update_copyright_items expects DataFrame or list of dicts.
        # overwrite=True ensures that the values from dashboard (which are user's explicit changes)
        # take precedence for the fields being updated.
        await update_copyright_items(
            updated_item_data_list,
            overwrite=True, # To ensure dashboard changes are applied directly for modified fields
            user_info=user_info_for_log,
            update_relations=False # Relations typically not modified by simple field updates from dashboard
        )
        refresh_copyright_df() # Refresh global cache after updates
        logger.info("Item changes stored and global DataFrame refreshed.")
    else:
        logger.info("No valid items with changes to store.")


async def get_item_history_for_dashboard(material_id: int) -> List[Dict[str, Any]]:
    """
    Retrieves and formats the edit history for a given material_id for dashboard display.

    Args:
        material_id (int): The material ID to fetch history for.

    Returns:
        List[Dict[str, Any]]: A list of dictionaries, where each dictionary represents
                              a formatted history entry.
    """
    logger.debug(f"Fetching item history for material_id: {material_id}")
    history_records: List[ItemUpdate] = await retrieve_item_history([material_id]) # retrieve_item_history is async

    formatted_history: List[Dict[str, Any]] = []
    for record in history_records:
        changes_dict = record.change_details # This is already a dict due to JSONField

        # Format for display (example: list of changed fields)
        changed_fields_summary_list: List[str] = []
        if isinstance(changes_dict, dict):
            for field, change_detail in changes_dict.items():
                if field not in ["material_id", "update_time", "modified_by"]:
                    if isinstance(change_detail, dict):
                        changed_fields_summary_list.append(
                            f"{field}: '{change_detail.get('old')}' -> '{change_detail.get('new')}'"
                        )
                    else: # Should not happen if format is consistent
                         changed_fields_summary_list.append(f"{field}: {change_detail}")

        formatted_history.append({
            "timestamp": record.created_at.strftime("%Y-%m-%d %H:%M:%S") if record.created_at else "N/A",
            "modified_by": changes_dict.get("modified_by", "Unknown"),
            "material_id_log": changes_dict.get("material_id", record.material_id), # From log or record
            "changes_summary": "; ".join(changed_fields_summary_list) if changed_fields_summary_list else "No specific field changes logged.",
            "full_details_json": json.dumps(changes_dict, indent=2) # For raw view or debugging
        })
    logger.info(f"Retrieved {len(formatted_history)} history entries for material_id {material_id}.")
    return formatted_history
```
