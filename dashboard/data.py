# data related classes / functions for the dashboard
# e.g. filtering, sorting, retrieval, storage, etc
import contextlib
import math
import re
from copy import copy, deepcopy
from dataclasses import asdict, dataclass
from functools import lru_cache
from typing import Any

import polars as pl
from loguru import logger

from easy_access.dashboard.constants import (
    BADGE_STYLES,
    FILTERABLE_COLUMNS,
)
from dashboard.web import AppState, load_app_state
from easy_access.db.retrieve import retrieve_copyright_items, retrieve_item_history
from easy_access.db.update import update_copyright_items
from easy_access.settings import SETTINGS

EMPTY_FILTERS: dict[str, list[str]] = {}


# -----------------------------
#       data state management
# -----------------------------
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


def process_state(
    session: dict[str, Any], request_params: dict[str, Any]
) -> tuple[dict, AppState]:
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
    logger.debug(
        f"Processing state with params: {request_params}"
    )  # Debug incoming params

    # 1. Load initial state
    app_state = load_app_state(session)
    initial_state_dict = asdict(app_state)  # For change detection

    filter_change_occurred = False  # Specific flag for filter changes to reset page

    if request_params.get("action") == "reset":
        # Only apply reset if there are filters or page is not 1
        if app_state.filters or app_state.page != 1:
            app_state.clear_filters()
            app_state.set_page(1)
            filter_change_occurred = True  # Resetting filters counts
        else:
            logger.debug("No change needed for reset.")

    elif "page" in request_params:
        try:
            page = int(request_params["page"])
            if page != app_state.page:
                app_state.set_page(page)
        except (ValueError, TypeError):
            logger.warning(
                f"Invalid page parameter '{request_params['page']}'. Ignoring."
            )
    elif "sort_by" in request_params:
        sort_by = request_params.get("sort_by")
        sort_desc = request_params.get("sort_desc", "False").lower() == "true"
        if sort_by != app_state.sort_by or sort_desc != app_state.sort_desc:
            app_state.set_sort(sort_by if sort_by else None, sort_desc)
            app_state.set_page(1)
    elif "per_page" in request_params:
        try:
            per_page = int(request_params["per_page"])
            if per_page != app_state.per_page:
                app_state.set_per_page(per_page)
                app_state.set_page(1)  # Reset page when per_page changes
        except (ValueError, TypeError):
            logger.warning(
                f"Invalid per_page parameter '{request_params['per_page']}'. Ignoring."
            )

    else:  # Check for filter changes by comparing with current state
        is_checkbox_change = "checked" in request_params
        current_filters = app_state.get_active_filters()  # Get a copy

        # Iterate through request params to find the one representing the change
        for param_key, param_value in request_params.items():
            if param_key.startswith("filter_"):
                filter_key = param_key[len("filter_") :]

                if filter_key not in FILTERABLE_COLUMNS:
                    # skipping unknown filter key
                    continue  # Silently ignore unknown keys

                current_value_str = current_filters.get(filter_key, "")

                # --- Logic for Checkbox Filters ---
                if is_checkbox_change:
                    # The relevant param_value is the value of the *specific checkbox* clicked.
                    # We need this value to simulate the change.
                    is_checked = request_params["checked"].lower() == "true"
                    clicked_value = (
                        param_value  # The value from the clicked checkbox param
                    )

                    # Simulate the change based on the clicked value and its intended state
                    temp_values = (
                        set(current_value_str.split("|"))
                        if current_value_str
                        else set()
                    )
                    if is_checked:
                        temp_values.add(clicked_value)
                    else:
                        temp_values.discard(clicked_value)
                    temp_values.discard("")  # Clean up potential empty strings
                    new_simulated_value_str = (
                        "|".join(sorted(list(temp_values))) if temp_values else ""
                    )

                    # If the simulated new string is different from the current one, this is our change.
                    if new_simulated_value_str != current_value_str:
                        # Apply the change using the specific clicked value and intended state
                        if app_state.apply_filter_change(
                            filter_key, clicked_value, is_checked
                        ):
                            filter_change_occurred = True
                        break  # Process only the first detected difference for checkboxes

                # --- Logic for Text Filters ---
                elif not is_checkbox_change:
                    # For text filters, the param_value is the new text value (can be empty).
                    # Compare directly with the current value string for this key.
                    if param_value != current_value_str:
                        if app_state.apply_text_filter(filter_key, param_value):
                            filter_change_occurred = True
                        # Important: Break here assumes only one text filter changes per request.
                        # This holds true if text inputs use hx-vals with getElementById.
                        break

        # Reset page if a filter changed (and it wasn't handled by per_page already)
        if filter_change_occurred and "per_page" not in request_params:
            app_state.set_page(1)

    # 3. Save potentially modified state back to session *if* a change occurred
    final_state_dict = asdict(app_state)
    if final_state_dict != initial_state_dict:
        logger.debug("Saving updated state to session.")
        session["app_state"] = final_state_dict
    else:
        logger.debug("No state change detected, session not updated.")

    # 4. Return the session and the AppState reflecting the request intent
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
    # 1. Get Auth Details from Session
    auth_details = session.get("auth", {})

    # 2. Determine Auth Constraint
    faculty_constraint = None
    user_faculty = auth_details.get("faculty")
    user_role = auth_details.get("role")
    if user_role != "admin" and user_faculty and user_faculty != "all":
        faculty_constraint = {"faculty": user_faculty}

    # 3. Filter Data (using adapted get_filtered_sorted_df)
    filtered_df = get_filtered_sorted_df(
        app_state, extra_constraints=faculty_constraint
    )
    total_filtered_rows = filtered_df.height

    # 4. Calculate Total Pages & *Validate* Page Number in a *copy* of the state
    validated_app_state = deepcopy(app_state)
    total_pages = (
        math.ceil(total_filtered_rows / validated_app_state.per_page)
        if validated_app_state.per_page > 0
        else 1
    )
    validated_app_state.page = max(
        1, min(validated_app_state.page, total_pages if total_pages > 0 else 1)
    )

    # 5. Calculate Filter Counts (using adapted calculate_counts_for_ui)
    # Pass the *requested* state (app_state) because counts show potential changes
    # Pass global_df (copyright_df_global assumed available) and constraint
    filter_counts = calculate_counts_for_ui(app_state, faculty_constraint)

    # 6. Prepare Slice using *Validated* State
    offset = (validated_app_state.page - 1) * validated_app_state.per_page
    df_slice = filtered_df.slice(offset, validated_app_state.per_page)

    # 7. Return results with the *Validated* State
    return ProcessedDataResult(
        app_state=validated_app_state,  # Return the state with corrected page number
        df_slice=df_slice,
        total_filtered_rows=total_filtered_rows,
        total_pages=total_pages,
        filter_counts=filter_counts,
    )


def calculate_counts_for_ui(
    app_state: AppState, faculty_constraint: dict | None
) -> dict[str, dict[str, int]]:
    """
    Calculates counts for UI filter display based on current filters.

    Args:
        app_state: The current AppState reflecting user's choices.
        global_df: The complete, unfiltered DataFrame.
        faculty_constraint: An optional dictionary specifying the faculty constraint (e.g., {"faculty": "EEMCS"}).

    Returns:
        A dictionary where keys are filter columns and values are dicts
        mapping filter options to their calculated current counts.
    """
    full_df = get_filtered_sorted_df(
        app_state=app_state, extra_constraints=faculty_constraint
    )
    if not EMPTY_FILTERS:
        tmp_app_state = copy(app_state)
        tmp_app_state.clear_filters()
        get_filter_counts(tmp_app_state)  # Initialize EMPTY_FILTERS if not set

    applied_filters = app_state.get_active_filters()

    df = full_df.lazy()
    # for each col in FILTERABLE_COLUMNS, calculate the counts for each option defined
    # in the keys of dict BADGE_STYLES[col]
    # skipping col values that are in EMPTY_FILTERS[col]

    # then process the dicts to only include valid options
    results = {}
    for col in FILTERABLE_COLUMNS:
        if col not in df.columns:
            continue

        if col == "manual_classification":
            options = list(BADGE_STYLES["classification"].keys())
        else:
            if col not in BADGE_STYLES:
                continue
            options = list(BADGE_STYLES[col].keys())
        if col in applied_filters:
            # if the col is currently being filtered on, we'll do the following:
            # - for each option in options, check if the option is in the applied filter
            # if it is, set the count to the current total
            # else, temporarily add the option to the applied filter and get the total row count
            # then remove the filter option again
            # repeat for all options
            # directly store each found count in the result dict
            # then add it to results and continue
            result = {}
            for option in options:
                if option in EMPTY_FILTERS.get(col, []):
                    # if the option is in EMPTY_FILTERS, set count to '-'
                    # this means the value is not in the df at all
                    result[option] = "-"
                    continue
                if option in applied_filters:
                    # if the option is already in the applied filter, set count to '-'
                    # this means the value is in the filtered df
                    result[option] = df.collect().height

                else:
                    temp_filters = copy(applied_filters)
                    temp_filters[col] += f"|{option}"
                    count = _apply_filters_for_count(temp_filters)
                    result[option] = count
            results[col] = result
            continue
        result = df.group_by(col).len().collect().to_dicts()
        result = [x for x in result if x.get(col) in options]
        for option in options:
            if option not in [x[col] for x in result]:
                if option in EMPTY_FILTERS.get(col, []):
                    # if the option is in EMPTY_FILTERS, set count to '-'
                    # this means the value is not in the df at all
                    result.append({col: option, "len": "-"})
                else:
                    # else set count to 0
                    # this means the value is in the df, but not in the filtered df
                    result.append({col: option, "len": 0})
        results[col] = {x[col]: x["len"] for x in result}

    # logger.debug(f"Final Counts: {all_counts}")
    return results


@lru_cache
def get_item_df() -> pl.DataFrame:
    """
    Returns the global copyright DataFrame defined in this module, for use in other modules.
    """
    return retrieve_copyright_items(SETTINGS)


def refresh_copyright_df() -> pl.DataFrame:
    """
    Refreshes the global copyright DataFrame by re-fetching the data from the database.

    Returns:
        A Polars DataFrame containing the latest copyright data.
    """
    get_item_df.cache_clear()
    return get_item_df()


# -----------------------
#       filtering
# -----------------------


def _create_filter_expression(
    df: pl.DataFrame,
    user_filters: dict[str, str],
    extra_constraints: dict[str, str] | None = None,
) -> pl.Expr:
    filter_expressions = []
    for col, value in user_filters.items():
        actual_col = col  # Assume col name matches DataFrame column directly
        if actual_col not in df.columns:
            logger.warning(f"Filter column '{actual_col}' not in DataFrame. Skipping.")
            continue
        if value is None or value == "":  # Skip empty filter values explicitly
            continue

        try:
            or_values = (
                value.split("|") if isinstance(value, str) and "|" in value else [value]
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
                        | (pl.col(actual_col) == "-")  # Include hyphen check if needed
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
                            pl.col(actual_col).str.to_lowercase() == or_value.lower()
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
            logger.warning(f"Filter warning processing '{actual_col}'='{value}': {e}")

    if extra_constraints:
        for col, value in extra_constraints.items():
            if col not in df.columns:
                logger.warning(f"Constraint column '{col}' not in DataFrame. Skipping.")
                continue
            if value:  # Only apply if constraint value is non-empty
                # Assuming simple equality constraint for now (e.g., faculty)
                # Adapt if more complex constraints are needed
                filter_expressions.append(
                    pl.col(col).str.to_lowercase() == str(value).lower()
                )  # Ensure comparison is case-insensitive
    if not filter_expressions:
        return pl.lit(True)
    return pl.all_horizontal(filter_expressions)


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
    df = get_item_df()  # Start with the global DataFrame
    user_filters = app_state.get_active_filters()
    sort_by = app_state.sort_by
    sort_desc = app_state.sort_desc

    expression = _create_filter_expression(df, user_filters, extra_constraints)

    # --- Apply combined filters ---

    try:
        df = df.filter(expression)
    except Exception as e:
        logger.error(f"Error applying filters: {e}")
        # Fallback: return empty DataFrame on error? Or original?
        return df.clear()

    # --- Apply sorting ---
    if sort_by and sort_by in df.columns:
        try:
            df = df.sort(by=sort_by, descending=sort_desc, nulls_last=True)
        except Exception as e:
            logger.warning(f"Sort warning on '{sort_by}': {e}")
    elif sort_by:
        logger.warning(
            f"Sort column '{sort_by}' not found in DataFrame. Skipping sort."
        )
    return df


def _apply_filters_for_count(filters: dict[str, str]) -> int:
    """
    Helper to apply filters to the provided global DataFrame and return the count.
    Optimized for counting.

    Args:
        filters: Dictionary of filters to apply.
        global_df: The complete, unfiltered DataFrame to filter on.

    Returns:
        The count of rows matching the filters.
    """
    df = get_item_df()  # Use the passed global DataFrame

    if not filters:
        return df.height  # Return total count if no filters

    filter_expr = _create_filter_expression(df, filters)
    try:
        # Apply filters and get count directly using lazy evaluation
        count = df.lazy().filter(filter_expr).select(pl.len()).collect().item()
        return count
    except Exception as e:
        logger.error(f"Count Filter error applying filters {filters}: {e}")
        return 0  # Return 0 on error


def get_filter_counts(app_state: AppState) -> dict[str, dict[str, int]]:
    """Calculates the counts for each filter option based on current state."""
    all_counts = {}
    current_filters = app_state.get_active_filters()
    # Define the options we need counts for (reuse from page_header_component)
    # Combine all options into one structure for iteration

    for filter_key, options in BADGE_STYLES.items():
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
                all_counts[filter_key][option_value] = "-"
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

    if not EMPTY_FILTERS:
        # Only update EMPTY_FILTERS if it was not previously set
        _update_empty_filters(all_counts)
    return all_counts


def _update_empty_filters(filter_counts: dict[str, dict[str, int]]):
    """
    input filter_counts is the result of get_filter_counts() run on a non-filtered global df -- so on init and on refresh.
    Stores the values that are missing from the global df for each filterable column in EMPTY_FILTERS.
    """
    global EMPTY_FILTERS

    # Update EMPTY_FILTERS based on filter_counts
    for col, counts in filter_counts.items():
        if col in FILTERABLE_COLUMNS:
            EMPTY_FILTERS[col] = [key for key, count in counts.items() if count == 0]


# -----------------------
#       storage
# -----------------------


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
            full_item_data = (
                get_item_df().filter(pl.col("material_id") == material_id).to_dicts()
            )
            if full_item_data:
                logger.info(f"Updating material_id {material_id}")
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
                            logger.info(
                                f"Updating {key}: {full_item_data[key]} --> {value}"
                            )
                        full_item_data[key] = value
                full_data_list.append(full_item_data)

    await update_copyright_items(
        SETTINGS,
        full_data_list,
        update_relations=False,
        overwrite=True,
        user_info=user_info,
    )

    # the global_df is now stale, reload it


# ----------------------
#     retrieval
# ----------------------


async def get_item_history(material_id: int):
    """
    retrieves the items edit history and parses it
    """
    history = await retrieve_item_history([material_id], SETTINGS)

    return history
