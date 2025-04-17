# This file contains old / temp / misc code that is not used (directly) in the main codebase.
# this will eventually be removed, for now it's something like a recycling bin


# --- monsterui fixes ---
# NOTE: no longer required, merged into main library


class LabelT(VEnum):
    """FIX: replaces danger with destructive"""

    def _generate_next_value_(name, start, count, last_values):
        return str2ukcls("label", name)

    primary = auto()
    secondary = auto()
    destructive = auto()


# A fix for the monsterui library to enable theme switching in the webdashboard
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


# unused filter-related functions


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


@rt("/update_filter_counts", methods=["POST"])
async def update_filter_counts_endpoint(session: dict, request: Request):
    """
    Updates session state based on checkbox change, calculates all filter counts,
    and returns OOB swaps for all checkbox filter groups.
    """
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
