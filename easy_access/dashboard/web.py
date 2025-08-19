"""
This module provides core web functionalities for the dashboard, including:
- Application state management (`AppState`) for user interface state (filters, pagination, sorting).
- Basic user authentication and session management.
- Middleware for protecting routes.
- Static file parameter registration for `fasthtml`.
"""

import json
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any  # Added Set, Union, Tuple

from fasthtml.common import Beforeware, RedirectResponse, Request  # For type hints

# from rich import print # Removed, using logging
from easy_access.dashboard.constants import (  # db from constants
    DEFAULT_PER_PAGE,
    db,
)
from easy_access.settings import SETTINGS, DirSetting  # For SCRIPT_DATA dir

logger = logging.getLogger(__name__)

# --- Configuration Candidates ---
USERS_FILENAME: str = "users.json.secret"  # Candidate for settings.yaml
# --- End Configuration Candidates ---

# --------------------
# Application State Management
# --------------------


@dataclass
class AppState:
    """
    Manages the UI state of the application, including pagination, sorting, and filters.
    Instances are typically stored in the user's session.
    """

    page: int = 1
    per_page: int = DEFAULT_PER_PAGE
    sort_by: str | None = None
    sort_desc: bool = False
    filters: dict[str, str] = field(
        default_factory=dict
    )  # Stores active filters as {col_name: "val1|val2"}

    def set_page(self, page_num: int) -> None:
        """Sets the current page number, ensuring it's at least 1."""
        self.page = max(1, page_num)

    def set_per_page(self, num_per_page: int) -> None:
        """Sets the number of items to display per page, ensuring it's at least 1."""
        self.per_page = max(1, num_per_page)

    def set_sort(self, sort_by_col: str | None, is_descending: bool) -> None:
        """Sets the sorting column and direction."""
        # TODO: Optionally validate sort_by_col against a list of sortable columns
        self.sort_by = sort_by_col
        self.sort_desc = is_descending

    def apply_filter_change(
        self, filter_key: str, toggled_value: str, is_checked: bool
    ) -> bool:
        """
        Updates filters based on a checkbox change (multi-select OR logic for a single key).

        Args:
            filter_key: The filter category (e.g., 'workflow_status').
            toggled_value: The specific value within the category that was toggled.
            is_checked: The new state of the checkbox (True if checked, False if unchecked).

        Returns:
            bool: True if the filter state for this key was effectively changed, False otherwise.
        """
        current_filter_values_str = self.filters.get(filter_key, "")
        current_values_set: set[str] = (
            set(current_filter_values_str.split("|"))
            if current_filter_values_str
            else set()
        )

        original_set_size = len(current_values_set)
        # Remove empty strings that might arise from split('') or initial state
        current_values_set.discard("")

        if is_checked:
            current_values_set.add(toggled_value)
        else:
            current_values_set.discard(toggled_value)

        if not current_values_set:  # No values selected for this filter key
            if filter_key in self.filters:
                del self.filters[filter_key]
                return True  # Changed from something to nothing
            return False  # Was already nothing
        else:
            new_filter_str = "|".join(sorted(list(current_values_set)))
            if new_filter_str != current_filter_values_str:
                self.filters[filter_key] = new_filter_str
                return True  # Value changed
            return False  # No effective change

    def apply_text_filter(self, filter_key: str, new_text_value: str) -> bool:
        """
        Sets or removes a text-based filter value for a given filter key.

        Args:
            filter_key: The filter category (e.g., 'department').
            new_text_value: The new text value from the input. An empty string removes the filter.

        Returns:
            bool: True if the filter value for this key changed, False otherwise.
        """
        original_value = self.filters.get(filter_key)
        cleaned_new_text_value = new_text_value.strip()

        if cleaned_new_text_value:  # If new value is not empty
            if cleaned_new_text_value != original_value:
                self.filters[filter_key] = cleaned_new_text_value
                return True
            return False  # No change
        else:  # Empty new value means remove the filter for this key
            if filter_key in self.filters:
                del self.filters[filter_key]
                return True  # Filter was removed
            return False  # Filter was not present anyway

    def clear_filters(self) -> None:
        """Clears all active user-defined filters and resets page to 1."""
        self.filters = {}
        self.set_page(1)  # Resetting page on filter clear is common UX
        logger.info("All filters cleared and page reset.")

    def nav_params_dict(self) -> dict[str, Any]:
        """Returns a dictionary of navigation parameters (page, sort, per_page) for URL/form generation."""
        return {
            "page": self.page,
            "per_page": self.per_page,
            "sort_by": self.sort_by or "",  # Ensure string for URL generation
            "sort_desc": str(self.sort_desc).lower(),  # Consistent 'true'/'false'
        }

    def filter_params_dict(self) -> dict[str, str]:
        """Returns a dictionary of currently active filters, prefixed with 'filter_' for forms/URLs."""
        if not self.filters:
            return {}
        return {
            f"filter_{k}": v for k, v in self.filters.items() if v
        }  # Ensure only non-empty filter values

    def all_params_dict(self) -> dict[str, Any]:
        """Returns a combined dictionary of all navigation and filter parameters."""
        return {**self.nav_params_dict(), **self.filter_params_dict()}

    def get_active_filters(self) -> dict[str, str]:
        """Returns a copy of the currently active user-defined filters."""
        return self.filters.copy()


def load_app_state(session: dict[str, Any]) -> AppState:
    """
    Loads the AppState object from the session dictionary.
    If 'app_state' is not found in the session or is invalid (e.g., due to stale data
    or structure mismatch), it returns a default AppState instance.

    Args:
        session (Dict[str, Any]): The current user session dictionary.

    Returns:
        AppState: An AppState instance, either loaded from session or a new default.
    """
    app_state_from_session: dict[str, Any] | None = session.get("app_state")
    if isinstance(app_state_from_session, dict):
        try:
            # Filter out keys not in AppState.__annotations__ before **
            valid_keys = AppState.__annotations__.keys()
            filtered_state_dict = {
                k: v for k, v in app_state_from_session.items() if k in valid_keys
            }
            state = AppState(**filtered_state_dict)
            logger.debug("AppState loaded from session.")
            return state
        except (
            TypeError,
            ValueError,
        ) as e:  # Catch errors if dict structure doesn't match AppState
            logger.warning(
                f"Invalid 'app_state' structure in session: {e}. Using default AppState."
            )
    else:
        if app_state_from_session is not None:  # Log if it exists but isn't a dict
            logger.warning(
                f"Found 'app_state' in session but it's not a dictionary (type: {type(app_state_from_session)}). Using default AppState."
            )
        else:  # Not in session
            logger.debug("'app_state' not found in session. Using default AppState.")

    return AppState()  # Return default state


# --------------------
# Authentication
# --------------------


@dataclass
class Login:
    """Represents login credentials."""

    email: str
    pwd: str  # Password (plain text from form, to be hashed for comparison)


def create_users_from_secrets() -> None:
    """
    Loads user definitions from a JSON secret file (defined by USERS_FILENAME)
    and ensures they exist in the database. If a user from the file is not in
    the database, they are inserted.

    This function is typically called at application startup to synchronize
    the users table with the contents of the secret file.

    Note:
        This implies user management is primarily done via the JSON file.
        Password hashing should be implemented for security if not already.
        The current implementation stores/compares plain text passwords from the file.
    """
    # TODO: Implement password hashing for security.
    users_secret_file_path: Path = (
        SETTINGS.dirs[DirSetting.SCRIPT_DATA].full / USERS_FILENAME
    )

    if not users_secret_file_path.exists():
        logger.error(
            f"Users secret file '{USERS_FILENAME}' not found at {users_secret_file_path}. Cannot create/update users."
        )
        return

    try:
        with open(users_secret_file_path, encoding="utf-8") as f:
            users_data_from_file: dict[str, dict[str, str]] = json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        logger.error(
            f"Error reading or parsing users secret file {users_secret_file_path}: {e}"
        )
        return

    try:
        existing_users_emails_q = db.q(
            f"SELECT email FROM {users.tname}"
        )  # Use users.tname for table name
        existing_users_emails: list[str] = [
            user_row.get("email") for user_row in existing_users_emails_q if user_row
        ]  # type: ignore
    except Exception as e_db_query:  # Catch potential DB query errors
        logger.error(f"Error querying existing users from database: {e_db_query}")
        return

    users_added_count = 0
    for email, data in users_data_from_file.items():
        if email not in existing_users_emails:
            try:
                # TODO: Hash data["password"] before inserting
                users.insert(
                    name=data.get("name"),
                    email=email,
                    pwd=data.get("password"),  # PWD should be hashed
                    faculty=data.get("faculty"),
                    role=data.get("role"),
                )
                users_added_count += 1
                logger.info(f"Added new user to database: {email}")
            except Exception as e_insert:  # Catch DB insert errors
                logger.error(f"Error inserting user {email} into database: {e_insert}")

    if users_added_count > 0:
        logger.info(
            f"Synchronized users from secrets file. Added {users_added_count} new users."
        )
    else:
        logger.debug("User database is already synchronized with the secrets file.")


# `users` table definition using fastlite
# This should ideally be defined once. The `if users not in db.t:` check handles recreation.
users: Any = db.t.users  # Get table object from fastlite db instance
if (
    users.tname not in db.tnames
):  # Check if table name exists in db's list of table names
    logger.info(f"Table '{users.tname}' not found, creating it.")
    users.create(
        email=str,
        name=str,
        pwd=str,  # PWD should store hash
        faculty=str,
        role=str,
        pk="email",
    )
else:
    logger.debug(f"Table '{users.tname}' already exists.")

# Populate users from secrets file at startup
create_users_from_secrets()

# Redirect response for login failures
LOGIN_REDIRECT_RESPONSE = RedirectResponse("/login", status_code=303)


def auth_middleware_before_request(
    request: Request, session: dict[str, Any]
) -> RedirectResponse | None:
    """
    FastHTML Beforeware function to check user authentication for protected routes.
    If the user is not authenticated (no 'auth' key in session), they are redirected
    to the login page.

    Args:
        request (Request): The incoming HTTP request.
        session (Dict[str, Any]): The user's session data.

    Returns:
        Optional[RedirectResponse]: A RedirectResponse to the login page if not authenticated,
                                    otherwise None (allowing request to proceed).
    """
    # Store auth details from session into request scope for easier access in route handlers
    request.scope["auth"] = session.get("auth")

    if not request.scope["auth"]:  # If 'auth' details are not in session
        logger.debug(
            f"User not authenticated for path: {request.url.path}. Redirecting to login."
        )
        return LOGIN_REDIRECT_RESPONSE
    logger.debug(
        f"User authenticated: {request.scope['auth'].get('email')}. Allowing access to {request.url.path}"
    )
    return None  # User is authenticated, proceed with the request


# Middleware instance, skipping static files, favicon, and login page itself
# Ensure regex patterns for skip are correct
auth_bware = Beforeware(
    auth_middleware_before_request,
    skip=[
        r"/favicon\.ico$",
        r"/static/.*",  # Skips anything under /static/
        r"/login$",  # Skips the /login path itself
        r".*\.css$",  # Skip CSS files (if served dynamically and not under /static/)
        # Add other specific public paths if needed
    ],
)


# --------------------
# Static File Parameter Registration (fasthtml specific)
# --------------------
# This is a fasthtml utility to help its router recognize common static file extensions.
# It does not serve files but helps in pattern matching for routes if used.
def _register_static_file_extensions() -> None:
    """Registers common static file extensions with fasthtml's parameter regex system."""
    from fasthtml.xtend import (
        reg_re_param,  # Local import as it's a specific setup call
    )

    common_static_exts = "ico|gif|jpg|jpeg|webm|css|js|woff|png|svg|mp4|webp|ttf|otf|eot|woff2|txt|xml|html|pdf|md"
    try:
        reg_re_param("static_common_exts", common_static_exts)
        logger.debug(
            f"Registered common static file extensions for fasthtml: {common_static_exts}"
        )
    except Exception as e_reg:  # pragma: no cover
        logger.warning(
            f"Could not register static file extensions with fasthtml: {e_reg}"
        )


_register_static_file_extensions()
