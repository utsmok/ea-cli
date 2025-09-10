# web-related classes / functions for the dashboard
# e.g. state handling, auth, etc
import json
from dataclasses import dataclass, field
from typing import Any

from fasthtml.common import *
from loguru import logger

from dashboard.constants import DEFAULT_PER_PAGE, db
from easy_access.settings import SETTINGS, DirSetting

# --------------------
#       state
# --------------------


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
        logger.warning("Invalid or missing app_state in session. Using default.")
        state = AppState()  # Return default state
    return state


# --------------------
#       auth
# --------------------


@dataclass
class Login:
    email: str
    pwd: str


def create_users_from_secrets():
    # read in users.json.secret from root
    with open(
        SETTINGS.dirs[DirSetting.SCRIPT_DATA].full / "users.json.secret",
        encoding="utf-8",
    ) as f:
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


def before(req: Request, sess: dict):
    auth = req.scope["auth"] = sess.get("auth")
    if not auth:
        return login_redir


bware = Beforeware(
    before, skip=[r"/favicon\.ico", r"/static/.*", r"/static/.*", r".*\.css", "/login"]
)

# --------------------
#       static files
# --------------------
reg_re_param(
    "static",
    "ico|gif|jpg|jpeg|webm|css|js|woff|png|svg|mp4|webp|ttf|otf|eot|woff2|txt|xml|html|pdf|md",
)
