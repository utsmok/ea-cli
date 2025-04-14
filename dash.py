# dash.py

import contextlib
import json
import math
import traceback
from collections import defaultdict
from dataclasses import dataclass, field
from enum import auto
from pathlib import Path
from typing import Any

import fasthtml.common as fh
import polars as pl
from fastcore.utils import *
from fastcore.xml import FT
from fasthtml.common import *
from fasthtml.components import Button
from fastlite import database
from monsterui.all import *
from monsterui.foundations import VEnum, str2ukcls
from rich import print
from starlette.staticfiles import StaticFiles

from easy_access.db.retrieve import retrieve_copyright_items, retrieve_osiris_data
from easy_access.db.update import update_copyright_items
from easy_access.settings import SETTINGS, DirSetting

# --- monsterui fixes ---

"""
UT Logo Variants
<!-- Logo Variant 1: Black / Green -->
<div class="ut-logo-base logo-variant-1">
  <p>UNIVERSITY</p>
  <p>OF TWENTE</p>
  <p>CDD//UT.</p>
</div>

<hr> <!-- Separator for clarity -->

<!-- Logo Variant 2: Black / Red -->
<div class="ut-logo-base logo-variant-2">
  <p>UNIVERSITY</p>
  <p>OF TWENTE</p>
  <p>CDD//UT.</p>
</div>

<hr> <!-- Separator for clarity -->

<!-- Logo Variant 3: Blue / Blue-Orange Split -->
<!-- Note the <span> around //UT. for split color -->
<div class="ut-logo-base logo-variant-3">
  <p>UNIVERSITY</p>
  <p>OF TWENTE</p>
  <p>CDD<span class="split">//UT.</span></p>
</div>

<hr> <!-- Separator for clarity -->

<!-- Logo Variant 4: Blue / Blue-Purple-Red Split -->
<!-- Note the <span> around //UT. for split color -->
<div class="ut-logo-base logo-variant-4">
  <p>UNIVERSITY</p>
  <p>OF TWENTE</p>
  <p>CDD<span class="split">//UT.</span></p>
</div>

"""


class LabelT(VEnum):
    """FIX: replaces danger with destructive"""

    def _generate_next_value_(name, start, count, last_values):
        return str2ukcls("label", name)

    primary = auto()
    secondary = auto()
    destructive = auto()


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


# --- globals ---


copyright_df_global: pl.DataFrame = retrieve_copyright_items()

db = database("db.sqlite3")

items = db.t.items

PORT = 8000
ROOT_URL = f"http://localhost:{PORT}"
DEFAULT_PER_PAGE = 15
MAX_CELL_LENGTH = 35


@dataclass
class Login:
    email: str
    pwd: str


@dataclass
class AppState:
    """
    data structure to hold the state of the app
    """

    page: int = 1
    per_page: int = DEFAULT_PER_PAGE
    sort_by: str | None = None
    sort_desc: bool = False
    filters: dict[str, str] = field(default_factory=dict)

    def update_from_req(
        self,
        request_params: dict[str, Any],
        form_data: FormData,
        auth_details: dict = None,
    ):
        self.page = int(request_params.get("page", self.page))
        self.per_page = int(request_params.get("per_page", self.per_page))
        new_sort_by = request_params.get("sort_by")
        if new_sort_by:
            self.sort_by = new_sort_by
        self.sort_desc = (
            str(request_params.get("sort_desc", str(self.sort_desc))).lower() == "true"
        )

        new_filters = {}
        all_filter_keys = [  # Define all possible filter keys
            "workflow_status",
            "status",
            "classification",
            "manual_classification",
            "faculty",
            "department",
            "course_name",  # Add any other text filter keys
        ]

        for key in all_filter_keys:
            filter_param_key = f"filter_{key}"
            if filter_param_key in request_params and request_params[filter_param_key]:
                # Add if the key exists and has a non-empty value
                new_filters[key] = request_params[filter_param_key]

        # Replace the old filters entirely
        self.filters = new_filters

        if auth_details:
            auth_faculty = auth_details.get("faculty")
            auth_role = auth_details.get("role")
            # Apply faculty filter ONLY if user is not admin and has a specific faculty
            if auth_faculty and auth_faculty != "all" and auth_role != "admin":
                self.filters["faculty"] = auth_faculty

    # Helper to generate params for HTMX links (excluding filters)
    def nav_params_dict(self) -> dict[str, Any]:
        return {
            "page": self.page,
            "per_page": self.per_page,
            "sort_by": self.sort_by or "",
            "sort_desc": str(self.sort_desc),
        }

    # Helper to generate filter params for HTMX links/forms
    def filter_params_dict(self) -> dict[str, str]:
        return {f"filter_{k}": v for k, v in self.filters.items()}

    # Helper to get combined params for full state links/forms if needed
    def all_params_dict(self) -> dict[str, Any]:
        return {**self.nav_params_dict(), **self.filter_params_dict()}

    # Helper to get filter dict compatible with get_filtered_sorted_df
    def get_active_filters(self) -> dict[str, str]:
        # Return a copy of the filters dict
        return self.filters.copy()


# --- auth ---


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
    FILTER_JS_HELPER,
)


# for static files, use the following regex to match the file extensions
reg_re_param(
    "static",
    "ico|gif|jpg|jpeg|webm|css|js|woff|png|svg|mp4|webp|ttf|otf|eot|woff2|txt|xml|html|pdf|md",
)

bware = Beforeware(
    before, skip=[r"/favicon\.ico", r"/static/.*", r"/imgs/.*", r".*\.css", "/login"]
)

app, rt = fast_app(
    before=bware,
    hdrs=(
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
    ),
    exts="loading-states",
    debug=True,
)

# Mount static files directory AFTER creating the app instance
app.mount("/imgs", StaticFiles(directory="."), name="images")

setup_toasts(app)

# --- constants / mappings ---

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


# --- component templates ---
def ItemDetailCard(
    title: str,
    *body_content: Any,
    card_id: str,
    col_span: int = 1,
    start_collapsed: bool = False,
    color: str = "[var(--ring)]",
    lazy_load_url: str | None = None,
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

    summary_classes = f"p-3 bg-{color}-100 hover:bg-{color}-200 cursor-pointer list-none flex items-center justify-between"
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
            "hx_indicator": f"#{card_id}-loading",
        }
    else:
        actual_body_content = body_content

    content_div = Div(*actual_body_content, id=content_target_id, cls=content_classes)
    summary_element = Summary(
        H5(title, cls="font-semibold text-sm m-0"),
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
    # NEW: Pass the calculated counts
    counts: Optional[dict[str, int]] = None,
    default_option_style: str = LabelT.secondary,
) -> FT:
    """Creates a group of styled checkboxes with counts and HTMX trigger."""
    selected_values = set(current_values.split("|")) if current_values else set()
    checkbox_items = []
    form_field_name = f"filter_{filter_key}"
    counts = counts or {}  # Ensure counts is a dict
    print(counts)
    sorted_options = sorted(options.items(), key=lambda item: item[0].lower())

    for value, style_info in sorted_options:
        is_checked = value in selected_values
        safe_value = "".join(c if c.isalnum() else "-" for c in value)
        checkbox_id = f"filter-cb-{filter_key}-{safe_value}"
        count_span_id = f"count-cb-{filter_key}-{safe_value}"  # ID for the count span
        style_class = (
            str(style_info)
            if isinstance(style_info, LabelT)
            else str(default_option_style)
        )

        badge_label = Label(value, cls=f"{style_class} badge-sm cursor-pointer")

        # Get the count for this specific option
        count = counts.get(value, "?")  # Default to '?' if count not found

        # Create the span to display the count
        count_span = Span(
            f"({count})",
            id=count_span_id,  # Assign the ID
            cls="text-xs text-base-content/70 ml-1",  # Styling
        )

        checkbox_input = fh.CheckboxX(
            name=form_field_name,
            value=value,
            id=checkbox_id,
            checked=is_checked,
            cls="checkbox checkbox-primary checkbox-xs align-middle mr-1",
        )

        # Wrap checkbox, badge, and count span in the label
        wrapper_label = fh.Label(
            checkbox_input,
            badge_label,
            count_span,  # Add the count span here
            fr=checkbox_id,
            cls="inline-flex items-center mr-2 mb-1",
        )
        checkbox_items.append(wrapper_label)

    return Div(cls="form-control w-full mb-3")(
        fh.Label(
            label_text,
            cls="label-text pb-1 text-xs font-medium text-base-content/90 block mb-1",
        ),
        Div(*checkbox_items, cls="flex flex-wrap gap-1"),
        # Optional: Add a hidden loading indicator for count updates
        # Span(cls="loading loading-xs loading-dots htmx-indicator-counts")
    )


# --- filter dataframe ---


def get_filtered_sorted_df(
    app_state: AppState,  # Use AppState object
) -> pl.DataFrame:
    """Applies filtering and sorting based on AppState."""
    df = copyright_df_global
    filters = app_state.get_active_filters()  # Get filters from AppState
    sort_by = app_state.sort_by
    sort_desc = app_state.sort_desc

    if filters:
        filter_expressions = []
        for col, value in filters.items():
            actual_col = col
            if value and actual_col in df.columns:
                try:
                    # --- START OR Logic ---
                    or_values = (
                        value.split("|")
                        if isinstance(value, str) and "|" in value
                        else [value]
                    )
                    or_expressions = []
                    # --- END OR Logic ---

                    for or_value in or_values:  # Iterate through potential OR values
                        or_value = or_value.strip()  # Clean up whitespace
                        if not or_value:
                            continue  # Skip empty strings resulting from split

                        if df[actual_col].dtype == pl.Utf8:
                            # Handle specific 'status' capitalization if needed
                            if col in ["status", "workflow_status", "classification"]:
                                or_expressions.append(
                                    pl.col(actual_col).str.to_lowercase()
                                    == or_value.lower()
                                )
                            else:
                                # General case-insensitive contains match for other text fields
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
                                or_expressions.append(
                                    pl.col(actual_col) == float(or_value)
                                )
                        # Add more type handling if needed

                    # --- Combine OR expressions ---
                    if len(or_expressions) > 1:
                        filter_expressions.append(pl.any_horizontal(or_expressions))
                    elif len(or_expressions) == 1:
                        filter_expressions.append(or_expressions[0])
                    # --- End Combine OR ---

                except Exception as e:
                    print(f"Filter warning on '{actual_col}' for value '{value}': {e}")

        if filter_expressions:
            try:
                # Combine all column filters with AND
                df = df.filter(pl.all_horizontal(filter_expressions))
            except Exception as e:
                print(f"Filter error applying filters {filters}: {e}")
                df = copyright_df_global  # Fallback
    if sort_by and sort_by in df.columns:
        try:
            df = df.sort(by=sort_by, descending=sort_desc, nulls_last=True)
        except Exception as e:
            print(f"Sort warning on '{sort_by}': {e}")
    return df


def _apply_filters_for_count(filters: dict[str, str]) -> int:
    """Helper to apply filters and return only the count. Optimized."""
    df = copyright_df_global  # Start with the full dataset

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

                    if df[actual_col].dtype == pl.Utf8:
                        if col in [
                            "status",
                            "workflow_status",
                            "classification",
                            "manual_classification",
                            "faculty",
                        ]:  # Assuming exact match for these
                            or_expressions.append(
                                pl.col(actual_col).str.to_lowercase()
                                == or_value.lower()
                            )
                        else:  # Contains match for others
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
                    # Add more types if needed

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
            # Apply filters and get count directly
            # Use lazy frame for potentially better optimization
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
        # No valid filters were generated, return total count
        return df.height


@rt(
    "/update_filter_counts", methods=["POST"]
)  # Use POST as it's triggered by a form change
async def update_filter_counts_endpoint(session: dict, request: Request):
    """Calculates filter counts based on submitted form state and returns OOB updates."""
    form_data = await request.form()
    request_params = dict(form_data)  # Use form data directly for state update

    # Create a temporary AppState based on the submitted form to calculate counts
    temp_app_state = session.get(
        "app_state", AppState()
    )  # Get base state (page/sort etc.)
    if isinstance(temp_app_state, dict):  # Ensure it's an instance if retrieved as dict
        temp_app_state = AppState(**temp_app_state)

    # Update the temporary state's filters based *only* on the incoming form data
    # (Don't pass auth_details here, just reflect the form's current state)
    temp_app_state.update_from_req(
        request_params, form_data
    )  # Pass form_data again for getlist

    # Calculate counts based on this potentially intermediate state
    counts = get_filter_counts(temp_app_state)

    # Generate OOB swap fragments for each count
    oob_fragments = []
    for filter_key, option_counts in counts.items():
        for option_value, count in option_counts.items():
            safe_value = "".join(c if c.isalnum() else "-" for c in option_value)
            count_span_id = f"count-cb-{filter_key}-{safe_value}"
            # Create the span with the count, targeting the correct ID
            count_span = Span(
                f"({count})",  # Display format
                id=count_span_id,  # The ID that will be targeted
                hx_swap_oob="innerHTML",  # Swap innerHTML of the target span
                cls="text-xs text-base-content/70 ml-1",  # Styling for the count
            )
            oob_fragments.append(count_span)

    # Return all OOB fragments
    return Div(*oob_fragments), HtmxResponseHeaders(
        reswap="none"
    )  # Prevent default swap


# Function to get counts for all options based on current selections
def get_filter_counts(app_state: AppState) -> dict[str, dict[str, int]]:
    """Calculates the counts for each filter option based on current state."""
    all_counts = {}
    current_filters = app_state.get_active_filters()

    # Define the options we need counts for (reuse from page_header_component)
    # Combine all options into one structure for iteration
    filter_options_map = {
        "workflow_status": {
            k: v.value for k, v in WORKFLOW_STYLES.items()
        },  # Value: Style map needed? No, just keys.
        "status": {k: v.value for k, v in STATUS_STYLES.items()},
        "classification": {
            **{v: LabelT.primary.value for v in PRIMARY_CLASSIFICATIONS},
            **{v: LabelT.secondary.value for v in SECONDARY_CLASSIFICATIONS},
            **{v: LabelT.destructive.value for v in DESTRUCTIVE_CLASSIFICATIONS},
        },
        "manual_classification": {
            **{v: LabelT.primary.value for v in PRIMARY_CLASSIFICATIONS},
            **{v: LabelT.secondary.value for v in SECONDARY_CLASSIFICATIONS},
            **{v: LabelT.destructive.value for v in DESTRUCTIVE_CLASSIFICATIONS},
        },
        "faculty": {
            f: LabelT.secondary.value for f in ["BMS", "EEMCS", "ET", "ITC", "TNW"]
        },
        # Add other filterable columns if they were checkbox groups (e.g., department?)
    }

    for filter_key, options in filter_options_map.items():
        all_counts[filter_key] = {}
        current_selections_for_key = (
            set(current_filters.get(filter_key, "").split("|"))
            if current_filters.get(filter_key)
            else set()
        )

        for option_value in options:
            # --- Calculate count IF this option were toggled ---
            temp_filters = current_filters.copy()
            temp_selections = current_selections_for_key.copy()

            # Simulate toggling this option
            if option_value in temp_selections:
                temp_selections.remove(option_value)  # Toggle OFF
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

    # Calculate counts for text filters maybe? (Less common UX) - Skipping for now.

    return all_counts


# --- save changes ---
async def store_item_changes(
    input_data: list[dict[str, str | int]] | dict[str, str | int],
) -> None:
    """
    For a list of dicts containing at least a material_id and updated fields,
    for each item, retrieve the full item data for the other fields from the copyright_df_global
    make sure the updated fields in the dict are not overwritten with the original values

    then send the new list of dicts to the database for update
    """
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

    await update_copyright_items(full_data_list, update_relations=False, overwrite=True)


@rt("/save_details", methods=["POST"])
async def save_item_details(
    request: Request,
    session,
    # Editable fields
    material_id: int,
    workflow_status: str = "",
    manual_classification: str = "",
    remarks: str = "",
):
    """
    Route to store an edited item back to the database.
    Update logic is in the store_item_changes function, and of course in db.update.update_copyright_items.
    Uses the material_id as a primary key to identify the item.
    For now, only three field can be updated: workflow_status, manual_classification and remarks.
    This route should only be called from the item detail view at the moment.
    Might be extended to do bulk updates from the table view in the future.
    """
    global copyright_df_global
    app_state_dict = session.get("app_state", {})
    app_state = AppState(**app_state_dict)  # Convert dict back to instance

    print(f"Saving changes for material_id: {material_id}")
    update_data = {
        "material_id": material_id,
        "workflow_status": workflow_status,
        "manual_classification": manual_classification,
        "remarks": remarks,
    }

    oob_grid_swap = Div()  # Default empty div

    try:
        await store_item_changes(update_data)

        print("Reloading global DataFrame after save...")
        copyright_df_global = retrieve_copyright_items()
        print("Global DataFrame reloaded.")

        # Re-filter and slice for the current page view
        filtered_sorted_df = get_filtered_sorted_df(app_state)  # Pass the state object
        total_filtered_rows = filtered_sorted_df.height
        total_pages = (
            math.ceil(total_filtered_rows / app_state.per_page)
            if app_state.per_page > 0
            else 1
        )
        app_state.page = max(
            1, min(app_state.page, total_pages if total_pages > 0 else 1)
        )
        session["app_state"] = asdict(app_state)  # Convert back to dict

        offset = (app_state.page - 1) * app_state.per_page
        df_slice = filtered_sorted_df.slice(
            offset, per_page=app_state.per_page
        )  # Use per_page from state

        # Render the grid component HTML using AppState
        updated_grid_component = render_data_grid_component(
            df_slice=df_slice,
            app_state=app_state,  # Pass the whole state object
            total_filtered_rows=total_filtered_rows,
            total_pages=total_pages,
        )

        # --- Create OOB Swap Div for the Grid ---
        oob_grid_swap = Div(
            to_xml(updated_grid_component),
            hx_swap_oob="outerHTML:#data-grid-component",
        )

        add_toast(session, f"Successfully updated {material_id}", "success")

    except Exception as e:
        print(f"Error saving changes for {material_id}: {e}")
        traceback.print_exc()
        add_toast(session, f"Error saving changes for {material_id}: {e}", "error")

    return oob_grid_swap, HtmxResponseHeaders(
        reswap="none",
    )


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
    all_contact_emails = []
    courses = item_data.get("courses") or []
    for course in courses:
        persons = course.get("persons") or []
        for person in persons:
            if person.get("role") == "contact":
                contact_persons.append(person)
                if person.get("email"):
                    all_contact_emails.append(person["email"])

    contact_info_items = []
    emails_str = ";".join(all_contact_emails)
    copy_js = f"navigator.clipboard.writeText('{emails_str}');"
    contact_info_items.append(
        Div(
            fh.Button(
                UkIcon("copy", cls="w-4 h-4 mr-1"),
                "Copy all email addresses",
                cls="btn btn-primary btn-sm btn-block",
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
                "hx_indicator": "#modal-loading-indicator",  # Add indicator for modal load
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
) -> FT:
    """Renders the header area with Title, User Info, and a Sidebar for Filters."""
    current_filters = app_state.get_active_filters()

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
    }
    manual_classification_options = classification_options
    faculty_options = {
        f: LabelT.secondary for f in ["BMS", "EEMCS", "ET", "ITC", "TNW"]
    }

    # --- Filter Checkbox Groups ---
    filter_checkbox_groups = [
        create_checkbox_filter_group(
            filter_key="workflow_status",
            label_text="Workflow Status",
            options=workflow_options,
            current_values=current_filters.get("workflow_status"),
            counts=filter_counts.get("workflow_status"),  # Pass counts for this key
        ),
        create_checkbox_filter_group(
            filter_key="status",
            label_text="Status",
            options=status_options,
            current_values=current_filters.get("status"),
            counts=filter_counts.get("status"),  # Pass counts
        ),
        create_checkbox_filter_group(
            filter_key="classification",
            label_text="ML Classification",
            options=classification_options,
            current_values=current_filters.get("classification"),
            counts=filter_counts.get("classification"),  # Pass counts
        ),
        create_checkbox_filter_group(
            filter_key="manual_classification",
            label_text="Manual Classification",
            options=manual_classification_options,
            current_values=current_filters.get("manual_classification"),
            counts=filter_counts.get("manual_classification"),  # Pass counts
        ),
    ]
    # Conditionally add Faculty filter group
    is_admin = user_details.get("role") == "admin"
    user_faculty = user_details.get("faculty")
    # Show faculty filter if admin OR user has no specific faculty assigned
    if is_admin or not user_faculty or user_faculty == "all":
        filter_checkbox_groups.append(
            create_checkbox_filter_group(
                filter_key="faculty",
                label_text="Faculty",
                options=faculty_options,
                current_values=current_filters.get("faculty"),
            )
        )

    # --- Filter Inputs (Text - Remaining) ---
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
        actual_col_name = col
        if actual_col_name not in copyright_df_global.columns:
            continue
        label = f"{col.replace('_', ' ').title()}"
        if "department" in label.lower():
            label = "Programme"

        text_filter_inputs.append(
            LabelInput(
                label,
                name=f"filter_{col}",
                value=current_filters.get(col, ""),
                placeholder="Filter...",
                id=f"filter-input-{col}",
                input_cls="input input-bordered input-xs w-full focus:input-primary",
                label_cls="label-text pb-1 text-xs font-medium text-base-content/90",
                cls="form-control w-full",
            )
        )

    # --- Hidden State Inputs ---
    hidden_state_inputs = [
        Input(type="hidden", name=k, value=str(v))
        for k, v in app_state.nav_params_dict().items()
    ]
    reset_dict = app_state.nav_params_dict()
    reset_dict["page"] = 1  # Reset page on filter reset
    # --- Filter Form (Content of the Sidebar) ---
    filter_form = Form(
        H4("Standard Filters", cls="text-md font-semibold mb-3"),
        # Use checkbox groups instead of dropdowns/toggles
        *filter_checkbox_groups,
        H4("Specific Filters", cls="text-md font-semibold mb-3 mt-4"),
        Div(*text_filter_inputs, cls="grid grid-cols-1 gap-y-3"),
        *hidden_state_inputs,
        # --- Action Buttons (Keep as is) ---
        Div(cls="mt-6 space-y-2")(
            Button(
                "Apply Filters",
                type="button",
                # REMOVE hx-vals attribute
                # Add an ID
                id="apply-filters-btn",
                hx_get=data_grid.to(),  # Base URL remains
                hx_target="#data-grid-component",
                # REMOVE hx-include="closest form" (not needed)
                hx_indicator="#grid-loading-indicator",
                **{"@click": "slideOverOpen=false"},
                cls="btn btn-sm btn-primary w-full",
            ),
            Button(
                "Reset Filters",
                type="button",
                hx_get=data_grid.to(action="reset", **reset_dict),
                hx_target="#content-area",
                hx_indicator="#grid-loading-indicator",
                **{"@click": "slideOverOpen=false"},
                cls="btn btn-sm btn-ghost w-full",
            ),
        ),
        id="filter-form",
        cls="p-4",
        hx_post=update_filter_counts_endpoint.to(),
        # Trigger specifically on checkbox changes within this form
        hx_trigger="change from:input[type=checkbox] delay:300ms",
        # Include data from the form that changed
        hx_include="this",  # Include data from the form itself
        # Target doesn't really matter for OOB swaps, but needs to be valid
        hx_target="this",
        hx_swap="none",  # Rely on OOB swaps returned from the endpoint
    )
    # --- Sidebar Structure (Alpine.js) ---
    sidebar_component = Div(
        x_data="{ slideOverOpen: false }", cls="relative z-50 w-auto h-auto"
    )(
        # Sidebar Trigger Button (keep as is)
        Button(
            UkIcon("filter", cls="mr-1"),
            "Filters",
            **{"@click": "slideOverOpen=true"},
            cls="btn btn-sm btn-outline btn-primary ml-4",
        ),
        # Teleport the sidebar content to the body
        Template(x_teleport="body")(
            # Outermost container for the sidebar, managed by x-show
            Div(
                x_show="slideOverOpen",
                **{"@keydown.window.escape": "slideOverOpen=false"},
                cls="relative z-[99]",
                x_cloak=True,
            )(  # Add x-cloak
                # Overlay
                Div(
                    x_show="slideOverOpen",
                    x_transition_opacity_duration_600ms=True,
                    **{"@click": "slideOverOpen = false"},
                    cls="fixed inset-0 bg-black bg-opacity-10",
                ),  # Overlay needs transition too
                # Sidebar Container
                Div(cls="fixed inset-0 overflow-hidden")(
                    Div(cls="absolute inset-0 overflow-hidden")(
                        # Positioning from the Left
                        Div(cls="fixed inset-y-0 left-0 flex max-w-full pr-10")(
                            # The actual sliding panel
                            Div(
                                x_show="slideOverOpen",  # Let x-show control visibility
                                **{"@click.away": "slideOverOpen = false"},
                                x_transition_enter="transform transition ease-in-out duration-500 sm:duration-700",
                                x_transition_enter_start="-translate-x-full",
                                x_transition_enter_end="translate-x-0",
                                x_transition_leave="transform transition ease-in-out duration-500 sm:duration-700",
                                x_transition_leave_start="translate-x-0",
                                x_transition_leave_end="-translate-x-full",
                                cls="w-screen max-w-md",
                                # Remove inline style: style="display: none;"
                            )(
                                # Sidebar Content Area
                                # Use overflow-y-auto (addressed in point 4)
                                Div(
                                    cls="flex flex-col h-full py-5 overflow-y-auto bg-base-100 border-r border-base-300 shadow-lg"
                                )(
                                    # Sidebar Header (keep as is)
                                    Div(cls="px-4 sm:px-5")(
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
                                    # Main Filter Form Area (keep as is)
                                    Div(cls="relative flex-1 mt-5")(filter_form),
                                )
                            )
                        )
                    )
                ),
            )
        ),
    )

    # --- Header Structure ---
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
            # Dashboard Title and Last Update
            Div(
                H4("Copyright Data Dashboard", cls="text-2xl font-bold text-primary"),
                H6(
                    f"Last update from Qlik: {last_update_time}",
                    cls="text-xs text-base-content/70 mt-1",
                ),
            ),
            # Sidebar Trigger and User Info/Logout
            Div(cls="flex items-center space-x-4")(
                sidebar_component,  # Add the sidebar trigger button here
                # User Avatar/Greeting (Keep existing)
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
    ]
    for col in cols_in_header:
        header_text = "ID" if col == "material_id" else col.replace("_", " ").title()

        sort_indicator_icon, next_sort_desc_str = "", "False"
        is_current_sort_col = col == app_state.sort_by

        if is_current_sort_col:
            next_sort_desc_bool = not app_state.sort_desc
            next_sort_desc_str = str(next_sort_desc_bool)
            sort_indicator_icon = "arrow-down" if app_state.sort_desc else "arrow-up"
        else:
            next_sort_desc_str = "False"

        sort_link_params = {
            "page": app_state.page,  # Keep current page for sorting
            "per_page": app_state.per_page,
            "sort_by": col,
            "sort_desc": next_sort_desc_str,
            **app_state.filter_params_dict(),  # Include current filters
        }
        header_content = Button(
            header_text,
            UkIcon(
                sort_indicator_icon, cls="inline-block ml-1.5 w-3 h-3 stroke-current"
            )
            if sort_indicator_icon
            else "",
            hx_get=data_grid.to(**sort_link_params),
            hx_target="#data-grid-component",
            hx_indicator="#grid-loading-indicator",
            # Add hx_include to send filter form data if needed, though params might be enough
            # hx_include="#filter-form", # Maybe not needed if params cover filters
            cls=f"{ButtonT.ghost} text-xs uppercase tracking-wider p-1.5 h-auto min-h-0 font-bold text-base-content/70 hover:text-primary transition-colors duration-150",
        )
        th_class = f"col-{col.replace('_', '-')}"
        header_cells.append(Th(header_content, cls=f"px-4 py-2 {th_class}"))

    header = Thead(Tr(*header_cells))

    body = Tbody(
        *render_table_rows(df_slice, app_state),
        id="data-table-body",
    )

    table_wrapper = Div(
        Table(header, body, role="grid", id="data-table", cls="table table-sm w-full"),
        id="table-wrapper",
        cls="overflow-x-auto border border-base-300 rounded-lg shadow-sm bg-base-100",
    )

    # --- Pagination ---
    # Pass AppState to render_paginatlabel_text=ion
    pagination_html = render_pagination(app_state, total_pages)

    # --- Final Component ---
    return Div(
        P(
            Strong(str(total_filtered_rows)),
            Span(" items found", cls="text-base-content text-bold"),
            cls=" mb-2 text-sm",
        ),
        pagination_html,
        Div(
            table_wrapper,
            cls="htmx-loading:opacity-50 htmx-loading:transition-opacity htmx-loading:duration-300 ease-in-out",
        ),
        Div(id="grid-loading-indicator", cls="htmx-indicator text-center p-4")(
            "Loading..."
        ),
        id="data-grid-component",
        cls="flex flex-col text-center justify-center",  # Existing classes
    )


def render_pagination(
    app_state: AppState,  # Accept AppState object
    total_pages: int,
) -> FT:
    """Renders pagination controls using AppState."""
    if total_pages <= 1:
        return Div(cls="h-12")  # No pagination needed

    current_page = app_state.page
    # Base parameters for pagination links include current sort and filters
    base_params = {
        "per_page": app_state.per_page,
        "sort_by": app_state.sort_by or "",
        "sort_desc": str(app_state.sort_desc),
        **app_state.filter_params_dict(),  # Include current filters
    }

    pagination_items = []

    # --- Previous Button ---
    prev_disabled, prev_page = current_page <= 1, max(1, current_page - 1)
    prev_link_params = {"page": prev_page, **base_params}
    prev_attrs = {
        "hx_get": data_grid.to(**prev_link_params),
        "hx_target": "#data-grid-component",
        "hx_indicator": "#grid-loading-indicator",
        "role": "button",
        # Add hx_include if filters aren't reliably passed via params alone
        # "hx_include": "#filter-form"
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

    # --- Page Indicator ---
    pagination_items.append(
        Span(
            f"Page {current_page} of {total_pages}",
            cls="underline decoration-pink-500 bg-base-200 text-base-content font-semibold mr-5 ml-5 mb-1",
        )  # Adjusted styling/spacing
    )

    # --- Next Button ---
    next_disabled, next_page = (
        current_page >= total_pages,
        min(total_pages, current_page + 1),
    )
    next_link_params = {"page": next_page, **base_params}
    next_attrs = {
        "hx_get": data_grid.to(**next_link_params),
        "hx_target": "#data-grid-component",
        "hx_indicator": "#grid-loading-indicator",
        "role": "button",
        # "hx_include": "#filter-form" # If needed
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

    return Div(
        *pagination_items,
        aria_label="pagination",
        class_="flex justify-center items-center space-x-5 gap-4 pt-4 pb-2",
    )  # Keep existing layout


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
            return Span(display_text, cls="text-base-content/70 text-sm"), "span"
        elif val_str_norm in PRIMARY_CLASSIFICATIONS:
            style = LabelT.primary
        elif val_str_norm in SECONDARY_CLASSIFICATIONS:
            style = LabelT.secondary
        elif val_str_norm in DESTRUCTIVE_CLASSIFICATIONS:
            style = LabelT.destructive
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


@rt("/data")
async def data_grid(session: dict, request: Request):
    app_state_dict = session.get("app_state", {})
    app_state = AppState(**app_state_dict)

    query_params = dict(request.query_params)
    form_data = FormData()
    is_post = request.method == "POST"
    content_type = request.headers.get("content-type", "")
    if (
        is_post
        or "multipart/form-data" in content_type
        or "application/x-www-form-urlencoded" in content_type
    ):
        with contextlib.suppress(Exception):
            form_data = await request.form()
    request_params = {**query_params, **dict(form_data)}
    auth_details = session.get("auth", {})

    is_reset_action = request_params.get("action") == "reset"
    if is_reset_action:
        print("Reset action detected, clearing filters.")
        app_state.filters = {}  # Clear state filters FIRST
        request_params.pop("action", None)
        keys_to_remove = [k for k in request_params if k.startswith("filter_")]
        for k in keys_to_remove:
            request_params.pop(k, None)
        form_data = FormData()  # Clear form data as well for update_from_req

    # Pass both combined params AND original form_data for getlist()
    app_state.update_from_req(request_params, form_data, auth_details)

    # --- Recalculate, Validate, Store state ---
    filtered_sorted_df = get_filtered_sorted_df(app_state)
    total_filtered_rows = filtered_sorted_df.height
    total_pages = (
        math.ceil(total_filtered_rows / app_state.per_page)
        if app_state.per_page > 0
        else 1
    )
    app_state.page = max(1, min(app_state.page, total_pages if total_pages > 0 else 1))
    session["app_state"] = asdict(app_state)
    filter_counts = get_filter_counts(app_state)

    # --- Prepare Slice and Render Components ---
    offset = (app_state.page - 1) * app_state.per_page
    df_slice = filtered_sorted_df.slice(offset, app_state.per_page)
    grid_component = render_data_grid_component(
        df_slice=df_slice,
        app_state=app_state,
        total_filtered_rows=total_filtered_rows,
        total_pages=total_pages,
    )
    header_component = page_header_component(auth_details, app_state, filter_counts)

    # --- Determine Return Value ---
    is_htmx = request.headers.get("hx-request", "false").lower() == "true"

    if not is_htmx:
        header_component = page_header_component(auth_details, app_state, filter_counts)

        modal_placeholder = Dialog(
            id="modal-placeholder", cls="modal modal-bottom sm:modal-middle"
        )
        modal_trigger_script = Script("""
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

        modal_interaction_script = Script(f"""
            // Function to update a pill's appearance and hidden input value
            function updatePill(fieldName, newValue, newText, newStyleClass) {{
                const pillElement = document.getElementById(`pill-display-${{fieldName}}`);
                const inputElement = document.getElementById(`input-${{fieldName}}`);
                if (pillElement && inputElement) {{
                    // Update hidden input
                    inputElement.value = newValue;
                    // Update visible pill text
                    pillElement.textContent = newText;
                    // Update visible pill style (remove old, add new)
                    // Assumes style classes are like 'uk-label-primary', 'uk-label-secondary', etc.
                    pillElement.classList.remove('uk-label-primary', 'uk-label-secondary', 'uk-label-destructive');
                    if (newStyleClass) {{ // Add new style if provided
                    pillElement.classList.add(newStyleClass);
                    }} else {{ // Fallback if no specific style maps (e.g., for 'N/A')
                        pillElement.classList.add('uk-label-secondary'); // Or your default
                    }}
                    // Close the dropdown (assuming uk-drop is used)
                    const drop = UIkit.drop(pillElement.closest('[uk-drop]'));
                    if (drop) {{ drop.hide(false); }}
                    markDirty(); // Mark form as dirty
                }} else {{
                    console.error(`Cannot find pill or input elements for ${{fieldName}}`);
                }}
            }}
            function markDirty() {{
                const indicator = document.getElementById('save-indicator');
                const saveButton = document.getElementById('modal-save-btn');
                if (indicator) indicator.classList.remove('hidden');
                if (saveButton) saveButton.disabled = false; // Make sure it's enabled
            }}

            // Function to reset the modal form fields to original values
            function resetModalForm() {{
                console.log('Resetting modal form'); // Debug log
                const form = document.getElementById('modal-details-form');
                if (!form) return;

                const pillInputs = form.querySelectorAll('input[data-original-value][id^="input-"]'); // Target only pill inputs
                pillInputs.forEach(input => {{
                    const originalValue = input.dataset.originalValue;
                    const originalText = input.dataset.originalText || originalValue; // Fallback text
                    const originalStyle = input.dataset.originalStyle || '{str(DEFAULT_PILL_STYLE)}'; // Fallback style
                    const fieldName = input.id.replace('input-', '');

                    // Reset hidden input
                    input.value = originalValue;

                    // Reset visible pill
                    const pillElement = document.getElementById(`pill-display-${{fieldName}}`);
                    if (pillElement) {{
                        pillElement.textContent = originalText;
                        pillElement.classList.remove('uk-label-primary', 'uk-label-secondary', 'uk-label-destructive');
                        if (originalStyle && originalStyle.startsWith('uk-label-')) {{ // Ensure it's a valid style class
                            pillElement.classList.add(originalStyle);
                        }} else {{
                             pillElement.classList.add('{str(DEFAULT_PILL_STYLE)}'); // Default
                        }}
                    }}
                }});

                const remarksTextarea = form.querySelector('#modal_remarks');
                if (remarksTextarea && typeof remarksTextarea.dataset.originalValue !== 'undefined') {{ // Check if attribute exists
                    remarksTextarea.value = remarksTextarea.dataset.originalValue;
                }}

                // Hide save indicator and disable save button
                const indicator = document.getElementById('save-indicator');
                const saveButton = document.getElementById('modal-save-btn');
                if (indicator) {{
                    indicator.classList.add('hidden');
                }}
                if (saveButton) {{
                    saveButton.disabled = true;
                }}
            }}
        """)

        return (
            Title("CDD//UT"),
            Div(
                Div(header_component, grid_component, id="content-area"),
                id="page-container",
            ),
            modal_placeholder,
            modal_trigger_script,
            modal_interaction_script,
            FILTER_JS_LISTENER,
        )

    else:
        # --- HTMX Partial Update ---
        if is_reset_action:
            # Reset needs to update the sidebar form, so return header+grid
            print("HTMX Reset: Returning Header + Grid")
            header_component = page_header_component(
                auth_details, app_state, filter_counts
            )
            return (header_component, grid_component)
        else:
            # Apply Filters (or sort/paginate) only needs to update the grid
            print("HTMX Apply/Sort/Page: Returning Grid Only")
            return grid_component


@rt("/modal/{material_id:int}")
async def show_item_details(session: dict, material_id: int):  # Simplified signature
    """Fetches data and returns structured INNER content for the modal dialog."""

    app_state_dict = session.get("app_state", {})
    app_state = AppState(**app_state_dict)  # Convert dict back to instance

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
            ):  # If no style applied, ensure badge-sm is used for consistency if it's a Label
                if isinstance(content_component, fh.fastcore.xml.Gen):
                    if (
                        content_component.tag == "span"
                        and "uk-label" in content_component.attrs.get("cls", "")
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
        filtered_sorted_df = get_filtered_sorted_df(app_state)
        ordered_ids = filtered_sorted_df.get_column("material_id").to_list()

        try:
            current_index = ordered_ids.index(material_id)
        except ValueError:
            current_index = -1  # Item not found in current filtered/sorted list

        prev_id = ordered_ids[current_index - 1] if current_index > 0 else None
        next_id = (
            ordered_ids[current_index + 1]
            if current_index != -1 and current_index < len(ordered_ids) - 1
            else None
        )

        # --- Fetch Item Data  ---
        nested_data_list = retrieve_osiris_data([material_id])
        if not nested_data_list:  # Handle case where item data isn't found
            raise ValueError(f"Material ID {material_id} not found in Osiris data.")
        item_data = nested_data_list[0]
        if "faculty_id" in item_data:
            item_data["faculty"] = item_data.pop("faculty_id")  # Alias faculty_id

        manual_classification_options = {}
        for val in PRIMARY_CLASSIFICATIONS:
            manual_classification_options[val] = LabelT.primary
        for val in SECONDARY_CLASSIFICATIONS:
            manual_classification_options[val] = LabelT.secondary
        for val in DESTRUCTIVE_CLASSIFICATIONS:
            manual_classification_options[val] = LabelT.destructive
        if "onbekend" not in manual_classification_options:
            manual_classification_options["onbekend"] = LabelT.secondary

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
                UkIcon("file-key-2", cls="mr-1"),
                str(material_id),
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
        indicator_attrs = {"hx_indicator": "#modal-loading-indicator"}

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
                    **indicator_attrs,
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
                    **indicator_attrs,
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
                    # Loading indicator for modal transitions (prev/next/save)
                    Div(
                        Span("Loading...", cls="loading loading-lg"),
                        id="modal-loading-indicator",
                        cls="htmx-indicator absolute inset-0 bg-base-100/50 flex items-center justify-center z-50",
                    ),
                    # Form for saving details
                    Form(
                        # Essential hidden field for the save endpoint
                        Input(type="hidden", name="material_id", value=material_id),
                        # --- REMOVE HIDDEN STATE INPUTS ---
                        # *[Input(type="hidden", name=k, value=v) for k, v in app_state.filter_params_dict().items()],
                        # *[Input(type="hidden", name=k, value=v) for k, v in app_state.nav_params_dict().items()],
                        modal_cards_grid,  # The main content grid
                        id="modal-details-form",
                        hx_post=save_item_details.to(),  # Save endpoint doesn't need state params in URL
                        hx_indicator="#modal-loading-indicator",
                        # Consider hx_target="body" and hx_swap="none" if only relying on OOB swaps from save_item_details
                        # Or target a specific small element inside the modal for status messages if needed
                        # hx_target="#save-status-message", hx_swap="innerHTML"
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
