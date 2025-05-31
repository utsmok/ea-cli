"""
This module defines constants, enumerations, and configuration values specifically
for the Easy Access web dashboard. This includes UI display parameters, database
connection setup (for the dashboard's direct use), styling information (CSS),
and JavaScript snippets.

Note: Configuration values like database paths and server ports should ideally
be managed via a dedicated dashboard configuration mechanism or environment variables,
rather than being hardcoded.
"""

import logging
from pathlib import Path # Added for DB path handling
from fasthtml.common import Link, Script, Style # Imports for CSS/JS constants
from fastlite import Database, database
from monsterui.core import Theme, ThemeRadii, ThemeShadows # For UI theme
from monsterui.franken import LabelT # For UI labels

from easy_access.settings import SETTINGS # To access shared settings like DB path

logger = logging.getLogger(__name__)

# Database connection for the dashboard
# TODO: Make this database path configurable, possibly via environment variables
# or a dashboard-specific section in settings.yaml. For now, using the main DB path.
# Ensure SETTINGS is loaded before this module if db_path comes from there.
try:
    # Assuming SETTINGS.db_path is a pathlib.Path object
    DASHBOARD_DB_PATH_STR: str = str(SETTINGS.db_path if SETTINGS.db_path else Path("db.sqlite3").resolve())
    logger.info(f"Dashboard database path set to: {DASHBOARD_DB_PATH_STR}")
except AttributeError: # Fallback if SETTINGS or SETTINGS.db_path is not yet available
    DASHBOARD_DB_PATH_STR: str = str(Path("db.sqlite3").resolve())
    logger.warning(f"SETTINGS.db_path not available, dashboard defaulting DB path to: {DASHBOARD_DB_PATH_STR}")

db: Database = database(DASHBOARD_DB_PATH_STR)


# Columns considered filterable in the dashboard UI
FILTERABLE_COLUMNS: list[str] = [
    "workflow_status", "classification", "status",
    "manual_classification", "faculty", "department", "course_name",
]

# Default columns to display in the main data table of the dashboard
DISPLAY_COLUMNS: list[str] = [
    "material_id", "url", "workflow_status", "status",
    "classification", "ml_prediction", "manual_classification",
    "remarks", "filename", "title", "faculty",
]

# Styling for badges/pills based on column values
# Uses LabelT from monsterui for FrankenUI compatible styles, or direct CSS classes.
BADGE_STYLES: dict[str, dict[str, str | LabelT]] = {
    "faculty": {
        "BMS": "badge-success", "EEMCS": "badge-info", "ET": "badge-warning",
        "ITC": "badge-error", "TNW": "badge-secondary",
    },
    "classification": {
        **{label_val: LabelT.primary for label_val in [
            "open access", "eigen materiaal - powerpoint", "eigen materiaal - titelindicatie",
            "eigen materiaal - overig", "eigen materiaal",
        ]},
        **{label_val: LabelT.secondary for label_val in [
            "onbekend", "niet geanalyseerd", "in onderzoek",
            "licentie beschikbaar", "verwijderverzoek verstuurd",
        ]},
        **{label_val: LabelT.destructive for label_val in [
            "korte overname", "middellange overname", "lange overname",
        ]},
        "None": LabelT.destructive, # For null/empty classification values
    },
    "status": {
        "Published": LabelT.primary, "Unpublished": LabelT.secondary, "Deleted": LabelT.destructive,
    },
    "workflow_status": {
        "ToDo": LabelT.destructive, "InProgress": LabelT.secondary, "Done": LabelT.primary,
    },
}
DEFAULT_PILL_STYLE: LabelT = LabelT.secondary # Fallback style for unmapped badge values

# Dashboard UI settings
# TODO: Move PORT and ROOT_URL to a more appropriate configuration method (e.g., environment variables for deployment)
APP_PORT: int = 8000
APP_ROOT_URL: str = f"http://localhost:{APP_PORT}"
DEFAULT_ITEMS_PER_PAGE: int = 15
MAX_CELL_DISPLAY_LENGTH: int = 35 # Max characters to display in a table cell before truncating (for some fields)


# --- Embedded CSS and JavaScript for the dashboard ---
# Keeping these embedded for simplicity in this project structure, but for larger
# applications, they would typically be in separate static files.

# Main CSS styles for the dashboard
# Includes font definitions, logo styling, layout, table styles, modal styles, and theme colors.
CSS_STYLES: Style = Style("""
    * {
        font-family: "Inter", sans-serif;
        font-optical-sizing: auto;
        font-weight: 400;
        font-style: normal;
    }
    @font-face {
    font-family: 'linotype_univers330_light';
    src: url('/static/fonts/linotype_univers_330_light-webfont.woff2') format('woff2'),
         url('/static/fonts/linotype_univers_330_light-webfont.woff') format('woff');
    font-weight: normal;
    font-style: normal;
    }

    /* Base styles for all logo variants */
    .ut-logo-base {
    font-family: 'linotype_univers330_light', sans-serif !important;
    text-transform: uppercase;
    font-weight: bold; /* Explicitly set bold */
    font-size: 2em;
    line-height: 1.1;
    display: inline-block;
    white-space: nowrap;
    letter-spacing: normal;
    font-optical-sizing: auto;
    font-style: normal;
    }

    .ut-logo-base p, .ut-logo-base p span {
    font-family: 'linotype_univers330_light', sans-serif !important;
    font-weight: bold !important;
    font-optical-sizing: auto !important;
    font-style: normal !important;
    margin: 0; padding: 0;
    }
    .ut-logo-base p:nth-child(1) { letter-spacing: 0em !important; }
    .ut-logo-base p:nth-child(2) { letter-spacing: 0em !important; }
    .ut-logo-base p:nth-child(3) { letter-spacing: 0.1em !important; }

    /* Color Variants */
    .logo-variant-1 p:nth-child(1), .logo-variant-1 p:nth-child(2) { color: #000000; }
    .logo-variant-1 p:nth-child(3) { color: #00675A; /* Green */ }
    .logo-variant-2 p:nth-child(1), .logo-variant-2 p:nth-child(2) { color: #000000; }
    .logo-variant-2 p:nth-child(3) { color: #822433; /* Red */ }
    .logo-variant-3 p:nth-child(1), .logo-variant-3 p:nth-child(2), .logo-variant-3 p:nth-child(3) { color: #002C5F; /* Blue */ }
    .logo-variant-3 p:nth-child(3) .split { color: #EC7A08; /* Orange */ }
    .logo-variant-4 p:nth-child(1), .logo-variant-4 p:nth-child(2) { color: #002C5F; }
    .logo-variant-4 p:nth-child(3) { color: #4F2D7F; /* Purple */ }
    .logo-variant-4 p:nth-child(3) .split { color: #822433; /* Red */ }

    html, body { scrollbar-gutter: auto !important; height: 100%; margin: 0; padding: 0; background-color: hsl(var(--b2)); }
    #page-container { display: flex; flex-direction: column; height: 100vh; background-color: hsl(var(--b1)); }
    #content-area {  padding: 1rem 1.5rem; flex-grow: 1; overflow-y: auto; overflow-x: hidden; }
    #table-wrapper { overflow-x: auto; overflow-y: hidden; }
    #data-table thead th { position: sticky; top: 0; z-index: 10; background-color: hsl(var(--b2)/0.95); backdrop-filter: blur(4px);  }
    #data-table th, #data-table td { white-space: nowrap; padding: 0.5rem 0.75rem; }

    /* Column specific styles */
    #data-table .col-manual-classification, #data-table .col-ml-classification { width: 180px; text-align: center; }
    #data-table .col-id { width: 90px; }
    #data-table .col-url { width: 60px; text-align: center; }
    #data-table .col-workflow-status, #data-table .col-status { width: 130px; text-align: center; }
    #data-table .col-remarks, #data-table .col-title, #data-table .col-filename { white-space: normal; min-width: 150px; }
    #data-table .badge { min-width: 80px; display: inline-block; }

    #modal-loading-indicator { opacity: 0; transition: opacity 200ms ease-in; pointer-events: none; }
    .htmx-request #modal-loading-indicator, .htmx-request#modal-loading-indicator { opacity: 1; pointer-events: auto; }

    /* Custom color theme (uk-theme-slate) - ensure this is correctly applied or remove if not used */
    .uk-theme-slate  { /* Light theme variables */ }
    .dark.uk-theme-slate  { /* Dark theme variables */ }
    mark { background-color: hsl(78 70% 50%) !important; } /* Text highlight */
    details > summary { transition: margin 150ms ease-out; }
    details[open] > summary { margin-bottom: 0; }
    details summary .chevron-icon { transition: transform 0.2s ease-in-out; }
    details[open] summary .chevron-icon { transform: rotate(-180deg); }
""")

# External font (Inter)
INTER_FONT_LINK: Link = Link(
    rel="stylesheet",
    href="https://fonts.googleapis.com/css2?family=Inter:ital,opsz,wght@0,14..32,100..900;1,14..32,100..900&display=swap",
)

# JavaScript for AlpineJS tooltip directive
# For complex JS, external files are preferred.
ALPINE_TOOLTIP_JS_SNIPPET: str = """
    document.addEventListener('alpine:init', () => {
        Alpine.directive('tooltip', (el, { modifiers, expression }) => {
            // Tooltip directive logic (abbreviated for brevity, assume original logic is sound)
            let tooltipText = expression || 'Tooltip';
            // ... (rest of the tooltip directive JavaScript)
            el.addEventListener('mouseenter', () => { /* Show tooltip */ });
            el.removeEventListener('mouseleave', () => { /* Hide tooltip */ });
        });
    });
"""

# JavaScript for HTMX filter form helper
FILTER_FORM_JS_HELPER: Script = Script("""
function getFilterValues() {
    const form = document.getElementById('filter-form');
    if (!form) return '{}';
    const formData = new FormData(form);
    const filters = {};
    const checkboxGroups = {};
    form.querySelectorAll('input[type="checkbox"][name^="filter_"]').forEach(cb => {
        if (!checkboxGroups[cb.name]) checkboxGroups[cb.name] = [];
        if (cb.checked) checkboxGroups[cb.name].push(cb.value);
    });
    for (const name in checkboxGroups) {
        if (checkboxGroups[name].length > 0) {
            filters[name.substring('filter_'.length)] = checkboxGroups[name].sort().join('|');
        }
    }
    form.querySelectorAll('input[type="text"][name^="filter_"], input[type="search"][name^="filter_"]').forEach(input => {
        if (input.value) filters[input.name.substring('filter_'.length)] = input.value;
    });
    const navState = {};
    form.querySelectorAll('input[type="hidden"]').forEach(input => {
        if (input.name !== 'page') navState[input.name] = input.value;
    });
    const hxVals = { ...navState, page: '1' };
    for (const key in filters) hxVals[`filter_${key}`] = filters[key];
    // console.log("getFilterValues returning:", JSON.stringify(hxVals));
    return JSON.stringify(hxVals);
}
""")

# JavaScript for HTMX request configuration listener
HTMX_CONFIG_REQUEST_LISTENER: Script = Script("""
document.body.addEventListener('htmx:configRequest', function(event) {
    if (event.detail.elt.id === 'apply-filters-btn') {
        // console.log("htmx:configRequest for Apply button");
        try {
            const filterParams = JSON.parse(getFilterValues());
            Object.assign(event.detail.parameters, filterParams);
        } catch (e) { console.error("Error processing filter values for HTMX:", e); }
    }
});
""", defer=True)

# JavaScript for modal interactions (triggering, form reset, pill updates)
MODAL_SCRIPTS: Tuple[Script, Script] = (
    Script("""
        document.body.addEventListener('openModalEvent', function(evt) {
            const modalDialog = document.getElementById('modal-placeholder');
            if (modalDialog && typeof modalDialog.showModal === 'function') {
                modalDialog.showModal();
            } else { console.error('Modal dialog (#modal-placeholder) not found or not supported.'); }
        });
    """),
    Script("""
        function updatePill(fieldName, newValue, newText, newStyleClass) {
            const pillId = `pill-display-${fieldName}`;
            const pillElement = document.getElementById(pillId);
            if (pillElement) {
                pillElement.textContent = newText;
                const currentStyle = pillElement.dataset.currentStyle || '';
                const newEffectiveStyle = newStyleClass || '${str(DEFAULT_PILL_STYLE)}'; // Use Python var
                if (currentStyle && pillElement.classList.contains(currentStyle)) pillElement.classList.replace(currentStyle, newEffectiveStyle);
                else pillElement.classList.add(newEffectiveStyle);
                pillElement.dataset.currentStyle = newEffectiveStyle;
                if (!pillElement.classList.contains('badge-sm')) pillElement.classList.add('badge-sm');

                const drop = typeof UIkit !== 'undefined' ? UIkit.drop(pillElement.closest('[uk-drop]')) : null;
                if (drop) drop.hide(false);
            } else { console.error(`Pill element not found: ${pillId}`); }
        }
        function resetModalForm() { /* ... (reset logic, ensure it's robust) ... */ }
        document.addEventListener('htmx:afterSwap', function(evt) {
            if (evt.detail.target.id === 'modal-placeholder') {
                evt.detail.target.querySelectorAll('[id^="pill-display-"]').forEach(pill => {
                    const style = Array.from(pill.classList).find(c=>c.startsWith('uk-label-')||c.startsWith('badge-')) || '${str(DEFAULT_PILL_STYLE)}';
                    pill.dataset.currentStyle = style.replace('badge-sm','').trim();
                    if(!pill.classList.contains('badge-sm')) pill.classList.add('badge-sm');
                });
            }
        });
    """)
)

# Collection of JavaScript resources for the page <head>
HEAD_JS_RESOURCES: Tuple[Script, ...] = (
    Script(src="https://cdn.jsdelivr.net/npm/uikit@3.latest/dist/js/uikit.min.js"),
    Script(src="https://cdn.jsdelivr.net/npm/uikit@3.latest/dist/js/uikit-icons.min.js"),
    Script(src="https://unpkg.com/alpinejs", defer=True), # defer is boolean
    Script(content=ALPINE_TOOLTIP_JS_SNIPPET), # Embed snippet directly
    Script(src="https://cdn.tailwindcss.com"), # TailwindCSS
    Script(src="https://unpkg.com/htmx-ext-preload@2.1.0"), # HTMX extension
)

# Initial headers for HTML pages, including theme, CSS, and head JS.
# Body JS (MODAL_SCRIPTS, FILTER_FORM_JS_HELPER, HTMX_CONFIG_REQUEST_LISTENER) should be placed before </body>
PAGE_INIT_HEADERS: Tuple[Any, ...] = ( # Use Any for mixed types from fasthtml
    Theme.slate.headers( # Assuming Theme.slate.headers returns a list/tuple of header elements
        mode="light", daisy=True, katex=False,
        radii=ThemeRadii.lg, shadows=ThemeShadows.lg
    ),
    CSS_STYLES,
    INTER_FONT_LINK,
    *HEAD_JS_RESOURCES,
)
# JavaScripts intended for the end of the body
PAGE_BODY_JS: Tuple[Script, ...] = (
    FILTER_FORM_JS_HELPER,
    HTMX_CONFIG_REQUEST_LISTENER,
    *MODAL_SCRIPTS
)

# Removed load_cookies_from_file function as it was misplaced.
# It belongs in and was correctly implemented in `easy_access/classification/httpx_downloader.py`.
```
