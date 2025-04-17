from fasthtml.common import Link, Script, Style
from fastlite import Database, database
from monsterui.core import Theme, ThemeRadii, ThemeShadows
from monsterui.franken import LabelT

db: Database = database("db.sqlite3")

# holds constants/enums/etc used in the dashboard and webui


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

# this dict defines the badge styles for various values in specific columns
# the keys are the column names, and the values are dicts with the
# colvalues as keys and the style to use as value
# most styles use LabelT class to style using FrankenUI,
# but some use strings to define classes (e.g. daisyUI badges).
BADGE_STYLES: dict[str, dict[str, str | LabelT]] = {
    "faculty": {
        "BMS": "badge-success",
        "EEMCS": "badge-info",
        "ET": "badge-warning",
        "ITC": "badge-error",
        "TNW": "badge-secondary",
    },
    "classification": {
        **{
            v: LabelT.primary
            for v in [
                "open access",
                "eigen materiaal - powerpoint",
                "eigen materiaal - titelindicatie",
                "eigen materiaal - overig",
                "eigen materiaal",
            ]
        },
        **{
            v: LabelT.secondary
            for v in [
                "onbekend",
                "niet geanalyseerd",
                "in onderzoek",
                "licentie beschikbaar",
                "verwijderverzoek verstuurd",
            ]
        },
        **{
            v: LabelT.destructive
            for v in [
                "korte overname",
                "middellange overname",
                "lange overname",
            ]
        },
        "None": LabelT.destructive,
    },
    "status": {
        "Published": LabelT.primary,
        "Unpublished": LabelT.secondary,
        "Deleted": LabelT.destructive,
    },
    "workflow_status": {
        "ToDo": LabelT.destructive,
        "InProgress": LabelT.secondary,
        "Done": LabelT.primary,
    },
}
# default styles for pills not defined in BADGE_STYLES or elsewhere
DEFAULT_PILL_STYLE = LabelT.secondary

PORT = 8000
ROOT_URL = f"http://localhost:{PORT}"
DEFAULT_PER_PAGE = 15
MAX_CELL_LENGTH = 35


# JS / CSS / headers / HTML


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
    Script("https://unpkg.com/htmx-ext-preload@2.1.0"),
)

BODY_JS = (MODAL_TRIGGER, MODAL_INTERACTION)

INIT_HEADERS = (
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
