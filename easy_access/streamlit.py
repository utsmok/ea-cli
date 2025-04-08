import asyncio
import hashlib
from datetime import datetime

import pandas as pd
import polars as pl
import streamlit as st
from db.retrieve import retrieve_copyright_items
from db.update import update_copyright_items
from rich import print
from settings import SETTINGS

if False:
    # Auth logic -- temp disabled

    def authenticate_user(username: str, password: str) -> bool:
        userdict = {
            "admin": "test",
            "eemcs_user": "test",
        }
        print(f"Authenticating: {username} with password: {password}")
        if username in userdict and userdict[username] == password:
            return True
        return False

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    def show_login_form():
        st.title("Login")
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Login")
            if submitted:
                if authenticate_user(username, password):
                    st.session_state.authenticated = True
                    st.rerun()  # Rerun to hide login and show app
                else:
                    st.error("Invalid username or password")

    if not st.session_state.authenticated:
        show_login_form()
    else:
        # run main app logic
        ...


# Constants

FACULTIES: list[str] = list(SETTINGS.university_settings.faculty_abbreviations)
if "UNM" in FACULTIES:
    FACULTIES.remove("UNM")  # Remove UNM from the list of faculties

WORKFLOW_STATUS_CHOICES = ["ToDo", "InProgress", "Done"]
MANUAL_CLASSIFICATION_CHOICES = [
    "open access",
    "eigen materiaal - powerpoint",
    "eigen materiaal - overig",
    "lange overname",
    "eigen materiaal - titelindicatie",
    "anders",
    "korte overname",
    "middellange overname",
    "-",
]
COLUMNS_TO_DISPLAY = [
    "material_id",
    "url",
    "filename",
    "title",
    "owner",
    "ml_prediction",
    "author",
    "publisher",
    "doi",
    "isbn",
    "workflow_status",
    "manual_classification",
    "remarks",
    "faculty",
]

# async task functions

tasks: dict[str, asyncio.Task] = {}

try:
    loop = asyncio.get_running_loop()
except RuntimeError:
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)


def schedule_task(key, coro) -> None:
    """Schedules an async task and stores it with a unique key."""
    if key not in tasks:
        tasks[key] = loop.create_task(coro)


def process_tasks() -> None:
    """Process pending tasks on the event loop."""
    pending = [task for task in tasks.values() if not task.done()]
    if pending:
        loop.run_until_complete(asyncio.gather(*pending))


def generate_task_key(*args) -> str:
    """Generate a unique hash-based key for a task."""
    return hashlib.sha256("-".join(map(str, args)).encode()).hexdigest()


def save_user_edits(material_id: int, updates: dict | pl.DataFrame) -> bool:
    """
    Stores edits in the database.
    """
    if not isinstance(updates, pl.DataFrame):
        updates["material_id"] = material_id
        polars_update = pl.from_dict(updates)
    else:
        polars_update = updates
    try:
        task_key = generate_task_key(datetime.now(), material_id)
        schedule_task(
            task_key, update_copyright_items(polars_update, update_relations=False)
        )
        process_tasks()
        task_keys = list(tasks.keys())
        for task_key in task_keys:
            if tasks[task_key].done():
                del tasks[task_key]
        return True
    except Exception as e:
        print(f"Error during update: {e}")
        return False


def all_items() -> pl.DataFrame:
    """
    wrapper to retrieve all items from the database.
    can be used for caching for example
    """
    return retrieve_copyright_items()


def get_materials(
    faculty: str,
) -> tuple[pl.DataFrame, int]:
    """
    fetches materials for a given faculty
    can be expanded to include filters, sorting, limits, etc.
    """
    df = all_items().select(COLUMNS_TO_DISPLAY).filter(pl.col("faculty") == faculty)
    total_count = df.shape[0]
    return df, total_count


# main streamlit app


# init session vars for streamlit
st.set_page_config(page_title="Easy Access @UT", layout="wide")

if "selected_material_id" not in st.session_state:
    st.session_state.selected_material_id = None
if "edit_form_key" not in st.session_state:
    st.session_state.edit_form_key = 0

for faculty in FACULTIES:
    prefix = f"{faculty}_"
    if f"{prefix}page" not in st.session_state:
        st.session_state[f"{prefix}page"] = 1
    if f"{prefix}filters" not in st.session_state:
        st.session_state[f"{prefix}filters"] = {
            "workflow_status": [],
            "manual_classification": [],
        }
    if f"{prefix}sort_by" not in st.session_state:
        st.session_state[f"{prefix}sort_by"] = "material_id"
    if f"{prefix}sort_ascending" not in st.session_state:
        st.session_state[f"{prefix}sort_ascending"] = True
    if "selected_item" not in st.session_state:
        st.session_state["selected_item"] = {}


full_data: dict[str, pd.DataFrame] = {}


st.title("Material Classification Workflow")

st.header("Edit Material")

with st.form(key=f"edit_form_{st.session_state.edit_form_key}"):
    wf_status = None
    material_id = None
    manual_class = None

    if st.session_state["selected_item"]:
        item = st.session_state["selected_item"]
        print(item)
        material_id = item.get("material_id")
        wf_status = item.get("workflow_status")
        manual_class = item.get("manual_classification")

        material_id = st.write(
            "Selected material_id:",
            material_id,
        )
    else:
        st.info("Select a row in the table below to edit.")
    print(wf_status, manual_class)
    print(WORKFLOW_STATUS_CHOICES.index(wf_status) if wf_status else 0)
    print(MANUAL_CLASSIFICATION_CHOICES.index(manual_class) if manual_class else 0)
    wf_status = st.selectbox(
        "Workflow Status",
        options=WORKFLOW_STATUS_CHOICES,
        index=WORKFLOW_STATUS_CHOICES.index(wf_status) if wf_status else 0,
    )
    manual_class = st.selectbox(
        "Manual Classification",
        options=MANUAL_CLASSIFICATION_CHOICES,
        index=MANUAL_CLASSIFICATION_CHOICES.index(manual_class) if manual_class else 0,
    )
    remarks = st.text_area("Remarks")

    save_button = st.form_submit_button("Save Changes")

    if save_button:
        updates = {
            "workflow_status": wf_status,
            "manual_classification": manual_class,
            "remarks": remarks,
        }
        try:
            success = save_user_edits(st.session_state["selected_material_id"], updates)
            if success:
                st.success(
                    f"Material {st.session_state['selected_material_id']} updated successfully!"
                )
                st.rerun()
            else:
                st.error(
                    f"Failed to update material {st.session_state['selected_material_id']}."
                )
        except Exception as e:
            st.error(f"An error occurred during save: {e}")

faculty_tabs = st.tabs(FACULTIES)

for i, tab in enumerate(faculty_tabs):
    faculty = FACULTIES[i]
    prefix = f"{faculty}_"

    with tab:
        st.subheader(f"Materials for {faculty}")
        try:
            df, total_items = get_materials(
                faculty=faculty,
            )

        except Exception as e:
            st.error(f"Error fetching data for {faculty}: {e}")
            df = pl.DataFrame()  # Empty dataframe on error
            total_items = 0

        full_data[faculty] = st.dataframe(
            df,
            use_container_width=True,
            key=f"table_{faculty}",
            on_select="rerun",
            hide_index=True,
            selection_mode="single-row",
        )

        selection = full_data[faculty].get("selection").get("rows")
        if selection:
            st.session_state["selected_item"] = df[selection].to_dicts()[0]
            st.rerun()
        else:
            st.session_state["selected_item"] = {}
        if False:
            # functionality for storing changes when using st.data_editor
            edited_rows = st.session_state[f"table_{faculty}"].get("edited_rows")
            if edited_rows:
                st.write(
                    f"Storing changes for {len(edited_rows)} item{'s' if len(edited_rows) > 1 else ''}."
                )
                # edited_rows = {int: dict} where int is the index of the row in the original dataframe

                # grab the full rows from the original dataframe using the index
                for row in edited_rows:
                    save_user_edits(
                        full_data[faculty][row]["material_id"],
                        full_data[faculty][row],
                    )
