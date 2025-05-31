"""
This module provides functionalities for reading, creating, and manipulating
Excel sheets, particularly for copyright data management. It includes functions
for reading raw copyright exports, finalizing sheets with data entry capabilities,
storing data, and creating specific export formats.
"""

import json
import logging
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any # For type hints

import openpyxl
import openpyxl.worksheet
import openpyxl.worksheet.datavalidation
import openpyxl.worksheet.worksheet
import polars as pl
import typer # Used for typer.Exit
from openpyxl.styles import Alignment, NamedStyle
from openpyxl.worksheet.table import Table as ExcelTable # Alias to avoid confusion
from openpyxl.worksheet.table import TableStyleInfo

from easy_access.db.retrieve import (
    retrieve_copyright_items, # Used if data not passed to create_export_sheet
    retrieve_llm_classifications,
)
from easy_access.settings import DEPARTMENT_MAPPING, SETTINGS, ColInfo, DirSetting
from easy_access.utils import Directory, File

logger = logging.getLogger(__name__)


def read_copyright_export(file: File | None = None) -> tuple[str, pl.DataFrame]:
    """
    Reads data from the latest copyright export Excel file or a specified file.

    The input Excel sheet should be a direct export from the Copyright tool.
    This function standardizes column names, adds metadata columns (e.g.,
    retrieval date, default workflow status), and performs initial filtering
    (e.g., removing items with null material_id, specific filetypes).

    Args:
        file (File | None, optional): A `File` object pointing to a specific Excel file.
            If None, the function searches for the newest .xlsx file in the
            `RAW_COPYRIGHT_DATA` directory specified in settings. Defaults to None.

    Returns:
        tuple[str, pl.DataFrame]: A tuple containing:
            - str: The date of the processed file (YYYY-MM-DD format).
            - pl.DataFrame: A Polars DataFrame with the processed copyright data.
                            Returns an empty DataFrame if no file is found or on error.

    Raises:
        typer.Exit: If no file is found, or there are permission/value errors during processing.
    """
    processed_df: pl.DataFrame = pl.DataFrame()
    file_date_str: str = datetime.now().strftime("%Y-%m-%d") # Default date

    try:
        if not file:
            raw_data_dir = SETTINGS.dirs.get(DirSetting.RAW_COPYRIGHT_DATA)
            if not raw_data_dir or not raw_data_dir.exists:
                logger.warning(f"Raw copyright data directory not found or configured. Searched at: {raw_data_dir.full if raw_data_dir else 'N/A'}")
                raise typer.Exit(code=1)

            excel_files = [f for f in raw_data_dir.files if f.extension in ['.xlsx', '.xls'] and not f.name.startswith("~$")]
            if not excel_files:
                logger.warning(f"No Excel files found in {raw_data_dir.full}")
                raise typer.Exit(code=1)
            file = max(excel_files, key=lambda x: x.created)

        logger.info(f"Reading copyright data from: {file.name}")
        file_date_str = file.created.strftime("%Y-%m-%d")
        raw_df = pl.read_excel(file.path)

        # Standardize column names and add initial fields
        processed_df = (
            raw_df.with_columns(pl.all().cast(pl.Utf8, strict=False)) # Cast all to string first for safety
            .rename(lambda c: str(c).replace(" ", "_").replace("#", "count_").replace("*", "x").lower())
            .with_columns(
                pl.lit(file_date_str).alias("retrieved_from_copyright_on"),
                pl.lit("ToDo").alias("workflow_status"), # Default workflow status
                pl.col("last_change").str.replace(r"^- ভারতবর্ষ$", None) # Specific cleaning for placeholder
                    .str.strip_chars().str.strptime(pl.Date, "%Y-%m-%d", strict=False).dt.strftime("%Y-%m-%d"),
                pl.col("classification").str.to_lowercase(),
                faculty=pl.col("department").replace_strict(DEPARTMENT_MAPPING, default="Unmapped")
            )
        )
        logger.info(f"Retrieved {len(processed_df)} items initially from {file.name}.")

        # Filter out rows with null material_id or unwanted filetypes
        # Ensure material_id is treated as string for is_not_null checks if it might be mixed
        processed_df = processed_df.filter(pl.col("material_id").is_not_null() & (pl.col("material_id") != ""))

        # Define relevant filetypes, ensure 'filetype' column exists
        if "filetype" in processed_df.columns:
            relevant_filetypes = ["pdf", "ppt", "pptx", "doc", "docx", "-"] # Added pptx, docx
            processed_df = processed_df.filter(
                (pl.col("filetype").str.to_lowercase().is_in(relevant_filetypes)) |
                (pl.col("filetype").is_null()) |
                (pl.col("filetype") == "")
            )
        logger.info(f"{len(processed_df)} items remaining after initial filtering from {file.name}.")
        return file_date_str, processed_df

    except FileNotFoundError: # Should be caught by initial checks, but as safeguard
        logger.error(f"Critical: Copyright data file/directory not found. Path: {SETTINGS.dirs.get(DirSetting.RAW_COPYRIGHT_DATA, 'N/A')}")
        raise typer.Exit(code=1)
    except PermissionError:
        logger.error(f"Permission denied reading file: {file.name if file else 'N/A'}")
        raise typer.Exit(code=1)
    except ValueError as e: # e.g., max() on empty sequence if no files found and not handled
        logger.error(f"ValueError during copyright export processing (e.g. no files found): {e}")
        raise typer.Exit(code=1)
    except Exception as e: # Catch any other Polars or unexpected errors
        logger.error(f"Unexpected error in read_copyright_export for file {file.name if file else 'N/A'}: {e}")
        logger.debug(traceback.format_exc()) # For more detailed error in logs
        raise typer.Exit(code=1) # Exit on unexpected errors


@dataclass
class DataEntrySheet:
    """
    Manages the creation and formatting of a data entry sheet within an Excel workbook.
    This class handles adding data from a Polars DataFrame, styling the sheet as a table,
    and applying data validation rules.

    Attributes:
        sheet_name (str): The name of the sheet to be created.
        cols (list[ColInfo]): A list of `ColInfo` objects describing each column in the sheet.
        table_style (TableStyleInfo): Openpyxl table style to apply.
        workbook (openpyxl.Workbook): The parent Excel workbook object.
        file_path (str): The file path where the workbook will be saved. Used for saving.
        sheet (openpyxl.worksheet.worksheet.Worksheet): The created worksheet object (init=False).
        max_row (int): Maximum number of data rows added to the sheet (init=False).
        word_wrap_style (Alignment): Openpyxl Alignment style for word wrapping.
    """
    sheet_name: str
    cols: list[ColInfo]
    table_style: TableStyleInfo
    workbook: openpyxl.Workbook
    file_path: str # Path to save the workbook
    sheet: openpyxl.worksheet.worksheet.Worksheet = field(init=False)
    max_row: int = field(default=0, init=False)
    word_wrap_style: Alignment = field(default_factory=lambda: NamedStyle(name="wordwrap_dynamic", alignment=Alignment(wrapText=True)))


    def __post_init__(self) -> None:
        """Creates the sheet within the workbook after initialization."""
        self.sheet = self.workbook.create_sheet(self.sheet_name, index=1) # Create after 'Complete Data'

    def add_data(self, data: pl.DataFrame) -> None:
        """
        Populates the sheet with data from a Polars DataFrame and applies column formatting.

        Args:
            data (pl.DataFrame): The DataFrame containing data to add to the sheet.
        """
        if data.is_empty():
            logger.warning(f"Data provided to DataEntrySheet '{self.sheet_name}' is empty.")
            # Still create headers
            for col_idx, col_info in enumerate(self.cols):
                self.sheet.cell(1, col_idx + 1).value = col_info.new_name or col_info.name
            self.max_row = 0
        else:
            self.max_row = data.height # Polars DataFrame height for row count

        col_offset: int = 1 # Start writing from column 1 (A)
        for col_info in self.cols:
            col_name_in_sheet = col_info.new_name or col_info.name
            self.sheet.cell(1, col_offset).value = col_name_in_sheet # Write header

            column_data_list: list[Any]
            if col_info.is_new:
                column_data_list = [col_info.default_val] * self.max_row
            elif col_info.name in data.columns:
                # Apply default value to empty/null cells from existing column
                series_data = data.get_column(col_info.name)
                if col_info.default_val != "":
                    column_data_list = series_data.apply(lambda x: col_info.default_val if (x is None or str(x).strip() == "") else x).to_list()
                else:
                    column_data_list = series_data.to_list()
            else: # Column not in data and not new, fill with empty or log warning
                logger.warning(f"Column '{col_info.name}' not found in DataFrame for sheet '{self.sheet_name}'. Filling with empty strings.")
                column_data_list = [""] * self.max_row

            # Write data cells
            for row_idx, cell_val in enumerate(column_data_list, start=2): # Data starts at row 2
                current_cell = self.sheet.cell(row_idx, col_offset)
                current_cell.value = cell_val

                if col_info.is_url and isinstance(cell_val, str) and "/" in cell_val:
                    current_cell.value = ".../" + cell_val.split("/")[-1] # Display shortened URL
                    current_cell.hyperlink = cell_val

                # Track width for auto-sizing (simplified)
                # Actual width calculation in openpyxl is more complex. This is a heuristic.
                cell_display_value = str(current_cell.value if current_cell.value is not None else "")
                if len(cell_display_value) > col_info.max_width:
                    col_info.max_width = len(cell_display_value)
                if len(cell_display_value) > 40:
                    col_info.count_max_width_over_40 += 1

            col_offset += 1

        # Apply column-level formatting (dropdowns, width)
        self._format_columns()
        self.create_table() # Add Excel table structure
        self.save() # Save workbook

    def _format_columns(self) -> None:
        """Applies data validation (dropdowns) and column width adjustments."""
        for col_idx, col_info in enumerate(self.cols, start=1):
            col_letter = openpyxl.utils.get_column_letter(col_idx)

            if col_info.has_dropdown and self.max_row > 0: # Only add validation if there's data
                dv = openpyxl.worksheet.datavalidation.DataValidation(
                    type="list", formula1=f'"{col_info.dropdown_options}"', allowBlank=True
                )
                dv.error = "Please select a valid option from the list."
                dv.errorTitle = "Invalid Option"
                dv.prompt = "Please select from the list."
                dv.promptTitle = "List Selection"
                self.sheet.add_data_validation(dv)
                dv.add(f"{col_letter}2:{col_letter}{self.max_row + 1}") # Apply to data rows

            # Column width adjustment heuristic
            dim = self.sheet.column_dimensions[col_letter]
            if col_info.max_width > 40 and (col_info.count_max_width_over_40 > 5 or col_info.count_max_width_over_40 > self.max_row * 0.2):
                dim.width = 40
                # Apply word wrap to cells in this column if many are long
                if self.max_row > 0:
                    for row_num in range(2, self.max_row + 2):
                        self.sheet.cell(row_num, col_idx).alignment = self.word_wrap_style.alignment # Apply alignment part
            else:
                dim.width = max(col_info.max_width + 2, len(col_info.new_name or col_info.name) + 2, 10) # Min width 10
            dim.bestFit = True # Let openpyxl try to best fit after setting width

    def create_table(self) -> None:
        """Formats the populated data range as an Excel table."""
        if self.max_row == 0 and not self.cols: # No data and no columns
            return

        max_col_letter = openpyxl.utils.get_column_letter(len(self.cols) if self.cols else 1)
        # Table ref should be at least A1 if no cols/rows, or cover headers if only headers
        table_ref_end_row = self.max_row + 1 if self.max_row > 0 else 1

        table = ExcelTable(
            displayName=self.sheet_name.replace(" ", "_").replace("-", "_")[:30], # Sanitize name
            ref=f"A1:{max_col_letter}{table_ref_end_row}",
        )
        table.tableStyleInfo = self.table_style
        self.sheet.add_table(table)

    def save(self) -> None:
        """Saves the parent workbook to the specified file path."""
        try:
            self.workbook.save(filename=self.file_path)
        except Exception as e:
            logger.error(f"Error saving workbook to {self.file_path}: {e}")


def finalize_sheet(file: File, data: pl.DataFrame, style_iter: int) -> int:
    """
    Adds a standardized 'Data Entry' sheet to an existing Excel workbook.
    The workbook is expected to already contain a 'Complete Data' sheet.
    This function also handles adding LLM classification data to a separate sheet
    if the processed file is an overview sheet.

    Args:
        file (File): The `File` object representing the Excel workbook to modify.
        data (pl.DataFrame): DataFrame to populate the 'Data Entry' sheet with.
                             This data is typically a subset or processed version of 'Complete Data'.
        style_iter (int): An integer used to cycle through table styles for visual distinction.

    Returns:
        int: The updated `style_iter` value (incremented).
    """
    try:
        wb = openpyxl.load_workbook(filename=str(file.path))
    except FileNotFoundError:
        logger.error(f"File not found for finalizing sheet: {file.path}")
        return style_iter # Return original style_iter on error

    # Ensure 'Complete Data' sheet exists or rename active if it's the only one
    complete_data_sheet_name = SETTINGS.data_settings.complete_data_name
    if complete_data_sheet_name not in wb.sheetnames:
        if len(wb.sheetnames) == 1: # If only one sheet, assume it's the complete data
            wb.active.title = complete_data_sheet_name
        else: # If multiple sheets and 'Complete Data' is missing, log warning
            logger.warning(f"Sheet '{complete_data_sheet_name}' not found in {file.name}. Active sheet titled '{wb.active.title}'.")
            # Optionally, one could try to find a suitable sheet or create it.

    # Define table style for the new Data Entry sheet
    # Cycle through available styles (e.g., TableStyleMedium1 to TableStyleMedium21)
    table_style_number = (style_iter % 21) + 1 # Keep it within common range of styles
    tabstyle = TableStyleInfo(
        name=f"TableStyleMedium{table_style_number}", showRowStripes=True,
    )

    data_entry_sheet = DataEntrySheet(
        workbook=wb,
        sheet_name=SETTINGS.data_settings.data_entry_name,
        cols=SETTINGS.data_settings.data_entry_cols, # Assumes ColInfo objects are correctly defined in settings
        table_style=tabstyle,
        file_path=str(file.path),
    )

    # Ensure data for data entry sheet is unique by material_id
    data_for_entry_sheet = data.unique(subset=["material_id"], keep="first", maintain_order=True)
    data_entry_sheet.add_data(data_for_entry_sheet) # This now also calls .save()
    logger.info(f"Added/updated 'Data Entry' sheet in {file.name}")

    # Handle LLM classification data for overview sheets
    if "overview" in file.path.stem.lower(): # Check if "overview" is in the filename stem
        logger.debug(f"Processing LLM classification for overview file: {file.name}")
        llm_data = enrich_with_llm_classifications(data_for_entry_sheet) # Uses data from data entry
        if not llm_data.is_empty():
            llm_sheet_name = "LLM_Classification_Data" # Standardized name
            llm_table_name = "LLMClassificationTable"
            # Ensure llm_sheet_path is correctly formed
            llm_sheet_path = file.path.parent / f"{file.path.stem}_llm_data.xlsx" # Storing in a separate file now

            try:
                llm_data.write_excel(
                    workbook=str(llm_sheet_path), # write_excel expects string path for new workbook
                    worksheet=llm_sheet_name,
                    table_name=llm_table_name,
                    table_style=f"TableStyleMedium{(table_style_number + 1) % 21 + 1}", # Use next style
                    autofit=True,
                )
                logger.info(f"Stored LLM classification data to separate sheet: {llm_sheet_path.name}")
            except Exception as e:
                 logger.error(f"Failed to write LLM classification sheet for {file.name}: {e}")
        else:
            logger.info(f"No LLM classification data to enrich for {file.name}.")

    return style_iter + 1 # Increment style iterator for next sheet


def store_complete_data(output_file_path: Path | File, data: pl.DataFrame) -> None:
    """
    Stores the provided DataFrame to an Excel file, typically as the 'Complete Data' sheet.
    It uses a predefined column order from settings and ensures `material_id` is unique.

    Args:
        output_file_path (Path | File): The path (or File object) where the Excel file will be saved.
        data (pl.DataFrame): The DataFrame to store.
    """
    path_obj: Path = output_file_path.path if isinstance(output_file_path, File) else output_file_path

    # Ensure parent directory exists
    path_obj.parent.mkdir(parents=True, exist_ok=True)

    if path_obj.exists() and path_obj.stat().st_size > 0:
        logger.info(f"File {path_obj} already exists and is not empty. It will be overwritten.")
        # No explicit delete needed as write_excel typically overwrites.
        # However, if issues occur, uncomment: File(path_obj).delete()

    # Select and order columns as per settings
    cols_to_select = [col for col in SETTINGS.data_settings.final_data_col_order if col in data.columns]
    if not cols_to_select: # Fallback to all columns if none of the specified are present
        logger.warning("No columns from final_data_col_order found in data. Writing all columns.")
        cols_to_select = data.columns

    processed_data = data.select(cols_to_select).unique(subset=["material_id"], keep="first", maintain_order=True)

    try:
        processed_data.write_excel(
            workbook=str(path_obj), # Ensure path is string for write_excel
            worksheet=SETTINGS.data_settings.complete_data_name,
            table_style="TableStyleMedium1", # Default style for complete data
            autofit=True
        )
        logger.info(f"Stored {processed_data.height} rows to {path_obj}")
    except Exception as e:
        logger.error(f"Failed to store complete data to {path_obj}: {e}")


def read_export_sheets() -> pl.DataFrame:
    """
    Reads and concatenates all Excel and CSV files from the `EXPORT_TO_SURF` directory.

    Returns:
        pl.DataFrame: A single DataFrame containing data from all export sheets.
                      Returns an empty DataFrame if no files are found or on error.
    """
    all_data: list[pl.DataFrame] = []
    export_dir = SETTINGS.dirs.get(DirSetting.EXPORT_TO_SURF)

    if not export_dir or not export_dir.exists:
        logger.warning(f"Export directory not found or configured: {export_dir.full if export_dir else 'N/A'}")
        return pl.DataFrame()

    for file_obj in export_dir.files: # Assumes .files gives File objects
        try:
            if file_obj.extension in [".xlsx", ".xls"]:
                all_data.append(pl.read_excel(file_obj.path))
            elif file_obj.extension == ".csv":
                all_data.append(pl.read_csv(file_obj.path))
        except Exception as e:
            logger.warning(f"Could not read or process export file {file_obj.name}: {e}")

    if not all_data:
        logger.info("No data found in export sheets.")
        return pl.DataFrame()

    return pl.concat(all_data, how="diagonal_relaxed")


def create_export_sheet(
    data: pl.DataFrame, # Made non-optional as per previous refactor
    faculty_name: str | None = None,
    # print_overview: bool = True, # Parameter seems unused
) -> None:
    """
    Creates an export Excel sheet suitable for re-import into the Copyright tool.
    The sheet contains specific columns and is named with faculty and date.
    A second sheet with full details is also created.

    Args:
        data (pl.DataFrame): The DataFrame containing items to be exported. Must not be empty.
        faculty_name (str | None, optional): Name of the faculty for naming the export file
                                            and creating a subdirectory. If None or "ALL_FACULTIES",
                                            a global export is assumed.
    """
    # These columns are specific to the target import format.
    EXPORT_COL_NAMES: list[str] = [
        "material_id", "filename", "manual_classification",
        "auditor", "remarks", "scope",
    ]
    TODAY_STR: str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S") # More descriptive name

    if data.is_empty(): # data is now required
        logger.warning("No data provided to create_export_sheet. Skipping export.")
        return

    # Filter for "Done" items if this is a strict requirement for exports
    export_data_subset = data.filter(pl.col("workflow_status") == "Done")
    if export_data_subset.is_empty():
        logger.warning(f"No items with workflow_status 'Done' found for '{faculty_name or 'ALL'}'. Skipping export.")
        return

    base_export_dir_path: Path = SETTINGS.dirs[DirSetting.EXPORT_TO_SURF].full
    current_export_dir_path: Path
    if faculty_name and faculty_name != "ALL_FACULTIES":
        current_export_dir_path = base_export_dir_path / faculty_name
    else: # Global export
        current_export_dir_path = base_export_dir_path

    Directory(current_export_dir_path).mkdir(parents=True, exist_ok=True) # Ensure export dir exists

    file_prefix_str = f"utwente_{faculty_name or 'GLOBAL'}_{TODAY_STR}_{export_data_subset.height}_items"

    main_export_file_path = current_export_dir_path / f"{file_prefix_str}_copyright_import.xlsx"
    details_export_file_path = current_export_dir_path / f"{file_prefix_str}_copyright_import_full_details.xlsx"

    logger.info(f"Exporting {export_data_subset.height} items for '{faculty_name or 'GLOBAL'}' to {current_export_dir_path}")

    # Create the main, minimal export sheet
    try:
        # Ensure only existing columns from EXPORT_COL_NAMES are selected
        cols_for_main_export = [col for col in EXPORT_COL_NAMES if col in export_data_subset.columns]
        export_data_subset.select(cols_for_main_export).write_excel(str(main_export_file_path))
        logger.info(f"Main export sheet saved to: {main_export_file_path.name}")
    except Exception as e:
        logger.error(f"Failed to write main export sheet {main_export_file_path.name}: {e}")


    # Store the sheet with full details using store_complete_data
    try:
        store_complete_data(file=details_export_file_path, data=export_data_subset) # store_complete_data handles its own logging
        logger.info(f"Full details export sheet saved to: {details_export_file_path.name}")
    except Exception as e:
        logger.error(f"Failed to write full details export sheet {details_export_file_path.name}: {e}")


def enrich_with_llm_classifications(data: pl.DataFrame) -> pl.DataFrame:
    """
    Enriches a DataFrame with LLM classification data. (Currently Disabled)
    If enabled, it would retrieve LLM classifications for material IDs in the input
    DataFrame and join them.

    Args:
        data (pl.DataFrame): The input DataFrame of copyright items.

    Returns:
        pl.DataFrame: The input DataFrame, potentially enriched with LLM data columns.
                      Returns an empty DataFrame or the original if enrichment is disabled or fails.
    """
    logger.warning("Enriching data with LLM classification data is currently disabled.")
    return pl.DataFrame() # Return empty or original `data` if preferred when disabled.

    # Unreachable code:
    # llm_data = retrieve_llm_classifications(
    #     selected_material_ids=data.select("material_id").unique().to_series().to_list()
    # )
    # if llm_data is None or llm_data.is_empty(): # retrieve_llm_classifications now returns empty DF with schema
    #     logger.info("No LLM classification data found for enrichment.")
    #     return data # Return original data if no LLM data
    #
    # joined_data = data.join(llm_data, on="material_id", how="left")
    #
    # # Define expected column order for the enriched data (subset or all)
    # # This list can be extensive and should ideally match a defined output schema.
    # llm_enriched_col_order = [
    #     "material_id", "url", "filename", "manual_classification", "remarks", # Core cols
    #     "ml_prediction", "allowed_usage_llm", "allowed_usage_reasoning_llm", # LLM specific
    #     # ... other relevant columns from both dataframes
    # ]
    # final_cols = [col for col in llm_enriched_col_order if col in joined_data.columns]
    # missing_cols = [col for col in llm_enriched_col_order if col not in joined_data.columns]
    # if missing_cols:
    #     logger.debug(f"Missing expected columns after LLM enrichment: {missing_cols}")
    # return joined_data.select(final_cols)


def retrieve_all_classifications() -> pl.DataFrame:
    """
    Retrieves all LLM classification data from JSON files stored in the classifications directory.
    Handles ".replace" files which indicate a material ID was superseded by another.

    Returns:
        pl.DataFrame: A DataFrame containing all processed LLM classification data.
                      Returns an empty DataFrame if no classification files are found or on error.
    """
    classifications_dir = SETTINGS.dirs.get(DirSetting.CLASSIFICATIONS)
    if not classifications_dir or not classifications_dir.exists:
        logger.warning(f"Classifications directory not found: {classifications_dir.full if classifications_dir else 'N/A'}")
        return pl.DataFrame()

    all_files_in_dir: list[File] = classifications_dir.files_r # Get all files recursively

    all_json_files: list[File] = [
        f for f in all_files_in_dir
        if f.extension == ".json" and "_" not in f.name and f.name.rstrip(".json").isdigit()
    ]
    all_replacement_files_info: dict[int, int] = {} # old_id -> new_id
    for f in all_files_in_dir:
        if f.extension == ".replace":
            parts = f.name.replace(".replace", "").split("_")
            if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                all_replacement_files_info[int(parts[0])] = int(parts[1])
            else:
                logger.warning(f"Malformed .replace file found: {f.name}")

    logger.info(
        f"Found {len(all_json_files)} primary JSON classification files and "
        f"{len(all_replacement_files_info)} replacement directives."
    )

    classification_data_list: list[dict[str, Any]] = []
    processed_mat_ids: set[int] = set()

    for json_file in all_json_files:
        mat_id_str = json_file.name.replace(".json", "")
        try:
            current_mat_id = int(mat_id_str)
            if current_mat_id in processed_mat_ids: continue # Already processed (e.g. via a replacement)

            with open(json_file.path, encoding="utf-8", errors="replace") as f_json:
                single_classification_data = json.load(f_json)

            # Apply replacements: if this mat_id was replaced, skip its direct data.
            # If another mat_id was replaced by this one, this data will be used for that original ID.
            # This logic needs to be careful. The current structure seems to imply data is stored under the NEW ID.

            # Store data under its actual material_id from filename
            single_classification_data["material_id"] = current_mat_id
            # Add _llm suffix to all keys from JSON, except for material_id
            transformed_data: dict[str, Any] = {"material_id": current_mat_id}
            for key, value in single_classification_data.items():
                if key != "material_id":
                    transformed_data[key + "_llm"] = value

            classification_data_list.append(transformed_data)
            processed_mat_ids.add(current_mat_id)

        except json.JSONDecodeError:
            logger.warning(f"Could not decode JSON from: {json_file.name}")
        except ValueError: # For int conversion
            logger.warning(f"Invalid material_id from filename: {json_file.name}")
        except Exception as e:
            logger.error(f"Error processing classification file {json_file.name}: {e}")

    # Handle replacements: if an old_id was replaced by new_id, and we have data for new_id,
    # create an entry for old_id using new_id's data.
    # This assumes data in JSON files is stored under the *new* (replacement) ID.
    final_data_for_df: list[dict[str, Any]] = []
    temp_data_by_id: dict[int, dict[str, Any]] = {item["material_id"]: item for item in classification_data_list}

    for item_data in classification_data_list:
        final_data_for_df.append(item_data) # Add data for the ID it was found under

    for old_id, new_id in all_replacement_files_info.items():
        if old_id in temp_data_by_id: # Data for old_id exists, maybe it's also a replacement target
            logger.debug(f"Data for {old_id} already processed directly. Replacement by {new_id} might be redundant or a chain.")
            # If old_id's data should NOT be used because it's replaced, needs more complex logic.
            # Current assumption: JSON files are named with the ID of the data they contain.
            # If old_id.json exists, it's used for old_id. If old_id.replace points to new_id,
            # and new_id.json exists, we might create an *additional* entry for old_id using new_id's data if not careful.

        # If old_id's data is NOT directly present, but it's replaced by new_id for which we DO have data:
        elif old_id not in temp_data_by_id and new_id in temp_data_by_id:
            replacement_data_content = temp_data_by_id[new_id].copy()
            replacement_data_content["material_id"] = old_id # Attribute this data to the old_id
            final_data_for_df.append(replacement_data_content)
            logger.info(f"Applied replacement: Data for {new_id} is now also used for {old_id}.")
        elif new_id not in temp_data_by_id:
             logger.warning(f"Replacement for {old_id} by {new_id} specified, but no data found for {new_id}.")


    if not final_data_for_df:
        return pl.DataFrame()
    return pl.from_dicts(final_data_for_df, infer_schema_length=None)
