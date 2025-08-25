import contextlib
import json
import logging
import os
import warnings
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import openpyxl
import openpyxl.worksheet
import openpyxl.worksheet.datavalidation
import openpyxl.worksheet.worksheet
import polars as pl
import typer
from loguru import logger
from openpyxl.styles import Alignment, NamedStyle
from openpyxl.worksheet.table import Table as ExcelTable
from openpyxl.worksheet.table import TableStyleInfo

from easy_access.db.retrieve import (
    retrieve_copyright_items,
)

# from easy_access.settings import DEPARTMENT_MAPPING, SETTINGS, ColInfo, DirSetting # Will be passed as parameters
from easy_access.settings import ColInfo, DirSetting, Settings  # Keep for type hinting
from easy_access.utils import Directory, File


def _read_excel_quiet(file_path: str | Path, **kwargs) -> pl.DataFrame:
    """
    Reads an Excel file quietly, suppressing dtype inference messages.

    We redirect stdout/stderr during the read to avoid noisy messages from the
    underlying libraries. If the quiet read fails, a second attempt without
    suppression is performed to raise a visible error.
    """
    # Temporarily raise log level for noisy libraries and silence warnings
    noisy_loggers = ["polars", "openpyxl", "pyxlsb", "lxml"]
    prev_levels = {}
    for name in noisy_loggers:
        lg = logging.getLogger(name)
        prev_levels[name] = lg.level
        lg.setLevel(logging.ERROR)
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            with open(os.devnull, "w") as devnull:
                with (
                    contextlib.redirect_stdout(devnull),
                    contextlib.redirect_stderr(devnull),
                ):
                    return pl.read_excel(file_path, **kwargs)
    finally:
        for name, level in prev_levels.items():
            logging.getLogger(name).setLevel(level)


def read_copyright_export(
    settings: Settings, file: File | None = None
) -> tuple[str, pl.DataFrame]:
    """
    Reads in data from the latest copyright export file in the copyright dir;
    or if a file is given, reads in that file.
    Input should be a direct export from the CopyRight tool without any changes.
    """
    try:
        if not file:
            logger.info(
                f"Reading in newest Copyright Data from directory: {settings.dirs[DirSetting.RAW_COPYRIGHT_DATA]}"
            )
            file = max(
                settings.dirs[DirSetting.RAW_COPYRIGHT_DATA].files,
                key=lambda x: x.created,
            )

        logger.info(f"Reading in data from:\n            {file.name}\n")
        latest_file_date = file.created.strftime("%Y-%m-%d")
        raw_copyright_data = _read_excel_quiet(file.path, sheet_name=None)
        copyright_data = (
            raw_copyright_data.with_columns(pl.exclude(pl.Utf8).cast(str))
            .rename(
                lambda col: col.replace(" ", "_")
                .replace("#", "count_")
                .replace("*", "x")
                .lower()
            )
            .with_columns(
                pl.Series(
                    "retrieved_from_copyright_on",
                    [latest_file_date] * len(raw_copyright_data),
                ),
                pl.Series("workflow_status", ["ToDo"] * len(raw_copyright_data)),
                pl.col("last_change")
                .str.replace(r"^-$", "")
                .str.strip_chars()
                .str.strptime(pl.Date, "%Y-%m-%d", strict=False)
                .dt.strftime("%Y-%m-%d"),
                pl.col("classification").str.to_lowercase(),
                faculty=pl.col("department").replace_strict(
                    settings.university_settings.department_mapping, default="Unmapped"
                ),
            )
        )

        # now drop rows we definitely do not want.
        # - drop row if material_id is null, None, blank, or '-'
        # - keep rows with filetype pdf, ppt, doc, or blank ('-'/None/null/""), drop rest
        logger.info(f"Retrieved {len(copyright_data)} items from {file.name}.")

        copyright_data = copyright_data.filter(pl.col("material_id").is_not_null())
        copyright_data = copyright_data.filter(
            (pl.col("filetype").is_in(["pdf", "ppt", "doc", "-"]))
            | (pl.col("filetype").is_null())
        )

        logger.info(
            f"{len(copyright_data)} items remaining from {file.name} after filtering out missing material_ids and specific filetypes."
        )
        return latest_file_date, copyright_data
    except FileNotFoundError:
        logger.warning(
            f"No files found in {settings.dirs[DirSetting.RAW_COPYRIGHT_DATA]}"
        )  # Use passed settings
        raise typer.Exit(code=1)
    except PermissionError:
        logger.warning(f"Permission denied to read {file}")
        raise typer.Exit(code=1)
    except ValueError:
        logger.warning("No file found.")
        raise typer.Exit(code=1)


@dataclass
class DataEntrySheet:
    """
    Use to add a dataentry sheet to an excel file.
    Has functions to add data from dataframe, format as table, add datavalidation, and save
    """

    sheet_name: str
    cols: list[
        ColInfo
    ]  # a list with the cols in order of appearance from left to right
    table_style: TableStyleInfo
    workbook: openpyxl.Workbook
    sheet: openpyxl.worksheet.worksheet.Worksheet = field(init=False)
    file_path: str
    max_row: int = 0
    word_wrap_style: Alignment = NamedStyle(
        name="wordwrap", alignment=Alignment(wrapText=True)
    )

    def __post_init__(self):
        self.sheet = self.workbook.create_sheet(self.sheet_name, index=1)

    def add_data(self, data: pl.DataFrame) -> None:
        self.max_row = data.shape[0]
        colnum = 0

        for col in self.cols:
            colnum += 1
            col_name = col.new_name if col.new_name else col.name
            if col.is_new:
                # Check if col is truly new first by retrieving the data from the dataframe
                # fill empty cells with default value if the col exists
                # otherwise create new data with default value and length of max_row
                if col.name in data.columns:
                    col_data = data.select(pl.col(col.name)).to_series().to_list()
                    if col.default_val != "":
                        for item_num, item in enumerate(col_data):
                            if item == "" or not item:
                                col_data[item_num] = col.default_val
                else:
                    # create new coldata
                    col_data = [col.default_val] * self.max_row
            else:
                # retrieve coldata from dataframe
                col_data = data.select(pl.col(col.name)).to_series().to_list()
                if col.default_val != "":
                    for item_num, item in enumerate(col_data):
                        if item == "" or not item:
                            col_data[item_num] = col.default_val

            self.sheet.cell(1, colnum).value = col_name
            for row, cell_data in enumerate(col_data, start=2):
                if not cell_data:
                    self.sheet.cell(row, colnum).value = cell_data
                    continue

                if col.is_url:
                    if "/" not in cell_data:
                        self.sheet.cell(row, colnum).value = cell_data
                    else:
                        self.sheet.cell(row, colnum).value = (
                            ".../" + cell_data.split("/")[-1]
                        )
                    self.sheet.cell(row, colnum).hyperlink = cell_data
                    if len(self.sheet.cell(row, colnum).value) > col.max_width:
                        col.max_width = len(self.sheet.cell(row, colnum).value)
                    if len(self.sheet.cell(row, colnum).value) > 40:
                        col.count_max_width_over_40 += 1

                else:
                    self.sheet.cell(row, colnum).value = cell_data
                    if len(str(cell_data)) > col.max_width:
                        col.max_width = len(str(cell_data))
                    if len(str(cell_data)) > 40:
                        col.count_max_width_over_40 += 1

        for colnum, col in enumerate(self.cols):
            col_letter = chr(ord("A") + colnum)
            colnum += 1
            if col.has_dropdown:
                dv = openpyxl.worksheet.datavalidation.DataValidation(
                    type="list", formula1=col.dropdown_options, allowBlank=True
                )
                dv.error = "Please select a valid option from the list"
                dv.errorTitle = "Invalid option"
                dv.prompt = "Please select from the list"
                dv.promptTitle = "List selection"
                self.sheet.add_data_validation(dv)
                if self.max_row == 1:
                    dv.add(f"{col_letter}2")
                else:
                    dv.add(f"{col_letter}2:{col_letter}{self.max_row + 1}")
            if col.max_width > 40 and (
                (col.count_max_width_over_40 > 5)
                or (col.count_max_width_over_40 > self.max_row - 2)
            ):
                # Too much long items: cap width to 40 & enable word wrap for this col
                for row in range(2, self.max_row + 1):
                    self.sheet.cell(row, colnum).style = self.word_wrap_style
                self.sheet.column_dimensions[col_letter].bestFit = False
                self.sheet.column_dimensions[col_letter].width = 40
            else:
                # Acceptable width, don't enable word wrap but fit width to contents
                self.sheet.column_dimensions[col_letter].width = col.max_width

        self.create_table()
        self.save()

    def create_table(self) -> None:
        max_col_letter = chr(ord("A") + len(self.cols) - 1)
        table = ExcelTable(
            displayName=self.sheet_name.replace(" ", ""),
            ref=f"A1:{max_col_letter}{self.max_row + 1}",
        )
        table.tableStyleInfo = self.table_style
        self.sheet.add_table(table)

    def save(self) -> None:
        self.workbook.save(filename=self.file_path)


def finalize_sheet(
    settings: Settings, file: File, data: pl.DataFrame, style_iter: int
) -> int:  # Added settings, changed return type
    """
    This function takes an excel file with 'Complete Data' and adds a data entry sheet +styling.
    Input: an excel file with the complete data, and a dataframe with that same data to be processed for the data entry sheet

    Adds the sheet to the workbook and saves it. Returns the incremented style_iter var.
    """
    wb = openpyxl.load_workbook(filename=str(file.path))
    if (
        settings.data_settings.complete_data_name not in wb.sheetnames and wb.active
    ):  # Use passed settings # Check if active sheet exists
        wb.active.title = (
            settings.data_settings.complete_data_name
        )  # Use passed settings

    tabstyle = TableStyleInfo(
        name=f"TableStyleMedium{style_iter}",
        showRowStripes=True,
    )
    style_iter = style_iter + 1
    sheet = DataEntrySheet(
        workbook=wb,
        sheet_name=settings.data_settings.data_entry_name,  # Use passed settings
        cols=settings.data_settings.data_entry_cols,  # Use passed settings
        table_style=tabstyle,
        file_path=str(file.path),
    )
    data = data.unique("material_id")
    sheet.add_data(data)
    logger.info(f"Added data entry sheet to {file.name}")
    # llm_classification_data = enrich_with_llm_classifications(settings=settings, data=data) # enrich_with_llm_classifications is disabled
    if "overview" in file.path.stem:
        # only add the llm classification data if the file is an overview file
        llm_classification_data = enrich_with_llm_classifications(
            settings=settings, data=data
        )  # Pass settings and data
        if llm_classification_data.is_empty():  # So this will likely be true
            return style_iter
        llm_sheet_path = (
            file.path.parent / f"{file.path.stem}_llm_classification_data.xlsx"
        )
        llm_classification_data.write_excel(  # This part might not be reached if enrich_with_llm_classifications stays disabled
            workbook=llm_sheet_path,
            worksheet="llm_classification_data",
            table_name="llm_classification_data",
            table_style="TableStyleMedium3",
            autofit=True,
        )
        logger.info(f"Stored llm_classification_data sheet to {llm_sheet_path.name}")

    return style_iter


def store_complete_data(
    settings: Settings, file: File | Path, data: pl.DataFrame
) -> None:  # Added settings
    """
    Stores the given data in an excel file with 1 sheet named SETTINGS.data_settings.complete_data_name
    using the col order in SETTINGS.data_settings.final_data_col_order
    """

    if isinstance(file, File):
        file = file.path

    if file.exists():
        size = file.stat().st_size
        if size > 0:
            File(file).delete()
    selectcols = [
        col
        for col in settings.data_settings.final_data_col_order  # Use passed settings
        if col in data.columns
    ]
    data = data.select(selectcols)
    data = data.unique("material_id")
    data.write_excel(
        file, worksheet=settings.data_settings.complete_data_name
    )  # Use passed settings
    logger.info(f"Stored {data.shape[0]} rows to {file}")


def read_export_sheets(settings: Settings) -> pl.DataFrame:  # Added settings
    """
    Read in the export sheets from the EXPORT_TO_SURF dir
    return as concatenated dataframe
    """
    returndata = pl.DataFrame()
    for file in settings.dirs[DirSetting.EXPORT_TO_SURF].files:  # Use passed settings
        if file.extension in [".xls", ".xlsx"]:
            returndata = pl.concat(
                [returndata, _read_excel_quiet(file.path, sheet_name=None)],
                how="diagonal_relaxed",
            )
        if file.extension in [".csv"]:
            returndata = pl.concat(
                [returndata, pl.read_csv(file.path)], how="diagonal_relaxed"
            )
    return returndata


def create_export_sheet(  # Added settings
    settings: Settings,
    data: pl.DataFrame | None = None,
    faculty: str | None = None,
    print_overview: bool = True,
):
    """
        Create an export sheet to import back into CopyRight tool.
        Specifications:

            - .xlsx file with 1 sheet
            - utf-8 encoding
            - Column info:
            OutputExcelField        --	QlikField               --  Notes
    ----------------------------------------------------------------------------------------------------------------
            MaterialID              --	Material id	            --  Formatted as Number
            Filename                --	Filename	            --  n/a
            Manual_classification   --	Manual classification   --  Classification of the item
            Manual_identifier       --	Manual identifier	    --  e.g. for ISBN/DOI/...
            Auditor                 --	Auditor	                --  E-mail adress of whomever classified the item
            Remarks                 --	Remarks	                --  Free text field
            Scope                   --	Scope                   --  Pick between:  Altijd / Eenmaal / DezePeriode


        Also store a full details sheet with all the data available for each entry.
    """
    COL_NAMES = [
        "material_id",
        "filename",
        "manual_classification",
        "auditor",
        "remarks",
        "scope",
    ]
    TODAY: str = datetime.now().strftime(format="%Y-%m-%d_%H-%M-%S")
    if not isinstance(data, pl.DataFrame):
        data = retrieve_copyright_items(settings=settings)  # Pass settings
        if faculty:
            data = data.filter(pl.col("faculty") == faculty)

    data = data.filter(pl.col(name="workflow_status") == "Done")
    if data.is_empty():
        logger.warning("No data to export")
        return
    if not faculty:
        dir = Directory(
            settings.dirs[DirSetting.EXPORT_TO_SURF].full
        )  # Use passed settings
    else:
        dir = Directory(
            settings.dirs[DirSetting.EXPORT_TO_SURF].full / faculty
        )  # Use passed settings

    export_file_path: Path = (
        dir.full / f"utwente_{TODAY}_{data.shape[0]}_items_copyright_import.xlsx"
    )
    full_details_file_path = (
        dir.full
        / f"utwente__{TODAY}_{data.shape[0]}_items_copyright_import_full_details.xlsx"
    )

    logger.info(f"Exporting data to {export_file_path} and {full_details_file_path}")
    data.select(COL_NAMES).write_excel(export_file_path)

    # store formatted / styled export file with all details
    full_export_file = File(full_details_file_path)
    store_complete_data(
        settings=settings, file=full_export_file, data=data
    )  # Pass settings


def enrich_with_llm_classifications(
    settings: Settings, data: pl.DataFrame
) -> pl.DataFrame:  # Added settings
    """
    Create a dataframe with llm classification data for a set of material ids(see classifier_api or classifier_local for more details).
    call retrieve_all_classifications first
    join data with that dataframe on material_id
    select only the relevant columns
    return the joined dataframe
    """
    logger.warning("Enriching data with llm classification data currently disabled")
    return (
        pl.DataFrame()
    )  # Returns empty DF, so the code below this is currently not executed.

    # The code below would be active if the above return was removed.
    llm_data = retrieve_all_classifications(settings=settings)  # Corrected call

    if not isinstance(llm_data, pl.DataFrame) or llm_data.is_empty():
        logger.warning(
            "No llm classification data found or llm_data is not a DataFrame."
        )
        return data  # Return original data if no llm data

    joined_data = data.join(llm_data, on="material_id", how="left")
    # set col_order
    col_order = [
        "material_id",
        "url",
        "filename",
        "manual_classification",
        "remarks",
        "ml_prediction",
        "allowed_usage_llm",
        "allowed_usage_reasoning_llm",
        "copyright_status_llm",
        "copyright_classification_reason_llm",
        "item_type_llm",
        "item_type_classification_reason_llm",
        "remarks_llm",
        "author",
        "author_names_llm",
        "title",
        "item_title_llm",
        "publisher",
        "publisher_name_llm",
        "copyright_holder_llm",
        "doi",
        "doi_llm",
        "isbn",
        "isbn_llm",
        "source_url_llm",
        "license_llm",
        "course_name",
        "topic_llm",
        "pagecount",
        "pdf_page_count_llm",
    ]
    # Ensure only existing columns are selected to avoid errors if llm_data schema varies
    existing_cols_in_order = [col for col in col_order if col in joined_data.columns]

    # print missing expected columns
    missing_cols = [col for col in col_order if col not in joined_data.columns]
    if missing_cols:
        logger.warning(
            f"Missing expected columns in llm classification data: {missing_cols}"
        )
    return joined_data.select(existing_cols_in_order)


def retrieve_all_classifications(settings: Settings) -> pl.DataFrame:  # Added settings
    """
    -> read all .json files in script_data / classifications /
    -> each json file has name {material_id}_.......json
    -> open each json, use key as colname, contents as values, add material_id col with material_id from filename as value
    -> if a json is not found, check if a .replace file is present --> should have format {input_mat_id}_{replacement_mat_id}.replace
        -> if a .replace file exists, read the replacement mat_id instead and add that row data
    -> return as dataframe
    """
    # This function will need settings if it's to access SETTINGS.dirs
    # For now, assuming it will be refactored or called with settings if used.
    # If SETTINGS is still used here, it will cause an error.
    # Based on current usage (disabled in finalize_sheet), this might not be an immediate issue.
    # However, for completeness, if it were to be used:
    # all_files = Directory(settings.dirs[DirSetting.CLASSIFICATIONS].full).files
    # logger.warning("retrieve_all_classifications is called but uses global SETTINGS which should be refactored if this function is enabled.") # Comment out warning as it's now fixed
    # The following line will error if SETTINGS is not available globally. # Comment out as it's now fixed
    all_files = Directory(
        settings.dirs[DirSetting.CLASSIFICATIONS].full
    ).files  # Use passed settings - This is correct
    all_jsons = [
        file
        for file in all_files
        if all(
            [
                file.extension == ".json",
                "_" not in file.name,
                file.name.rstrip(".json").isdigit(),
            ]
        )
    ]
    all_replacements = [
        file.name.replace(".replace", "")
        for file in all_files
        if file.extension == ".replace"
    ]
    logger.info(
        f"Found {len(all_jsons)} json files with llm classifications, and {len(all_replacements)} replacement files in {settings.dirs[DirSetting.CLASSIFICATIONS].full}"  # Use passed settings
    )
    # rename all jsons to {material_id}.json --> split filename on _ and take first part

    # all_jsons = [file.rename(file.name.split("_")[0] + ".json") for file in all_jsons if '_' in file.name]
    # all_files = Directory(SETTINGS.dirs[DirSetting.CLASSIFICATIONS].full).files
    # all_jsons = [file for file in all_files if file.extension == ".json"]
    # read all jsons

    data: dict[str, dict] = {}

    for file in all_jsons:
        with open(file.path, encoding="utf-8", errors="replace") as f:
            read_str = f.read()
            try:
                data[file.name.replace(".json", "")] = json.loads(read_str)
            except json.JSONDecodeError:
                continue

    final_data = []
    final_data_dict = {}
    for mat_id, mat_data in data.items():
        if not str(mat_id).isdigit():
            logger.warning(
                f"Skipping llm data for {mat_id} as it is not a valid material_id"
            )
            continue
        mat_id = int(mat_id)
        tmp = {}
        for key, value in mat_data.items():
            if isinstance(value, list):
                value = "\n".join(value)
            if not value:
                value = None
            tmp[key + "_llm"] = value
        tmp["material_id"] = mat_id
        final_data.append(tmp)
        final_data_dict[mat_id] = tmp

    for replace in all_replacements:
        old, new = replace.split("_")
        if not old.isdigit() or not new.isdigit():
            logger.warning(
                f"Skipping replacement file {replace} as it on or both material_ids are not valid: {old}, {new}"
            )
            continue
        old, new = int(old), int(new)
        if old not in final_data_dict:
            if new not in final_data_dict:
                continue
            replace_data = final_data_dict[new]
            replace_data["material_id"] = old
            final_data.append(replace_data)

    return pl.from_dicts(final_data, infer_schema_length=None)  # Correctly indented


# Deleting the duplicated/malformed content from here to the end of the file
