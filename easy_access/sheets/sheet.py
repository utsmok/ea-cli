from copy import copy
from dataclasses import dataclass, field
from datetime import datetime
from itertools import batched
from pathlib import Path

import openpyxl
import openpyxl.worksheet
import openpyxl.worksheet.datavalidation
import openpyxl.worksheet.table
import openpyxl.worksheet.worksheet
import polars as pl
import typer
from openpyxl.styles import Alignment, NamedStyle
from openpyxl.worksheet.table import Table as ExcelTable
from openpyxl.worksheet.table import TableStyleInfo

from easy_access.db.retrieve import (
    retrieve_copyright_items,
    retrieve_llm_classifications,
)
from easy_access.settings import DEPARTMENT_MAPPING, SETTINGS, ColInfo, DirSetting
from easy_access.utils import File, info, warn


def read_copyright_export(file: File | None = None) -> tuple[str, pl.DataFrame]:
    """
    Reads in data from the latest copyright export file in the copyright dir;
    or if a file is given, reads in that file.
    Input should be a direct export from the CopyRight tool without any changes.
    """
    try:
        if not file:
            info(
                f"Reading in newest Copyright Data from directory: {SETTINGS.dirs[DirSetting.RAW_COPYRIGHT_DATA]}"
            )
            file = max(
                SETTINGS.dirs[DirSetting.RAW_COPYRIGHT_DATA].files,
                key=lambda x: x.created,
            )

        info(f"Reading in data from:\n            {file.name}\n")
        latest_file_date = file.created.strftime("%Y-%m-%d")
        raw_copyright_data = pl.read_excel(file.path)
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
                    DEPARTMENT_MAPPING, default="Unmapped"
                ),
            )
        )

        # now drop rows we definitely do not want.
        # - drop row if material_id is null, None, blank, or '-'
        # - keep rows with filetype pdf, ppt, doc, or blank ('-'/None/null/""), drop rest
        info(f"Retrieved {len(copyright_data)} items from {file.name}.")

        copyright_data = copyright_data.filter(pl.col("material_id").is_not_null())
        copyright_data = copyright_data.filter(
            (pl.col("filetype").is_in(["pdf", "ppt", "doc", "-"]))
            | (pl.col("filetype").is_null())
        )

        info(
            f"{len(copyright_data)} items remaining from {file.name} after filtering out missing material_ids and specific filetypes."
        )
        return latest_file_date, copyright_data
    except FileNotFoundError:
        warn(f"No files found in {SETTINGS.dirs[DirSetting.RAW_COPYRIGHT_DATA]}")
        raise typer.Exit(code=1)
    except PermissionError:
        warn(f"Permission denied to read {file.name}")
        raise typer.Exit(code=1)
    except ValueError:
        warn("No file found.")
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
            if col.new_name:
                col_name = col.new_name
            else:
                col_name = col.name
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


def finalize_sheet(file: File, data: pl.DataFrame, style_iter: int) -> None:
    """
    This function takes an excel file with 'Complete Data' and adds a data entry sheet +styling.
    Input: an excel file with the complete data, and a dataframe with that same data to be processed for the data entry sheet

    Adds the sheet to the workbook and saves it. Returns the incremented style_iter var.
    """

    wb = openpyxl.load_workbook(filename=str(file.path))
    if SETTINGS.data_settings.complete_data_name not in wb.sheetnames:
        wb.active.title = SETTINGS.data_settings.complete_data_name

    tabstyle = TableStyleInfo(
        name=f"TableStyleMedium{style_iter}",
        showRowStripes=True,
    )
    style_iter = style_iter + 1
    sheet = DataEntrySheet(
        workbook=wb,
        sheet_name=SETTINGS.data_settings.data_entry_name,
        cols=SETTINGS.data_settings.data_entry_cols,
        table_style=tabstyle,
        file_path=str(file.path),
    )

    data = data.unique("material_id")
    sheet.add_data(data)
    info(f"Added data entry sheet to {file.name}")
    llm_classification_data = enrich_with_llm_classifications(data)
    if isinstance(llm_classification_data, pl.DataFrame):
        llm_sheet_path = (
            file.path.parent / f"{file.path.stem}_llm_classification_data.xlsx"
        )
        llm_classification_data.write_excel(
            workbook=llm_sheet_path,
            worksheet="llm_classification_data",
            table_name="llm_classification_data",
            table_style="TableStyleMedium3",
            autofit=True,
        )
        info(f"Stored llm_classification_data sheet to {llm_sheet_path.name}")

    return style_iter


def store_complete_data(file: File | Path, data: pl.DataFrame) -> None:
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
        for col in SETTINGS.data_settings.final_data_col_order
        if col in data.columns
    ]
    data = data.select(selectcols)
    data = data.unique("material_id")
    data.write_excel(file, worksheet=SETTINGS.data_settings.complete_data_name)
    info(f"Stored {data.shape[0]} rows to {file}")


def read_export_sheets() -> pl.DataFrame:
    """
    Read in the export sheets from the EXPORT_TO_SURF dir
    return as concatenated dataframe
    """
    returndata = pl.DataFrame()
    for file in SETTINGS.dirs[DirSetting.EXPORT_TO_SURF].files:
        if file.extension in [".xls", ".xlsx"]:
            returndata = pl.concat(
                [returndata, pl.read_excel(file.path)], how="diagonal_relaxed"
            )
        if file.extension in [".csv"]:
            returndata = pl.concat(
                [returndata, pl.read_csv(file.path)], how="diagonal_relaxed"
            )
    return returndata


def create_export_sheet(data: pl.DataFrame | None = None, print_overview: bool = True):
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
            Owner                   --	Owner	                --  E-mail adress of whomever classified the item aanpasser
            Remarks                 --	Remarks	                --  Free text field
            Scope                   --	Scope                   --  Pick between:  Altijd / Eenmaal / DezePeriode


        Also store a full details sheet with all the data available for each entry.

        Returns a list of material_ids for the items that were exported, for further processing.

        #TODO: Determine if for 'Owner' the actual 'Owner' field in copyright data should be used, or if it should be 'auditor' instead??

    """
    COL_NAMES = [
        "material_id",
        "filename",
        "manual_classification",
        "owner",
        "remarks",
        "scope",
    ]
    TODAY: str = datetime.now().strftime(format="%Y-%m-%d_%H-%M-%S")
    if False:
        # previous implementation
        data = data.filter(pl.col(name="workflow_status") == "Done")
        full_data: pl.DataFrame = copy(x=data)

        data = data.select([col for col in COL_NAMES if col in data.columns])
        data = data.rename(
            mapping={col: col.replace("_", "").lower() for col in data.columns}
        )

        existing_data: pl.DataFrame = read_export_sheets()
        if not existing_data.is_empty():
            data = data.join(other=existing_data, on="materialid", how="anti")

        if data.is_empty():
            warn(text="No new data found to export!")

        full_data = full_data.filter(
            (pl.col(name="material_id").is_in(other=data["materialid"]))
            & (pl.col(name="workflow_status") == "Done")
        )
        material_ids_exported: list[str] = data["Material id"].to_list()

        if not print_overview:
            return material_ids_exported

        # messy code to print some data on what is being exported
        overview = {
            "Number of items per faculty": full_data.group_by("faculty")
            .len()
            .sort("len", descending=True)
            .to_dicts(),
            "Number of items per classification": full_data.group_by(
                "manual_classification"
            )
            .len()
            .sort("len", descending=True)
            .to_dicts(),
            "Number of items per ml_prediction": full_data.group_by("ml_prediction")
            .len()
            .sort("len", descending=True)
            .to_dicts(),
            "man_class per ml_pred": full_data.group_by(
                ["ml_prediction", "manual_classification"]
            )
            .agg(pl.len().alias(name="len"))
            .sort("ml_prediction", "len", descending=[False, True])
            .to_dicts(),
        }
        print(overview)
        info(text=f"Creating export sheet with {data.shape[0]} rows.")
        print()
        print()
        for key, value in overview.items():
            print(
                "    -------------------------------------------------------------------"
            )
            print(f"                                   {key}")
            print(
                "    -------------------------------------------------------------------"
            )
            if "man_class per ml_pred" in key:
                batch_num = 3

            else:
                batch_num = 2
            maxgap: int = max(
                [max([len(str(object=val)) for val in item.values()]) for item in value]
            )
            if batch_num == 3:
                maxgap = maxgap * 2
            for n, item in enumerate(value):
                item: dict[str, int] = item
                if n == 0:
                    if batch_num == 3:
                        print(
                            f"      {list(item.keys())[0]}:{list(item.keys())[1]}{' ' * (maxgap - len(list(item.keys())[0]) - len(list(item.keys())[1]))} |     {list(item.keys())[2]}"
                        )
                        print(
                            f" {'-' * (len(str(list(item.keys())[0]) + ':' + str(list(item.keys())[1])) + 5)}{'-' * (maxgap - len(str(list(item.keys())[0]) + ':' + str(list(item.keys())[1])) + 4)}|{'-' * (len(list(item.keys())[2]) + 10)}"
                        )

                    else:
                        print(
                            f"      {list(item.keys())[0]}{' ' * (maxgap - len(list(item.keys())[0]) - 3)} |     {list(item.keys())[1]}"
                        )
                        print(
                            f" {'-' * (len(list(item.keys())[0]) + 5)}{'-' * (maxgap - len(list(item.keys())[0]) + 4)}|{'-' * (len(list(item.keys())[1]) + 10)}"
                        )
                for results in batched(item.items(), batch_num):
                    if batch_num != 3:
                        key = results[0][1]
                        value = results[1][1]
                    else:
                        key = f"{results[0][1]} --> {results[1][1]}"
                        value = results[2][1]
                    print(f"      {key}{' ' * (maxgap - len(key) + 3)} |     {value}")

    data = retrieve_copyright_items()
    # get all data from db
    # filter out only 'Done' items (?)
    # write to the two excel files with different sets of cols
    export_file_path: Path = (
        SETTINGS.dirs[DirSetting.EXPORT_TO_SURF].full
        / f"utwente_{TODAY}_{data.shape[0]}_items_copyright_import.xlsx"
    )
    full_details_file_path: Path = (
        SETTINGS.dirs[DirSetting.EXPORT_TO_SURF].full
        / f"utwente_{TODAY}_{data.shape[0]}_items_copyright_import_full_details.xlsx"
    )
    info(f"Exporting data to {export_file_path} and {full_details_file_path}")
    data = data.filter(pl.col(name="workflow_status") == "Done")
    data.select(COL_NAMES).write_excel(export_file_path)

    # store formatted / styled export file with all details
    full_export_file = File(full_details_file_path)
    store_complete_data(file=full_export_file, data=data)


def enrich_with_llm_classifications(data: pl.DataFrame) -> pl.DataFrame:
    """
    Create a dataframe with llm classification data for a set of material ids(see classifier_api or classifier_local for more details).
    call retrieve_all_classifications first
    join data with that dataframe on material_id
    select only the relevant columns
    return the joined dataframe
    """

    llm_data = retrieve_llm_classifications(
        selected_material_ids=data.select("material_id")
        .unique("material_id")
        .to_series()
        .to_list()
    )
    if not isinstance(llm_data, pl.DataFrame):
        warn("No llm classification data found.")
        return None
    else:
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
    col_order = [col for col in col_order if col in joined_data.columns]
    # print missing expected columns
    missing_cols = [col for col in col_order if col not in joined_data.columns]
    if missing_cols:
        warn(f"Missing expected columns in llm classification data: {missing_cols}")
    return joined_data.select(col_order)
