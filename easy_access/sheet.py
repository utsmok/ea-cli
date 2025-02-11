from copy import copy
import polars as pl

from easy_access.settings import SETTINGS, ColInfo, DirSetting, DEPARTMENT_MAPPING
from dataclasses import dataclass, field
from easy_access.utils import File, info, warn, cool
from pathlib import Path
from datetime import datetime
from itertools import batched

import openpyxl
from openpyxl.styles import NamedStyle, Alignment
import openpyxl.worksheet
import openpyxl.worksheet.datavalidation
import openpyxl.worksheet.table
import openpyxl.worksheet.worksheet
from openpyxl.worksheet.table import TableStyleInfo
from openpyxl.worksheet.table import Table as ExcelTable
import typer

def read_other_sheet(file: File) -> pl.DataFrame:
        """
        Reads in the data from another sheet as the datasource, instead of using CopyRight data.
        Sheet should be formatted in the same way as the faculty output sheets.
        It will read in the first sheet in the .xlsx file.
        It will do a quick check on the columns in the sheets to prevent the most basic errors.
        """

        info(f"Reading in data from {file.name}")
        copyright_data = pl.read_excel(file.path)
        latest_file_date = file.modified.strftime("%Y-%m-%d")
        info(
            f"Read {len(copyright_data)} items from {file.name}. Item was lasted changed on {latest_file_date}"
        )

        if "workflow_status" not in copyright_data.columns:
            copyright_data = copyright_data.with_columns(
                pl.Series("workflow_status", ["ToDo"] * len(copyright_data))
            )
        if "retrieved_from_copyright_on" not in copyright_data.columns:
            if "added_to_sheet_on" not in copyright_data.columns:
                copyright_data = copyright_data.with_columns(
                    pl.Series(
                        "retrieved_from_copyright_on",
                        [latest_file_date] * len(copyright_data),
                    )
                )
            else:
                copyright_data = copyright_data.rename(
                    {"added_to_sheet_on": "retrieved_from_copyright_on"}
                )

        latest_file_date = max(
            copyright_data.select(pl.col("retrieved_from_copyright_on"))
            .to_series()
            .to_list()
        )

        return latest_file_date, copyright_data.select(SETTINGS.data_settings.complete_data_cols)

def read_copyright_export() -> tuple[str, pl.DataFrame]:
        """
        Reads in data from the latest copyright export file in the copyright dir.
        """

        info(
            f"Reading in newest Copyright Data from directory: {SETTINGS.dirs[DirSetting.RAW_COPYRIGHT_DATA]}"
        )
        try:
            all_files = SETTINGS.dirs[DirSetting.RAW_COPYRIGHT_DATA].files
            latest_file = max(all_files, key=lambda x: x.created)
            latest_file_date = latest_file.created.strftime("%Y-%m-%d")
            info(
                f"Selected newest copyright export file:\n          {latest_file.name}\n          created @ {latest_file_date}"
            )
            raw_copyright_data = pl.read_excel(latest_file.path)
            # cast all columns to str
            raw_copyright_data =  raw_copyright_data.with_columns(
                pl.exclude(pl.Utf8).cast(str)
            )

            copyright_data =  raw_copyright_data.rename(
                lambda col: col.replace(" ", "_")
                .replace("#", "count_")
                .replace("*", "x")
                .lower()
            ).with_columns(
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
                faculty=pl.col("department").replace_strict(
                    DEPARTMENT_MAPPING, default="Unmapped"
                ),
            )

            return latest_file_date, copyright_data

        except FileNotFoundError:
            warn(f"No files found in {SETTINGS.dirs[DirSetting.RAW_COPYRIGHT_DATA]}")
            raise typer.Exit(code=1)
        except PermissionError:
            warn(f"Permission denied to read {SETTINGS.latest_file.name}")
            raise typer.Exit(code=1)
        except ValueError:
            warn(f"No files found in {SETTINGS.dirs[DirSetting.RAW_COPYRIGHT_DATA]}")
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
                    if len(cell_data) > col.max_width:
                        col.max_width = len(cell_data)
                    if len(cell_data) > 40:
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

    sheet.add_data(data)
    return style_iter

def store_complete_data(file: File | Path, data: pl.DataFrame) -> None:
    """
    Stores the given data in an excel file with 1 sheet named SETTINGS.data_settings.complete_data_name
    using the col order in SETTINGS.data_settings.final_data_col_order
    """
    if isinstance(file, File):
        file = file.path

    selectcols = [col for col in SETTINGS.data_settings.final_data_col_order if col in data.columns]
    data = data.select(selectcols)
    data.write_excel(file, worksheet=SETTINGS.data_settings.complete_data_name)
    info(f'Stored {data.shape[0]} rows to {file}')

def read_export_sheets() -> pl.DataFrame:
    """
    Read in the export sheets from the EXPORT_TO_SURF dir
    return as concatenated dataframe
    """
    returndata = pl.DataFrame()
    for file in SETTINGS.dirs[DirSetting.EXPORT_TO_SURF].files:
        if file.extension in [".xls", ".xlsx"]:
            returndata = pl.concat([returndata, pl.read_excel(file.path)], how="diagonal_relaxed")
        if file.extension in [".csv"]:
            returndata = pl.concat([returndata, pl.read_csv(file.path)], how="diagonal_relaxed")
    return returndata

def create_export_sheet(data: pl.DataFrame) -> list[str]:
    """
    Create an export sheet to import back into CopyRight tool


    Will store an excel sheet with the following columns:
    Material id
    Filename
    Manual classification
    Owner
    Remarks
    Scope

    And also store a full details sheet with all the data available for each entry.

    Returns a list of material_ids for the items that were exported.
    """

    col_name_mapping: dict[str, str] = {
        'material_id':"Material id",
        'filename':"Filename",
        'manual_classification':"Manual classification",
        'owner':"Owner",
        'remarks':"Remarks",
        'scope':"Scope"
    }

    alt_col_names: dict[str, str] = { # 'expected field name':'alternative name'
        'owner':'uploaded_by',
    }


    # extract the cols from data using col_name_mapping
    # if any cols are missing, try using alt_col_names
    final_selected_colnames: dict[str,str] = {}
    for col in col_name_mapping:
        if col not in data.columns:
            if col not in alt_col_names:
                raise Exception(f"While building export sheet:Could not find column {col} in data: {data.head(n=5)} with columns {data.columns}")
            if alt_col_names[col] in data.columns:
                final_selected_colnames[alt_col_names[col]] = col_name_mapping[col]
            else:
                raise Exception(f"While building export sheet:Could not find alternative column name {alt_col_names[col]} in data: {data.head(n=5)} with columns {data.columns}")
        else:
            final_selected_colnames[col] = col_name_mapping[col]

    data = data.filter(pl.col(name='workflow_status') == 'Done')

    full_data: pl.DataFrame = copy(x=data)
    data = data.select(final_selected_colnames.keys()).rename(mapping=final_selected_colnames)
    existing: pl.DataFrame = read_export_sheets()
    if not existing.is_empty():
        data = data.join(other=existing, on='Material id', how='anti')

    if data.is_empty():
        warn(text="No new data found to export!")

    full_data = full_data.filter((pl.col(name='material_id').is_in(other=data['Material id'])) & (pl.col(name='workflow_status') == 'Done'))

    overview = {

        "Number of items per faculty": full_data.group_by("faculty").len().sort("len", descending=True).to_dicts(),
        "Number of items per classification": full_data.group_by("manual_classification").len().sort("len", descending=True).to_dicts(),
        "Number of items per ml_prediction": full_data.group_by("ml_prediction").len().sort("len", descending=True).to_dicts(),
        "man_class per ml_pred": full_data.group_by(["ml_prediction", "manual_classification"]).agg(pl.len().alias(name='len')).sort("ml_prediction","len", descending=[False, True]).to_dicts(),
    }
    print(overview)
    info(text=f'Creating export sheet with {data.shape[0]} rows.')
    print()
    print()
    for key, value in overview.items():
        print("    -------------------------------------------------------------------")
        print(f"                                   {key}")
        print("    -------------------------------------------------------------------")
        if 'man_class per ml_pred' in key:
            batch_num = 3

        else:
            batch_num = 2
        maxgap: int = max([max([len(str(object=val)) for val in item.values()]) for item in value])
        if batch_num == 3:
            maxgap = maxgap * 2
        for n, item in enumerate(value):
            item: dict[str, int] = item
            if n == 0:
                if batch_num == 3:
                    print(f"      {list(item.keys())[0]}:{list(item.keys())[1]}{" "*(maxgap-len(list(item.keys())[0])-len(list(item.keys())[1]))} |     {list(item.keys())[2]}")
                    print(f" {'-'*(len(str(list(item.keys())[0])+":"+str(list(item.keys())[1]))+5)}{"-"*(maxgap-len(str(list(item.keys())[0])+":"+str(list(item.keys())[1]))+4)}|{'-'*(len(list(item.keys())[2])+10)}")

                else:
                    print(f"      {list(item.keys())[0]}{" "*(maxgap-len(list(item.keys())[0])-3)} |     {list(item.keys())[1]}")
                    print(f" {'-'*(len(list(item.keys())[0])+5)}{"-"*(maxgap-len(list(item.keys())[0])+4)}|{'-'*(len(list(item.keys())[1])+10)}")
            for results in batched(item.items(),batch_num):
                if batch_num != 3:
                    key = results[0][1]
                    value = results[1][1]
                else:
                    key = f"{results[0][1]} --> {results[1][1]}"
                    value = results[2][1]
                print(f"      {key}{" "*(maxgap-len(key)+3)} |     {value}")

    material_ids_exported: list[str] = data["Material id"].to_list()

    today: str = datetime.now().strftime(format="%Y-%m-%d_%H-%M-%S")

    export_file_path: Path = SETTINGS.dirs[DirSetting.EXPORT_TO_SURF].full / f"utwente_{today}_{data.shape[0]}_items_copyright_import.xlsx"
    data.write_excel(workbook=export_file_path)
    full_details_file_path: Path = SETTINGS.dirs[DirSetting.EXPORT_TO_SURF].full / f"utwente_{today}_{data.shape[0]}_items_copyright_import_full_details.xlsx"
    full_data.write_excel(workbook=full_details_file_path)

    return material_ids_exported
