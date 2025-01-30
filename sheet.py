from dataclasses import dataclass, field
from utils import File, Directory, info, warn, cool
import openpyxl
import polars as pl
from openpyxl.styles import NamedStyle, Alignment
from openpyxl.worksheet.table import TableStyleInfo
from openpyxl.worksheet.table import Table as ExcelTable

import openpyxl
import openpyxl.worksheet
import openpyxl.worksheet.datavalidation
import openpyxl.worksheet.table
import openpyxl.worksheet.worksheet

def finalize_sheet(file: File, data: pl.DataFrame, style_iter: int) -> None:
    """
    This function takes an excel file with 'Complete Data' and adds a data entry sheet +styling.
    Input: an excel file with the complete data, and a dataframe with that same data to be processed for the data entry sheet

    Adds the sheet to the workbook and saves it. Returns the incremented style_iter var.
    """

    @dataclass
    class ColInfo:
        """
        contains the info for a single col used in a DataEntrySheet
        """

        name: str  # the colname as included in the sheet (e.g. 'manual_classification')
        dropdown_options: str = (
            ""  # the options for the dropdown; if not applicable, an empty str
        )
        is_url: bool = False  # format as url or not?
        is_new: bool = (
            False  # if True, this col is not present in the original data
        )
        is_editable: bool = False  # if True, this col can be edited
        new_name: str = ""  # if not empty, this col will be renamed to this name
        default_val: str = (
            ""  # if 'is_new' is True, use this as the default value for the new col
        )
        max_width: int = 8  # the max length of any value present in this col, to be set while processing. Min width is this initial number.
        count_max_width_over_40: int = 0  # the number of items in this col that are longer than 40 chars, to be set while processing

        @property
        def has_dropdown(self) -> bool:
            return len(self.dropdown_options) > 0

    @dataclass
    class DataEntrySheet:
        """
        Use to add a dateentry sheet to an excel file.
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
            self.sheet = wb.create_sheet(self.sheet_name, index=1)

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

            info(f"Added data to {self.sheet_name} in file {self.file_path}.")
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
            info(
                f"Created table with {self.max_row} rows and {len(self.cols)} cols in sheet {self.sheet_name} of file {self.file_path}"
            )

        def save(self) -> None:
            self.workbook.save(filename=self.file_path)
            info(f"Saved .xlsx file with DataEntrySheet to {self.file_path}")

    wb = openpyxl.load_workbook(filename=str(file.path))
    wb.active.title = "Complete data"

    tabstyle = TableStyleInfo(
        name=f"TableStyleMedium{style_iter}",
        showRowStripes=True,
    )
    style_iter = style_iter + 1

    sheet = DataEntrySheet(
        workbook=wb,
        sheet_name="Data entry",
        cols=[
            ColInfo("material_id"),
            ColInfo("url", is_url=True),
            ColInfo(
                "workflow_status",
                is_new=True,
                is_editable=True,
                dropdown_options='"ToDo,Done,InProgress"',
                default_val="ToDo",
            ),
            ColInfo(
                "manual_classification",
                is_editable=True,
                default_val="-",
                dropdown_options='"open access,eigen materiaal - powerpoint,eigen materiaal - overig,lange overname,eigen materiaal - titelindicatie,anders,korte overname,middellange overname,-"',
            ),
            ColInfo("remarks", is_editable=True),
            ColInfo("ml_prediction"),
            ColInfo("filename"),
            ColInfo("title"),
            ColInfo("owner", new_name="uploaded_by"),
            ColInfo("author", new_name="detected_author"),
            ColInfo("contact_name"),
            ColInfo("contact_email"),
            ColInfo("contact_org"),
            ColInfo("osiris_catalogue_url", is_url=True),
            ColInfo("course_name", new_name="course_name_canvas"),
            ColInfo("department", new_name="programme_canvas"),
            ColInfo("osiris_programme", new_name="programme_osiris"),
            ColInfo("osiris_course_codes_found"),
            ColInfo("osiris_course_code_data_selected"),
        ],
        table_style=tabstyle,
        file_path=str(file.path),
    )

    sheet.add_data(data)
    return style_iter
