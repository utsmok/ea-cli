import contextlib
import hashlib
import logging
import os
import warnings
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import openpyxl.worksheet
import openpyxl.worksheet.datavalidation
import openpyxl.worksheet.worksheet
import polars as pl
import typer
from loguru import logger
from openpyxl.styles import Alignment, NamedStyle
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table as ExcelTable
from openpyxl.worksheet.table import TableStyleInfo

# from easy_access.settings import DEPARTMENT_MAPPING, SETTINGS, ColInfo, DirSetting # Will be passed as parameters
from easy_access.settings import ColInfo, DirSetting, Settings  # Keep for type hinting
from easy_access.utils import File, standardize_dataframe


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
            with (
                open(os.devnull, "w") as devnull,
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
        copyright_data = standardize_dataframe(raw_copyright_data)

        # Only set default workflow_status if the column doesn't exist or is all null/empty
        columns_to_add = {
            "retrieved_from_copyright_on": [latest_file_date] * len(copyright_data),
        }

        if (
            "workflow_status" not in copyright_data.columns
            or copyright_data["workflow_status"].is_null().all()
            or (copyright_data["workflow_status"].str.strip_chars().eq("").all())
        ):
            columns_to_add["workflow_status"] = ["ToDo"] * len(copyright_data)

        copyright_data = copyright_data.with_columns(
            **{k: pl.Series(k, v) for k, v in columns_to_add.items()},
        ).with_columns(
            pl.col("last_change")
            .str.replace(r"^-", "")
            .str.strip_chars()
            .str.strptime(pl.Date, "%Y-%m-%d", strict=False)
            .dt.strftime("%Y-%m-%d"),
            pl.col("classification").str.to_lowercase(),
            faculty=pl.col("department").replace_strict(
                settings.university_settings.department_mapping, default="Unmapped"
            ),
        )

        # now drop rows we definitely do not want.
        # - drop row if material_id is null, None, blank, or '-'
        # - keep rows with filetype pdf, ppt, doc, or blank ('-'/None/null/"'), drop rest
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
        )
        raise typer.Exit(code=1)
    except PermissionError:
        logger.warning(f"Permission denied to read {file}")
        raise typer.Exit(code=1)
    except ValueError:
        logger.warning("No file found.")
        raise typer.Exit(code=1)


def add_v2_classification(data: pl.DataFrame) -> pl.DataFrame:
    """
    Given a dataframe with a 'manual_classification' column, adds a v2 classification cols
    using the mapping in db.models.CLASSIFICATION_MAPPING_V1_TO_V2.
    Adds these cols:
    - manual_classification_v2
    - lengte
    - overname_status

    all are defined as enumstrs in db.models.ClassificationV2, LengthCategory, OvernameStatus
    """

    from easy_access.db.enums import CLASSIFICATION_MAPPING_V1_TO_V2, Classification

    input_data = data.select(["manual_classification", "material_id"]).to_dicts()
    update_data = []
    for entry in input_data:
        current_classification = entry.get("manual_classification", "onbekend")
        if not current_classification or current_classification == "-":
            current_classification = "onbekend"
        mapped_data = CLASSIFICATION_MAPPING_V1_TO_V2.get(
            Classification(current_classification)
        )

        if not mapped_data:
            mapped_data = CLASSIFICATION_MAPPING_V1_TO_V2[Classification.ONBEKEND]

        new_data = {
            "manual_classification_v2": mapped_data.classification.value,
            "length": mapped_data.length.value,
            "overname_status": mapped_data.overname_status.value,
        }
        entry.update(new_data)
        update_data.append(entry)
        new_data = pl.DataFrame(update_data, infer_schema_length=len(update_data))
    data = data.join(new_data, on="material_id", how="left")
    return data


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
    word_wrap_style: NamedStyle = NamedStyle(
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

        # Build a deterministic mapping from column index -> Excel column letter
        header_col_letters: list[str] = [
            get_column_letter(i + 1) for i in range(len(self.cols))
        ]

        for idx, col in enumerate(self.cols):
            # use column position (order in self.cols) to determine target column
            col_index_1based = idx + 1
            target_col_letter = header_col_letters[idx]

            if col.has_dropdown:
                wb = self.workbook
                list_sheet_name = "_ea_lists"

                # Normalize dropdown source: support list/tuple or comma-separated string, strip outer quotes
                raw_opts = col.dropdown_options
                items: list[str] | None = None
                formula1: str | None = None

                if isinstance(raw_opts, list | tuple):
                    items = [
                        str(x).strip().strip('"').strip("'")
                        for x in raw_opts
                        if x is not None
                    ]
                else:
                    opts = (str(raw_opts) if raw_opts is not None else "").strip()
                    # If this looks like a formula/range or already quoted literal, use directly
                    if (
                        (opts.startswith('"') and opts.endswith('"'))
                        or opts.startswith("=")
                        or "!" in opts
                        or opts.startswith("{")
                    ):
                        formula1 = opts
                        # If it's a quoted literal like "A,B,C", derive items for cleaning later
                        if opts.startswith('"') and opts.endswith('"'):
                            inner = opts[1:-1]
                            items = [
                                s.strip().strip('"').strip("'")
                                for s in inner.split(",")
                                if s.strip()
                            ]
                    else:
                        # plain comma-separated string
                        inner = opts
                        items = [
                            s.strip().strip('"').strip("'")
                            for s in inner.split(",")
                            if s.strip()
                        ]

                # If we have a list of items, create (or reuse) a named range on a hidden sheet
                if items is not None and len(items) > 0:
                    if list_sheet_name in wb.sheetnames:
                        list_ws = wb[list_sheet_name]
                    else:
                        list_ws = wb.create_sheet(list_sheet_name)
                        list_ws.sheet_state = "hidden"

                    # deterministic name for identical lists
                    name_hash = hashlib.md5(
                        ",".join(items).encode("utf-8")
                    ).hexdigest()[:8]
                    list_name = f"_ea_list_{name_hash}"

                    # Check existing defined names by name (wb.defined_names yields names)
                    existing_names = list(wb.defined_names)
                    if list_name not in existing_names:
                        # find first empty column in the lists sheet
                        col_idx = 1
                        while list_ws.cell(row=1, column=col_idx).value is not None:
                            col_idx += 1

                        for ridx, val in enumerate(items, start=1):
                            list_ws.cell(row=ridx, column=col_idx).value = val

                        list_col_letter = get_column_letter(col_idx)
                        # Use quote_sheetname to handle special chars and spaces
                        ref = f"{quote_sheetname(list_sheet_name)}!${list_col_letter}$1:${list_col_letter}${len(items)}"
                        dn = DefinedName(list_name, attr_text=ref)
                        # add the defined name to the workbook
                        wb.defined_names.add(dn)

                    # Use the named range as the validation source
                    formula1 = f"={list_name}"

                # If formula1 is still None for some reason, fall back to empty literal (will produce an invalid DV)
                if not formula1:
                    formula1 = '""'

                # create DataValidation and enforce strict selection
                dv = openpyxl.worksheet.datavalidation.DataValidation(
                    type="list",
                    formula1=formula1,
                    allow_blank=False,
                    showDropDown=False,
                )
                dv.error = "Please select a valid option from the list"
                dv.errorTitle = "Invalid option"
                dv.prompt = "Please select from the list"
                dv.promptTitle = "List selection"
                # prefer a strict stop style so Excel validates properly
                dv.errorStyle = "stop"
                dv.showErrorMessage = True
                dv.showInputMessage = True

                self.sheet.add_data_validation(dv)
                start_cell = f"{target_col_letter}2"
                end_cell = f"{target_col_letter}{self.max_row + 1}"
                if self.max_row == 1:
                    dv.add(start_cell)
                else:
                    dv.add(f"{start_cell}:{end_cell}")

            # Apply column width / wrap behavior for the target column
            if col.max_width > 40 and (
                (col.count_max_width_over_40 > 5)
                or (col.count_max_width_over_40 > self.max_row - 2)
            ):
                # Too much long items: cap width to 40 & enable word wrap for this col
                for row in range(2, self.max_row + 1):
                    self.sheet.cell(row, col_index_1based).style = self.word_wrap_style
                self.sheet.column_dimensions[target_col_letter].bestFit = False
                self.sheet.column_dimensions[target_col_letter].width = 40
            else:
                # Acceptable width, don't enable word wrap but fit width to contents
                self.sheet.column_dimensions[target_col_letter].width = col.max_width

            if col.style and (
                isinstance(col.style.activate_on, list)
                and len(col.style.activate_on) > 0
            ):
                try:
                    rule = col.style.cf_rule
                    self.sheet.conditional_formatting.add(
                        f"{target_col_letter}2:{target_col_letter}{self.max_row + 1}",
                        rule,
                    )
                except Exception as e:
                    logger.warning(
                        f"Could not add conditional formatting to col {col.name} ({target_col_letter}): {e}"
                    )

        self.create_table()
        self.save()

    def create_table(self) -> None:
        # use get_column_letter to support multi-letter column names beyond 'Z'
        max_col_letter = get_column_letter(len(self.cols))
        table = ExcelTable(
            displayName=self.sheet_name.replace(" ", ""),
            ref=f"A1:{max_col_letter}{self.max_row + 1}",
        )
        table.tableStyleInfo = self.table_style
        self.sheet.add_table(table)

    def save(self) -> None:
        try:
            self.workbook.active = self.sheet
        except Exception:
            self.workbook.active = 1

        self.sheet.sheet_view.tabSelected = True
        self.workbook["Complete data"].sheet_view.tabSelected = False
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
    if settings.data_settings.complete_data_name not in wb.sheetnames and wb.active:
        wb.active.title = settings.data_settings.complete_data_name

    tabstyle = TableStyleInfo(
        name=f"TableStyleMedium{style_iter}",
        showRowStripes=True,
    )
    style_iter = style_iter + 1
    sheet = DataEntrySheet(
        workbook=wb,
        sheet_name=settings.data_settings.data_entry_name,
        cols=settings.data_settings.data_entry_cols,
        table_style=tabstyle,
        file_path=str(file.path),
    )
    data = data.unique("material_id")
    sheet.add_data(data)
    logger.info(f"Added data entry sheet to {file.name}")

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

    def validate_export_dataframe(df: pl.DataFrame, required_cols: set[str]) -> None:
        missing = required_cols - set(df.columns)
        if missing:
            raise ValueError(
                f"Export dataframe missing required columns: {sorted(missing)}"
            )

    # Validate minimal required columns before any column selection/unique operations
    validate_export_dataframe(data, required_cols={"material_id"})

    selectcols = [
        col
        for col in settings.data_settings.final_data_col_order
        if col in data.columns
    ]
    # Ensure several related course/contact columns are preserved in the Complete Data
    extra_export_cols = [
        "course_contacts_faculties",
        "course_contacts_names",
        "course_contacts_emails",
        "course_contacts_organizations",
        "course_names",
        "cursuscodes",
        "programmes",
    ]
    for c in extra_export_cols:
        if c in data.columns and c not in selectcols:
            selectcols.append(c)
    data = data.select(selectcols)
    # Deduplicate by material_id (material_id guaranteed to exist due to earlier validation)
    data = data.unique("material_id")

    # Remove internal-only columns from the exported Complete Data
    for drop_col in ("is_duplicate", "replacement_id"):
        if drop_col in data.columns:
            data = data.drop(drop_col)

    # Use atomic write: write to temp file then rename into place
    target_path = Path(file)
    tmp_path = target_path.with_suffix(target_path.suffix + ".tmp")
    try:
        data.write_excel(tmp_path, worksheet=settings.data_settings.complete_data_name)
        # os.replace is atomic on most platforms
        os.replace(tmp_path, target_path)
    finally:
        if tmp_path.exists():
            with contextlib.suppress(Exception):
                tmp_path.unlink()

    logger.info(f"Stored {data.shape[0]} rows to {file}")


def protect_workbook(
    file_path: str | Path,
    protect_sheets: list[str] | None = None,
    active_sheet: str | None = None,
    password: str | None = None,
) -> None:
    """Protect specified sheets in an existing workbook and set the active sheet.

    This is intended to be run after the atomic write has placed the final file.
    It will attempt to open the workbook with openpyxl, enable protection on the
    requested sheets (if present) and set the workbook.active to the requested
    sheet name if available.
    """
    try:
        wb = openpyxl.load_workbook(filename=str(file_path))
    except Exception as e:
        logger.warning(f"Could not open workbook for protection: {file_path}: {e}")
        return

    protect_sheets = protect_sheets or []

    for name in protect_sheets:
        if name in wb.sheetnames:
            ws = wb[name]
            try:
                # enable sheet protection; set a password if provided
                ws.protection.sheet = True
                if password:
                    # openpyxl will hash the password as needed
                    ws.protection.set_password(password)
            except Exception as e:
                logger.debug(f"Could not set protection on sheet {name}: {e}")

    if active_sheet and active_sheet in wb.sheetnames:
        try:
            wb.active = wb[active_sheet]
        except Exception:
            # best-effort: ignore if setting the active sheet fails
            pass

    try:
        wb.save(filename=str(file_path))
    except Exception as e:
        logger.warning(f"Failed saving workbook after protection step: {e}")


def read_faculty_sheets(settings: Settings) -> pl.DataFrame:
    """
    Reads all faculty sheets and returns a single DataFrame.
    """
    all_dfs = []
    select_cols: list[str] = [
        "material_id",
        "workflow_status",
        "remarks",
        "manual_classification",
    ]
    pl.DataFrame(schema={col: pl.Utf8 for col in select_cols})

    for faculty_dir in settings.dirs[DirSetting.FACULTIES_DIR].dirs():
        for file in faculty_dir.files_r:
            if (
                file.extension == ".xlsx"
                and "overview" not in file.name
                and "llm" not in file.name
            ):
                try:
                    df = _read_excel_quiet(
                        file.path, sheet_name=settings.data_settings.data_entry_name
                    )
                    for col_name in select_cols:
                        if col_name not in df.columns:
                            df = df.with_columns(
                                pl.lit(None).alias(col_name).cast(pl.Utf8)
                            )
                        else:
                            df = df.with_columns(pl.col(col_name).cast(pl.Utf8))
                    df = df.select(
                        select_cols
                    )  # Ensure correct column order and selection

                    all_dfs.append(df)
                except Exception as e:
                    logger.warning(f"Error reading {file.path}: {e}")
                    continue
    if not all_dfs:
        return pl.DataFrame()

    return pl.concat(all_dfs)
