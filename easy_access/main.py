from collections import defaultdict
import asyncio
import os
from datetime import datetime
import polars as pl
import typer
from loguru import logger
from easy_access.utils import Directory, File, info, cool, warn, print
from easy_access.enrichment import enrich_df_with_osiris_data, update_osiris_data
from easy_access.sheet import (
    finalize_sheet,
    store_complete_data,
    read_copyright_export,
    read_other_sheet,
    create_export_sheet
)
from easy_access.analysis import create_faculty_overviews
from easy_access.settings import (
    SETTINGS,
    FileSetting,
    DirSetting,
    EasyAccessSettings,
    Functions,
    COURSE_MAPPING
)




class EasyAccessTool:
    """
    Main class to run the entire app.
    Runs functions based on settings (see settings.yaml and settings.py) & cli input (see run.py)

    """

    def __init__(self, settings: EasyAccessSettings) -> None:
        """
        Parameters:
            setting:  str | None
                Pick which functions to run when self.run() is called. If no argument is passed, it will run all functions.
                pick from one of the presets below:
                    'none' -> don't run any functions
                    'all' -> run all functions
                    'new_data' -> read in new data, process it, create new faculty and all itemssheets
                    'read_sheets' -> read in faculty sheet data, process, create import sheet
            dirs: dict[str,str] | None
                A dict containing the str path to the directories to use. If not provided, it will use the default dirs from settings.env.
            only_changes: bool = True
                True (default): only add items that have been changed to the created sheets
                False: add all items from the CopyRight export to the created sheets
            other_sheets: list[str] | None
                A list of paths to additional .xlsx files to ingest instead the raw data from CopyRight.
        """

        self.files: dict[FileSetting, File] = SETTINGS.files
        self.faculties: list[str] = []
        # latest copyright export file & when it was created
        self.latest_file_date: str = ""

        self.settings = settings
        self.functions: list[callable] = []
        self.dirs = settings.dirs

        # Initialize data structures
        self.copyright_data = pl.DataFrame()
        self.faculty_sheet_data = pl.DataFrame()
        self.all_items_sheet_data = pl.DataFrame()
        self.import_sheet_data = pl.DataFrame()

        # Initialize other attributes
        self.other_sheet = File(settings.other_sheet) if settings.other_sheet else None
        self.only_changes = settings.only_changes
        self.disable_writes = not settings.save_files
        self.retrieve_all = settings.retrieve_all
        self.refresh_osiris_data = settings.refresh_osiris_data
        self.enrich_with_osiris_data = settings.enrich_with_osiris_data
        self.only_retrieve_missing_osiris_data = settings.only_retrieve_missing_osiris_data
        self.no_new_items = False
        self.style_iter = 2

        # Set functions to run
        self.set_functions(settings.functions)

    def set_functions(self, functions: Functions | None) -> None:
        """
        Sets the functions to run based on the input parameter 'functions'.
        stores it in self.settings as a list of functions to run.
        Returns None.
        """
        if functions is None:
            return

        # Export only
        if functions == Functions.export:
            self.functions = [
                self.read_faculty_sheets,
                self.create_import_sheet,
            ]
            return


        # Common functions
        self.functions.extend([
            self.process_copyright_export,
            # self.read_all_items_sheets, # Disabled because the data is not used anywhere at the moment
            ])



        # Common functions
        self.functions.extend([
            self.create_faculty_sheets,
            self.create_all_items_sheet,
            self.create_overviews,
            ])

        # exclusive for 'Both'
        if functions == Functions.both:
            self.functions.extend([
            self.create_export_sheet
        ])
    def run(self) -> None:
        """
        Runs the functions as specified in the settings dict.
        """
        if self.retrieve_all:
            # will retrieve all data from the directories where users can enter data
            # and store it as a parquet file and csv file in the root dir
            self.retrieve_all_data()
        for func in self.functions:
            info(f'running {func}')
            func()

    def process_copyright_export(self) -> None:
        """
        Process the raw copyright data:
        rename column headers, add extra columns, format some data, and match to faculty.

        If 'only_changes' it will compare this data to the items present in the faculty sheets,
        and only include new items in the export.
        """
        if self.copyright_data.is_empty():
            if self.other_sheet:
                self.latest_file_date, self.copyright_data = read_other_sheet(self.other_sheet)
            else:
                self.latest_file_date, self.copyright_data = read_copyright_export()
            if self.copyright_data.is_empty():
                warn("No new Copyright data found to process! Exiting...")
                raise typer.Exit(code=1)


        # enrich copyright_data with OSIRIS data if bool is set
        if self.enrich_with_osiris_data:
            self.copyright_data = enrich_df_with_osiris_data(self.copyright_data,'full data')
            self.copyright_data = self.clean_and_validate_df(self.copyright_data)


        self.faculties = (
            self.copyright_data.select(pl.col("faculty").unique()).to_series().sort().to_list()
        )
        self.copyright_data = self.clean_and_validate_df(self.copyright_data)
        if self.only_changes:
            self.read_faculty_sheets(include_overview=False)
            if self.faculty_sheet_data.is_empty():
                info(
                    "No faculty sheets found. Adding all items without checking for changes."
                )
            else:
                """
                In this part, all items in self.copyright_data that are not present in self.faculty_sheet_data
                will be added to self.faculty_sheet_data.
                This is done by comparing columns material_id and last_change.
                Items are added to faculty_sheet_data if:
                - material_id is not present in self.faculty_sheet_data
                - material_id is found but last_change date is different
                """
                info(f'number of rows in copyright_data: {self.copyright_data.shape[0]}')
                info(f'number of rows in faculty sheet data: {self.faculty_sheet_data.shape[0]}')

                not_in_faculty = self.copyright_data.join(
                    self.faculty_sheet_data, on="material_id", how="anti"
                )
                info(f'copyright_data rows not found in faculty sheet data: {not_in_faculty.shape[0]}')

                matching_id_diff_change = (
                    self.copyright_data.join(
                        self.faculty_sheet_data, on="material_id", how="inner"
                    )
                    .filter(pl.col("last_change") != pl.col("last_change_right"))
                    .select(pl.all().exclude("last_change_right"))
                    .drop_nulls(pl.col("material_id"))
                    .filter(pl.col("status") == "Deleted")
                )

                if not_in_faculty.is_empty():
                    if matching_id_diff_change.is_empty():
                        info("No new items to add!")
                        self.no_new_items = True
                    else:
                        self.copyright_data = matching_id_diff_change
                if not matching_id_diff_change.is_empty():
                    self.copyright_data = pl.concat(
                        [not_in_faculty, matching_id_diff_change]
                    )
                else:
                    self.copyright_data = not_in_faculty
                self.copyright_data = self.clean_and_validate_df(self.copyright_data)
                self.read_all_items_sheets()

        cool(f'process copyright export done. {self.copyright_data.shape[0]} rows in self.copyright_data.')

    def create_faculty_sheets(self) -> None:
        """
        Splits the processed copyright data into one sheet per faculty
        and exports the result to excel sheets.
        """
        info(f"Exporting new items to faculty sheets for date {self.latest_file_date}")
        if self.faculties:
            self.faculties.sort()
        if not self.copyright_data.is_empty():
            self.copyright_data = enrich_df_with_osiris_data(self.copyright_data)
        for faculty in self.faculties:
            if faculty in COURSE_MAPPING:
                self.create_programme_sheets(faculty)

            faculty_dir = Directory(self.dirs[DirSetting.FACULTIES_DIR].full / faculty)
            if faculty is None or faculty == "":
                faculty = "no_faculty_found"

            filename = f"{faculty}_{self.latest_file_date}.xlsx"
            i = 1
            while os.path.exists(faculty_dir.full / filename):
                filename = f"{faculty}_{self.latest_file_date}_{i}.xlsx"
                i += 1

            faculty_data = self.copyright_data.filter(pl.col("faculty") == faculty)
            gap = " " * (15 - len(faculty))
            if faculty_data.is_empty():
                warn(f"{faculty}:{gap}{faculty_data.shape[0]} (no new items, skipping)")
                continue
            else:
                info(f"{faculty}:{gap}{faculty_data.shape[0]}")
            store_complete_data(faculty_dir.full / filename, faculty_data)
            self.style_iter = finalize_sheet(
                File(str(faculty_dir.full / filename)), faculty_data, self.style_iter
            )

    def create_programme_sheets(self, faculty: str) -> None:
        """
        For a given faculty, split processed copyright data into one sheet per programme.
        Export to faculty_dir / per_programme / programme_name}_{date}.xlsx
        """

        programme_dir = Directory(
            self.dirs[DirSetting.FACULTIES_DIR].full / faculty / "per_programme"
        )
        course_to_sheet: dict[str, str] = COURSE_MAPPING[faculty]
        data: list[dict[str, pl.DataFrame]] = []
        info(f"creating programme sheets for {faculty}")
        for course, group in course_to_sheet.items():
            course_data = self.copyright_data.filter(pl.col("department") == course)
            gap = " " * (40 - len(course))
            if course_data.is_empty():
                warn(f"{course}:{gap}{course_data.shape[0]} (no new items, skipping)")
                continue
            else:
                info(f"retrieved programme sheet data for {course}")
                info(f"{course}:{gap}{course_data.shape[0]}")
                data.append({"sheet": group, "data": course_data})
        final_data: dict[str, pl.DataFrame] = {}
        for item in data:
            if item.get("sheet") in final_data:
                final_data[item.get("sheet")] = pl.concat(
                    [final_data[item.get("sheet")], item["data"]]
                )
            else:
                final_data[item.get("sheet")] = item["data"]

        for groupname, df in final_data.items():
            filename = (
                programme_dir.full / f"{groupname}_{self.latest_file_date}.xlsx"
            )

            store_complete_data(filename, df)
            self.style_iter = finalize_sheet(File(str(filename)), df, self.style_iter)
            info(
                f"created programme sheet {groupname}_{self.latest_file_date}.xlsx"
            )

    def create_all_items_sheet(self) -> None:
        """
        Add all items in the current Copyright data to a single sheet.
        """
        if not self.no_new_items:
            filename = f"all_items_{self.latest_file_date}.xlsx"
            i = 1
            while os.path.exists(self.dirs[DirSetting.ALL_ITEMS_DIR].full / filename):
                filename = f"all_items_{self.latest_file_date}_{i}.xlsx"
                i += 1
            if not self.disable_writes:
                store_complete_data(self.dirs[DirSetting.ALL_ITEMS_DIR].full / filename, self.copyright_data)
                info(f"Created sheet: {self.dirs[DirSetting.ALL_ITEMS_DIR].full / filename}")

    def read_all_items_sheets(self) -> None:
        """
        Reads in all data from all 'all_items' sheets
        and stores it in self.all_items_sheet_data as a single concatted dataframe.

        # NOTE: currently this data is not used anywhere!!
        """

        self.all_items_sheet_data = self.read_complete_data_from_sheets(
            self.dirs[DirSetting.ALL_ITEMS_DIR].files_r
        )
        warn(f'self.all_items_sheet_data has been set, but is not currently used.')

    def read_complete_data_from_sheets(
        self, files: list[File], sheetname: str = SETTINGS.data_settings.complete_data_name
    ) -> pl.DataFrame:
        """
        Reads the data from the Complete data sheet for each file in 'files'.
        """
        file_data = []
        for file in files:
            if file.extension not in [".xls", ".xlsx"]:
                continue
            if "overview" in file.name:
                info(f"skipping {file.path}")
                continue
            try:
                current_data = pl.read_excel(
                    file.path, sheet_name=sheetname, infer_schema_length=None
                )
            except Exception:
                current_data = pl.read_excel(file.path, infer_schema_length=None)

            current_data = self.clean_and_validate_df(current_data)
            if current_data.is_empty():
                continue
            else:
                file_data.append(current_data)
        if file_data:
            result: pl.DataFrame = pl.concat(file_data, how="diagonal_relaxed")
        else:
            result = pl.DataFrame()
        return self.clean_and_validate_df(result.unique())

    def clean_and_validate_df(self, df: pl.DataFrame) -> pl.DataFrame:
        """
        Current implementation is bare:
        - if sheet is not empty:
            - set all columns to type str
            - replace truncated url values

        TODO: Implement this function fully.


        """

        if not df.is_empty():
            # set all columns to type str
            df = df.with_columns(pl.exclude(pl.Utf8).cast(str))

            # replace truncated url values
            if 'url' in df.columns:
                df = df.with_columns(pl.col('url').str.replace(r'\.{3}','https://utwente.instructure.com/files'))
            if 'osiris_catalogue_url' in df.columns:
                df = df.with_columns(pl.col('osiris_catalogue_url').str.replace(r'\.{3}','https://utwente.instructure.com/files'))
        return df

    def create_import_sheet(self) -> None:
        """
        combine self.faculty_sheet_data and self.all_items_sheet_data
        clean it up
        change from UT Easy Access format to SURF CopyRight format
        create & export an .xlsx sheet that can be sent to SURF to be imported into CopyRight.
        """
        # TODO
        ...

    def read_faculty_sheets(self, include_overview: bool = True) -> None:
        """
        Reads in all data from all sheets in the faculties dir
        and stores it in self.faculty_sheet_data as a single concatted dataframe.
        """
        self.faculty_sheet_data = self.get_all_faculty_data(
            include_overview=include_overview
        )

    def get_all_faculty_data(self, include_overview: bool = True) -> pl.DataFrame:
        """
        Read in all available faculty sheets
        and merge the 'complete data' and 'data entry' sheets for each one.
        concat all the data into a single dataframe and return it.
        """
        all_faculty_data = pl.DataFrame()
        for faculty in self.faculties:
            info(f"getting data for faculty {faculty}")
            faculty_data = self.get_faculty_data(
                faculty, include_overview=include_overview
            )
            if faculty_data.is_empty():
                continue
            all_faculty_data = pl.concat(
                [all_faculty_data, faculty_data], how="diagonal_relaxed"
            )

        final_data =  self.clean_and_validate_df(all_faculty_data.unique())
        info(f'data from all faculties has {final_data.shape[0]} unique rows')
        return final_data

    def get_faculty_data(
        self, faculty: str, del_overview: bool = False, include_overview: bool = True
    ) -> pl.DataFrame:
        """
        TODO: Change merge algorithm.Now it's on sheet-by-sheet basis. Change to a item-by-item comparison.
        for a given faculty, read in all available faculty sheets
        and merge the 'complete data' and 'data entry' sheets for each one.
        concat all the data into a single dataframe and return it.

        Parameters:
            faculty: str
                the faculty to get the data for. Will scan through all sheets in path self.dirs['faculties'].full / faculty.
            del_overview: bool
                if True, delete the existing overview sheets for this faculty.
        """

        def join_coalesce_all(
            df1: pl.DataFrame, df2: pl.DataFrame, on: str, prefer_right=set()
        ) -> pl.DataFrame:
            to_coalesce = set(df1.columns) & set(df2.columns) - set([on])
            coalesced = {
                c: pl.coalesce(pl.col(c + "_right"), pl.col(c))
                if c in prefer_right
                else pl.coalesce(pl.col(c), pl.col(c + "_right"))
                for c in to_coalesce
            }
            return (
                df1.join(df2, on=on, how="full", suffix="_right")
                .with_columns(**coalesced)
                .drop([c + "_right" for c in to_coalesce])
                .drop(["material_id_right"])
            )

        if faculty is None or faculty == "":
            return pl.DataFrame()

        faculty_dir = Directory(self.dirs[DirSetting.FACULTIES_DIR].full / faculty)
        faculty_files = faculty_dir.files_r

        all_faculty_data = pl.DataFrame()

        prefer_overview_cols = False  # set to True to give edits in total_overview file higher priority than edits in each weekly faculty excel

        total_overview = pl.DataFrame()
        overview_file: File = None
        latest_mod_date = None
        if include_overview:
            for file in faculty_files:
                if "total_overview" in file.name and faculty in file.name:
                    try:
                        total_overview_complete = pl.read_excel(
                            file.path, sheet_name="Complete data"
                        )
                        total_overview_data_entry = pl.read_excel(
                            file.path, sheet_name="Data entry"
                        )
                        total_overview = join_coalesce_all(
                            total_overview_complete,
                            total_overview_data_entry,
                            on="material_id",
                            prefer_right=set(total_overview_data_entry.columns)
                            - {"material_id"},
                        )
                    except ValueError:
                        total_overview = pl.read_excel(file.path)
                    overview_file = file
                    logger.debug(f'{faculty} data retrieval:  {total_overview.shape[0]} rows in overview file')
        for file in faculty_files:
            if file.extension not in [".xls", ".xlsx"]:
                continue
            elif "overview" in file.name or 'llm_classification' in file.name:
                continue
            else:
                latest_mod_date = (
                    file.modified
                    if latest_mod_date is None
                    else max(latest_mod_date, file.modified)
                )
                full_data = pl.read_excel(file.path, sheet_name=SETTINGS.data_settings.complete_data_name)
                data_entry = pl.read_excel(file.path, sheet_name=SETTINGS.data_settings.data_entry_name)

                full_data = self.clean_and_validate_df(full_data)
                data_entry = self.clean_and_validate_df(data_entry)

                # merge data_entry into full_data on column material_id.
                # data from data_entry will overwrite data from full_data
                # if a col is present in data_entry, but not in full_data, it will be added
                # keep the columns in full_data that are not in data_entry

                merged_data = join_coalesce_all(
                    full_data,
                    data_entry,
                    on="material_id",
                    prefer_right=set(data_entry.columns) - {"material_id"},
                )
                merged_data = merged_data.unique(subset="material_id")
                all_faculty_data = pl.concat(
                    [all_faculty_data, merged_data], how="diagonal_relaxed"
                )
                all_faculty_data = all_faculty_data.unique(subset="material_id")

        if not total_overview.is_empty():
            if latest_mod_date:
                if (latest_mod_date < overview_file.modified) and (
                    not prefer_overview_cols
                ):
                    all_overview_man_class = (
                        total_overview.select(pl.col("manual_classification"))
                        .to_series()
                        .to_list()
                    )
                    all_overview_man_class = [
                        i for i in all_overview_man_class if i not in [None, "", "-", " "]
                    ]
                    all_faculty_data_man_class = (
                        all_faculty_data.select(pl.col("manual_classification"))
                        .to_series()
                        .to_list()
                    )
                    all_faculty_data_man_class = [
                        i
                        for i in all_faculty_data_man_class
                        if i not in [None, "", "-", " "]
                    ]
                    if len(all_faculty_data_man_class) < len(all_overview_man_class):
                        warn(
                            f"all_faculty_data has less manual classifications than total_overview. Will prefer overview columns for {faculty}."
                        )
                        prefer_overview_cols = True
            if prefer_overview_cols:
                preffered_cols = set(total_overview.columns) - {"material_id"}
                all_faculty_data = join_coalesce_all(
                    all_faculty_data,
                    total_overview,
                    on="material_id",
                    prefer_right=preffered_cols,
                )
            else:
                preffered_cols = set(all_faculty_data.columns) - {"material_id"}
                all_faculty_data = join_coalesce_all(
                    total_overview,
                    all_faculty_data,
                    on="material_id",
                    prefer_right=preffered_cols,
                )
            all_faculty_data = all_faculty_data.unique(subset="material_id")

        overview_fac_dir = Directory(self.dirs[DirSetting.OVERVIEWS_BACKUP].full / faculty)
        overview_file_name = None
        if del_overview:
            if overview_file:
                overview_file_name = overview_file.name
                if SETTINGS.backup_settings.backup_overviews:
                    overview_file.move(overview_fac_dir.full / overview_file.name)
                else:
                    overview_file.delete()

            for file in faculty_files:
                if file.name == overview_file_name:
                    continue
                if "total_overview" in file.name and faculty in file.name:
                    if SETTINGS.backup_settings.backup_overviews and "llm" not in file.name:
                        file.move(overview_fac_dir.full / file.name)
                    else:
                        file.delete()

        return self.clean_and_validate_df(all_faculty_data)

    def create_overviews(self) -> None:
        """
        1. Creates overview sheets with data per faculty (and per programme if found in COURSE_MAPPING)
        2. Creates tables with summary data and prints them to the console + stores them as .html files

        No parameters, will pull the data from disk for each faculty in self.faculties.
        """
        faculty_dict: dict[str, pl.DataFrame] = {}
        if not self.faculties:
            self.process_copyright_export()
            if not self.faculties:
                warn(f'No faculties detected in current data. Cannot produce overviews.')
                return

        # store the full current data used to create the overviews to self.import_data as well
        self.import_sheet_data = pl.DataFrame()
        for faculty in self.faculties:
            if not faculty or faculty == "" or faculty == "Unmapped":
                continue
            data = self.get_faculty_data(faculty, del_overview=True)
            if data.is_empty():
                continue
            enrich_df_with_osiris_data(data, faculty)
            self.import_sheet_data = pl.concat([self.import_sheet_data, data], how="diagonal_relaxed")
            faculty_dict[faculty] = data
        self.style_iter = create_faculty_overviews(faculty_dict, self.style_iter)


    def retrieve_all_data(self) -> pl.DataFrame:
        """
        Goes through all files to retrieve all available data.
        Then, for each material_id, grab only unique rows.
        Keep track of where the data came from.

        Returns a dataframe with all unique rows including provenance.
        """
        found_dfs = dict()
        today = datetime.now().strftime("%Y-%m-%d")
        cool("Retrieving all data. Please wait, this can take a while.")
        numfiles = 0
        dirs = {
            "faculties": self.dirs.get(DirSetting.FACULTIES_DIR),
        }
        data_entry_info = defaultdict(list)
        for name, dir in dirs.items():
            cur_df = pl.DataFrame()
            for file in dir.files_r:
                if file.extension not in [".xls", ".xlsx", ".csv"]:
                    continue
                if file.extension in [".xls", ".xlsx"]:
                    try:
                        file_content:dict[str,pl.DataFrame|list[pl.DataFrame]] = {'other':list()}
                        try:
                            file_content[SETTINGS.data_settings.complete_data_name] = pl.read_excel(file.path, sheet_name=SETTINGS.data_settings.complete_data_name)
                        except Exception:
                            ...
                        try:
                            file_content[SETTINGS.data_settings.data_entry_name] = pl.read_excel(file.path, sheet_name=SETTINGS.data_settings.data_entry_name)
                        except Exception:
                            ...
                        if SETTINGS.data_settings.complete_data_name not in file_content:
                            try:
                                file_content['other'].append(pl.read_excel(file.path, sheet_id=1))
                            except Exception:
                                ...
                        if SETTINGS.data_settings.data_entry_name not in file_content:
                            try:
                                file_content['other'].append(pl.read_excel(file.path, sheet_id=2))
                            except Exception:
                                ...
                        numfiles += 1
                    except Exception as e:
                        print(f"Couldnt read file {file.path}: {e}")
                        continue


                    for sheetname, dataframe in file_content.items():
                        if isinstance(dataframe, pl.DataFrame):
                            if dataframe.is_empty():
                                continue
                        if sheetname is not SETTINGS.data_settings.data_entry_name:
                            if not isinstance(dataframe, list):
                                dataframe = [dataframe]
                            elif len(dataframe) == 0:
                                continue
                            for df in dataframe:
                                if df.is_empty():
                                    continue
                                df = df.with_columns(
                                    pl.lit(str(file.name)).alias("from_file")
                                )
                                if cur_df.is_empty():
                                    cur_df = df
                                    continue
                                cur_df = pl.concat([cur_df, df], how="diagonal_relaxed")
                        else:
                            df_as_dict = dataframe.to_dicts()
                            for row in df_as_dict:
                                if row.get("material_id"):
                                    data_entry_info[row.get("material_id")].append(
                                        {
                                            "from_file": file.name,
                                            "manual_classification": row.get("manual_classification"),
                                            "remarks": row.get("remarks"),
                                            "workflow_status": row.get("workflow_status"),
                                        }
                                    )


            info(f"Retrieved {cur_df.shape[0]} rows from dir {name}")
            cur_df = cur_df.unique()
            info(f"{cur_df.shape[0]} remaining after removing duplicate rows")
            if "material_id" in cur_df.columns:
                cur_df_mat_ids = cur_df.select("material_id").to_series().to_list()
                unique_cur_df_mat_ids = set(cur_df_mat_ids)
                info(
                    f"found {len(cur_df_mat_ids)} rows in dir {name} for {len(unique_cur_df_mat_ids)} unique material ids."
                )
            found_dfs[name] = cur_df

        info(f"Done retrieving data from {numfiles} files. Now merging all.")
        full_df = pl.DataFrame()
        for df in found_dfs.values():
            if full_df.is_empty():
                full_df = df
                continue
            full_df = pl.concat([full_df, df], how="diagonal_relaxed")

        info(f"after concatting all dfs, full_df has {full_df.shape[0]} rows")
        select_cols = ["from_file"]
        if "material_id" in full_df.columns:
            select_cols.append("material_id")
        if "Material id" in full_df.columns:
            select_cols.append("Material id")
        if "manual_classification" in full_df.columns:
            select_cols.append("manual_classification")
        if "remarks" in full_df.columns:
            select_cols.append("remarks")
        if "workflow_status" in full_df.columns:
            select_cols.append("workflow_status")
        df_subset = full_df.select(select_cols).to_dicts()

        files_final_dict: dict[int, list] = dict()
        man_class_final_dict: dict[str, str] = dict()
        remarks_final_dict: dict[str, str] = dict()
        workflow_status_final_dict: dict[str, str] = dict()

        for row in df_subset:
            if row.get("material_id"):
                mat_id = int(row.get("material_id"))
            elif row.get("Material id"):
                mat_id = int(row.get("Material id"))
            if not mat_id:
                continue
            if mat_id in files_final_dict:
                if row.get("from_file") not in files_final_dict.get(mat_id):
                    files_final_dict[mat_id].append(row.get("from_file"))
            else:
                files_final_dict[mat_id] = list()
                files_final_dict[mat_id].append(row.get("from_file"))
            mat_id = str(mat_id)
            if mat_id  in data_entry_info:
                for entry in data_entry_info.get(mat_id):
                    man_class = entry.get("manual_classification")
                    remark = entry.get("remarks")
                    workflow_status = entry.get("workflow_status")
                    if man_class:
                        if man_class != "-" and man_class != row.get("manual_classification"):
                            man_class_final_dict[mat_id] = man_class
                    if remark:
                        if remark != "-" and remark != row.get("remarks"):
                            remarks_final_dict[mat_id] = remark
                    if workflow_status:
                        if workflow_status != row.get("workflow_status") and workflow_status != "ToDo":
                            workflow_status_final_dict[mat_id] = workflow_status

        from_file_update_dict: dict[str, str] = dict()
        for material_id, filenames in files_final_dict.items():
            if len(filenames) > 1:
                from_file_update_dict[str(material_id)] = ", ".join(filenames)
            else:
                from_file_update_dict[str(material_id)] = filenames[0]

        full_df = full_df.drop("from_file")

        full_df = full_df.unique(
            subset=[
                "material_id",
                "manual_classification",
                "remarks",
                "workflow_status",
            ]
        )
        full_df = full_df.drop(["manual_classification", "remarks", "workflow_status"])
        full_df = full_df.with_columns(
            [
                pl.col("material_id").replace(from_file_update_dict).alias("from_file"),
                pl.col("material_id").replace(man_class_final_dict).alias("manual_classification"),
                pl.col("material_id").replace(remarks_final_dict).alias("remarks"),
                pl.col("material_id").replace(workflow_status_final_dict).alias("workflow_status"),
                pl.lit(today).alias("last_sheet_update"),
            ]
        )

        def merge_rows(df: pl.DataFrame, unique_col: str) -> pl.DataFrame:
            # Cast all columns to string and preprocess
            df = df.with_columns(pl.exclude(pl.Utf8).cast(str))

            # Replace empty-like values with null
            preprocessed_exprs = [
                pl.when(
                    pl.col(col).is_null() |
                    (pl.col(col) == "") |
                    (pl.col(col) == "-")
                )
                .then(None)
                .otherwise(pl.col(col))
                .alias(col)
                for col in df.columns
            ]
            df_preprocessed = df.select(preprocessed_exprs)

            # Generate aggregation expressions
            agg_exprs = []
            for col in df_preprocessed.columns:
                if col == unique_col:
                    continue

                # Build single expression with null handling
                expr = (
                    pl.when(pl.col(col).is_null().all())
                    .then(None)  # All null case
                    .when(pl.col(col).drop_nulls().unique().len() == 1)
                    .then(pl.col(col).drop_nulls().unique().first())  # Single unique value
                    .otherwise(pl.col(col).drop_nulls().first())  # Multiple unique values
                    .alias(col)
                )
                agg_exprs.append(expr)

            # Group and aggregate with proper null handling
            return df_preprocessed.group_by(unique_col).agg(agg_exprs)



        #full_df = merge_rows(full_df, unique_col="material_id")
        info(
            f"{full_df.shape[0]} rows remaining after selecting unique rows based on material_id, manual classification, remarks, and workflow_status."
        )
        info("Now comparing data with previously stored items.")
        df_merged = pl.DataFrame()
        try:
            stored_df = pl.read_parquet(SETTINGS.files.get(FileSetting.FULL_DATA_PARQUET).path)
        except Exception as e:
            warn(f"Error reading full_df.parquet: {e}.")
            df_merged = full_df

        if df_merged.is_empty():
            if "last_sheet_update" not in stored_df.columns:
                info(
                    f"stored_df has no last_sheet_update data. Overwriting stored data with new data"
                )
                df_merged = full_df
            else:
                compare_cols = [
                    "manual_classification",
                    "remarks",
                    "workflow_status",
                    "retrieved_from_copyright_on",
                    "last_change",
                    "status",
                ]

                # Compare stored_df and full_df.
                # stored_df is the data currently on disk in full_df.parquet ('old'), full_df is the data we just retrieved ('new')
                # Compare rows with the same material_id.
                #

                df_merged = (
                    full_df.join(
                        stored_df, on="material_id", how="left", suffix="_stored"
                    )
                    .with_columns(
                        [
                            pl.fold(
                                True,
                                lambda acc, x: acc & x,
                                [
                                    (pl.col(c) == pl.col(f"{c}_stored"))
                                    for c in compare_cols
                                ],
                            ).alias("all_match")
                        ]
                    )
                    .with_columns(
                        [
                            pl.fold(
                                False,
                                lambda acc, x: acc | x,
                                [
                                    (
                                        pl.col(c).is_null()
                                        & pl.col(f"{c}_stored").is_not_null()
                                    )
                                    for c in compare_cols
                                ],
                            ).alias("any_missing_in_full_df")
                        ]
                    )
                    .with_columns(
                        [
                            pl.when(pl.col("any_missing_in_full_df"))
                            .then(pl.col(f"{c}_stored"))
                            .otherwise(pl.col(c))
                            .alias(c)
                            for c in compare_cols
                        ]
                        + [
                            pl.when(pl.col("any_missing_in_full_df"))
                            .then(pl.col("last_sheet_update_stored"))
                            .otherwise(
                                pl.when(pl.col("all_match"))
                                .then(pl.col("last_sheet_update_stored"))
                                .otherwise(pl.col("last_sheet_update"))
                            )
                            .alias("last_sheet_update")
                        ]
                    )
                )
                dropcols = [c for c in df_merged.columns if "_stored" in c]
                df_merged = df_merged.drop(dropcols).drop(
                    ["all_match", "any_missing_in_full_df"]
                )
                df_merged = df_merged.unique('material_id')

        # refresh OSIRIS data if bool is set
        if self.refresh_osiris_data:
            info("Refreshing OSIRIS data. This will take a while!")
            asyncio.run(update_osiris_data(df_merged, self.only_retrieve_missing_osiris_data))
        self.clean_and_validate_df(df_merged)
        df_merged.write_parquet(SETTINGS.files.get(FileSetting.FULL_DATA_PARQUET).path)
        df_merged.write_csv(SETTINGS.files.get(FileSetting.FULL_DATA_CSV).path)

    def create_export_sheet(self) -> None:
        material_ids: list[str] = create_export_sheet(data=self.import_sheet_data)
        if material_ids:
            info(f"Updating export status for {len(material_ids)} material ids.")
            warn(f'NOT YET IMPLEMENTED')
            item_data: pl.DataFrame = self.import_sheet_data.filter(pl.col(name='material_id').is_in(other=material_ids))

            # group by faculty
            item_data_per_faculty: dict[str, pl.DataFrame] = {faculty: item_data.filter(pl.col(name='faculty') == faculty) for faculty in item_data.select(pl.col(name='faculty')).unique().to_series().to_list()}
            for faculty, data in item_data_per_faculty.items():
                # for each file in the faculty dir, open it
                # read the data
                # if the material_id is in the data, add col 'exported_on' with the current date (YYYY-MM-DD)
                # save the file
                info(f'Updating export status for {faculty}')
                print(data)
                warn(f'NOT YET IMPLEMENTED')

            # once done, run 'update_export_sheets' to create/update the export sheets for each faculty
            # these sheets show all the exported item for that faculty
            # items in these sheets should be removed from the overview sheets of that faculty
