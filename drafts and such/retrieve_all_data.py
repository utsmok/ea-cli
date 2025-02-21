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
