def read_faculty_sheets(settings: Settings) -> pl.DataFrame:
    """
    Reads all faculty sheets and returns a single DataFrame.
    """
    all_dfs = []
    for faculty_dir in settings.dirs[DirSetting.FACULTIES_DIR].dirs():
        for file in faculty_dir.files_r:
            if file.extension == ".xlsx" and "overview" not in file.name:
                try:
                    df = _read_excel_quiet(
                        file.path, sheet_name=settings.data_settings.data_entry_name
                    )
                    all_dfs.append(df)
                except Exception as e:
                    logger.warning(f"Error reading {file.path}: {e}")
                    continue
    if not all_dfs:
        return pl.DataFrame()
    return pl.concat(all_dfs)