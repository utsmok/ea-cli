import time
from functools import partial
from pathlib import Path

import aiometer
import httpx
import polars as pl
from loguru import logger


async def check_file_exists(
    api_token: str,
    df: pl.DataFrame | None = None,
    excel_file: str | Path | None = None,
    sheetname: str = "Complete data",
    url_col: str = "url",
) -> pl.DataFrame:
    """
    Arguments:
        - api_token: The API token to use for authentication (Canvas API).
        - df: The DataFrame containing the data to process (if not provided, will read from excel_file).
        - excel_file: The path to the Excel file to read. (will not be used if df is provided)
        - sheetname: The name of the sheet to read from/write to (default is "Complete data").
        - url_col: The name of the column containing the URLs (default is "url").
    Returns:


    This function does the following:
    Step 1: If no df is given, read in data from the given excel file
    Step 2: Check if the 'url' column exists, otherwise raise an error
    Step 3: For each row, extract the file_id
    Step 4: For each row, use this file_id to construct the url
    Step 5: For each row, check if the file exists and store result as bool in `file_exists` col
    Step 6: Store the results as the same file again, only now with the new column(s) added. Drop unnecessary columns.
    """

    async def check_file_exists(
        data: dict[str, str | bool], session: httpx.AsyncClient
    ) -> dict[str, str | bool]:
        """Checks if a Canvas file exists via the API"""
        file_url = data.get("file_url")
        if not file_url or not isinstance(file_url, str):
            return {"material_id": data["material_id"], "file_exists": False}
        try:
            response = await session.get(file_url)
            data["file_exists"] = response.status_code == 200
            return data
        except Exception as e:
            logger.error(f"Error checking file existence for {file_url}: {e}")
            return {"material_id": data["material_id"], "file_exists": False}

    if not isinstance(df, pl.DataFrame):
        if excel_file:
            df = pl.read_excel(excel_file, sheet_name=sheetname)
        else:
            raise ValueError("Either df or excel_file must be provided.")

    if url_col not in df.columns:
        raise ValueError(f"Excel must contain a column named '{url_col}'")

    # Extract urls
    files_to_check = (
        df.with_columns(
            pl.col(url_col).str.extract(r"/files/([^/?]+)\?", 1).alias("file_id")
        )
        .with_columns(
            pl.when(pl.col("file_id").is_not_null())
            .then(
                pl.concat_str(
                    [
                        pl.lit("https://utwente.instructure.com/api/v1/files/"),
                        pl.col("file_id"),
                    ]
                )
            )
            .otherwise(None)
            .alias("file_url")
        )
        .select("material_id", "file_url")
        .to_dicts()
    )

    # Use API to check file existence
    header = {"Authorization": f"Bearer {api_token}"}
    processed_results = []
    counterdict = {"true": 0, "false": 0, "none": 0}
    async with httpx.AsyncClient(
        headers=header, follow_redirects=True, timeout=20
    ) as session:
        logger.info(f"Checking {len(files_to_check)} URLs")
        start_time = time.time()
        # do a batchwise check
        counter = 0
        async with aiometer.amap(
            partial(check_file_exists, session=session),
            files_to_check,
            max_at_once=100,
            max_per_second=200,
        ) as temp_results:
            async for result in temp_results:
                if not isinstance(result, BaseException):
                    processed_results.append(result)
                    if result.get("file_exists") is True:
                        counterdict["true"] += 1
                    elif result.get("file_exists") is False:
                        counterdict["false"] += 1
                    else:
                        counterdict["none"] += 1
                else:
                    counterdict["none"] += 1
    logger.info(
        f"Received data for {len(processed_results)} URLs after checking {len(files_to_check)} items in {time.time() - start_time:.2f} seconds."
    )
    logger.info(f"File existence check results: {counterdict}")
    results = pl.from_dicts(processed_results)
    results.write_excel("file_existence_check_results.xlsx")
    df = df.drop("file_exists")
    df = df.join(results, on="material_id", how="left")

    if excel_file:
        df.write_excel(excel_file, worksheet=sheetname)
        logger.info(f"Results written back to {excel_file}.")

    logger.debug(f"Checked {df.shape[0]} items for file existence on Canvas.")
    return df
