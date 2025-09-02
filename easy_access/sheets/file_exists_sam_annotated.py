# import pandas as pd
# import requests
# from time import sleep # we don't need this when using httpx

from pathlib import Path

import httpx  # switched to httpx as it's async capable and already used in this code
import polars as pl  # moved to polars as we're using that everywhere, it's faster and easier to use
from loguru import (
    logger,
)  # imported the logger used throughout the code to log messages instead of using `print`

# first, we will turn the main part of this file into a function so we can call it from other files


# we will add an argument for each of the parameters/variables, instead of hard-coding them
# we will change the vars to lower case by convention, as these are no longer constants
def check_file_exists(
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
    Step 6: if excel_file was given:
            Store the results as the same file with the `file_exists` col.
    Step 7: return the modified DataFrame.

    """
    '''
    changed extract_file_id to a simpler function, see below
    we won't use that one either, as I use  a polars function directly instead
    so this is just for reference :)
    # ORIGINAL CODE
    def extract_file_id(url):
        """Extracts the file ID from a Canvas file URL"""
        if not url or (isinstance(url, tuple) and not url[0]):  # handle empty/tuple cases
            return None
        match = re.search(r'/(\d+)(?:\D*$|$)', str(url))
        return match.group(1) if match else None
    '''

    def extract_file_id(url: str) -> str | None:
        """Helper function to extract the file id from an URL."""
        if "/files/" not in url:
            logger.warning(f"Invalid URL format: {url}. Expected '/files/'.")
            return None
        try:
            return url.split("files/")[1].split("?")[0].strip("/")
        except IndexError as e:
            logger.error(
                f"Could not extract material_id from URL {url}, with error message {e}"
            )
            return None

    # Modified check function: added type hints, settings for sessions move to initialization of client (see later)
    # also: input is no longer the file_id but a direct url we pre-constructed
    def check_file_exists(file_url: str, session: httpx.Client) -> bool:
        """Checks if a Canvas file exists via the API"""
        if not file_url:
            logger.warning(f"Invalid or missing file_url: {file_url}. Returning False.")
            return False
        try:
            response = session.get(file_url)
            logger.debug(
                f"Checked file_url {file_url}, got status code {response.status_code}, returning {response.status_code == 200}."
            )
            return response.status_code == 200
        except httpx.HTTPError as e:
            logger.error(f"Error checking file existence for {file_url}: {e}")
            return False

    # Actual code starts here
    # step 1
    # same code for pandas as polars; but we added the sheet_name to ensure we grab the correct sheet
    # because we don't want the data entry sheet

    # if no df is provided, read in data from excel
    if not isinstance(df, pl.DataFrame):
        if excel_file:
            df = pl.read_excel(excel_file, sheet_name=sheetname)
        else:
            raise ValueError("Either df or excel_file must be provided.")

    # step 2
    # immediately check if we actually have urls so we can fail before doing work
    if url_col not in df.columns:
        raise ValueError(f"Excel must contain a column named '{url_col}'")

    # step 3
    # Extract the file_ids before doing the checks
    df = df.with_columns(  # we use `with_columns` to create a new column from an input
        # this lets us run a series of expressions/functions on each row at the same time
        pl.col(url_col)  # we use the `url` column as input
        .str.extract(
            r"/files/([^/?]+)\?", 1
        )  # we apply our regex and select the second hit
        .alias("file_id")  # we store the result in column `file_id`
    )

    # now we should have a 'file_id' col with the extracted file IDs
    logger.info(
        f"Extracted file IDs from {excel_file}. First 10 file_ids: {df['file_id'].head(10).to_list()}"
    )

    # step 4
    # Construct the urls directly and store as a column
    df = df.with_columns(  # create a new col
        pl.when(pl.col("file_id").is_not_null())  # if we have a file_id for a row
        .then(
            pl.concat_str(
                [  # then join the file_id with the canvas api url
                    pl.lit("https://utwente.instructure.com/api/v1/files/"),
                    pl.col("file_id"),
                ]
            )
        )
        .otherwise(None)  # if not return none
        .alias("file_url")  # store as `file_url`
    )

    # We won't use this code:
    # ORIGINAL CODE
    # Start session with API token
    # session = requests.Session()
    # session.headers.update(HEADERS)

    # instead we use a `with` statement for starting a httpx session
    # this ensures the session is closed automatically at the end of the function
    # we also initialize all settings directly here instead when we call the `get`

    header = {"Authorization": f"Bearer {api_token}"}
    with httpx.Client(headers=header, follow_redirects=True, timeout=10) as session:
        logger.info(f"Checking {df.shape[0]} URLs from {excel_file}.")

        # step 5

        df = df.with_columns(  # again we use with_columns to create a new column
            pl.col("file_url")
            .map_elements(  # we use map_elements to apply a function to each element, using `file_url` as input
                lambda file_url: check_file_exists(
                    file_url, session
                ),  # we pass the file_url and session to the function
                return_dtype=pl.Boolean,  # we expect a boolean return type
            )
            .alias("file_exists")  # store as `file_exists`
        )
        # note: this is not ideal, we should probably use async for this to speed things up, but fine for now
        # another note: this will automatically skip over rows with empty file_urls
        # so those `file_exists` value will not be set to false/true but instead to null (empty)

        # you don't want to loop over dataframes, very slow and inefficient
        # when possible use dataframe-native expressions, or apply functions, etc.
        # as shown above
        # ORIGINAL CODE
        # results = []
        # for i, url in enumerate(df["url"], start=1):
        #    file_id = extract_file_id(url)
        #    exists = check_file_exists(file_id, session)
        #    results.append(exists)
        #    print(f"{i}/{len(df)} | {url} -> {exists}")
        #    sleep(0.1)  # avoid hammering the server too fast

    # we won't use this code
    # ORIGINAL CODE
    # df["file_exists"] = results
    # df.to_excel(OUTPUT_FILE, index=False)
    # print(f"\nResults saved to {OUTPUT_FILE}")

    # step 6
    # drop columns we don't want in the final sheet
    df = df.drop(["file_url", "file_id"])
    # fill empty cells in `file_exists` with False
    df = df.with_columns(
        pl.col("file_exists").cast(pl.Boolean)  # first we ensure the coltype is correct
    ).with_columns(
        pl.col("file_exists").fill_null(False)  # then we fill empty cells
    )
    if excel_file:
        # will only write back data if an excel_file is provided at the start
        # of course using polars + correct sheetname
        df.write_excel(excel_file, worksheet=sheetname)
        logger.info(f"Results written back to {excel_file}.")

    # step 7
    return df
