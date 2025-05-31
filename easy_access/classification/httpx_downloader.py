"""
This module provides utilities for downloading files, primarily from a Canvas LMS,
using asynchronous HTTP requests with `httpx`. It includes features for:
- Loading cookies from a JSON file for authenticated sessions.
- Downloading individual files with error handling and progress information.
- Batch downloading multiple files concurrently using a semaphore for rate limiting.
- Helper functions to identify files to download based on database records and
  previously downloaded files.
- A utility to rename downloaded files from Canvas's internal ID to the application's
  `material_id`.
"""

import asyncio
import json
import logging # Added
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple, Optional # Added Optional

import httpx
import polars as pl
# from rich import print # Removed, using logging

from easy_access.db.base import CopyrightItem, init as init_tortoise_orm # init renamed
from easy_access.db.models import Status # For filtering
from easy_access.settings import SETTINGS, DirSetting
# from easy_access.utils import cool, info # Removed, using logging

logger = logging.getLogger(__name__)

# --- Configuration Candidates (Consider moving to settings.yaml) ---
DEFAULT_HTTP_TIMEOUT: float = 30.0
DEFAULT_MAX_CONCURRENT_DOWNLOADS: int = 10
# The following were hardcoded in HttpxDownloader.download_file:
# CANVAS_BASE_DOMAIN: str = "https://utwente.instructure.com"
# CANVAS_USER_ID_FOR_DOWNLOAD: str = "66733" # This is highly specific and problematic
# COOKIE_FILENAME: str = "cookies.secret" # Currently hardcoded in load_cookies_from_file
# --- End Configuration Candidates ---


def load_cookies_from_file() -> httpx.Cookies:
    """
    Loads cookies from a JSON file named 'cookies.secret' located in the SCRIPT_DATA directory.
    The JSON file should be a list of cookie objects, compatible with browser cookie formats
    (e.g., from a browser extension like "Get cookies.txt LOCALLY").

    Each cookie object should at least have "name" and "value". "domain" and "path" are optional.

    Returns:
        httpx.Cookies: An httpx.Cookies object populated with the loaded cookies.

    Raises:
        FileNotFoundError: If the cookie file is not found.
        ValueError: If the cookie file is not a valid JSON list or contains invalid cookie objects.
        json.JSONDecodeError: If the cookie file is not valid JSON.
    """
    # Consider making COOKIE_FILENAME a setting
    cookie_file_path: Path = SETTINGS.dirs[DirSetting.SCRIPT_DATA].full / "cookies.secret"
    logger.info(f"Loading cookies from JSON file: {cookie_file_path}")

    cookies = httpx.Cookies()
    if not cookie_file_path.exists():
        logger.error(f"Cookie file not found: {cookie_file_path}")
        raise FileNotFoundError(f"Cookie file not found: {cookie_file_path}")

    try:
        with open(cookie_file_path, "r", encoding="utf-8") as f: # Specify read mode 'r'
            cookie_data: List[Dict[str, Any]] = json.load(f)

        if not isinstance(cookie_data, list):
            raise ValueError("Cookie file content is not a JSON list.")

        for cookie_obj in cookie_data:
            if isinstance(cookie_obj, dict) and "name" in cookie_obj and "value" in cookie_obj:
                name: str = cookie_obj["name"]
                value: str = cookie_obj["value"]
                domain: Optional[str] = cookie_obj.get("domain")
                path: str = cookie_obj.get("path", "/")
                cookies.set(name, value, domain=domain, path=path)
            else:
                logger.warning(f"Skipping invalid cookie object: {cookie_obj}")

        logger.info(f"Loaded {len(cookies)} cookies from {cookie_file_path.name}.")
        return cookies
    except json.JSONDecodeError as e:
        logger.error(f"Failed to decode JSON from cookie file {cookie_file_path.name}: {e}")
        raise
    except Exception as e: # Catch other potential errors like file read issues
        logger.error(f"Failed to load or parse cookie file '{cookie_file_path.name}': {e}")
        raise


class HttpxDownloader:
    """
    A downloader class using httpx for asynchronous file downloads,
    handling cookies and basic browser-like headers.
    """
    def __init__(self, timeout: float = DEFAULT_HTTP_TIMEOUT) -> None:
        """
        Initializes the HttpxDownloader.

        Args:
            timeout (float, optional): Default timeout for HTTP requests in seconds.
                                       Defaults to DEFAULT_HTTP_TIMEOUT.
        """
        self.download_dir: Directory = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]
        self.cookies: httpx.Cookies = load_cookies_from_file()

        headers: Dict[str, str] = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        self.client: httpx.AsyncClient = httpx.AsyncClient(
            cookies=self.cookies, headers=headers, follow_redirects=True, timeout=timeout
        )
        logger.info(f"HttpxDownloader client initialized with timeout {timeout}s and loaded cookies.")

    async def close_client(self) -> None:
        """Closes the persistent httpx client. Should be called when done."""
        if hasattr(self, 'client') and self.client and not self.client.is_closed:
            await self.client.aclose()
            logger.info("HttpxDownloader client closed.")

    async def download_file(self, initial_url: str) -> Tuple[bool, Optional[Path], Optional[str]]:
        """
        Downloads a file, expecting a Canvas LMS URL structure.
        The actual download URL is constructed based on an assumed pattern.

        Args:
            initial_url (str): The initial Canvas URL (e.g., a file page URL).

        Returns:
            Tuple[bool, Optional[Path], Optional[str]]: A tuple containing:
                - success (bool): True if download was successful, False otherwise.
                - downloaded_filepath (Optional[Path]): Path to the downloaded file if successful.
                - error_message (Optional[str]): Error message if download failed.
        """
        if "/files/" not in initial_url:
            error_msg = f"Invalid URL format: {initial_url}. Expected '/files/' part."
            logger.warning(error_msg)
            return False, None, error_msg

        try:
            # Extracts 'material_id_from_url' which is Canvas's internal file ID.
            material_id_from_url: str = initial_url.split("files/")[1].split("?")[0].strip("/")
        except IndexError:
            error_msg = f"Could not extract Canvas file ID from URL: {initial_url}"
            logger.warning(error_msg)
            return False, None, error_msg

        # TODO: These should be settings from settings.yaml
        canvas_base_domain: str = "https://utwente.instructure.com"
        canvas_user_id_for_download: str = "66733" # This is highly specific and needs configuration

        # Construct the direct download URL (this pattern is Canvas-specific)
        download_url = (
            f"{canvas_base_domain}/users/{canvas_user_id_for_download}/files/"
            f"{material_id_from_url}/download?download_frd=1"
        )

        # Filename will be {canvas_file_id}.pdf initially.
        # A separate step (`replace_canvas_id_with_material_id`) renames these later using the DB material_id.
        # This is because the DB material_id might not be known at this download stage if URL is the only input.
        final_filename = f"{material_id_from_url}.pdf" # Using Canvas ID for now
        output_filepath = self.download_dir.full / final_filename

        logger.info(f"Attempting download for Canvas ID {material_id_from_url} from: {download_url}")
        try:
            async with self.client.stream("GET", download_url) as response:
                response.raise_for_status() # Check for HTTP errors

                # Ensure download directory exists
                self.download_dir.mkdir(parents=True, exist_ok=True)

                bytes_downloaded: int = 0
                with open(output_filepath, "wb") as f:
                    async for chunk in response.aiter_bytes():
                        f.write(chunk)
                        bytes_downloaded += len(chunk)

                if bytes_downloaded > 0:
                    logger.info(f"Successfully downloaded {bytes_downloaded} bytes to {output_filepath.name} (Canvas ID: {material_id_from_url})")
                    return True, output_filepath, None
                else: # pragma: no cover (empty file is unlikely if status is 200)
                    logger.warning(f"Downloaded an empty file for Canvas ID {material_id_from_url} from {download_url}.")
                    # Clean up empty file
                    if output_filepath.exists(): output_filepath.unlink(missing_ok=True)
                    return False, None, "Downloaded empty file."


        except httpx.HTTPStatusError as e_status:
            error_msg = f"HTTP error {e_status.response.status_code} for Canvas ID {material_id_from_url} ({e_status.request.url})"
            logger.warning(error_msg)
            return False, None, error_msg
        except httpx.RequestError as e_req: # Covers network errors, timeouts, etc.
            error_msg = f"Request error for Canvas ID {material_id_from_url} ({e_req.request.url}): {e_req.__class__.__name__}"
            logger.warning(error_msg)
            return False, None, error_msg
        except Exception as e_general:
            error_msg = f"Unexpected error downloading Canvas ID {material_id_from_url}: {e_general.__class__.__name__} - {e_general}"
            logger.error(error_msg, exc_info=True)
            if output_filepath.exists(): output_filepath.unlink(missing_ok=True) # Cleanup
            return False, None, error_msg


async def main_download_all(max_concurrent: int = DEFAULT_MAX_CONCURRENT_DOWNLOADS) -> Tuple[List[Path], List[Tuple[str, str]]]:
    """
    Main orchestrator for downloading all relevant PDF files based on URLs in the database.
    It identifies URLs not yet downloaded and uses the HttpxDownloader to fetch them concurrently.

    Args:
        max_concurrent (int, optional): Maximum number of concurrent downloads.
                                        Defaults to DEFAULT_MAX_CONCURRENT_DOWNLOADS.

    Returns:
        Tuple[List[Path], List[Tuple[str, str]]]: A tuple containing:
            - list[Path]: Paths of successfully downloaded files.
            - list[tuple[str, str]]: Tuples of (URL, error_message) for failed downloads.
    """
    downloader = HttpxDownloader() # Uses default timeout
    downloaded_files_paths: List[Path] = []
    failed_url_errors: List[Tuple[str, str]] = []

    try:
        # This function renames files based on DB material_id. Should run *after* downloads if files are named by Canvas ID.
        # Or, if downloads are named by material_id directly, it might not be needed here.
        # Current download_file saves as {canvas_id}.pdf. So, this is important.
        await replace_canvas_id_with_material_id()

        semaphore = asyncio.Semaphore(max_concurrent)

        # Get URLs to download (filters out already downloaded ones)
        urls_df: pl.DataFrame = await get_urls_from_full_data() # This function logs and handles its own DB connection
        if urls_df.is_empty():
            logger.info("No URLs identified for download after filtering.")
            return [], []

        urls_to_download_list: List[str] = urls_df.get_column("url").to_list()

        logger.info(f"Starting bulk download of {len(urls_to_download_list)} files (max concurrent: {max_concurrent})...")

        download_tasks: List[asyncio.Task[Tuple[bool, Optional[Path], Optional[str]]]] = []
        async def _download_with_semaphore(url: str) -> Tuple[bool, Optional[Path], Optional[str]]:
            async with semaphore:
                return await downloader.download_file(url)

        for url_str in urls_to_download_list:
            download_tasks.append(asyncio.create_task(_download_with_semaphore(url_str)))

        results = await asyncio.gather(*download_tasks, return_exceptions=True)

        success_count: int = 0
        for i, result_item in enumerate(results):
            original_url = urls_to_download_list[i]
            if isinstance(result_item, Exception):
                logger.warning(f"Download task for {original_url} raised an exception: {result_item}")
                failed_url_errors.append((original_url, str(result_item)))
            elif isinstance(result_item, tuple): # Expected (success, filepath, error_msg)
                success, filepath, error_msg = result_item
                if success and filepath:
                    success_count += 1
                    downloaded_files_paths.append(filepath)
                else:
                    failed_url_errors.append((original_url, error_msg or "Unknown download error"))
            else: # Should not happen
                 logger.error(f"Unexpected result type from download task for {original_url}: {type(result_item)}")
                 failed_url_errors.append((original_url, "Unknown internal error"))


        logger.info(f"Download process complete. Success: {success_count}, Failed: {len(failed_url_errors)}")
        if failed_url_errors:
            logger.warning("Failed URLs:")
            for url_fail, err_msg_fail in failed_url_errors:
                logger.warning(f"  - {url_fail} (Error: {err_msg_fail})")

    except Exception as e_main: # Catch-all for main_download_all setup
        logger.error(f"Critical error in main_download_all: {e_main}", exc_info=True)
    finally:
        await downloader.close_client() # Ensure client is always closed

    return downloaded_files_paths, failed_url_errors


def get_already_downloaded_material_ids() -> List[str]:
    """
    Scans the PDF download directory and extracts material IDs from filenames.
    Assumes filenames are either `{material_id}.pdf` or `{material_id}_*.pdf`.

    Returns:
        List[str]: A list of material IDs corresponding to existing PDF files.
    """
    pdf_download_dir = SETTINGS.dirs.get(DirSetting.PDF_DOWNLOADS)
    if not pdf_download_dir or not pdf_download_dir.exists:
        logger.warning("PDF download directory not configured or found. Cannot get downloaded IDs.")
        return []

    existing_pdf_filenames: List[str] = [file.name for file in pdf_download_dir.files if file.extension == ".pdf"]

    material_ids_on_disk: List[str] = []
    for pdf_name in existing_pdf_filenames:
        if "_" in pdf_name:
            material_ids_on_disk.append(pdf_name.split("_")[0])
        else: # Assumes format is material_id.pdf
            material_ids_on_disk.append(pdf_name.split(".pdf")[0])

    # Filter for valid numeric IDs if necessary, though material_id can be non-numeric from source
    valid_material_ids = [mid for mid in material_ids_on_disk if mid.strip()] # Basic check for non-empty
    logger.info(f"Found {len(valid_material_ids)} material IDs from files in download directory.")
    return valid_material_ids


async def get_urls_from_full_data() -> pl.DataFrame:
    """
    Retrieves items from the database that have URLs and are not yet downloaded,
    and are not marked as 'Deleted'.

    Returns:
        pl.DataFrame: A DataFrame with columns "url", "material_id", "filename"
                      for items that need to be downloaded.
    """
    await init_tortoise_orm() # Ensure Tortoise is initialized for DB access
    try:
        full_item_count = await CopyrightItem.all().count()

        # Fetch items that have a URL, are not deleted, and might need downloading
        items_with_urls_qs = (
            CopyrightItem.filter(url__isnull=False, status__not=Status.DELETED.value) # Use enum value
            .exclude(url__in=["", "-"]) # Exclude empty or placeholder URLs
        )
        all_relevant_items_list = await items_with_urls_qs.values("material_id", "url", "filename")

        logger.info(f"Fetched {len(all_relevant_items_list)} relevant items with URLs from DB (total items: {full_item_count}).")
        if not all_relevant_items_list:
            return pl.DataFrame() # Return empty if no relevant items

        all_items_df = pl.from_dicts(all_relevant_items_list, schema={
            "material_id": pl.Int64, "url": pl.Utf8, "filename": pl.Utf8
        }) # Specify schema

        # Identify already downloaded files
        # Material IDs from filenames might be Canvas IDs or internal material_ids depending on when renaming occurs
        # Assuming get_already_downloaded_material_ids returns internal material_ids after potential renaming.
        material_ids_already_on_disk_str: List[str] = get_already_downloaded_material_ids()
        # Convert to int if material_id in DB is int for comparison
        material_ids_on_disk_int: Set[int] = {int(mid) for mid in material_ids_already_on_disk_str if mid.isdigit()}

        # Filter out items already downloaded
        # The 'material_id' column in all_items_df is Int64.
        items_to_download_df = all_items_df.filter(
            ~pl.col("material_id").is_in(list(material_ids_on_disk_int))
        )

        # Select necessary columns and ensure unique URLs
        final_urls_df = items_to_download_df.select(["url", "material_id", "filename"]).unique(subset=["url"], maintain_order=True)

        downloaded_count = len(material_ids_on_disk_int)
        total_relevant_url_items = len(all_items_df) # Count before filtering out downloaded

        logger.info(
            f"{final_urls_df.height} URLs remaining for download. "
            f"({total_relevant_url_items} total relevant, {downloaded_count} already on disk)."
        )
        # Removed input() for non-interactive operation
        # input("Press any key to continue...")
        return final_urls_df
    except Exception as e:
        logger.error(f"Error in get_urls_from_full_data: {e}")
        logger.debug(traceback.format_exc())
        return pl.DataFrame() # Return empty on error
    finally:
        await Tortoise.close_connections()


async def replace_canvas_id_with_material_id() -> None:
    """
    Renames downloaded PDF files from `{canvas_id}.pdf` or `{canvas_id}_....pdf`
    to `{material_id}.pdf` or `{material_id}_....pdf`.
    It fetches URL and material_id from all CopyrightItems in the database
    to perform this mapping.
    """
    await init_tortoise_orm()
    try:
        all_items_db_info = await CopyrightItem.all().values("material_id", "url")

        download_dir: Directory = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]
        if not download_dir.exists:
            logger.warning(f"PDF download directory {download_dir.full} not found. Cannot rename files.")
            return

        all_files_in_download_dir: List[File] = download_dir.files # Get File objects
        renamed_count: int = 0

        logger.info(f"Starting filename normalization for {len(all_files_in_download_dir)} files in {download_dir.full}...")
        for item_map in all_items_db_info:
            url, material_id_db = item_map.get("url"), item_map.get("material_id")
            if not url or not material_id_db or "/files/" not in url:
                continue # Skip if URL is invalid or no material_id

            try:
                # Canvas ID is the number after '/files/' and before any '?' or '/'
                canvas_id_from_url: str = url.split("files/")[1].split("?")[0].split("/")[0].strip("/")
                if not canvas_id_from_url.isdigit(): # Ensure extracted ID is numeric
                    logger.debug(f"Non-numeric Canvas ID '{canvas_id_from_url}' from URL '{url}'. Skipping.")
                    continue
            except IndexError:
                logger.debug(f"Could not extract Canvas file ID from URL: {url}")
                continue

            # Find files that start with this canvas_id
            for file_obj in all_files_in_download_dir:
                if file_obj.name.startswith(canvas_id_from_url):
                    # Construct new name: material_id + rest of the old name after canvas_id part
                    rest_of_filename = file_obj.name[len(canvas_id_from_url):] # e.g., "_some_suffix.pdf" or ".pdf"
                    new_filename = f"{material_id_db}{rest_of_filename}"

                    if file_obj.name != new_filename:
                        try:
                            file_obj.rename(new_filename) # File.rename handles moving to new name
                            logger.info(f"Renamed: {file_obj.name} -> {new_filename}")
                            renamed_count += 1
                        except Exception as e_rename:
                            logger.error(f"Error renaming file {file_obj.name} to {new_filename}: {e_rename}")

        logger.info(f"Filename normalization complete. Renamed {renamed_count} files.")
    except Exception as e:
        logger.error(f"An error occurred during replace_canvas_id_with_material_id: {e}")
        logger.debug(traceback.format_exc())
    finally:
        await Tortoise.close_connections()


# --- Example Usage (typically called from CLI or another orchestrator) ---
async def download_pdfs_example_usage(max_concurrent_downloads: int = DEFAULT_MAX_CONCURRENT_DOWNLOADS) -> None: # type: ignore
    """
    Example function demonstrating how to use HttpxDownloader to download all relevant PDFs.
    This function is intended for standalone execution or testing of the download process.
    """
    logger.info(f"Starting PDF download example with max {max_concurrent_downloads} concurrent downloads.")
    # Ensure Tortoise is initialized if DB operations are needed before/during download prep
    await init_tortoise_orm()

    downloaded_paths, failed_downloads = await main_download_all(max_concurrent=max_concurrent_downloads)

    logger.info(f"\n--- Download Summary ---")
    logger.info(f"Successfully downloaded files ({len(downloaded_paths)}):")
    # for p in downloaded_paths: logger.info(f"  - {p.name}") # Can be very verbose

    if failed_downloads:
        logger.warning(f"Failed downloads ({len(failed_downloads)}):")
        for url_fail, err_msg_fail in failed_downloads:
            logger.warning(f"  - URL: {url_fail}, Error: {err_msg_fail}")

    await Tortoise.close_connections() # Close Tortoise connections when done


if __name__ == "__main__": # pragma: no cover
    # This allows running the download process directly for testing.
    # Ensure proper logging setup if run this way.
    logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

    # Example of how it might be run:
    # asyncio.run(download_pdfs_example_usage(max_concurrent_downloads=5))

    # Or for renaming test:
    async def test_rename():
        await init_tortoise_orm()
        await replace_canvas_id_with_material_id()
        await Tortoise.close_connections()
    # asyncio.run(test_rename())
    logger.info("httpx_downloader.py executed directly (likely for testing).")
```
