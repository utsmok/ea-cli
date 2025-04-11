import asyncio
import json
from datetime import datetime
from pathlib import Path

import httpx
import polars as pl
from rich import print

from easy_access.db.base import CopyrightItem, init
from easy_access.db.models import Status
from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import cool, info, warn

SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]
SETTINGS.dirs[DirSetting.SCRIPT_DATA]


def load_cookies_from_file() -> httpx.Cookies:
    """
    Loads cookies from a Netscape format cookie file (cookies.txt).
    Handles comments, empty lines, and basic structure.
    """
    cookie_file = SETTINGS.dirs[DirSetting.SCRIPT_DATA].full / "cookies.secret"
    cool(f"Loading cookies from JSON file: {cookie_file}")
    cookies = httpx.Cookies()
    if not cookie_file.exists():
        raise FileNotFoundError(f"Cookie file not found: {cookie_file}")

    try:
        with open(cookie_file, encoding="utf-8") as f:
            cookie_data = json.load(f)  # Parse the JSON file

        if not isinstance(cookie_data, list):
            raise ValueError("Cookie file is not a JSON list.")

        for cookie_obj in cookie_data:
            # Ensure it's a dictionary and has the minimum required keys
            if (
                isinstance(cookie_obj, dict)
                and "name" in cookie_obj
                and "value" in cookie_obj
            ):
                name = cookie_obj["name"]
                value = cookie_obj["value"]
                domain = cookie_obj.get("domain")  # Use .get for optional keys
                path = cookie_obj.get("path", "/")  # Default path to '/' if missing

                # httpx's cookie jar primarily uses name, value, domain, path
                # It handles secure attribute implicitly based on request URL scheme
                # Expiry is not directly managed by the simple httpx.Cookies jar
                # Note: Leading dots in domain (e.g., ".example.com") are handled correctly by httpx
                cookies.set(name, value, domain=domain, path=path)
            else:
                info(f"Skipping invalid cookie object: {cookie_obj}")

        cool(f"Loaded {len(cookies)} cookies from JSON.")
        return cookies
    except json.JSONDecodeError as e:
        print(f"[ERROR] Failed to decode JSON from cookie file: {e}")
        raise
    except Exception as e:
        print(f"[ERROR] Failed to load or parse cookie file '{cookie_file}': {e}")
        raise


# --- Enhanced Downloader Class using httpx ---
class HttpxDownloader:
    def __init__(
        self,
    ) -> None:
        self.download_dir = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]
        self.cookies = load_cookies_from_file()
        # Set up a persistent client for connection pooling and cookie management
        # Add headers to mimic a browser somewhat
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        self.client = httpx.AsyncClient(
            cookies=self.cookies, headers=headers, follow_redirects=True, timeout=30.0
        )
        cool("httpx client initialized with loaded cookies.")

    async def close_client(self):
        """Closes the httpx client."""
        await self.client.aclose()
        cool("httpx client closed.")

    async def download_file(self, url: str) -> tuple[bool, Path | None, str | None]:
        """
        Downloads a file from a canvas URL using httpx.

        Args:
            url: The *initial* URL (e.g., https://.../files/XXXX)

        Returns:
            tuple: (success: bool, downloaded_filepath: Path | None, error_message: str | None)
        """
        if "/files/" not in url:
            error_msg = f"Invalid URL format: {url}. Expected '/files/'."
            info(error_msg)
            return False, None, error_msg

        try:
            material_id = url.split("files/")[1].split("?")[0].strip("/")
        except IndexError:
            error_msg = f"Could not extract material_id from URL: {url}"
            info(error_msg)
            return False, None, error_msg

        # Construct the download URL (adjust if the pattern differs for your LMS)
        # The provided example seems to use this pattern, verify it works consistently.
        # Assuming the domain is part of the input url or known:
        # Example: Infer domain if needed, or hardcode if always the same
        base_domain = (
            "https://utwente.instructure.com"  # Make this configurable if needed
        )
        user_id = (
            "66733"  # You might need to make this dynamic or configurable if it changes
        )

        download_url = (
            f"{base_domain}/users/{user_id}/files/{material_id}/download?download_frd=1"
        )
        now = datetime.now()
        timestamp = now.strftime("%Y%m%d%H%M%S")

        cool(f"Attempting download: {download_url}")
        try:
            async with self.client.stream("GET", download_url) as response:
                # Check for successful response before proceeding
                response.raise_for_status()  # Raises HTTPStatusError for 4xx/5xx responses

                final_filename = f"{material_id}.pdf"
                filepath = self.download_dir.full / final_filename

                # Stream download to file
                cool(f"Downloading to: {filepath}")
                bytes_downloaded = 0
                with open(filepath, "wb") as f:
                    async for chunk in response.aiter_bytes():
                        f.write(chunk)
                        bytes_downloaded += len(chunk)

                info(
                    f"Successfully downloaded {bytes_downloaded} bytes to {filepath.name}"
                )
                return True, filepath, None

        except httpx.HTTPStatusError as e:
            error_msg = f"HTTP error downloading {material_id}: {e.response.status_code} - {e.request.url} "
            print(f"[ERROR] {error_msg}")
            return False, None, error_msg
        except httpx.RequestError as e:
            error_msg = f"Network error downloading {material_id}: {e.__class__.__name__} - {e.request.url}"
            print(f"[ERROR] {error_msg}")
            return False, None, error_msg
        except Exception as e:
            error_msg = f"Unexpected error downloading {material_id}: {e.__class__.__name__} - {e}"
            print(f"[ERROR] {error_msg}")
            # Clean up potentially incomplete file
            if "filepath" in locals() and filepath.exists():
                try:
                    filepath.unlink()
                except OSError:
                    pass  # Ignore error during cleanup
            return False, None, error_msg


async def main_download_all(max_concurrent: int = 10):
    """
    Downloads files from a list of URLs concurrently using HttpxDownloader.
    """
    downloader = HttpxDownloader()

    semaphore = asyncio.Semaphore(max_concurrent)
    tasks = []
    await init()

    async def download_with_semaphore(url):
        async with semaphore:
            return await downloader.download_file(url)

    full_df: pl.DataFrame = await get_urls_from_full_data()
    urls_to_download = full_df["url"].to_list()

    cool(
        f"Starting bulk download of {len(urls_to_download)} files (max concurrent: {max_concurrent})..."
    )
    for url in urls_to_download:
        tasks.append(download_with_semaphore(url))

    results = await asyncio.gather(*tasks)

    # --- Process results ---
    success_count = 0
    failed_count = 0
    downloaded_files = []
    failed_urls = []

    for i, result in enumerate(results):
        success, filepath, error_msg = result
        if success and filepath:
            success_count += 1
            downloaded_files.append(filepath)
        else:
            failed_count += 1
            failed_urls.append((urls_to_download[i], error_msg))

    cool(f"Download complete. Success: {success_count}, Failed: {failed_count}")
    if failed_urls:
        info("Failed URLs:")
        for url, err in failed_urls:
            info(f"  - {url} (Error: {err})")

    # Ensure client is closed
    await downloader.close_client()

    # You could return results for further processing
    return downloaded_files, failed_urls


def get_already_downloaded_material_ids() -> list[str]:
    """
    Retrieves material_ids from the pdfs in the download dir.
    """
    existing_pdfs = [
        file.name
        for file in SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].files
        if file.extension == ".pdf"
    ]
    material_ids_downloaded = list()
    for pdf in existing_pdfs:
        if "_" in pdf:
            material_ids_downloaded.append(pdf.split("_")[0])
        else:
            warn(f"Could not extract material_id from {pdf}?")
    info(f"Found {len(material_ids_downloaded)} material_ids in download dir.")
    return material_ids_downloaded


async def get_urls_from_full_data() -> pl.DataFrame:
    """
    From all CopyrightItems in the db, grab "material_id", "url", "workflow_status", "filename", and "status" for all non-deleted items.
    Returns a dataframe with those columns.
    """
    full_item_len = await CopyrightItem.all().count()
    all_items = (
        await CopyrightItem.filter(url__isnull=False)
        .filter(url__not_in=["", "-"])
        .filter(
            status__in=[
                Status.PUBLISHED,
                Status.UNPUBLISHED,
                Status.PUBLISHED.value,
                Status.UNPUBLISHED.value,
            ]
        )
        .all()
        .values("material_id", "url", "workflow_status", "filename", "status")
    )
    all_item_len = len(all_items)
    info(
        f"Loaded {all_item_len} items of {full_item_len} total amount of items in db. Filtered out url-less items and DELETED items."
    )
    all_items = pl.from_dicts(all_items)
    info("Loaded into dataframe.")
    print(all_items.head())
    material_ids_downloaded = [int(x) for x in get_already_downloaded_material_ids()]
    # cast col 'material_id' to int
    all_items = all_items.with_columns(pl.col("material_id").cast(pl.Int32))
    amount_downloaded = len(
        all_items.filter(all_items["material_id"].is_in(material_ids_downloaded))
    )
    all_items = all_items.filter(
        ~all_items["material_id"].is_in(material_ids_downloaded)
    )

    urls = all_items.with_columns(
        pl.col("url").replace(old="-", new=None).replace("", None)
    ).select(["url", "material_id", "filename"])
    urls = urls.drop_nulls("url").unique("url")
    info(
        f"{len(urls)}/{all_item_len} urls remaining to download after filtering out {len(material_ids_downloaded)} already downloaded files ({amount_downloaded}) ."
    )
    return urls


# --- Example Usage ---
def download_pdfs():
    downloaded, failed = asyncio.run(main_download_all(max_concurrent=15))
    info(f"\nDownloaded Files ({len(downloaded)})")
    info(f"Failed Files ({len(failed)})")
