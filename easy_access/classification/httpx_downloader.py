import asyncio
import contextlib
import json
from pathlib import Path

import httpx
import polars as pl
from loguru import logger
from tortoise.expressions import Q, Subquery
from tortoise.transactions import in_transaction

from easy_access.db.base import CopyrightItem, init
from easy_access.db.models import PDF
from easy_access.settings import SETTINGS, DirSetting, Settings

SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]
SETTINGS.dirs[DirSetting.SCRIPT_DATA]


def load_cookies_from_file() -> httpx.Cookies:
    """
    Loads cookies from a Netscape format cookie file (cookies.txt).
    Handles comments, empty lines, and basic structure.
    """
    cookie_file = SETTINGS.dirs[DirSetting.SCRIPT_DATA].full / "cookies.secret"
    logger.success(f"Loading cookies from JSON file: {cookie_file}")
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
                # httpx.Cookies.set expects a domain string; only pass domain if present
                if domain:
                    cookies.set(name, value, domain=domain, path=path)
                else:
                    cookies.set(name, value, path=path)
            else:
                logger.info(f"Skipping invalid cookie object: {cookie_obj}")

        logger.success(f"Loaded {len(cookies)} cookies from JSON.")
        return cookies
    except json.JSONDecodeError as e:
        logger.error(f"Failed to decode JSON from cookie file: {e}")
        raise
    except Exception as e:
        logger.error(f"Failed to load or parse cookie file '{cookie_file}': {e}")
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
        logger.success("httpx client initialized with loaded cookies.")

    async def close_client(self):
        """Closes the httpx client."""
        await self.client.aclose()
        logger.success("httpx client closed.")

    async def download_file(self, url: str) -> tuple[bool, Path | None, str | None]:
        """
        Downloads a file from a canvas URL using httpx.

        Args:
            url: The *initial* URL

        Returns:
            tuple: (success: bool, downloaded_filepath: Path | None, error_message: str | None)
        """
        if "/files/" not in url:
            error_msg = f"Invalid URL format: {url}. Expected '/files/'."
            logger.info(error_msg)
            return False, None, error_msg

        try:
            material_id: str = url.split("files/")[1].split("?")[0].strip("/")
        except IndexError:
            error_msg = f"Could not extract material_id from URL: {url}"
            logger.info(error_msg)
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
        logger.success(f"Attempting download: {download_url}")
        try:
            async with self.client.stream("GET", download_url) as response:
                # Check for successful response before proceeding
                response.raise_for_status()  # Raises HTTPStatusError for 4xx/5xx responses

                final_filename = f"{material_id}.pdf"
                filepath = self.download_dir.full / final_filename

                # Stream download to file
                logger.success(f"Downloading to: {filepath}")
                bytes_downloaded = 0
                with open(filepath, "wb") as f:
                    async for chunk in response.aiter_bytes():
                        f.write(chunk)
                        bytes_downloaded += len(chunk)

                logger.info(
                    f"Successfully downloaded {bytes_downloaded} bytes to {filepath.name}"
                )

                # Persist download success in DB inside a transaction so the
                # download flags reflect the real outcome immediately.
                try:
                    try:
                        mat_id_int = int(material_id)
                    except Exception:
                        mat_id_int = None

                    if mat_id_int is not None:
                        async with in_transaction():
                            related_item = await CopyrightItem.get_or_none(
                                material_id=mat_id_int
                            )
                            pdf_obj = await PDF.get_or_none(material_id=mat_id_int)
                            if pdf_obj:
                                pdf_obj.current_file_name = filepath.name
                                pdf_obj.download_attempted = True
                                pdf_obj.download_succeeded = True
                                if related_item:
                                    pdf_obj.original_file_name = related_item.filename
                                    pdf_obj.original_page_count = related_item.pagecount
                                await pdf_obj.save()
                            else:
                                pdf_kwargs = {
                                    "material_id": mat_id_int,
                                    "current_file_name": filepath.name,
                                    "download_attempted": True,
                                    "download_succeeded": True,
                                }
                                if related_item:
                                    pdf_kwargs["original_file_name"] = (
                                        related_item.filename
                                    )
                                    pdf_kwargs["original_page_count"] = (
                                        related_item.pagecount
                                    )
                                await PDF.create(**pdf_kwargs)
                except Exception as e:
                    logger.warning(f"Could not persist download success to DB: {e}")

                return True, filepath, None

        except httpx.HTTPStatusError as e:
            error_msg = f"HTTP error downloading {material_id}: {e.response.status_code} - {e.request.url} "
            logger.error(error_msg)
            # Persist failure state
            try:
                try:
                    mat_id_int = int(material_id)
                except Exception:
                    mat_id_int = None

                if mat_id_int is not None:
                    async with in_transaction():
                        pdf_obj = await PDF.get_or_none(material_id=mat_id_int)
                        if pdf_obj:
                            pdf_obj.download_attempted = True
                            pdf_obj.download_succeeded = False
                            await pdf_obj.save()
                        else:
                            await PDF.create(
                                material_id=mat_id_int,
                                current_file_name=f"{mat_id_int}.pdf",
                                download_attempted=True,
                                download_succeeded=False,
                            )
            except Exception as ex:
                logger.warning(f"Could not persist download failure to DB: {ex}")
            return False, None, error_msg
        except httpx.RequestError as e:
            error_msg = f"Network error downloading {material_id}: {e.__class__.__name__} - {e.request.url}"
            logger.error(error_msg)
            # Persist failure state (same as above)
            try:
                try:
                    mat_id_int = int(material_id)
                except Exception:
                    mat_id_int = None

                if mat_id_int is not None:
                    async with in_transaction():
                        pdf_obj = await PDF.get_or_none(material_id=mat_id_int)
                        if pdf_obj:
                            pdf_obj.download_attempted = True
                            pdf_obj.download_succeeded = False
                            await pdf_obj.save()
                        else:
                            await PDF.create(
                                material_id=mat_id_int,
                                current_file_name=f"{mat_id_int}.pdf",
                                download_attempted=True,
                                download_succeeded=False,
                            )
            except Exception as ex:
                logger.warning(f"Could not persist download failure to DB: {ex}")
            return False, None, error_msg
        except Exception as e:
            error_msg = f"Unexpected error downloading {material_id}: {e.__class__.__name__} - {e}"
            logger.error(error_msg)
            # Clean up potentially incomplete file
            if "filepath" in locals() and filepath.exists():
                with contextlib.suppress(OSError):
                    filepath.unlink()  # Ignore error during cleanup
            # Persist failure state
            try:
                try:
                    mat_id_int = int(material_id)
                except Exception:
                    mat_id_int = None

                if mat_id_int is not None:
                    async with in_transaction():
                        pdf_obj = await PDF.get_or_none(material_id=mat_id_int)
                        if pdf_obj:
                            pdf_obj.download_attempted = True
                            pdf_obj.download_succeeded = False
                            await pdf_obj.save()
                        else:
                            await PDF.create(
                                material_id=mat_id_int,
                                current_file_name=f"{mat_id_int}.pdf",
                                download_attempted=True,
                                download_succeeded=False,
                            )
            except Exception as ex:
                logger.warning(f"Could not persist download failure to DB: {ex}")
            return False, None, error_msg


async def main_download_all(settings: Settings, max_concurrent: int = 10):
    """
    Downloads files from a list of URLs concurrently using HttpxDownloader.
    """
    try:
        downloader = HttpxDownloader()
        await replace_canvas_id_with_material_id()

        semaphore = asyncio.Semaphore(max_concurrent)
        tasks = []
        await init(settings=settings)

        async def download_with_semaphore(url):
            async with semaphore:
                return await downloader.download_file(url)

        full_df: pl.DataFrame = await get_urls_from_full_data()
        urls_to_download = full_df["url"].to_list()

        logger.success(
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

        logger.success(
            f"Download complete. Success: {success_count}, Failed: {failed_count}"
        )
        if failed_urls:
            logger.info("Failed URLs:")
            for url, err in failed_urls:
                logger.info(f"  - {url} (Error: {err})")

    finally:
        # Ensure client is closed
        await downloader.close_client()
    return downloaded_files, failed_urls


def get_currently_downloaded_material_ids() -> list[str]:
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
            material_ids_downloaded.append(pdf.split(".pdf")[0])
    logger.info(f"Found {len(material_ids_downloaded)} material_ids in download dir.")

    return material_ids_downloaded


async def get_urls_from_full_data() -> pl.DataFrame:
    """
    From all CopyrightItems in the db, grab "material_id", "url", "workflow_status", "filename", and "status" for all non-deleted items.
    Returns a dataframe with those columns.
    """
    # Exclude CopyrightItems that already have a PDF row marked as downloaded
    # using a subquery so the filtering is executed in SQL (more efficient).
    downloaded_subq = Subquery(
        PDF.filter(download_succeeded=True).values("material_id")
    )
    all_items = (
        await CopyrightItem.filter(url__isnull=False)
        .exclude(material_id__in=downloaded_subq)
        .values("material_id", "url", "workflow_status", "filename", "status")
    )
    logger.info(
        f"Retrieved {len(all_items)} items from the database (excluded already-downloaded PDFs)."
    )
    all_item_len = len(all_items)

    all_items = pl.from_dicts(all_items)
    logger.info("Loaded items into dataframe.")
    logger.debug(all_items.head())
    # Parse material ids found on disk; skip any non-integer names
    currently_downloaded_ids: list[int] = []
    for x in get_currently_downloaded_material_ids():
        try:
            currently_downloaded_ids.append(int(x))
        except Exception:
            logger.debug(f"Skipping non-integer downloaded id from filename: {x}")
    # cast col 'material_id' to int
    all_items = all_items.with_columns(pl.col("material_id").cast(pl.Int32))
    amount_downloaded = len(
        all_items.filter(all_items["material_id"].is_in(currently_downloaded_ids))
    )
    items_not_yet_downloaded = all_items.filter(
        ~all_items["material_id"].is_in(currently_downloaded_ids)
    )

    urls = items_not_yet_downloaded.with_columns(
        pl.col("url").replace(old="-", new=None).replace("", None)
    ).select(["url", "material_id", "filename"])
    urls = urls.drop_nulls("url").unique("url")
    logger.info(
        f"{len(urls)}/{all_item_len} urls remaining to download after filtering out {len(currently_downloaded_ids)} already downloaded files ({amount_downloaded}) ."
    )

    if amount_downloaded > 0:
        logger.info(
            "Making sure to set 'download_attempted' and 'download_succeeded' to True for files that are already downloaded..."
        )
        # grab all PDF files with material_id found in `all_items`
        # where:
        # pdf.download_attempted == False
        # or
        # pdf.download_succeeded == None or False
        # Build a queryset for PDFs that correspond to files present on disk
        # but are not yet marked as attempted/succeeded in the DB.
        queryset = PDF.filter(
            Q(material_id__in=currently_downloaded_ids)
            & (Q(download_attempted=False) | Q(download_succeeded__in=[None, False]))
        )

        # Perform a bulk update using the queryset.update(...) method which is
        # executed on the DB side and returns the number of updated rows.
        try:
            to_mark = await queryset.count()
            if to_mark:
                updated = await queryset.update(
                    download_attempted=True, download_succeeded=True
                )
                logger.info(
                    f"Marked {updated} PDF row(s) as download_attempted=True and download_succeeded=True in DB."
                )
            else:
                logger.debug("No PDF rows needed marking as downloaded in DB.")
        except Exception as e:
            logger.warning(
                f"Error while marking existing PDFs as downloaded in DB: {e}"
            )

    # Non-interactive: don't block in async code. Continue execution.
    logger.debug("Finished marking existing PDFs as downloaded (if any).")
    return urls


async def replace_canvas_id_with_material_id() -> None:
    """
    pull all items from db
    get material_id and url
    extract the "canvas_id" from the url
    then go through all files in the download dir
    and replace each occurrence of the canvas_id in any filename with the material_id
    """

    await init(settings=SETTINGS)
    all_items = await CopyrightItem.all().values("material_id", "url")

    download_dir = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]
    all_files = download_dir.files
    logger.info(f"Fixing the filenames of {len(all_files)} files...")
    for item in all_items:
        url = item["url"]
        if not url or "files/" not in url or "?" not in url:
            continue
        canvas_id: str = url.split("files/")[1].split("?")[0].strip("/")
        material_id = item["material_id"]
        [
            file.rename(file.name.replace(canvas_id, str(material_id)))
            for file in all_files
            if canvas_id in file.name
        ]
    logger.info(f"Done renaming {len(all_files)} files.")
    return


# --- Example Usage ---
def download_pdfs():
    downloaded, failed = asyncio.run(
        main_download_all(settings=SETTINGS, max_concurrent=15)
    )
    logger.info(f"\nDownloaded Files ({len(downloaded)})")
    logger.info(f"Failed Files ({len(failed)})")
