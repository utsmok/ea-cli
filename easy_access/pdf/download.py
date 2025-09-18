import asyncio
import datetime
from pathlib import Path

import httpx
from loguru import logger
from tqdm.asyncio import tqdm_asyncio

from easy_access.db.base import ensure_db_inited
from easy_access.db.models import PDF, CopyrightItem, PDFCanvasMetadata
from easy_access.settings import DirSetting, Settings
from easy_access.utils import File


async def download_pdf_from_canvas(
    url: str, filepath: Path, session: httpx.AsyncClient
) -> tuple[File, PDFCanvasMetadata] | None:
    """
    Downloads a PDF from the given URL (which should be a canvas file url)
    and saves it to the given filepath.
    Returns a File object representing the downloaded file if successful, else None.
    """
    if not url or "/files/" not in url:
        logger.error(f"Invalid URL format: {url}")
        return None

    # TODO: check if any files have `usage_rights` that we need to store in the metadata as well, so far all files have null usage_rights
    file_id = url.split("/files/")[1].split("/")[0].split("?")[0]
    file_url = f"https://utwente.instructure.com/api/v1/files/{file_id}"
    try:
        response = await session.get(
            file_url, params={"include[]": ["usage_rights", "user"]}
        )
        response.raise_for_status()
        rate_limit_remaining = response.headers.get("X-Rate-Limit-Remaining", "1000000")
        if float(rate_limit_remaining) < 10:
            logger.warning(
                f"Rate limit critically low: {rate_limit_remaining:.2f}. Pausing for 10 seconds."
            )
            await asyncio.sleep(10)

        metadata = response.json()
        pdf_metadata = {
            "id": int(metadata.get("id")),
            "uuid": metadata.get("uuid"),
            "folder_id": int(metadata.get("folder_id")),
            "display_name": metadata.get("display_name"),
            "filename": metadata.get("filename"),
            "upload_status": metadata.get("upload_status"),
            "content_type": metadata.get("content-type"),
            "mime_class": metadata.get("mime_class"),
            "category": metadata.get("category"),
            "download_url": metadata.get("url"),
            "size": int(metadata.get("size")),
            "thumbnail_url": metadata.get("thumbnail_url"),
            "canvas_created_at": metadata.get("created_at"),
            "canvas_updated_at": metadata.get("updated_at"),
            "canvas_modified_at": metadata.get("modified_at"),
            "locked": metadata.get("locked"),
            "hidden": metadata.get("hidden"),
            "lock_at": metadata.get("lock_at"),
            "unlock_at": metadata.get("unlock_at"),
            "visibility_level": metadata.get("visibility_level"),
        }
        if metadata.get("user"):
            pdf_metadata.update(
                {
                    "user_id": metadata["user"].get("id"),
                    "user_anonymous_id": metadata["user"].get("anonymous_id"),
                    "user_display_name": metadata["user"].get("display_name"),
                    "user_avatar_image_url": metadata["user"].get("avatar_image_url"),
                    "user_html_url": metadata["user"].get("html_url"),
                    "user_pronouns": metadata["user"].get("pronouns"),
                }
            )

        pdf_metadata_obj = await PDFCanvasMetadata.create(**pdf_metadata)

        download_link = metadata.get("url")
        if not download_link:
            logger.error(f"No download URL found in metadata for file ID {file_id}")
            return None

        async with session.stream("GET", download_link) as file_response:
            file_response.raise_for_status()
            rate_limit_remaining = file_response.headers.get(
                "X-Rate-Limit-Remaining", "1000000"
            )
            if float(rate_limit_remaining) < 10:
                logger.warning(
                    f"Rate limit critically low: {rate_limit_remaining:.2f}. Pausing for 10 seconds."
                )
                await asyncio.sleep(10)
            with open(filepath, "wb") as f:
                async for chunk in file_response.aiter_bytes():
                    f.write(chunk)
        return File(filepath), pdf_metadata_obj
    except httpx.HTTPStatusError as e:
        logger.error(f"HTTP error downloading from {url}: {e.response.status_code}")
        return None
    except Exception as e:
        logger.error(f"Error downloading from {url}: {e}")
        return None


async def download_pdfs(settings: Settings, limit: int = 0) -> None:
    """
    Downloads all undownloaded PDFs from canvas, stores them in the default pdf_download_dir,
    and creates corresponding PDF entries in the database.
    If limit > 0, only downloads up to 'limit' PDFs.
    """

    await ensure_db_inited(settings)
    download_dir = settings.dirs[DirSetting.PDF_DOWNLOADS].full

    api_token = getattr(settings.university_settings, "canvas_api_token", None)
    if not api_token:
        logger.error("Canvas API token not found in settings, cannot download PDFs.")
        logger.debug(settings.university_settings.__dict__)
        return

    # filter all CopyrightItems that do not have a related PDF yet
    # and file_exists is True
    undownloaded_items = await CopyrightItem.filter(pdf__isnull=True, file_exists=True)

    if not undownloaded_items:
        logger.info("No undownloaded PDFs to download!")
        return
    if limit > 0:
        undownloaded_items = undownloaded_items[:limit]
    logger.info(f"Retrieving {len(undownloaded_items)} undownloaded PDFs.")

    semaphore = asyncio.Semaphore(value=5)

    async def download_single(item: CopyrightItem, session: httpx.AsyncClient):
        try:
            async with semaphore:
                filename = item.filename or f"{item.material_id}.pdf"
                # sanitize filename
                safe_filename = "".join(
                    c for c in filename if c.isalnum() or c in "._- "
                ).strip()
                if not safe_filename:
                    safe_filename = f"{item.material_id}.pdf"
                filepath = download_dir / f"{item.material_id}_{safe_filename}.pdf"
                result = await download_pdf_from_canvas(item.url, filepath, session)
                if result:
                    file, pdf_metadata_obj = result
                    pdf_dict = {
                        "copyright_item": item,
                        "current_file_name": file.name,
                        "filename": pdf_metadata_obj.filename,
                        "url": item.url,
                        "file_size": filepath.stat().st_size,
                        "retrieved_on": datetime.datetime.now(datetime.UTC),
                        "canvas_metadata": pdf_metadata_obj,
                    }
                    await PDF.create(**pdf_dict)
                else:
                    logger.error(f"Failed to download {item.material_id}")
        except Exception as e:
            logger.error(f"Error processing item {item.material_id}: {e}")
            return

    headers = {"Authorization": f"Bearer {api_token}"}
    async with httpx.AsyncClient(
        headers=headers, follow_redirects=True, timeout=20.0
    ) as session:
        tasks = [download_single(item, session) for item in undownloaded_items]
        await tqdm_asyncio.gather(*tasks)
    logger.info("Download complete")
