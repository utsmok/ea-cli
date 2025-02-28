import time
from datetime import datetime, timedelta
from pathlib import Path

import Levenshtein
import polars as pl
from selenium import webdriver
from selenium.webdriver.common.by import By

from easy_access.db.base import init
from easy_access.db.ingest import load_pdfs
from easy_access.db.models import CopyrightItem, Status
from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import File, cool, info, warn

default_download_dir = str(Path.home() / "Downloads")
default_profile_path = str(
    Path.home() / "AppData" / "Local" / "Google" / "Chrome" / "User Data"
)


class Downloader:
    driver: webdriver.Chrome

    def __init__(self) -> None:
        self.download_dir = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]

    def setup_selenium(self) -> None:
        """
        Sets up the selenium chrome driver with the correct options and settings.
        """
        cool("setting up selenium")

        my_options = webdriver.ChromeOptions()
        profile_path = default_profile_path
        profile_name = "Default"  # Use "Profile N" if you have other accounts in Chrome
        my_options.add_argument(f"--user-data-dir={profile_path}")
        my_options.add_argument(f"--profile-directory={profile_name}")
        my_options.add_argument("--headless")
        my_options.add_argument("--enable-downloads")
        my_options.add_argument("--no-sandbox")
        my_options.add_argument("--disable-extensions")
        my_options.enable_downloads = True

        my_options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": str(self.download_dir.full),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True,
                "safebrowsing.enabled": False,
                "safebrowsing_for_trusted_sources_enabled": False,
            },
        )
        cool("initiating chrome driver")
        self.driver = webdriver.Chrome(
            options=my_options,
        )

        with open(
            SETTINGS.dirs[DirSetting.SCRIPT_DATA].full / "stealth.min.js",
            "r",
            encoding="utf8",
        ) as f:
            js = f.read()
        cool("injecting stealth.js")
        self.driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument", {"source": js}
        )
        info(f"Download dir: {self.download_dir.full}. profile path: {profile_path}.")
        self.driver.implicitly_wait(5)
        cool(text="done setting up selenium")

    def reset_chrome(self) -> None:
        """
        Resets chrome back to default settings and closes the driver.
        """
        my_options = webdriver.ChromeOptions()
        my_options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": default_download_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": False,
            },
        )
        self.driver.close()
        self.driver = webdriver.Chrome(
            options=my_options,
        )
        time.sleep(5)
        self.driver.close()

    def get_already_downloaded_material_ids(self) -> list[str]:
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
        return material_ids_downloaded

    async def get_urls_from_full_data(
        self, subset: list[str] | list[int]
    ) -> pl.DataFrame:
        """
        From all CopyrightItems in the db, grab "material_id", "url", "workflow_status", "filename", and "status" for all non-deleted items.
        Returns a dataframe with those columns.
        """
        if subset:
            subset = [int(x) for x in subset]
        full_item_len = await CopyrightItem.all().count()
        if not subset:
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
        else:
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
                .filter(material_id__in=subset)
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
        material_ids_downloaded = [
            int(x) for x in self.get_already_downloaded_material_ids()
        ]
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

    async def rename_pdfs(self) -> None:
        """
        Renames the downloaded PDF files based on their material_id and original filename.
        If a file with the exact name is not found, it attempts to find the closest match.

        This should not be necessary if the download function is working correctly.
        """
        urls: pl.DataFrame = await self.get_urls_from_full_data()
        all_files = list(self.download_dir.files)
        selected_files: dict[str, File] = {file.name: file for file in all_files}
        file_info: list[dict[str, str | datetime]] = urls.to_dicts()
        for result in file_info:
            print(".", end="")
            file: File | None = selected_files.get(result["filename"])
            if not file:
                file = selected_files.get(
                    result["filename"].rstrip(".pdf") + " (1).pdf"
                )
                if not file:
                    closest_match = max(
                        selected_files.keys(),
                        key=lambda x: Levenshtein.ratio(
                            result["filename"], x, processor=lambda x: x.lower()
                        ),
                        default=None,
                    )
                    if closest_match:
                        ratio = Levenshtein.ratio(result["filename"], closest_match)
                        if ratio < 0.9:
                            continue
                        file = selected_files.get(closest_match)
            if file:
                try:
                    file.rename(
                        f"{result['material_id']}_{result['filename'].rstrip('.pdf').replace(':', '')}.pdf"
                    )
                except Exception as e:
                    warn(f"Error renaming file with info {result}: {e}")

            print("\n")

    async def download_pdfs(
        self, subset: list[str] | None, max_amount: int = None
    ) -> bool:
        """
        Downloads PDFs based on the URLs fetched from the database, potentially filtered by a subset of material IDs.
        If subset is None, all items are downloaded.
        If max_amount is set, only that amount of items are downloaded.
        """
        await init()
        deleted = 0
        for file in self.download_dir.files:
            if not str(file.name.split("_")[0]).isdigit():
                file.delete()
                deleted += 1
        if deleted:
            info(f"Deleted {deleted} incorrectly named pdf files.")

        urls: pl.DataFrame = await self.get_urls_from_full_data(subset)
        if urls.is_empty():
            warn("No urls to download.")
            return True
        skiplist = await self.download(urls=urls, max_amount=max_amount)
        # if items were skipped, keep retrying until done
        while skiplist:
            skiplist = await self.download(urls=None, skiplist=skiplist)
        await load_pdfs()
        return True

    async def download(
        self,
        urls: pl.DataFrame | None,
        max_amount: int = None,
        skiplist: list[dict] = list(),
    ) -> bool:
        """
        Downloads PDFs based on the URLs fetched from the database, potentially filtered by a subset of material IDs.
        If max_amount is set, only that amount of items are downloaded.

        Will rename the files in the download dir to match the material_id_filename_created.pdf format.
        If file name is not directly found, returns the closest match based on Levenshtein distance.

        Will also handle items that were deleted on canvas by marking them as deleted in the database.

        Ends by resetting chrome and removing misnamed pdfs.
        """

        async def process_results(results, first_datetime, start_time):
            """
            Renames the files in the download dir to match the material_id_filename_created.pdf format
            If file name is not directly found, returns the closest match based on Levenshtein distance.
            """
            missing = 0
            info("Waiting until downloads are completed...")
            while any(
                file.name.endswith((".crdownload", ".tmp"))
                for file in self.download_dir.files
            ):
                time.sleep(5)
                info("Zzz...")
            info(
                f"Done! retrieved {len(results)} files in {time.time() - start_time:.2f} seconds"
            )
            selected_files: dict[str, File] = {
                file.name: file
                for file in self.download_dir.files
                if file.created > first_datetime - timedelta(minutes=5)
            }

            all_files = [
                f
                for f in self.download_dir.files
                if f.name.endswith(".pdf") and str(f.name.split("_")[0]).isdigit()
            ]
            all_files: list[File] = sorted(all_files, key=lambda x: x.created)
            info(f"first_datetime: {first_datetime}. all_files len: {len(all_files)}")

            info(f"selected files len: {len(selected_files)}")
            if selected_files:
                info(f"# of files found in dir: {len(selected_files)}")
                if len(selected_files) < len(results):
                    warn(
                        f"Not all files were downloaded?. {len(selected_files)}/{len(results)}"
                    )
                    missing = len(results) - len(selected_files)
                elif len(selected_files) > len(results):
                    warn(
                        f"More files were downloaded than expected. {len(selected_files)}/{len(results)}"
                    )
            if not selected_files:
                warn("No files were found in download dir.")
                return
            # sort results by 'created' datetime
            results = sorted(results, key=lambda x: x.get("created"))
            skipped = 0
            info(f"Now renaming {len(results)} files")
            for result in results:
                print(".", end="")
                file: File | None = selected_files.get(result["filename"])
                if not file:
                    closest_match = max(
                        selected_files.keys(),
                        key=lambda x: Levenshtein.ratio(
                            result["filename"], x, processor=lambda x: x.lower()
                        ),
                        default=None,
                    )
                    if closest_match:
                        ratio = Levenshtein.ratio(result["filename"], closest_match)
                        if ratio < 0.9:
                            skipped += 1
                            continue
                        file = selected_files.get(closest_match)
                if file:
                    try:
                        file.rename(
                            f"{result['material_id']}_{result['filename']}_{result['created'].strftime('%Y-%m-%d_%H-%M-%S')}.pdf"
                        )
                    except Exception as e:
                        warn(f"Error renaming file with info {result}: {e}")
            if skipped:
                if skipped == missing:
                    cool(
                        "Number of skipped files matches number of missing files. All good."
                    )
                else:
                    warn(
                        f"Number of skipped files does not match number of missing files. Please check: {skipped=}, {missing=}"
                    )
            cool("Done with renaming! Resuming downloads...")

        self.setup_selenium()
        if not skiplist:
            info(f"Downloading {len(urls)} files")
            urls: list[dict] = urls.to_dicts()
        if skiplist:
            info(f"Retrying {len(skiplist)} skipped items.")
            urls = skiplist
        start_time = time.time()
        batch_start_time = time.time()
        step_len = max(min(len(urls) // 20, 20), 10)
        results = []
        first_datetime = None
        no_url = 0
        deleted_items = []
        total = 0
        skiplist = []
        if len(urls) < 20:
            print("Downloading these urls: ")
            for url in urls:
                print("    ", str(url))
        try:
            for index, item in enumerate(urls, start=1):
                try:
                    if max_amount and index >= max_amount:
                        print("\n")
                        info(
                            f"{len(results)} files downloaded in {time.time() - batch_start_time:.2f} seconds. Max amount of downloads reached, stopping."
                        )
                        break
                    if len(results) >= step_len:
                        print("\n")
                        info(
                            f"{len(results)} files downloaded in {time.time() - batch_start_time:.2f} seconds"
                        )
                        if results:
                            await process_results(
                                results, first_datetime, batch_start_time
                            )
                        first_datetime = None
                        results = []
                        batch_start_time = time.time()

                    if not item.get("url"):
                        no_url += 1
                        continue
                    if item["filename"] in self.download_dir.files:
                        skiplist.append(item)
                        continue
                    succes, item["created"] = await self.download_file(item.get("url"))
                    if not succes:
                        deleted_items.append(item)
                        continue
                    if not first_datetime:
                        first_datetime = item["created"]
                    results.append(item)
                    total += 1
                    print(".", end="")
                    items_per_sec = 0
                    if index:
                        items_per_sec = index / (time.time() - start_time)
                    if items_per_sec > 1:
                        time.sleep(5)
                except Exception as e:
                    warn(f"Error downloading file: {e}")
                    continue
        except Exception as e:
            warn(f"Error downloading file: {e}")
        finally:
            info(
                f"Done downloading.\n  Input urls:  {len(urls)}\n  Downloaded: {total}\n  Skipped: {no_url}\n  Deleted: {len(deleted_items)}"
            )
            try:
                if results:
                    await process_results(results, first_datetime, batch_start_time)
            except Exception as e:
                warn(f"error {e} while processing download results")

        info("Resetting & closing chrome, hold on...")
        self.reset_chrome()
        for file in self.download_dir.files:
            if not str(file.name.split("_")[0]).isdigit():
                file.delete()
            elif len(str(file.name.split("_")[0])) != 8:
                file.delete()
        cool("Done resetting chrome & removing misnamed pdfs!")

        if deleted_items:
            info(f"Now processing {len(deleted_items)} deleted items.")
            await self.handle_deleted_items(deleted_items)
        if skiplist:
            info(f"{len(skiplist)} items were skipped, retrying...")
        return skiplist

    async def download_file(self, url: str) -> tuple[bool, datetime]:
        """
        Downloads a file from a canvas url and stores it in the download dir (set during init) as:
        {material_id}_{orig_name}_{date_retrieved}.pdf
        example input urls:
        https://utwente.instructure.com/files/4393077?
        https://utwente.instructure.com/files/4659492?

        example download urls:
        https://utwente.instructure.com/users/66733/files/4393077/download?download_frd=1
        https://utwente.instructure.com/users/66733/files/4659492/download?download_frd=1
        """
        material_id = url.split("files/")[1].replace("?", "").strip()

        download_url = f"https://utwente.instructure.com/users/66733/files/{material_id}/download?download_frd=1"

        now = datetime.now()

        result_done = self.driver.get(download_url)
        # check if "Page Not Found" and/or "This file has been deleted" is on page

        if False:
            try:
                if (
                    "Page Not Found" in self.driver.page_source
                    and "This file has been deleted" in self.driver.page_source
                ):
                    info(
                        f"[by source] File {material_id} has been deleted from canvas."
                    )
                    return False, now
            except Exception as e:
                warn(
                    f"Error while checking if file {material_id} has been deleted: {e}"
                )
        # alternative option, disabled for now
        if False:
            try:
                if self.driver.find_element(by=By.CLASS_NAME, value="ic-Error-page"):
                    text = self.driver.find_element(
                        by=By.CLASS_NAME, value="ic-Error-page"
                    ).text
                    if "deleted" in text:
                        info(
                            f"[by id] File {material_id} has been deleted from canvas."
                        )
                        return False, now
            except Exception as e:
                warn(
                    f"error {e} while checking by id if file {material_id} has been deleted."
                )
        return True, now

    async def handle_deleted_items(self, items: list[dict] | list[str]) -> None:
        """
        Handles items that were deleted on canvas by marking them as deleted in the database.
        Input is a list of material ids, or dicts with a "material_id" key.
        """
        changed = 0
        not_found = 0
        already_deleted = 0
        for item in items:
            if isinstance(item, str):
                mat_id = item
            elif isinstance(item, dict):
                mat_id = item.get("material_id")

            if not mat_id:
                warn(f"Could not find material_id in item {item}. Skipping.")
                continue
            try:
                c_item = await CopyrightItem.get_or_none(material_id=int(mat_id))
                if c_item:
                    if c_item.status == Status.DELETED:
                        already_deleted += 1
                        continue
                    c_item.status = Status.DELETED
                    await c_item.save()
                    changed += 1
                else:
                    not_found += 1
            except Exception as e:
                warn(f"Error while marking item {mat_id} as deleted: {e}")
                continue

        cool(f"Marked {changed}/{len(items)} items as deleted.")
        cool(f"{already_deleted}/{len(items)} items were already marked as deleted.")
        warn(f"{not_found}/{len(items)} items not found in the database.")
        return
