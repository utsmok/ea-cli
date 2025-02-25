
from selenium import webdriver

from easy_access.db.models import CopyrightItem
from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import info, warn, cool,  File
import polars as pl
import time
from datetime import datetime
import Levenshtein

default_download_dir = r"C:\Users\MokS\Downloads"
default_profile_path = r"C:\Users\MokS\AppData\Local\Google\Chrome\User Data"

class Downloader:
    driver: webdriver.Chrome
    def __init__(self) -> None:
        self.download_dir = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]
    def setup_selenium(self) -> None:
        cool(f'setting up selenium')

        my_options = webdriver.ChromeOptions()
        profile_path = default_profile_path
        profile_name = "Default" # Use "Profile N" if you have other accounts in Chrome
        my_options.add_argument(f"--user-data-dir={profile_path}")
        my_options.add_argument(f"--profile-directory={profile_name}")
        my_options.add_argument("--headless")
        my_options.add_argument("--enable-downloads")
        my_options.add_argument("--no-sandbox")
        my_options.enable_downloads = True

        my_options.add_experimental_option("prefs", {
            "download.default_directory":str(self.download_dir.full),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            "safebrowsing.enabled": False,
            "safebrowsing_for_trusted_sources_enabled": False,

        })
        cool(f'initiating chrome driver')
        self.driver = webdriver.Chrome(
            options = my_options,
        )

        with open(SETTINGS.dirs[DirSetting.SCRIPT_DATA].full / 'stealth.min.js', 'r', encoding='utf8') as f:
            js = f.read()
        cool(f'injecting stealth.js')
        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {'source': js})
        info(f'Download dir: {self.download_dir.full}. profile path: {profile_path}.')
        cool(text=f'done setting up selenium')

    def get_already_downloaded_material_ids(self) -> list[str]:
        existing_pdfs =[file.name for file in SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].files if file.extension == '.pdf']
        material_ids_downloaded = list()
        for pdf in existing_pdfs:
            if '_' in pdf:
                material_ids_downloaded.append(pdf.split('_')[0])
            else:
                warn(f'Could not extract material_id from {pdf}?')
        return material_ids_downloaded

    def get_already_classified_material_ids(self) -> list[str]:
        return [file.name.rstrip('.json') for file in SETTINGS.dirs[DirSetting.CLASSIFICATIONS].files if file.extension == '.json' and '_old' not in file.name]

    async def get_urls_from_full_data(self) -> pl.DataFrame:
        #all_items: pl.DataFrame = pl.read_parquet(source=SETTINGS.dirs[DirSetting.SCRIPT_DATA].full / 'full_data.parquet', columns=['material_id', 'url', 'workflow_status', 'filename'])
        all_items = await CopyrightItem.all().values('material_id', 'url', 'workflow_status', 'filename')
        all_item_len = len(all_items)
        all_items = pl.from_dicts(all_items)
        material_ids_downloaded = [int(x) for x in self.get_already_downloaded_material_ids()]
        material_ids_classified = [int(x) for x in self.get_already_classified_material_ids()]
        amount_done = len(all_items.filter(all_items['workflow_status'] == 'Done'))
        amount_downloaded = len(all_items.filter(all_items['material_id'].is_in(material_ids_downloaded)))
        #amount_classified = len(all_items.filter(all_items['material_id'].is_in(material_ids_classified)))
        #all_items = all_items.filter(all_items['workflow_status']!= 'Done')
        all_items = all_items.filter(~all_items['material_id'].is_in(material_ids_downloaded))
        #all_items = all_items.filter(~all_items['material_id'].is_in(material_ids_classified))

        urls = all_items.with_columns(
            pl.col('url').replace(old="-",new=None).replace("", None)
        ).select(['url', 'material_id', 'filename'])
        urls = urls.drop_nulls('url').unique('url')
        info(f'{len(urls)}/{all_item_len} urls remaining to download after filtering out {len(material_ids_downloaded)} already downloaded files ({amount_downloaded}) .')
        return urls

    async def rename_pdfs(self) -> None:

        urls: pl.DataFrame = await self.get_urls_from_full_data()
        all_files = list(self.download_dir.files)
        selected_files: dict[str, File] = {file.name:file for file in all_files}
        file_info: list[dict[str, str|datetime]] = urls.to_dicts()
        for result in file_info:
            print(f'.', end='')
            file: File | None = selected_files.get(result['filename'])
            if not file:
                file = selected_files.get(result['filename'].rstrip('.pdf')+' (1).pdf')
                if not file:
                    closest_match = max(selected_files.keys(), key=lambda x: Levenshtein.ratio(result['filename'], x, processor=lambda x: x.lower()), default=None)
                    if closest_match:
                        ratio = Levenshtein.ratio(result["filename"], closest_match)
                        if ratio < 0.9:
                            continue
                        file = selected_files.get(closest_match)
            if file:
                try:
                    file.rename(f'{result["material_id"]}_{result["filename"].rstrip('.pdf').replace(":","")}.pdf')
                except Exception as e:
                    warn(f'Error renaming file with info {result}: {e}')

            print(f'\n')

    async def download_pdfs(self, subset: list[str] | None, max_amount: int = None) -> webdriver.Chrome:
        deleted = 0
        for file in self.download_dir.files:
            if not str(file.name.split('_')[0]).isdigit():
                file.delete()
                deleted +=1
        if deleted:
            info(f'Deleted {deleted} incorrectly named pdf files.')
        urls: pl.DataFrame = await self.get_urls_from_full_data()
        if subset:
            urls = urls.filter(urls['material_id'].is_in(subset))
        if urls.is_empty():
            warn(f'No urls to download.')
            return True
        return self.download(urls=urls, max_amount=max_amount)

    def download(self, urls: pl.DataFrame, max_amount: int = None) -> None:
        def process_results(results, first_datetime, start_time):
            missing = 0
            info(f'Waiting until downloads are completed...')
            while any(file.name.endswith(('.crdownload', '.tmp')) for file in self.download_dir.files):
                time.sleep(5)
                info('Zzz...')
            info(f'Done! retrieved {len(results)} files in {time.time() - start_time:.2f} seconds')
            all_files = [f for f in self.download_dir.files if f.name.endswith('.pdf') and str(f.name.split('_')[0]).isdigit()]
            all_files: list[File] = sorted(all_files, key=lambda x: x.created)
            info(f"first_datetime: {first_datetime}. all_files len: {len(all_files)}")

            selected_files: dict[str, File] = {file.name:file for file in all_files if file.created > first_datetime}
            info(f"selected files len: {len(selected_files)}")
            if selected_files:
                info(f'# of files found in dir: {len(selected_files)}')
                if len(selected_files) < len(results):
                    warn(f'Not all files were downloaded?. {len(selected_files)}/{len(results)}')
                    missing = len(results) - len(selected_files)
                elif len(selected_files) > len(results):
                    warn(f'More files were downloaded than expected. {len(selected_files)}/{len(results)}')

            # sort results by 'created' datetime
            results = sorted(results, key=lambda x: x.get('created'))
            skipped = 0
            info(f'Now renaming {len(results)} files')
            for result in results:
                print(f'.', end='')
                file: File | None = selected_files.get(result['filename'])
                if not file:
                    closest_match = max(selected_files.keys(), key=lambda x: Levenshtein.ratio(result['filename'], x, processor=lambda x: x.lower()), default=None)
                    if closest_match:
                        ratio = Levenshtein.ratio(result["filename"], closest_match)
                        if ratio < 0.9:
                            skipped += 1
                            continue
                        file = selected_files.get(closest_match)
                if file:
                    try:
                        file.rename(f'{result["material_id"]}_{result["filename"]}_{result["created"].strftime("%Y-%m-%d_%H-%M-%S")}.pdf')
                    except Exception as e:
                        warn(f'Error renaming file with info {result}: {e}')
            if skipped:
                if skipped == missing:
                    cool("Number of skipped files matches number of missing files. All good.")
                else:
                    warn(f"Number of skipped files does not match number of missing files. Please check: {skipped=}, {missing=}")

        self.setup_selenium()
        info(f'Downloading {len(urls)} files')
        urls = urls.to_dicts()

        start_time = time.time()
        batch_start_time = time.time()
        step_len = min(len(urls) // 20, 20)
        results = []
        first_datetime = None
        try:
            for index, item in enumerate(urls, start=1):
                try:
                    x = 0
                    while item['filename'] in self.download_dir.files:
                        x += 1
                        item['filename'] = item['filename'] + f'_{x}'

                    item['created'] = self.download_file(item.get('url'))
                    if not first_datetime:
                        first_datetime = item['created']
                    results.append(item)
                    print(f'.', end='')
                    items_per_sec = index / (time.time() - start_time)
                    if items_per_sec > 1:
                        time.sleep(5)
                    if max_amount and index >= max_amount:
                        print(f'\n')
                        info(f'[{len(results)}/{len(urls)}] files downloaded in {time.time() - start_time:.2f} seconds. Max amount of downloads reached, stopping.')
                        break
                    if index % step_len == 0:
                        print(f'\n')
                        info(f'[{len(results)}/{len(urls)}] files downloaded in {time.time() - start_time:.2f} seconds')
                        process_results(results, first_datetime, batch_start_time)
                        first_datetime = None
                        results = []
                        batch_start_time = time.time()
                except Exception as e:
                    warn(f'Error downloading file: {e}')
                    continue
        except Exception as e:
            warn(f'Error downloading file: {e}')
        finally:
            if results:
                process_results(results, first_datetime, batch_start_time)

        cool(f'Done! Downloaded {len(urls)} files in {time.time() - start_time:.2f} seconds')
        for file in self.download_dir.files:
            if not str(file.name.split('_')[0]).isdigit():
                file.delete()
        return True

    def reset_chrome(self) -> None:
        my_options = webdriver.ChromeOptions()

        my_options.add_experimental_option("prefs", {
            "download.default_directory":default_download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": False
        })
        self.driver.close()
        self.driver = webdriver.Chrome(
            options = my_options,
        )
        self.driver.close()

    def download_file(self, url: str) -> None:
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
        material_id = url.split('files/')[1].replace("?","").strip()

        download_url = f'https://utwente.instructure.com/users/66733/files/{material_id}/download?download_frd=1'


        now = datetime.now()
        result_done = self.driver.get(download_url)
        self.driver.implicitly_wait(20) # is this necessary?
        return now
