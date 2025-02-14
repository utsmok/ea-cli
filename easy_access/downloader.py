
from selenium import webdriver

from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import info, warn, cool, Directory, File
import polars as pl
import time
from datetime import datetime
import Levenshtein

class Downloader:
    driver: webdriver.Chrome
    def __init__(self) -> None:
        self.download_dir = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]
    def setup_selenium(self) -> None:
        cool(f'setting up selenium')

        my_options = webdriver.ChromeOptions()
        profile_path = r"C:\Users\MokS\AppData\Local\Google\Chrome\User Data"
        profile_name = "Default" # Use "Profile N" if you have other accounts in Chrome
        my_options.add_argument(f"--user-data-dir={profile_path}")
        my_options.add_argument(f"--profile-directory={profile_name}")
        my_options.add_argument("--headless")

        my_options.add_experimental_option("prefs", {
            "download.default_directory":str(self.download_dir.full),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True
        })
        cool(f'initiating chrome driver')
        self.driver = webdriver.Chrome(
            options = my_options,
        )

        with open(SETTINGS.dirs[DirSetting.SCRIPT_DATA].full / 'stealth.min.js', 'r', encoding='utf8') as f:
            js = f.read()
        cool(f'injecting stealth.js')
        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {'source': js})
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

    def get_urls_from_full_data(self) -> pl.DataFrame:
        all_items: pl.DataFrame =pl.read_parquet(source=SETTINGS.dirs[DirSetting.SCRIPT_DATA].full / 'full_data.parquet', columns=['material_id', 'url', 'workflow_status', 'filename'])
        all_item_len = len(all_items)

        material_ids_downloaded = self.get_already_downloaded_material_ids()
        amount_done = len(all_items.filter(all_items['workflow_status'] == 'Done'))
        amount_downloaded = len(all_items.filter(all_items['material_id'].is_in(material_ids_downloaded)))
        #all_items = all_items.filter(all_items['workflow_status']!= 'Done')
        all_items = all_items.filter(~all_items['material_id'].is_in(material_ids_downloaded))

        urls = all_items.with_columns(
            pl.col('url').replace(old="-",new=None).replace("", None)
        ).select(['url', 'material_id', 'filename'])
        urls = urls.drop_nulls('url').unique('url')
        info(f'{len(urls)}/{all_item_len} urls remaining to download after filtering out {len(material_ids_downloaded)} already downloaded files ({amount_downloaded}) and files marked as done ({amount_done}).')
        return urls

    def download_pdfs(self, subset: list[str], max_amount: int = None) -> webdriver.Chrome:

        urls: pl.DataFrame = self.get_urls_from_full_data()
        if subset:
            urls = urls.filter(urls['material_id'].is_in(subset))
        if urls.is_empty():
            warn(f'No urls to download.')
            return
        return self.download(urls=urls, max_amount=max_amount)

    def download(self, urls: pl.DataFrame, max_amount: int = None) -> None:
        self.setup_selenium()
        info(f'Downloading {len(urls)} files')
        urls = urls.to_dicts()

        start_time = time.time()
        step_len = max(len(urls) // 20, 100)
        results = []
        first_datetime = None
        try:
            for index, item in enumerate(urls, start=1):
                x = 0
                while item['filename'] in self.download_dir.files:
                    x += 1
                    item['filename'] = item['filename'] + f'_{x}'

                item['created'] = self.download_file(item.get('url'))
                if not first_datetime:
                    first_datetime = item['created']
                results.append(item)

                items_per_sec = index / (time.time() - start_time)
                if items_per_sec > 1:
                    time.sleep(5)
                if max_amount and index >= max_amount:
                    info(f'[{index}/{len(urls)}] files downloaded in {time.time() - start_time:.2f} seconds. Max amount of downloads reached, stopping.')
                    break
                if index % step_len == 0:
                    info(f'[{index}/{len(urls)}] files downloaded in {time.time() - start_time:.2f} seconds')
        except Exception as e:
            warn(f'Error downloading file: {e}')
        finally:
            if results:
                info(f'Waiting until downloads are completed...')
                while any(file.name.endswith(('.crdownload', '.tmp')) for file in self.download_dir.files):
                    time.sleep(5)
                    info('Zzz...')
                info(f'Done! retrieved {len(results)} files in {time.time() - start_time:.2f} seconds')
                all_files = list(self.download_dir.files)
                all_files: list[File] = sorted(all_files, key=lambda x: x.created)
                selected_files: dict[str, File] = {file.name:file for file in all_files if file.created > first_datetime}
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
                for result in results:
                    file: File | None = selected_files.get(result['filename'])
                    if not file:
                        warn(f'Could not find {result["filename"]} in download dir.')
                        closest_match = max(selected_files.keys(), key=lambda x: Levenshtein.ratio(result['filename'], x, processor=lambda x: x.lower()), default=None)
                        if closest_match:
                            ratio = Levenshtein.ratio(result["filename"], closest_match)
                            if ratio < 0.9:
                                warn(f'Closest match has ratio below 0.9. Skipping.')
                                skipped += 1
                                continue
                            file = selected_files.get(closest_match)
                    if file:
                        file.rename(f'{result["material_id"]}_{result["filename"]}_{result["created"].strftime("%Y-%m-%d_%H-%M-%S")}.pdf')
                if skipped:
                    if skipped == missing:
                        cool("Number of skipped files matches number of missing files. All good.")
                    else:
                        warn(f"Number of skipped files does not match number of missing files. Please check: {skipped=}, {missing=}")
            else:
                warn(f'No files downloaded.')

        cool(f'Done! Downloaded {len(urls)} files in {time.time() - start_time:.2f} seconds')
        my_options = webdriver.ChromeOptions()

        my_options.add_experimental_option("prefs", {
            "download.default_directory":r"C:\Users\Sam\Downloads",
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": False
        })
        self.driver.close()
        self.driver = webdriver.Chrome(
            options = my_options,
        )
        self.driver.close()

        return self.driver


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


        info(f'downloading {material_id} from {download_url}')
        now = datetime.now()
        result_done = self.driver.get(download_url)
        self.driver.implicitly_wait(20) # is this necessary?
        return now
