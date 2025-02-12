# Import the required modules
from selenium import webdriver
from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import Directory, File
import time

class Downloader:
    def __init__(self):
        self.setup_selenium()

    def setup_selenium(self) -> None:
        self.download_dir = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]
        my_options = webdriver.ChromeOptions()
        profile_path = r"C:\Users\MokS\AppData\Local\Google\Chrome\User Data"
        profile_name = "Default" # Use "Profile N" if you have other accounts in Chrome
        my_options.add_argument(f"--user-data-dir={profile_path}")
        my_options.add_argument(f"--profile-directory={profile_name}")
        my_options.add_argument("--headless")
        webdriver_path = r'C:\\dev\\chromedriver-win64\\chromedriver.exe',

        my_options.add_experimental_option("prefs", {
            "download.default_directory":str(self.download_dir.full),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True
        })
        self.driver = webdriver.Chrome(

            options = my_options,
        )

        with open(SETTINGS.dirs[DirSetting.SCRIPT_DATA].full / 'stealth.min.js', 'r', encoding='utf8') as f:
            js = f.read()

        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {'source': js})


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

        num_files_in_dir = len(list(self.download_dir.files))

        now = time.strftime("%Y-%m-%d_%H-%M-%S")
        self.driver.get(download_url)
        # Wait for the download to complete
        time_waited = 0
        while num_files_in_dir == len(list(self.download_dir.files)):
            time.sleep(5)
            time_waited += 5
            if time_waited > 120:
                print(f'Download timed out. Skipping {url}.')
                return

        file_name = self.download_dir.newest_file().name
        if '.pdf' in file_name:
            file_name = file_name.split('.pdf')[0]
            self.download_dir.rename_latest_file(f'{material_id}_{file_name}_{now}.pdf')
        else:
            print(f'File {file_name} is not a pdf. Skipping...')








