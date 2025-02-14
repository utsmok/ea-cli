from easy_access.classifier_api import main, delete_files
from easy_access.downloader import Downloader
import easy_access.settings
import asyncio

if __name__ == "__main__":
    delete_files()
    #downloader = Downloader()
    #downloader.download_pdfs(subset=[])

    asyncio.run(main())

