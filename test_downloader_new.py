from rich import print

from easy_access.classification.httpx_downloader import (
    download_pdfs,
)
from easy_access.db.retrieve import retrieve_osiris_data

if False:
    download_pdfs()

print(retrieve_osiris_data([24597148]))

# asyncio.get_event_loop().run_until_complete(replace_canvas_id_with_material_id())


# process_items()
#
# try:
#   parse_pdfs()
# except Exception as e:
#   print(f"An error occurred: {e}")
#
