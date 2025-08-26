import pandas as pd
import requests
from time import sleep
import re
import math

INPUT_FILE = "ET_Combined_Data.xlsx"
OUTPUT_FILE = "ET_output_checked.xlsx"

API_TOKEN = "PUT API TOKEN HERE"
HEADERS = {"Authorization": f"Bearer {API_TOKEN}"}


def extract_file_id(url):
    """Extracts the file ID from a Canvas file URL"""
    if not url or (isinstance(url, tuple) and not url[0]):  # handle empty/tuple cases
        return None
    match = re.search(r"/(\d+)(?:\D*$|$)", str(url))
    return match.group(1) if match else None


def check_file_exists(file_id, session):
    """Checks if a Canvas file exists via the API"""
    if not file_id:  # empty URL or no ID
        return False
    api_url = f"https://utwente.instructure.com/api/v1/files/{file_id}"
    try:
        response = session.get(api_url, allow_redirects=True, timeout=10)
        return response.status_code == 200
    except requests.RequestException:
        return False


# Load Excel
df = pd.read_excel(INPUT_FILE)

if "url" not in df.columns:
    raise ValueError("Excel must contain a column named 'url'")

# Start session with API token
session = requests.Session()
session.headers.update(HEADERS)

results = []
for i, url in enumerate(df["url"], start=1):
    file_id = extract_file_id(url)
    exists = check_file_exists(file_id, session)
    results.append(exists)
    print(f"{i}/{len(df)} | {url} -> {exists}")
    # sleep(0.1)  # avoid hammering the server too fast

df["file_exists"] = results
df.to_excel(OUTPUT_FILE, index=False)
print(f"\nResults saved to {OUTPUT_FILE}")
