from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import json
from pathlib import Path
import uvicorn
import re
app = FastAPI()

# Get the absolute path to the classify_webapp directory
BASE_DIR = Path(__file__).parent.absolute()
print(f"Base directory: {BASE_DIR}")

# Mount static directories
app.mount("/static", StaticFiles(directory=BASE_DIR), name="static")
app.mount("/pdf_downloads", StaticFiles(directory="pdf_downloads"), name="pdfs")

@app.get("/")
async def read_root():
    try:
        html_path = BASE_DIR / "index.html"
        print(f"Looking for HTML at: {html_path}")
        if not html_path.exists():
            print("HTML file not found!")
            raise HTTPException(status_code=404)
        return FileResponse(html_path)
    except Exception as e:
        print(f"Error serving index.html: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/list-pdfs")
async def list_pdfs():  # removed strict return type annotation
    pdf_dir = Path("pdf_downloads")
    pdfs = [f.name for f in pdf_dir.glob("*.pdf")]
    print(f"found {len(pdfs)} pdfs")
    pdf_without_extension = [name.rstrip(".pdf") for name in pdfs]
    json_names = [
        "".join([c for c in name if re.match(r"\w", c)]) + "_gemini_classification.json"
        for name in pdf_without_extension
    ]
    pdf_with_json = list(zip(pdfs, json_names))
    pdfs_with_json = []

    for pdf_filename, json_filename in pdf_with_json:
        try:
            json_path = pdf_dir / json_filename
            if json_path.exists():
                # Try UTF-8 first, fall back to Latin-1 if decoding fails
                with open(json_path, "rb") as jf:
                    raw_bytes = jf.read()
                try:
                    j_data = json.loads(raw_bytes.decode("utf-8"))
                except UnicodeDecodeError:
                    j_data = json.loads(raw_bytes.decode("latin-1", errors="replace"))
                pdfs_with_json.append({
                    "pdf": pdf_filename,
                    "json": json_filename,
                    "json_data": j_data
                })
        except Exception as e:
            print(f"Error loading JSON for {pdf_filename}: {e}")
            continue

    print(f"found {len(pdfs_with_json)} pdfs with json")
    return pdfs_with_json

@app.get("/pdf_downloads/{filename}")
async def serve_file(filename: str):
    file_path = Path("pdf_downloads") / filename
    print(f'Serving file: {file_path}')
    if not file_path.exists():
        raise HTTPException(status_code=404)
    if filename.endswith('.json'):
        return FileResponse(file_path, media_type="application/json; charset=utf-8")
    return FileResponse(file_path)

@app.post("/save-json/{filename}")
async def save_json(filename: str, data: dict):
    json_path = Path("pdf_downloads") / filename
    print(f"Saving JSON to {json_path}")
    try:
        # Try UTF-8 first, fall back to Latin-1 if needed
        with open(json_path, "rb") as f:
            raw_bytes = f.read()
        try:
            existing_data = json.loads(raw_bytes.decode("utf-8"))
        except UnicodeDecodeError:
            existing_data = json.loads(raw_bytes.decode("latin-1", errors="replace"))

        # Only update the specific fields
        existing_data["item_type"] = data["item_type"]
        existing_data["copyright_status"] = data["copyright_status"]
        print(f"Updated JSON: {existing_data}")
        # Write back with UTF-8 encoding
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(existing_data, f, indent=2, ensure_ascii=False)

        return {"status": "success", "data": existing_data}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)
