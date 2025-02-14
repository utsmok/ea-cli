from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import json
from pathlib import Path
import uvicorn

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
async def list_pdfs():
    pdf_dir = Path("pdf_downloads")
    pdfs =  [f.name for f in pdf_dir.glob("*.pdf")]
    pdfs_with_json = [f for f in pdfs if Path(pdf_dir / (f.rstrip('.pdf') + '_gemini_classification.json')).exists()]
    return pdfs_with_json

@app.get("/pdf_downloads/{filename}")
async def serve_file(filename: str):
    file_path = Path("pdf_downloads") / filename
    if not file_path.exists():
        raise HTTPException(status_code=404)
    if filename.endswith('.json'):
        return FileResponse(file_path, media_type="application/json; charset=utf-8")
    return FileResponse(file_path)

@app.post("/save-json/{filename}")
async def save_json(filename: str, data: dict):
    json_path = Path("pdf_downloads") / filename

    # Load existing JSON with UTF-8 encoding
    if json_path.exists():
        with open(json_path, encoding='utf-8') as f:
            existing_data = json.load(f)
    else:
        existing_data = {}

    # Update only item_type and copyright_status
    existing_data.update({
        'item_type': data['item_type'],
        'copyright_status': data['copyright_status']
    })

    # Save back to file with UTF-8 encoding
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(existing_data, f, indent=2, ensure_ascii=False)

    return {"status": "success"}

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)

