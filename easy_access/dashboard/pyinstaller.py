import threading
import time

import webview

from .dash import start


def start_app():
    uvicorn_thread = threading.Thread(target=start, daemon=True)
    uvicorn_thread.start()
    time.sleep(3)  # Adjust as needed, or implement a proper check
    window_title = "Easy Access Dashboard"
    window_url = "http://localhost:8000/"
    window = webview.create_window(window_title, window_url, width=1024, height=768)
    webview.start(debug=True, gui="edgechromium")
