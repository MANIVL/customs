#!/usr/bin/env python3
"""Web launcher for ТаможенФормат.

Starts FastAPI/uvicorn and opens the app in the default browser.
"""
import os
import sys
import time
import threading
import urllib.request
import webbrowser

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BACKEND_DIR = os.path.join(BASE_DIR, "backend")
sys.path.insert(0, BACKEND_DIR)


def wait_until_ready(url: str, attempts: int = 80, delay: float = 0.25) -> bool:
    for _ in range(attempts):
        try:
            with urllib.request.urlopen(url, timeout=1) as resp:
                if resp.status == 200:
                    return True
        except Exception:
            pass
        time.sleep(delay)
    return False


def main() -> None:
    import uvicorn
    import server  # backend/server.py

    host = os.environ.get("APP_HOST", "127.0.0.1")
    port = int(os.environ.get("APP_PORT", "8756"))
    open_browser = os.environ.get("APP_OPEN_BROWSER", "1") != "0"
    local_url = f"http://127.0.0.1:{port}/"

    def open_when_ready() -> None:
        if wait_until_ready(local_url + "api/health") and open_browser:
            webbrowser.open(local_url)

    threading.Thread(target=open_when_ready, daemon=True).start()

    print(f"ТаможенФормат: {local_url}")
    if host not in ("127.0.0.1", "localhost"):
        print(f"Сервер слушает {host}:{port} (доступ из сети).")
    print("Чтобы остановить, нажмите Ctrl+C.")
    uvicorn.run(server.app, host=host, port=port, log_level="info")


if __name__ == "__main__":
    main()
