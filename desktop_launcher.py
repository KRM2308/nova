from __future__ import annotations

import os
import socket
import threading
import time
import webbrowser

import requests
import uvicorn

from server import app


def _start_server(port: int) -> None:
    uvicorn.run(app, host="127.0.0.1", port=port, reload=False, log_level="warning")


def _find_free_port(preferred_port: int) -> int:
    # If preferred port is busy, choose a free ephemeral port.
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        if probe.connect_ex(("127.0.0.1", preferred_port)) != 0:
            return preferred_port
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _open_browser(url: str) -> None:
    time.sleep(1.2)
    webbrowser.open(url)


def _wait_server(url: str, timeout: float = 20.0) -> bool:
    started = time.time()
    while time.time() - started < timeout:
        try:
            r = requests.get(url, timeout=1.5)
            if r.status_code == 200 and r.json().get("app") == "pdf_nova":
                return True
        except Exception:
            pass
        time.sleep(0.2)
    return False


def main() -> None:
    preferred_port = int(os.getenv("PORT", "8091"))
    port = _find_free_port(preferred_port)
    t = threading.Thread(target=_start_server, args=(port,), daemon=True)
    t.start()
    base_url = f"http://127.0.0.1:{port}"
    health = f"{base_url}/api/health"
    if not _wait_server(health):
        raise RuntimeError("Le serveur local ne demarre pas.")

    threading.Thread(target=_open_browser, args=(base_url,), daemon=True).start()

    try:
        while True:
            time.sleep(60)
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
