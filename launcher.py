from __future__ import annotations

import os
import threading
import time
import webbrowser

import uvicorn

from server import app


def _open_browser(port: int) -> None:
    time.sleep(1.5)
    webbrowser.open(f"http://127.0.0.1:{port}")


def main() -> None:
    port = int(os.getenv("PORT", "8091"))
    threading.Thread(target=_open_browser, args=(port,), daemon=True).start()
    uvicorn.run(app, host="0.0.0.0", port=port, reload=False)


if __name__ == "__main__":
    main()
