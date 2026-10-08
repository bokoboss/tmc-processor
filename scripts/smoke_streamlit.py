"""Start Streamlit headlessly and verify its health endpoint on localhost."""

from __future__ import annotations

import os
from pathlib import Path
import socket
import subprocess
import sys
import tempfile
import time
from urllib.error import URLError
from urllib.request import urlopen


def _available_local_port() -> int:
    with socket.socket() as listener:
        listener.bind(("127.0.0.1", 0))
        return int(listener.getsockname()[1])


def main() -> int:
    project_root = Path(__file__).resolve().parents[1]
    port = _available_local_port()
    environment = os.environ.copy()
    environment["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    command = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(project_root / "app.py"),
        "--server.headless=true",
        "--server.address=127.0.0.1",
        f"--server.port={port}",
        "--browser.gatherUsageStats=false",
    ]
    with tempfile.TemporaryFile() as startup_log:
        process = subprocess.Popen(
            command,
            cwd=project_root,
            env=environment,
            stdout=startup_log,
            stderr=subprocess.STDOUT,
        )
        health_url = f"http://127.0.0.1:{port}/_stcore/health"
        deadline = time.monotonic() + 30
        try:
            while time.monotonic() < deadline:
                if process.poll() is not None:
                    startup_log.seek(0)
                    details = startup_log.read().decode("utf-8", errors="replace").strip()
                    raise RuntimeError(
                        f"Streamlit exited before becoming healthy (exit {process.returncode}).\n{details}"
                    )
                try:
                    with urlopen(health_url, timeout=1) as response:
                        body = response.read().decode("utf-8", errors="replace").strip().lower()
                        if response.status == 200 and body == "ok":
                            print(f"Streamlit healthy at {health_url}")
                            return 0
                        raise RuntimeError(
                            f"Unexpected Streamlit health response: {response.status} {body!r}."
                        )
                except URLError:
                    time.sleep(0.25)
            startup_log.seek(0)
            details = startup_log.read().decode("utf-8", errors="replace").strip()
            raise TimeoutError(
                f"Streamlit health endpoint did not respond within 30 seconds: {health_url}\n{details}"
            )
        finally:
            if process.poll() is None:
                process.terminate()
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait(timeout=5)


if __name__ == "__main__":
    raise SystemExit(main())
