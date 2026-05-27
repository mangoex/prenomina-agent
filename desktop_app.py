from __future__ import annotations

import os
import socket
import sys
import threading
import time
import traceback
import urllib.request
from pathlib import Path

import uvicorn
import webview
from dotenv import load_dotenv


APP_NAME = "Prenomina"
HOST = "127.0.0.1"
STARTUP_TIMEOUT_SECONDS = float(os.getenv("PRENOMINA_STARTUP_TIMEOUT_SECONDS", "90"))


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def user_data_dir() -> Path:
    if os.name == "nt":
        documents = Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Documents"
    else:
        documents = Path.home() / "Documents"
    return documents / APP_NAME


def prepare_environment() -> Path:
    data_dir = user_data_dir()
    (data_dir / "input").mkdir(parents=True, exist_ok=True)
    (data_dir / "output").mkdir(parents=True, exist_ok=True)
    os.environ.setdefault("PRENOMINA_DESKTOP", "true")

    packaged_env = app_base_dir() / ".env"
    user_env = data_dir / ".env"
    if user_env.exists():
        load_dotenv(user_env, override=True)
    elif packaged_env.exists():
        load_dotenv(packaged_env, override=True)

    os.environ.setdefault("PRENOMINA_DATA_DIR", str(data_dir))
    return data_dir


def find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind((HOST, 0))
        return int(sock.getsockname()[1])


def wait_for_health(
    port: int,
    server_holder: dict[str, object],
    timeout_seconds: float = STARTUP_TIMEOUT_SECONDS,
) -> None:
    deadline = time.monotonic() + timeout_seconds
    url = f"http://{HOST}:{port}/health"
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        startup_error = server_holder.get("startup_error")
        if startup_error:
            raise RuntimeError(f"No se pudo iniciar {APP_NAME}.\n\n{startup_error}")
        try:
            with urllib.request.urlopen(url, timeout=1.0) as response:
                if response.status == 200:
                    return
        except Exception as exc:
            last_error = exc
            time.sleep(0.25)
    raise RuntimeError(f"No se pudo iniciar {APP_NAME}. Ultimo error: {last_error}")


def run_server(port: int, server_holder: dict[str, object]) -> None:
    try:
        from app.main import app

        config = uvicorn.Config(
            app,
            host=HOST,
            port=port,
            log_level="warning",
            reload=False,
        )
        server = uvicorn.Server(config)
        server_holder["server"] = server
        server.run()
    except Exception:
        server_holder["startup_error"] = traceback.format_exc()


def main() -> None:
    prepare_environment()
    port = find_free_port()
    server_holder: dict[str, object] = {}
    server_thread = threading.Thread(
        target=run_server,
        args=(port, server_holder),
        daemon=True,
    )
    server_thread.start()
    wait_for_health(port, server_holder)

    if "--smoke-test" in sys.argv:
        server = server_holder.get("server")
        if isinstance(server, uvicorn.Server):
            server.should_exit = True
        server_thread.join(timeout=5)
        return

    window = webview.create_window(
        APP_NAME,
        f"http://{HOST}:{port}",
        width=1180,
        height=820,
        min_size=(980, 680),
    )

    def stop_server() -> None:
        server = server_holder.get("server")
        if isinstance(server, uvicorn.Server):
            server.should_exit = True

    window.events.closed += stop_server
    webview.start()


if __name__ == "__main__":
    main()
