from __future__ import annotations

import argparse
import ctypes
import logging
import os
import socket
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

from .config import PROJECT_ROOT, get_settings
from .outlook_listener import listener_is_running


DASHBOARD_PORT = 8501
DASHBOARD_URL = f"http://127.0.0.1:{DASHBOARD_PORT}"
logger = logging.getLogger(__name__)


def configure_logging() -> None:
    """Persist startup diagnostics because this module normally runs through pythonw."""
    if logger.handlers:
        return
    log_path = PROJECT_ROOT / "data" / "startup.log"
    log_path.parent.mkdir(parents=True, exist_ok=True)
    handler = logging.FileHandler(log_path, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    logger.propagate = False


def port_is_open(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.25)
        return sock.connect_ex(("127.0.0.1", port)) == 0


def start_process(arguments: list[str]) -> None:
    """Start a detached local service using this project's virtual-environment Python."""
    creation_flags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen(
        [sys.executable, *arguments],
        cwd=PROJECT_ROOT,
        creationflags=creation_flags,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def interactive_desktop_is_ready() -> bool:
    """Return whether Windows Explorer has created the signed-in user's desktop."""
    if sys.platform != "win32":
        return True
    try:
        return bool(ctypes.windll.user32.GetShellWindow())
    except (AttributeError, OSError):
        return False


def open_browser_tab(url: str) -> bool:
    """Ask for a visible new browser tab, with a Windows-shell fallback."""
    try:
        if webbrowser.open(url, new=2, autoraise=True):
            return True
    except webbrowser.Error:
        logger.exception("The Python browser launcher could not open the dashboard tab.")

    start_file = getattr(os, "startfile", None)
    if sys.platform == "win32" and start_file is not None:
        try:
            start_file(url)
            return True
        except OSError:
            logger.exception("The Windows URL launcher could not open the dashboard tab.")
    return False


def ensure_dashboard_tab(attempts: int = 30, delay_seconds: float = 2.0) -> bool:
    """Wait for the app and interactive desktop, then reliably request a new tab."""
    for attempt in range(1, attempts + 1):
        dashboard_ready = port_is_open(DASHBOARD_PORT)
        desktop_ready = interactive_desktop_is_ready()
        if dashboard_ready and desktop_ready:
            if open_browser_tab(DASHBOARD_URL):
                logger.info("Dashboard tab opened: %s", DASHBOARD_URL)
                return True
            logger.warning("Dashboard tab launch attempt %s of %s failed.", attempt, attempts)
        time.sleep(delay_seconds)

    logger.error(
        "Dashboard is running=%s and interactive desktop is ready=%s, but its tab could not be opened.",
        port_is_open(DASHBOARD_PORT),
        interactive_desktop_is_ready(),
    )
    return False


def start_local_services(open_browser: bool = True) -> None:
    """Start the dashboard and local webhook listener only when they are not already running."""
    settings = get_settings()
    if settings.mail_source == "graph" and not port_is_open(settings.webhook_port):
        start_process(["-m", "email_analytics.webhook"])
    if settings.mail_source == "outlook_desktop" and not listener_is_running():
        start_process(["-m", "email_analytics.outlook_listener"])

    if not port_is_open(DASHBOARD_PORT):
        start_process(
            [
                "-m",
                "streamlit",
                "run",
                "app.py",
                "--server.headless=true",
                f"--server.port={DASHBOARD_PORT}",
            ]
        )

    if open_browser:
        ensure_dashboard_tab()


def main() -> None:
    configure_logging()
    parser = argparse.ArgumentParser(description="Start the local email dashboard and webhook listener.")
    parser.add_argument("--no-browser", action="store_true", help="Do not open the dashboard browser tab.")
    args = parser.parse_args()
    start_local_services(open_browser=not args.no_browser)


if __name__ == "__main__":
    main()
