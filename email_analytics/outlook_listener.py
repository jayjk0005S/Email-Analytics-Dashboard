from __future__ import annotations

import argparse
import atexit
import logging
import os
import time
from pathlib import Path

import pythoncom
import win32api
import win32com.client
import win32event
import winerror
import pywintypes

from .config import get_settings
from .database import initialize_database, upsert_emails
from .outlook_desktop import inbox_folder, outlook_item_to_message, sync_outlook_inbox


logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)


# Windows releases this mutex automatically if the listener exits unexpectedly.
# It prevents two startup commands from both importing/listening at the same time.
_listener_mutex = None


def listener_pid_path() -> Path:
    return get_settings().database_path.parent / "outlook_listener.pid"


def listener_is_running() -> bool:
    """Return whether a PID is alive using the Windows process API."""
    try:
        pid = int(listener_pid_path().read_text(encoding="utf-8"))
        handle = win32api.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    except (FileNotFoundError, ValueError, pywintypes.error):
        return False
    try:
        handle.Close()
        return True
    except pywintypes.error:
        return False


def _clear_pid_file() -> None:
    listener_pid_path().unlink(missing_ok=True)


def claim_listener_slot() -> bool:
    """Reserve the one allowed Outlook listener slot for this Windows user."""
    global _listener_mutex
    mutex = win32event.CreateMutex(None, False, "Local\\EmailAnalyticsOutlookListener")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        mutex.Close()
        return False
    _listener_mutex = mutex
    return True


def release_listener_slot() -> None:
    global _listener_mutex
    if _listener_mutex is not None:
        _listener_mutex.Close()
        _listener_mutex = None


class InboxEvents:
    """Receives classic Outlook's immediate ItemAdd event for new Inbox mail."""

    def OnItemAdd(self, item) -> None:  # noqa: N802 - Outlook COM event name
        settings = get_settings()
        try:
            message = outlook_item_to_message(item, settings.store_body_preview)
            if message:
                stored = upsert_emails(settings.database_path, [message])
                logger.info("New Outlook email stored (%s row affected).", stored)
        except Exception:
            logger.exception("Could not save a new Outlook Inbox item. The next catch-up sync will retry it.")


def listen_forever() -> None:
    settings = get_settings()
    initialize_database(settings.database_path)
    pythoncom.CoInitialize()
    try:
        inbox = inbox_folder()
        # Keep this COM event object alive for the entire process lifetime.
        events = win32com.client.WithEvents(inbox.Items, InboxEvents)
        logger.info("Listening for new classic Outlook Inbox emails.")
        while True:
            pythoncom.PumpWaitingMessages()
            _ = events
            time.sleep(0.2)
    finally:
        pythoncom.CoUninitialize()


def main() -> None:
    parser = argparse.ArgumentParser(description="Listen for new emails in classic Outlook Desktop.")
    parser.add_argument("--skip-catch-up", action="store_true")
    args = parser.parse_args()
    # Write this marker before the potentially long first import.  Startup may
    # be invoked again while that import is still running.
    if listener_is_running():
        logger.info("Another Outlook listener is already running; this copy will exit.")
        return
    if not claim_listener_slot():
        logger.info("Another Outlook listener is already running; this copy will exit.")
        return

    listener_pid_path().write_text(str(os.getpid()), encoding="utf-8")
    try:
        if not args.skip_catch_up:
            saved = sync_outlook_inbox()
            logger.info("Outlook catch-up complete (%s row affected).", saved)
        listen_forever()
    finally:
        _clear_pid_file()
        release_listener_slot()


if __name__ == "__main__":
    main()
