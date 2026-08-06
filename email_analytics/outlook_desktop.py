from __future__ import annotations

from datetime import datetime, timedelta, timezone
from typing import Any

import pythoncom
import win32com.client

from .config import get_settings
from .database import get_state, initialize_database, set_state, upsert_emails


OL_FOLDER_INBOX = 6
OL_MAIL_ITEM = 43


class OutlookDesktopError(RuntimeError):
    pass


def inbox_folder():
    """Return the default Inbox from classic Outlook on this Windows computer."""
    try:
        application = win32com.client.Dispatch("Outlook.Application")
        namespace = application.GetNamespace("MAPI")
        return namespace.GetDefaultFolder(OL_FOLDER_INBOX)
    except Exception as error:
        raise OutlookDesktopError(
            "Could not connect to classic Outlook. Open Outlook, complete its normal sign-in, and try again."
        ) from error


def _as_utc(value: datetime) -> str:
    if value.tzinfo is None:
        value = value.astimezone()
    return value.astimezone(timezone.utc).isoformat()


def _sender_email(item: Any) -> str:
    address = str(getattr(item, "SenderEmailAddress", "") or "")
    if getattr(item, "SenderEmailType", "") == "EX":
        try:
            exchange_user = item.Sender.GetExchangeUser()
            if exchange_user and exchange_user.PrimarySmtpAddress:
                return str(exchange_user.PrimarySmtpAddress).lower()
        except Exception:
            pass
    return address.lower() or "unknown@local-outlook"


def outlook_item_to_message(item: Any, store_body_preview: bool) -> dict[str, Any] | None:
    """Convert a classic Outlook MailItem into the database's common message shape."""
    if getattr(item, "Class", None) != OL_MAIL_ITEM:
        return None
    try:
        received_at = _as_utc(item.ReceivedTime)
    except Exception:
        return None
    try:
        stable_id = str(item.InternetMessageID or "")
    except Exception:
        stable_id = ""
    if not stable_id:
        stable_id = str(getattr(item, "EntryID", ""))
    if not stable_id:
        return None
    preview = ""
    if store_body_preview:
        try:
            preview = str(item.Body or "")[:1000]
        except Exception:
            preview = ""
    return {
        "id": f"outlook:{stable_id}",
        "from": {
            "emailAddress": {
                "name": str(getattr(item, "SenderName", "") or ""),
                "address": _sender_email(item),
            }
        },
        "subject": str(getattr(item, "Subject", "") or ""),
        "receivedDateTime": received_at,
        "bodyPreview": preview,
        "hasAttachments": bool(getattr(getattr(item, "Attachments", None), "Count", 0)),
        "importance": {0: "low", 1: "normal", 2: "high"}.get(getattr(item, "Importance", 1), "normal"),
        "conversationId": None,
        "webLink": None,
    }


def get_messages_since(since: datetime) -> list[dict[str, Any]]:
    """Read all newer default-Inbox mail from Outlook Desktop, newest first."""
    settings = get_settings()
    pythoncom.CoInitialize()
    try:
        items = inbox_folder().Items
        items.Sort("[ReceivedTime]", True)
        messages: list[dict[str, Any]] = []
        for position in range(1, items.Count + 1):
            item = items.Item(position)
            message = outlook_item_to_message(item, settings.store_body_preview)
            if not message:
                continue
            received = datetime.fromisoformat(message["receivedDateTime"])
            if received < since:
                break
            messages.append(message)
        return messages
    finally:
        pythoncom.CoUninitialize()


def sync_outlook_inbox(initial_days: int = 90) -> int:
    """Catch up Outlook Desktop mail and safely save it to SQLite."""
    settings = get_settings()
    initialize_database(settings.database_path)
    previous_sync = get_state(settings.database_path, "last_successful_sync_at")
    if previous_sync:
        since = datetime.fromisoformat(previous_sync).astimezone(timezone.utc) - timedelta(minutes=10)
    else:
        since = datetime.now(timezone.utc) - timedelta(days=initial_days)
    saved = upsert_emails(settings.database_path, get_messages_since(since))
    set_state(settings.database_path, "last_successful_sync_at", datetime.now(timezone.utc).isoformat())
    return saved
