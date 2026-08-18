from __future__ import annotations

from datetime import datetime, timedelta, timezone
from typing import Any

import pythoncom
import win32con
import win32com.client
import win32gui

from .config import get_settings
from .database import (
    get_outlook_status_targets,
    get_state,
    initialize_database,
    set_state,
    update_read_statuses,
    upsert_emails,
)


OL_FOLDER_INBOX = 6
OL_MAIL_ITEM = 43


class OutlookDesktopError(RuntimeError):
    pass


def _bring_inspector_to_front(inspector: Any) -> None:
    """Restore an Outlook Inspector and place it above the dashboard browser."""
    try:
        # olNormalWindow = 2. This restores a minimized Inspector before the
        # Windows-level foreground request below.
        inspector.WindowState = 2
    except Exception:
        pass

    inspector.Activate()

    try:
        window_handle = int(inspector.HWND)
    except Exception:
        try:
            window_handle = int(win32gui.FindWindow(None, str(inspector.Caption)))
        except Exception:
            window_handle = 0

    if not window_handle:
        return

    try:
        flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
        win32gui.ShowWindow(window_handle, win32con.SW_RESTORE)
        win32gui.BringWindowToTop(window_handle)
        # Briefly make the Inspector topmost so Windows places it above the
        # browser, then immediately return it to ordinary window behavior.
        win32gui.SetWindowPos(window_handle, win32con.HWND_TOPMOST, 0, 0, 0, 0, flags)
        win32gui.SetWindowPos(window_handle, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, flags)
        win32gui.SetForegroundWindow(window_handle)
    except Exception:
        # Outlook's own Activate call above remains the safe fallback when a
        # corporate Windows policy blocks foreground-window manipulation.
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
    # Outlook's COM DATE values are local wall-clock times.  PyWin32 can attach
    # a misleading tzinfo.  Discard only that COM tag, then let the operating
    # system apply its own configured local timezone before storing UTC.
    local_wall_time = value.replace(tzinfo=None)
    return local_wall_time.astimezone(timezone.utc).isoformat()


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
    entry_id = str(getattr(item, "EntryID", "") or "")
    try:
        store_id = str(item.Parent.StoreID or "")
    except Exception:
        store_id = ""
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
        "isRead": not bool(getattr(item, "UnRead", False)),
        "importance": {0: "low", 1: "normal", 2: "high"}.get(getattr(item, "Importance", 1), "normal"),
        "conversationId": None,
        "webLink": None,
        "outlookEntryId": entry_id or None,
        "outlookStoreId": store_id or None,
    }


def display_outlook_item(entry_id: str, store_id: str | None = None, reply: bool = False) -> None:
    """Open an Outlook message or display a reply draft for it."""
    if not entry_id:
        raise OutlookDesktopError(
            "This email does not yet have an Outlook item identifier. Restart the listener to refresh it."
        )

    pythoncom.CoInitialize()
    try:
        application = win32com.client.Dispatch("Outlook.Application")
        namespace = application.GetNamespace("MAPI")
        item = namespace.GetItemFromID(entry_id, store_id) if store_id else namespace.GetItemFromID(entry_id)
        displayed_item = item.Reply() if reply else item
        displayed_item.Display(False)
        inspector = displayed_item.GetInspector
        _bring_inspector_to_front(inspector)
    except Exception as error:
        action = "create a reply for" if reply else "open"
        raise OutlookDesktopError(
            f"Could not {action} this message in Outlook. It may have been moved or deleted."
        ) from error
    finally:
        pythoncom.CoUninitialize()


def get_outlook_item_details(entry_id: str, store_id: str | None = None) -> dict[str, Any]:
    """Read extra details for one message directly from classic Outlook."""
    if not entry_id:
        raise OutlookDesktopError(
            "This email does not yet have an Outlook item identifier. Restart the listener to refresh it."
        )

    pythoncom.CoInitialize()
    try:
        application = win32com.client.Dispatch("Outlook.Application")
        namespace = application.GetNamespace("MAPI")
        item = namespace.GetItemFromID(entry_id, store_id) if store_id else namespace.GetItemFromID(entry_id)

        attachments: list[dict[str, Any]] = []
        attachment_collection = getattr(item, "Attachments", None)
        attachment_count = int(getattr(attachment_collection, "Count", 0) or 0)
        for position in range(1, attachment_count + 1):
            attachment = attachment_collection.Item(position)
            attachments.append(
                {
                    "name": str(getattr(attachment, "FileName", "") or "Attachment"),
                    "size": int(getattr(attachment, "Size", 0) or 0),
                }
            )

        received_at = getattr(item, "ReceivedTime", None)
        return {
            "sender_name": str(getattr(item, "SenderName", "") or ""),
            "sender_email": _sender_email(item),
            "to": str(getattr(item, "To", "") or ""),
            "cc": str(getattr(item, "CC", "") or ""),
            "subject": str(getattr(item, "Subject", "") or ""),
            "received_at": _as_utc(received_at) if received_at else None,
            "unread": bool(getattr(item, "UnRead", False)),
            "importance": {0: "low", 1: "normal", 2: "high"}.get(
                getattr(item, "Importance", 1), "normal"
            ),
            "categories": str(getattr(item, "Categories", "") or ""),
            "attachments": attachments,
            "body_preview": str(getattr(item, "Body", "") or "")[:4000],
            "size": int(getattr(item, "Size", 0) or 0),
        }
    except Exception as error:
        raise OutlookDesktopError(
            "Could not load this message from Outlook. It may have been moved or deleted."
        ) from error
    finally:
        pythoncom.CoUninitialize()


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


def refresh_outlook_read_statuses(limit: int = 500) -> int:
    """Reconcile stored read/unread flags against current Classic Outlook state."""
    settings = get_settings()
    targets = get_outlook_status_targets(settings.database_path, limit)
    if not targets:
        return 0

    pythoncom.CoInitialize()
    try:
        application = win32com.client.Dispatch("Outlook.Application")
        namespace = application.GetNamespace("MAPI")
        statuses: list[tuple[str, bool]] = []
        for target in targets:
            try:
                store_id = target.get("outlook_store_id") or None
                item = (
                    namespace.GetItemFromID(target["outlook_entry_id"], store_id)
                    if store_id
                    else namespace.GetItemFromID(target["outlook_entry_id"])
                )
                is_read = not bool(getattr(item, "UnRead", False))
            except Exception:
                # Moved or deleted messages should not prevent other statuses
                # from being refreshed.
                continue
            if int(is_read) != int(target["is_read"]):
                statuses.append((str(target["message_id"]), is_read))
        return update_read_statuses(settings.database_path, statuses)
    except Exception as error:
        raise OutlookDesktopError(
            "Could not refresh read statuses from Classic Outlook."
        ) from error
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
