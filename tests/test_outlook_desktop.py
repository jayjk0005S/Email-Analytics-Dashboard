from datetime import datetime, timezone

from email_analytics.outlook_desktop import (
    _as_utc,
    display_outlook_item,
    get_outlook_item_bodies,
    get_outlook_item_details,
    outlook_item_to_message,
)


def test_outlook_wall_clock_time_uses_the_computers_local_timezone():
    # PyWin32 may attach a misleading zone to Outlook's local wall time.  The
    # expected value deliberately uses the host timezone, so this test is
    # portable to computers outside India.
    outlook_time = datetime(2026, 8, 18, 19, 56, 35, tzinfo=timezone.utc)
    expected = outlook_time.replace(tzinfo=None).astimezone(timezone.utc).isoformat()

    assert _as_utc(outlook_time) == expected


class Attachments:
    Count = 1


class ParentFolder:
    StoreID = "default-store-id"


class DemoMailItem:
    Class = 43
    SenderEmailAddress = "sender@example.com"
    SenderEmailType = "SMTP"
    SenderName = "Desktop Outlook Sender"
    Subject = "Outlook Desktop email"
    ReceivedTime = datetime(2026, 8, 4, 9, 30)
    InternetMessageID = "<desktop-outlook@example.com>"
    EntryID = "fallback-entry-id"
    Parent = ParentFolder()
    Attachments = Attachments()
    UnRead = True
    Importance = 2
    Body = "Private email body that should not be stored in metadata mode."


def test_outlook_item_stores_full_body_for_priority_rules():
    message = outlook_item_to_message(DemoMailItem(), store_body_preview=False)
    assert message is not None
    assert message["id"] == "outlook:<desktop-outlook@example.com>"
    assert message["from"]["emailAddress"]["address"] == "sender@example.com"
    assert message["hasAttachments"] is True
    assert message["isRead"] is False
    assert message["importance"] == "high"
    assert message["bodyPreview"] == "Private email body that should not be stored in metadata mode."
    assert message["outlookEntryId"] == "fallback-entry-id"
    assert message["outlookStoreId"] == "default-store-id"


def test_display_outlook_item_opens_message_or_reply(monkeypatch):
    displayed: list[str] = []
    activated: list[str] = []

    class Inspector:
        def __init__(self, kind: str):
            self.kind = kind
            self.HWND = 100
            self.Caption = f"{kind} window"
            self.WindowState = 1

        def Activate(self) -> None:
            activated.append(self.kind)

    class ReplyDraft:
        GetInspector = Inspector("reply")

        def Display(self, modal: bool) -> None:
            assert modal is False
            displayed.append("reply")

    class MailItem:
        GetInspector = Inspector("open")

        def Display(self, modal: bool) -> None:
            assert modal is False
            displayed.append("open")

        def Reply(self):
            return ReplyDraft()

    class Namespace:
        def GetItemFromID(self, entry_id: str, store_id: str):
            assert entry_id == "entry-id"
            assert store_id == "store-id"
            return MailItem()

    class Application:
        def GetNamespace(self, name: str):
            assert name == "MAPI"
            return Namespace()

    monkeypatch.setattr(
        "email_analytics.outlook_desktop.win32com.client.Dispatch",
        lambda name: Application() if name == "Outlook.Application" else None,
    )
    monkeypatch.setattr("email_analytics.outlook_desktop.pythoncom.CoInitialize", lambda: None)
    monkeypatch.setattr("email_analytics.outlook_desktop.pythoncom.CoUninitialize", lambda: None)
    monkeypatch.setattr("email_analytics.outlook_desktop.win32gui.ShowWindow", lambda *_args: None)
    monkeypatch.setattr("email_analytics.outlook_desktop.win32gui.BringWindowToTop", lambda *_args: None)
    monkeypatch.setattr("email_analytics.outlook_desktop.win32gui.SetWindowPos", lambda *_args: None)
    monkeypatch.setattr("email_analytics.outlook_desktop.win32gui.SetForegroundWindow", lambda *_args: None)

    display_outlook_item("entry-id", "store-id")
    display_outlook_item("entry-id", "store-id", reply=True)

    assert displayed == ["open", "reply"]
    assert activated == ["open", "reply"]


def test_get_outlook_item_details_reads_live_message_without_saving_it(monkeypatch):
    class Attachment:
        FileName = "report.xlsx"
        Size = 2048

    class DetailAttachments:
        Count = 1

        def Item(self, position: int):
            assert position == 1
            return Attachment()

    class MailItem:
        SenderName = "Desktop Outlook Sender"
        SenderEmailAddress = "sender@example.com"
        SenderEmailType = "SMTP"
        To = "Friend <friend@example.com>"
        CC = ""
        Subject = "Detailed Outlook email"
        ReceivedTime = datetime(2026, 8, 4, 9, 30)
        UnRead = True
        Importance = 2
        Categories = "Customer"
        Attachments = DetailAttachments()
        Body = "Live Outlook message body"
        Size = 4096

    class Namespace:
        def GetItemFromID(self, entry_id: str, store_id: str):
            assert entry_id == "entry-id"
            assert store_id == "store-id"
            return MailItem()

    class Application:
        def GetNamespace(self, name: str):
            assert name == "MAPI"
            return Namespace()

    monkeypatch.setattr(
        "email_analytics.outlook_desktop.win32com.client.Dispatch",
        lambda name: Application() if name == "Outlook.Application" else None,
    )
    monkeypatch.setattr("email_analytics.outlook_desktop.pythoncom.CoInitialize", lambda: None)
    monkeypatch.setattr("email_analytics.outlook_desktop.pythoncom.CoUninitialize", lambda: None)

    details = get_outlook_item_details("entry-id", "store-id")

    assert details["sender_email"] == "sender@example.com"
    assert details["subject"] == "Detailed Outlook email"
    assert details["unread"] is True
    assert details["importance"] == "high"
    assert details["attachments"] == [{"name": "report.xlsx", "size": 2048}]
    assert details["body_preview"] == "Live Outlook message body"


def test_get_outlook_item_bodies_uses_one_com_session_and_skips_missing_items(monkeypatch):
    requested: list[tuple[str, str | None]] = []

    class MailItem:
        def __init__(self, body: str):
            self.Body = body

    class Namespace:
        def GetItemFromID(self, entry_id: str, store_id: str | None = None):
            requested.append((entry_id, store_id))
            if entry_id == "missing-entry":
                raise RuntimeError("Message moved")
            return MailItem(f"Full body for {entry_id}")

    class Application:
        def GetNamespace(self, name: str):
            assert name == "MAPI"
            return Namespace()

    dispatches: list[str] = []

    def dispatch(name: str):
        dispatches.append(name)
        return Application()

    monkeypatch.setattr("email_analytics.outlook_desktop.win32com.client.Dispatch", dispatch)
    monkeypatch.setattr("email_analytics.outlook_desktop.pythoncom.CoInitialize", lambda: None)
    monkeypatch.setattr("email_analytics.outlook_desktop.pythoncom.CoUninitialize", lambda: None)

    bodies = get_outlook_item_bodies(
        [
            {"message_id": "one", "outlook_entry_id": "entry-one", "outlook_store_id": "store-one"},
            {"message_id": "two", "outlook_entry_id": "entry-two", "outlook_store_id": ""},
            {"message_id": "missing", "outlook_entry_id": "missing-entry", "outlook_store_id": "store-one"},
            {"message_id": "no-outlook-id", "outlook_entry_id": ""},
        ]
    )

    assert dispatches == ["Outlook.Application"]
    assert requested == [
        ("entry-one", "store-one"),
        ("entry-two", None),
        ("missing-entry", "store-one"),
    ]
    assert bodies == {
        "one": "Full body for entry-one",
        "two": "Full body for entry-two",
    }
