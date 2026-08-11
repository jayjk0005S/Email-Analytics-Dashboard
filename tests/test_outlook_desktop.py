from datetime import datetime

from email_analytics.outlook_desktop import (
    display_outlook_item,
    get_outlook_item_details,
    outlook_item_to_message,
)


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


def test_outlook_item_becomes_a_duplicate_safe_metadata_record():
    message = outlook_item_to_message(DemoMailItem(), store_body_preview=False)
    assert message is not None
    assert message["id"] == "outlook:<desktop-outlook@example.com>"
    assert message["from"]["emailAddress"]["address"] == "sender@example.com"
    assert message["hasAttachments"] is True
    assert message["isRead"] is False
    assert message["importance"] == "high"
    assert message["bodyPreview"] == ""
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
