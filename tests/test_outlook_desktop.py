from datetime import datetime

from email_analytics.outlook_desktop import outlook_item_to_message


class Attachments:
    Count = 1


class DemoMailItem:
    Class = 43
    SenderEmailAddress = "sender@example.com"
    SenderEmailType = "SMTP"
    SenderName = "Desktop Outlook Sender"
    Subject = "Outlook Desktop email"
    ReceivedTime = datetime(2026, 8, 4, 9, 30)
    InternetMessageID = "<desktop-outlook@example.com>"
    EntryID = "fallback-entry-id"
    Attachments = Attachments()
    Importance = 2
    Body = "Private email body that should not be stored in metadata mode."


def test_outlook_item_becomes_a_duplicate_safe_metadata_record():
    message = outlook_item_to_message(DemoMailItem(), store_body_preview=False)
    assert message is not None
    assert message["id"] == "outlook:<desktop-outlook@example.com>"
    assert message["from"]["emailAddress"]["address"] == "sender@example.com"
    assert message["hasAttachments"] is True
    assert message["importance"] == "high"
    assert message["bodyPreview"] == ""
