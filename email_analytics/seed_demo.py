from __future__ import annotations

import random
from datetime import datetime, timedelta, timezone

from .config import get_settings
from .database import initialize_database, upsert_emails


SENDERS = [
    ("HR India", "hr.india@example.com"),
    ("Finance Operations", "finance.ops@example.com"),
    ("Project Updates", "projects@example.com"),
    ("Customer Success", "customer.success@example.com"),
    ("Security Alerts", "security@example.com"),
    ("Marketing Team", "marketing@example.com"),
    ("IT Service Desk", "it.helpdesk@example.com"),
    ("Product Announcements", "product@example.com"),
    ("Vendor Updates", "vendors@example.com"),
]
SUBJECTS = [
    "Weekly status update", "Action required: review request", "Project milestone update",
    "Meeting follow-up and next steps", "Service notification", "Invoice available for review",
    "New announcement", "Customer request received",
]


def make_demo_messages(count: int = 511) -> list[dict]:
    randomizer = random.Random(42)
    start = datetime.now(timezone.utc) - timedelta(days=120)
    messages: list[dict] = []
    for index in range(count):
        sender_name, sender_email = randomizer.choices(
            SENDERS, weights=[18, 12, 10, 9, 7, 6, 5, 4, 3]
        )[0]
        received = start + timedelta(
            days=randomizer.randint(0, 120), hours=randomizer.randint(0, 23), minutes=randomizer.randint(0, 59)
        )
        messages.append(
            {
                "id": f"demo-message-{index:04d}",
                "from": {"emailAddress": {"name": sender_name, "address": sender_email}},
                "subject": randomizer.choice(SUBJECTS),
                "receivedDateTime": received.isoformat(),
                "bodyPreview": "Demo email record. Connect Microsoft Graph to view your actual Inbox data.",
                "hasAttachments": randomizer.random() < 0.18,
                "importance": randomizer.choices(["low", "normal", "high"], weights=[1, 15, 2])[0],
                "conversationId": f"demo-conversation-{index // 2}",
                "webLink": None,
            }
        )
    return messages


def main() -> None:
    settings = get_settings()
    initialize_database(settings.database_path)
    saved = upsert_emails(settings.database_path, make_demo_messages())
    print(f"Saved {saved} demo records to {settings.database_path}")


if __name__ == "__main__":
    main()
