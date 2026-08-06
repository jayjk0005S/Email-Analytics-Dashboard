from datetime import date

from email_analytics.database import count_emails, get_dashboard_data, get_date_bounds, get_senders, initialize_database, upsert_emails
from email_analytics.refresh_signal import read_refresh_signal, refresh_signal_path


def message(message_id: str, sender: str, received_at: str) -> dict:
    return {"id": message_id, "from": {"emailAddress": {"name": "Test sender", "address": sender}}, "subject": "Test email", "receivedDateTime": received_at, "bodyPreview": "Test preview", "hasAttachments": False, "importance": "normal"}


def test_upsert_is_idempotent_and_filters_data(tmp_path):
    db_path = tmp_path / "emails.db"
    initialize_database(db_path)
    assert read_refresh_signal(db_path) is None
    messages = [message("one", "one@example.com", "2026-08-01T08:00:00Z"), message("two", "two@example.com", "2026-08-02T08:00:00Z")]
    assert upsert_emails(db_path, messages) == 2
    first_revision = read_refresh_signal(db_path)
    assert first_revision
    upsert_emails(db_path, messages)
    assert read_refresh_signal(db_path) != first_revision
    assert not list(refresh_signal_path(db_path).parent.glob("*.tmp"))
    assert count_emails(db_path) == 2
    assert get_senders(db_path) == ["one@example.com", "two@example.com"]
    assert get_date_bounds(db_path) == (date(2026, 8, 1), date(2026, 8, 2))
    filtered = get_dashboard_data(db_path, sender_email="one@example.com")
    assert len(filtered) == 1
    assert filtered[0]["subject"] == "Test email"
