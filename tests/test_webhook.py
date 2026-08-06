from email_analytics.config import Settings
from email_analytics.database import count_emails, initialize_database
from email_analytics.webhook import DemoEmail, demo_new_email


def test_simulated_trigger_immediately_creates_a_database_row(tmp_path, monkeypatch):
    database_path = tmp_path / "emails.db"
    settings = Settings(
        database_path=database_path,
        client_id=None,
        tenant_id="organizations",
        public_webhook_url=None,
        webhook_client_state=None,
        webhook_port=8787,
        dashboard_backup_refresh_seconds=300,
        store_body_preview=False,
        mail_source="outlook_desktop",
    )
    monkeypatch.setattr("email_analytics.webhook.get_settings", lambda: settings)
    initialize_database(database_path)

    result = demo_new_email(DemoEmail(subject="Triggered now"))

    assert result["status"] == "saved"
    assert count_emails(database_path) == 1
