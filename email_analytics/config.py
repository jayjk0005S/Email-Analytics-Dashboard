from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


PROJECT_ROOT = Path(__file__).resolve().parents[1]


@dataclass(frozen=True)
class Settings:
    database_path: Path
    client_id: str | None
    tenant_id: str
    public_webhook_url: str | None
    webhook_client_state: str | None
    webhook_port: int
    dashboard_backup_refresh_seconds: int
    store_body_preview: bool
    mail_source: str


def get_settings() -> Settings:
    """Load local configuration without ever placing secrets in source code."""
    load_dotenv(PROJECT_ROOT / ".env")
    configured_path = os.getenv("APP_DB_PATH", "data/email_analytics.db")
    database_path = Path(configured_path)
    if not database_path.is_absolute():
        database_path = PROJECT_ROOT / database_path

    store_body_preview = os.getenv("STORE_BODY_PREVIEW", "false").strip().lower() in {"1", "true", "yes"}
    mail_source = os.getenv("MAIL_SOURCE", "outlook_desktop").strip().lower()
    if mail_source not in {"outlook_desktop", "graph"}:
        raise ValueError("MAIL_SOURCE must be either 'outlook_desktop' or 'graph'.")
    return Settings(
        database_path=database_path,
        client_id=os.getenv("GRAPH_CLIENT_ID") or None,
        tenant_id=os.getenv("GRAPH_TENANT_ID", "organizations"),
        public_webhook_url=os.getenv("PUBLIC_WEBHOOK_URL") or None,
        webhook_client_state=os.getenv("WEBHOOK_CLIENT_STATE") or None,
        webhook_port=int(os.getenv("WEBHOOK_PORT", "8787")),
        dashboard_backup_refresh_seconds=int(os.getenv("DASHBOARD_BACKUP_REFRESH_SECONDS", "300")),
        store_body_preview=store_body_preview,
        mail_source=mail_source,
    )
