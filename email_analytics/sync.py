from __future__ import annotations

import argparse
from datetime import datetime, timedelta, timezone

from .config import get_settings
from .database import get_state, initialize_database, set_state, upsert_emails
from .graph_client import GraphMailClient


def sync_inbox(initial_days: int = 90) -> int:
    """Fetch recent Inbox messages and upsert them into the local database."""
    settings = get_settings()
    initialize_database(settings.database_path)
    previous_sync = get_state(settings.database_path, "last_successful_sync_at")
    if previous_sync:
        since = datetime.fromisoformat(previous_sync).astimezone(timezone.utc) - timedelta(minutes=10)
    else:
        since = datetime.now(timezone.utc) - timedelta(days=initial_days)

    client = GraphMailClient(
        client_id=settings.client_id,
        tenant_id=settings.tenant_id,
        cache_path=settings.database_path.parent / "token_cache.bin",
        store_body_preview=settings.store_body_preview,
    )
    messages = client.get_messages_since(since)
    stored = upsert_emails(settings.database_path, messages)
    set_state(settings.database_path, "last_successful_sync_at", datetime.now(timezone.utc).isoformat())
    return stored


def main() -> None:
    parser = argparse.ArgumentParser(description="Synchronize Outlook Inbox messages to the local database.")
    parser.add_argument("--initial-days", type=int, default=90, help="Inbox history to fetch on first run.")
    args = parser.parse_args()
    stored = sync_inbox(initial_days=args.initial_days)
    print(f"Sync complete. Added or updated {stored} email record(s).")


if __name__ == "__main__":
    main()
