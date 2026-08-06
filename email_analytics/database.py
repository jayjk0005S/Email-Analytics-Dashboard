from __future__ import annotations

import sqlite3
import logging
from collections.abc import Iterable
from contextlib import contextmanager
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from .refresh_signal import update_refresh_signal


logger = logging.getLogger(__name__)


@contextmanager
def connection(database_path: Path):
    database_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(database_path)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def initialize_database(database_path: Path) -> None:
    with connection(database_path) as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS emails (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                graph_message_id TEXT NOT NULL UNIQUE,
                sender_name TEXT,
                sender_email TEXT NOT NULL,
                subject TEXT NOT NULL DEFAULT '',
                received_at TEXT NOT NULL,
                body_preview TEXT NOT NULL DEFAULT '',
                has_attachments INTEGER NOT NULL DEFAULT 0,
                importance TEXT NOT NULL DEFAULT 'normal',
                conversation_id TEXT,
                web_link TEXT,
                synced_at TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_emails_received_at
                ON emails(received_at);
            CREATE INDEX IF NOT EXISTS idx_emails_sender_email
                ON emails(sender_email);

            CREATE TABLE IF NOT EXISTS app_state (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            """
        )


def upsert_emails(database_path: Path, messages: Iterable[dict[str, Any]]) -> int:
    """Save messages idempotently. The Graph message ID prevents duplicate rows."""
    now = datetime.now(timezone.utc).isoformat()
    rows: list[tuple[Any, ...]] = []
    for message in messages:
        sender = message.get("from") or {}
        sender_address = sender.get("emailAddress") or {}
        graph_id = message.get("id")
        if not graph_id or not sender_address.get("address") or not message.get("receivedDateTime"):
            continue
        rows.append(
            (
                graph_id,
                sender_address.get("name") or "",
                sender_address["address"].lower(),
                message.get("subject") or "",
                message["receivedDateTime"],
                message.get("bodyPreview") or "",
                int(bool(message.get("hasAttachments"))),
                message.get("importance") or "normal",
                message.get("conversationId"),
                message.get("webLink"),
                now,
            )
        )

    if not rows:
        return 0

    with connection(database_path) as conn:
        before = conn.total_changes
        conn.executemany(
            """
            INSERT INTO emails (
                graph_message_id, sender_name, sender_email, subject, received_at,
                body_preview, has_attachments, importance, conversation_id, web_link, synced_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(graph_message_id) DO UPDATE SET
                sender_name = excluded.sender_name,
                sender_email = excluded.sender_email,
                subject = excluded.subject,
                body_preview = excluded.body_preview,
                has_attachments = excluded.has_attachments,
                importance = excluded.importance,
                conversation_id = excluded.conversation_id,
                web_link = excluded.web_link,
                synced_at = excluded.synced_at
            """,
            rows,
        )
        changed = conn.total_changes - before

    # The connection context has committed at this point, so an open dashboard
    # can safely reload the new data as soon as it sees this revision change.
    if changed:
        try:
            update_refresh_signal(database_path)
        except OSError:
            # Email storage is authoritative.  A missed signal is recovered by
            # the dashboard's five-minute backup refresh.
            logger.exception("Email data was committed, but the dashboard refresh signal could not be updated.")
    return changed


def get_state(database_path: Path, key: str) -> str | None:
    with connection(database_path) as conn:
        row = conn.execute("SELECT value FROM app_state WHERE key = ?", (key,)).fetchone()
    return row["value"] if row else None


def set_state(database_path: Path, key: str, value: str) -> None:
    with connection(database_path) as conn:
        conn.execute(
            """
            INSERT INTO app_state(key, value) VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (key, value),
        )


def get_dashboard_data(
    database_path: Path,
    sender_email: str | None = None,
    start_date: date | None = None,
    end_date: date | None = None,
) -> list[dict[str, Any]]:
    clauses: list[str] = []
    values: list[Any] = []
    if sender_email:
        clauses.append("sender_email = ?")
        values.append(sender_email)
    if start_date:
        clauses.append("date(received_at) >= date(?)")
        values.append(start_date.isoformat())
    if end_date:
        clauses.append("date(received_at) <= date(?)")
        values.append(end_date.isoformat())

    where = f" WHERE {' AND '.join(clauses)}" if clauses else ""
    query = (
        "SELECT sender_name, sender_email, subject, received_at, body_preview, "
        "has_attachments, importance, web_link FROM emails"
        f"{where} ORDER BY received_at DESC"
    )
    with connection(database_path) as conn:
        rows = conn.execute(query, values).fetchall()
    return [dict(row) for row in rows]


def get_senders(database_path: Path) -> list[str]:
    with connection(database_path) as conn:
        rows = conn.execute(
            "SELECT DISTINCT sender_email FROM emails ORDER BY sender_email COLLATE NOCASE"
        ).fetchall()
    return [row["sender_email"] for row in rows]


def get_date_bounds(database_path: Path) -> tuple[date, date] | None:
    with connection(database_path) as conn:
        row = conn.execute(
            "SELECT min(date(received_at)) AS minimum, max(date(received_at)) AS maximum FROM emails"
        ).fetchone()
    if not row or not row["minimum"]:
        return None
    return date.fromisoformat(row["minimum"]), date.fromisoformat(row["maximum"])


def count_emails(database_path: Path) -> int:
    with connection(database_path) as conn:
        return int(conn.execute("SELECT count(*) FROM emails").fetchone()[0])
