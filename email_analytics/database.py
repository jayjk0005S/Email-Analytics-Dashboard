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
LOCAL_RECEIVED_DATE_SQL = "date(received_at, '+5 hours', '+30 minutes')"
PRIORITY_RULE_TYPES = {"critical", "high", "normal"}


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
                is_read INTEGER NOT NULL DEFAULT 0,
                importance TEXT NOT NULL DEFAULT 'normal',
                conversation_id TEXT,
                web_link TEXT,
                outlook_entry_id TEXT,
                outlook_store_id TEXT,
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

            CREATE TABLE IF NOT EXISTS priority_rules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                layout_type TEXT NOT NULL DEFAULT 'two' CHECK(layout_type IN ('two', 'three')),
                rule_type TEXT NOT NULL CHECK(rule_type IN ('critical', 'high', 'normal')),
                pattern TEXT NOT NULL,
                normalized_pattern TEXT NOT NULL,
                created_at TEXT NOT NULL,
                UNIQUE(layout_type, rule_type, normalized_pattern)
            );

            CREATE INDEX IF NOT EXISTS idx_priority_rules_type
                ON priority_rules(rule_type);
            """
        )

        priority_rule_sql_row = conn.execute(
            "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'priority_rules'"
        ).fetchone()
        priority_rule_sql = str(priority_rule_sql_row[0] or "") if priority_rule_sql_row else ""
        if "critical" not in priority_rule_sql.casefold() or "layout_type" not in priority_rule_sql.casefold():
            legacy_columns = {row[1] for row in conn.execute("PRAGMA table_info(priority_rules)")}
            legacy_rows = conn.execute(
                "SELECT id, rule_type, pattern, normalized_pattern, created_at"
                + (", layout_type" if "layout_type" in legacy_columns else "")
                + " FROM priority_rules ORDER BY id"
            ).fetchall()
            conn.execute("ALTER TABLE priority_rules RENAME TO priority_rules_legacy")
            conn.execute(
                """
                CREATE TABLE priority_rules (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    layout_type TEXT NOT NULL DEFAULT 'two' CHECK(layout_type IN ('two', 'three')),
                    rule_type TEXT NOT NULL CHECK(rule_type IN ('critical', 'high', 'normal')),
                    pattern TEXT NOT NULL,
                    normalized_pattern TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    UNIQUE(layout_type, rule_type, normalized_pattern)
                )
                """
            )
            for row in legacy_rows:
                layout_type = str(row["layout_type"]) if "layout_type" in legacy_columns else "two"
                conn.execute(
                    """
                    INSERT INTO priority_rules
                        (id, layout_type, rule_type, pattern, normalized_pattern, created_at)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (row["id"], layout_type, row["rule_type"], row["pattern"], row["normalized_pattern"], row["created_at"]),
                )
            conn.execute("DROP TABLE priority_rules_legacy")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_priority_rules_type ON priority_rules(rule_type)")

        columns = {row[1] for row in conn.execute("PRAGMA table_info(emails)")}
        added_outlook_identifiers = False
        if "outlook_entry_id" not in columns:
            try:
                conn.execute("ALTER TABLE emails ADD COLUMN outlook_entry_id TEXT")
                added_outlook_identifiers = True
            except sqlite3.OperationalError:
                current = {row[1] for row in conn.execute("PRAGMA table_info(emails)")}
                if "outlook_entry_id" not in current:
                    raise
        if "outlook_store_id" not in columns:
            try:
                conn.execute("ALTER TABLE emails ADD COLUMN outlook_store_id TEXT")
                added_outlook_identifiers = True
            except sqlite3.OperationalError:
                current = {row[1] for row in conn.execute("PRAGMA table_info(emails)")}
                if "outlook_store_id" not in current:
                    raise
        if "is_read" not in columns:
            try:
                conn.execute("ALTER TABLE emails ADD COLUMN is_read INTEGER NOT NULL DEFAULT 0")
                # Re-import recent mail on the next listener start so migrated
                # rows receive their current read/unread state from Outlook.
                conn.execute("DELETE FROM app_state WHERE key = 'last_successful_sync_at'")
            except sqlite3.OperationalError:
                current = {row[1] for row in conn.execute("PRAGMA table_info(emails)")}
                if "is_read" not in current:
                    raise
        if added_outlook_identifiers:
            # The next listener start performs a full catch-up so existing rows
            # receive the identifiers required by the Open and Reply actions.
            conn.execute("DELETE FROM app_state WHERE key = 'last_successful_sync_at'")


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
                int(bool(message.get("isRead"))),
                message.get("importance") or "normal",
                message.get("conversationId"),
                message.get("webLink"),
                message.get("outlookEntryId"),
                message.get("outlookStoreId"),
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
                body_preview, has_attachments, is_read, importance, conversation_id, web_link,
                outlook_entry_id, outlook_store_id, synced_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(graph_message_id) DO UPDATE SET
                sender_name = excluded.sender_name,
                sender_email = excluded.sender_email,
                subject = excluded.subject,
                received_at = excluded.received_at,
                body_preview = excluded.body_preview,
                has_attachments = excluded.has_attachments,
                is_read = excluded.is_read,
                importance = excluded.importance,
                conversation_id = excluded.conversation_id,
                web_link = excluded.web_link,
                outlook_entry_id = excluded.outlook_entry_id,
                outlook_store_id = excluded.outlook_store_id,
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
        clauses.append(f"{LOCAL_RECEIVED_DATE_SQL} >= date(?)")
        values.append(start_date.isoformat())
    if end_date:
        clauses.append(f"{LOCAL_RECEIVED_DATE_SQL} <= date(?)")
        values.append(end_date.isoformat())

    where = f" WHERE {' AND '.join(clauses)}" if clauses else ""
    query = (
        "SELECT graph_message_id AS message_id, sender_name, sender_email, subject, "
        "received_at, body_preview, has_attachments, is_read, importance, web_link, "
        "outlook_entry_id, outlook_store_id FROM emails"
        f"{where} ORDER BY received_at DESC"
    )
    with connection(database_path) as conn:
        rows = conn.execute(query, values).fetchall()
    return [dict(row) for row in rows]


def get_outlook_status_targets(database_path: Path, limit: int = 500) -> list[dict[str, Any]]:
    """Return recent stored Outlook identifiers used for read-state reconciliation."""
    safe_limit = min(max(int(limit), 1), 5000)
    with connection(database_path) as conn:
        rows = conn.execute(
            "SELECT graph_message_id AS message_id, outlook_entry_id, outlook_store_id, is_read "
            "FROM emails WHERE outlook_entry_id IS NOT NULL AND outlook_entry_id <> '' "
            "ORDER BY received_at DESC LIMIT ?",
            (safe_limit,),
        ).fetchall()
    return [dict(row) for row in rows]


def update_read_statuses(database_path: Path, statuses: Iterable[tuple[str, bool]]) -> int:
    """Persist changed Outlook read states and notify open dashboard sessions."""
    now = datetime.now(timezone.utc).isoformat()
    rows = [(int(is_read), now, message_id, int(is_read)) for message_id, is_read in statuses]
    if not rows:
        return 0

    with connection(database_path) as conn:
        before = conn.total_changes
        conn.executemany(
            "UPDATE emails SET is_read = ?, synced_at = ? "
            "WHERE graph_message_id = ? AND is_read <> ?",
            rows,
        )
        changed = conn.total_changes - before

    if changed:
        try:
            update_refresh_signal(database_path)
        except OSError:
            logger.exception("Read states changed, but the dashboard refresh signal could not be updated.")
    return changed


def update_email_bodies(database_path: Path, bodies: dict[str, str]) -> int:
    """Store full Outlook body text for existing messages."""
    if not bodies:
        return 0
    now = datetime.now(timezone.utc).isoformat()
    rows = [(body, now, message_id, body) for message_id, body in bodies.items()]
    with connection(database_path) as conn:
        before = conn.total_changes
        conn.executemany(
            "UPDATE emails SET body_preview = ?, synced_at = ? "
            "WHERE graph_message_id = ? AND body_preview <> ?",
            rows,
        )
        return conn.total_changes - before


def get_senders(
    database_path: Path,
    start_date: date | None = None,
    end_date: date | None = None,
) -> list[str]:
    """Return only senders that occur inside the selected calendar range."""
    clauses: list[str] = []
    values: list[Any] = []
    if start_date:
        clauses.append(f"{LOCAL_RECEIVED_DATE_SQL} >= date(?)")
        values.append(start_date.isoformat())
    if end_date:
        clauses.append(f"{LOCAL_RECEIVED_DATE_SQL} <= date(?)")
        values.append(end_date.isoformat())

    where = f" WHERE {' AND '.join(clauses)}" if clauses else ""
    with connection(database_path) as conn:
        rows = conn.execute(
            "SELECT DISTINCT sender_email FROM emails"
            f"{where} ORDER BY sender_email COLLATE NOCASE",
            values,
        ).fetchall()
    return [row["sender_email"] for row in rows]


def get_date_bounds(database_path: Path) -> tuple[date, date] | None:
    with connection(database_path) as conn:
        row = conn.execute(
            f"SELECT min({LOCAL_RECEIVED_DATE_SQL}) AS minimum, "
            f"max({LOCAL_RECEIVED_DATE_SQL}) AS maximum FROM emails"
        ).fetchone()
    if not row or not row["minimum"]:
        return None
    return date.fromisoformat(row["minimum"]), date.fromisoformat(row["maximum"])


def count_emails(database_path: Path) -> int:
    with connection(database_path) as conn:
        return int(conn.execute("SELECT count(*) FROM emails").fetchone()[0])


def get_priority_rules(
    database_path: Path,
    rule_type: str | None = None,
    layout_type: str | None = None,
) -> list[dict[str, Any]]:
    """Return saved priority rules in the order they were created."""
    values: list[Any] = []
    clauses: list[str] = []
    if rule_type:
        if rule_type not in PRIORITY_RULE_TYPES:
            raise ValueError("rule_type must be 'critical', 'high', or 'normal'")
        clauses.append("rule_type = ?")
        values.append(rule_type)
    if layout_type:
        if layout_type not in {"two", "three"}:
            raise ValueError("layout_type must be 'two' or 'three'")
        clauses.append("layout_type = ?")
        values.append(layout_type)
    where = f" WHERE {' AND '.join(clauses)}" if clauses else ""

    with connection(database_path) as conn:
        rows = conn.execute(
            "SELECT id, layout_type, rule_type, pattern, created_at FROM priority_rules"
            f"{where} ORDER BY id",
            values,
        ).fetchall()
    return [dict(row) for row in rows]


def add_priority_rule(
    database_path: Path,
    rule_type: str,
    pattern: str,
    layout_type: str = "two",
) -> bool:
    """Add a case-insensitive keyword or sender email rule."""
    if rule_type not in PRIORITY_RULE_TYPES:
        raise ValueError("rule_type must be 'critical', 'high', or 'normal'")
    if layout_type not in {"two", "three"}:
        raise ValueError("layout_type must be 'two' or 'three'")
    cleaned = " ".join(pattern.split())
    if not cleaned:
        return False

    with connection(database_path) as conn:
        before = conn.total_changes
        conn.execute(
            """
            INSERT OR IGNORE INTO priority_rules
                (layout_type, rule_type, pattern, normalized_pattern, created_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (layout_type, rule_type, cleaned, cleaned.casefold(), datetime.now(timezone.utc).isoformat()),
        )
        return conn.total_changes > before


def delete_priority_rule(database_path: Path, rule_id: int) -> None:
    with connection(database_path) as conn:
        conn.execute("DELETE FROM priority_rules WHERE id = ?", (rule_id,))
