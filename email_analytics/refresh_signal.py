from __future__ import annotations

import os
from pathlib import Path
from uuid import uuid4


SIGNAL_FILE_NAME = "dashboard_refresh.signal"


def refresh_signal_path(database_path: Path) -> Path:
    """Keep the refresh signal beside the private dashboard database."""
    return database_path.parent / SIGNAL_FILE_NAME


def read_refresh_signal(database_path: Path) -> str | None:
    """Return the latest committed-data revision, or None before the first write."""
    try:
        return refresh_signal_path(database_path).read_text(encoding="utf-8").strip() or None
    except OSError:
        return None


def update_refresh_signal(database_path: Path) -> str:
    """Atomically publish a new revision after a successful database commit."""
    signal_path = refresh_signal_path(database_path)
    signal_path.parent.mkdir(parents=True, exist_ok=True)
    revision = uuid4().hex
    temporary_path = signal_path.with_name(f".{signal_path.name}.{revision}.tmp")
    try:
        temporary_path.write_text(revision, encoding="utf-8")
        os.replace(temporary_path, signal_path)
    finally:
        temporary_path.unlink(missing_ok=True)
    return revision
