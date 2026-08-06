from __future__ import annotations

import argparse

from .config import get_settings
from .outlook_desktop import sync_outlook_inbox
from .sync import sync_inbox


def sync_mail(initial_days: int = 90) -> int:
    """Run the configured email source's duplicate-safe catch-up sync."""
    settings = get_settings()
    if settings.mail_source == "outlook_desktop":
        return sync_outlook_inbox(initial_days)
    return sync_inbox(initial_days)


def main() -> None:
    parser = argparse.ArgumentParser(description="Synchronize the configured mail source to SQLite.")
    parser.add_argument("--initial-days", type=int, default=90)
    args = parser.parse_args()
    saved = sync_mail(args.initial_days)
    print(f"Sync complete. Added or updated {saved} email record(s).")


if __name__ == "__main__":
    main()
