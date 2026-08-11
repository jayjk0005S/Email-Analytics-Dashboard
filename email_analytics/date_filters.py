from __future__ import annotations

from datetime import date


def parse_requested_date_range(
    start_value: object,
    end_value: object,
) -> tuple[date, date] | None:
    """Parse Rio's ISO date query parameters without accepting invalid ranges."""
    try:
        start = date.fromisoformat(str(start_value))
        end = date.fromisoformat(str(end_value))
    except (TypeError, ValueError):
        return None
    return (start, end) if start <= end else None
