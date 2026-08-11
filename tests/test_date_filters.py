from datetime import date

from email_analytics.date_filters import parse_requested_date_range


def test_parse_requested_date_range_accepts_iso_values():
    assert parse_requested_date_range("2026-08-16", "2026-08-21") == (
        date(2026, 8, 16),
        date(2026, 8, 21),
    )


def test_parse_requested_date_range_rejects_invalid_or_reversed_values():
    assert parse_requested_date_range("2026-02-30", "2026-03-01") is None
    assert parse_requested_date_range("2026-08-21", "2026-08-16") is None
