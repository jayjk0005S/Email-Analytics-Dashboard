from email_analytics.priority_inbox_dashboard import _classify


def test_full_body_rules_use_critical_high_normal_precedence():
    record = {
        "message_id": "one",
        "sender_name": "Operations",
        "sender_email": "operations@example.com",
        "subject": "Daily update",
        "received_at": "2026-08-19T08:00:00Z",
        "body_preview": "A critical outage is active. This is also urgent and informational.",
        "importance": "normal",
    }
    rules = [
        {"rule_type": "normal", "pattern": "informational"},
        {"rule_type": "high", "pattern": "urgent"},
        {"rule_type": "critical", "pattern": "critical outage"},
    ]

    three_queue = _classify([record], rules, three_queue_layout=True)
    two_queue = _classify([record], rules, three_queue_layout=False)

    assert three_queue.iloc[0]["Priority"] == "Critical"
    assert three_queue.iloc[0]["Match"] == "Critical rule: critical outage"
    assert two_queue.iloc[0]["Priority"] == "High Priority"
    assert two_queue.iloc[0]["Match"] == "High rule: urgent"


def test_email_address_rule_can_match_inside_the_full_body():
    record = {
        "message_id": "one",
        "sender_email": "sender@example.com",
        "subject": "Contact details",
        "received_at": "2026-08-19T08:00:00Z",
        "body_preview": "Escalate this request to incident.owner@example.com immediately.",
        "importance": "normal",
    }
    rules = [{"rule_type": "critical", "pattern": "incident.owner@example.com"}]

    frame = _classify([record], rules, three_queue_layout=True)

    assert frame.iloc[0]["Priority"] == "Critical"


def test_explicit_normal_body_rule_overrides_outlook_high_importance():
    record = {
        "message_id": "one",
        "sender_email": "sender@example.com",
        "subject": "Status",
        "received_at": "2026-08-19T08:00:00Z",
        "body_preview": "This message is a routine newsletter.",
        "importance": "high",
    }
    rules = [{"rule_type": "normal", "pattern": "routine newsletter"}]

    frame = _classify([record], rules, three_queue_layout=True)

    assert frame.iloc[0]["Priority"] == "Normal"
    assert frame.iloc[0]["Match"] == "Normal rule: routine newsletter"
