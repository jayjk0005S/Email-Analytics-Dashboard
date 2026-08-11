from email_analytics.advanced_dashboard import _email_selection_url


def test_selected_email_row_toggles_off_and_other_rows_select():
    assert _email_selection_url("message-one", "message-one") == "?"
    assert _email_selection_url("message-two", "message-one") == "?selected_email=message-two"
