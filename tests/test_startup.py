import socket

from email_analytics import startup


def test_port_is_open_reports_closed_ephemeral_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.bind(("127.0.0.1", 0))
        unused_port = probe.getsockname()[1]
    assert startup.port_is_open(unused_port) is False


def test_ensure_dashboard_tab_waits_for_app_and_desktop(monkeypatch):
    readiness = iter([False, True, True])
    opened_urls: list[str] = []

    monkeypatch.setattr(startup, "port_is_open", lambda _port: next(readiness))
    monkeypatch.setattr(startup, "interactive_desktop_is_ready", lambda: True)
    monkeypatch.setattr(startup, "open_browser_tab", lambda url: opened_urls.append(url) or True)
    monkeypatch.setattr(startup.time, "sleep", lambda _seconds: None)

    assert startup.ensure_dashboard_tab(attempts=3, delay_seconds=0) is True
    assert opened_urls == [startup.DASHBOARD_URL]


def test_open_browser_tab_uses_windows_fallback(monkeypatch):
    fallback_urls: list[str] = []

    monkeypatch.setattr(startup.webbrowser, "open", lambda *_args, **_kwargs: False)
    monkeypatch.setattr(startup.sys, "platform", "win32")
    monkeypatch.setattr(startup.os, "startfile", lambda url: fallback_urls.append(url), raising=False)

    assert startup.open_browser_tab(startup.DASHBOARD_URL) is True
    assert fallback_urls == [startup.DASHBOARD_URL]
