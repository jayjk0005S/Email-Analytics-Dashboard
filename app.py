from __future__ import annotations

import base64
import time
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from email_analytics.config import PROJECT_ROOT, get_settings
from email_analytics.database import count_emails, get_dashboard_data, get_date_bounds, get_senders, get_state, initialize_database
from email_analytics.date_filters import parse_requested_date_range
from email_analytics.outlook_desktop import OutlookDesktopError, display_outlook_item
from email_analytics.refresh_signal import read_refresh_signal


PURPLE = "#A855F7"
CYAN = "#22D3EE"
PAGE_BACKGROUND = "#08111F"
CARD_BACKGROUND = "#101D31"
TEXT_PRIMARY = "#FFFFFF"
TEXT_MUTED = "#D1DDF0"
GRID = "#30425E"
BORDER_COLOR = "#3D5877"
BORDER_FOCUS = "#22D3EE"
ZETA_LOGO_PATH = PROJECT_ROOT / "assets" / "zeta-logo-primary.svg"
ZETA_LOGO_BASE64 = base64.b64encode(ZETA_LOGO_PATH.read_bytes()).decode("ascii")

st.set_page_config(page_title="Email Analytics Dashboard", page_icon="✉️", layout="wide")
st.markdown(
    f"""
    <style>
      /* Keep the dashboard compact while making text and controls easy to read. */
      :root {{ --filter-control-height: 42px; }}
      html {{ font-size: 15px !important; }}
      body {{ overflow-x: hidden; }}
      .stApp {{ background: {PAGE_BACKGROUND}; color: {TEXT_PRIMARY}; font-size: 1rem; overflow-x: hidden; }}
      .block-container, .stMainBlockContainer, [data-testid="stMainBlockContainer"], section[data-testid="stMain"] .block-container {{ box-sizing: border-box; width: 96% !important; max-width: 96% !important; margin-left: auto !important; margin-right: auto !important; padding-top: 1.6rem; padding-bottom: 2rem; }}
      .hero-title {{ font-size: 2.15rem; font-weight: 800; letter-spacing: -.04em; line-height: 1.15; color: {TEXT_PRIMARY}; text-shadow: 0 0 18px rgba(34, 211, 238, .22); margin: 0; }}
      .hero-copy {{ color: {TEXT_MUTED}; font-size: 1.05rem; line-height: 1.45; margin: .25rem 0 0; }}
      .zeta-logo-plate {{ background: #EAF3FF; border: 1px solid {BORDER_COLOR}; border-radius: 9px; box-shadow: 0 6px 18px rgba(0, 0, 0, .22); display: inline-flex; align-items: center; justify-content: center; min-width: 108px; min-height: 46px; padding: .35rem .55rem; box-sizing: border-box; transform: translateY(-5px); }}
      .zeta-logo-plate img {{ display: block; width: 88px; height: auto; }}
      .status-ribbon {{ background: {CARD_BACKGROUND}; border: 1px solid {BORDER_COLOR}; border-radius: 12px; padding: .8rem 1rem; margin: .9rem 0 1.2rem; box-shadow: 0 6px 18px rgba(0, 0, 0, .24); display: flex; align-items: center; flex-wrap: wrap; gap: .45rem 1.25rem; font-size: .95rem; line-height: 1.4; color: {TEXT_MUTED}; }}
      .status-dot {{ display: inline-block; width: 10px; height: 10px; border-radius: 50%; margin-right: .38rem; }}
      .status-item {{ white-space: nowrap; }}
      [data-testid="stMetric"] {{ background: {CARD_BACKGROUND}; border: 1px solid {BORDER_COLOR}; border-radius: 12px; padding: 1.05rem 1.2rem; box-shadow: 0 6px 18px rgba(0, 0, 0, .24); }}
      [data-testid="stMetricLabel"] {{ color: {TEXT_PRIMARY} !important; font-size: 1rem !important; font-weight: 750; text-shadow: 0 0 12px rgba(34, 211, 238, .16); }}
      [data-testid="stMetricValue"] {{ color: {TEXT_PRIMARY} !important; font-size: 2rem !important; line-height: 1.15 !important; font-weight: 750; text-shadow: 0 0 18px rgba(34, 211, 238, .22); }}
      div[data-testid="stDataFrame"] {{ background: {CARD_BACKGROUND}; border: 1px solid #293F5B; border-radius: 8px; overflow: hidden; }}
      div[data-testid="stDataFrame"] * {{ color: {TEXT_PRIMARY}; }}
      [data-testid="stVerticalBlockBorderWrapper"] {{ background: {CARD_BACKGROUND}; border: 1px solid {BORDER_COLOR}; border-radius: 14px; box-shadow: 0 6px 18px rgba(0, 0, 0, .24); min-height: 410px; overflow: hidden; }}
      [data-testid="stVerticalBlockBorderWrapper"] > div {{ padding: 1.25rem 1.35rem .65rem; }}
      .card-title {{ color: {TEXT_PRIMARY}; font-size: 1.3rem; line-height: 1.3; font-weight: 800; text-shadow: 0 0 12px rgba(34, 211, 238, .14); margin: 0; }}
      .card-subtitle {{ color: {TEXT_MUTED}; font-size: 1rem; line-height: 1.4; margin: .18rem 0 .7rem; }}
      .filter-heading, .filter-heading strong {{ color: {TEXT_PRIMARY}; font-size: 1.2rem; line-height: 1.3; font-weight: 850 !important; margin: 0 0 .42rem; text-shadow: 0 0 10px rgba(34, 211, 238, .14); }}
      .empty-state {{ background: {CARD_BACKGROUND}; border: 1px dashed {BORDER_COLOR}; border-radius: 16px; padding: 2rem; text-align: center; color: {TEXT_MUTED}; margin-top: 1rem; }}
      [data-testid="stSelectbox"] label, [data-testid="stDateInput"] label {{ color: {TEXT_PRIMARY} !important; font-size: 1.22rem !important; font-weight: 850 !important; margin-bottom: .35rem !important; }}
      [data-testid="stSelectbox"] > div, [data-testid="stDateInput"] > div, [data-testid="stSelectbox"] [data-baseweb="select"], [data-testid="stDateInput"] {{ width: 100% !important; }}
      /* Give each filter one border on the interactive control itself. */
      [data-testid="stSelectbox"] {{ outline: none !important; border: 0 !important; box-shadow: none !important; overflow: visible !important; }}
      [data-testid="stSelectbox"] div[data-baseweb="select"],
      [data-testid="stSelectbox"] div[data-baseweb="select"] > div,
      [data-testid="stDateInput"] div[data-baseweb="input"] {{ height: var(--filter-control-height) !important; min-height: var(--filter-control-height) !important; box-sizing: border-box !important; }}
      [data-testid="stSelectbox"] div[data-baseweb="select"] > div,
      [data-testid="stDateInput"] div[data-baseweb="input"] {{ background: {CARD_BACKGROUND} !important; border: 1px solid {BORDER_COLOR} !important; border-radius: 9px !important; box-shadow: none !important; color: {TEXT_PRIMARY} !important; }}
      [data-testid="stSelectbox"] div[data-baseweb="select"] > div,
      [data-testid="stSelectbox"] [role="combobox"],
      [data-testid="stSelectbox"] [aria-haspopup="listbox"] {{ display: flex !important; align-items: center !important; min-height: var(--filter-control-height) !important; }}
      [data-testid="stDateInput"] input {{ height: calc(var(--filter-control-height) - 2px) !important; min-height: calc(var(--filter-control-height) - 2px) !important; }}
      [data-testid="stSelectbox"] [role="combobox"], [data-testid="stSelectbox"] [aria-haspopup="listbox"] {{ border: 0 !important; box-shadow: none !important; }}
      [data-testid="stSelectbox"] div[data-baseweb="select"] > div:hover,
      [data-testid="stDateInput"] div[data-baseweb="input"]:hover {{ border-color: #5F7EA1 !important; }}
      [data-testid="stSelectbox"] div[data-baseweb="select"] > div:focus-within,
      [data-testid="stDateInput"] div[data-baseweb="input"]:focus-within {{ border-color: {BORDER_FOCUS} !important; box-shadow: 0 0 0 2px rgba(34, 211, 238, .16) !important; }}
      [data-testid="stSelectbox"] input, [data-testid="stSelectbox"] [role="combobox"], [data-testid="stDateInput"] input {{ color: {TEXT_PRIMARY} !important; font-size: 1.05rem !important; line-height: 1.35 !important; }}
      [data-baseweb="select"] svg {{ fill: {CYAN}; }}
      /* Streamlit renders filter menus and calendars in a body-level portal. Keep
         those popups aligned with the dashboard instead of inheriting the black theme. */
      [data-baseweb="popover"] [role="listbox"],
      [data-baseweb="popover"] [role="listbox"] > div,
      [data-baseweb="popover"] [data-baseweb="menu"],
      [data-testid="stSelectboxVirtualDropdown"],
      [data-testid="stSelectboxVirtualDropdown"] [role="listbox"],
      [data-baseweb="calendar"],
      [data-testid="stPopoverBody"] {{
        background: #FFFFFF !important;
        border-color: #C7D2FE !important;
        color: #172554 !important;
      }}
      [data-baseweb="popover"]:has([role="listbox"]) > div,
      [data-baseweb="popover"]:has([data-baseweb="calendar"]) > div {{
        background: #FFFFFF !important;
        border: 1px solid #C7D2FE !important;
        border-radius: 10px !important;
        box-shadow: 0 14px 34px rgba(49, 46, 129, .18) !important;
        overflow: hidden !important;
      }}
      [data-testid="stSelectboxVirtualDropdown"] {{
        border: 1px solid #C7D2FE !important;
        border-radius: 10px !important;
        box-shadow: 0 14px 34px rgba(49, 46, 129, .18) !important;
        overflow: hidden !important;
      }}
      [data-baseweb="popover"] [role="option"],
      [data-baseweb="popover"] [role="option"] *,
      [data-testid="stSelectboxVirtualDropdown"] [role="option"],
      [data-testid="stSelectboxVirtualDropdown"] [role="option"] *,
      [data-baseweb="calendar"] * {{ color: #172554 !important; }}
      [data-baseweb="popover"] [role="option"],
      [data-testid="stSelectboxVirtualDropdown"] [role="option"] {{ background: #FFFFFF !important; }}
      [data-baseweb="popover"] [role="option"]:hover,
      [data-baseweb="popover"] [role="option"][aria-selected="true"],
      [data-testid="stSelectboxVirtualDropdown"] [role="option"]:hover,
      [data-testid="stSelectboxVirtualDropdown"] [role="option"][aria-selected="true"] {{
        background: #EEF2FF !important;
        color: #312E81 !important;
      }}
      [data-baseweb="calendar"] button:hover {{ background: #EEF2FF !important; }}
      [data-baseweb="calendar"] [aria-selected="true"],
      [data-baseweb="calendar"] [aria-selected="true"] * {{
        background: #4F46E5 !important;
        color: #FFFFFF !important;
      }}
      [data-baseweb="calendar"] [aria-disabled="true"],
      [data-baseweb="calendar"] [aria-disabled="true"] * {{ color: #94A3B8 !important; }}
      .stButton > button {{ background: {CARD_BACKGROUND}; border: 1px solid {BORDER_COLOR}; color: {TEXT_PRIMARY}; border-radius: 9px; font-size: 1rem; min-height: 39px; box-shadow: none; }}
      .stButton > button:hover {{ background: #123353; border-color: {BORDER_FOCUS}; color: white; }}
      [data-testid="stCaptionContainer"], .stCaption {{ color: {TEXT_MUTED}; font-size: .95rem; }}
      hr {{ border-color: #293F5B; }}
      @media (max-width: 800px) {{
        html {{ font-size: 16px !important; }}
        :root {{ --filter-control-height: 44px; }}
        .block-container, .stMainBlockContainer, [data-testid="stMainBlockContainer"], section[data-testid="stMain"] .block-container {{ width: 94% !important; max-width: 94% !important; padding-top: 1rem; }}
        .hero-title {{ font-size: 1.75rem; }}
        .status-ribbon {{ font-size: .9rem; }}
      }}
    </style>
    """, unsafe_allow_html=True,
)


def section_heading(title: str, subtitle: str) -> None:
    st.markdown(f'<p class="card-title">{title}</p><p class="card-subtitle">{subtitle}</p>', unsafe_allow_html=True)


def apply_chart_theme(chart) -> None:
    """Apply the selected midnight visual theme without changing chart data or structure."""
    chart.update_layout(
        paper_bgcolor=CARD_BACKGROUND,
        plot_bgcolor=CARD_BACKGROUND,
        font=dict(color=TEXT_PRIMARY, family="Segoe UI, Arial, sans-serif", size=15),
        hoverlabel=dict(bgcolor="#172940", font_color=TEXT_PRIMARY, font_size=14),
    )
    chart.update_xaxes(
        color=TEXT_PRIMARY,
        gridcolor=GRID,
        zerolinecolor=GRID,
        tickfont=dict(color="#FFFFFF", family="Segoe UI Semibold, Arial, sans-serif", size=15),
        title_font=dict(color="#FFFFFF", family="Segoe UI Semibold, Arial, sans-serif", size=17),
        title_standoff=18,
        ticklabelstandoff=9,
        ticks="outside",
        ticklen=6,
        tickwidth=1.5,
        tickcolor="#9FB5D1",
        showline=True,
        linecolor="#6E87A8",
        automargin=True,
    )
    chart.update_yaxes(
        color=TEXT_PRIMARY,
        gridcolor=GRID,
        zerolinecolor=GRID,
        tickfont=dict(color="#FFFFFF", family="Segoe UI Semibold, Arial, sans-serif", size=15),
        title_font=dict(color="#FFFFFF", family="Segoe UI Semibold, Arial, sans-serif", size=17),
        title_standoff=18,
        ticklabelstandoff=9,
        ticks="outside",
        ticklen=6,
        tickwidth=1.5,
        tickcolor="#9FB5D1",
        showline=True,
        linecolor="#6E87A8",
        automargin=True,
    )


@st.fragment(run_every=1)
def watch_for_dashboard_refresh(database_path: str, backup_refresh_seconds: int) -> None:
    """Rerun this browser session after a committed data change or backup interval."""
    resolved_path = Path(database_path)
    session_suffix = str(resolved_path.resolve())
    initialized_key = f"dashboard_signal_initialized:{session_suffix}"
    revision_key = f"dashboard_signal_revision:{session_suffix}"
    last_render_key = f"dashboard_last_full_render:{session_suffix}"
    current_revision = read_refresh_signal(resolved_path)

    if not st.session_state.get(initialized_key, False):
        st.session_state[initialized_key] = True
        st.session_state[revision_key] = current_revision
        return

    signal_changed = current_revision != st.session_state.get(revision_key)
    last_full_render = st.session_state.get(last_render_key, time.monotonic())
    backup_refresh_due = time.monotonic() - last_full_render >= backup_refresh_seconds
    if signal_changed or backup_refresh_due:
        # Record the observed revision before rerunning to prevent a refresh loop.
        st.session_state[revision_key] = current_revision
        st.rerun()


def render_dashboard() -> None:
    settings = get_settings()
    # Streamlit can retain an older Settings class during a hot reload.  Use a
    # safe default so the dashboard stays available while its process reloads.
    mail_source = getattr(settings, "mail_source", "outlook_desktop")
    initialize_database(settings.database_path)
    total_in_database = count_emails(settings.database_path)
    logo_column, title_column, refresh_column, next_column = st.columns(
        [1.05, 4.45, 1, 1.45], vertical_alignment="center"
    )
    with logo_column:
        st.markdown(
            f'<div class="zeta-logo-plate"><img src="data:image/svg+xml;base64,{ZETA_LOGO_BASE64}" alt="Zeta Global"></div>',
            unsafe_allow_html=True,
        )
    with title_column:
        st.markdown(
            '<h1 class="hero-title">Email Analytics Dashboard</h1>'
            '<p class="hero-copy">Dashboard 2 of 2 · Track Inbox volume, sender activity, and engagement over time.</p>',
            unsafe_allow_html=True,
        )
    with refresh_column:
        if st.button("Refresh now", width="stretch"):
            st.rerun()
    with next_column:
        if st.button("Next dashboard", icon=":material/arrow_forward:", width="stretch", key="overview_next"):
            st.query_params.clear()
            st.rerun()
    last_sync = get_state(settings.database_path, "last_successful_sync_at")
    connection_state = "Classic Outlook Desktop"
    privacy_state = "Body previews stored" if settings.store_body_preview else "Metadata-only storage"
    connection_color = "#17a36b" if mail_source == "outlook_desktop" else "#a35f00"
    st.markdown(
        '<div class="status-ribbon">'
        f'<span class="status-item"><span class="status-dot" style="background:{connection_color}"></span>{connection_state}</span>'
        f'<span class="status-item"><span class="status-dot" style="background:#8f00ff"></span>{privacy_state}</span>'
        '<span class="status-item"><span class="status-dot" style="background:#2384d6"></span>Email-triggered refresh · 5 min backup</span>'
        f'<span class="status-item">Last sync: {last_sync or "not connected"}</span>'
        '</div>',
        unsafe_allow_html=True,
    )

    if total_in_database == 0:
        st.markdown(
            '<div class="empty-state"><strong>Your local database is ready.</strong><br>'
            'Open classic Outlook to import Inbox records, or add demo data to preview the charts.</div>',
            unsafe_allow_html=True,
        )
        st.code(".\\.venv\\Scripts\\python.exe -m email_analytics.seed_demo", language="powershell")
        return

    bounds = get_date_bounds(settings.database_path)
    assert bounds is not None
    today = date.today()
    default_start = today - timedelta(days=1)
    default_end = today
    requested_dates = parse_requested_date_range(
        st.query_params.get("rio_start"),
        st.query_params.get("rio_end"),
    )
    selectable_start = min(bounds[0], default_start, *(requested_dates or ()))
    selectable_end = max(bounds[1], default_end, *(requested_dates or ()))
    if requested_dates:
        request_signature = (
            requested_dates,
            str(st.query_params.get("rio_date_token") or ""),
        )
        if st.session_state.get("overview_rio_date_signature") != request_signature:
            st.session_state["overview_date_range"] = requested_dates
            st.session_state["overview_rio_date_signature"] = request_signature
    selected_dates = st.session_state.get(
        "overview_date_range",
        (default_start, default_end),
    )
    start_date: date | None = None
    end_date: date | None = None
    if isinstance(selected_dates, tuple) and len(selected_dates) == 2:
        start_date, end_date = selected_dates
    elif isinstance(selected_dates, date):
        start_date = end_date = selected_dates

    sender_options = [
        "All senders",
        *get_senders(settings.database_path, start_date, end_date),
    ]
    if st.session_state.get("overview_sender_filter") not in sender_options:
        st.session_state.overview_sender_filter = "All senders"

    filter_left, filter_center, filter_right = st.columns([1.85, 0.7, 1.85], vertical_alignment="bottom")
    with filter_left:
        st.markdown('<div class="filter-heading"><strong>From</strong></div>', unsafe_allow_html=True)
        selected_sender = st.selectbox(
            "From",
            sender_options,
            key="overview_sender_filter",
            label_visibility="collapsed",
        )
    with filter_center:
        total_metric = st.empty()
    with filter_right:
        st.markdown('<div class="filter-heading"><strong>Date range</strong></div>', unsafe_allow_html=True)
        selected_dates = st.date_input(
            "Date range",
            value=(default_start, default_end),
            min_value=selectable_start,
            max_value=selectable_end,
            key="overview_date_range",
            label_visibility="collapsed",
        )

    start_date: date | None = None
    end_date: date | None = None
    if isinstance(selected_dates, tuple) and len(selected_dates) == 2:
        start_date, end_date = selected_dates
    elif isinstance(selected_dates, date):
        start_date = end_date = selected_dates
    records = get_dashboard_data(settings.database_path, None if selected_sender == "All senders" else selected_sender, start_date, end_date)
    total_metric.metric(
        "Total Emails",
        f"{len(records):,}",
        help="Emails matching the selected sender and date filters",
    )
    if not records:
        st.warning("No emails match the selected filters.")
        return

    frame = pd.DataFrame(records)
    # Streamlit can hot-reload this file while retaining an older imported
    # database module. Keep the table renderable until the process restarts
    # and begins returning the new Outlook action identifier columns.
    for action_column in ("outlook_entry_id", "outlook_store_id"):
        if action_column not in frame.columns:
            frame[action_column] = None
    frame["received_at"] = pd.to_datetime(frame["received_at"], utc=True, errors="coerce")
    frame["sender"] = frame["sender_name"].where(frame["sender_name"].str.len() > 0, frame["sender_email"])
    frame["date"] = frame["received_at"].dt.date
    if len(frame) != total_in_database:
        st.caption(f"Showing {len(frame):,} of {total_in_database:,} locally stored emails after filters.")

    left, right = st.columns(2, gap="large")
    with left:
        with st.container(border=True):
            section_heading("Emails by Sender", "Breakdown of emails received per sender")
            sender_counts = frame.groupby("sender", as_index=False).size().sort_values("size", ascending=False).head(15)
            chart = px.bar(sender_counts, x="sender", y="size", text="size", color_discrete_sequence=[PURPLE])
            chart.update_layout(height=400, margin=dict(l=100, r=35, t=25, b=130), showlegend=False, xaxis_title="Sender", yaxis_title="Email count")
            apply_chart_theme(chart)
            chart.update_traces(texttemplate="%{text:,.0f}", textposition="outside", textfont=dict(color="#FFFFFF", size=15), cliponaxis=False)
            chart.update_xaxes(tickangle=-35, showgrid=False)
            chart.update_yaxes(tickformat=",d", nticks=6)
            st.plotly_chart(chart, width="stretch", config={"displayModeBar": False})
    with right:
        with st.container(border=True):
            section_heading(
                "Email Details",
                "All emails matching the selected sender and calendar dates",
            )
            details = frame[["sender_email", "received_at", "subject", "has_attachments", "importance"]].copy()
            details["received_at"] = details["received_at"].dt.tz_convert("Asia/Kolkata").dt.strftime("%d %b %Y, %I:%M %p")
            details["has_attachments"] = details["has_attachments"].map({1: "Yes", 0: "No"})
            details.columns = ["From", "Received time", "Subject", "Attachments", "Importance"]
            visible_details = details.reset_index(drop=True)
            action_details = frame[
                ["sender_email", "subject", "outlook_entry_id", "outlook_store_id"]
            ].reset_index(drop=True)
            st.caption(f"{len(visible_details):,} matching email(s) · Select one row for Outlook actions")
            styled_details = visible_details.style.set_properties(**{"font-size": "14px"}).set_table_styles(
                [{"selector": "th", "props": [("font-size", "14px"), ("font-weight", "700")]}]
            )
            selection = st.dataframe(
                styled_details,
                width="stretch",
                hide_index=True,
                height=255,
                row_height=31,
                key="email_details_table",
                on_select="rerun",
                selection_mode="single-row",
            )
            selected_rows = list(selection.selection.rows)
            if selected_rows and selected_rows[0] < len(action_details):
                selected = action_details.iloc[selected_rows[0]]
                entry_id = selected["outlook_entry_id"]
                store_id = selected["outlook_store_id"]
                entry_id = str(entry_id) if pd.notna(entry_id) and entry_id else ""
                store_id = str(store_id) if pd.notna(store_id) and store_id else None
                st.caption(f'Selected: {selected["subject"] or "(no subject)"} · {selected["sender_email"]}')
                open_column, reply_column = st.columns(2)
                with open_column:
                    open_message = st.button(
                        "Open in Outlook",
                        icon=":material/open_in_new:",
                        width="stretch",
                        disabled=not entry_id,
                    )
                with reply_column:
                    reply_message = st.button(
                        "Reply in Outlook",
                        icon=":material/reply:",
                        width="stretch",
                        disabled=not entry_id,
                    )

                if not entry_id:
                    st.warning("Restart the Outlook listener once to enable actions for this existing email.")
                elif open_message or reply_message:
                    try:
                        display_outlook_item(entry_id, store_id, reply=reply_message)
                    except OutlookDesktopError as error:
                        st.error(str(error))
                    else:
                        st.toast("Reply draft opened in Outlook." if reply_message else "Email opened in Outlook.")

    lower_left, lower_right = st.columns(2, gap="large")
    daily_counts = frame.groupby("date", as_index=False).size().rename(columns={"size": "emails"})
    with lower_left:
        with st.container(border=True):
            section_heading("Email Trends Over Time", "Daily Inbox activity based on selected filters")
            chart = px.line(daily_counts, x="date", y="emails", markers=True, color_discrete_sequence=[PURPLE])
            chart.update_layout(height=375, margin=dict(l=100, r=35, t=25, b=85), showlegend=False, xaxis_title="Date", yaxis_title="Email count")
            apply_chart_theme(chart)
            chart.update_traces(line=dict(width=3), marker=dict(size=8))
            chart.update_xaxes(tickformat="%d %b", nticks=7)
            chart.update_yaxes(tickformat=",d", nticks=6)
            st.plotly_chart(chart, width="stretch", config={"displayModeBar": False})
    with lower_right:
        with st.container(border=True):
            section_heading("Daily Email Volume", "Number of emails received per day")
            latest_days = daily_counts.sort_values("date", ascending=False).head(12).sort_values("date")
            chart = px.bar(latest_days, x="emails", y="date", orientation="h", text="emails", color_discrete_sequence=[PURPLE])
            chart.update_layout(height=375, margin=dict(l=110, r=60, t=25, b=85), showlegend=False, xaxis_title="Email count", yaxis_title="Date")
            apply_chart_theme(chart)
            chart.update_traces(texttemplate="%{text:,.0f}", textposition="outside", textfont=dict(color="#FFFFFF", size=15), cliponaxis=False)
            chart.update_xaxes(tickformat=",d", nticks=6)
            chart.update_yaxes(type="category")
            st.plotly_chart(chart, width="stretch", config={"displayModeBar": False})
    st.divider()
    st.caption(f"Local database: `{settings.database_path}` · Dashboard time zone: Asia/Kolkata")


def main() -> None:
    settings = get_settings()
    session_suffix = str(settings.database_path.resolve())
    st.session_state[f"dashboard_last_full_render:{session_suffix}"] = time.monotonic()
    watch_for_dashboard_refresh(
        str(settings.database_path),
        getattr(settings, "dashboard_backup_refresh_seconds", 300),
    )
    if st.query_params.get("view") == "overview":
        render_dashboard()
    else:
        from email_analytics.advanced_dashboard import render_advanced_dashboard

        render_advanced_dashboard()


if __name__ == "__main__":
    main()
