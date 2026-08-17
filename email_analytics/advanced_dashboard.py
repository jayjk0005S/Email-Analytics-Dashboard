from __future__ import annotations

import base64
import html
from datetime import date, timedelta
from urllib.parse import quote

import pandas as pd
import plotly.express as px
import streamlit as st

from .config import PROJECT_ROOT, get_settings
from .date_filters import parse_requested_date_range
from .database import get_dashboard_data, get_date_bounds, initialize_database
from .outlook_desktop import OutlookDesktopError, display_outlook_item, get_outlook_item_details


PURPLE = "#7C3AED"
CYAN = "#06B6D4"
CARD_BACKGROUND = "#FFFFFF"
TEXT_PRIMARY = "#172554"
TEXT_MUTED = "#64748B"
BORDER_COLOR = "#D8DEFF"
GRID = "#E7EAF8"
ZETA_LOGO_PATH = PROJECT_ROOT / "assets" / "zeta-logo-primary.svg"
ZETA_LOGO_BASE64 = base64.b64encode(ZETA_LOGO_PATH.read_bytes()).decode("ascii")


def _format_size(size: int) -> str:
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.1f} MB"


def _email_selection_url(message_id: str, selected_message_id: str) -> str:
    """Clicking the selected row clears it; clicking another row selects it."""
    return "?" if message_id and message_id == selected_message_id else f"?selected_email={quote(message_id, safe='')}"


def _date_selection(value: object) -> tuple[date | None, date | None]:
    if isinstance(value, tuple) and len(value) == 2:
        return value[0], value[1]
    if isinstance(value, date):
        return value, value
    return None, None


def _reset_advanced_filters() -> None:
    for key in (
        "advanced_search",
        "advanced_sender",
        "advanced_senders",
        "advanced_dates",
        "advanced_day",
        "advanced_importance",
        "advanced_attachments",
        "advanced_read_status",
    ):
        st.session_state.pop(key, None)


def _render_header() -> None:
    with st.container(key="advanced_header"):
        logo_column, title_column, shortcut_column, save_column, saved_column, next_column = st.columns(
            [0.94, 3.62, 0.82, 1.05, 1.15, 0.55],
            vertical_alignment="center",
        )
        with logo_column:
            st.markdown(
                f'<div class="advanced-zeta-logo"><img src="data:image/svg+xml;base64,{ZETA_LOGO_BASE64}" '
                'alt="Zeta Global"></div>',
                unsafe_allow_html=True,
            )
        with title_column:
            st.markdown(
                '<div class="advanced-title-line">'
                '<div class="advanced-dashboard-name">Advanced Email Explorer</div>'
                '<div class="advanced-dashboard-badge">Dashboard 2 of 3</div>'
                '</div>',
                unsafe_allow_html=True,
            )
        with shortcut_column:
            st.markdown('<div class="advanced-shortcut"><kbd>/</kbd><span>to search</span></div>', unsafe_allow_html=True)
        with save_column:
            if st.button("Save search", icon=":material/star:", width="stretch", key="advanced_save_search"):
                st.toast("Search view saved for this session.")
        with saved_column:
            with st.popover("Saved searches", icon=":material/bookmark:", width="stretch"):
                st.markdown("**Saved searches**")
                st.caption("Your saved filter combinations will appear here.")
        with next_column:
            if st.button(
                "",
                icon=":material/arrow_forward:",
                width="stretch",
                key="advanced_next",
                help="Open Dashboard 2",
            ):
                st.query_params["view"] = "overview"
                st.rerun()


def _apply_chart_theme(chart) -> None:
    chart.update_layout(
        paper_bgcolor=CARD_BACKGROUND,
        plot_bgcolor=CARD_BACKGROUND,
        font=dict(color=TEXT_PRIMARY, family="Segoe UI, Arial, sans-serif", size=15),
        hoverlabel=dict(bgcolor="#EEF2FF", font_color=TEXT_PRIMARY, font_size=14),
    )
    chart.update_xaxes(
        color=TEXT_PRIMARY,
        gridcolor=GRID,
        zerolinecolor=GRID,
        tickfont=dict(color="#172554", family="Segoe UI Semibold, Arial, sans-serif", size=15),
        title_font=dict(color="#172554", family="Segoe UI Semibold, Arial, sans-serif", size=17),
        title_standoff=18,
        ticklabelstandoff=9,
        ticks="outside",
        ticklen=6,
        tickwidth=1.5,
        tickcolor="#64748B",
        showline=True,
        linecolor="#94A3B8",
        automargin=True,
    )
    chart.update_yaxes(
        color=TEXT_PRIMARY,
        gridcolor=GRID,
        zerolinecolor=GRID,
        tickfont=dict(color="#172554", family="Segoe UI Semibold, Arial, sans-serif", size=15),
        title_font=dict(color="#172554", family="Segoe UI Semibold, Arial, sans-serif", size=17),
        title_standoff=18,
        ticklabelstandoff=9,
        ticks="outside",
        ticklen=6,
        tickwidth=1.5,
        tickcolor="#64748B",
        showline=True,
        linecolor="#94A3B8",
        automargin=True,
    )


def _section_heading(title: str, subtitle: str) -> None:
    st.markdown(
        f'<p class="card-title">{title}</p><p class="card-subtitle">{subtitle}</p>',
        unsafe_allow_html=True,
    )


def _render_kpi_cards(frame: pd.DataFrame) -> None:
    total = len(frame)
    unique_senders = frame["sender_email"].nunique() if total else 0
    attachment_count = int(frame["has_attachments"].astype(bool).sum()) if total else 0
    important_count = int(frame["importance"].fillna("").str.lower().eq("high").sum()) if total else 0
    attachment_share = round(attachment_count / total * 100) if total else 0
    important_share = round(important_count / total * 100) if total else 0

    icons = {
        "mail": '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 6.5A2.5 2.5 0 0 1 5.5 4h13A2.5 2.5 0 0 1 21 6.5v11a2.5 2.5 0 0 1-2.5 2.5h-13A2.5 2.5 0 0 1 3 17.5v-11Zm2 .3 7 5.1 7-5.1"/></svg>',
        "people": '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="9" cy="8" r="3"/><path d="M3.5 19c.4-4 2.2-6 5.5-6s5.1 2 5.5 6M16 5.5a3 3 0 0 1 0 5.9M16 13c3 0 4.6 2 5 5"/></svg>',
        "clip": '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m8.5 12.5 6.7-6.7a3.5 3.5 0 1 1 5 5l-8.4 8.4a5 5 0 0 1-7-7l8-8M8 16l8.2-8.2"/></svg>',
        "flag": '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 21V4m0 1h11l-1.5 3L16 11H5"/></svg>',
    }
    cards = (
        ("Matching emails", total, "Total results", "#3157F5", icons["mail"], "0,26 12,23 23,25 35,16 47,19 59,10 71,15 84,5 100,9"),
        ("Unique senders", unique_senders, "Distinct senders", "#8B5CF6", icons["people"], "0,24 12,19 24,22 36,10 48,13 60,8 72,20 84,7 100,11"),
        ("With attachments", attachment_count, f"{attachment_share}% of results", "#06B6D4", icons["clip"], "0,25 12,21 24,24 36,14 48,17 60,6 72,20 84,10 100,4"),
        ("High importance", important_count, f"{important_share}% of results", "#D946EF", icons["flag"], "0,25 12,23 24,16 36,20 48,8 60,12 72,22 84,11 100,5"),
    )
    markup = ['<div class="electric-kpi-grid">']
    for label, value, subtitle, accent, icon, points in cards:
        markup.append(
            f'<div class="electric-kpi-card" style="--accent:{accent}">'
            f'<div class="electric-kpi-icon">{icon}</div>'
            '<div class="electric-kpi-copy">'
            f'<div class="electric-kpi-label">{label}</div>'
            f'<div class="electric-kpi-value">{value:,}</div>'
            f'<div class="electric-kpi-subtitle">{subtitle}</div>'
            '</div>'
            f'<svg class="electric-sparkline" viewBox="0 0 100 30" preserveAspectRatio="none"><polyline points="{points}"/></svg>'
            '</div>'
        )
    markup.append('</div>')
    st.markdown("".join(markup), unsafe_allow_html=True)


def _render_email_table(visible: pd.DataFrame, frame: pd.DataFrame, visible_columns: list[str]) -> int | None:
    rows_per_page = 8
    page_count = max(1, (len(visible) + rows_per_page - 1) // rows_per_page)
    page_key = "advanced_table_page"
    current_page = min(max(int(st.session_state.get(page_key, 0)), 0), page_count - 1)
    st.session_state[page_key] = current_page
    start = current_page * rows_per_page
    end = min(start + rows_per_page, len(visible))
    selected_message_id = str(st.query_params.get("selected_email") or "")

    markup = ['<div class="electric-table-scroll"><table class="electric-data-table"><thead><tr>']
    markup.append('<th class="electric-selector-column"><span class="electric-check"></span></th>')
    for column in visible_columns:
        column_class = column.lower().replace(" ", "-")
        markup.append(f'<th class="electric-column-{column_class}">{html.escape(column)}</th>')
    markup.append('</tr></thead><tbody>')

    page_rows = visible.iloc[start:end]
    for display_index, row in page_rows.iterrows():
        message_id = str(frame.iloc[display_index].get("message_id") or "")
        is_selected = bool(message_id and message_id == selected_message_id)
        selected_class = " electric-row-selected" if is_selected else ""
        row_url = _email_selection_url(message_id, selected_message_id)
        selection_label = "Clear email selection" if is_selected else "Select email"
        markup.append(f'<tr class="electric-email-row{selected_class}">')
        markup.append(
            f'<td class="electric-selector-column"><a href="{row_url}" target="_self" aria-label="{selection_label}">'
            f'<span class="electric-check">{"✓" if selected_class else ""}</span></a></td>'
        )
        for column in visible_columns:
            raw_value = str(row[column])
            safe_value = html.escape(raw_value)
            cell_class = f'electric-cell-{column.lower().replace(" ", "-")}'
            if column == "Importance":
                importance_class = " electric-importance-high" if raw_value == "High" else " electric-importance-normal"
                safe_value = f'<span class="electric-badge{importance_class}">{safe_value}</span>'
            elif column == "Attachments" and raw_value == "Yes":
                safe_value = f'<span class="electric-attachment">↗ {safe_value}</span>'
            elif column == "Read status":
                status_class = " electric-status-unread" if raw_value == "Not Read" else " electric-status-read"
                status_icon = "●" if raw_value == "Not Read" else "✓"
                safe_value = f'<span class="electric-status{status_class}">{status_icon} {safe_value}</span>'
            markup.append(
                f'<td class="{cell_class}" title="{html.escape(raw_value, quote=True)}">'
                f'<a href="{row_url}" target="_self">{safe_value}</a></td>'
            )
        markup.append('</tr>')
    markup.append('</tbody></table></div>')
    st.markdown("".join(markup), unsafe_allow_html=True)

    footer_left, footer_spacer, previous_column, page_column, next_column = st.columns(
        [2.4, 4.2, 0.55, 0.48, 0.55],
        vertical_alignment="center",
    )
    with footer_left:
        st.caption(f"Showing {start + 1:,}–{end:,} of {len(visible):,} results")
    with previous_column:
        if st.button(
            "",
            icon=":material/chevron_left:",
            key="advanced_table_previous",
            help="Previous page",
            disabled=current_page == 0,
            width="stretch",
        ):
            st.session_state[page_key] = current_page - 1
            st.rerun()
    with page_column:
        st.markdown(f'<div class="electric-page-number">{current_page + 1}</div>', unsafe_allow_html=True)
    with next_column:
        if st.button(
            "",
            icon=":material/chevron_right:",
            key="advanced_table_next",
            help="Next page",
            disabled=current_page >= page_count - 1,
            width="stretch",
        ):
            st.session_state[page_key] = current_page + 1
            st.rerun()

    if selected_message_id:
        matches = frame.index[frame["message_id"].fillna("").astype(str).eq(selected_message_id)]
        if len(matches):
            return int(matches[0])
    return None


def _render_visual_overview(frame: pd.DataFrame) -> None:
    local_dates = frame["received_at"].dt.tz_convert("Asia/Kolkata").dt.date
    daily_counts = (
        frame.assign(date=local_dates)
        .groupby("date", as_index=False)
        .size()
        .rename(columns={"size": "emails"})
    )
    sender_counts = (
        frame.groupby("sender", as_index=False)
        .size()
        .rename(columns={"size": "emails"})
        .nlargest(8, "emails")
        .sort_values("emails")
    )

    left, right = st.columns(2, gap="large")
    with left:
        with st.container(border=True):
            _section_heading("Filtered Email Activity", "Daily volume for the filters currently selected")
            activity_chart = px.line(
                daily_counts,
                x="date",
                y="emails",
                markers=True,
                color_discrete_sequence=[CYAN],
            )
            activity_chart.update_traces(
                line=dict(width=4),
                marker=dict(size=9, color=PURPLE, line=dict(color=CYAN, width=2)),
                fill="tozeroy",
                fillcolor="rgba(34, 211, 238, 0.12)",
            )
            activity_chart.update_layout(
                height=380,
                margin=dict(l=100, r=35, t=24, b=85),
                showlegend=False,
                xaxis_title="Date",
                yaxis_title="Email count",
            )
            _apply_chart_theme(activity_chart)
            activity_chart.update_xaxes(tickformat="%d %b", nticks=7)
            activity_chart.update_yaxes(tickformat=",d", nticks=6)
            st.plotly_chart(activity_chart, width="stretch", config={"displayModeBar": False})
    with right:
        with st.container(border=True):
            _section_heading("Top Senders", "Most active senders in the filtered results")
            sender_chart = px.bar(
                sender_counts,
                x="emails",
                y="sender",
                orientation="h",
                text="emails",
                color="emails",
                color_continuous_scale=[CYAN, PURPLE],
            )
            sender_chart.update_traces(
                texttemplate="%{text:,.0f}",
                textposition="outside",
                textfont=dict(color="#172554", size=15),
                cliponaxis=False,
            )
            sender_chart.update_layout(
                height=380,
                margin=dict(l=165, r=65, t=24, b=85),
                coloraxis_showscale=False,
                xaxis_title="Email count",
                yaxis_title="Sender",
            )
            _apply_chart_theme(sender_chart)
            sender_chart.update_xaxes(tickformat=",d", nticks=6)
            sender_chart.update_yaxes(showgrid=False)
            st.plotly_chart(sender_chart, width="stretch", config={"displayModeBar": False})


def _render_detail_panel(selected: pd.Series) -> None:
    entry_id_value = selected.get("outlook_entry_id")
    store_id_value = selected.get("outlook_store_id")
    entry_id = str(entry_id_value) if pd.notna(entry_id_value) and entry_id_value else ""
    store_id = str(store_id_value) if pd.notna(store_id_value) and store_id_value else None

    st.markdown('<p class="card-title">Selected Email</p>', unsafe_allow_html=True)
    if entry_id:
        try:
            details = get_outlook_item_details(entry_id, store_id)
        except OutlookDesktopError as error:
            st.warning(str(error))
            details = None
    else:
        details = None
        st.info("This stored email has no Outlook link yet. Restart the listener once to add it.")

    if details:
        received = pd.to_datetime(details.get("received_at"), utc=True, errors="coerce")
        received_text = (
            received.tz_convert("Asia/Kolkata").strftime("%d %b %Y, %I:%M %p")
            if pd.notna(received)
            else "Unknown"
        )
        st.markdown(f"### {details['subject'] or '(no subject)'}")
        left, right = st.columns(2)
        with left:
            st.markdown(f"**From**  \n{details['sender_name'] or details['sender_email']}  \n{details['sender_email']}")
            st.markdown(f"**To**  \n{details['to'] or 'Not shown'}")
            if details["cc"]:
                st.markdown(f"**CC**  \n{details['cc']}")
        with right:
            st.markdown(f"**Received**  \n{received_text}")
            st.markdown(
                f"**Status / importance**  \n{'Not Read' if details['unread'] else 'Read'} · "
                f"{details['importance'].title()}"
            )
            st.markdown(f"**Message size**  \n{_format_size(details['size'])}")
        if details["categories"]:
            st.markdown(f"**Categories**  \n{details['categories']}")
        if details["attachments"]:
            attachment_names = ", ".join(
                f"{attachment['name']} ({_format_size(attachment['size'])})"
                for attachment in details["attachments"]
            )
            st.markdown(f"**Attachments**  \n{attachment_names}")
        else:
            st.markdown("**Attachments**  \nNone")
        st.markdown("**Body preview (read live from Outlook)**")
        st.text_area(
            "Body preview",
            value=details["body_preview"] or "No plain-text body is available.",
            height=220,
            disabled=True,
            label_visibility="collapsed",
        )
    else:
        st.markdown(f"### {selected.get('subject') or '(no subject)'}")
        st.markdown(f"**From**  \n{selected.get('sender_email') or 'Unknown'}")
        stored_preview = str(selected.get("body_preview") or "")
        if stored_preview:
            st.markdown("**Stored preview**")
            st.text_area(
                "Stored preview",
                value=stored_preview,
                height=180,
                disabled=True,
                label_visibility="collapsed",
            )

    open_column, reply_column = st.columns(2)
    with open_column:
        open_message = st.button(
            "Open in Outlook",
            icon=":material/open_in_new:",
            width="stretch",
            disabled=not entry_id,
            key="advanced_open_message",
        )
    with reply_column:
        reply_message = st.button(
            "Reply in Outlook",
            icon=":material/reply:",
            width="stretch",
            disabled=not entry_id,
            key="advanced_reply_message",
        )
    if entry_id and (open_message or reply_message):
        try:
            display_outlook_item(entry_id, store_id, reply=reply_message)
        except OutlookDesktopError as error:
            st.error(str(error))
        else:
            st.toast("Reply draft opened in Outlook." if reply_message else "Email opened in Outlook.")


def render_advanced_dashboard() -> None:
    settings = get_settings()
    initialize_database(settings.database_path)
    st.markdown(
        f"""
        <style>
          :root {{ color-scheme: light; }}
          .stApp {{ background: #F7F8FF !important; color: {TEXT_PRIMARY} !important; }}
          [data-testid="stMainBlockContainer"] {{ padding-top: .85rem !important; }}
          [data-testid="stHeader"] {{ background: transparent !important; }}
          .st-key-advanced_header {{
            background: linear-gradient(105deg, #1535BE 0%, #1924BC 48%, #4C10C7 100%);
            border-radius: 14px; box-shadow: 0 12px 30px rgba(49, 46, 129, .18);
            padding: .9rem 1.05rem; margin-bottom: .9rem;
          }}
          .advanced-zeta-logo {{
            align-items: center; background: white; border-radius: 9px; display: flex;
            box-sizing: border-box; height: 60px; justify-content: center; line-height: 0; overflow: hidden;
            padding: 11px 13px; width: 100%;
            box-shadow: 0 4px 12px rgba(14, 20, 76, .18);
          }}
          .advanced-zeta-logo img {{
            display: block; height: 36px; margin: 0 auto; max-height: 36px; max-width: 118px;
            object-fit: contain; object-position: center; width: 100%;
          }}
          .advanced-title-line {{ align-items: center; display: flex; flex-wrap: nowrap; gap: .7rem; min-width: 0; white-space: nowrap; }}
          .advanced-dashboard-name {{ color: white; font-size: 1.58rem; font-weight: 800; letter-spacing: -.035em; line-height: 1.1; }}
          .advanced-dashboard-badge {{
            background: rgba(255,255,255,.12); border: 1px solid rgba(255,255,255,.2);
            border-radius: 999px; color: #EEF2FF; font-size: .82rem; font-weight: 700; padding: .32rem .64rem;
          }}
          .advanced-shortcut {{ align-items: center; color: #E0E7FF; display: flex; font-size: .8rem; gap: .35rem; justify-content: center; white-space: nowrap; }}
          .advanced-shortcut kbd {{ background: rgba(255,255,255,.12); border: 1px solid rgba(255,255,255,.28); border-radius: 5px; color: white; padding: .18rem .38rem; }}
          .st-key-advanced_header .stButton > button,
          .st-key-advanced_header [data-testid="stPopover"] > button {{
            background: white !important; border: 0 !important; color: #312E81 !important;
            font-size: .83rem !important; font-weight: 750 !important; height: 48px !important; min-height: 48px !important;
            box-shadow: 0 5px 14px rgba(12, 18, 83, .18) !important;
          }}
          .st-key-advanced_header button {{
            background: white !important; border-color: white !important; color: #312E81 !important;
            height: 48px !important; min-height: 48px !important;
          }}
          .st-key-advanced_header [data-testid="stButton"],
          .st-key-advanced_header [data-testid="stPopover"] {{ align-items: center; display: flex; width: 100%; }}
          .st-key-advanced_header [data-testid="stButton"] > button,
          .st-key-advanced_header [data-testid="stPopover"] button {{ width: 100%; }}
          .advanced-section-heading {{ color: {TEXT_PRIMARY}; font-size: 1.2rem; font-weight: 850; margin: .55rem 0 .55rem; }}
          [data-testid="stTextInput"] label, [data-testid="stMultiSelect"] label,
          [data-testid="stSelectbox"] label, [data-testid="stDateInput"] label {{
            color: {TEXT_PRIMARY} !important; font-size: .9rem !important; font-weight: 750 !important;
          }}
          [data-testid="stTextInput"] input, [data-testid="stDateInput"] input {{
            background: white !important; color: {TEXT_PRIMARY} !important; -webkit-text-fill-color: {TEXT_PRIMARY} !important;
          }}
          [data-testid="stTextInput"] input::placeholder {{ color: #94A3B8 !important; opacity: 1 !important; }}
          [data-testid="stTextInput"] div[data-baseweb="input"],
          [data-testid="stDateInput"] div[data-baseweb="input"],
          [data-testid="stSelectbox"] div[data-baseweb="select"] > div,
          [data-testid="stMultiSelect"] div[data-baseweb="select"] > div {{
            background: white !important; border: 1px solid {BORDER_COLOR} !important;
            border-radius: 9px !important; color: {TEXT_PRIMARY} !important; box-shadow: 0 2px 8px rgba(49,46,129,.035) !important;
          }}
          [data-testid="stSelectbox"] [data-baseweb="select"] *, [data-testid="stMultiSelect"] span {{ color: {TEXT_PRIMARY} !important; }}
          [data-testid="stSelectbox"] [role="combobox"],
          [data-testid="stSelectbox"] [aria-haspopup="listbox"],
          [data-testid="stSelectbox"] div[data-baseweb="select"],
          [data-testid="stSelectbox"] div[data-baseweb="select"] > div {{
            background-color: white !important; color: {TEXT_PRIMARY} !important;
          }}
          [data-testid="stSelectbox"] svg {{ fill: #4338CA !important; color: #4338CA !important; }}
          [data-baseweb="tag"] {{ background: #EEF2FF !important; border: 1px solid #C7D2FE !important; }}
          .stButton > button, [data-testid="stDownloadButton"] > button,
          [data-testid="stPopover"] > button {{
            background: white; border: 1px solid {BORDER_COLOR}; border-radius: 9px;
            color: #3730A3; font-weight: 720; box-shadow: 0 3px 10px rgba(49,46,129,.055);
          }}
          .stButton > button:hover, [data-testid="stDownloadButton"] > button:hover,
          [data-testid="stPopover"] > button:hover {{ background: #EEF2FF; border-color: #818CF8; color: #312E81; }}
          .electric-kpi-grid {{ display: grid; gap: 1rem; grid-template-columns: repeat(4, minmax(0, 1fr)); margin: .75rem 0 1rem; }}
          .electric-kpi-card {{
            align-items: center; background: linear-gradient(135deg, white 55%, color-mix(in srgb, var(--accent) 7%, white));
            border: 1px solid color-mix(in srgb, var(--accent) 35%, #D8DEFF); border-radius: 13px;
            box-shadow: 0 8px 22px rgba(49,46,129,.075); display: flex; min-height: 122px;
            overflow: hidden; padding: 1rem 1rem; position: relative; transition: transform .18s ease, box-shadow .18s ease;
          }}
          .electric-kpi-card:hover {{ box-shadow: 0 13px 28px color-mix(in srgb, var(--accent) 18%, transparent); transform: translateY(-3px); }}
          .electric-kpi-icon {{
            align-items: center; background: linear-gradient(145deg, color-mix(in srgb, var(--accent) 88%, white), var(--accent));
            border-radius: 50%; box-shadow: 0 8px 18px color-mix(in srgb, var(--accent) 25%, transparent);
            color: white; display: flex; flex: 0 0 56px; height: 56px; justify-content: center;
          }}
          .electric-kpi-icon svg {{ fill: none; height: 29px; stroke: currentColor; stroke-linecap: round; stroke-linejoin: round; stroke-width: 1.8; width: 29px; }}
          .electric-kpi-copy {{ margin-left: .8rem; position: relative; z-index: 2; }}
          .electric-kpi-label {{ color: var(--accent); font-size: .9rem; font-weight: 800; }}
          .electric-kpi-value {{ color: {TEXT_PRIMARY}; font-size: 1.75rem; font-weight: 850; line-height: 1.1; margin-top: .18rem; }}
          .electric-kpi-subtitle {{ color: {TEXT_MUTED}; font-size: .78rem; margin-top: .24rem; }}
          .electric-sparkline {{ bottom: 15px; height: 38px; position: absolute; right: 12px; width: 86px; }}
          .electric-sparkline polyline {{ fill: none; stroke: var(--accent); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2; vector-effect: non-scaling-stroke; }}
          .st-key-advanced_table_card > div[data-testid="stVerticalBlockBorderWrapper"] {{
            background: white !important; border: 1px solid #DCE1F5 !important; border-radius: 14px !important;
            box-shadow: 0 10px 26px rgba(49,46,129,.085) !important; min-height: 0 !important; overflow: hidden;
          }}
          .st-key-advanced_table_card > div[data-testid="stVerticalBlockBorderWrapper"] > div {{ padding: .75rem .85rem !important; }}
          .st-key-advanced_table_card button {{ background: white !important; border-color: {BORDER_COLOR} !important; color: #3730A3 !important; }}
          .advanced-table-title {{ align-items: center; display: flex; gap: .55rem; }}
          .advanced-table-title strong {{ color: {TEXT_PRIMARY}; font-size: 1.12rem; }}
          .advanced-table-title span {{ background: #EEF2FF; border-radius: 999px; color: #4338CA; font-size: .76rem; font-weight: 750; padding: .22rem .52rem; }}
          .electric-table-scroll {{
            border: 1px solid #E0E5F5; border-radius: 10px; max-height: 406px; overflow: auto;
            scrollbar-color: #C7D2FE #F8FAFF; scrollbar-width: thin;
          }}
          .electric-data-table {{ background: white; border-collapse: separate; border-spacing: 0; table-layout: fixed; width: 100%; }}
          .electric-data-table th {{
            background: #EEF2FF; border-bottom: 1px solid #D8DEFF; color: #3730A3;
            font-size: .76rem; font-weight: 800; letter-spacing: .02em; padding: .72rem .68rem;
            position: sticky; text-align: left; text-transform: uppercase; top: 0; z-index: 3;
          }}
          .electric-data-table td {{
            background: white; border-bottom: 1px solid #EDF0FA; color: {TEXT_PRIMARY};
            font-size: .82rem; padding: .68rem .68rem; transition: background .14s ease;
          }}
          .electric-data-table tr:last-child td {{ border-bottom: 0; }}
          .electric-data-table tr:hover td {{ background: #F5F7FF; }}
          .electric-data-table tr.electric-row-selected td {{ background: #E9F9FC; }}
          .electric-data-table tr.electric-row-selected td:first-child {{ box-shadow: inset 3px 0 0 #06B6D4; }}
          .electric-data-table th:first-child {{ border-radius: 9px 0 0 0; }}
          .electric-data-table th:last-child {{ border-radius: 0 9px 0 0; }}
          .electric-data-table a {{ color: inherit; display: block; overflow: hidden; text-decoration: none; text-overflow: ellipsis; white-space: nowrap; }}
          .electric-selector-column {{ background: #F8FAFF !important; box-sizing: border-box; text-align: center !important; width: 44px; }}
          .electric-check {{
            align-items: center; border: 1.5px solid #A5B4D4; border-radius: 5px; color: white;
            display: inline-flex; font-size: .7rem; height: 18px; justify-content: center; width: 18px;
          }}
          .electric-row-selected .electric-check {{ background: #06B6D4; border-color: #06B6D4; box-shadow: 0 2px 7px rgba(6,182,212,.28); }}
          .electric-column-received, .electric-cell-received {{ width: 15%; }}
          .electric-column-subject, .electric-cell-subject {{ width: 34%; }}
          .electric-column-from, .electric-cell-from {{ width: 17%; }}
          .electric-column-importance, .electric-cell-importance {{ width: 11%; }}
          .electric-column-attachments, .electric-cell-attachments {{ width: 10%; }}
          .electric-column-read-status, .electric-cell-read-status {{ width: 9%; }}
          .electric-badge, .electric-status {{ border-radius: 999px; display: inline-flex; font-size: .75rem; font-weight: 750; padding: .2rem .48rem; }}
          .electric-importance-high {{ background: #FDF2F8; color: #DB2777; }}
          .electric-importance-normal {{ background: #EEF2FF; color: #4338CA; }}
          .electric-status-read {{ background: #ECFDF5; color: #15803D; }}
          .electric-status-unread {{ background: #EFF6FF; color: #1D4ED8; }}
          .electric-attachment {{ color: #0891B2; font-weight: 750; }}
          .electric-page-number {{
            align-items: center; background: #4F46E5; border-radius: 8px; color: white; display: flex;
            font-size: .82rem; font-weight: 800; height: 38px; justify-content: center;
          }}
          div[data-testid="stDataFrame"] {{ background: white !important; border: 1px solid #E2E5F4 !important; border-radius: 9px !important; }}
          div[data-testid="stDataFrame"] * {{ color: {TEXT_PRIMARY} !important; }}
          [data-testid="stVerticalBlockBorderWrapper"] {{
            background: white; border-color: {BORDER_COLOR} !important; box-shadow: 0 7px 20px rgba(49,46,129,.07) !important; min-height: 0 !important;
          }}
          [data-testid="stTextArea"] textarea:disabled {{ background: #F8FAFF !important; color: {TEXT_PRIMARY} !important; -webkit-text-fill-color: {TEXT_PRIMARY} !important; opacity: 1 !important; }}
          .card-title {{ color: {TEXT_PRIMARY} !important; text-shadow: none !important; }}
          .card-subtitle, [data-testid="stCaptionContainer"] {{ color: {TEXT_MUTED} !important; }}
          [data-testid="stExpander"] {{
            background: linear-gradient(105deg, #173ED4 0%, #3520D1 49%, #7C2BDF 100%) !important;
            border: 0 !important; border-radius: 14px !important; box-shadow: 0 12px 26px rgba(63, 48, 199, .2); margin-top: .85rem;
          }}
          [data-testid="stExpander"] summary {{ color: white !important; font-size: 1.1rem; font-weight: 850; min-height: 72px; }}
          [data-testid="stExpander"] summary p::after {{
            color: #E0E7FF; content: "Explore activity trends, top senders, and email volumes over time.";
            display: block; font-size: .78rem; font-weight: 500; margin-top: .2rem;
          }}
          [data-testid="stExpander"] summary svg {{ fill: white !important; color: white !important; }}
          [data-testid="stExpander"] details[open] > div {{ background: #F8FAFF; border-radius: 0 0 13px 13px; color: {TEXT_PRIMARY}; padding: .8rem; }}
          [data-testid="stAlert"] {{ color: {TEXT_PRIMARY} !important; }}
          @media (max-width: 900px) {{
            .advanced-shortcut {{ display: none; }}
            .advanced-dashboard-name {{ font-size: 1.22rem; }}
            .electric-kpi-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
          }}
        </style>
        """,
        unsafe_allow_html=True,
    )
    _render_header()

    bounds = get_date_bounds(settings.database_path)
    if bounds is None:
        st.markdown(
            '<div class="empty-state"><strong>No emails are available yet.</strong><br>'
            'Open Classic Outlook and start the listener, then return to this page.</div>',
            unsafe_allow_html=True,
        )
        return

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
        if st.session_state.get("advanced_rio_date_signature") != request_signature:
            st.session_state["advanced_dates"] = requested_dates
            st.session_state["advanced_rio_date_signature"] = request_signature

    st.markdown('<div class="advanced-section-heading">Search and filters</div>', unsafe_allow_html=True)
    search_column, date_column, day_column = st.columns([2.05, 1.05, 0.72])
    with search_column:
        search_text = st.text_input(
            "Search emails",
            placeholder="Search emails, senders, domains, or keywords…",
            key="advanced_search",
        )
    with date_column:
        selected_dates = st.date_input(
            "Date range",
            value=(default_start, default_end),
            min_value=selectable_start,
            max_value=selectable_end,
            key="advanced_dates",
        )
    start_date, end_date = _date_selection(selected_dates)
    records = get_dashboard_data(
        settings.database_path,
        None,
        start_date,
        end_date,
    )
    frame = pd.DataFrame(records)
    for required_column, default in (
        ("sender_name", ""),
        ("sender_email", ""),
        ("subject", ""),
        ("body_preview", ""),
        ("has_attachments", 0),
        ("is_read", 0),
        ("importance", "normal"),
        ("outlook_entry_id", None),
        ("outlook_store_id", None),
    ):
        if required_column not in frame.columns:
            frame[required_column] = default

    if not frame.empty:
        frame["received_at"] = pd.to_datetime(frame["received_at"], utc=True, errors="coerce")
        frame["received_day"] = frame["received_at"].dt.tz_convert("Asia/Kolkata").dt.day_name()
        frame["sender"] = frame["sender_name"].fillna("").astype(str)
        blank_names = frame["sender"].str.strip().eq("")
        frame.loc[blank_names, "sender"] = frame.loc[blank_names, "sender_email"]
        if search_text.strip():
            needle = search_text.strip()
            matches = (
                frame["sender_name"].fillna("").astype(str).str.contains(needle, case=False, regex=False)
                | frame["sender_email"].fillna("").astype(str).str.contains(needle, case=False, regex=False)
                | frame["subject"].fillna("").astype(str).str.contains(needle, case=False, regex=False)
            )
            frame = frame[matches]

    weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    available_days = set(frame["received_day"].dropna().astype(str)) if not frame.empty else set()
    day_options = ["All days", *[day for day in weekday_order if day in available_days]]
    if st.session_state.get("advanced_day", "All days") not in day_options:
        st.session_state.advanced_day = "All days"
    with day_column:
        selected_day = st.selectbox("Day", day_options, key="advanced_day")
    if selected_day != "All days" and not frame.empty:
        frame = frame[frame["received_day"].eq(selected_day)]

    sender_column, importance_column, attachment_column, status_column, clear_column = st.columns(
        [1.35, 1.2, 1.05, 0.95, 0.62], vertical_alignment="bottom"
    )
    sender_options = sorted(
        frame["sender_email"].dropna().astype(str).unique().tolist(),
        key=str.casefold,
    ) if not frame.empty else []
    current_senders = list(st.session_state.get("advanced_senders", []))
    valid_senders = [sender for sender in current_senders if sender in sender_options]
    if valid_senders != current_senders:
        st.session_state.advanced_senders = valid_senders
    with sender_column:
        selected_senders = st.multiselect(
            "Sender",
            sender_options,
            default=[],
            placeholder="All senders",
            key="advanced_senders",
        )
    if selected_senders and not frame.empty:
        frame = frame[frame["sender_email"].isin(selected_senders)]

    importance_order = ["Low", "Normal", "High"]
    available_importance = {
        value.title()
        for value in frame["importance"].fillna("normal").astype(str).str.lower().unique()
    } if not frame.empty else set()
    importance_options = [value for value in importance_order if value in available_importance]
    current_importance = st.session_state.get("advanced_importance")
    if current_importance is not None:
        valid_importance = [value for value in current_importance if value in importance_options]
        if valid_importance != list(current_importance):
            st.session_state.advanced_importance = valid_importance
    with importance_column:
        selected_importance = st.multiselect(
            "Importance",
            importance_options,
            default=importance_options,
            placeholder="All importance",
            key="advanced_importance",
        )
    importance_values = {value.lower() for value in selected_importance}
    if selected_importance and not frame.empty:
        frame = frame[frame["importance"].fillna("normal").str.lower().isin(importance_values)]

    attachment_options = ["All emails"]
    if not frame.empty and frame["has_attachments"].astype(bool).any():
        attachment_options.append("With attachments")
    if not frame.empty and (~frame["has_attachments"].astype(bool)).any():
        attachment_options.append("Without attachments")
    if st.session_state.get("advanced_attachments", "All emails") not in attachment_options:
        st.session_state.advanced_attachments = "All emails"
    with attachment_column:
        attachment_filter = st.selectbox(
            "Attachments",
            attachment_options,
            key="advanced_attachments",
        )
    if attachment_filter == "With attachments" and not frame.empty:
        frame = frame[frame["has_attachments"].astype(bool)]
    elif attachment_filter == "Without attachments" and not frame.empty:
        frame = frame[~frame["has_attachments"].astype(bool)]

    read_status_options = ["All emails"]
    if not frame.empty and frame["is_read"].astype(bool).any():
        read_status_options.append("Read")
    if not frame.empty and (~frame["is_read"].astype(bool)).any():
        read_status_options.append("Not Read")
    if st.session_state.get("advanced_read_status", "All emails") not in read_status_options:
        st.session_state.advanced_read_status = "All emails"
    with status_column:
        read_status_filter = st.selectbox(
            "Read status",
            read_status_options,
            key="advanced_read_status",
        )
    if read_status_filter == "Not Read" and not frame.empty:
        frame = frame[~frame["is_read"].astype(bool)]
    elif read_status_filter == "Read" and not frame.empty:
        frame = frame[frame["is_read"].astype(bool)]

    with clear_column:
        st.button(
            "Clear",
            icon=":material/close:",
            width="stretch",
            key="advanced_clear_filters",
            on_click=_reset_advanced_filters,
        )

    frame = frame.reset_index(drop=True)

    _render_kpi_cards(frame)

    if frame.empty:
        st.markdown('<div class="advanced-section-heading">Matching emails</div>', unsafe_allow_html=True)
        st.warning("No emails match these filters. Try changing the search text or filters.")
        return

    visible = frame[["received_at", "subject", "sender", "importance", "has_attachments", "is_read"]].copy()
    visible["received_at"] = visible["received_at"].dt.tz_convert("Asia/Kolkata").dt.strftime(
        "%d %b %Y, %I:%M %p"
    )
    visible["has_attachments"] = visible["has_attachments"].astype(bool).map({True: "Yes", False: "No"})
    visible["importance"] = visible["importance"].fillna("normal").str.title()
    visible["is_read"] = visible["is_read"].astype(bool).map({True: "Read", False: "Not Read"})
    visible.columns = ["Received", "Subject", "From", "Importance", "Attachments", "Read status"]

    with st.container(border=True, key="advanced_table_card"):
        title_column, clear_column, export_column, columns_column = st.columns(
            [3.45, 0.95, 0.72, 0.78], vertical_alignment="center"
        )
        with title_column:
            st.markdown(
                f'<div class="advanced-table-title"><strong>Matching emails</strong><span>{len(visible):,} results</span></div>',
                unsafe_allow_html=True,
            )
        with clear_column:
            has_selection = bool(st.query_params.get("selected_email"))
            if st.button(
                "Clear selection",
                icon=":material/deselect:",
                width="stretch",
                key="advanced_clear_selection",
                disabled=not has_selection,
            ):
                del st.query_params["selected_email"]
                st.rerun()
        with export_column:
            st.download_button(
                "Export",
                data=visible.to_csv(index=False).encode("utf-8"),
                file_name="filtered-emails.csv",
                mime="text/csv",
                icon=":material/upload:",
                width="stretch",
                key="advanced_export",
            )
        with columns_column:
            with st.popover("Columns", icon=":material/view_column:", width="stretch"):
                visible_columns = st.multiselect(
                    "Visible columns",
                    list(visible.columns),
                    default=list(visible.columns),
                    key="advanced_visible_columns",
                )
        if not visible_columns:
            visible_columns = ["Received", "Subject", "From"]

        st.caption("Select a row to inspect the email and use Outlook actions.")
        selected_index = _render_email_table(visible, frame, visible_columns)

    with st.expander("Visual overview", icon=":material/insert_chart:", expanded=False):
        st.caption("Click Visual overview again to hide these graphs.")
        _render_visual_overview(frame)

    if selected_index is not None and selected_index < len(frame):
        with st.container(border=True, key="advanced_selected_email"):
            _render_detail_panel(frame.iloc[selected_index])


__all__ = ["render_advanced_dashboard"]
