from __future__ import annotations

import base64
import html
from datetime import date, timedelta
from typing import Any
from urllib.parse import quote

import pandas as pd
import plotly.express as px
import streamlit as st

from .config import PROJECT_ROOT, get_settings
from .database import (
    add_priority_rule,
    count_emails,
    delete_priority_rule,
    get_dashboard_data,
    get_date_bounds,
    get_priority_rules,
    get_state,
    initialize_database,
)
from .mail_sync import sync_mail
from .outlook_desktop import OutlookDesktopError, display_outlook_item, get_outlook_item_details


ZETA_LOGO_PATH = PROJECT_ROOT / "assets" / "zeta-logo-primary.svg"
ZETA_LOGO_BASE64 = base64.b64encode(ZETA_LOGO_PATH.read_bytes()).decode("ascii")
HIGH_COLOR = "#EA332A"
CRITICAL_COLOR = "#DC2648"
HIGH_THREE_COLOR = "#F97316"
NORMAL_COLOR = "#2E6CE6"
INK = "#172554"
MUTED = "#7180A4"
PIE_COLORS = {"Critical": "#FF3366", "High Priority": "#FF9F1C", "Normal": "#00C2FF"}
DAILY_COLORS = {"Critical": "#D946EF", "High Priority": "#FF7A45", "Normal": "#22C55E"}
CHART_RAINBOW = ["#FF4D6D", "#FF8A34", "#F6C945", "#20C997", "#00A8E8", "#3A6FF7", "#7C3AED", "#D946EF"]
TEMPLATE_OPTIONS = ("Zeta Purple", "Ocean Blue", "Sunset Coral")
TEMPLATE_SLUGS = {"Zeta Purple": "zeta", "Ocean Blue": "ocean", "Sunset Coral": "sunset"}
REFRESH_MODES = ("Scheduled", "Live", "Manual")
LAYOUT_OPTIONS = ("2 queues", "3 queues")
LAYOUT_SLUGS = {"2 queues": "two", "3 queues": "three"}
DIALOG_STATE_VERSION = 3
TEMPLATE_STYLES = {
    "Zeta Purple": {
        "page": "radial-gradient(circle at 96% 2%, rgba(226,193,255,.56), transparent 31%), radial-gradient(circle at 3% 0%, rgba(186,207,255,.61), transparent 32%), #F4F6FF",
        "header": "linear-gradient(107deg, #1C2A5E 0%, #312E81 50%, #7530E7 100%)",
        "action": "linear-gradient(105deg, #6337EC, #7037E7 54%, #168EBB)",
        "underline": "linear-gradient(90deg, #6B35E5, #28BCD2, transparent)",
    },
    "Ocean Blue": {
        "page": "radial-gradient(circle at 96% 2%, rgba(125,211,252,.40), transparent 31%), radial-gradient(circle at 3% 0%, rgba(186,230,253,.65), transparent 32%), #F2F9FF",
        "header": "linear-gradient(107deg, #12315B 0%, #075985 48%, #0891B2 100%)",
        "action": "linear-gradient(105deg, #075985, #0E7490 54%, #22A6C8)",
        "underline": "linear-gradient(90deg, #0369A1, #22D3EE, transparent)",
    },
    "Sunset Coral": {
        "page": "radial-gradient(circle at 96% 2%, rgba(253,186,116,.35), transparent 31%), radial-gradient(circle at 3% 0%, rgba(254,202,202,.55), transparent 32%), #FFF8F6",
        "header": "linear-gradient(107deg, #4C1D3D 0%, #9F2D55 50%, #F06543 100%)",
        "action": "linear-gradient(105deg, #A82458, #E54B4B 54%, #F28C3B)",
        "underline": "linear-gradient(90deg, #D6295D, #F59E42, transparent)",
    },
}


def _safe(value: object) -> str:
    return html.escape(str(value), quote=True)


def _initialize_dialog_state() -> None:
    """Keep dialogs closed until this browser session produces a real UI event."""
    if st.session_state.get("priority_dialog_state_version") != DIALOG_STATE_VERSION:
        st.session_state["priority_dialog_state_version"] = DIALOG_STATE_VERSION
        st.session_state["priority_graph_dialog_open"] = False
        st.session_state["priority_email_dialog_open"] = False
        st.session_state["priority_email_dialog_id"] = ""

    st.session_state.setdefault("priority_graph_dialog_open", False)
    st.session_state.setdefault("priority_email_dialog_open", False)
    st.session_state.setdefault("priority_email_dialog_id", "")

    incoming_popup = str(st.query_params.get("priority_popup") or "")
    incoming_message_id = str(st.query_params.get("selected_priority_email") or "")

    # Treat the URL as a one-time click event, then remove it immediately so a
    # refresh or later Streamlit rerun cannot reopen the dialog.
    if incoming_popup == "email" and incoming_message_id:
        st.session_state["priority_email_dialog_id"] = incoming_message_id
        st.session_state["priority_email_dialog_open"] = True
        st.session_state["priority_graph_dialog_open"] = False

    for parameter in ("priority_popup", "selected_priority_email", "priority_click"):
        if parameter in st.query_params:
            del st.query_params[parameter]


def _open_graph_dialog() -> None:
    st.session_state["priority_graph_dialog_open"] = True
    st.session_state["priority_email_dialog_open"] = False


def _close_graph_dialog() -> None:
    st.session_state["priority_graph_dialog_open"] = False


def _close_email_dialog() -> None:
    st.session_state["priority_email_dialog_open"] = False
    st.session_state["priority_email_dialog_id"] = ""


def _initialize_header_preferences() -> None:
    template_by_slug = {slug: name for name, slug in TEMPLATE_SLUGS.items()}
    requested_template = template_by_slug.get(str(st.query_params.get("template") or "").casefold())
    if requested_template:
        st.session_state["priority_template"] = requested_template
    elif st.session_state.get("priority_template") not in TEMPLATE_OPTIONS:
        st.session_state["priority_template"] = "Zeta Purple"

    refresh_by_slug = {mode.casefold(): mode for mode in REFRESH_MODES}
    requested_refresh = refresh_by_slug.get(str(st.query_params.get("refresh_mode") or "").casefold())
    if requested_refresh:
        st.session_state["priority_schedule"] = requested_refresh
    elif st.session_state.get("priority_schedule") not in REFRESH_MODES:
        st.session_state["priority_schedule"] = "Scheduled"

    layout_by_slug = {slug: name for name, slug in LAYOUT_SLUGS.items()}
    requested_layout = layout_by_slug.get(str(st.query_params.get("layout") or "").casefold())
    if requested_layout:
        st.session_state["priority_layout"] = requested_layout
    elif st.session_state.get("priority_layout") not in LAYOUT_OPTIONS:
        st.session_state["priority_layout"] = "2 queues"


def _persist_template_preference() -> None:
    selected = str(st.session_state.get("priority_template") or "Zeta Purple")
    st.query_params["template"] = TEMPLATE_SLUGS.get(selected, "zeta")


def _persist_refresh_preference() -> None:
    selected = str(st.session_state.get("priority_schedule") or "Scheduled")
    st.query_params["refresh_mode"] = selected.casefold()


def _persist_layout_preference() -> None:
    selected = str(st.session_state.get("priority_layout") or "2 queues")
    st.query_params["layout"] = LAYOUT_SLUGS.get(selected, "two")


def _install_styles() -> None:
    st.markdown(
        """
        <style>
          :root { color-scheme: light; }
          .stApp {
            background:
              radial-gradient(circle at 96% 2%, rgba(226, 193, 255, .56), transparent 31%),
              radial-gradient(circle at 3% 0%, rgba(186, 207, 255, .61), transparent 32%),
              #F4F6FF !important;
            color: #172554 !important;
          }
          [data-testid="stHeader"] { background: transparent !important; }
          [data-testid="stMainBlockContainer"] { max-width: 1600px !important; padding-top: 1.25rem !important; }

          .st-key-priority_inbox_header {
            background: linear-gradient(107deg, #1C2A5E 0%, #312E81 50%, #7530E7 100%);
            border-radius: 28px; box-shadow: 0 18px 38px rgba(53, 42, 139, .22);
            margin-bottom: 1.65rem; min-height: 142px; padding: 1rem 1.15rem;
          }
          .priority-logo-plate {
            align-items: center; background: #FFFFFF; border-radius: 18px; display: flex;
            height: 74px; justify-content: center; padding: 13px; width: 100%;
            box-shadow: 0 7px 18px rgba(16, 25, 82, .18);
          }
          .priority-logo-plate img { display: block; max-height: 48px; max-width: 130px; width: 100%; }
          .priority-header-title { color: #FFFFFF !important; font-size: 1.7rem; font-weight: 850; letter-spacing: -.045em; line-height: 1.08; margin: 0 0 .42rem; }
          .priority-title-row { align-items: center; display: flex; flex-wrap: wrap; gap: .7rem; }
          .priority-dashboard-badge { background: rgba(255,255,255,.13); border: 1px solid rgba(255,255,255,.22); border-radius: 999px; color: #EEF2FF; font-size: .8rem; font-weight: 750; padding: .33rem .65rem; }
          .priority-header-copy { color: #DCE3FF !important; font-size: .84rem; line-height: 1.38; margin: .2rem 0 0; }
          .priority-listener { color: #FFFFFF; line-height: 1.35; padding-left: .2rem; }
          .priority-listener-main { align-items: center; display: flex; font-size: .91rem; font-weight: 800; gap: .52rem; }
          .priority-listener-dot { background: #46DA83; border-radius: 50%; box-shadow: 0 0 0 4px rgba(70,218,131,.15); display: inline-block; height: 10px; width: 10px; }
          .priority-listener-sub { color: #C9D2F5 !important; font-size: .78rem; margin: .35rem 0 0 1.15rem; }
          .st-key-priority_inbox_header [data-testid="stSelectbox"] label { display: none !important; }
          .st-key-priority_inbox_header [data-testid="stSelectbox"] div[data-baseweb="select"] > div,
          .st-key-priority_inbox_header [data-testid="stSelectbox"] .react-aria-ComboBox [role="group"],
          .st-key-priority_inbox_header [data-baseweb="select"] > div {
            background: #FFFFFF !important; border: 0 !important; border-radius: 12px !important;
            color: #172554 !important; min-height: 48px !important;
            box-shadow: 0 5px 14px rgba(12,18,83,.16) !important;
          }
          .st-key-priority_inbox_header [data-testid="stSelectbox"] div[data-baseweb="select"] *,
          .st-key-priority_inbox_header [data-testid="stSelectbox"] .react-aria-ComboBox input,
          .st-key-priority_inbox_header [data-testid="stSelectbox"] .react-aria-ComboBox button,
          .st-key-priority_inbox_header [data-baseweb="select"] * { color: #172554 !important; -webkit-text-fill-color: #172554 !important; }
          .st-key-priority_inbox_header .stButton > button {
            border-radius: 12px !important; font-size: .9rem !important; font-weight: 750 !important;
            min-height: 48px !important; width: 100% !important;
          }
          .st-key-priority_inbox_header .st-key-priority_refresh button {
            background: linear-gradient(115deg, #6840E8, #168FBD) !important;
            border: 1px solid rgba(255,255,255,.42) !important; color: #FFFFFF !important;
          }
          .st-key-priority_inbox_header .st-key-priority_next button {
            background: #FFFFFF !important; border: 0 !important; color: #172554 !important;
            box-shadow: 0 5px 14px rgba(12,18,83,.16) !important;
          }
          .st-key-priority_inbox_header .st-key-priority_next button p,
          .st-key-priority_inbox_header .st-key-priority_next button span,
          .st-key-priority_inbox_header .st-key-priority_next button svg { color: #172554 !important; fill: #172554 !important; }

          .priority-section-heading { color: #172554 !important; font-size: 1.5rem; font-weight: 850; letter-spacing: -.025em; margin: .35rem 0 .1rem; }
          .priority-section-heading::after { background: linear-gradient(90deg, #6B35E5, #28BCD2, transparent); border-radius: 999px; content: ""; display: block; height: 4px; margin-top: .85rem; width: 175px; }
          .priority-section-copy { color: #7180A4 !important; font-size: .9rem; margin: .55rem 0 1.05rem; }
          .st-key-priority_filters label { color: #172554 !important; font-size: .95rem !important; font-weight: 650 !important; }
          .st-key-priority_filters [data-testid="stTextInput"] div[data-baseweb="input"],
          .st-key-priority_filters [data-testid="stTextInputRootElement"],
          .st-key-priority_filters [data-testid="stDateInput"] div[data-baseweb="input"],
          .st-key-priority_filters [data-testid="stSelectbox"] div[data-baseweb="select"] > div,
          .st-key-priority_filters [data-testid="stSelectbox"] .react-aria-ComboBox [role="group"],
          .st-key-priority_filters [data-baseweb="input"],
          .st-key-priority_filters [data-baseweb="select"] > div {
            background: #FFFFFF !important; border: 1px solid #C9D3ED !important; border-radius: 12px !important;
            color: #172554 !important; min-height: 56px !important;
          }
          .st-key-priority_filters [data-testid="stTextInput"] input,
          .st-key-priority_filters [data-testid="stDateInput"] input,
          .st-key-priority_filters [data-testid="stSelectbox"] div[data-baseweb="select"] *,
          .st-key-priority_filters [data-testid="stSelectbox"] .react-aria-ComboBox input,
          .st-key-priority_filters [data-testid="stSelectbox"] .react-aria-ComboBox button,
          .st-key-priority_filters input, .st-key-priority_filters [data-baseweb="select"] * { color: #172554 !important; -webkit-text-fill-color: #172554 !important; }
          .st-key-priority_filters input::placeholder { color: #8C9AB8 !important; opacity: 1 !important; }
          .st-key-priority_graph_toggle button {
            background: linear-gradient(105deg, #6337EC, #7037E7 54%, #168EBB) !important;
            border: 1px solid #8998F7 !important; border-radius: 12px !important; color: #FFFFFF !important;
            font-weight: 800 !important; min-height: 56px !important; width: 100% !important;
            box-shadow: 0 9px 20px rgba(91,54,215,.18) !important;
          }
          .st-key-priority_graph_toggle button p, .st-key-priority_graph_toggle button span,
          .st-key-priority_graph_toggle button svg { color: #FFFFFF !important; fill: #FFFFFF !important; }

          .priority-kpi-grid { display: grid; gap: .9rem; grid-template-columns: repeat(4, minmax(0,1fr)); margin: 1rem 0 1.55rem; }
          .priority-kpi-grid.priority-kpi-grid-five { grid-template-columns: repeat(5, minmax(0,1fr)); }
          .priority-kpi-card {
            background: rgba(255,255,255,.76); border: 1px solid #A8B9FF; border-top: 5px solid var(--accent);
            border-radius: 17px; box-shadow: 0 10px 23px rgba(55,66,137,.11); min-height: 116px; padding: .82rem 1rem;
          }
          .priority-kpi-label { color: var(--accent) !important; font-size: .75rem; font-weight: 850; letter-spacing: .015em; text-transform: uppercase; }
          .priority-kpi-value { color: #172554 !important; font-size: 1.85rem; font-weight: 850; letter-spacing: -.05em; line-height: 1.05; margin: .28rem 0 .22rem; }
          .priority-kpi-note { color: #7180A4 !important; font-size: .75rem; }

          .priority-queue-card {
            background: rgba(255,255,255,.40); border: 1px solid #BBC7E5; border-radius: 16px;
            min-height: 480px; padding: 1.15rem 1.2rem; box-shadow: 0 9px 23px rgba(55,66,137,.07);
          }
          .priority-queue-heading { align-items: center; color: #172554 !important; display: flex; font-size: 1.05rem; font-weight: 850; gap: .6rem; margin-bottom: .95rem; }
          .priority-queue-dot { background: var(--queue-color); border-radius: 50%; display: inline-block; height: 16px; width: 16px; }
          .priority-queue-count { background: rgba(255,255,255,.78); border-radius: 999px; color: #5E6E93 !important; font-size: .78rem; margin-left: auto; padding: .28rem .58rem; }
          .priority-queue-list { max-height: 515px; overflow-y: auto; padding-right: .3rem; scrollbar-color: #B7C4E8 transparent; scrollbar-width: thin; }
          .priority-mail-link { color: inherit !important; display: block; text-decoration: none !important; }
          .priority-mail-link:hover { color: inherit !important; text-decoration: none !important; }
          .priority-mail-row { align-items: flex-start; background: #FFFFFF; border: 1px solid #E0E5F3; border-radius: 14px; display: flex; gap: .75rem; margin-bottom: .55rem; padding: .8rem .9rem; }
          .priority-mail-row:hover { border-color: #A6B7F1; box-shadow: 0 5px 14px rgba(56,69,145,.08); transform: translateY(-1px); }
          .priority-mail-row.selected { background: #EEF4FF; border-color: var(--queue-color); box-shadow: inset 4px 0 0 var(--queue-color), 0 5px 14px rgba(56,69,145,.08); }
          .priority-mail-radio { border: 1.5px solid #C8D0E1; border-radius: 50%; flex: 0 0 auto; height: 20px; margin-top: .2rem; width: 20px; }
          .priority-mail-radio.unread { background: var(--queue-color); border: 5px solid #FFFFFF; box-shadow: 0 0 0 1.5px var(--queue-color); }
          .priority-mail-main { color: #172554 !important; font-size: .84rem; line-height: 1.4; min-width: 0; }
          .priority-mail-sender { font-weight: 800; }
          .priority-mail-subject { font-weight: 600; }
          .priority-mail-meta { color: #7180A4 !important; font-size: .76rem; margin-top: .25rem; }
          .priority-match-pill { background: #F0F3FF; border-radius: 999px; color: #4F46A5 !important; display: inline-block; font-size: .69rem; font-weight: 700; margin-top: .38rem; padding: .18rem .42rem; }
          .priority-empty-queue { align-items: center; background: rgba(255,255,255,.6); border: 1px dashed #C4CEE6; border-radius: 14px; color: #7180A4 !important; display: flex; font-size: .95rem; justify-content: center; min-height: 330px; text-align: center; }
          .priority-queue-more { color: #7180A4 !important; font-size: .75rem; padding: .35rem .1rem; text-align: center; }

          .st-key-priority_rules { margin-top: 1rem; }
          .st-key-priority_high_rule_card,
          .st-key-priority_normal_rule_card {
            background: rgba(255,255,255,.82) !important; border: 1px solid #D2DAF1 !important;
            border-radius: 18px !important; box-shadow: 0 10px 24px rgba(55,66,137,.08) !important;
            min-height: 205px !important; padding: .78rem .95rem .65rem !important;
          }
          .st-key-priority_critical_rule_card {
            background: rgba(255,255,255,.82) !important; border: 1px solid #D2DAF1 !important;
            border-radius: 18px !important; border-top: 5px solid #DC2648 !important;
            box-shadow: 0 10px 24px rgba(55,66,137,.08) !important; min-height: 205px !important;
            padding: .78rem .95rem .65rem !important;
          }
          .st-key-priority_high_rule_card { border-top: 5px solid #EA332A !important; }
          .st-key-priority_normal_rule_card { border-top: 5px solid #2E6CE6 !important; }
          .st-key-priority_high_rule_card > div[data-testid="stVerticalBlockBorderWrapper"],
          .st-key-priority_normal_rule_card > div[data-testid="stVerticalBlockBorderWrapper"] {
            background: rgba(255,255,255,.82) !important; border: 1px solid #D2DAF1 !important;
            border-radius: 18px !important; box-shadow: 0 10px 24px rgba(55,66,137,.08) !important;
            min-height: 205px !important;
          }
          .st-key-priority_critical_rule_card > div[data-testid="stVerticalBlockBorderWrapper"] {
            background: rgba(255,255,255,.82) !important; border: 1px solid #D2DAF1 !important;
            border-radius: 18px !important; border-top: 5px solid #DC2648 !important;
            box-shadow: 0 10px 24px rgba(55,66,137,.08) !important; min-height: 205px !important;
          }
          .st-key-priority_critical_rule_card > div[data-testid="stVerticalBlockBorderWrapper"] > div { padding: .78rem .95rem .65rem !important; }
          .st-key-priority_high_rule_card > div[data-testid="stVerticalBlockBorderWrapper"] { border-top: 5px solid #EA332A !important; }
          .st-key-priority_normal_rule_card > div[data-testid="stVerticalBlockBorderWrapper"] { border-top: 5px solid #2E6CE6 !important; }
          .st-key-priority_high_rule_card > div[data-testid="stVerticalBlockBorderWrapper"] > div,
          .st-key-priority_normal_rule_card > div[data-testid="stVerticalBlockBorderWrapper"] > div { padding: .78rem .95rem .65rem !important; }
          .priority-rule-card { background: transparent; border: 0; box-shadow: none; padding: 0; }
          .priority-rule-card.high, .priority-rule-card.normal { border-top: 0; }
          .priority-rule-card h3 { color: #172554 !important; font-size: 1rem; margin: 0 0 .2rem; }
          .priority-rule-card p { color: #7180A4 !important; font-size: .77rem; margin: 0 0 .55rem; }
          .st-key-priority_rules [data-testid="stTextInput"] div[data-baseweb="input"],
          .st-key-priority_rules [data-testid="stTextInputRootElement"],
          .st-key-priority_rules [data-baseweb="input"] { background: #FFFFFF !important; border: 1px solid #C9D3ED !important; border-radius: 11px !important; }
          .st-key-priority_rules [data-testid="stTextInput"] input,
          .st-key-priority_rules input { color: #172554 !important; -webkit-text-fill-color: #172554 !important; }
          .st-key-priority_rules [data-testid="stFormSubmitButton"] button,
          .st-key-priority_rules .stButton > button {
            background: linear-gradient(105deg, #6337EC, #287FD0) !important; border: 0 !important;
            border-radius: 10px !important; color: #FFFFFF !important; font-weight: 800 !important; min-height: 43px !important;
          }
          .st-key-priority_rules button p, .st-key-priority_rules button span, .st-key-priority_rules button svg { color: #FFFFFF !important; fill: #FFFFFF !important; }
          [class*="st-key-priority_rule_delete_"] button {
            background: #FFFFFF !important; border: 1px solid #F2B8B8 !important; color: #B4232B !important;
            box-shadow: none !important; min-height: 38px !important;
          }
          [class*="st-key-priority_rule_delete_"] button p,
          [class*="st-key-priority_rule_delete_"] button span,
          [class*="st-key-priority_rule_delete_"] button svg { color: #B4232B !important; fill: #B4232B !important; }
          .priority-rule-list-label { color: #7180A4 !important; font-size: .72rem; font-weight: 850; letter-spacing: .08em; margin-top: .65rem; text-transform: uppercase; }
          .priority-rule-chip { background: #F3F5FF; border: 1px solid #D7DEF3; border-radius: 999px; color: #172554 !important; display: inline-block; font-size: .79rem; padding: .35rem .58rem; }
          .priority-rule-chip.high { background: #FFF4F4; border-color: #FFD0D0; color: #A61F1F !important; }
          .priority-rule-chip.critical { background: #FFF0F4; border-color: #FFB8CA; color: #A30F35 !important; }
          .priority-rule-chip.normal { background: #F1F6FF; border-color: #CBDAFC; color: #2254A4 !important; }
          .priority-rule-preview-group { margin: .35rem 0 .8rem; }
          .priority-rule-preview-label { color: #526187 !important; font-size: .72rem; font-weight: 850; letter-spacing: .06em; margin-bottom: .38rem; text-transform: uppercase; }
          .priority-rule-preview-items { display: flex; flex-wrap: wrap; gap: .42rem; }
          [class*="st-key-priority_rule_item_"] {
            background: #F8FAFF; border: 1px solid #E0E6F6; border-radius: 12px;
            margin-top: .45rem; padding: .35rem .38rem .2rem;
          }
          [class*="st-key-priority_rule_item_"] .priority-rule-chip {
            max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
          }

          .st-key-priority_selected_email {
            background: rgba(255,255,255,.88) !important; border: 1px solid #CDD7F1 !important;
            border-radius: 20px !important; box-shadow: 0 14px 30px rgba(55,66,137,.11) !important;
            margin: 1.25rem 0 2rem !important; padding: 1.2rem 1.35rem !important;
          }
          [data-baseweb="modal"] { backdrop-filter: blur(5px); background: rgba(17, 24, 62, .34) !important; }
          [data-testid="stDialog"] {
            background: linear-gradient(145deg, #F8FAFF 0%, #F4F0FF 58%, #EFFBFF 100%) !important;
            border: 1px solid rgba(255,255,255,.82) !important; border-radius: 24px !important;
            box-shadow: 0 30px 80px rgba(21,29,78,.34) !important; color: #172554 !important;
            max-height: 92vh !important; max-width: 1240px !important; overflow-y: auto !important;
            overscroll-behavior: contain; width: min(94vw, 1240px) !important;
          }
          [data-testid="stDialog"] div[role="dialog"] {
            background: linear-gradient(145deg, #F8FAFF 0%, #F4F0FF 58%, #EFFBFF 100%) !important;
            border: 1px solid rgba(255,255,255,.82) !important; border-radius: 24px !important;
            box-shadow: 0 30px 80px rgba(21,29,78,.34) !important; max-height: 92vh !important;
            max-width: 1240px !important; overflow-y: auto !important; overscroll-behavior: contain;
            width: min(94vw, 1240px) !important;
          }
          [data-testid="stDialog"] div[role="dialog"] > div {
            background: transparent !important; color: #172554 !important;
          }
          [role="dialog"] {
            background: linear-gradient(145deg, #F8FAFF 0%, #F4F0FF 58%, #EFFBFF 100%) !important;
            color: #172554 !important; max-height: 92vh !important; overflow-y: auto !important;
          }
          [role="dialog"] > div { background: transparent !important; color: #172554 !important; }
          [data-testid="stDialog"] [data-testid="stDialogHeader"],
          [data-testid="stDialog"] [data-testid="stDialogBody"] { background: transparent !important; color: #172554 !important; }
          [data-testid="stDialog"] h1, [data-testid="stDialog"] h2,
          [data-testid="stDialog"] h3, [data-testid="stDialog"] p { color: #172554 !important; }
          [data-testid="stDialog"] .st-key-priority_selected_email {
            margin: 0 !important; box-shadow: none !important;
          }
          [data-testid="stDialog"] .priority-section-heading { font-size: 1.28rem !important; margin-top: 0 !important; }
          [data-testid="stDialog"] .priority-section-heading::after { margin-top: .55rem !important; }
          [data-testid="stDialog"] .priority-email-body { max-height: 34vh !important; }
          [data-testid="stDialog"] .stButton > button {
            border-radius: 11px !important; font-weight: 800 !important; min-height: 44px !important;
          }
          .st-key-priority_graph_dialog_close button {
            background: #FFFFFF !important; border: 1px solid #AFC0E8 !important; color: #25345D !important;
            width: 100% !important;
          }
          .st-key-priority_graph_dialog_close button p,
          .st-key-priority_graph_dialog_close button span,
          .st-key-priority_graph_dialog_close button svg { color: #25345D !important; fill: #25345D !important; }
          .priority-detail-header { align-items: flex-start; display: flex; gap: 1rem; justify-content: space-between; margin-bottom: .85rem; }
          .priority-detail-eyebrow { color: #6D3CEB !important; font-size: .72rem; font-weight: 850; letter-spacing: .09em; text-transform: uppercase; }
          .priority-detail-subject { color: #172554 !important; font-size: 1.35rem; font-weight: 850; letter-spacing: -.025em; line-height: 1.25; margin: .28rem 0; }
          .priority-detail-from { color: #7180A4 !important; font-size: .84rem; }
          .priority-detail-kpi-grid { display: grid; gap: .8rem; grid-template-columns: repeat(4, minmax(0,1fr)); margin: .9rem 0 1.1rem; }
          .priority-detail-kpi-card { background: #F7F8FF; border: 1px solid #DCE3F5; border-radius: 14px; min-height: 86px; padding: .75rem .85rem; }
          .priority-detail-kpi-label { color: #7180A4 !important; font-size: .68rem; font-weight: 800; letter-spacing: .065em; text-transform: uppercase; }
          .priority-detail-kpi-value { color: #172554 !important; font-size: 1rem; font-weight: 850; line-height: 1.25; margin-top: .35rem; }
          .priority-detail-meta { color: #526182 !important; font-size: .83rem; line-height: 1.55; margin: .35rem 0 .85rem; }
          .priority-detail-meta strong { color: #172554 !important; }
          .priority-body-label { color: #172554 !important; font-size: .82rem; font-weight: 850; margin: .4rem 0 .45rem; }
          .priority-email-body { background: #FFFFFF; border: 1px solid #DCE3F5; border-radius: 14px; color: #2D3A5D !important; font-size: .84rem; line-height: 1.6; max-height: 300px; min-height: 130px; overflow-y: auto; padding: 1rem; white-space: pre-wrap; }
          .st-key-priority_selected_email .stButton > button {
            background: linear-gradient(105deg, #6337EC, #287FD0) !important; border: 0 !important;
            border-radius: 11px !important; color: #FFFFFF !important; font-weight: 800 !important; min-height: 46px !important; width: 100% !important;
          }
          .st-key-priority_selected_email .stButton > button:disabled { background: #D6DDEE !important; color: #73809C !important; opacity: 1 !important; }
          .st-key-priority_selected_email .stButton > button p,
          .st-key-priority_selected_email .stButton > button span,
          .st-key-priority_selected_email .stButton > button svg { color: inherit !important; fill: currentColor !important; }

          .st-key-priority_graph_queue, .st-key-priority_graph_daily,
          .st-key-priority_graph_trend, .st-key-priority_graph_senders {
            background: rgba(255,255,255,.88) !important; border: 1px solid #D4DCF2 !important;
            border-radius: 18px !important; box-shadow: 0 10px 25px rgba(55,66,137,.08) !important;
            margin-bottom: 1rem !important; padding: .95rem 1rem .35rem !important;
          }
          .priority-chart-title { color: #172554 !important; font-size: 1.05rem; font-weight: 850; margin: 0; }
          .priority-chart-copy { color: #7180A4 !important; font-size: .8rem; margin: .2rem 0 .6rem; }
          .priority-preview-badge { background: #EEE9FF; border: 1px solid #D9CEFF; border-radius: 999px; color: #6437D7 !important; display: inline-block; font-size: .65rem; font-weight: 850; letter-spacing: .07em; margin-bottom: .5rem; padding: .22rem .45rem; text-transform: uppercase; }
          [data-testid="stCaptionContainer"], [data-testid="stAlert"] { color: #64749B !important; }
          [data-baseweb="popover"] [role="listbox"], [data-baseweb="calendar"] { background: #FFFFFF !important; color: #172554 !important; }

          @media (max-width: 1050px) {
            .priority-kpi-grid { grid-template-columns: repeat(2, minmax(0,1fr)); }
            .priority-kpi-grid.priority-kpi-grid-five { grid-template-columns: repeat(2, minmax(0,1fr)); }
            .priority-detail-kpi-grid { grid-template-columns: repeat(2, minmax(0,1fr)); }
            .priority-header-title { font-size: 1.55rem; }
          }
          @media (max-width: 720px) {
            .priority-kpi-grid { grid-template-columns: 1fr; }
            .priority-kpi-grid.priority-kpi-grid-five { grid-template-columns: 1fr; }
            .priority-detail-kpi-grid { grid-template-columns: 1fr; }
            .priority-logo-plate { height: 68px; }
            .st-key-priority_inbox_header { border-radius: 20px; }
            [data-testid="stDialog"], [role="dialog"] { max-height: 94vh !important; width: 96vw !important; }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _install_template_styles(template_name: str) -> None:
    theme = TEMPLATE_STYLES.get(template_name, TEMPLATE_STYLES["Zeta Purple"])
    st.markdown(
        """
        <style>
          .stApp { background: %s !important; }
          .st-key-priority_inbox_header { background: %s !important; }
          .st-key-priority_inbox_header .st-key-priority_refresh button,
          .st-key-priority_graph_toggle button,
          .st-key-priority_selected_email .stButton > button,
          .st-key-priority_rules [data-testid="stFormSubmitButton"] button,
          .st-key-priority_rules .stButton > button { background: %s !important; }
          .priority-section-heading::after { background: %s !important; }
        </style>
        """
        % (theme["page"], theme["header"], theme["action"], theme["underline"]),
        unsafe_allow_html=True,
    )


def _rule_matches(pattern: str, record: dict[str, Any]) -> bool:
    needle = pattern.casefold()
    if "@" in needle:
        return needle in str(record.get("sender_email") or "").casefold()
    searchable = " ".join(
        str(record.get(field) or "")
        for field in ("sender_name", "sender_email", "subject", "body_preview")
    ).casefold()
    return needle in searchable


def _classify(
    records: list[dict[str, Any]],
    rules: list[dict[str, Any]],
    *,
    three_queue_layout: bool = False,
) -> pd.DataFrame:
    critical_rules = [rule["pattern"] for rule in rules if rule["rule_type"] == "critical"]
    high_rules = [rule["pattern"] for rule in rules if rule["rule_type"] == "high"]
    normal_rules = [rule["pattern"] for rule in rules if rule["rule_type"] == "normal"]
    rows: list[dict[str, Any]] = []
    for record in records:
        matched_critical = next((rule for rule in critical_rules if _rule_matches(rule, record)), None)
        matched_high = next((rule for rule in high_rules if _rule_matches(rule, record)), None)
        matched_normal = next((rule for rule in normal_rules if _rule_matches(rule, record)), None)
        if three_queue_layout and matched_critical:
            priority, reason = "Critical", f"Critical rule: {matched_critical}"
        elif matched_high:
            priority, reason = "High Priority", f"High rule: {matched_high}"
        elif str(record.get("importance") or "").casefold() == "high":
            priority, reason = "High Priority", "Outlook marked high"
        elif matched_normal:
            priority, reason = "Normal", f"Normal rule: {matched_normal}"
        else:
            priority, reason = "Normal", "Default inbox"

        received = pd.to_datetime(record.get("received_at"), utc=True, errors="coerce")
        received_text = (
            received.tz_convert("Asia/Kolkata").strftime("%d %b, %I:%M %p")
            if not pd.isna(received)
            else "Unknown time"
        )
        rows.append(
            {
                "Sender": record.get("sender_name") or record.get("sender_email") or "Unknown sender",
                "Email": record.get("sender_email") or "",
                "Subject": record.get("subject") or "(No subject)",
                "Received": received_text,
                "ReceivedAt": received,
                "Match": reason,
                "Priority": priority,
                "IsRead": bool(record.get("is_read")),
                "MessageId": record.get("message_id") or "",
                "BodyPreview": record.get("body_preview") or "",
                "HasAttachments": bool(record.get("has_attachments")),
                "Importance": record.get("importance") or "normal",
                "OutlookEntryId": record.get("outlook_entry_id") or "",
                "OutlookStoreId": record.get("outlook_store_id") or "",
            }
        )
    return pd.DataFrame(
        rows,
        columns=[
            "Sender", "Email", "Subject", "Received", "ReceivedAt", "Match", "Priority", "IsRead",
            "MessageId", "BodyPreview", "HasAttachments", "Importance", "OutlookEntryId", "OutlookStoreId",
        ],
    )


def _render_header(last_sync: str | None) -> None:
    with st.container(key="priority_inbox_header"):
        logo_col, title_col, listener_col, template_col, layout_col, schedule_col, refresh_col, next_col = st.columns(
            [.9, 3.0, 1.2, 1.0, 1.0, 1.05, .9, 1.05],
            vertical_alignment="center",
        )
        with logo_col:
            st.markdown(
                f'<div class="priority-logo-plate"><img src="data:image/svg+xml;base64,{ZETA_LOGO_BASE64}" alt="Zeta Global"></div>',
                unsafe_allow_html=True,
            )
        with title_col:
            st.markdown(
                '<div class="priority-header-title">Priority Inbox Analytics</div>'
                '<div class="priority-title-row"><span class="priority-dashboard-badge">Dashboard 1 of 3</span></div>'
                '<p class="priority-header-copy">Review high-priority messages first, then act without leaving the dashboard.</p>',
                unsafe_allow_html=True,
            )
        with listener_col:
            sync_note = "Synced" if last_sync else "Waiting for first sync"
            st.markdown(
                '<div class="priority-listener"><div class="priority-listener-main">'
                '<span class="priority-listener-dot"></span><span>COM listener<br>running</span></div>'
                f'<div class="priority-listener-sub">Classic Outlook COM · {sync_note}</div></div>',
                unsafe_allow_html=True,
            )
        with template_col:
            st.selectbox(
                "Dashboard template",
                TEMPLATE_OPTIONS,
                key="priority_template",
                label_visibility="collapsed",
                help="Changes the colors and visual style of Dashboard 1 only.",
                on_change=_persist_template_preference,
            )
        with layout_col:
            st.selectbox(
                "Queue layout",
                LAYOUT_OPTIONS,
                key="priority_layout",
                label_visibility="collapsed",
                help="Choose the standard High Priority/Normal layout or the Critical/High Priority/Normal layout.",
                on_change=_persist_layout_preference,
            )
        with schedule_col:
            st.selectbox(
                "Refresh mode",
                REFRESH_MODES,
                key="priority_schedule",
                label_visibility="collapsed",
                help="Scheduled uses the configured backup interval. Live checks for updates continuously. Manual updates only when you select Refresh.",
                on_change=_persist_refresh_preference,
            )
        with refresh_col:
            if st.button("Refresh", icon=":material/refresh:", width="stretch", key="priority_refresh"):
                try:
                    with st.spinner("Syncing Outlook…"):
                        saved = sync_mail(initial_days=90)
                except Exception as error:
                    st.session_state["priority_refresh_notice"] = ("error", str(error))
                else:
                    st.session_state["priority_refresh_notice"] = (
                        "success",
                        f"Outlook refresh complete. {saved:,} message record(s) added or updated.",
                    )
                st.rerun()
        with next_col:
            if st.button("Next dashboard", icon=":material/arrow_forward:", width="stretch", key="priority_next"):
                st.query_params["view"] = "advanced"
                st.rerun()


def _kpi_markup(
    total: int,
    unread: int,
    critical: int,
    high: int,
    normal: int,
    matched: int,
    *,
    three_queue_layout: bool,
) -> str:
    cards = [("Total mail", total, f"{unread:,} unread", "#4438B8")]
    if three_queue_layout:
        cards.append(("Critical", critical, "Immediate attention", CRITICAL_COLOR))
    cards.extend(
        [
            ("High priority", high, "Review first", HIGH_THREE_COLOR if three_queue_layout else HIGH_COLOR),
            ("Normal", normal, "Routine messages", NORMAL_COLOR),
            ("Rule matches", matched, "Classified automatically", "#7C3AED"),
        ]
    )
    grid_class = " priority-kpi-grid-five" if three_queue_layout else ""
    markup = [f'<div class="priority-kpi-grid{grid_class}">']
    for label, value, note, accent in cards:
        markup.append(
            f'<div class="priority-kpi-card" style="--accent:{accent}">'
            f'<div class="priority-kpi-label">{_safe(label)}</div>'
            f'<div class="priority-kpi-value">{value:,}</div>'
            f'<div class="priority-kpi-note">{_safe(note)}</div></div>'
        )
    markup.append("</div>")
    return "".join(markup)


def _queue_markup(
    frame: pd.DataFrame,
    title: str,
    kind: str,
    color: str,
    selected_message_id: str,
    template_name: str,
    refresh_mode: str,
    layout_mode: str,
) -> str:
    markup = [
        f'<div class="priority-queue-card" style="--queue-color:{color}">',
        f'<div class="priority-queue-heading"><span class="priority-queue-dot"></span>{_safe(title)}'
        f'<span class="priority-queue-count">{len(frame):,}</span></div>',
    ]
    if frame.empty:
        markup.append('<div class="priority-empty-queue">No matching messages</div>')
    else:
        markup.append('<div class="priority-queue-list">')
        display_frame = frame.head(50)
        for _, row in display_frame.iterrows():
            read_class = "" if bool(row["IsRead"]) else " unread"
            message_id = str(row.get("MessageId") or "")
            selected_class = " selected" if message_id and message_id == selected_message_id else ""
            preference_query = (
                f"&template={quote(TEMPLATE_SLUGS.get(template_name, 'zeta'), safe='')}"
                f"&refresh_mode={quote(refresh_mode.casefold(), safe='')}"
                f"&layout={quote(LAYOUT_SLUGS.get(layout_mode, 'two'), safe='')}"
            )
            selection_url = (
                f"?view=priority&selected_priority_email={quote(message_id, safe='')}"
                f"&priority_popup=email{preference_query}"
            )
            sender = _safe(row["Sender"])
            email = _safe(row["Email"])
            subject = _safe(row["Subject"])
            received = _safe(row["Received"])
            reason = _safe(row["Match"])
            sender_title = f'{sender} <span style="font-weight:500;color:#7180A4">({email})</span>' if email else sender
            markup.append(
                f'<a class="priority-mail-link" href="{selection_url}" target="_self">'
                f'<div class="priority-mail-row{selected_class}">'
                f'<span class="priority-mail-radio{read_class}"></span>'
                '<div class="priority-mail-main">'
                f'<div><span class="priority-mail-sender">{sender_title}</span> — '
                f'<span class="priority-mail-subject">{subject}</span></div>'
                f'<div class="priority-mail-meta">{received}</div>'
                f'<span class="priority-match-pill">{reason}</span>'
                '</div></div></a>'
            )
        if len(frame) > len(display_frame):
            markup.append(f'<div class="priority-queue-more">Showing 50 of {len(frame):,} messages</div>')
        markup.append("</div>")
    markup.append("</div>")
    return "".join(markup)


def _show_rules(database_path, rules: list[dict[str, Any]], rule_type: str) -> None:
    selected = [rule for rule in rules if rule["rule_type"] == rule_type]
    if not selected:
        st.caption("No rules added yet.")
        return

    label = "High Priority" if rule_type == "high" else rule_type.title()
    st.markdown(
        f'<div class="priority-rule-list-label">Active rules · {len(selected)}</div>',
        unsafe_allow_html=True,
    )
    chip_class = rule_type
    email_rules = [rule for rule in selected if "@" in str(rule["pattern"])]
    word_rules = [rule for rule in selected if "@" not in str(rule["pattern"])]
    with st.popover(
        f"Preview {label} rules",
        icon=":material/visibility:",
        width="stretch",
    ):
        st.caption("The dashboard searches these saved values without case sensitivity.")
        for group_label, group_rules in (("Email addresses", email_rules), ("Words and phrases", word_rules)):
            st.markdown(
                f'<div class="priority-rule-preview-label">{group_label}</div>',
                unsafe_allow_html=True,
            )
            if not group_rules:
                st.caption("None saved")
                continue
            for offset in range(0, len(group_rules), 4):
                row_rules = group_rules[offset : offset + 4]
                columns = st.columns(len(row_rules), gap="small")
                for column, rule in zip(columns, row_rules):
                    with column:
                        with st.container(key=f"priority_rule_item_{rule['id']}"):
                            label_col, delete_col = st.columns([5, 1], vertical_alignment="center", gap="small")
                            with label_col:
                                st.markdown(
                                    f'<span class="priority-rule-chip {chip_class}" title="{_safe(rule["pattern"])}">{_safe(rule["pattern"])}</span>',
                                    unsafe_allow_html=True,
                                )
                            with delete_col:
                                if st.button("×", key=f"priority_rule_delete_{rule['id']}", help="Remove this rule"):
                                    delete_priority_rule(database_path, int(rule["id"]))
                                    st.rerun()


def _render_rule_card(
    database_path,
    rules: list[dict[str, Any]],
    *,
    layout_type: str,
    rule_type: str,
    title: str,
    copy: str,
    placeholder: str,
    button_label: str,
) -> None:
    with st.container(border=True, key=f"priority_{rule_type}_rule_card"):
        st.markdown(
            f'<div class="priority-rule-card {rule_type}"><h3>{_safe(title)}</h3><p>{_safe(copy)}</p></div>',
            unsafe_allow_html=True,
        )
        with st.form(f"priority_{layout_type}_{rule_type}_rule_form", clear_on_submit=True):
            pattern = st.text_input(
                f"{title} input",
                placeholder=placeholder,
                label_visibility="collapsed",
            )
            submitted = st.form_submit_button(button_label, width="stretch")
        if submitted:
            if not pattern.strip():
                st.warning("Enter a keyword, phrase, or sender email.")
            elif add_priority_rule(database_path, rule_type, pattern, layout_type=layout_type):
                st.rerun()
            else:
                st.warning(f"That {title} entry is already active in this layout.")
        _show_rules(database_path, rules, rule_type)


def _render_rule_controls(
    database_path,
    rules: list[dict[str, Any]],
    *,
    three_queue_layout: bool,
) -> None:
    layout_type = "three" if three_queue_layout else "two"
    precedence_copy = (
        "Critical rules take precedence, followed by High Priority and Normal. Rules are saved only for this three-queue layout."
        if three_queue_layout
        else "High Priority rules take precedence over Normal. Rules are saved only for this two-queue layout."
    )
    st.markdown(
        '<div class="priority-section-heading">Automatic sorting rules</div>'
        f'<p class="priority-section-copy">Add a keyword, phrase, or sender email. {_safe(precedence_copy)}</p>',
        unsafe_allow_html=True,
    )
    cards = []
    if three_queue_layout:
        cards.append(
            {
                "rule_type": "critical",
                "title": "Critical rules",
                "copy": "Send matches to the red Critical queue.",
                "placeholder": "e.g. outage or incident@company.com",
                "button_label": "Add critical",
            }
        )
    cards.extend(
        [
            {
                "rule_type": "high",
                "title": "High priority rules",
                "copy": "Send matches to the High Priority queue.",
                "placeholder": "e.g. urgent or ceo@company.com",
                "button_label": "Add high priority",
            },
            {
                "rule_type": "normal",
                "title": "Normal rules",
                "copy": "Keep matches in the blue Normal queue.",
                "placeholder": "e.g. newsletter or updates@company.com",
                "button_label": "Add normal",
            },
        ]
    )
    with st.container(key="priority_rules"):
        columns = st.columns(len(cards), gap="large")
        for column, card in zip(columns, cards):
            with column:
                _render_rule_card(
                    database_path,
                    rules,
                    layout_type=layout_type,
                    **card,
                )


def _format_size(size: object) -> str:
    try:
        value = max(float(size or 0), 0)
    except (TypeError, ValueError):
        return "Unknown"
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:.0f} {unit}" if unit == "B" else f"{value:.1f} {unit}"
        value /= 1024
    return "Unknown"


def _render_selected_email(selected: pd.Series) -> None:
    entry_id = str(selected.get("OutlookEntryId") or "")
    store_id = str(selected.get("OutlookStoreId") or "") or None
    details: dict[str, Any] | None = None
    outlook_error = ""
    if entry_id:
        try:
            details = get_outlook_item_details(entry_id, store_id)
        except OutlookDesktopError as error:
            outlook_error = str(error)

    subject = str((details or {}).get("subject") or selected.get("Subject") or "(No subject)")
    sender_name = str((details or {}).get("sender_name") or selected.get("Sender") or "Unknown sender")
    sender_email = str((details or {}).get("sender_email") or selected.get("Email") or "")
    received_value = (details or {}).get("received_at") or selected.get("ReceivedAt")
    received = pd.to_datetime(received_value, utc=True, errors="coerce")
    received_text = (
        received.tz_convert("Asia/Kolkata").strftime("%d %b %Y, %I:%M %p")
        if pd.notna(received)
        else str(selected.get("Received") or "Unknown")
    )
    is_unread = bool((details or {}).get("unread")) if details else not bool(selected.get("IsRead"))
    importance = str((details or {}).get("importance") or selected.get("Importance") or "normal").title()
    attachments = list((details or {}).get("attachments") or [])
    attachment_count = len(attachments) if details else int(bool(selected.get("HasAttachments")))
    body = str((details or {}).get("body_preview") or selected.get("BodyPreview") or "")

    with st.container(border=True, key="priority_selected_email"):
        heading_col, close_col = st.columns([8, 1], vertical_alignment="top")
        with heading_col:
            st.markdown(
                '<div class="priority-detail-header"><div>'
                '<div class="priority-detail-eyebrow">Selected email</div>'
                f'<div class="priority-detail-subject">{_safe(subject)}</div>'
                f'<div class="priority-detail-from">From {_safe(sender_name)}'
                + (f' &lt;{_safe(sender_email)}&gt;' if sender_email else "")
                + '</div></div></div>',
                unsafe_allow_html=True,
            )
        with close_col:
            close_selection = st.button(
                "Close email",
                icon=":material/close:",
                width="stretch",
                key="priority_close_selection",
            )

        kpis = (
            ("Queue", str(selected.get("Priority") or "Normal")),
            ("Read status", "Unread" if is_unread else "Read"),
            ("Attachments", f"{attachment_count:,}"),
            ("Received", received_text),
        )
        kpi_markup = ['<div class="priority-detail-kpi-grid">']
        for label, value in kpis:
            kpi_markup.append(
                '<div class="priority-detail-kpi-card">'
                f'<div class="priority-detail-kpi-label">{_safe(label)}</div>'
                f'<div class="priority-detail-kpi-value">{_safe(value)}</div></div>'
            )
        kpi_markup.append("</div>")
        st.markdown("".join(kpi_markup), unsafe_allow_html=True)

        if details:
            metadata = []
            if details.get("to"):
                metadata.append(f'<strong>To:</strong> {_safe(details["to"])}')
            if details.get("cc"):
                metadata.append(f'<strong>CC:</strong> {_safe(details["cc"])}')
            metadata.append(f'<strong>Importance:</strong> {_safe(importance)}')
            metadata.append(f'<strong>Message size:</strong> {_safe(_format_size(details.get("size")))}')
            if details.get("categories"):
                metadata.append(f'<strong>Categories:</strong> {_safe(details["categories"])}')
            if attachments:
                attachment_names = ", ".join(
                    f'{_safe(item.get("name") or "Attachment")} ({_safe(_format_size(item.get("size")))})'
                    for item in attachments
                )
                metadata.append(f'<strong>Files:</strong> {attachment_names}')
            st.markdown(f'<div class="priority-detail-meta">{"<br>".join(metadata)}</div>', unsafe_allow_html=True)
        elif outlook_error:
            st.warning(f"Outlook details are temporarily unavailable: {outlook_error}")
        elif not entry_id:
            st.info("This stored email has no Outlook link yet. Its saved content preview is shown below.")

        body_text = body or "No stored mail content is available for this message."
        body_label = "Mail content from Outlook" if details else "Stored mail content preview"
        st.markdown(
            f'<div class="priority-body-label">{_safe(body_label)}</div>'
            f'<div class="priority-email-body">{html.escape(body_text)}</div>',
            unsafe_allow_html=True,
        )

        open_col, reply_col = st.columns(2)
        with open_col:
            open_message = st.button(
                "Open in Outlook",
                icon=":material/open_in_new:",
                width="stretch",
                disabled=not entry_id,
                key="priority_open_outlook",
            )
        with reply_col:
            reply_message = st.button(
                "Reply in Outlook",
                icon=":material/reply:",
                width="stretch",
                disabled=not entry_id,
                key="priority_reply_outlook",
            )

        if close_selection:
            _close_email_dialog()
            st.rerun(scope="app")
        if entry_id and (open_message or reply_message):
            try:
                display_outlook_item(entry_id, store_id, reply=reply_message)
            except OutlookDesktopError as error:
                st.error(str(error))
            else:
                st.toast("Reply draft opened in Outlook." if reply_message else "Email opened in Outlook.")


def _apply_preview_axes(chart, x_title: str, y_title: str, *, height: int = 245) -> None:
    chart.update_layout(
        height=height,
        margin=dict(l=72, r=28, t=12, b=68),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
        font=dict(color=INK, size=12),
        xaxis_title=x_title,
        yaxis_title=y_title,
        legend=dict(orientation="h", y=1.14, x=0, font=dict(color=INK, size=11)),
        hoverlabel=dict(bgcolor="#FFFFFF", font_color=INK),
    )
    axis_style = dict(
        showticklabels=True,
        tickfont=dict(color=INK, size=11),
        title_font=dict(color=INK, size=12),
        gridcolor="#E6EAF5",
        linecolor="#AAB6D0",
        showline=True,
        ticks="outside",
        tickcolor="#64748B",
        automargin=True,
        zeroline=False,
    )
    chart.update_xaxes(**axis_style)
    chart.update_yaxes(**axis_style)


def _chart_heading(title: str, copy: str) -> None:
    st.markdown(
        '<span class="priority-preview-badge">Graph preview</span>'
        f'<p class="priority-chart-title">{_safe(title)}</p>'
        f'<p class="priority-chart-copy">{_safe(copy)}</p>',
        unsafe_allow_html=True,
    )


def _render_graphs(
    frame: pd.DataFrame,
    *,
    show_heading: bool = True,
    three_queue_layout: bool = False,
) -> None:
    if show_heading:
        st.markdown(
            '<div class="priority-section-heading">Priority graphs</div>'
            '<p class="priority-section-copy">Visual breakdown of the mail currently included by the filters.</p>',
            unsafe_allow_html=True,
        )
    if frame.empty:
        st.info("No messages match the current filters, so graph previews are not available.")
        return

    valid = frame.dropna(subset=["ReceivedAt"]).copy()
    if not valid.empty:
        valid["Date"] = valid["ReceivedAt"].dt.tz_convert("Asia/Kolkata").dt.date

    chart_left, chart_right = st.columns(2, gap="large")
    with chart_left:
        with st.container(border=True, key="priority_graph_queue"):
            queue_order = ["Critical", "High Priority", "Normal"] if three_queue_layout else ["High Priority", "Normal"]
            _chart_heading("Queue distribution", "Message totals across the active priority queues")
            counts = frame["Priority"].value_counts().reindex(queue_order, fill_value=0).rename_axis("Queue").reset_index(name="Messages")
            chart = px.pie(
                counts,
                names="Queue",
                values="Messages",
                hole=.62,
                color="Queue",
                color_discrete_map=PIE_COLORS,
            )
            chart.update_layout(height=245, margin=dict(l=20, r=20, t=8, b=28), paper_bgcolor="rgba(0,0,0,0)", font=dict(color=INK, size=12), legend=dict(orientation="h", y=-.05, x=.05, font=dict(color=INK, size=11)))
            chart.update_traces(textinfo="label+percent+value", textfont=dict(color="#FFFFFF", size=12), marker=dict(line=dict(color="#FFFFFF", width=2)))
            st.plotly_chart(chart, width="stretch", config={"displayModeBar": False})
    with chart_right:
        with st.container(border=True, key="priority_graph_daily"):
            _chart_heading("Daily email volume", "Messages per queue over the selected date range")
            if valid.empty:
                st.info("No dated messages are available for this chart.")
            else:
                daily = valid.groupby(["Date", "Priority"], as_index=False).size().rename(columns={"size": "Messages"})
                chart = px.bar(
                    daily,
                    x="Date",
                    y="Messages",
                    color="Priority",
                    text="Messages",
                    barmode="stack",
                    color_discrete_map=DAILY_COLORS,
                )
                _apply_preview_axes(chart, "Date", "Email count")
                daily_dates = sorted(daily["Date"].unique().tolist())
                chart.update_xaxes(
                    tickmode="array",
                    tickvals=daily_dates,
                    ticktext=[value.strftime("%d %b") for value in daily_dates],
                )
                chart.update_yaxes(tickformat=",d", dtick=1 if int(daily["Messages"].max()) <= 12 else None)
                chart.update_traces(textposition="inside", textfont=dict(color="#FFFFFF", size=11))
                st.plotly_chart(chart, width="stretch", config={"displayModeBar": False})

    chart_left, chart_right = st.columns(2, gap="large")
    with chart_left:
        with st.container(border=True, key="priority_graph_trend"):
            _chart_heading("Email trend", "Total incoming messages by day")
            if valid.empty:
                st.info("No dated messages are available for this chart.")
            else:
                trend = valid.groupby("Date", as_index=False).size().rename(columns={"size": "Messages"})
                chart = px.line(trend, x="Date", y="Messages", markers=True, text="Messages")
                chart.update_traces(
                    line=dict(color="#7C3AED", width=4),
                    marker=dict(color="#00B8D9", size=10, line=dict(color="#FFFFFF", width=2)),
                    fill="tozeroy",
                    fillcolor="rgba(124,58,237,.13)",
                    textposition="top center",
                    textfont=dict(color=INK, size=11),
                )
                _apply_preview_axes(chart, "Date", "Email count")
                trend_dates = sorted(trend["Date"].unique().tolist())
                chart.update_xaxes(
                    tickmode="array",
                    tickvals=trend_dates,
                    ticktext=[value.strftime("%d %b") for value in trend_dates],
                )
                chart.update_yaxes(tickformat=",d", dtick=1 if int(trend["Messages"].max()) <= 12 else None, rangemode="tozero")
                st.plotly_chart(chart, width="stretch", config={"displayModeBar": False})
    with chart_right:
        with st.container(border=True, key="priority_graph_senders"):
            _chart_heading("Top senders", "Senders with the most messages in this view")
            senders = frame.groupby("Sender", as_index=False).size().rename(columns={"size": "Messages"}).nlargest(8, "Messages").sort_values("Messages")
            chart = px.bar(senders, x="Messages", y="Sender", orientation="h", text="Messages")
            _apply_preview_axes(chart, "Email count", "Sender")
            chart.update_layout(margin=dict(l=155, r=28, t=12, b=68), showlegend=False)
            chart.update_xaxes(tickformat=",d", dtick=1 if int(senders["Messages"].max()) <= 12 else None)
            sender_colors = CHART_RAINBOW[: len(senders)]
            chart.update_traces(
                marker=dict(color=sender_colors, line=dict(color="#FFFFFF", width=1)),
                textposition="outside",
                textfont=dict(color=INK, size=11),
                cliponaxis=False,
            )
            st.plotly_chart(chart, width="stretch", config={"displayModeBar": False})


@st.dialog("Email details", width="large")
def _show_email_dialog(selected: pd.Series) -> None:
    if not st.session_state.get("priority_email_dialog_open", False):
        st.rerun(scope="app")
    _render_selected_email(selected)


@st.dialog("Priority graph previews", width="large")
def _show_graph_dialog(frame: pd.DataFrame, three_queue_layout: bool) -> None:
    if not st.session_state.get("priority_graph_dialog_open", False):
        st.rerun(scope="app")

    copy_col, close_col = st.columns([4, 1.4], vertical_alignment="center")
    with copy_col:
        st.markdown(
            '<p class="priority-section-copy">Colorful previews for the messages currently included by your filters.</p>',
            unsafe_allow_html=True,
        )
    with close_col:
        if st.button(
            "Close popup",
            icon=":material/close:",
            width="stretch",
            key="priority_graph_dialog_close",
        ):
            _close_graph_dialog()
            st.rerun(scope="app")
    _render_graphs(frame, show_heading=False, three_queue_layout=three_queue_layout)


def render_priority_dashboard() -> None:
    settings = get_settings()
    initialize_database(settings.database_path)
    _initialize_dialog_state()
    _initialize_header_preferences()
    _install_styles()
    _install_template_styles(str(st.session_state.get("priority_template") or "Zeta Purple"))
    layout_mode = str(st.session_state.get("priority_layout") or "2 queues")
    three_queue_layout = layout_mode == "3 queues"
    layout_type = "three" if three_queue_layout else "two"
    rules = get_priority_rules(settings.database_path, layout_type=layout_type)
    total_in_database = count_emails(settings.database_path)
    _render_header(get_state(settings.database_path, "last_successful_sync_at"))
    refresh_notice = st.session_state.pop("priority_refresh_notice", None)
    if refresh_notice:
        notice_type, notice_text = refresh_notice
        if notice_type == "success":
            st.toast(notice_text, icon="✅")
        else:
            st.error(f"Refresh failed: {notice_text}")

    bounds = get_date_bounds(settings.database_path)
    st.markdown('<div class="priority-section-heading">Inbox filters</div>', unsafe_allow_html=True)
    with st.container(key="priority_filters"):
        search_col, status_col, date_col, graph_col = st.columns([1.75, .85, 1.2, .78], vertical_alignment="bottom")
        with search_col:
            search = st.text_input("Search sender or subject", placeholder="Search messages…", key="priority_search")
        with status_col:
            read_status = st.selectbox("Read status", ["All mail", "Unread", "Read"], key="priority_read_status")
        with date_col:
            if bounds:
                default_end = bounds[1]
                default_start = max(bounds[0], default_end - timedelta(days=1))
                selected_dates = st.date_input("Date range", value=(default_start, default_end), min_value=bounds[0], max_value=bounds[1], key="priority_dates")
            else:
                selected_dates = st.date_input("Date range", value=date.today(), key="priority_dates")
        with graph_col:
            st.button(
                "View graphs",
                icon=":material/insert_chart:",
                width="stretch",
                key="priority_graph_toggle",
                on_click=_open_graph_dialog,
            )

    start_date = end_date = None
    if isinstance(selected_dates, tuple) and len(selected_dates) == 2:
        start_date, end_date = selected_dates
    elif isinstance(selected_dates, date):
        start_date = end_date = selected_dates
    records = get_dashboard_data(settings.database_path, start_date=start_date, end_date=end_date)

    if read_status == "Unread":
        records = [record for record in records if not bool(record.get("is_read"))]
    elif read_status == "Read":
        records = [record for record in records if bool(record.get("is_read"))]
    if search.strip():
        needle = search.strip().casefold()
        records = [
            record for record in records
            if needle in " ".join(str(record.get(field) or "") for field in ("sender_name", "sender_email", "subject", "body_preview")).casefold()
        ]

    frame = _classify(records, rules, three_queue_layout=three_queue_layout)
    selected_message_id = str(st.session_state.get("priority_email_dialog_id") or "")
    template_name = str(st.session_state.get("priority_template") or "Zeta Purple")
    refresh_mode = str(st.session_state.get("priority_schedule") or "Scheduled")
    critical_frame = frame[frame["Priority"] == "Critical"].reset_index(drop=True) if not frame.empty else frame.copy()
    high_frame = frame[frame["Priority"] == "High Priority"].reset_index(drop=True) if not frame.empty else frame.copy()
    normal_frame = frame[frame["Priority"] == "Normal"].reset_index(drop=True) if not frame.empty else frame.copy()
    unread_count = int((~frame["IsRead"]).sum()) if not frame.empty else 0
    matched_count = int((frame["Match"] != "Default inbox").sum()) if not frame.empty else 0

    st.markdown(
        _kpi_markup(
            len(frame),
            unread_count,
            len(critical_frame),
            len(high_frame),
            len(normal_frame),
            matched_count,
            three_queue_layout=three_queue_layout,
        ),
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="priority-section-heading">Priority queues</div>'
        '<p class="priority-section-copy">Unread messages have a filled dot. Work from the most urgent queue toward Normal mail.</p>',
        unsafe_allow_html=True,
    )
    if three_queue_layout:
        queue_columns = st.columns(3, gap="large")
        queue_specs = (
            (critical_frame, "Critical", "critical", CRITICAL_COLOR),
            (high_frame, "High Priority", "high", HIGH_THREE_COLOR),
            (normal_frame, "Normal", "normal", NORMAL_COLOR),
        )
    else:
        queue_columns = st.columns(2, gap="large")
        queue_specs = (
            (high_frame, "High Priority", "high", HIGH_COLOR),
            (normal_frame, "Normal", "normal", NORMAL_COLOR),
        )
    for column, (queue_frame, queue_title, queue_kind, queue_color) in zip(queue_columns, queue_specs):
        with column:
            st.markdown(
                _queue_markup(
                    queue_frame,
                    queue_title,
                    queue_kind,
                    queue_color,
                    selected_message_id,
                    template_name,
                    refresh_mode,
                    layout_mode,
                ),
                unsafe_allow_html=True,
            )

    if st.session_state.get("priority_email_dialog_open", False) and selected_message_id and not st.session_state.get("priority_graph_dialog_open", False):
        selected_rows = frame[frame["MessageId"].fillna("").astype(str) == selected_message_id]
        if not selected_rows.empty:
            _show_email_dialog(selected_rows.iloc[0])
        else:
            st.warning("The selected email is outside the current filters. Clear the selection or adjust the filters.")
            if st.button("Clear selected email", icon=":material/close:", key="priority_clear_missing_selection"):
                _close_email_dialog()
                st.rerun(scope="app")

    _render_rule_controls(
        settings.database_path,
        rules,
        three_queue_layout=three_queue_layout,
    )

    if st.session_state.get("priority_graph_dialog_open", False):
        _show_graph_dialog(frame, three_queue_layout)

    if total_in_database == 0:
        st.info("Your local database is empty. Start the Outlook listener or seed demo data to populate this dashboard.")
    st.caption(f"Local database: `{settings.database_path}` · Rules are isolated to the active {layout_mode} layout.")


__all__ = ["render_priority_dashboard"]
