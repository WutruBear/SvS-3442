import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

import networkx as nx

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

st.set_page_config(page_title="SvS #3442 tool", page_icon="⚔️", layout="wide")

# ═══════════════════════════════════════════════════════════════════════════════
# CSS
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

    :root {
        --bg:        #080b11;
        --bg-1:      #0d1018;
        --bg-2:      #111520;
        --bg-3:      #151a28;
        --border:    #1c2336;
        --border-2:  #243048;
        --text:      #c4cedd;
        --text-dim:  #7a8aa0;
        --text-faint:#3e4e65;
        --accent:    #4a9eff;
        --accent-bg: #0a1828;
        --green:     #3ec97a;
        --green-bg:  #071510;
        --amber:     #f0a840;
        --amber-bg:  #140e04;
        --red:       #e05555;
        --red-bg:    #130808;
        --radius:    10px;
        --radius-sm: 6px;
    }

    html, body, [class*="css"] {
        font-family: 'Sora', sans-serif;
        -webkit-font-smoothing: antialiased;
    }
    .stApp { background: var(--bg); color: var(--text); }
    #MainMenu, footer, header { visibility: hidden; }

    h1, h2, h3, h4 {
        font-family: 'Sora', sans-serif;
        font-weight: 700; color: #dde6f0; letter-spacing: -0.025em;
    }
    label {
        color: var(--text-dim) !important; font-size: 0.78rem !important;
        font-weight: 600 !important; text-transform: uppercase !important;
        letter-spacing: 0.07em !important;
    }

    section[data-testid="stSidebar"],
    button[data-testid="collapsedControl"] { display: none !important; }

    .block-container {
        padding-top: 1.5rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 1200px !important;
    }
    @media (max-width: 640px) {
        .block-container {
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
            padding-top: 0.75rem !important;
        }
    }

    .section-label {
        font-size: 0.68rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.12em; color: var(--text-faint); margin-bottom: 0.55rem;
    }

    .topnav {
        display: flex; align-items: center; justify-content: space-between;
        background: var(--bg-1); border: 1px solid var(--border);
        border-radius: 14px; padding: 0.65rem 1.1rem;
        margin-bottom: 1.5rem; gap: 0.75rem;
    }
    .topnav-brand { display: flex; align-items: center; gap: 0.55rem; flex-shrink: 0; }
    .topnav-brand-icon { font-size: 1.1rem; line-height: 1; }
    .topnav-brand-name { font-size: 0.9rem; font-weight: 700; color: #dde6f0; letter-spacing: -0.01em; }
    .topnav-badge {
        font-size: 0.72rem; font-weight: 600; color: var(--green);
        background: var(--green-bg); border: 1px solid #1a3d2a;
        border-radius: 20px; padding: 0.18rem 0.65rem; white-space: nowrap;
    }
    @media (max-width: 480px) {
        .topnav-brand-name { font-size: 0.82rem; }
        .topnav-badge { font-size: 0.68rem; padding: 0.14rem 0.5rem; }
    }

    .stepper {
        display: flex; align-items: center;
        background: var(--bg-1); border: 1px solid var(--border);
        border-radius: 12px; padding: 0.7rem 1.1rem;
        margin-bottom: 1.4rem; overflow-x: auto; gap: 0;
        scrollbar-width: none;
    }
    .stepper::-webkit-scrollbar { display: none; }
    .step { display: flex; align-items: center; gap: 0.45rem; flex-shrink: 0; }
    .step-circle {
        width: 24px; height: 24px; border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        font-size: 0.68rem; font-weight: 700; flex-shrink: 0;
        transition: all 0.2s ease;
    }
    .step-circle.done   { background: #0e2a1c; color: var(--green);  border: 1.5px solid #1e5535; }
    .step-circle.active { background: #0a1d35; color: var(--accent);  border: 1.5px solid #1e4580; box-shadow: 0 0 0 3px rgba(74,158,255,0.12); }
    .step-circle.idle   { background: var(--bg-2); color: var(--text-faint); border: 1.5px solid var(--border); }
    .step-label { font-size: 0.75rem; font-weight: 600; white-space: nowrap; }
    .step-label.done   { color: var(--green); }
    .step-label.active { color: #90c4f8; }
    .step-label.idle   { color: var(--text-faint); }
    .step-arrow { color: var(--border-2); font-size: 0.7rem; margin: 0 0.6rem; flex-shrink: 0; }
    @media (max-width: 600px) {
        .step-label { display: none; }
        .step-circle { width: 28px; height: 28px; font-size: 0.72rem; }
        .step-arrow { margin: 0 0.35rem; }
        .stepper { padding: 0.6rem 0.9rem; }
    }

    .stats-row {
        display: grid; grid-template-columns: repeat(4, 1fr);
        gap: 0.65rem; margin-bottom: 1.5rem;
    }
    @media (max-width: 640px) { .stats-row { grid-template-columns: repeat(2, 1fr); } }
    @media (max-width: 340px) { .stats-row { grid-template-columns: 1fr; } }
    .stat-card {
        background: var(--bg-2); border: 1px solid var(--border);
        border-radius: var(--radius); padding: 0.9rem 1rem;
        transition: border-color 0.2s;
    }
    .stat-card:hover { border-color: var(--border-2); }
    .stat-card .stat-value {
        font-size: 1.75rem; font-weight: 700; color: #dde6f0;
        line-height: 1; margin-bottom: 0.28rem; letter-spacing: -0.03em;
    }
    .stat-card .stat-label {
        font-size: 0.68rem; font-weight: 600; text-transform: uppercase;
        letter-spacing: 0.09em; color: var(--text-faint);
    }
    .stat-card.warn { border-color: rgba(240,168,64,0.35); background: #120e04; }
    .stat-card.warn .stat-value { color: var(--amber); }

    .alert-banner {
        background: var(--amber-bg); border: 1px solid rgba(240,168,64,0.25);
        border-left: 3px solid var(--amber); border-radius: var(--radius-sm);
        padding: 0.8rem 1rem; margin: 0.65rem 0; font-size: 0.83rem;
        color: #e8c068; line-height: 1.55;
    }
    .error-banner {
        background: var(--red-bg); border: 1px solid rgba(224,85,85,0.2);
        border-left: 3px solid var(--red); border-radius: var(--radius-sm);
        padding: 0.75rem 1rem; margin: 0.4rem 0; font-size: 0.82rem;
        color: #df8080; line-height: 1.5;
    }
    .info-banner {
        background: var(--accent-bg); border: 1px solid rgba(74,158,255,0.2);
        border-left: 3px solid var(--accent); border-radius: var(--radius-sm);
        padding: 0.75rem 1rem; margin: 0.4rem 0; font-size: 0.82rem;
        color: #80b8f0; line-height: 1.5;
    }
    .success-banner {
        background: var(--green-bg); border: 1px solid rgba(62,201,122,0.2);
        border-left: 3px solid var(--green); border-radius: var(--radius-sm);
        padding: 0.75rem 1rem; margin: 0.4rem 0; font-size: 0.82rem;
        color: #60d090; line-height: 1.5;
    }
    .field-warn { font-size: 0.72rem; color: var(--amber); margin-top: -0.3rem; margin-bottom: 0.5rem; }

    textarea {
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.78rem !important; line-height: 1.7 !important;
        background: var(--bg-1) !important; color: #90aac8 !important;
        border: 1px solid var(--border) !important;
        border-radius: var(--radius-sm) !important;
    }
    textarea:focus {
        border-color: rgba(74,158,255,0.5) !important;
        box-shadow: 0 0 0 3px rgba(74,158,255,0.1) !important;
    }

    input[type="text"], .stTextInput input, .stNumberInput input {
        background: var(--bg-2) !important; border: 1px solid var(--border) !important;
        border-radius: var(--radius-sm) !important; color: var(--text) !important;
        font-family: 'Sora', sans-serif !important; font-size: 0.85rem !important;
        min-height: 40px !important;
    }
    input:focus {
        border-color: rgba(74,158,255,0.5) !important;
        box-shadow: 0 0 0 3px rgba(74,158,255,0.1) !important;
    }

    .stSelectbox > div > div {
        background: var(--bg-2) !important; border: 1px solid var(--border) !important;
        border-radius: var(--radius-sm) !important; color: var(--text) !important;
        min-height: 40px !important;
    }
    .stMultiSelect > div > div {
        background: var(--bg-2) !important; border: 1px solid var(--border) !important;
        border-radius: var(--radius-sm) !important;
    }

    .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
        background: var(--bg-3); color: #a8c0dc;
        border: 1px solid var(--border-2); border-radius: var(--radius-sm);
        font-family: 'Sora', sans-serif; font-weight: 600;
        font-size: 0.83rem; letter-spacing: 0.01em;
        padding: 0.55rem 1.25rem; min-height: 40px;
        transition: all 0.15s ease;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: #1a2640; border-color: #3a5878; color: #c8dcf0;
        transform: translateY(-1px); box-shadow: 0 4px 12px rgba(0,0,0,0.3);
    }
    .stButton > button:active, .stDownloadButton > button:active {
        transform: translateY(0); box-shadow: none;
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #1a3a6a 0%, #0f2448 100%) !important;
        color: #90c4f8 !important; border-color: #2a5090 !important;
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #20468a 0%, #132e5a 100%) !important;
        border-color: #3a68b8 !important; color: #b0d4ff !important;
    }
    @media (max-width: 480px) {
        .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
            font-size: 0.8rem; padding: 0.5rem 0.9rem;
        }
    }

    .streamlit-expanderHeader {
        background: var(--bg-2) !important; border: 1px solid var(--border) !important;
        border-radius: var(--radius-sm) !important; color: var(--text) !important;
        font-size: 0.84rem !important; font-weight: 600 !important;
        padding: 0.65rem 0.9rem !important;
    }
    .streamlit-expanderHeader:hover { border-color: var(--border-2) !important; }
    .streamlit-expanderContent {
        background: #0c1020 !important; border: 1px solid var(--border) !important;
        border-top: none !important; border-radius: 0 0 var(--radius-sm) var(--radius-sm) !important;
    }

    .stDataFrame { border: 1px solid var(--border) !important; border-radius: var(--radius) !important; overflow: hidden; }
    .stDataFrame > div { overflow-x: auto !important; }

    .stTabs [data-baseweb="tab-list"] {
        background: transparent !important; gap: 0.25rem;
        border-bottom: 1px solid var(--border) !important;
    }
    .stTabs [data-baseweb="tab"] {
        background: transparent !important; color: var(--text-dim) !important;
        font-family: 'Sora', sans-serif !important; font-size: 0.82rem !important;
        font-weight: 600 !important;
        border-radius: var(--radius-sm) var(--radius-sm) 0 0 !important;
        border: 1px solid transparent !important; border-bottom: none !important;
        padding: 0.5rem 0.9rem !important; transition: all 0.15s ease !important;
    }
    .stTabs [aria-selected="true"] {
        background: var(--bg-2) !important; color: #90c4f8 !important;
        border-color: var(--border) !important; border-bottom-color: var(--bg-2) !important;
    }
    @media (max-width: 480px) {
        .stTabs [data-baseweb="tab"] { font-size: 0.75rem !important; padding: 0.4rem 0.6rem !important; }
    }

    .stRadio [data-testid="stWidgetLabel"] { color: var(--text-dim) !important; }
    .stRadio label { text-transform: none !important; letter-spacing: 0 !important; font-size: 0.85rem !important; font-weight: 500 !important; }

    [data-testid="metric-container"] {
        background: var(--bg-2); border: 1px solid var(--border);
        border-radius: var(--radius); padding: 0.75rem 1rem;
    }
    [data-testid="metric-container"] label { font-size: 0.7rem !important; color: var(--text-faint) !important; }
    [data-testid="metric-container"] [data-testid="stMetricValue"] { color: #dde6f0 !important; font-weight: 700 !important; }

    .fc-raw {
        font-family: 'JetBrains Mono', monospace; font-size: 0.76rem; color: var(--text-dim);
        margin-bottom: 0.9rem; padding: 0.5rem 0.75rem;
        background: var(--bg); border-radius: var(--radius-sm); border: 1px solid var(--border);
    }
    .fc-raw span { color: #6ea8d8; }

    hr { border: none; border-top: 1px solid var(--border); margin: 1.5rem 0; }

    .guide-grid {
        display: grid; grid-template-columns: repeat(3, 1fr);
        gap: 0.75rem; margin-bottom: 1.5rem;
    }
    @media (max-width: 700px) { .guide-grid { grid-template-columns: 1fr; } }
    .guide-card {
        background: var(--bg-1); border: 1px solid var(--border);
        border-radius: var(--radius); padding: 1rem 1.1rem; transition: border-color 0.2s;
    }
    .guide-card:hover { border-color: var(--border-2); }
    .guide-card-icon  { font-size: 1.25rem; margin-bottom: 0.5rem; line-height: 1; }
    .guide-card-title { font-size: 0.82rem; font-weight: 700; color: #dde6f0; margin-bottom: 0.3rem; }
    .guide-card-body  { font-size: 0.77rem; color: var(--text-dim); line-height: 1.6; }
    .guide-card-body code {
        font-family: 'JetBrains Mono', monospace; font-size: 0.7rem;
        color: #6ea8d8; background: var(--bg-2); padding: 0.1rem 0.3rem; border-radius: 3px;
    }

    .fmt-table { width: 100%; border-collapse: collapse; font-size: 0.78rem; }
    .fmt-table th {
        text-align: left; padding: 0.45rem 0.75rem; color: var(--text-faint);
        font-size: 0.68rem; text-transform: uppercase; letter-spacing: 0.08em;
        border-bottom: 1px solid var(--border);
    }
    .fmt-table td { padding: 0.4rem 0.75rem; color: #8fa8c0; border-bottom: 1px solid var(--bg-2); }
    .fmt-table td:first-child { color: var(--text); font-weight: 600; }
    .fmt-table code {
        font-family: 'JetBrains Mono', monospace; font-size: 0.7rem;
        color: #6ea8d8; background: var(--bg); padding: 0.1rem 0.3rem; border-radius: 3px;
    }
    @media (max-width: 600px) {
        .fmt-table { font-size: 0.72rem; }
        .fmt-table th, .fmt-table td { padding: 0.35rem 0.5rem; }
    }

    .empty-state {
        text-align: center; padding: 3rem 1rem;
        border: 1px dashed var(--border); border-radius: 14px;
        background: var(--bg-1); margin: 1rem 0;
    }
    .empty-state-icon  { font-size: 2.25rem; margin-bottom: 0.8rem; line-height: 1; }
    .empty-state-title { font-size: 1rem; font-weight: 700; color: #dde6f0; margin-bottom: 0.4rem; }
    .empty-state-body  { font-size: 0.82rem; color: var(--text-faint); max-width: 380px; margin: 0 auto; line-height: 1.6; }

    .stSpinner > div { border-top-color: var(--accent) !important; }

    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: var(--bg); }
    ::-webkit-scrollbar-thumb { background: #2a3550; border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: #3a4870; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# SHARED CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════

NUM = r'\d+(?:[.,]\d+)?'
HIGH, MEDIUM, LOW = "high", "medium", "low"
MAX_REASONABLE_DAYS   = 365
MIN_TIME_WINDOW_SLOTS = 6

DISPLAY_FIELDS = ["User ID", "Level", "Construction", "Research", "Troops",
                  "FCs", "FC Shards", "Time UTC", "Days"]

FIELD_HINTS = {
    "Level":        "e.g. FC3, FC5",
    "Construction": "e.g. 24d 3h  or  42d  or  35",
    "Research":     "e.g. 42d  or  20",
    "Troops":       "e.g. 100d 10h  or  50",
    "FCs":          "Number of FCs, e.g. 2693 or 2700",
    "FC Shards":    "Number of shards, e.g. 434",
    "Time UTC":     "e.g. 16:00–19:00  or  14:30-17  or  13",
    "Days":         "e.g. Mon, Thu  or  1, 4",
}

FIELD_WARN_MSG = {
    LOW:    "⚠ Could not parse — please verify",
    MEDIUM: "⚠ Parsed with low confidence — please check",
}

SAMPLE_RAW = """\
User ID: 1
Level: FC1
CONSTRUCTION (Monday): 1
RESEARCH (Tuesday):  1
TROOPS (Thursday): 1
How many FCs and FC shards you have: FC 1 shards 1
Desired time UTC (minimum 3 hour window): 16.00-19.00
Desired day(s) (1,.2,.4): Mon, Tue

User ID: 2
Level: FC2
CONSTRUCTION (Monday): 2d 2h
RESEARCH (Tuesday): 2d 2m
TROOPS (Thursday): 2d 2h
How many FCs and FC shards you have: FC: 2, FC shards: 2  
Desired time UTC (minimum 3 hour window): 7utc till 21utc
Desired day(s): 2, 4

User ID: 3
Level: FC3
CONSTRUCTION (Monday): 3 days
RESEARCH (Tuesday): 3 day 1 hour
TROOPS (Thursday): 3 days
How many FCs and FC shards you have: 3 FCs, 3 FC shards
Desired time UTC (minimum 3 hour window): 00, 10, 11, 23
Desired day(s): Monday, Tuesday

User ID: 4
Level: FC4
CONSTRUCTION (Monday): 4d
RESEARCH (Tuesday): 4d 4h
TROOPS (Thursday): 4d
How many FCs and FC shards you have: 4 FC, 4 shards
Desired time UTC (minimum 3 hour window): 14-17
Desired day(s): 1, 2, 4

User ID: 5
Level: FC5
CONSTRUCTION (Monday): 7200min
RESEARCH (Tuesday): 5 days, 5 hours
TROOPS (Thursday): 5
How many FCs and FC shards you have: 5 Crystals, 5 shards
Desired time UTC (minimum 3 hour window): 00utc - 4utc, 9utc, 20-23
Desired day(s): Mon, Thu
"""

DAY_CONFIG = [
    {"day": 1, "label": "Day 1 — VP",  "col": "Construction"},
    {"day": 2, "label": "Day 2 — VP",  "col": "Research"},
    {"day": 4, "label": "Day 4 — MoE", "col": "Troops"},
]

# Day 4 VP is a virtual 5th config: re-runs Day 4 but only for unassigned players
DAY4_VP_CONFIG = {"day": 4, "label": "Day 4 — VP", "col": "Troops"}

# Extra per-day column shown in the timeline view
_DAY_EXTRA_COL = {1: "FCs", 2: "FC Shards", 4: None}


# ═══════════════════════════════════════════════════════════════════════════════
# UI HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def banner(kind: str, html: str) -> None:
    """Render an alert/info/success/error banner.

    Args:
        kind: One of 'alert', 'info', 'success', 'error'.
        html: Inner HTML content for the banner.
    """
    st.markdown(f'<div class="{kind}-banner">{html}</div>', unsafe_allow_html=True)


def stat_card(value, label: str, warn: bool = False) -> str:
    """Return HTML for a single stat card."""
    cls = "stat-card warn" if warn else "stat-card"
    return (
        f'<div class="{cls}">'
        f'<div class="stat-value">{value}</div>'
        f'<div class="stat-label">{label}</div>'
        f'</div>'
    )


def stat_row(*cards: str) -> None:
    """Render a row of stat cards."""
    st.markdown(
        f'<div class="stats-row">{"".join(cards)}</div>',
        unsafe_allow_html=True,
    )


def render_stepper(steps: list[tuple[str, str]]) -> None:
    """Render the workflow stepper.

    Args:
        steps: List of (label, state) tuples where state is 'done', 'active', or 'idle'.
    """
    parts = []
    for i, (label, state) in enumerate(steps):
        if i:
            parts.append('<span class="step-arrow">&#8250;</span>')
        parts.append(
            f'<div class="step">'
            f'<div class="step-circle {state}">{i + 1}</div>'
            f'<span class="step-label {state}">{label}</span>'
            f'</div>'
        )
    st.markdown(
        f'<div class="stepper">{"".join(parts)}</div>',
        unsafe_allow_html=True,
    )


# ═══════════════════════════════════════════════════════════════════════════════
# PARSER HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _parse_number_str(val: str) -> str:
    """Normalise a numeric string with thousand-separators to a plain decimal."""
    if "." in val and len(val.split(".")[-1]) == 3:
        return val.replace(".", "")
    if "," in val and len(val.split(",")[-1]) == 3:
        return val.replace(",", "")
    return val.replace(",", ".")


def normalize_duration(raw: str) -> tuple:
    """Parse a speedup duration string into (days_float, confidence)."""
    if not raw:
        return "", LOW
    raw = raw.strip().lower()
    days = 0.0

    # Explicit range like "3d-5d" — take the lower bound
    range_m = re.match(rf'({NUM})\s*d?\s*[-\u2013]\s*({NUM})\s*d', raw)
    if range_m:
        val = float(range_m.group(1).replace(",", "."))
        conf = MEDIUM if val <= MAX_REASONABLE_DAYS else LOW
        return round(val, 2), conf

    day_m  = re.search(rf'({NUM})\s*(?:d|day|days)\b', raw)
    hour_m = re.search(rf'({NUM})\s*(?:h|hr|hour|hours)\b', raw)
    min_m  = re.search(rf'({NUM})\s*(?:m|min|minute|minutes)\b', raw)

    if day_m:  days += float(day_m.group(1).replace(",", "."))
    if hour_m: days += float(hour_m.group(1).replace(",", ".")) / 24
    if min_m:  days += float(min_m.group(1).replace(",", ".")) / 1440

    if days > 0:
        conf = MEDIUM if days > MAX_REASONABLE_DAYS else HIGH
        return round(days, 2), conf

    # Bare number fallback
    num_m = re.search(rf'({NUM})', raw)
    if num_m:
        val = float(num_m.group(1).replace(",", "."))
        conf = MEDIUM if val > MAX_REASONABLE_DAYS else HIGH
        return val, conf

    return "", LOW


def parse_fc_shards(raw: str) -> tuple:
    """Parse an FC / shards line into (fc_val, fc_conf, shard_val, shard_conf)."""
    if not raw:
        return "", LOW, "", LOW

    text = raw.strip().lower()
    fc_val = shard_val = None

    # 1. Update the regex to capture the punctuation (\d+(?:[.,]\d+)?)
    for pat in [r'shards?\s*[:\-]?\s*(\d+(?:[.,]\d+)?)', r'(\d+(?:[.,]\d+)?)\s*(?:fc\s+)?shards?']:
        m = re.search(pat, text, re.I)
        if m:
            try:
                # 2. Now m.group(1) will be "1.900", allowing _parse_number_str to clean it up!
                shard_val = int(float(_parse_number_str(m.group(1))))
            except ValueError:
                pass
            break

    for pat in [
        r'fcs?\s*[:\-]\s*(\d+(?:[.,]\d+)?)',
        r'fcs?\s+(\d+(?:[.,]\d+)?)',
        r'(\d+(?:[.,]\d+)?)\s*fcs?\b',
        r'(\d+(?:[.,]\d+)?)\s*crystals?',
    ]:
        m = re.search(pat, text, re.I)
        if m:
            try:
                fc_val = int(float(_parse_number_str(m.group(1))))
            except ValueError:
                pass
            break

    # Single bare number — treat as FC count with 0 shards
    if fc_val is None and shard_val is None:
        m = re.match(r'^\s*(\d+(?:[.,]\d+)?)\s*$', text)
        if m:
            try:
                fc_val    = int(float(_parse_number_str(m.group(1))))
                shard_val = 0
            except ValueError:
                pass

    return (
        fc_val    if fc_val    is not None else "",
        HIGH if fc_val    is not None else LOW,
        shard_val if shard_val is not None else "",
        HIGH if shard_val is not None else LOW,
    )


def parse_hhmm(s: str) -> int | None:
    """Parse 'HH:MM', 'HH.MM', or bare 'HH' into total minutes since midnight."""
    s = s.strip()
    m = re.match(r'^(\d{1,2})[.:](\d{2})$', s)
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
    else:
        m = re.match(r'^(\d{1,2})$', s)
        if not m:
            return None
        h, mi = int(m.group(1)), 0
    return h * 60 + mi if (0 <= h <= 23 and 0 <= mi <= 59) else None


def normalize_time_utc(raw: str) -> tuple:
    """Parse an availability string into (slots_csv, confidence, slot_count)."""
    if not raw:
        return "", LOW, 0

    text = raw.lower().strip()
    text = text.replace("till", "-").replace(" to ", "-").replace("\u2013", "-")
    slots: set[int] = set()

    TIME_PAT  = r'\d{1,2}(?:[.:]\d{2})?'
    range_pat = rf'({TIME_PAT})\s*(?:utc)?\s*-\s*({TIME_PAT})\s*(?:utc)?'

    for m in re.finditer(range_pat, text):
        start = parse_hhmm(m.group(1))
        end   = parse_hhmm(m.group(2))
        if start is None or end is None:
            continue
        start = (start // 30) * 30
        end   = (end   // 30) * 30
        if start == end:
            continue
        # Walk 30-minute steps, wrapping midnight
        t = start
        while t != end:
            slots.add(t)
            t = (t + 30) % (24 * 60)

    text_no_ranges = re.sub(range_pat, " ", text)
    for m in re.finditer(rf'\b({TIME_PAT})\s*(?:utc)?\b', text_no_ranges):
        t = parse_hhmm(m.group(1))
        if t is None:
            continue
        slot = (t // 30) * 30
        if 0 <= slot < 24 * 60:
            slots.add(slot)
            next_slot = slot + 30
            if next_slot < 24 * 60:
                slots.add(next_slot)

    if not slots:
        return "", LOW, 0

    valid     = sorted(s for s in slots if 0 <= s < 24 * 60)
    slots_str = ",".join(f"{s // 60:02d}:{s % 60:02d}" for s in valid)
    return slots_str, HIGH, len(valid)


def normalize_days(raw: str) -> tuple:
    """Parse a days string into (csv_of_day_numbers, confidence)."""
    if not raw:
        return "", LOW

    text = raw.lower()
    day_map = {
        "mon": "1", "monday": "1", "tue": "2", "tuesday": "2",
        "wed": "3", "wednesday": "3", "thu": "4", "thursday": "4",
        "fri": "5", "friday": "5", "sat": "6", "saturday": "6",
        "sun": "7", "sunday": "7",
    }
    found = list(re.findall(r'\b([1-7])\b', text))
    for k, v in day_map.items():
        if re.search(rf'\b{k}\b', text):
            found.append(v)
    found = sorted(set(found), key=int)
    return (",".join(found), HIGH) if found else ("", LOW)


def extract_field(block: str, patterns: list) -> str:
    """Return the first capture group from the first matching pattern."""
    for pat in patterns:
        m = re.search(pat, block, re.I | re.MULTILINE)
        if m:
            return m.group(1).strip()
    return ""


# Field-specific patterns (module-level to avoid repeated allocation)
_CONSTRUCTION_PATS = [
    r'CONSTRUCTION\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
    r'Constr[a-z]{2,10}\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
]
_RESEARCH_PATS = [
    r'RESEARCH\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
    r'Rese?[a-z]{2,7}\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
]
_TROOPS_PATS = [
    r'TROOPS?\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
    r'Troop[a-z]{0,4}\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
]

_SPEEDUP_FIELDS: list[tuple[str, list]] = [
    ("Construction", _CONSTRUCTION_PATS),
    ("Research",     _RESEARCH_PATS),
    ("Troops",       _TROOPS_PATS),
]

_FC_PATS = [
    r'(?:How many FCs[^:\n]*|FC[s]?\s*/\s*[Ss]hard[s]?[^:\n]*)\s*[:\-]?\s*([^\n]+)',
    r'FC[s]?\s+and[^:\n]*[:\-]?\s*([^\n]+)',
    r'(?:Crystal[s]?[^:\n]*)\s*[:\-]?\s*([^\n]+)',
]

_TIME_PATS = [
    r'Desired\s+time\s+UTC[^:\n]*[:\-]?\s*([^\n]+)',
    r'Time\s+UTC[^:\n]*[:\-]?\s*([^\n]+)',
    r'UTC\s*[:\-]\s*([^\n]+)',
]

_DAYS_PATS = [
    r'Desired\s+day(?:\(s\))?\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
    r'Day(?:s)?\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
]


def parse_block(block: str) -> dict:
    """Parse a single player block into a record dict with confidence metadata."""
    r: dict = {}

    r["User ID"] = extract_field(block, [
        r'User\s*ID\s*[:\-]?\s*(\d+)', r'\bID\s*[:\-]?\s*(\d+)',
    ])

    r["Level"]        = extract_field(block, [r'Level\s*[:\-]?\s*(\S+)', r'LVL\s*[:\-]?\s*(\S+)'])
    r["_conf_Level"]  = HIGH if r["Level"] else LOW

    for field, pats in _SPEEDUP_FIELDS:
        val, conf          = normalize_duration(extract_field(block, pats))
        r[field]           = val
        r[f"_conf_{field}"] = conf

    fc_line          = extract_field(block, _FC_PATS)
    r["_fc_raw"]     = fc_line
    fc_v, fc_c, sh_v, sh_c = parse_fc_shards(fc_line)
    r["FCs"]          = fc_v;  r["_conf_FCs"]       = fc_c
    r["FC Shards"]    = sh_v;  r["_conf_FC Shards"] = sh_c

    tv, tc, slot_count = normalize_time_utc(extract_field(block, _TIME_PATS).strip())
    r["Time UTC"] = tv
    if tv and slot_count < MIN_TIME_WINDOW_SLOTS:
        r["_conf_Time UTC"] = MEDIUM
        r["_warn_Time UTC"] = f"Only {slot_count / 2:.4g}h window — minimum is 3h"
    else:
        r["_conf_Time UTC"] = tc

    dv, dc     = normalize_days(extract_field(block, _DAYS_PATS))
    r["Days"]  = dv
    r["_conf_Days"] = dc

    return r


def parse_input(text: str) -> tuple[list, list]:
    """Split raw text into player blocks and parse each one.

    Returns:
        (records, warnings) where records is a list of dicts and warnings is a
        list of human-readable issue strings.
    """
    parts = re.split(r'(?=User\s*ID\s*[:\-]?\s*\d)', text.strip(), flags=re.I)
    parts = [p.strip() for p in parts if p.strip()]

    records: list  = []
    warnings: list = []
    seen_ids: dict = {}

    for i, part in enumerate(parts, 1):
        rec = parse_block(part)
        if not rec["User ID"]:
            warnings.append(f"Block {i}: Could not find User ID — skipped.")
            continue
        uid = rec["User ID"]
        if uid in seen_ids:
            warnings.append(f"Duplicate User ID {uid} in blocks {seen_ids[uid]} and {i} — both kept.")
        seen_ids[uid] = i
        rec["_raw_block"] = part
        rec["_manual"]    = False
        records.append(rec)

    return records, warnings


# ═══════════════════════════════════════════════════════════════════════════════
# EXCEL BUILDER (Parser export)
# ═══════════════════════════════════════════════════════════════════════════════

def build_excel(df_export: pd.DataFrame, flagged_cells: set) -> bytes:
    """Build a styled openpyxl workbook and return it as bytes."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SvS Data"

    HEADER_FILL  = PatternFill("solid", fgColor="1A2535")
    FLAGGED_FILL = PatternFill("solid", fgColor="1E1608")
    EMPTY_FILL   = PatternFill("solid", fgColor="130E0E")
    THIN         = Side(style="thin", color="1E2533")
    BORDER       = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

    cols     = list(df_export.columns)
    uid_col  = cols.index("User ID") if "User ID" in cols else None

    # Header row
    for ci, col in enumerate(cols, 1):
        cell           = ws.cell(row=1, column=ci, value=col)
        cell.font      = Font(bold=True, color="7AABDC", size=10)
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = BORDER
    ws.row_dimensions[1].height = 22

    # Data rows
    for ri, (_, row) in enumerate(df_export.iterrows(), 2):
        uid = str(row["User ID"]) if uid_col is not None else None
        for ci, col in enumerate(cols, 1):
            val  = row[col]
            cell = ws.cell(row=ri, column=ci, value=None if val in ("", "—") else val)
            cell.alignment = Alignment(vertical="center")
            cell.border    = BORDER
            if val in ("", "—", None):
                cell.fill = EMPTY_FILL
                cell.font = Font(color="4A5568")
            elif uid and (uid, col) in flagged_cells:
                cell.fill = FLAGGED_FILL
                cell.font = Font(color="D4A855")
            else:
                cell.font = Font(color="C9D1DC")
        ws.row_dimensions[ri].height = 18

    # Auto-fit column widths
    for ci, col in enumerate(cols, 1):
        max_len = max(len(str(col)), *[len(str(r)) for r in df_export[col].fillna("")])
        ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 4, 42)

    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
# SCHEDULER HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def slots_str_to_hours(slots_str: str) -> list[int]:
    """Convert parser Time UTC output ("14:00,14:30,15:00") to unique integer hours."""
    if not slots_str:
        return []
    hours: set[int] = set()
    for tok in str(slots_str).split(","):
        tok = tok.strip()
        if ":" in tok:
            try:
                hours.add(int(tok.split(":")[0]))
            except ValueError:
                pass
        elif tok.isdigit():
            hours.add(int(tok))
    return sorted(hours)


def parse_ints(s) -> list[int]:
    """Parse a comma-separated string of integers, silently skipping bad tokens."""
    result = []
    for x in str(s).split(","):
        x = x.strip()
        if x:
            try:
                result.append(int(x))
            except ValueError:
                pass
    return result


def slot_label(slot: int) -> str:
    """Return a human-readable 30-minute window label, e.g. '14:00 – 14:30'."""
    h, half = divmod(slot, 2)
    end_h   = (h + 1) % 24 if half else h
    end_m   = "00" if half else "30"
    return f"{h:02d}:{'30' if half else '00'} – {end_h:02d}:{end_m}"


def _user_slots(user: dict) -> set[int]:
    """Return the set of 30-min slot indices for a user (h*2 and h*2+1 per hour)."""
    return {h * 2 + f for h in user["hours"] for f in (0, 1)}


# ─────────────────────────────────────────────────────────────────────────────
# MIN-COST MAX-FLOW SCHEDULER
#
# Objective (strict priority order):
#   1. Maximise the number of users assigned a slot  (max cardinality)
#   2. Among all maximum-cardinality solutions, maximise total score
#      — i.e. the highest-scoring users always win when slots are contested.
#
# Why MCMF instead of Hopcroft-Karp + swap pass?
# ────────────────────────────────────────────────
# The old HK approach ran BFS augmentation for cardinality, then tried a
# one-level score-swap pass to improve fairness.  That swap pass was too
# shallow: it could only displace one user at a time, so any inversion
# requiring a chain of re-routings was left unsolved.
#
# MCMF solves both objectives simultaneously in a single pass:
#
#   Graph:  S → user_i  (cap 1, cost 0)
#           user_i → slot_j  (cap 1, cost ∝ max_score − score_i)
#           slot_j → T  (cap 1, cost 0)
#
#   networkx.max_flow_min_cost() pushes the maximum feasible flow first
#   (= max users placed), then, among all maximum flows, minimises total
#   edge cost (= maximises total score of placed users).
#
# This is provably optimal: no assignment can simultaneously place more
# users AND rank higher-scorers better.
#
# Remaining "inversions" (an unplaced user scores higher than the lowest
# placed user) are mathematically unavoidable.  They only occur when the
# unplaced user's entire time window is saturated by higher-scorers who
# ALSO need those same slots, and no rerouting frees a slot without
# dropping someone — MCMF already tried every possible rerouting.
# ─────────────────────────────────────────────────────────────────────────────

def run_day(users: list, dc: dict) -> dict:
    """Run the MCMF scheduler for a single day config and return results dict."""
    col = dc["col"]
    day = dc["day"]

    # ── 1. Deduplicate by User ID ──────────────────────────────────────────────
    # Keep the highest-scoring entry when the same User ID appears twice.
    best_by_uid: dict = {}
    for u in users:
        uid = u["User ID"]
        if uid not in best_by_uid or u[col] > best_by_uid[uid][col]:
            best_by_uid[uid] = u

    eligible = [u for u in best_by_uid.values() if day in u["days"] and u["hours"]]
    if not eligible:
        return {
            "day": day, "col": col,
            "slot_occ": {}, "user_slot": {},
            "via_reshuffle": set(), "moved": {}, "chains": [],
            "unassigned": [], "eligible": [],
        }

    eligible = sorted(eligible, key=lambda u: u[col], reverse=True)

    # ── 2. Build slot universe ─────────────────────────────────────────────────
    all_slots = sorted({h * 2 + f for u in eligible for h in u["hours"] for f in (0, 1)})
    slot_set  = set(all_slots)

    # ── 3. Edge costs ──────────────────────────────────────────────────────────
    # cost_i = (max_score − score_i) * SCALE + i
    # • Lower cost  ↔  higher score  ↔  preferred by min-cost solver
    # • The +i tiebreak ensures deterministic results for identical scores.
    SCALE  = 10_000
    max_s  = eligible[0][col] if eligible else 1.0
    costs  = [int((max_s - u[col]) * SCALE) + i for i, u in enumerate(eligible)]

    # ── 4. Build the flow network ──────────────────────────────────────────────
    S, T         = "S", "T"
    user_nodes   = [f"U{i}"     for i in range(len(eligible))]
    slot_node    = {s: f"SL{s}" for s in all_slots}

    G = nx.DiGraph()
    for i in range(len(eligible)):
        G.add_edge(S, user_nodes[i], capacity=1, weight=0)

    for i, u in enumerate(eligible):
        u_slots = _user_slots(u) & slot_set
        for s in u_slots:
            G.add_edge(user_nodes[i], slot_node[s], capacity=1, weight=costs[i])

    for s in all_slots:
        G.add_edge(slot_node[s], T, capacity=1, weight=0)

    # ── 5. Solve ───────────────────────────────────────────────────────────────
    flow_dict = nx.max_flow_min_cost(G, S, T)

    # ── 6. Extract assignments ─────────────────────────────────────────────────
    slot_occ:  dict = {}
    user_slot: dict = {}

    for i, u in enumerate(eligible):
        uid = u["User ID"]
        un  = user_nodes[i]
        for s in all_slots:
            if flow_dict.get(un, {}).get(slot_node[s], 0) == 1:
                slot_occ[s]    = uid
                user_slot[uid] = s
                break

    # ── 7. Build unassigned list with blocker detail ───────────────────────────
    top_48_scores    = [u[col] for u in eligible[:48]]
    min_top_48_score = top_48_scores[-1] if top_48_scores else float("-inf")

    unassigned = []
    for u in eligible:
        uid = u["User ID"]
        if uid in user_slot:
            continue
        all_u_slots = _user_slots(u)
        blockers    = {slot_occ[s] for s in all_u_slots if s in slot_occ}
        names       = [f"{b}" for b in sorted(str(b) for b in blockers)]

        if u[col] <= min_top_48_score:
            unassigned.append({
                "user":   u,
                "reason": "not enough speedups",
                "detail": "-",
            })
        else:
            unassigned.append({
                "user":   u,
                "reason": "Window saturated",
                "detail": (
                    f"All {len(all_u_slots)} slot(s) in their time window are taken — "
                    f"Blocked by: {', '.join(names) or 'unknown'}."
                ),
            })

    return {
        "day":           day,
        "col":           col,
        "slot_occ":      slot_occ,
        "user_slot":     user_slot,
        "via_reshuffle": set(),
        "moved":         {},
        "chains":        [],
        "unassigned":    unassigned,
        "eligible":      eligible,
    }


def run_scheduler(users: list) -> list:
    """Run the scheduler for all DAY_CONFIG entries and return a list of results."""
    return [run_day(users, dc) for dc in DAY_CONFIG]


def run_day4_vp(users: list, day4_result: dict) -> dict:
    """Run a second VP scheduling pass for Day 4 using only unassigned players.

    Players who were unassigned in the Day 4 MoE run (window saturated or low
    score) get a second chance here.  Scheduling logic is identical to run_day
    (MCMF over Troops score) but the eligible pool is restricted to those
    unassigned IDs only.
    """
    unassigned_ids = {e["user"]["User ID"] for e in day4_result["unassigned"]}
    if not unassigned_ids:
        return {
            "day": 4, "col": "Troops",
            "slot_occ": {}, "user_slot": {},
            "via_reshuffle": set(), "moved": {}, "chains": [],
            "unassigned": [], "eligible": [],
            "is_vp4": True,
        }

    # Restrict to users who were unassigned on Day 4 and have time windows
    vp4_users = [u for u in users if u["User ID"] in unassigned_ids and 4 in u["days"]]

    result = run_day(vp4_users, DAY4_VP_CONFIG)
    result["is_vp4"] = True
    return result


# ─────────────────────────────────────────────────────────────────────────────
# DataFrame builders
# ─────────────────────────────────────────────────────────────────────────────

def build_timeline_df(users: list, dr: dict, extra_col: str | None = None) -> pd.DataFrame:
    """Build the slot-timeline DataFrame for a single day result."""
    col        = dr["col"]
    avail_slots = sorted({
        h * 2 + f
        for u in dr["eligible"]
        for h in u["hours"]
        for f in (0, 1)
    })
    # Build a lookup for quick user access
    user_by_id = {u["User ID"]: u for u in users}

    rows = []
    for slot in avail_slots:
        auid = dr["slot_occ"].get(slot)
        au   = user_by_id.get(auid)
        if au:
            assigned_id    = f"{au['User ID']}"
            assigned_score = au[col]
        else:
            assigned_id, assigned_score = "empty", None

        row = {
            "Time Slot": slot_label(slot),
            "Assigned":  assigned_id,
            "Speedups":  f"{assigned_score:.2f}" if assigned_score is not None else "—",
        }
        if extra_col:
            row[extra_col] = (
                str(int(au[extra_col]))
                if au and au.get(extra_col) not in (None, "", "—")
                else "—"
            )
        rows.append(row)
    return pd.DataFrame(rows)


def build_unassigned_df(dr: dict) -> pd.DataFrame:
    """Build the unassigned-users DataFrame for a single day result."""
    return pd.DataFrame([{
        "User ID":  f"{e['user']['User ID']}",
        "Speedups": f"{e['user'][dr['col']]:.2f}",
        "Reason":   e["reason"],
        "Detail":   e["detail"],
    } for e in dr["unassigned"]])


def build_summary_df(users: list, day_results: list, vp4_result: dict | None = None) -> pd.DataFrame:
    """Build the cross-day summary DataFrame."""
    day_map = {dr["day"]: dr for dr in day_results}
    rows = []
    for u in users:
        row = {"User ID": u["User ID"], "Level": u["Level"]}
        for dc in DAY_CONFIG:
            dr  = day_map[dc["day"]]
            col = dc["col"]
            if dc["day"] not in u["days"]:
                row[dc["label"]] = "—"
            elif u["User ID"] in dr["user_slot"]:
                slot = dr["user_slot"][u["User ID"]]
                row[dc["label"]] = f"{slot_label(slot)}  [{u[col]:.2f}]"
            else:
                row[dc["label"]] = f"❌ unassigned  [{u[col]:.2f}]"

        # Day 4 — VP column: only meaningful for players unassigned in Day 4 MoE
        if vp4_result is not None:
            day4_dr = day_map.get(4)
            col     = "Troops"
            if 4 not in u["days"]:
                row[DAY4_VP_CONFIG["label"]] = "—"
            elif day4_dr and u["User ID"] in day4_dr["user_slot"]:
                # Already placed in MoE — not eligible for VP
                row[DAY4_VP_CONFIG["label"]] = "— (MoE placed)"
            elif u["User ID"] in vp4_result["user_slot"]:
                slot = vp4_result["user_slot"][u["User ID"]]
                row[DAY4_VP_CONFIG["label"]] = f"{slot_label(slot)}  [{u[col]:.2f}]"
            elif any(e["user"]["User ID"] == u["User ID"] for e in (day4_dr["unassigned"] if day4_dr else [])):
                row[DAY4_VP_CONFIG["label"]] = f"❌ unassigned  [{u[col]:.2f}]"
            else:
                row[DAY4_VP_CONFIG["label"]] = "—"

        rows.append(row)
    return pd.DataFrame(rows)


def to_excel_schedule(users: list, day_results: list, vp4_result: dict | None = None) -> io.BytesIO:
    """Export scheduler results to a multi-sheet xlsx file."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        build_summary_df(users, day_results, vp4_result).to_excel(
            writer, sheet_name="Summary", index=False
        )
        for dr in day_results:
            dc    = next(d for d in DAY_CONFIG if d["day"] == dr["day"])
            sheet = dc["label"].replace(" — ", " ").replace(" ", "_")[:31]
            build_timeline_df(users, dr).to_excel(
                writer, sheet_name=sheet, index=False
            )
            ua = build_unassigned_df(dr)
            if not ua.empty:
                ua.to_excel(
                    writer, sheet_name=f"Unassigned_Day{dc['day']}", index=False
                )
        # Day 4 — VP sheet
        if vp4_result is not None and (vp4_result["user_slot"] or vp4_result["unassigned"]):
            build_timeline_df(users, vp4_result).to_excel(
                writer, sheet_name="Day_4_VP", index=False
            )
            ua_vp4 = build_unassigned_df(vp4_result)
            if not ua_vp4.empty:
                ua_vp4.to_excel(writer, sheet_name="Unassigned_Day4_VP", index=False)
    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════════════════════
# SESSION STATE & BROWSER STORAGE INIT
# ═══════════════════════════════════════════════════════════════════════════════

from streamlit_local_storage import LocalStorage

local_storage = LocalStorage()

browser_saved_data = local_storage.getItem("svs_raw_input")
initial_text       = browser_saved_data if browser_saved_data and browser_saved_data.strip() else SAMPLE_RAW

_SESSION_DEFAULTS = {
    "page":           "parser",
    "raw_input":      initial_text,
    "manual_records": [],
    "excluded_ids":   set(),
    "corrections":    {},
    "_last_hash":     None,
    "_clipboard_csv": None,
    "parsed_df":      None,
}
for _k, _v in _SESSION_DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ═══════════════════════════════════════════════════════════════════════════════
# TOP NAVIGATION BAR
# ═══════════════════════════════════════════════════════════════════════════════

if st.session_state["parsed_df"] is not None:
    n           = len(st.session_state["parsed_df"])
    badge_html  = (
        f'<span class="topnav-badge">&#10003; {n} player{"s" if n != 1 else ""} parsed</span>'
    )
else:
    badge_html  = (
        '<span class="topnav-badge" style="color:#5c6a82;background:#0e1117;border-color:#1e2533;">'
        'no data yet</span>'
    )

st.markdown(f"""
<div class="topnav">
  <div class="topnav-brand">
    <span class="topnav-brand-icon">&#9876;&#65039;</span>
    <span class="topnav-brand-name">SvS #3442</span>
  </div>
  <div style="display:flex;align-items:center;gap:0.75rem;">
    {badge_html}
  </div>
</div>
""", unsafe_allow_html=True)

nav_col1, nav_col2, _ = st.columns([1.5, 1.5, 5])
with nav_col1:
    if st.button(
        "📋  Parser", use_container_width=True,
        type="primary" if st.session_state["page"] == "parser" else "secondary",
    ):
        st.session_state["page"] = "parser"
        st.rerun()
with nav_col2:
    if st.button(
        "🗓️  Scheduler", use_container_width=True,
        type="primary" if st.session_state["page"] == "scheduler" else "secondary",
    ):
        st.session_state["page"] = "scheduler"
        st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — PARSER
# ═══════════════════════════════════════════════════════════════════════════════

if st.session_state["page"] == "parser":

    has_data = st.session_state["parsed_df"] is not None
    render_stepper([
        ("Paste player data",  "active"),
        ("Review & correct",   "active"),
        ("Send to Scheduler",  "done" if has_data else "idle"),
        ("Run &amp; export",   "idle"),
    ])

    with st.expander("&#128218;  How to use this tool", expanded=(not has_data)):
        st.markdown("""
        <div class="guide-grid">
          <div class="guide-card">
            <div class="guide-card-icon">1&#65039;&#8419;</div>
            <div class="guide-card-title">Paste player responses</div>
            <div class="guide-card-body">
              Copy the raw sign-up replies from WOS
              and paste the whole block into the <b>Raw Input</b> box. Each player entry
              must start with <code>User ID: &lt;number&gt;</code>.
            </div>
          </div>
          <div class="guide-card">
            <div class="guide-card-icon">2&#65039;&#8419;</div>
            <div class="guide-card-title">Fix flagged fields</div>
            <div class="guide-card-body">
              Fields the parser couldn't read confidently are highlighted in orange below.
              Click the field and type the correct value. You can also add missing players
              with the <b>Add player manually</b> form.
            </div>
          </div>
          <div class="guide-card">
            <div class="guide-card-icon">3&#65039;&#8419;</div>
            <div class="guide-card-title">Send to Scheduler</div>
            <div class="guide-card-body">
              Once everything looks right, click <b>Send to Scheduler →</b> at the bottom.
              Switch to the <b>Scheduler</b> tab, then hit <b>Run scheduler</b> to assign
              every player a 30-minute slot on each of their days.
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <table class="fmt-table">
          <tr><th>Field</th><th>What it means</th><th>Accepted formats</th></tr>
          <tr><td>User ID</td><td>Unique player number</td><td><code>1</code>, <code>12345</code></td></tr>
          <tr><td>Level</td><td>FC tier</td><td><code>FC1</code> … <code>FC5</code></td></tr>
          <tr><td>Construction / Research / Troops</td><td>Speedups in days</td>
              <td><code>4d</code>, <code>3d 12h</code>, <code>7200min</code>, <code>5</code></td></tr>
          <tr><td>FCs / FC Shards</td><td>Resource counts</td>
              <td><code>FC 2693 shards 434</code>, <code>2,700 FCs</code></td></tr>
          <tr><td>Time UTC</td><td>Available hours (min 3 h window)</td>
              <td><code>14-17</code>, <code>16:00-19:30</code>, <code>00utc - 4utc, 20-23</code></td></tr>
          <tr><td>Days</td><td>Which SvS days the player joins</td>
              <td><code>1, 2, 4</code>, <code>Mon, Thu</code>, <code>Monday</code></td></tr>
        </table>
        <div style="font-size:0.73rem;color:#3a4558;margin-top:0.6rem;">
          Days: 1 = Construction &nbsp;·&nbsp; 2 = Research &nbsp;·&nbsp; 4 = Troops
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<div style='margin-bottom:1rem'></div>", unsafe_allow_html=True)

    col_input, col_tools = st.columns([4, 1])

    def _save_to_browser():
        local_storage.setItem("svs_raw_input", st.session_state["raw_input"])

    with col_input:
        st.markdown('<div class="section-label">Raw Input</div>', unsafe_allow_html=True)
        raw_text = st.text_area(
            "raw",
            height=340,
            label_visibility="collapsed",
            key="raw_input",
            on_change=_save_to_browser,
        )

    with col_tools:
        st.markdown('<div class="section-label">Input File</div>', unsafe_allow_html=True)
        uploaded = st.file_uploader("Load .txt", type=["txt"], label_visibility="collapsed")
        if uploaded:
            st.session_state["raw_input"] = uploaded.read().decode("utf-8")
            st.rerun()
        st.download_button(
            "💾 Save raw input",
            data=st.session_state["raw_input"].encode("utf-8"),
            file_name="svs_input.txt",
            mime="text/plain",
            use_container_width=True,
        )

    if not raw_text.strip():
        st.info("Paste your player data in the text box above.")
        st.stop()

    records, parse_warnings = parse_input(raw_text)
    input_hash               = hash(raw_text)

    if st.session_state["_last_hash"] != input_hash:
        new_corr = {
            rec["User ID"]: {f: rec[f] for f in DISPLAY_FIELDS if f != "User ID"}
            for rec in records
        }
        for mrec in st.session_state["manual_records"]:
            uid = mrec["User ID"]
            if uid not in new_corr:
                new_corr[uid] = {f: mrec.get(f, "") for f in DISPLAY_FIELDS if f != "User ID"}
        st.session_state["corrections"] = new_corr
        st.session_state["_last_hash"]  = input_hash

    for mrec in st.session_state["manual_records"]:
        uid = mrec["User ID"]
        if uid not in st.session_state["corrections"]:
            st.session_state["corrections"][uid] = {
                f: mrec.get(f, "") for f in DISPLAY_FIELDS if f != "User ID"
            }

    for w in parse_warnings:
        banner("error", f"⚠ {w}")

    all_records     = records + st.session_state["manual_records"]
    visible_records = [
        r for r in all_records
        if r["User ID"] not in st.session_state["excluded_ids"]
    ]

    if not all_records:
        st.warning("No valid records found. Make sure each block starts with 'User ID: <number>'.")
        st.stop()

    uncertain = [
        rec for rec in visible_records
        if any(
            rec.get(f"_conf_{f}", HIGH) in (LOW, MEDIUM)
            for f in DISPLAY_FIELDS if f != "User ID"
        )
    ]
    n_flagged = len(uncertain)

    stat_row(
        stat_card(len(visible_records), "Players"),
        stat_card(n_flagged, "Need Review", warn=bool(n_flagged)),
        stat_card(len(st.session_state["excluded_ids"]), "Excluded"),
    )

    if uncertain:
        banner(
            "alert",
            f"<b>{n_flagged} record{'s' if n_flagged > 1 else ''} require attention</b> — "
            "some fields could not be parsed confidently. Review and correct below before exporting.",
        )

        for rec in uncertain:
            uid     = rec["User ID"]
            flagged = [
                f for f in DISPLAY_FIELDS
                if f != "User ID" and rec.get(f"_conf_{f}", HIGH) in (LOW, MEDIUM)
            ]
            label = f"User {uid}  ·  {len(flagged)} field{'s' if len(flagged) > 1 else ''} flagged"
            if rec.get("_manual"):
                label += "  [manual]"

            with st.expander(label, expanded=True):
                for f in flagged:
                    if rec.get(f"_warn_{f}"):
                        banner("info", f"ℹ {rec[f'_warn_{f}']}")

                if "_fc_raw" in rec and any(f in flagged for f in ("FCs", "FC Shards")):
                    st.markdown(
                        f'<div class="fc-raw">FC line: <span>{rec["_fc_raw"]}</span></div>',
                        unsafe_allow_html=True,
                    )

                n_cols  = min(len(flagged), 3)
                cols_w  = st.columns(n_cols)
                for i, field in enumerate(flagged):
                    conf = rec.get(f"_conf_{field}", HIGH)
                    with cols_w[i % n_cols]:
                        new_val = st.text_input(
                            label=field,
                            value=st.session_state["corrections"].get(uid, {}).get(field, ""),
                            key=f"fix_{uid}_{field}",
                            placeholder=FIELD_HINTS.get(field, ""),
                        )
                        st.markdown(
                            f'<div class="field-warn">{FIELD_WARN_MSG.get(conf, "")}</div>',
                            unsafe_allow_html=True,
                        )
                        st.session_state["corrections"].setdefault(uid, {})[field] = new_val

    with st.expander("➕ Add player manually"):
        with st.form("manual_add_form", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                m_uid    = st.text_input("User ID *",    placeholder="e.g. 99999")
                m_level  = st.text_input("Level",        placeholder="e.g. FC3")
                m_constr = st.text_input("Construction", placeholder="e.g. 24d 3h")
            with c2:
                m_res    = st.text_input("Research",     placeholder="e.g. 42d")
                m_troops = st.text_input("Troops",       placeholder="e.g. 5d")
                m_fcs    = st.text_input("FCs",          placeholder="e.g. 2700")
            with c3:
                m_shards = st.text_input("FC Shards",    placeholder="e.g. 434")
                m_time   = st.text_input("Time UTC",     placeholder="e.g. 14:00-17:30")
                m_days   = st.text_input("Days",         placeholder="e.g. 1, 2")
            submitted = st.form_submit_button("Add Player", use_container_width=True)

        if submitted:
            uid_str = m_uid.strip()
            if not uid_str:
                st.error("User ID is required.")
            elif uid_str in [r["User ID"] for r in all_records]:
                st.error(f"User ID {uid_str} already exists.")
            else:
                constr_v, constr_c      = normalize_duration(m_constr)
                res_v,    res_c         = normalize_duration(m_res)
                troops_v, troops_c      = normalize_duration(m_troops)
                fc_raw                  = f"FC {m_fcs} shards {m_shards}" if (m_fcs or m_shards) else ""
                fc_v, fc_c, sh_v, sh_c = parse_fc_shards(fc_raw)
                time_v, time_c, scount  = normalize_time_utc(m_time)
                if time_v and scount < MIN_TIME_WINDOW_SLOTS:
                    time_c = MEDIUM
                days_v, days_c = normalize_days(m_days)

                new_rec = {
                    "User ID": uid_str,
                    "Level":        m_level.strip(),  "_conf_Level":        HIGH if m_level.strip() else LOW,
                    "Construction": constr_v,          "_conf_Construction": constr_c,
                    "Research":     res_v,             "_conf_Research":     res_c,
                    "Troops":       troops_v,          "_conf_Troops":       troops_c,
                    "FCs":          fc_v,              "_conf_FCs":          fc_c,
                    "FC Shards":    sh_v,              "_conf_FC Shards":    sh_c,
                    "Time UTC":     time_v,            "_conf_Time UTC":     time_c,
                    "Days":         days_v,            "_conf_Days":         days_c,
                    "_fc_raw":      fc_raw,
                    "_raw_block":   "(manual entry)",
                    "_manual":      True,
                }
                if time_v and scount < MIN_TIME_WINDOW_SLOTS:
                    new_rec["_warn_Time UTC"] = f"Only {scount / 2:.4g}h window — minimum is 3h"

                st.session_state["manual_records"].append(new_rec)
                st.session_state["corrections"][uid_str] = {
                    f: new_rec[f] for f in DISPLAY_FIELDS if f != "User ID"
                }
                st.success(f"Player {uid_str} added.")
                st.rerun()

    with st.expander("🗂 Manage records"):
        all_uids       = [r["User ID"] for r in all_records]
        manual_uid_set = {r["User ID"] for r in st.session_state["manual_records"]}

        col_del, col_clear = st.columns([3, 1])
        with col_del:
            to_exclude = st.multiselect(
                "Exclude from results and export",
                options=all_uids,
                default=[uid for uid in st.session_state["excluded_ids"] if uid in all_uids],
                format_func=lambda uid: f"User {uid}" + (" [manual]" if uid in manual_uid_set else ""),
            )
            st.session_state["excluded_ids"] = set(to_exclude)
        with col_clear:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Clear all manual", use_container_width=True):
                manual_ids = {r["User ID"] for r in st.session_state["manual_records"]}
                st.session_state["manual_records"] = []
                st.session_state["excluded_ids"]  -= manual_ids
                st.rerun()

        if st.session_state["excluded_ids"]:
            banner(
                "info",
                f"ℹ {len(st.session_state['excluded_ids'])} record(s) excluded from results and export.",
            )

    # ── Build the final DataFrame ──────────────────────────────────────────────
    final: list      = []
    conf_lookup: dict = {}

    for rec in visible_records:
        uid = rec["User ID"]
        row = {f: rec[f] for f in DISPLAY_FIELDS}
        row.update(st.session_state["corrections"].get(uid, {}))
        final.append(row)
        for f in DISPLAY_FIELDS:
            conf_lookup[(uid, f)] = rec.get(f"_conf_{f}", HIGH)

    df         = pd.DataFrame(final, columns=DISPLAY_FIELDS)
    df_display = df.replace("", "—")
    df_display[""] = [
        "✎" if any(r["User ID"] == uid and r.get("_manual") for r in all_records) else ""
        for uid in df_display["User ID"]
    ]
    df_display["_Raw Input"] = [
        next((rec["_raw_block"] for rec in all_records if rec["User ID"] == uid), "")
        for uid in df_display["User ID"]
    ]

    st.markdown('<div class="section-label" style="margin-top:0.5rem">Results</div>',
                unsafe_allow_html=True)
    st.dataframe(
        df_display,
        use_container_width=True,
        hide_index=True,
        column_config={
            "User ID":      st.column_config.TextColumn("User ID",       width="medium"),
            "Level":        st.column_config.TextColumn("Level",         width="small"),
            "Construction": st.column_config.NumberColumn("Construction", format="%.2f", width="small"),
            "Research":     st.column_config.NumberColumn("Research",     format="%.2f", width="small"),
            "Troops":       st.column_config.NumberColumn("Troops",       format="%.2f", width="small"),
            "FCs":          st.column_config.NumberColumn("FCs",          format="%d",   width="small"),
            "FC Shards":    st.column_config.NumberColumn("Shards",       format="%d",   width="small"),
            "Time UTC":     st.column_config.TextColumn("Time (UTC)",    width="large"),
            "Days":         st.column_config.TextColumn("Days",          width="small"),
            "":             st.column_config.TextColumn("",              width="small",
                                help="✎ = manually added record"),
            "_Raw Input":   st.column_config.TextColumn("Original Input",
                                help="Hover to see raw player input", width="medium"),
        },
    )

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown('<div class="section-label">Export &amp; send</div>', unsafe_allow_html=True)

    if st.button("🗓️  Send to Scheduler →", use_container_width=True, type="primary"):
        st.session_state["parsed_df"] = df.copy()
        st.session_state["page"]      = "scheduler"
        st.rerun()

    now           = datetime.utcnow().strftime("%Y%m%d_%H%M")
    df_export     = df[DISPLAY_FIELDS].copy()
    flagged_cells = {
        (uid, field)
        for (uid, field), conf in conf_lookup.items()
        if conf in (LOW, MEDIUM)
    }

    col_csv, col_xlsx, col_clip = st.columns(3)
    with col_csv:
        st.download_button(
            "⬇ CSV",
            data=df_export.to_csv(index=False).encode("utf-8"),
            file_name=f"svs_3442_{now}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with col_xlsx:
        if HAS_OPENPYXL:
            st.download_button(
                "⬇ Excel",
                data=build_excel(df_export, flagged_cells),
                file_name=f"svs_3442_{now}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            banner("info", "Install <code>openpyxl</code> to enable Excel export.")
    with col_clip:
        if st.button("📋 Copy CSV", use_container_width=True):
            st.session_state["_clipboard_csv"] = df_export.to_csv(index=False)

    if st.session_state["_clipboard_csv"]:
        st.markdown(
            '<div class="section-label" style="margin-top:0.75rem">CSV — click icon to copy</div>',
            unsafe_allow_html=True,
        )
        st.code(st.session_state["_clipboard_csv"], language=None)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — SCHEDULER
# ═══════════════════════════════════════════════════════════════════════════════

elif st.session_state["page"] == "scheduler":

    has_data = st.session_state["parsed_df"] is not None
    render_stepper([
        ("Paste player data",  "done"),
        ("Review &amp; correct", "done"),
        ("Send to Scheduler",  "done" if has_data else "idle"),
        ("Run &amp; export",   "active"),
    ])

    if not has_data:
        st.markdown("""
        <div class="empty-state">
          <div class="empty-state-icon">&#128197;</div>
          <div class="empty-state-title">No player data yet</div>
          <div class="empty-state-body">
            Go to the <b>Parser</b> page, paste your player responses, fix any flagged fields,
            then click <b>Send to Scheduler &#8594;</b> to load data here.
          </div>
        </div>
        """, unsafe_allow_html=True)

    SCHEDULER_SAMPLE = pd.DataFrame([
        {"User ID": 1,  "Level": "FC1", "Construction": 1.0,  "Research": 1.0,  "Troops": 1.0,  "FCs": 1200, "FC Shards": 210, "Time UTC": "16:00,16:30,17:00,17:30,18:00,18:30", "Days": "1,2"},
        {"User ID": 2,  "Level": "FC2", "Construction": 2.08, "Research": 2.0,  "Troops": 2.08, "FCs": 2100, "FC Shards": 380, "Time UTC": "07:00,07:30,08:00,08:30,09:00,09:30,10:00,10:30,11:00,11:30,12:00,12:30,13:00,13:30,14:00,14:30,15:00,15:30,16:00,16:30,17:00,17:30,18:00,18:30,19:00,19:30,20:00,20:30", "Days": "2,4"},
        {"User ID": 3,  "Level": "FC3", "Construction": 3.0,  "Research": 3.04, "Troops": 3.0,  "FCs": 2450, "FC Shards": 290, "Time UTC": "00:00,00:30,10:00,10:30,11:00,11:30,23:00,23:30", "Days": "1,2"},
        {"User ID": 4,  "Level": "FC4", "Construction": 4.0,  "Research": 4.17, "Troops": 4.0,  "FCs": 2600, "FC Shards": 410, "Time UTC": "14:00,14:30,15:00,15:30,16:00,16:30", "Days": "1,2,4"},
        {"User ID": 5,  "Level": "FC5", "Construction": 5.0,  "Research": 5.21, "Troops": 5.0,  "FCs": 2693, "FC Shards": 434, "Time UTC": "00:00,00:30,01:00,01:30,02:00,02:30,03:00,03:30,09:00,09:30,20:00,20:30,21:00,21:30,22:00,22:30,23:00,23:30", "Days": "1,4"},
        {"User ID": 6,  "Level": "FC2", "Construction": 2.5,  "Research": 2.3,  "Troops": 2.1,  "FCs": 1950, "FC Shards": 310, "Time UTC": "16:00,16:30,17:00,17:30", "Days": "1"},
        {"User ID": 7,  "Level": "FC3", "Construction": 3.3,  "Research": 3.1,  "Troops": 2.8,  "FCs": 2300, "FC Shards": 350, "Time UTC": "16:00,16:30,17:00,17:30", "Days": "1"},
        {"User ID": 8,  "Level": "FC1", "Construction": 1.1,  "Research": 0.9,  "Troops": 1.0,  "FCs": 900,  "FC Shards": 150, "Time UTC": "16:00,16:30,17:00,17:30", "Days": "1"},
        {"User ID": 9,  "Level": "FC2", "Construction": 2.0,  "Research": 1.8,  "Troops": 1.9,  "FCs": 1800, "FC Shards": 270, "Time UTC": "16:00,16:30,17:00,17:30", "Days": "1"},
        {"User ID": 10, "Level": "FC4", "Construction": 4.5,  "Research": 4.2,  "Troops": 4.3,  "FCs": 2550, "FC Shards": 420, "Time UTC": "16:00,16:30,17:00,17:30", "Days": "1"},
        {"User ID": 11, "Level": "FC3", "Construction": 3.7,  "Research": 3.5,  "Troops": 3.4,  "FCs": 2400, "FC Shards": 390, "Time UTC": "16:00,16:30,17:00,17:30", "Days": "1,2"},
        {"User ID": 12, "Level": "FC2", "Construction": 2.2,  "Research": 2.4,  "Troops": 2.0,  "FCs": 2050, "FC Shards": 330, "Time UTC": "06:00,06:30,07:00,07:30,08:00,08:30", "Days": "2,4"},
        {"User ID": 13, "Level": "FC5", "Construction": 5.3,  "Research": 5.1,  "Troops": 5.2,  "FCs": 2700, "FC Shards": 440, "Time UTC": "12:00,12:30,13:00,13:30,14:00,14:30", "Days": "1,2,4"},
        {"User ID": 14, "Level": "FC1", "Construction": 1.3,  "Research": 1.1,  "Troops": 1.4,  "FCs": 1100, "FC Shards": 190, "Time UTC": "22:00,22:30,23:00,23:30", "Days": "2"},
        {"User ID": 15, "Level": "FC4", "Construction": 4.1,  "Research": 3.9,  "Troops": 4.0,  "FCs": 2580, "FC Shards": 415, "Time UTC": "16:00,16:30,17:00,17:30", "Days": "1"},
    ])

    st.markdown('<div class="section-label">Data source</div>', unsafe_allow_html=True)

    source_options = ["Use parsed data from Parser", "Use built-in sample data", "Upload CSV / Excel"]
    default_src    = 0 if st.session_state["parsed_df"] is not None else 1
    source = st.radio(
        "Data source", source_options, index=default_src,
        horizontal=True, label_visibility="collapsed",
    )

    raw_df = None
    if source == "Use parsed data from Parser":
        if st.session_state["parsed_df"] is not None:
            raw_df = st.session_state["parsed_df"].copy()
            for col in ["Construction", "Research", "Troops"]:
                raw_df[col] = pd.to_numeric(raw_df[col], errors="coerce")
            n_before = len(raw_df)
            raw_df   = raw_df.dropna(subset=["Construction", "Research", "Troops"])
            n_dropped = n_before - len(raw_df)
            banner(
                "success",
                f"✓ Using {len(raw_df)} player(s) from Parser"
                + (f" — {n_dropped} skipped (missing numeric fields)" if n_dropped else ""),
            )
        else:
            banner("alert",
                   'No parsed data yet — go to the Parser page first, then click "Send to Scheduler".')

    elif source == "Use built-in sample data":
        raw_df = SCHEDULER_SAMPLE.copy()
        banner("info", "ℹ Using built-in sample dataset (15 users).")

    else:
        upload = st.file_uploader("Upload file", type=["csv", "xlsx"])
        if upload:
            raw_df = (
                pd.read_csv(upload) if upload.name.endswith(".csv")
                else pd.read_excel(upload)
            )
            banner("success", f"✓ Loaded {len(raw_df)} rows.")

    if raw_df is not None:
        with st.expander("Preview loaded data", expanded=False):
            st.dataframe(raw_df, use_container_width=True)

        st.markdown(
            '<div class="section-label" style="margin-top:1rem">Column mapping</div>',
            unsafe_allow_html=True,
        )
        cols = raw_df.columns.tolist()

        def pick(label: str, default: str) -> str:
            return st.selectbox(label, cols, index=cols.index(default) if default in cols else 0)

        c1, c2, c3 = st.columns(3)
        with c1:
            id_col  = pick("User ID column",     "User ID")
            lvl_col = pick("Level column",        "Level")
            fcs_col = pick("FCs column",          "FCs")
        with c2:
            con_col    = pick("Construction column", "Construction")
            res_col    = pick("Research column",     "Research")
            trp_col    = pick("Troops column",       "Troops")
            shards_col = pick("FC Shards column",    "FC Shards")
        with c3:
            time_default = next(
                (c for c in ("Time", "Hours") if c in cols), cols[0]
            )
            time_col = pick("Time UTC / Hours column", time_default)
            days_col = pick("Days column", "Days")

        st.markdown("<hr>", unsafe_allow_html=True)

        banner("info", """
            <b>Scheduling objective:</b> Maximise the number of players assigned a slot.
            Among all solutions that place the same number of players, the highest-scoring
            players are preferred — guaranteed by min-cost max-flow (provably optimal score ranking).
            Any remaining unplaced high-scorers are mathematically unplaceable: their entire
            time window is saturated by even-higher-scorers with no alternative slots.
        """)

        if st.button("⚡ Run scheduler", type="primary", use_container_width=False):
            users = []
            for _, row in raw_df.iterrows():
                raw_time_val = str(row[time_col]) if pd.notna(row[time_col]) else ""
                hours = slots_str_to_hours(raw_time_val) if ":" in raw_time_val else parse_ints(raw_time_val)

                try:
                    con_val = float(row[con_col])
                    res_val = float(row[res_col])
                    trp_val = float(row[trp_col])
                except (ValueError, TypeError):
                    continue

                def _safe_int(col_name):
                    val = row[col_name]
                    try:
                        return int(float(val)) if pd.notna(val) and str(val).strip() not in ("", "—") else None
                    except (ValueError, TypeError):
                        return None

                users.append({
                    "User ID":      row[id_col],
                    "Level":        str(row[lvl_col]),
                    "Construction": con_val,
                    "Research":     res_val,
                    "Troops":       trp_val,
                    "FCs":          _safe_int(fcs_col),
                    "FC Shards":    _safe_int(shards_col),
                    "hours":        hours,
                    "days":         parse_ints(row[days_col]),
                })

            if not users:
                st.error("No valid users to schedule — check your column mapping and data.")
            else:
                with st.spinner("Running min-cost max-flow scheduler…"):
                    day_results = run_scheduler(users)
                    # Day 4 — VP: second pass for players unassigned in Day 4 MoE
                    day4_result = next((dr for dr in day_results if dr["day"] == 4), None)
                    vp4_result  = run_day4_vp(users, day4_result) if day4_result else None

                st.markdown('<div class="section-label" style="margin-top:1.5rem">Results</div>',
                            unsafe_allow_html=True)

                total_possible   = sum(len(dr["eligible"])  for dr in day_results)
                total_assigned   = sum(len(dr["user_slot"]) for dr in day_results)
                total_unassigned = sum(len(dr["unassigned"]) for dr in day_results)
                pct = round(100 * total_assigned / total_possible) if total_possible else 0

                stat_row(
                    stat_card(total_possible,   "Eligible slots-days"),
                    stat_card(total_assigned,   "Assigned"),
                    stat_card(f"{pct}%",        "Fill rate"),
                    stat_card(total_unassigned, "Unassignable", warn=bool(total_unassigned)),
                )

                 # --- 📊 NEW: GLOBAL AVAILABILITY CHART ---
                st.markdown("#### 📊 Peak Availability")
                st.caption("Total number of players available at each 30-minute time slot.")
                
                # Count available players per slot
                slot_counts = {s: 0 for s in range(48)}
                for u in users:
                    for s in _user_slots(u):
                        slot_counts[s] += 1
                        
                # Format for Streamlit chart
                chart_df = pd.DataFrame([
                    {"Time Slot": slot_label(s), "Players Available": count}
                    for s, count in slot_counts.items()
                ])
                
                # Display a cyan bar chart matching your theme
                st.bar_chart(chart_df.set_index("Time Slot"), color="#00e5cc")
                # ----------------------------------------- 

                st.markdown("#### 📋 User summary — all days")
                st.caption(
                    "Each cell shows the assigned slot and speedups. "
                    "❌ = window saturated (all their slots taken by higher-speedup players). "
                    "— = not participating on that day."
                )
                st.dataframe(build_summary_df(users, day_results, vp4_result), use_container_width=True, hide_index=True)

                st.markdown("#### 📅 Per-day detail")
                tab_labels = [dc["label"] for dc in DAY_CONFIG]
                if vp4_result is not None:
                    tab_labels.append(DAY4_VP_CONFIG["label"])
                tabs = st.tabs(tab_labels)
                for i, dc in enumerate(DAY_CONFIG):
                    dr = day_results[i]
                    with tabs[i]:
                        ca, cb, cc = st.columns(3)
                        ca.metric("Eligible users", len(dr["eligible"]))
                        cb.metric("Assigned",       len(dr["user_slot"]))
                        cc.metric("Unassigned",     len(dr["unassigned"]))

                        if dr["user_slot"] and dr["unassigned"]:
                            placed_scores   = [u[dc["col"]] for u in dr["eligible"] if u["User ID"] in dr["user_slot"]]
                            unplaced_scores = [e["user"][dc["col"]] for e in dr["unassigned"]]
                            min_p = min(placed_scores)
                            max_u = max(unplaced_scores)
                            if max_u > min_p + 0.001:
                                banner(
                                    "info",
                                    f"ℹ Some unassigned players have more speedups than the "
                                    f"lowest-placed player (highest unassigned: <b>{max_u:.2f}</b>, "
                                    f"lowest placed: <b>{min_p:.2f}</b>).",
                                )
                            else:
                                banner(
                                    "success",
                                    f"✓ Optimal speedup order: every placed player has ≥ speedups "
                                    f"than every unplaced player "
                                    f"(min placed {min_p:.2f} ≥ max unplaced {max_u:.2f}).",
                                )

                        st.markdown("**Slot timeline**")
                        extra_col = _DAY_EXTRA_COL.get(dc["day"])
                        timeline_df = build_timeline_df(users, dr, extra_col=extra_col)
                        col_cfg = {
                            "Speedups":  st.column_config.TextColumn(width="small"),
                            "Time Slot": st.column_config.TextColumn(width="medium"),
                            "Assigned":  st.column_config.TextColumn(width="medium"),
                        }
                        if extra_col:
                            col_cfg[extra_col] = st.column_config.TextColumn(extra_col, width="small")
                        st.dataframe(timeline_df, use_container_width=True, column_config=col_cfg)

                        ua_df = build_unassigned_df(dr)
                        if ua_df.empty:
                            banner("success", "✅ All eligible users assigned on this day.")
                        else:
                            st.markdown("**Unassigned users**")
                            st.dataframe(ua_df, use_container_width=True, hide_index=True)

                # ── Day 4 — VP tab ────────────────────────────────────────────────────
                if vp4_result is not None:
                    with tabs[len(DAY_CONFIG)]:
                        banner(
                            "info",
                            "ℹ <b>Day 4 — VP</b> gives players who were <b>unassigned in Day 4 — MoE</b> "
                            "a second-chance VP slot.",
                        )

                        vp4_eligible   = vp4_result["eligible"]
                        vp4_assigned   = vp4_result["user_slot"]
                        vp4_unassigned = vp4_result["unassigned"]

                        ca, cb, cc = st.columns(3)
                        ca.metric("Eligible (unassigned from MoE)", len(vp4_eligible))
                        cb.metric("Assigned VP slots",              len(vp4_assigned))
                        cc.metric("Still unassigned",               len(vp4_unassigned))

                        if not vp4_eligible:
                            banner("success", "✅ No unassigned players from Day 4 — MoE. Nothing to schedule.")
                        else:
                            if vp4_assigned and vp4_unassigned:
                                placed_scores   = [u["Troops"] for u in vp4_eligible if u["User ID"] in vp4_assigned]
                                unplaced_scores = [e["user"]["Troops"] for e in vp4_unassigned]
                                min_p = min(placed_scores)
                                max_u = max(unplaced_scores)
                                if max_u > min_p + 0.001:
                                    banner(
                                        "info",
                                        f"ℹ Some unassigned players have more speedups than the "
                                        f"lowest-placed player (highest unassigned: <b>{max_u:.2f}</b>, "
                                        f"lowest placed: <b>{min_p:.2f}</b>). "
                                        f"Their entire time window is saturated — the scheduler cannot "
                                        f"free a slot without dropping someone else.",
                                    )
                                else:
                                    banner(
                                        "success",
                                        f"✓ Optimal speedup order: every placed player has ≥ speedups "
                                        f"than every unplaced player "
                                        f"(min placed {min_p:.2f} ≥ max unplaced {max_u:.2f}).",
                                    )

                            st.markdown("**Slot timeline**")
                            vp4_timeline_df = build_timeline_df(users, vp4_result)
                            st.dataframe(
                                vp4_timeline_df,
                                use_container_width=True,
                                column_config={
                                    "Speedups":  st.column_config.TextColumn(width="small"),
                                    "Time Slot": st.column_config.TextColumn(width="medium"),
                                    "Assigned":  st.column_config.TextColumn(width="medium"),
                                },
                            )

                            vp4_ua_df = build_unassigned_df(vp4_result)
                            if vp4_ua_df.empty:
                                banner("success", "✅ All eligible players assigned a VP slot.")
                            else:
                                st.markdown("**Still unassigned after VP pass**")
                                st.dataframe(vp4_ua_df, use_container_width=True, hide_index=True)

                st.markdown("<hr>", unsafe_allow_html=True)
                st.markdown('<div class="section-label">Export</div>', unsafe_allow_html=True)
                st.download_button(
                    label="📥 Download full schedule (.xlsx)",
                    data=to_excel_schedule(users, day_results, vp4_result),
                    file_name="SvS_schedule.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
