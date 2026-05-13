import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

st.set_page_config(page_title="SvS #3442 tool", page_icon="⚔️", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

    html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }
    .stApp { background: #0e1117; color: #c9d1dc; }
    #MainMenu, footer, header { visibility: hidden; }

    h1, h2, h3, h4 {
        font-family: 'Plus Jakarta Sans', sans-serif;
        font-weight: 700; color: #e8ecf1; letter-spacing: -0.02em;
    }
    label { color: #9ba8bb !important; font-size: 0.8rem !important; font-weight: 500 !important;
            text-transform: uppercase; letter-spacing: 0.06em; }

    .page-header {
        display: flex; align-items: center; gap: 1rem;
        padding: 1.75rem 0 1.5rem; border-bottom: 1px solid #1e2533; margin-bottom: 2rem;
    }
    .page-header-icon {
        width: 40px; height: 40px; background: #1c2333; border: 1px solid #2a3348;
        border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 1.2rem;
    }
    .page-header-text h1 { font-size: 1.25rem; margin: 0; color: #e8ecf1; }
    .page-header-text p  { margin: 0.15rem 0 0; font-size: 0.82rem; color: #6b7a93; }

    .section-label {
        font-size: 0.72rem; font-weight: 600; text-transform: uppercase;
        letter-spacing: 0.1em; color: #5c6a82; margin-bottom: 0.6rem;
    }

    textarea {
        font-family: 'JetBrains Mono', monospace !important; font-size: 0.78rem !important;
        background: #131722 !important; color: #8fa3c0 !important;
        border: 1px solid #1e2a3a !important; border-radius: 8px !important; line-height: 1.65 !important;
    }
    textarea:focus { border-color: #334566 !important; box-shadow: 0 0 0 3px rgba(51,69,102,0.25) !important; }

    .stats-row { display: flex; gap: 0.75rem; margin-bottom: 1.75rem; }
    .stat-card {
        flex: 1; background: #131722; border: 1px solid #1e2533;
        border-radius: 10px; padding: 1rem 1.25rem;
    }
    .stat-card .stat-value { font-size: 1.6rem; font-weight: 700; color: #e8ecf1; line-height: 1; margin-bottom: 0.3rem; }
    .stat-card .stat-label { font-size: 0.72rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; color: #5c6a82; }
    .stat-card.warn .stat-value { color: #e09a3a; }

    .alert-banner {
        background: #161207; border: 1px solid #3a2e12; border-left: 3px solid #c8872a;
        border-radius: 8px; padding: 0.75rem 1rem; margin: 0.75rem 0; font-size: 0.85rem; color: #d4a855;
    }
    .error-banner {
        background: #160d0d; border: 1px solid #3a1a1a; border-left: 3px solid #b04040;
        border-radius: 8px; padding: 0.7rem 1rem; margin: 0.4rem 0; font-size: 0.82rem; color: #d07070;
    }
    .info-banner {
        background: #0d1520; border: 1px solid #1e3050; border-left: 3px solid #3a72c0;
        border-radius: 8px; padding: 0.7rem 1rem; margin: 0.4rem 0; font-size: 0.82rem; color: #7aaadc;
    }
    .field-warn { font-size: 0.73rem; color: #c8872a; margin-top: -0.35rem; margin-bottom: 0.5rem; }

    .streamlit-expanderHeader {
        background: #131722 !important; border: 1px solid #1e2533 !important;
        border-radius: 8px !important; color: #c9d1dc !important;
        font-size: 0.85rem !important; font-weight: 500 !important;
    }
    .streamlit-expanderContent {
        background: #0f1319 !important; border: 1px solid #1e2533 !important;
        border-top: none !important; border-radius: 0 0 8px 8px !important;
    }

    input[type="text"], .stTextInput input, .stNumberInput input {
        background: #131722 !important; border: 1px solid #1e2533 !important;
        border-radius: 7px !important; color: #c9d1dc !important;
        font-family: 'Plus Jakarta Sans', sans-serif !important; font-size: 0.85rem !important;
    }
    input:focus { border-color: #2e4066 !important; box-shadow: 0 0 0 3px rgba(46,64,102,0.3) !important; }

    .stSelectbox > div > div {
        background: #131722 !important; border: 1px solid #1e2533 !important;
        border-radius: 7px !important; color: #c9d1dc !important;
    }
    .stMultiSelect > div > div {
        background: #131722 !important; border: 1px solid #1e2533 !important; border-radius: 7px !important;
    }

    .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
        background: #1a2535; color: #b0c4de; border: 1px solid #2a3a54;
        border-radius: 7px; font-family: 'Plus Jakarta Sans', sans-serif;
        font-weight: 600; font-size: 0.85rem; letter-spacing: 0.02em;
        padding: 0.5rem 1.5rem; transition: all 0.15s ease;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: #1e2e45; border-color: #3a5070; color: #d0dcea;
    }

    .stDataFrame { border: 1px solid #1e2533 !important; border-radius: 10px !important; overflow: hidden; }

    .fc-raw {
        font-family: 'JetBrains Mono', monospace; font-size: 0.78rem; color: #6b7a93;
        margin-bottom: 0.9rem; padding: 0.5rem 0.75rem;
        background: #0e1117; border-radius: 6px; border: 1px solid #1a2230;
    }
    .fc-raw span { color: #7da8c8; }

    hr { border: none; border-top: 1px solid #1a2030; margin: 1.5rem 0; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════

NUM = r'\d+(?:[.,]\d+)?'
HIGH, MEDIUM, LOW = "high", "medium", "low"
MAX_REASONABLE_DAYS   = 365
MIN_TIME_WINDOW_HOURS = 3

DISPLAY_FIELDS = ["User ID", "Level", "Construction", "Research", "Troops",
                  "FCs", "FC Shards", "Time UTC", "Days"]

FIELD_HINTS = {
    "Level":        "e.g. FC3, FC5",
    "Construction": "e.g. 24d 3h  or  42d  or  35",
    "Research":     "e.g. 42d  or  20",
    "Troops":       "e.g. 100d 10h  or  50",
    "FCs":          "Number of FCs, e.g. 2693 or 2700",
    "FC Shards":    "Number of shards, e.g. 434",
    "Time UTC":     "e.g. 16:00–19:00  or  7–21",
    "Days":         "e.g. Mon, Thu  or  1, 4",
}

FIELD_WARN_MSG = {
    LOW:    "⚠ Could not parse — please verify",
    MEDIUM: "⚠ Parsed with low confidence — please check",
}

SAMPLE = """\
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


# ═══════════════════════════════════════════════════════════════════════════════
# PARSING ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

def normalize_duration(raw):
    """
    Convert duration strings to float days.
    - Handles ranges like "30-35d" → takes lower bound, returns MEDIUM confidence.
    - Values > MAX_REASONABLE_DAYS are downgraded to MEDIUM.
    """
    if not raw:
        return "", LOW

    raw = raw.strip().lower()
    days = 0.0

    # Range: "30-35d" or "30d - 35d" → lower bound, flag as MEDIUM
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

    num_m = re.search(rf'({NUM})', raw)
    if num_m:
        val = float(num_m.group(1).replace(",", "."))
        conf = MEDIUM if val > MAX_REASONABLE_DAYS else HIGH
        return val, conf

    return "", LOW


def parse_fc_shards(raw):
    """
    Parse FCs + shards.
    Correctly handles European thousand separators: 2.693 or 2,700 → 2693 / 2700.
    """
    if not raw:
        return "", LOW, "", LOW

    text = raw.strip().lower()
    fc_val    = None
    shard_val = None

    for pat in [r'shards?\s*[:\-]?\s*(\d+)', r'(\d+)\s*(?:fc\s+)?shards?']:
        m = re.search(pat, text, re.I)
        if m:
            shard_val = int(m.group(1))
            break

    for pat in [
        r'fcs?\s*[:\-]\s*(\d+(?:[.,]\d+)?)',
        r'fcs?\s+(\d+(?:[.,]\d+)?)',
        r'(\d+(?:[.,]\d+)?)\s*fcs?\b',
        r'(\d+(?:[.,]\d+)?)\s*crystals?',
    ]:
        m = re.search(pat, text, re.I)
        if m:
            val = m.group(1)
            # "2.693" or "2,700" with exactly 3 decimal places → thousand separator
            if "." in val and len(val.split(".")[-1]) == 3:
                val = val.replace(".", "")
            elif "," in val and len(val.split(",")[-1]) == 3:
                val = val.replace(",", "")
            else:
                val = val.replace(",", ".")
            try:
                fc_val = int(float(val))
            except ValueError:
                pass
            break

    return (
        fc_val    if fc_val    is not None else "",
        HIGH if fc_val    is not None else LOW,
        shard_val if shard_val is not None else "",
        HIGH if shard_val is not None else LOW,
    )


def normalize_time_utc(raw):
    """
    Returns (hours_str, confidence, hour_count).
    hour_count < MIN_TIME_WINDOW_HOURS triggers a validation warning.
    """
    if not raw:
        return "", LOW, 0

    text = raw.lower().strip()
    text = text.replace("till", "-").replace(" to ", "-")
    hours = set()

    range_pat = r'(\d{1,2})(?:[:.]\d{2})?\s*(?:utc)?\s*-\s*(\d{1,2})(?:[:.]\d{2})?\s*(?:utc)?'
    for m in re.finditer(range_pat, text):
        s, e = int(m.group(1)), int(m.group(2))
        if not (0 <= s <= 23 and 0 <= e <= 23):
            continue
        expanded = (list(range(s, 24)) + list(range(0, e + 1))
                    if s > e else list(range(s, e + 1)))
        hours.update(expanded)

    text_no_ranges = re.sub(range_pat, " ", text)
    for m in re.finditer(r'\b(\d{1,2})(?:[:.]\d{2})?\s*(?:utc)?\b', text_no_ranges):
        h = int(m.group(1))
        if 0 <= h <= 23:
            hours.add(h)

    if not hours:
        return "", LOW, 0

    return ",".join(map(str, sorted(hours))), HIGH, len(hours)


def normalize_days(raw):
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


def extract_field(block, patterns):
    for pat in patterns:
        m = re.search(pat, block, re.I | re.MULTILINE)
        if m:
            return m.group(1).strip()
    return ""


# Fuzzy patterns — catch common typos (Reseach, Contruction, Troop, etc.)
CONSTRUCTION_PATS = [
    r'CONSTRUCTION\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
    r'Constr[a-z]{2,10}\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
]
RESEARCH_PATS = [
    r'RESEARCH\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
    r'Rese?[a-z]{2,7}\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
]
TROOPS_PATS = [
    r'TROOPS?\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
    r'Troop[a-z]{0,4}\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
]


def parse_block(block):
    r = {}
    r["User ID"] = extract_field(block, [
        r'User\s*ID\s*[:\-]?\s*(\d+)', r'\bID\s*[:\-]?\s*(\d+)',
    ])
    r["Level"] = extract_field(block, [r'Level\s*[:\-]?\s*(\S+)', r'LVL\s*[:\-]?\s*(\S+)'])
    r["_conf_Level"] = HIGH if r["Level"] else LOW

    for field, pats in [
        ("Construction", CONSTRUCTION_PATS),
        ("Research",     RESEARCH_PATS),
        ("Troops",       TROOPS_PATS),
    ]:
        raw_val = extract_field(block, pats)
        val, conf = normalize_duration(raw_val)
        r[field] = val
        r[f"_conf_{field}"] = conf

    fc_line = extract_field(block, [
        r'(?:How many FCs[^:\n]*|FC[s]?\s*/\s*[Ss]hard[s]?[^:\n]*)\s*[:\-]?\s*([^\n]+)',
        r'FC[s]?\s+and[^:\n]*[:\-]?\s*([^\n]+)',
        r'(?:Crystal[s]?[^:\n]*)\s*[:\-]?\s*([^\n]+)',
    ])
    r["_fc_raw"] = fc_line
    fc_v, fc_c, sh_v, sh_c = parse_fc_shards(fc_line)
    r["FCs"] = fc_v;       r["_conf_FCs"] = fc_c
    r["FC Shards"] = sh_v; r["_conf_FC Shards"] = sh_c

    raw_time = extract_field(block, [
        r'Desired\s+time\s+UTC[^:\n]*[:\-]?\s*([^\n]+)',
        r'Time\s+UTC[^:\n]*[:\-]?\s*([^\n]+)',
        r'UTC\s*[:\-]\s*([^\n]+)',
    ]).strip()
    tv, tc, hour_count = normalize_time_utc(raw_time)
    r["Time UTC"] = tv
    if tv and hour_count < MIN_TIME_WINDOW_HOURS:
        r["_conf_Time UTC"] = MEDIUM
        r["_warn_Time UTC"] = f"Only {hour_count}h window — minimum is {MIN_TIME_WINDOW_HOURS}h"
    else:
        r["_conf_Time UTC"] = tc

    raw_days = extract_field(block, [
        r'Desired\s+day(?:\(s\))?\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
        r'Day(?:s)?\s*(?:\([^)]*\))?\s*[:\-]\s*([^\n]+)',
    ])
    dv, dc = normalize_days(raw_days)
    r["Days"] = dv; r["_conf_Days"] = dc
    return r


def parse_input(text):
    parts = re.split(r'(?=User\s*ID\s*[:\-]?\s*\d)', text.strip(), flags=re.I)
    parts = [p.strip() for p in parts if p.strip()]
    records, warnings = [], []
    seen_ids = {}

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
        rec["_manual"] = False
        records.append(rec)
    return records, warnings


# ═══════════════════════════════════════════════════════════════════════════════
# EXCEL BUILDER
# ═══════════════════════════════════════════════════════════════════════════════

def build_excel(df_export, flagged_cells):
    """
    df_export     — DataFrame with display columns only
    flagged_cells — set of (uid, col_name) that were LOW/MEDIUM confidence
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SvS Data"

    HEADER_FILL  = PatternFill("solid", fgColor="1A2535")
    FLAGGED_FILL = PatternFill("solid", fgColor="1E1608")
    EMPTY_FILL   = PatternFill("solid", fgColor="130E0E")
    THIN   = Side(style="thin", color="1E2533")
    BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

    cols = list(df_export.columns)

    # Header row
    for ci, col in enumerate(cols, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.font      = Font(bold=True, color="7AABDC", size=10)
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = BORDER
    ws.row_dimensions[1].height = 22

    # Data rows
    uid_col = cols.index("User ID") if "User ID" in cols else None
    for ri, (_, row) in enumerate(df_export.iterrows(), 2):
        uid = str(row["User ID"]) if uid_col is not None else None
        for ci, col in enumerate(cols, 1):
            val = row[col]
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

    # Auto column width
    for ci, col in enumerate(cols, 1):
        max_len = max(
            len(str(col)),
            *[len(str(r)) for r in df_export[col].fillna("")]
        )
        ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 4, 42)

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
# SESSION STATE INIT
# ═══════════════════════════════════════════════════════════════════════════════

if "raw_input"      not in st.session_state: st.session_state["raw_input"]      = SAMPLE
if "manual_records" not in st.session_state: st.session_state["manual_records"] = []
if "excluded_ids"   not in st.session_state: st.session_state["excluded_ids"]   = set()
if "corrections"    not in st.session_state: st.session_state["corrections"]    = {}
if "_last_hash"     not in st.session_state: st.session_state["_last_hash"]     = None
if "_clipboard_csv" not in st.session_state: st.session_state["_clipboard_csv"] = None


# ═══════════════════════════════════════════════════════════════════════════════
# HEADER
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="page-header">
  <div class="page-header-icon">⚔️</div>
  <div class="page-header-text">
    <h1>SvS Prep Ministry — #3442</h1>
    <p>Paste raw player data below. Uncertain fields are flagged for manual review before export.</p>
  </div>
</div>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# INPUT AREA
# ═══════════════════════════════════════════════════════════════════════════════

col_input, col_tools = st.columns([4, 1])

with col_input:
    st.markdown('<div class="section-label">Raw Input</div>', unsafe_allow_html=True)
    raw_text = st.text_area(
        "raw", height=340, label_visibility="collapsed", key="raw_input",
    )

with col_tools:
    st.markdown('<div class="section-label">Input File</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Load .txt", type=["txt"], label_visibility="collapsed")
    if uploaded:
        content = uploaded.read().decode("utf-8")
        st.session_state["raw_input"] = content
        st.rerun()

    st.download_button(
        "💾 Save raw input as .txt",
        data=st.session_state["raw_input"].encode("utf-8"),
        file_name="svs_input.txt",
        mime="text/plain",
        use_container_width=True,
    )

if not raw_text.strip():
    st.info("Paste your player data in the text box above.")
    st.stop()


# ═══════════════════════════════════════════════════════════════════════════════
# PARSE + CORRECTIONS MERGE
# ═══════════════════════════════════════════════════════════════════════════════

records, parse_warnings = parse_input(raw_text)

# Preserve manual corrections across re-parses — only reset entries for new UIDs
input_hash = hash(raw_text)
if st.session_state["_last_hash"] != input_hash:
    old_corr = st.session_state["corrections"]
    new_corr  = {}
    for rec in records:
        uid  = rec["User ID"]
        base = {f: rec[f] for f in DISPLAY_FIELDS if f != "User ID"}
        # Keep any previously saved corrections for this UID
        if uid in old_corr:
            base.update(old_corr[uid])
        new_corr[uid] = base
    # Keep manual record corrections untouched
    for mrec in st.session_state["manual_records"]:
        uid = mrec["User ID"]
        if uid in old_corr:
            new_corr[uid] = old_corr[uid]
    st.session_state["corrections"] = new_corr
    st.session_state["_last_hash"]  = input_hash

# Ensure manual records always have correction entries
for mrec in st.session_state["manual_records"]:
    uid = mrec["User ID"]
    if uid not in st.session_state["corrections"]:
        st.session_state["corrections"][uid] = {
            f: mrec.get(f, "") for f in DISPLAY_FIELDS if f != "User ID"
        }

for w in parse_warnings:
    st.markdown(f'<div class="error-banner">⚠ {w}</div>', unsafe_allow_html=True)

all_records     = records + st.session_state["manual_records"]
visible_records = [r for r in all_records
                   if r["User ID"] not in st.session_state["excluded_ids"]]

if not all_records:
    st.warning("No valid records found. Make sure each block starts with 'User ID: <number>'.")
    st.stop()


# ═══════════════════════════════════════════════════════════════════════════════
# STATS + VERIFICATION PANELS
# ═══════════════════════════════════════════════════════════════════════════════

uncertain = [
    rec for rec in visible_records
    if any(rec.get(f"_conf_{f}", HIGH) in (LOW, MEDIUM)
           for f in DISPLAY_FIELDS if f != "User ID")
]
n_flagged = len(uncertain)

st.markdown(f"""
<div class="stats-row">
  <div class="stat-card">
    <div class="stat-value">{len(visible_records)}</div>
    <div class="stat-label">Players</div>
  </div>
  <div class="stat-card {'warn' if n_flagged else ''}">
    <div class="stat-value">{n_flagged}</div>
    <div class="stat-label">Need Review</div>
  </div>
  <div class="stat-card">
    <div class="stat-value">{len(st.session_state['excluded_ids'])}</div>
    <div class="stat-label">Excluded</div>
  </div>
</div>
""", unsafe_allow_html=True)

if uncertain:
    st.markdown(f"""
    <div class="alert-banner">
      <b>{n_flagged} record{'s' if n_flagged > 1 else ''} require attention</b> —
      some fields could not be parsed confidently. Review and correct below before exporting.
    </div>""", unsafe_allow_html=True)

    for rec in uncertain:
        uid     = rec["User ID"]
        flagged = [f for f in DISPLAY_FIELDS if f != "User ID"
                   and rec.get(f"_conf_{f}", HIGH) in (LOW, MEDIUM)]
        label   = f"User {uid}  ·  {len(flagged)} field{'s' if len(flagged) > 1 else ''} flagged"
        if rec.get("_manual"):
            label += "  [manual]"

        with st.expander(label, expanded=True):
            for f in flagged:
                extra = rec.get(f"_warn_{f}")
                if extra:
                    st.markdown(f'<div class="info-banner">ℹ {extra}</div>', unsafe_allow_html=True)

            if "_fc_raw" in rec and any(f in flagged for f in ("FCs", "FC Shards")):
                st.markdown(
                    f'<div class="fc-raw">FC line: <span>{rec["_fc_raw"]}</span></div>',
                    unsafe_allow_html=True,
                )

            cols_w = st.columns(min(len(flagged), 3))
            for i, field in enumerate(flagged):
                conf = rec.get(f"_conf_{field}", HIGH)
                with cols_w[i % min(len(flagged), 3)]:
                    new_val = st.text_input(
                        label=field,
                        value=st.session_state["corrections"].get(uid, {}).get(field, ""),
                        key=f"fix_{uid}_{field}",
                        placeholder=FIELD_HINTS.get(field, ""),
                    )
                    st.markdown(f'<div class="field-warn">{FIELD_WARN_MSG.get(conf, "")}</div>',
                                unsafe_allow_html=True)
                    if uid not in st.session_state["corrections"]:
                        st.session_state["corrections"][uid] = {}
                    st.session_state["corrections"][uid][field] = new_val


# ═══════════════════════════════════════════════════════════════════════════════
# ADD PLAYER MANUALLY
# ═══════════════════════════════════════════════════════════════════════════════

with st.expander("➕ Add player manually"):
    with st.form("manual_add_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            m_uid    = st.text_input("User ID *",     placeholder="e.g. 99999")
            m_level  = st.text_input("Level",         placeholder="e.g. FC3")
            m_constr = st.text_input("Construction",  placeholder="e.g. 24d 3h")
        with c2:           
            m_res    = st.text_input("Research",      placeholder="e.g. 42d")
            m_troops = st.text_input("Troops",        placeholder="e.g. 5d")
            m_fcs    = st.text_input("FCs",           placeholder="e.g. 2700") 
        with c3:
            m_shards = st.text_input("FC Shards",     placeholder="e.g. 434")
            m_time   = st.text_input("Time UTC",      placeholder="e.g. 14-17")
            m_days   = st.text_input("Days",          placeholder="e.g. 1, 2")

        submitted = st.form_submit_button("Add Player", use_container_width=True)

    if submitted:
        uid_str = m_uid.strip()
        if not uid_str:
            st.error("User ID is required.")
        elif uid_str in [r["User ID"] for r in all_records]:
            st.error(f"User ID {uid_str} already exists.")
        else:
            constr_v, constr_c   = normalize_duration(m_constr)
            res_v,    res_c      = normalize_duration(m_res)
            troops_v, troops_c   = normalize_duration(m_troops)
            fc_raw               = f"FC {m_fcs} shards {m_shards}" if (m_fcs or m_shards) else ""
            fc_v, fc_c, sh_v, sh_c = parse_fc_shards(fc_raw)
            time_v, time_c, hcount = normalize_time_utc(m_time)
            if time_v and hcount < MIN_TIME_WINDOW_HOURS:
                time_c = MEDIUM
            days_v, days_c = normalize_days(m_days)

            new_rec = {
                "User ID": uid_str,
                "Level": m_level.strip(),    "_conf_Level": HIGH if m_level.strip() else LOW,
                "Construction": constr_v,    "_conf_Construction": constr_c,
                "Research": res_v,           "_conf_Research": res_c,
                "Troops": troops_v,          "_conf_Troops": troops_c,
                "FCs": fc_v,                 "_conf_FCs": fc_c,
                "FC Shards": sh_v,           "_conf_FC Shards": sh_c,
                "Time UTC": time_v,          "_conf_Time UTC": time_c,
                "Days": days_v,              "_conf_Days": days_c,
                "_fc_raw": fc_raw,
                "_raw_block": "(manual entry)",
                "_manual": True,
            }
            if time_v and hcount < MIN_TIME_WINDOW_HOURS:
                new_rec["_warn_Time UTC"] = f"Only {hcount}h window — minimum is {MIN_TIME_WINDOW_HOURS}h"

            st.session_state["manual_records"].append(new_rec)
            st.session_state["corrections"][uid_str] = {
                f: new_rec[f] for f in DISPLAY_FIELDS if f != "User ID"
            }
            st.success(f"Player {uid_str} added.")
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
# MANAGE RECORDS
# ═══════════════════════════════════════════════════════════════════════════════

with st.expander("🗂 Manage records"):
    all_uids = [r["User ID"] for r in all_records]

    col_del, col_clear = st.columns([3, 1])
    with col_del:
        manual_uid_set = {r["User ID"] for r in st.session_state["manual_records"]}
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
            st.session_state["excluded_ids"] -= manual_ids
            st.rerun()

    if st.session_state["excluded_ids"]:
        st.markdown(
            f'<div class="info-banner">ℹ {len(st.session_state["excluded_ids"])} '
            f'record(s) excluded from results and export.</div>',
            unsafe_allow_html=True,
        )


# ═══════════════════════════════════════════════════════════════════════════════
# RESULTS TABLE
# ═══════════════════════════════════════════════════════════════════════════════

final       = []
conf_lookup = {}   # {(uid, field): confidence} — used for Excel highlighting

for rec in visible_records:
    uid = rec["User ID"]
    row = {f: rec[f] for f in DISPLAY_FIELDS}
    row.update(st.session_state["corrections"].get(uid, {}))
    final.append(row)
    for f in DISPLAY_FIELDS:
        conf_lookup[(uid, f)] = rec.get(f"_conf_{f}", HIGH)

df         = pd.DataFrame(final, columns=DISPLAY_FIELDS)
df_display = df.replace("", "—")

df_display[""] = [           # tiny "manual" marker column
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
        "Time UTC":     st.column_config.TextColumn("Time (UTC)",    width="medium"),
        "Days":         st.column_config.TextColumn("Days",          width="small"),
        "":             st.column_config.TextColumn("",              width="small",
                            help="✎ = manually added record"),
        "_Raw Input":   st.column_config.TextColumn("Original Input",
                            help="Hover to see raw player input", width="medium"),
    },
)


# ═══════════════════════════════════════════════════════════════════════════════
# EXPORT
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown('<hr>', unsafe_allow_html=True)
st.markdown('<div class="section-label">Export</div>', unsafe_allow_html=True)

now       = datetime.utcnow().strftime("%Y%m%d_%H%M")
df_export = df[DISPLAY_FIELDS].copy()          # clean values, no "—"

# Build flagged cell set for Excel (uid, col_name)
flagged_cells = {
    (uid, field)
    for (uid, field), conf in conf_lookup.items()
    if conf in (LOW, MEDIUM)
}

col_csv, col_xlsx, col_clip = st.columns(3)

with col_csv:
    csv_data = df_export.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇ Download CSV",
        data=csv_data,
        file_name=f"svs_3442_{now}.csv",
        mime="text/csv",
        use_container_width=True,
    )

with col_xlsx:
    if HAS_OPENPYXL:
        excel_bytes = build_excel(df_export, flagged_cells)
        st.download_button(
            "⬇ Download Excel",
            data=excel_bytes,
            file_name=f"svs_3442_{now}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.markdown(
            '<div class="info-banner">Install <code>openpyxl</code> to enable Excel export.</div>',
            unsafe_allow_html=True,
        )

with col_clip:
    if st.button("📋 Copy CSV to clipboard", use_container_width=True):
        st.session_state["_clipboard_csv"] = df_export.to_csv(index=False)

if st.session_state["_clipboard_csv"]:
    st.markdown('<div class="section-label" style="margin-top:0.75rem">CSV — click icon to copy</div>',
                unsafe_allow_html=True)
    st.code(st.session_state["_clipboard_csv"], language=None)
