import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="SvS #3442 tool", page_icon="⚔️", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

    /* ── Reset & Base ──────────────────────────────────── */
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

    .stApp {
        background: #0e1117;
        color: #c9d1dc;
    }

    /* ── Remove Streamlit chrome ─────────────────────── */
    #MainMenu, footer, header { visibility: hidden; }

    /* ── Typography ───────────────────────────────────── */
    h1, h2, h3, h4 {
        font-family: 'Plus Jakarta Sans', sans-serif;
        font-weight: 700;
        color: #e8ecf1;
        letter-spacing: -0.02em;
    }

    label { color: #9ba8bb !important; font-size: 0.8rem !important; font-weight: 500 !important; text-transform: uppercase; letter-spacing: 0.06em; }

    /* ── Header ───────────────────────────────────────── */
    .page-header {
        display: flex;
        align-items: center;
        gap: 1rem;
        padding: 1.75rem 0 1.5rem;
        border-bottom: 1px solid #1e2533;
        margin-bottom: 2rem;
    }

    .page-header-icon {
        width: 40px; height: 40px;
        background: #1c2333;
        border: 1px solid #2a3348;
        border-radius: 10px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.2rem;
    }

    .page-header-text h1 {
        font-size: 1.25rem;
        margin: 0;
        color: #e8ecf1;
    }

    .page-header-text p {
        margin: 0.15rem 0 0;
        font-size: 0.82rem;
        color: #6b7a93;
    }

    /* ── Section labels ───────────────────────────────── */
    .section-label {
        font-size: 0.72rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        color: #5c6a82;
        margin-bottom: 0.6rem;
    }

    /* ── Textarea ─────────────────────────────────────── */
    textarea {
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.78rem !important;
        background: #131722 !important;
        color: #8fa3c0 !important;
        border: 1px solid #1e2a3a !important;
        border-radius: 8px !important;
        line-height: 1.65 !important;
    }

    textarea:focus {
        border-color: #334566 !important;
        box-shadow: 0 0 0 3px rgba(51, 69, 102, 0.25) !important;
    }

    /* ── Cards / Stats ────────────────────────────────── */
    .stats-row {
        display: flex;
        gap: 0.75rem;
        margin-bottom: 1.75rem;
    }

    .stat-card {
        flex: 1;
        background: #131722;
        border: 1px solid #1e2533;
        border-radius: 10px;
        padding: 1rem 1.25rem;
    }

    .stat-card .stat-value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #e8ecf1;
        line-height: 1;
        margin-bottom: 0.3rem;
    }

    .stat-card .stat-label {
        font-size: 0.72rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #5c6a82;
    }

    .stat-card.warn .stat-value { color: #e09a3a; }

    /* ── Alert / Banner ───────────────────────────────── */
    .alert-banner {
        background: #161207;
        border: 1px solid #3a2e12;
        border-left: 3px solid #c8872a;
        border-radius: 8px;
        padding: 0.75rem 1rem;
        margin: 0.75rem 0;
        font-size: 0.85rem;
        color: #d4a855;
    }

    .error-banner {
        background: #160d0d;
        border: 1px solid #3a1a1a;
        border-left: 3px solid #b04040;
        border-radius: 8px;
        padding: 0.7rem 1rem;
        margin: 0.4rem 0;
        font-size: 0.82rem;
        color: #d07070;
    }

    .field-warn {
        font-size: 0.73rem;
        color: #c8872a;
        margin-top: -0.35rem;
        margin-bottom: 0.5rem;
    }

    /* ── Expander (verification panels) ──────────────── */
    .streamlit-expanderHeader {
        background: #131722 !important;
        border: 1px solid #1e2533 !important;
        border-radius: 8px !important;
        color: #c9d1dc !important;
        font-size: 0.85rem !important;
        font-weight: 500 !important;
    }

    .streamlit-expanderContent {
        background: #0f1319 !important;
        border: 1px solid #1e2533 !important;
        border-top: none !important;
        border-radius: 0 0 8px 8px !important;
    }

    /* ── Inputs ───────────────────────────────────────── */
    input[type="text"], .stTextInput input {
        background: #131722 !important;
        border: 1px solid #1e2533 !important;
        border-radius: 7px !important;
        color: #c9d1dc !important;
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        font-size: 0.85rem !important;
    }

    input[type="text"]:focus {
        border-color: #2e4066 !important;
        box-shadow: 0 0 0 3px rgba(46, 64, 102, 0.3) !important;
    }

    /* ── Selectbox / Checkbox ─────────────────────────── */
    .stSelectbox > div > div {
        background: #131722 !important;
        border: 1px solid #1e2533 !important;
        border-radius: 7px !important;
        color: #c9d1dc !important;
    }

    /* ── Button ───────────────────────────────────────── */
    .stButton > button, .stDownloadButton > button {
        background: #1a2535;
        color: #b0c4de;
        border: 1px solid #2a3a54;
        border-radius: 7px;
        font-family: 'Plus Jakarta Sans', sans-serif;
        font-weight: 600;
        font-size: 0.85rem;
        letter-spacing: 0.02em;
        padding: 0.5rem 1.5rem;
        transition: all 0.15s ease;
    }

    .stButton > button:hover, .stDownloadButton > button:hover {
        background: #1e2e45;
        border-color: #3a5070;
        color: #d0dcea;
    }

    /* ── Dataframe ────────────────────────────────────── */
    .stDataFrame {
        border: 1px solid #1e2533 !important;
        border-radius: 10px !important;
        overflow: hidden;
    }

    /* ── FC raw line ──────────────────────────────────── */
    .fc-raw {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.78rem;
        color: #6b7a93;
        margin-bottom: 0.9rem;
        padding: 0.5rem 0.75rem;
        background: #0e1117;
        border-radius: 6px;
        border: 1px solid #1a2230;
    }
    .fc-raw span { color: #7da8c8; }

    /* ── Divider ──────────────────────────────────────── */
    hr { border: none; border-top: 1px solid #1a2030; margin: 1.5rem 0; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PARSING ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

NUM = r'\d+(?:[.,]\d+)?'
HIGH, MEDIUM, LOW = "high", "medium", "low"


def normalize_duration(raw):
    if not raw:
        return "", LOW

    raw = raw.strip().lower()
    days = 0.0

    day_m  = re.search(rf'({NUM})\s*(?:d|day|days)\b', raw)
    hour_m = re.search(rf'({NUM})\s*(?:h|hr|hour|hours)\b', raw)
    min_m  = re.search(rf'({NUM})\s*(?:m|min|minute|minutes)\b', raw)

    if day_m:  days += float(day_m.group(1).replace(",", "."))
    if hour_m: days += float(hour_m.group(1).replace(",", ".")) / 24
    if min_m:  days += float(min_m.group(1).replace(",", ".")) / 1440

    if days > 0:
        return round(days, 2), HIGH

    num_m = re.search(rf'({NUM})', raw)
    if num_m:
        return float(num_m.group(1).replace(",", ".")), HIGH

    return "", LOW


def parse_fc_shards(raw):
    if not raw:
        return "", LOW, "", LOW

    text = raw.strip().lower()
    fc_val = None
    shard_val = None

    shard_patterns = [
        r'shards?\s*(\d+)',
        r'(\d+)\s*(?:fc\s+)?shards?',
        r'shards:?\s*(\d+)',
    ]
    for pat in shard_patterns:
        m = re.search(pat, text, re.I)
        if m:
            shard_val = int(m.group(1))
            break

    fc_patterns = [
        r'fc\s*(\d+(?:[.,]\d+)?)',
        r'fc:\s*(\d+(?:[.,]\d+)?)',
        r'fcs:\s*(\d+(?:[.,]\d+)?)',
        r'(\d+(?:[.,]\d+)?)\s*fc\b',
        r'(\d+(?:[.,]\d+)?)\s*crystals?',
    ]
    for pat in fc_patterns:
        m = re.search(pat, text, re.I)
        if m:
            val = m.group(1)
            if "." in val and len(val.split(".")[-1]) == 3:
                val = val.replace(".", "")
            elif "," in val and len(val.split(",")[-1]) == 3:
                val = val.replace(",", "")
            else:
                val = float(val.replace(",", "."))
            fc_val = int(val)
            break

    return (
        fc_val if fc_val is not None else "",
        HIGH if fc_val is not None else LOW,
        shard_val if shard_val is not None else "",
        HIGH if shard_val is not None else LOW,
    )


def normalize_time_utc(raw):
    if not raw:
        return "", LOW

    text = raw.lower().strip()
    text = text.replace("till", "-").replace("to", "-")
    hours = set()

    range_pattern = r'(\d{1,2})(?:[:.]\d{2})?\s*(?:utc)?\s*-\s*(\d{1,2})(?:[:.]\d{2})?\s*(?:utc)?'
    for m in re.finditer(range_pattern, text):
        start, end = int(m.group(1)), int(m.group(2))
        if not (0 <= start <= 23 and 0 <= end <= 23):
            continue
        expanded = list(range(start, 24)) + list(range(0, end + 1)) if start > end else list(range(start, end + 1))
        hours.update(expanded)

    text_no_ranges = re.sub(range_pattern, " ", text)
    for m in re.finditer(r'\b(\d{1,2})(?:[:.]\d{2})?\s*(?:utc)?\b', text_no_ranges):
        hour = int(m.group(1))
        if 0 <= hour <= 23:
            hours.add(hour)

    if not hours:
        return "", LOW

    return ",".join(map(str, sorted(hours))), HIGH


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


def parse_block(block):
    r = {}
    r["User ID"] = extract_field(block, [r'User\s*ID\s*[:\-]?\s*(\d+)', r'\bID\s*[:\-]?\s*(\d+)'])
    r["Level"] = extract_field(block, [r'Level\s*[:\-]?\s*(\S+)', r'LVL\s*[:\-]?\s*(\S+)'])
    r["_conf_Level"] = HIGH if r["Level"] else LOW

    for field, pats in [
        ("Construction", [r'CONSTRUCTION\s*(?:\([^)]*\))?\s*[:\-]?\s*([^\n]+)', r'Construction\s*[:\-]?\s*([^\n]+)']),
        ("Research",     [r'RESEARCH\s*(?:\([^)]*\))?\s*[:\-]?\s*([^\n]+)',     r'Research\s*[:\-]?\s*([^\n]+)']),
        ("Troops",       [r'TROOPS\s*(?:\([^)]*\))?\s*[:\-]?\s*([^\n]+)',       r'Troops\s*[:\-]?\s*([^\n]+)']),
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
    tv, tc = normalize_time_utc(raw_time)
    r["Time UTC"] = tv; r["_conf_Time UTC"] = tc

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
    for i, part in enumerate(parts, 1):
        rec = parse_block(part)
        if not rec["User ID"]:
            warnings.append(f"Block {i}: Could not find User ID — skipped.")
            continue
        rec["_raw_block"] = part
        records.append(rec)
    return records, warnings


DISPLAY_FIELDS = ["User ID", "Level", "Construction", "Research", "Troops",
                  "FCs", "FC Shards", "Time UTC", "Days"]

FIELD_HINTS = {
    "Level":        "e.g. FC3, FC5",
    "Construction": "e.g. 24d 3h  or  42d  or  35",
    "Research":     "e.g. 42d  or  20",
    "Troops":       "e.g. 100d 10h  or  50",
    "FCs":          "Number of FCs, e.g. 2.693 or 2700",
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
# UI
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

col_input, col_opts = st.columns([3, 1])

with col_opts:
    st.markdown('<div class="section-label">Options</div>', unsafe_allow_html=True)
    show_empty = st.checkbox("Show empty fields as —", value=True)
    sort_col   = st.selectbox("Sort by", DISPLAY_FIELDS)

with col_input:
    st.markdown('<div class="section-label">Raw Input</div>', unsafe_allow_html=True)
    raw_text = st.text_area("raw", value=SAMPLE, height=340, label_visibility="collapsed")

if not raw_text.strip():
    st.info("Paste your player data in the text box above.")
    st.stop()

records, warnings = parse_input(raw_text)

if warnings:
    for w in warnings:
        st.markdown(f'<div class="error-banner">⚠ {w}</div>', unsafe_allow_html=True)

if not records:
    st.warning("No valid records found. Make sure each block starts with 'User ID: <number>'.")
    st.stop()

# ── Session state for corrections ─────────────────────────────────────────────
input_hash = hash(raw_text)
if st.session_state.get("_last_hash") != input_hash:
    st.session_state["corrections"] = {
        rec["User ID"]: {f: rec[f] for f in DISPLAY_FIELDS if f != "User ID"}
        for rec in records
    }
    st.session_state["_last_hash"] = input_hash

# ── Verification panels ────────────────────────────────────────────────────────
uncertain = [
    rec for rec in records
    if any(rec.get(f"_conf_{f}", HIGH) in (LOW, MEDIUM)
           for f in DISPLAY_FIELDS if f != "User ID")
]

n_flagged = len(uncertain)

# Stats
st.markdown(f"""
<div class="stats-row">
  <div class="stat-card">
    <div class="stat-value">{len(records)}</div>
    <div class="stat-label">Players</div>
  </div>
  <div class="stat-card {'warn' if n_flagged else ''}">
    <div class="stat-value">{n_flagged}</div>
    <div class="stat-label">Need Review</div>
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
        uid = rec["User ID"]
        flagged = [f for f in DISPLAY_FIELDS if f != "User ID"
                   and rec.get(f"_conf_{f}", HIGH) in (LOW, MEDIUM)]
        with st.expander(f"User {uid}  ·  {len(flagged)} field{'s' if len(flagged) > 1 else ''} flagged", expanded=True):
            if "_fc_raw" in rec and any(f in flagged for f in ("FCs", "FC Shards")):
                st.markdown(
                    f'<div class="fc-raw">FC line: <span>{rec["_fc_raw"]}</span></div>',
                    unsafe_allow_html=True,
                )
            cols = st.columns(min(len(flagged), 3))
            for i, field in enumerate(flagged):
                conf = rec.get(f"_conf_{field}", HIGH)
                with cols[i % min(len(flagged), 3)]:
                    new_val = st.text_input(
                        label=field,
                        value=st.session_state["corrections"][uid].get(field, ""),
                        key=f"fix_{uid}_{field}",
                        placeholder=FIELD_HINTS.get(field, ""),
                    )
                    st.markdown(f'<div class="field-warn">{FIELD_WARN_MSG.get(conf, "")}</div>', unsafe_allow_html=True)
                    st.session_state["corrections"][uid][field] = new_val

# ── Merge corrections & build display df ──────────────────────────────────────
final = []
for rec in records:
    uid = rec["User ID"]
    row = {f: rec[f] for f in DISPLAY_FIELDS}
    row.update(st.session_state["corrections"].get(uid, {}))
    final.append(row)

df = pd.DataFrame(final, columns=DISPLAY_FIELDS)
if show_empty:
    df = df.replace("", "—")

try:
    df_sorted = df.sort_values(sort_col, ascending=True,
                               key=lambda s: s.str.lower() if s.dtype == object else s)
except Exception:
    df_sorted = df

df_sorted["_Raw Input"] = [
    next((rec["_raw_block"] for rec in records if rec["User ID"] == uid), "")
    for uid in df_sorted["User ID"]
]

st.markdown('<div class="section-label" style="margin-top:0.5rem">Results</div>', unsafe_allow_html=True)

st.dataframe(
    df_sorted,
    use_container_width=True,
    hide_index=True,
    column_config={
        "User ID":      st.column_config.TextColumn("User ID",      width="medium"),
        "Level":        st.column_config.TextColumn("Level",        width="small"),
        "Construction": st.column_config.NumberColumn("Construction", format="%.2f", width="small"),
        "Research":     st.column_config.NumberColumn("Research",     format="%.2f", width="small"),
        "Troops":       st.column_config.NumberColumn("Troops",       format="%.2f", width="small"),
        "FCs":          st.column_config.NumberColumn("FCs",          format="%d",   width="small"),
        "FC Shards":    st.column_config.NumberColumn("Shards",       format="%d",   width="small"),
        "Time UTC":     st.column_config.TextColumn("Time (UTC)",   width="medium"),
        "Days":         st.column_config.TextColumn("Days",          width="small"),
        "_Raw Input":   st.column_config.TextColumn("Original Input",
                            help="Hover to see raw player input", width="medium"),
    },
)

# ── Export ─────────────────────────────────────────────────────────────────────
st.markdown('<hr>', unsafe_allow_html=True)
st.markdown('<div class="section-label">Export</div>', unsafe_allow_html=True)

from datetime import datetime
now = datetime.utcnow().strftime("%Y%m%d_%H%M")
default_name = f"svs_3442_export_{now}.csv"

file_name = st.text_input("Filename", value=default_name)
csv_data  = df_sorted.to_csv(index=False).encode("utf-8")

st.download_button(
    label="Download CSV",
    data=csv_data,
    file_name=file_name,
    mime="text/csv",
    use_container_width=True,
)
