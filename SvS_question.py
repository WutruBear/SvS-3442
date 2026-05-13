import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="SvS #3442 tool", page_icon="⚔️", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;600;700&family=Share+Tech+Mono&display=swap');
    html, body, [class*="css"] { font-family: 'Rajdhani', sans-serif; }
    .stApp { background: #0d0f1a; color: #e0e6f0; }
    h1, h2, h3 { font-family: 'Rajdhani', sans-serif; font-weight: 700; color: #7eb8f7; letter-spacing: 2px; }
    .header-block {
        background: linear-gradient(135deg, #111827 0%, #1a2340 100%);
        border: 1px solid #2a3a5c; border-left: 4px solid #4a9eff;
        border-radius: 8px; padding: 1.5rem 2rem; margin-bottom: 2rem;
    }
    .header-block h1 { margin: 0; font-size: 2.2rem; }
    .header-block p  { margin: 0.3rem 0 0; color: #7a90b8; font-size: 1rem; }
    textarea { font-family: 'Share Tech Mono', monospace !important; font-size: 0.82rem !important;
               background: #111827 !important; color: #a8c8ff !important;
               border: 1px solid #2a3a5c !important; border-radius: 6px !important; }
    .metric-row { display: flex; gap: 1rem; margin-bottom: 1.5rem; }
    .metric-card { flex: 1; background: #111827; border: 1px solid #2a3a5c;
                   border-radius: 8px; padding: 1rem 1.5rem; text-align: center; }
    .metric-card .val { font-size: 2rem; font-weight: 700; color: #4a9eff; }
    .metric-card .lbl { font-size: 0.85rem; color: #7a90b8; text-transform: uppercase; letter-spacing: 1px; }
    .stButton > button {
        background: linear-gradient(135deg, #1e3a6e, #2a5298); color: #fff;
        border: 1px solid #4a9eff; border-radius: 6px;
        font-family: 'Rajdhani', sans-serif; font-weight: 600;
        font-size: 1rem; letter-spacing: 1px; padding: 0.5rem 2rem;
    }
    .verify-banner {
        background: #1a1a0d; border: 1px solid #7a6020; border-left: 4px solid #f0b040;
        border-radius: 8px; padding: 0.8rem 1.2rem; margin: 0.6rem 0; color: #f0d080;
    }
    .field-warn { color: #f0b040; font-size: 0.82rem; margin-top: -0.5rem; margin-bottom: 0.5rem; }
    label { color: #a8c8ff !important; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PARSING ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

NUM = r'\d+(?:[.,]\d+)?'
HIGH, MEDIUM, LOW = "high", "medium", "low"


def normalize_duration(raw):
    """
    Convert durations to float days.

    Examples:
        30d -> 30
        24d 12h -> 24.5
        12h -> 0.5
        90m -> 0.0625
    """
    if not raw:
        return "", LOW

    raw = raw.strip().lower()

    days = 0.0

    day_m = re.search(rf'({NUM})\s*(?:d|day|days)\b', raw)
    hour_m = re.search(rf'({NUM})\s*(?:h|hr|hour|hours)\b', raw)
    min_m = re.search(rf'({NUM})\s*(?:m|min|minute|minutes)\b', raw)

    if day_m:
        days += float(day_m.group(1).replace(",", "."))

    if hour_m:
        days += float(hour_m.group(1).replace(",", ".")) / 24

    if min_m:
        days += float(min_m.group(1).replace(",", ".")) / 1440

    if days > 0:
        return round(days, 2), HIGH

    # fallback: plain number means days
    num_m = re.search(rf'({NUM})', raw)
    if num_m:
        return float(num_m.group(1).replace(",", ".")), HIGH

    return "", LOW


def parse_fc_shards(raw):
    """
    Parse FCs + shards as integers.

    Examples:
        FC 2700 shards 400
        2700 FC, 400 FC shards
        2.693 FC, 434 shards
        26 FC, 434 FCs
    """

    if not raw:
        return "", LOW, "", LOW

    text = raw.strip().lower()

    fc_val = None
    shard_val = None

    # ─────────────────────────────────────────────
    # SHARDS FIRST (more specific)
    # ─────────────────────────────────────────────
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

    # ─────────────────────────────────────────────
    # FCs
    # ─────────────────────────────────────────────
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

            # 2.693 -> 2693
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
    """
    Convert mixed UTC time expressions into expanded hour lists.

    Examples:
        12-14 -> 12,13,14
        16.00-19.00 -> 16,17,18,19
        7utc till 21utc -> 7,8,...,21
        00utc - 4utc, 9utc, 20-23
            -> 0,1,2,3,4,9,20,21,22,23
    """

    if not raw:
        return "", LOW

    text = raw.lower().strip()

    # Normalize wording
    text = text.replace("till", "-")
    text = text.replace("to", "-")

    hours = set()

    # ─────────────────────────────────────────────
    # FIND RANGES FIRST
    # ─────────────────────────────────────────────
    range_pattern = r'(\d{1,2})(?:[:.]\d{2})?\s*(?:utc)?\s*-\s*(\d{1,2})(?:[:.]\d{2})?\s*(?:utc)?'

    for m in re.finditer(range_pattern, text):
        start = int(m.group(1))
        end = int(m.group(2))

        if not (0 <= start <= 23 and 0 <= end <= 23):
            continue

        # Overnight range
        if start > end:
            expanded = list(range(start, 24)) + list(range(0, end + 1))
        else:
            expanded = list(range(start, end + 1))

        hours.update(expanded)

    # Remove ranges from text so standalone hours
    # don't duplicate incorrectly
    text_no_ranges = re.sub(range_pattern, " ", text)

    # ─────────────────────────────────────────────
    # FIND STANDALONE HOURS
    # ─────────────────────────────────────────────
    single_pattern = r'\b(\d{1,2})(?:[:.]\d{2})?\s*(?:utc)?\b'

    for m in re.finditer(single_pattern, text_no_ranges):
        hour = int(m.group(1))

        if 0 <= hour <= 23:
            hours.add(hour)

    if not hours:
        return "", LOW

    cleaned = sorted(hours)

    return ",".join(map(str, cleaned)), HIGH


def normalize_days(raw):
    """
    Convert days into numbers.

    Monday -> 1
    Tuesday -> 2
    ...
    Sunday -> 7
    """

    if not raw:
        return "", LOW

    text = raw.lower()

    day_map = {
        "mon": "1", "monday": "1",
        "tue": "2", "tuesday": "2",
        "wed": "3", "wednesday": "3",
        "thu": "4", "thursday": "4",
        "fri": "5", "friday": "5",
        "sat": "6", "saturday": "6",
        "sun": "7", "sunday": "7",
    }

    found = []

    # explicit numbers
    nums = re.findall(r'\b([1-7])\b', text)
    found.extend(nums)

    # text days
    for k, v in day_map.items():
        if re.search(rf'\b{k}\b', text):
            found.append(v)

    found = sorted(set(found), key=int)

    if found:
        return ",".join(found), HIGH

    return "", LOW


def extract_field(block, patterns):
    for pat in patterns:
        m = re.search(pat, block, re.I | re.MULTILINE)
        if m:
            return m.group(1).strip()
    return ""


def parse_block(block):
    r = {}
    r["User ID"] = extract_field(block, [
        r'User\s*ID\s*[:\-]?\s*(\d+)', r'\bID\s*[:\-]?\s*(\d+)',
    ])
    r["Level"] = extract_field(block, [r'Level\s*[:\-]?\s*(\S+)', r'LVL\s*[:\-]?\s*(\S+)'])
    r["_conf_Level"] = HIGH if r["Level"] else LOW

    for field, pats in [
        ("Construction", [r'CONSTRUCTION\s*(?:\([^)]*\))?\s*[:\-]?\s*([^\n]+)',
                          r'Construction\s*[:\-]?\s*([^\n]+)']),
        ("Research",     [r'RESEARCH\s*(?:\([^)]*\))?\s*[:\-]?\s*([^\n]+)',
                          r'Research\s*[:\-]?\s*([^\n]+)']),
        ("Troops",       [r'TROOPS\s*(?:\([^)]*\))?\s*[:\-]?\s*([^\n]+)',
                          r'Troops\s*[:\-]?\s*([^\n]+)']),
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
    r["FCs"] = fc_v;         r["_conf_FCs"] = fc_c
    r["FC Shards"] = sh_v;   r["_conf_FC Shards"] = sh_c

    raw_time = extract_field(block, [
    r'Desired\s+time\s+UTC[^:\n]*[:\-]?\s*([^\n]+)',
    r'Time\s+UTC[^:\n]*[:\-]?\s*([^\n]+)',
    r'UTC\s*[:\-]\s*([^\n]+)',
    ]).strip()

    tv, tc = normalize_time_utc(raw_time)

    r["Time UTC"] = tv
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
    "FCs":          "Number of FCs (crystals), e.g. 2.693 or 2700",
    "FC Shards":    "Number of FC shards, e.g. 434",
    "Time UTC":     "e.g. 16:00-19:00  or  7-21",
    "Days":         "e.g. Mon, Thu  or  1, 4",
}

FIELD_WARN_MSG = {
    LOW:    "⚠️ Could not parse — please verify",
    MEDIUM: "🔶 Parsed with low confidence — please check",
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
<div class="header-block">
  <h1>⚔️ SvS Prep Ministry tool #3442</h1>
  <p>Paste raw player data.
  Uncertain fields are flagged for manual review.</p>
</div>
""", unsafe_allow_html=True)

col_input, col_opts = st.columns([3, 1])

with col_opts:
    st.markdown("### ⚙️ Options")
    show_empty = st.checkbox("Show empty fields as —", value=True)
    sort_col   = st.selectbox("Sort by", DISPLAY_FIELDS)

with col_input:
    st.markdown("### 📋 Raw Input")
    raw_text = st.text_area("raw", value=SAMPLE, height=340, label_visibility="collapsed")

if not raw_text.strip():
    st.info("Paste your player data in the text box above.")
    st.stop()

records, warnings = parse_input(raw_text)

if warnings:
    for w in warnings:
        st.markdown(f'<div style="background:#2a1a1a;border:1px solid #7a3030;border-radius:6px;padding:0.6rem 1rem;color:#f08080;margin:0.3rem 0">⚠ {w}</div>', unsafe_allow_html=True)

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

if uncertain:
    n = len(uncertain)
    st.markdown(f"""
    <div class="verify-banner">
      <b>⚠️ {n} record{'s' if n>1 else ''} need your attention</b> —
      some fields could not be parsed confidently. Please review below before exporting.
    </div>""", unsafe_allow_html=True)

    for rec in uncertain:
        uid = rec["User ID"]
        flagged = [f for f in DISPLAY_FIELDS if f != "User ID"
                   and rec.get(f"_conf_{f}", HIGH) in (LOW, MEDIUM)]
        with st.expander(f"🔶 User {uid} — {len(flagged)} field{'s' if len(flagged)>1 else ''} flagged", expanded=True):
            if "_fc_raw" in rec and any(f in flagged for f in ("FCs", "FC Shards")):
                st.markdown(
                    f'<div style="font-family:monospace;font-size:0.82rem;color:#7a9ac0;margin-bottom:0.8rem">'
                    f'Original FC line: <span style="color:#c0d8f8">{rec["_fc_raw"]}</span></div>',
                    unsafe_allow_html=True
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
                    st.markdown(f'<div class="field-warn">{FIELD_WARN_MSG.get(conf,"")}</div>', unsafe_allow_html=True)
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
                               key=lambda s: s.str.lower() if s.dtype==object else s)
except Exception:
    df_sorted = df

n_flagged = len(uncertain)
st.markdown(f"""
<div class="metric-row">
  <div class="metric-card"><div class="val">{len(records)}</div><div class="lbl">Players included</div></div>
  <div class="metric-card"><div class="val">{n_flagged}</div><div class="lbl">Need Review</div></div>
</div>
""", unsafe_allow_html=True)

df_sorted["_Raw Input"] = [
    next(
        (rec["_raw_block"] for rec in records if rec["User ID"] == uid),
        ""
    )
    for uid in df_sorted["User ID"]
]

st.markdown("### 📊 Results Table")

st.dataframe(
    df_sorted,
    use_container_width=True,
    hide_index=True,

    column_config={

        "User ID": st.column_config.TextColumn(
            "User ID",
            width="medium",
        ),

        "Level": st.column_config.TextColumn(
            "Level",
            width="small",
        ),

        "Construction": st.column_config.NumberColumn(
            "Construction",
            format="%.2f",
            width="small",
        ),

        "Research": st.column_config.NumberColumn(
            "Research",
            format="%.2f",
            width="small",
        ),

        "Troops": st.column_config.NumberColumn(
            "Troops",
            format="%.2f",
            width="small",
        ),

        "FCs": st.column_config.NumberColumn(
            "FCs",
            format="%d",
            width="small",
        ),

        "FC Shards": st.column_config.NumberColumn(
            "Shards",
            format="%d",
            width="small",
        ),

        "Time UTC": st.column_config.TextColumn(
            "Time (UTC)",
            width="medium",
        ),

        "Days": st.column_config.TextColumn(
            "Days",
            width="small",
        ),

        "_Raw Input": st.column_config.TextColumn(
            "Original Input",
            help="Hover cells to see original raw player input",
            width="medium",
        ),
    }
)

from datetime import datetime

st.markdown("### 💾 Export")

# Default filename with timestamp
now = datetime.utcnow().strftime("%Y%m%d_%H%M")
default_name = f"svs_3442_export_{now}.csv"

file_name = st.text_input(
    "Filename",
    value=default_name,
    help="Choose the exported CSV filename"
)

csv_data = df_sorted.to_csv(index=False).encode("utf-8")

st.download_button(
    label="⬇ Save CSV",
    data=csv_data,
    file_name=file_name,
    mime="text/csv",
    use_container_width=True,
)
