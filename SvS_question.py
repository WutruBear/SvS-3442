import streamlit as st
import pandas as pd
import re

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(page_title="FC Data Parser", page_icon="⚔️", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;600;700&family=Share+Tech+Mono&display=swap');

    html, body, [class*="css"] { font-family: 'Rajdhani', sans-serif; }

    .stApp { background: #0d0f1a; color: #e0e6f0; }

    h1, h2, h3 { font-family: 'Rajdhani', sans-serif; font-weight: 700; color: #7eb8f7; letter-spacing: 2px; }

    .header-block {
        background: linear-gradient(135deg, #111827 0%, #1a2340 100%);
        border: 1px solid #2a3a5c;
        border-left: 4px solid #4a9eff;
        border-radius: 8px;
        padding: 1.5rem 2rem;
        margin-bottom: 2rem;
    }

    .header-block h1 { margin: 0; font-size: 2.2rem; }
    .header-block p  { margin: 0.3rem 0 0; color: #7a90b8; font-size: 1rem; }

    textarea { font-family: 'Share Tech Mono', monospace !important; font-size: 0.82rem !important;
               background: #111827 !important; color: #a8c8ff !important;
               border: 1px solid #2a3a5c !important; border-radius: 6px !important; }

    .stDataFrame { border: 1px solid #2a3a5c; border-radius: 8px; overflow: hidden; }

    .metric-row { display: flex; gap: 1rem; margin-bottom: 1.5rem; }
    .metric-card {
        flex: 1;
        background: #111827;
        border: 1px solid #2a3a5c;
        border-radius: 8px;
        padding: 1rem 1.5rem;
        text-align: center;
    }
    .metric-card .val { font-size: 2rem; font-weight: 700; color: #4a9eff; }
    .metric-card .lbl { font-size: 0.85rem; color: #7a90b8; text-transform: uppercase; letter-spacing: 1px; }

    .stButton > button {
        background: linear-gradient(135deg, #1e3a6e, #2a5298);
        color: #fff;
        border: 1px solid #4a9eff;
        border-radius: 6px;
        font-family: 'Rajdhani', sans-serif;
        font-weight: 600;
        font-size: 1rem;
        letter-spacing: 1px;
        padding: 0.5rem 2rem;
        transition: all 0.2s;
    }
    .stButton > button:hover { background: linear-gradient(135deg, #2a5298, #3a6bc8); }

    .parse-error { background: #2a1a1a; border: 1px solid #7a3030;
                   border-radius: 6px; padding: 0.8rem 1.2rem; margin: 0.4rem 0;
                   color: #f08080; font-size: 0.9rem; }
    .parse-ok    { background: #0d2a1a; border: 1px solid #2a6040;
                   border-radius: 6px; padding: 0.4rem 1rem; margin: 0.2rem 0;
                   color: #60c090; font-size: 0.88rem; }

    .stSelectbox div[data-baseweb="select"] { background: #111827 !important; border-color: #2a3a5c !important; }
    .stCheckbox { color: #a8c8ff; }
    label { color: #a8c8ff !important; }
</style>
""", unsafe_allow_html=True)


# ── Parser helpers ─────────────────────────────────────────────────────────────

def normalize_duration(raw: str) -> str:
    """Normalize duration strings like '24d 3h', '24 days', '42d', '35', '100d 10h' → uniform label."""
    if not raw:
        return ""
    raw = raw.strip()
    # already looks like a number (hours in game, maybe)
    days_match  = re.search(r'(\d+(?:\.\d+)?)\s*d(?:ay(?:s)?)?', raw, re.I)
    hours_match = re.search(r'(\d+(?:\.\d+)?)\s*h(?:our(?:s)?)?', raw, re.I)
    mins_match  = re.search(r'(\d+(?:\.\d+)?)\s*m(?:in(?:ute(?:s)?)?)?', raw, re.I)

    parts = []
    if days_match:  parts.append(f"{days_match.group(1)}d")
    if hours_match: parts.append(f"{hours_match.group(1)}h")
    if mins_match:  parts.append(f"{mins_match.group(1)}m")

    if parts:
        return " ".join(parts)

    # plain number — return as-is (probably speedup count)
    num_match = re.search(r'(\d+(?:\.\d+)?)', raw)
    return num_match.group(1) if num_match else raw


def parse_fc_shards(raw: str):
    """
    Handle many FC/crystal/shard formats:
      'FC 2700 shards 400'
      '2.693 FC, 434 FC shards'
      '2.693 Crystals, 434 shards'
      'FC 2 shards 400'
    Returns (fc_str, shards_str)
    """
    if not raw:
        return "", ""

    raw = raw.strip()

    # Normalise separator
    text = raw.replace(",", " ").replace(";", " ")

    # Look for FC count — number before or after 'FC' or 'Crystal(s)'
    fc_val = ""
    shard_val = ""

    fc_pattern     = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:FC|Crystal(?:s)?)\b', text, re.I)
    fc_pattern_rev = re.search(r'\bFC\s+(\d+(?:[.,]\d+)?)', text, re.I)

    if fc_pattern:
        fc_val = fc_pattern.group(1).replace(",", ".")
    elif fc_pattern_rev:
        fc_val = fc_pattern_rev.group(1)

    # Shards — prefer "shards <number>" pattern first, then "<number> [FC ]shards"
    # Use word boundary to avoid matching FC count (e.g. "FC 2700 shards 400")
    shard_after  = re.search(r'\bshard(?:s)?\s+(\d+(?:[.,]\d+)?)', text, re.I)
    shard_before = re.search(r'(\d+(?:[.,]\d+)?)\s+(?:FC\s+)?shard(?:s)?\b', text, re.I)

    if shard_after:
        shard_val = shard_after.group(1)
    elif shard_before:
        shard_val = shard_before.group(1)

    return fc_val, shard_val


def normalize_days(raw: str) -> str:
    """Convert day numbers or names to a consistent 'Mon, Thu' style."""
    if not raw:
        return ""

    day_map = {
        "1": "Mon", "monday": "Mon",
        "2": "Tue", "tuesday": "Tue",
        "3": "Wed", "wednesday": "Wed",
        "4": "Thu", "thursday": "Thu",
        "5": "Fri", "friday": "Fri",
        "6": "Sat", "saturday": "Sat",
        "7": "Sun", "sunday": "Sun",
    }

    tokens = re.split(r'[\s,;/]+', raw.lower())
    seen, result = set(), []
    for t in tokens:
        t = t.strip("().")
        # strip notes like "but not enough speedups"
        if re.match(r'[a-z]+', t) and t not in day_map:
            continue
        mapped = day_map.get(t)
        if mapped and mapped not in seen:
            result.append(mapped)
            seen.add(mapped)
    return ", ".join(result) if result else raw.strip()


def extract_field(block: str, patterns: list[str]) -> str:
    """Try multiple regex patterns against a block; return first match group 1."""
    for pat in patterns:
        m = re.search(pat, block, re.I | re.MULTILINE)
        if m:
            return m.group(1).strip()
    return ""


def parse_block(block: str) -> dict:
    """Parse a single user data block into a dict."""
    record = {}

    record["User ID"] = extract_field(block, [
        r'User\s*ID\s*[:\-]?\s*(\d+)',
        r'ID\s*[:\-]?\s*(\d+)',
    ])

    record["Level"] = extract_field(block, [
        r'Level\s*[:\-]?\s*(\S+)',
        r'LVL\s*[:\-]?\s*(\S+)',
    ])

    # Construction / Research / Troops — accept with or without day label
    record["Construction"] = normalize_duration(extract_field(block, [
        r'CONSTRUCTION\s*(?:\([^)]*\))?\s*[:\-]?\s*([^\n]+)',
        r'Construction\s*[:\-]?\s*([^\n]+)',
    ]))

    record["Research"] = normalize_duration(extract_field(block, [
        r'RESEARCH\s*(?:\([^)]*\))?\s*[:\-]?\s*([^\n]+)',
        r'Research\s*[:\-]?\s*([^\n]+)',
    ]))

    record["Troops"] = normalize_duration(extract_field(block, [
        r'TROOPS\s*(?:\([^)]*\))?\s*[:\-]?\s*([^\n]+)',
        r'Troops\s*[:\-]?\s*([^\n]+)',
    ]))

    # FC / Shards line
    fc_line = extract_field(block, [
        r'(?:How many FCs[^:\n]*|FC[s]?\s*/\s*[Ss]hard[s]?[^:\n]*)[:\-]?\s*([^\n]+)',
        r'FC[s]?\s+and[^:\n]*[:\-]?\s*([^\n]+)',
    ])
    fc_val, shard_val = parse_fc_shards(fc_line)
    record["FCs"]         = fc_val
    record["FC Shards"]   = shard_val

    record["Time UTC"] = extract_field(block, [
        r'Desired\s+time\s+UTC[^:\n]*[:\-]?\s*([^\n]+)',
        r'Time\s+UTC[^:\n]*[:\-]?\s*([^\n]+)',
        r'UTC\s*[:\-]\s*([^\n]+)',
    ]).strip()

    raw_days = extract_field(block, [
        r'Desired\s+day[s]?\s*(?:\([^)]*\))?\s*[:\-]?\s*([^\n]+)',
        r'Day[s]?\s*[:\-]?\s*([^\n]+)',
    ])
    record["Days"] = normalize_days(raw_days)

    return record


def parse_input(text: str) -> tuple[list[dict], list[str]]:
    """Split raw text into blocks and parse each. Returns (records, warnings)."""
    # Split on 'User ID' boundaries
    parts = re.split(r'(?=User\s*ID\s*[:\-]?\s*\d)', text.strip(), flags=re.I)
    parts = [p.strip() for p in parts if p.strip()]

    records, warnings = [], []
    for i, part in enumerate(parts, 1):
        rec = parse_block(part)
        if not rec["User ID"]:
            warnings.append(f"Block {i}: Could not find User ID — skipped.")
            continue
        records.append(rec)

    return records, warnings


# ── Sample data ────────────────────────────────────────────────────────────────
SAMPLE = """\
User ID: 574199347
Level: FC5
CONSTRUCTION (Monday): 35
RESEARCH (Tuesday):  20
TROOPS (Thursday): 50
How many FCs and FC shards you have: FC 2700 shards 400
Desired time UTC (minimum 3 hour window): 16.00-19.00
Desired day(s) (1,.2,.4): 1,4

User ID: 564134590
Level: FC3
CONSTRUCTION (Monday): 24d 3h
RESEARCH (Tuesday): 42d
TROOPS (Thursday): 100d 10h
How many FCs and FC shards you have: 2.693 FC, 434 FC shards 
Desired time UTC (minimum 3 hour window): 7utc till 21utc
Desired day(s): (1 but not enough speedups😓), 2, 4

User ID: 564134596
Level: FC3
CONSTRUCTION (Monday): 24 days
RESEARCH (Tuesday): 42 days
TROOPS (Thursday): 100 days
How many FCs and FC shards you have: 2.693 Crystals, 434 shards 
Desired time UTC (minimum 3 hour window): 10, 11, 23, 24
Desired day(s): Monday, Tuesday
"""

# ── Layout ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-block">
  <h1>⚔️ FC DATA PARSER</h1>
  <p>Paste raw player data in any format — the parser normalises everything into a clean table.</p>
</div>
""", unsafe_allow_html=True)

col_input, col_opts = st.columns([3, 1])

with col_opts:
    st.markdown("### ⚙️ Options")
    show_empty = st.checkbox("Show empty fields as —", value=True)
    sort_col   = st.selectbox("Sort by", ["User ID", "Level", "FCs", "FC Shards", "Construction", "Research", "Troops", "Days"])
    export_csv = st.button("⬇ Download CSV")

with col_input:
    st.markdown("### 📋 Raw Input")
    raw_text = st.text_area(
        label="Paste player data blocks below",
        value=SAMPLE,
        height=340,
        label_visibility="collapsed",
    )

# ── Parse & display ────────────────────────────────────────────────────────────
if raw_text.strip():
    records, warnings = parse_input(raw_text)

    if warnings:
        for w in warnings:
            st.markdown(f'<div class="parse-error">⚠ {w}</div>', unsafe_allow_html=True)

    if records:
        df = pd.DataFrame(records, columns=[
            "User ID", "Level", "Construction", "Research", "Troops",
            "FCs", "FC Shards", "Time UTC", "Days"
        ])

        if show_empty:
            df = df.replace("", "—")

        # Sort
        try:
            df_sorted = df.sort_values(sort_col, ascending=True, key=lambda s: s.str.lower() if s.dtype == object else s)
        except Exception:
            df_sorted = df

        # Metrics row
        st.markdown(f"""
        <div class="metric-row">
          <div class="metric-card"><div class="val">{len(records)}</div><div class="lbl">Players Parsed</div></div>
          <div class="metric-card"><div class="val">{df['Level'].nunique()}</div><div class="lbl">Unique Levels</div></div>
          <div class="metric-card"><div class="val">{len(warnings)}</div><div class="lbl">Parse Warnings</div></div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### 📊 Parsed Table")
        st.dataframe(
            df_sorted,
            use_container_width=True,
            hide_index=True,
            column_config={
                "User ID":      st.column_config.TextColumn("User ID",      width="medium"),
                "Level":        st.column_config.TextColumn("Level",        width="small"),
                "Construction": st.column_config.TextColumn("Construction 🏗", width="medium"),
                "Research":     st.column_config.TextColumn("Research 🔬",  width="medium"),
                "Troops":       st.column_config.TextColumn("Troops ⚔️",   width="medium"),
                "FCs":          st.column_config.TextColumn("FCs 💎",       width="small"),
                "FC Shards":    st.column_config.TextColumn("Shards 🔷",    width="small"),
                "Time UTC":     st.column_config.TextColumn("Time (UTC) 🕐", width="medium"),
                "Days":         st.column_config.TextColumn("Days 📅",      width="medium"),
            }
        )

        if export_csv:
            csv_data = df_sorted.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="Click to download",
                data=csv_data,
                file_name="fc_data.csv",
                mime="text/csv",
            )

        with st.expander("🔍 Raw parsed records (JSON)"):
            st.json(records)

    else:
        st.warning("No valid records found. Make sure each block starts with 'User ID: <number>'.")
else:
    st.info("Paste your player data in the text box above.")