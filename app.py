"""
Fleet Command Telemetry - V21 (Patched Full Build)
Architecture:
- Preserves original premium Streamlit UI
- Keeps doc -> docx conversion flow
- Keeps broad extraction style but adds section-aware fallbacks
- Adds Review Queue for ambiguous / invalid values
- Treats AUX as visible 9-component list
- Improves Other Equipment parsing for mixed date/hour/text rows
"""

import streamlit as st
st.set_page_config(
    page_title="Fleet Running Hours",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

import os
import re
import shutil
import tempfile
import subprocess
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from dateutil import parser as dt_parser

# ══════════════════════════════════════════════════════════════════
#  NATIVE UI THEME - PRESERVED / EXPANDED
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;600;700&family=Inter:wght@400;500;600&display=swap');

:root {
  --bg: #071019;
  --bg2: #0c1623;
  --bg3: #0f1d2d;
  --line: #1b2d44;
  --line2: #294867;
  --gold: #c99818;
  --t0: #ebf3ff;
  --t1: #a9bdd4;
  --t2: #71879f;
  --ok: #1f9d68;
  --warn: #d28d17;
  --bad: #c65348;
  --note: #4a78a8;
}

html, body, [class*="css"] {
  background: var(--bg)!important;
  color: var(--t1)!important;
  font-family: 'Inter', sans-serif!important;
}

.main, .block-container {
  background:
    radial-gradient(circle at top right, rgba(201,152,24,.055), transparent 25%),
    linear-gradient(180deg, rgba(255,255,255,.01), rgba(255,255,255,0)),
    var(--bg)!important;
}

[data-testid="stSidebar"], [data-testid="collapsedControl"] {
  display: none!important;
}

.hero-k {
  font-size: .66rem;
  letter-spacing: .24em;
  text-transform: uppercase;
  color: var(--gold);
  font-weight: 700;
}

.hero-h {
  font-family: 'Space Grotesk', sans-serif;
  font-size: 1.9rem;
  font-weight: 700;
  color: var(--t0);
  line-height: 1.1;
  margin-top: .2rem;
}

.hero-rule {
  height: 1px;
  margin: 1rem 0;
  background: linear-gradient(90deg, var(--gold), var(--line), transparent);
}

.metric-grid {
  display: grid;
  grid-template-columns: repeat(6, 1fr);
  gap: 1rem;
  margin: 1.25rem 0 1rem 0;
}

.metric {
  background: linear-gradient(180deg, var(--bg2), var(--bg3));
  border: 1px solid var(--line);
  border-radius: 12px;
  padding: 1rem;
  border-top: 2px solid var(--gold);
  box-shadow: 0 6px 18px rgba(0,0,0,.18);
}

.metric-v {
  font-family: 'Space Grotesk', sans-serif;
  font-size: 1.45rem;
  font-weight: 700;
  color: var(--t0);
  line-height: 1.1;
}

.metric-l {
  font-size: .6rem;
  text-transform: uppercase;
  letter-spacing: .15em;
  color: var(--t2);
  margin-top: 5px;
}

[data-testid="stFileUploadDropzone"] {
  background: rgba(201,152,24,.05)!important;
  border: 1.5px dashed rgba(201,152,24,.4)!important;
  border-radius: 12px!important;
}

div[data-baseweb="tab-list"] {
  gap: .45rem;
  border-bottom: 1px solid var(--line)!important;
}

button[data-baseweb="tab"] {
  background: linear-gradient(180deg, var(--bg2), #102131)!important;
  border: 1px solid var(--line)!important;
  border-radius: 12px 12px 0 0!important;
  color: var(--t1)!important;
  padding: .7rem 1rem!important;
}

button[aria-selected="true"][data-baseweb="tab"] {
  color: var(--t0)!important;
  border-color: var(--line2)!important;
}

.banner {
  border: 1px solid var(--line);
  background: linear-gradient(180deg, var(--bg2), #0d1826);
  border-radius: 12px;
  padding: .9rem 1rem;
  margin: .9rem 0 1rem 0;
}

.banner strong { color: var(--t0); }
.banner-ok   { border-left: 3px solid var(--ok); }
.banner-warn { border-left: 3px solid var(--warn); }
.banner-bad  { border-left: 3px solid var(--bad); }
.banner-note { border-left: 3px solid var(--note); }

.tbl-wrap {
  border: 1px solid var(--line);
  border-radius: 14px;
  overflow-x: auto;
  background: linear-gradient(180deg, #0a1420, #0d1826);
}

.tbl {
  width: 100%;
  border-collapse: collapse;
  min-width: 950px;
}

.tbl thead th {
  background: #102031;
  color: var(--t0);
  font-size: .74rem;
  text-transform: uppercase;
  letter-spacing: .08em;
  padding: .82rem .75rem;
  border-bottom: 1px solid var(--line);
  white-space: nowrap;
}

.tbl tbody td {
  padding: .78rem .75rem;
  border-bottom: 1px solid rgba(27,45,68,.8);
  color: var(--t1);
  vertical-align: middle;
  font-size: .92rem;
}

.tbl tbody tr:nth-child(even) td {
  background: rgba(255,255,255,.015);
}

.t-center { text-align: center; }
.t-right  { text-align: right; font-variant-numeric: tabular-nums; }

.status-chip {
  display: inline-flex;
  align-items: center;
  gap: .35rem;
  padding: .27rem .58rem;
  border-radius: 999px;
  font-size: .78rem;
  font-weight: 700;
  white-space: nowrap;
}

.status-ok {
  color: #dbfff1;
  background: rgba(31,157,104,.14);
  border: 1px solid rgba(31,157,104,.32);
}

.status-high {
  color: #ffedc7;
  background: rgba(210,141,23,.14);
  border: 1px solid rgba(210,141,23,.32);
}

.status-overdue {
  color: #ffe2dd;
  background: rgba(198,83,72,.14);
  border: 1px solid rgba(198,83,72,.32);
}

.status-nodata {
  color: #dce8f3;
  background: rgba(74,120,168,.14);
  border: 1px solid rgba(74,120,168,.32);
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
#  DOMAIN DEFINITIONS
# ══════════════════════════════════════════════════════════════════
ME_COMPONENTS = {
    "CYLINDER COVER", "PISTON ASSEMBLY", "STUFFING BOX", "PISTON CROWN",
    "CYLINDER LINER", "EXHAUST VALVE", "STARTING VALVE", "SAFETY VALVE",
    "FUEL VALVES", "FUEL PUMP", "PLUNGER AND BARREL(RENEWAL)",
    "PLUNGER AND BARREL", "FUEL PUMP SUCTION VALVE",
    "FUEL PUMP PUNCTURE VALVE", "CROSSHEAD BEARINGS",
    "BOTTOM END BEARINGS", "MAIN BEARINGS"
}

ME_COMPONENT_ORDER = [
    "CYLINDER COVER", "PISTON ASSEMBLY", "STUFFING BOX", "PISTON CROWN",
    "CYLINDER LINER", "EXHAUST VALVE", "STARTING VALVE", "SAFETY VALVE",
    "FUEL VALVES", "FUEL PUMP", "PLUNGER AND BARREL(RENEWAL)",
    "PLUNGER AND BARREL", "FUEL PUMP SUCTION VALVE",
    "FUEL PUMP PUNCTURE VALVE", "CROSSHEAD BEARINGS",
    "BOTTOM END BEARINGS", "MAIN BEARINGS"
]

AUX_COMPONENTS = {
    "CYLINDER HEAD", "PISTON", "CONNECTING ROD", "CYLINDER LINERS",
    "FUEL VALVES (1)", "FUEL PUMPS", "CRANK PIN BEARING",
    "MAIN BEARING", "ADJUST VALVE HEAD CLEARANCE"
}

AUX_COMPONENT_ORDER = [
    "CYLINDER HEAD", "PISTON", "CONNECTING ROD", "CYLINDER LINERS",
    "FUEL VALVES (1)", "FUEL PUMPS", "CRANK PIN BEARING",
    "MAIN BEARING", "ADJUST VALVE HEAD CLEARANCE"
]

DG_COMPONENTS = {
    "TURBOCHARGER (2)", "TURBOCHARGER (3)", "COOLING WATER PUMP",
    "COOL WATER THERMOSTAT VALVE", "L.O. THERMOSTAT VALVE",
    "THRUST BEARING", "AIR COOLER", "L.O. COOLER CLEAN",
    "F.W. COOLER CLEAN", "L.O. RENEWAL", "ALTERNATOR CLEANING"
}

DG_COMPONENT_ORDER = [
    "TURBOCHARGER (2)", "TURBOCHARGER (3)", "COOLING WATER PUMP",
    "COOL WATER THERMOSTAT VALVE", "L.O. THERMOSTAT VALVE",
    "THRUST BEARING", "AIR COOLER", "L.O. COOLER CLEAN",
    "F.W. COOLER CLEAN", "L.O. RENEWAL", "ALTERNATOR CLEANING"
]

OTHER_SIMPLE_COMPONENTS = {
    "GENERAL O/H", "BALANCING OF ROTOR SHAFT", "AIR COOLER CLEANING",
    "M/E L.O.", "JACKET FW", "JACKET FW NO.1", "PISTON L.O.",
    "ATMOSPHERIC CONDENSER", "AIR COND. COMPRESSOR NO.1",
    "AIR COND. COMPRESSOR NO.2", "AIR. COND. COOLER CLEANING",
    "REFRIGERATION COMPRESSOR NO.1", "REFRIGERATION COMPRESSOR NO.2",
    "FURNACE INSPECTION", "BURNER ATOMIZER", "FORCED DRAFT FAN",
    "FEED PUMPS NO.1", "FEED PUMPS NO.2", "WASHING THE TUBES",
    "O/H CIRC. PUMP NO.1", "O/H CIRC. PUMP NO.2",
    "STARTING MAIN AIR COMPRESSOR NO.1", "STARTING MAIN AIR COMPRESSOR NO.2",
    "SERVICE AIR COMPRESSOR", "EMERGENCY AIR COMPRESSOR NO."
}

OTHER_SIMPLE_ORDER = [
    "GENERAL O/H", "BALANCING OF ROTOR SHAFT", "AIR COOLER CLEANING",
    "M/E L.O.", "JACKET FW", "JACKET FW NO.1", "PISTON L.O.",
    "ATMOSPHERIC CONDENSER", "AIR COND. COMPRESSOR NO.1",
    "AIR COND. COMPRESSOR NO.2", "AIR. COND. COOLER CLEANING",
    "REFRIGERATION COMPRESSOR NO.1", "REFRIGERATION COMPRESSOR NO.2",
    "FURNACE INSPECTION", "BURNER ATOMIZER", "FORCED DRAFT FAN",
    "FEED PUMPS NO.1", "FEED PUMPS NO.2", "WASHING THE TUBES",
    "O/H CIRC. PUMP NO.1", "O/H CIRC. PUMP NO.2",
    "STARTING MAIN AIR COMPRESSOR NO.1", "STARTING MAIN AIR COMPRESSOR NO.2",
    "SERVICE AIR COMPRESSOR", "EMERGENCY AIR COMPRESSOR NO."
]

ALIASES = {
    "PERIODICTLY": "PERIODICITY",
    "EXAUST VALVE": "EXHAUST VALVE",
    "FUEL VALVES(1)": "FUEL VALVES (1)",
    "TURBOCHARGER(2)": "TURBOCHARGER (2)",
    "TURBOCHARGER(3)": "TURBOCHARGER (3)",
    "PLUNGER AND BARREL (RENEWAL)": "PLUNGER AND BARREL(RENEWAL)",
    "JACKET FW NO.1": "JACKET FW"
}

TEXTUAL_VALUES = {"N/A", "NO RECORD", "NOT WORKING", "CENTRAL", "COOLER"}

# ══════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════
def fl(txt: Any) -> str:
    if txt is None:
        return ""
    raw = str(txt).replace("\x07", "").replace("\xa0", " ").replace("\t", " ")
    lines = [line.strip() for line in raw.split("\n") if line.strip()]
    return " ".join(lines) if lines else ""

def normalize_comp(txt: Any) -> str:
    s = fl(txt).upper()
    s = s.replace("**", " ")
    s = s.replace("[", "").replace("]", "")
    s = re.sub(r"\s+", " ", s).strip(" :-#*")
    s = s.replace("SEPT", "SEP")
    return ALIASES.get(s, s)

def parse_num(txt: Any) -> Optional[float]:
    s = normalize_comp(txt)
    if not s:
        return None
    if s in TEXTUAL_VALUES or "OBSERVATION" in s:
        return None
    m = re.search(r"\d[\d,\.]*", s)
    if not m:
        return None
    token = m.group()
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", token) or re.fullmatch(r"\d{1,3}(,\d{3})+", token):
        token = token.replace(".", "").replace(",", "")
    elif token.count(",") == 1 and token.count(".") == 0 and len(token.split(",")[-1]) != 3:
        token = token.replace(",", ".")
    else:
        token = token.replace(",", "")
    try:
        return float(token)
    except:
        return None

def parse_date(txt: Any) -> Tuple[Optional[str], Optional[str], bool]:
    raw = fl(txt).replace("[", "").replace("]", "").strip()
    if not raw or raw in {"-", ""}:
        return None, None, False

    norm = normalize_comp(raw)
    if norm in {"1", "2"}:
        return None, None, False

    if norm in TEXTUAL_VALUES:
        return None, norm, False

    try:
        dt = dt_parser.parse(raw, dayfirst=True, fuzzy=False)
        return dt.date().isoformat(), None, False
    except:
        return None, raw, True

def format_hours(val: Optional[float]) -> str:
    if val is None:
        return "—"
    if float(val).is_integer():
        return f"{int(val):,}"
    return f"{val:,.1f}"

def format_percent(val: Optional[float]) -> str:
    if val is None:
        return "—"
    return f"{val:.1f}%"

def get_status(hrs: Optional[float], period: Optional[float]) -> str:
    if hrs is None or period is None or period <= 0:
        return "NO DATA"
    ratio = hrs / period
    if ratio >= 1.0:
        return "OVERDUE"
    if ratio >= 0.8:
        return "HIGH PRIORITY"
    return "OK"

def add_review(review: List[Dict[str, str]], section: str, item: str, issue: str, raw_value: str):
    review.append({
        "Section": section,
        "Item": item,
        "Issue": issue,
        "Raw Value": raw_value
    })

def sort_key(name: str, order: List[str]) -> int:
    try:
        return order.index(name)
    except ValueError:
        return 9999

def render_status_chip(status: str) -> str:
    s = (status or "").upper()
    if s == "OVERDUE":
        cls = "status-overdue"
    elif s == "HIGH PRIORITY":
        cls = "status-high"
    elif s == "OK":
        cls = "status-ok"
    else:
        cls = "status-nodata"
    return f'<span class="status-chip {cls}">{status}</span>'

def render_html_table(df: pd.DataFrame, cols: List[str], numeric_cols=None, center_cols=None):
    numeric_cols = numeric_cols or []
    center_cols = center_cols or []

    html = ['<div class="tbl-wrap"><table class="tbl"><thead><tr>']
    for c in cols:
        html.append(f"<th>{c}</th>")
    html.append("</tr></thead><tbody>")

    for _, row in df[cols].iterrows():
        html.append("<tr>")
        for c in cols:
            cls = ""
            if c in numeric_cols:
                cls = ' class="t-right"'
            elif c in center_cols:
                cls = ' class="t-center"'
            val = "—" if pd.isna(row[c]) else str(row[c])
            html.append(f"<td{cls}>{val}</td>")
        html.append("</tr>")

    html.append("</tbody></table></div>")
    st.markdown("".join(html), unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
#  DOC CONVERSION
# ══════════════════════════════════════════════════════════════════
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found.")

    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as t:
        t.write(raw)
        src = t.name

    outdir = tempfile.mkdtemp(prefix="lo_")
    out = os.path.join(outdir, Path(src).stem + ".docx")
    profile = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"

    try:
        proc = subprocess.run(
            [
                soffice,
                "--headless",
                "--norestore",
                "--nofirststartwizard",
                f"-env:UserInstallation={profile}",
                "--convert-to", "docx",
                src,
                "--outdir", outdir
            ],
            capture_output=True,
            timeout=120,
            text=True
        )
        if proc.returncode != 0 or not os.path.exists(out):
            raise RuntimeError((proc.stderr or proc.stdout or "LibreOffice conversion failed").strip())
        with open(out, "rb") as f:
            return f.read()
    finally:
        for p in [src, out]:
            try:
                os.unlink(p)
            except:
                pass
        shutil.rmtree(outdir, ignore_errors=True)

# ══════════════════════════════════════════════════════════════════
#  DOCX EXTRACTION
# ══════════════════════════════════════════════════════════════════
def read_docx_structure(docx_bytes: bytes):
    from docx import Document
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as t:
        t.write(docx_bytes)
        tp = t.name

    try:
        doc = Document(tp)
        paragraphs = [fl(p.text) for p in doc.paragraphs if fl(p.text)]
        tables = []
        for table in doc.tables:
            rows = []
            for row in table.rows:
                vals = [fl(cell.text) for cell in row.cells]
                if any(v.strip() for v in vals):
                    rows.append(vals)
            if rows:
                tables.append(rows)
        return paragraphs, tables
    finally:
        try:
            os.unlink(tp)
        except:
            pass

def flatten_for_legacy_scan(paragraphs, tables):
    cells = []
    for p in paragraphs:
        if fl(p):
            cells.append(fl(p))
    for table in tables:
        for row in table:
            for cell in row:
                cells.append(fl(cell))
    return cells

def all_rows(paragraphs, tables):
    rows = []
    for p in paragraphs:
        rows.append([p])
    for tbl in tables:
        rows.extend(tbl)
    return rows

def row_text(row: List[str]) -> str:
    return " | ".join([fl(x) for x in row if fl(x)])

def norm_row(row: List[str]) -> List[str]:
    return [normalize_comp(x) for x in row]

def find_first_row(rows, predicate):
    for i, row in enumerate(rows):
        if predicate(row):
            return i
    return None

# ══════════════════════════════════════════════════════════════════
#  HEADER EXTRACTION
# ══════════════════════════════════════════════════════════════════
def extract_header(rows, cells, review):
    text = " ".join(row_text(r) for r in rows)
    out = {
        "vessel": "UNKNOWN",
        "report_date": "—",
        "me_total_hours": None,
        "me_this_month": None
    }

    m = re.search(r"VESSEL['’]S NAME\s*:\s*(?:MV\s+)?(.+?)\s+DATE\s*:?\s*([A-Z0-9/ .-]+)", text, re.I)
    if m:
        out["vessel"] = re.sub(r"(?i)^MV\s+", "", fl(m.group(1))).strip()
        iso, raw, invalid = parse_date(m.group(2))
        out["report_date"] = iso or raw or "—"
        if invalid:
            add_review(review, "Header", "Report Date", "Invalid date format", raw or "")
    else:
        for c in cells:
            if "VESSEL" in c.upper() and "DATE" in c.upper():
                m2 = re.search(r"Name\s*:\s*(?:MV\s+)?(.+?)\s+Date\s*:?\s*(.+)$", c, re.I)
                if m2:
                    out["vessel"] = re.sub(r"(?i)^MV\s+", "", fl(m2.group(1))).strip()
                    iso, raw, invalid = parse_date(m2.group(2))
                    out["report_date"] = iso or raw or "—"
                    if invalid:
                        add_review(review, "Header", "Report Date", "Invalid date format", raw or "")
                    break

    m = re.search(r"TOTAL RUNNING HOURS[\s:ǀ|]*([\d,\.]+)", text, re.I)
    if m:
        out["me_total_hours"] = parse_num(m.group(1))

    m = re.search(r"THIS MONTH[\s:]*([\d,\.]+)", text, re.I)
    if m:
        out["me_this_month"] = parse_num(m.group(1))

    return out

# ══════════════════════════════════════════════════════════════════
#  MAIN ENGINE - TABLE AWARE
# ══════════════════════════════════════════════════════════════════
def find_me_start(rows):
    return find_first_row(rows, lambda r: "MAIN ENGINE" in normalize_comp(row_text(r)) or "CYL. NO." in normalize_comp(row_text(r)))

def find_other_start(rows):
    return find_first_row(rows, lambda r: "TURBOCHARGER" in normalize_comp(row_text(r)) and "COOLERS" in normalize_comp(row_text(r)))

def find_aux_start(rows):
    return find_first_row(rows, lambda r: "AUX. ENGINE MAKER / TYPE" in normalize_comp(row_text(r)))

def find_dg_start(rows):
    return find_first_row(rows, lambda r: "D/G NO1" in normalize_comp(row_text(r)) or "D/G NO.1" in normalize_comp(row_text(r)))

def extract_me_table(rows, me_start, other_start, review):
    if me_start is None:
        return []

    end = other_start if other_start is not None and other_start > me_start else len(rows)
    zone = rows[me_start:end]
    out = []

    for i in range(len(zone) - 1):
        r1, r2 = zone[i], zone[i + 1]
        if len(r1) < 4 or len(r2) < 4:
            continue

        comp = normalize_comp(r1[0])
        if comp not in ME_COMPONENTS:
            continue

        marker1 = normalize_comp(r1[2]) if len(r1) > 2 else ""
        marker2 = normalize_comp(r2[2]) if len(r2) > 2 else ""
        if marker1 != "1" or marker2 != "2":
            continue

        raw_period = r1[1] if len(r1) > 1 else ""
        period = parse_num(raw_period)
        obs_based = "OBSERVATION" in normalize_comp(raw_period)

        if obs_based:
            add_review(review, "Main Engine", comp, "Observation-based periodicity", fl(raw_period))

        max_cols = min(len(r1), len(r2))
        cyl_count = min(max_cols - 3, 7)

        for c_idx in range(cyl_count):
            raw_date = r1[3 + c_idx] if 3 + c_idx < len(r1) else ""
            raw_hrs = r2[3 + c_idx] if 3 + c_idx < len(r2) else ""

            iso, raw_text, invalid = parse_date(raw_date)
            hrs = parse_num(raw_hrs)

            if invalid:
                add_review(review, "Main Engine", f"{comp} / Cyl {c_idx+1}", "Invalid date", raw_text or "")

            if raw_text in TEXTUAL_VALUES:
                add_review(review, "Main Engine", f"{comp} / Cyl {c_idx+1}", "Textual date value kept", raw_text)

            if iso or raw_text or hrs is not None:
                used = (hrs / period * 100) if (hrs is not None and period and period > 0) else None
                out.append({
                    "Status": get_status(hrs, period),
                    "Component": comp,
                    "Engine": "ME",
                    "Unit": f"Cyl {c_idx+1}",
                    "Periodicity": "OBSERVATION" if obs_based else (
                        int(period) if period and float(period).is_integer() else (period if period is not None else "—")
                    ),
                    "Last O/H": iso or raw_text or "—",
                    "Hrs Since": format_hours(hrs),
                    "Used %": format_percent(used)
                })

    out.sort(key=lambda x: (sort_key(x["Component"], ME_COMPONENT_ORDER), int(re.search(r"\d+", x["Unit"]).group())))
    return out

# ══════════════════════════════════════════════════════════════════
#  AUX ENGINE - 9 COMPONENT BLOCK
# ══════════════════════════════════════════════════════════════════
def extract_aux_meta(zone):
    text = " ".join(row_text(r) for r in zone)
    out = {"aux_total_hours": None, "aux_this_month": None}

    m = re.search(r"TOTAL HOURS\s*:?\s*([\d,\.]+)", text, re.I)
    if m:
        out["aux_total_hours"] = parse_num(m.group(1))

    m = re.search(r"HOURS THIS MONTH\s*([\d,\.]+)", text, re.I)
    if m:
        out["aux_this_month"] = parse_num(m.group(1))

    return out

def find_aux_header(zone):
    for i, row in enumerate(zone):
        joined = " | ".join(norm_row(row))
        if "DESCRIPTION" in joined and "1" in joined and "2" in joined and ("PERIODICITY" in joined or "PERIODICTLY" in joined):
            return i
    return None

def extract_aux_table(rows, aux_start, dg_start, review):
    if aux_start is None:
        return [], {"aux_total_hours": None, "aux_this_month": None}, AUX_COMPONENT_ORDER.copy()

    end = dg_start if dg_start is not None and dg_start > aux_start else len(rows)
    zone = rows[aux_start:end]
    meta = extract_aux_meta(zone)

    header_idx = find_aux_header(zone)
    if header_idx is None:
        add_review(review, "Aux Engine", "Header", "AUX detail header not found", "DESCRIPTION | PERIODICITY | 1 | 2")
        return [], meta, AUX_COMPONENT_ORDER.copy()

    body = zone[header_idx + 1:]
    out = []
    i = 0

    while i < len(body) - 1:
        r1, r2 = body[i], body[i + 1]
        comp = normalize_comp(r1[0] if len(r1) > 0 else "")
        marker1 = normalize_comp(r1[2] if len(r1) > 2 else "")
        marker2 = normalize_comp(r2[2] if len(r2) > 2 else "")

        if comp in AUX_COMPONENTS and marker1 == "1" and marker2 == "2":
            period = parse_num(r1[1] if len(r1) > 1 else "")
            raw_date = r1[3] if len(r1) > 3 else ""
            raw_hrs = r2[3] if len(r2) > 3 else ""

            iso, raw_text, invalid = parse_date(raw_date)
            hrs = parse_num(raw_hrs)

            if invalid:
                add_review(review, "Aux Engine", comp, "Invalid date", raw_text or "")

            if raw_text in TEXTUAL_VALUES:
                add_review(review, "Aux Engine", comp, "Textual date value kept", raw_text)

            if raw_hrs and parse_num(raw_hrs) is None and normalize_comp(raw_hrs) not in {"", "2"}:
                add_review(review, "Aux Engine", comp, "Non-numeric running hours", fl(raw_hrs))

            used = (hrs / period * 100) if (hrs is not None and period and period > 0) else None
            out.append({
                "Status": get_status(hrs, period),
                "Component": comp,
                "Engine": "AUX",
                "Unit": "Engine",
                "Periodicity": int(period) if period and float(period).is_integer() else (period if period is not None else "—"),
                "Last O/H": iso or raw_text or "—",
                "Hrs Since": format_hours(hrs),
                "Used %": format_percent(used)
            })
            i += 2
            continue

        i += 1

    found = {x["Component"] for x in out}
    missing = [x for x in AUX_COMPONENT_ORDER if x not in found]
    for m in missing:
        add_review(review, "Aux Engine", m, "Expected AUX component not found", "")

    out.sort(key=lambda x: sort_key(x["Component"], AUX_COMPONENT_ORDER))
    return out, meta, missing

# ══════════════════════════════════════════════════════════════════
#  OTHER EQUIPMENT
# ══════════════════════════════════════════════════════════════════
def extract_other_simple_zone(zone, review):
    out = []

    for row in zone:
        cells = [fl(c) for c in row]
        for idx, cell in enumerate(cells):
            comp = normalize_comp(cell)
            if comp not in OTHER_SIMPLE_COMPONENTS:
                continue

            raw_period = cells[idx + 1] if idx + 1 < len(cells) else ""
            raw_date = cells[idx + 2] if idx + 2 < len(cells) else ""
            raw_hrs = cells[idx + 3] if idx + 3 < len(cells) else ""

            period = parse_num(raw_period)

            date_is_num_like = bool(re.fullmatch(r"\d+(\.\d+)?", normalize_comp(raw_date)))
            if raw_date and not date_is_num_like:
                iso, raw_text, invalid = parse_date(raw_date)
            else:
                iso, raw_text, invalid = None, None, False

            hrs = parse_num(raw_hrs)

            if invalid:
                add_review(review, "Other Equipment", comp, "Invalid date", raw_text or "")

            if raw_text in TEXTUAL_VALUES:
                add_review(review, "Other Equipment", comp, "Textual date value kept", raw_text)

            if raw_hrs and parse_num(raw_hrs) is None and normalize_comp(raw_hrs) not in {"", "1", "2"}:
                add_review(review, "Other Equipment", comp, "Non-numeric running hours", fl(raw_hrs))

            if period is None and raw_period and normalize_comp(raw_period) not in TEXTUAL_VALUES:
                add_review(review, "Other Equipment", comp, "Unclear periodicity", fl(raw_period))

            used = (hrs / period * 100) if (hrs is not None and period and period > 0) else None
            out.append({
                "Status": get_status(hrs, period),
                "Description": comp,
                "Unit": "—",
                "Periodicity": int(period) if period and float(period).is_integer() else (period if period is not None else "—"),
                "Last Date": iso or raw_text or "—",
                "Run Hrs": format_hours(hrs),
                "Used %": format_percent(used)
            })

    return out

def extract_dg_zone(zone, review):
    out = []

    for i in range(len(zone) - 1):
        r1, r2 = zone[i], zone[i + 1]
        comp = normalize_comp(r1[0] if len(r1) > 0 else "")
        if comp not in DG_COMPONENTS:
            continue

        marker1 = normalize_comp(r1[2] if len(r1) > 2 else "")
        marker2 = normalize_comp(r2[2] if len(r2) > 2 else "")
        if marker1 != "1" or marker2 != "2":
            continue

        period = parse_num(r1[1] if len(r1) > 1 else "")

        for dg_idx in range(3):
            c = 3 + dg_idx
            raw_date = r1[c] if c < len(r1) else ""
            raw_hrs = r2[c] if c < len(r2) else ""

            iso, raw_text, invalid = parse_date(raw_date)
            hrs = parse_num(raw_hrs)

            if invalid:
                add_review(review, "Other Equipment", f"{comp} / DG {dg_idx+1}", "Invalid date", raw_text or "")

            if raw_text in TEXTUAL_VALUES:
                add_review(review, "Other Equipment", f"{comp} / DG {dg_idx+1}", "Textual date value kept", raw_text)

            if raw_hrs and parse_num(raw_hrs) is None and normalize_comp(raw_hrs) not in {"", "1", "2"}:
                add_review(review, "Other Equipment", f"{comp} / DG {dg_idx+1}", "Non-numeric running hours", fl(raw_hrs))

            if iso or raw_text or hrs is not None:
                used = (hrs / period * 100) if (hrs is not None and period and period > 0) else None
                out.append({
                    "Status": get_status(hrs, period),
                    "Description": comp,
                    "Unit": f"DG {dg_idx+1}",
                    "Periodicity": int(period) if period and float(period).is_integer() else (period if period is not None else "—"),
                    "Last Date": iso or raw_text or "—",
                    "Run Hrs": format_hours(hrs),
                    "Used %": format_percent(used)
                })

    return out

def dedupe_records(records, keys):
    seen = set()
    out = []
    for r in records:
        k = tuple(r.get(x) for x in keys)
        if k not in seen:
            seen.add(k)
            out.append(r)
    return out

def extract_other_table(rows, other_start, aux_start, dg_start, review):
    parts = []

    if other_start is not None:
        end_a = aux_start if aux_start is not None and aux_start > other_start else len(rows)
        zone_a = rows[other_start:end_a]
        parts.extend(extract_other_simple_zone(zone_a, review))

    if dg_start is not None:
        zone_b = rows[dg_start:]
        parts.extend(extract_dg_zone(zone_b, review))

    parts = dedupe_records(parts, ["Description", "Unit", "Periodicity", "Last Date", "Run Hrs"])
    parts.sort(key=lambda x: (
        sort_key(x["Description"], OTHER_SIMPLE_ORDER + DG_COMPONENT_ORDER),
        999 if x["Unit"] == "—" else int(re.search(r"\d+", x["Unit"]).group())
    ))
    return parts

# ══════════════════════════════════════════════════════════════════
#  LEGACY FALLBACK SCAN
# ══════════════════════════════════════════════════════════════════
def extract_legacy_fallback(cells, review):
    me_rows = []
    vessel = "UNKNOWN"
    report_date = "—"
    me_total = None
    me_month = None

    doc_text = " ".join([c for c in cells if c]).upper()

    if m := re.search(r"TOTAL RUNNING HOURS[\s:ǀ|]*([\d,\.]+)", doc_text):
        me_total = parse_num(m.group(1))
    if m := re.search(r"THIS MONTH[\s:]*([\d,\.]+)", doc_text):
        me_month = parse_num(m.group(1))

    for c in cells:
        if 'VESSEL' in c.upper() and 'DATE' in c.upper():
            if m := re.search(r"Name\s*:\s*(?:MV\s+)?(.+?)\s+Date\s*:\s*(.+)$", c, re.I):
                vessel = re.sub(r'(?i)^MV\s+', '', fl(m.group(1)))
                iso, raw, invalid = parse_date(m.group(2))
                report_date = iso or raw or "—"
                if invalid:
                    add_review(review, "Header", "Report Date", "Invalid date format", raw or "")
                break

    return vessel, report_date, me_total, me_month, me_rows

# ══════════════════════════════════════════════════════════════════
#  UI
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero-k">Running Hours Management System</div>
<div class="hero-h">TEC-004 Extraction Matrix</div>
<div class="hero-rule"></div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("Upload TEC-004 Report (.doc)", type=['doc'])

if uploaded:
    with st.spinner("Executing Zero-Trust Extraction..."):
        try:
            docx_data = convert_doc_to_docx(uploaded.read())
            paragraphs, tables = read_docx_structure(docx_data)
            cells = flatten_for_legacy_scan(paragraphs, tables)
            rows = all_rows(paragraphs, tables)

            review = []

            header = extract_header(rows, cells, review)

            me_start = find_me_start(rows)
            other_start = find_other_start(rows)
            aux_start = find_aux_start(rows)
            dg_start = find_dg_start(rows)

            me_data = extract_me_table(rows, me_start, other_start, review)
            aux_data, aux_meta, aux_missing = extract_aux_table(rows, aux_start, dg_start, review)
            oe_data = extract_other_table(rows, other_start, aux_start, dg_start, review)

            if not me_data and not aux_data and not oe_data:
                legacy_vessel, legacy_date, legacy_me_total, legacy_me_month, _ = extract_legacy_fallback(cells, review)
                if header["vessel"] == "UNKNOWN":
                    header["vessel"] = legacy_vessel
                if header["report_date"] == "—":
                    header["report_date"] = legacy_date
                if header["me_total_hours"] is None:
                    header["me_total_hours"] = legacy_me_total
                if header["me_this_month"] is None:
                    header["me_this_month"] = legacy_me_month

            n_od = sum(1 for c in (me_data + aux_data + oe_data) if c['Status'] == 'OVERDUE')
            n_hp = sum(1 for c in (me_data + aux_data + oe_data) if c['Status'] == 'HIGH PRIORITY')
            n_rv = len(review)

            st.markdown(f"""
            <div class="metric-grid">
              <div class="metric"><div class="metric-v">{header.get('vessel') or 'UNKNOWN'}</div><div class="metric-l">Vessel</div></div>
              <div class="metric"><div class="metric-v">{header.get('report_date') or '—'}</div><div class="metric-l">Report Date</div></div>
              <div class="metric"><div class="metric-v">{format_hours(header.get('me_total_hours'))}</div><div class="metric-l">ME Total Hrs</div></div>
              <div class="metric"><div class="metric-v">{format_hours(header.get('me_this_month'))}</div><div class="metric-l">ME This Month</div></div>
              <div class="metric"><div class="metric-v">{n_od}</div><div class="metric-l">Overdue</div></div>
              <div class="metric"><div class="metric-v">{n_rv}</div><div class="metric-l">Review Items</div></div>
            </div>
            """, unsafe_allow_html=True)

            if aux_missing:
                st.markdown(
                    f'<div class="banner banner-warn"><strong>Aux parsed as the visible 9-component block.</strong> Missing expected items: {", ".join(aux_missing)}</div>',
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    '<div class="banner banner-ok"><strong>Aux extracted successfully.</strong> The app treats AUX as the single 9-component list shown in the report.</div>',
                    unsafe_allow_html=True
                )

            if review:
                st.markdown(
                    f'<div class="banner banner-note"><strong>Validation active.</strong> {len(review)} item(s) were flagged for operator review instead of being silently discarded.</div>',
                    unsafe_allow_html=True
                )

            tab1, tab2, tab3, tab4 = st.tabs([
                f"⚙ Main Engine ({len(me_data)})",
                f"🔩 Aux Engines ({len(aux_data)})",
                f"🛠 Other Equipment ({len(oe_data)})",
                f"🧾 Review Queue ({len(review)})"
            ])

            with tab1:
                if not me_data:
                    st.info("No Main Engine records found.")
                else:
                    df_me = pd.DataFrame(me_data)
                    df_me["Status"] = df_me["Status"].apply(render_status_chip)
                    render_html_table(
                        df_me,
                        ["Status", "Component", "Engine", "Unit", "Periodicity", "Last O/H", "Hrs Since", "Used %"],
                        numeric_cols=["Hrs Since", "Used %"],
                        center_cols=["Engine", "Unit"]
                    )

            with tab2:
                if not aux_data:
                    st.info("No Auxiliary Engine records found.")
                else:
                    df_aux = pd.DataFrame(aux_data)
                    df_aux["Status"] = df_aux["Status"].apply(render_status_chip)
                    render_html_table(
                        df_aux,
                        ["Status", "Component", "Engine", "Unit", "Periodicity", "Last O/H", "Hrs Since", "Used %"],
                        numeric_cols=["Hrs Since", "Used %"],
                        center_cols=["Engine", "Unit"]
                    )

                    st.markdown(
                        f'<div class="banner banner-note"><strong>Aux Meta:</strong> Total Hours = {format_hours(aux_meta.get("aux_total_hours"))}, Hours This Month = {format_hours(aux_meta.get("aux_this_month"))}</div>',
                        unsafe_allow_html=True
                    )

            with tab3:
                if not oe_data:
                    st.info("No Other Equipment records found.")
                else:
                    df_oe = pd.DataFrame(oe_data)
                    df_oe["Status"] = df_oe["Status"].apply(render_status_chip)
                    render_html_table(
                        df_oe,
                        ["Status", "Description", "Unit", "Periodicity", "Last Date", "Run Hrs", "Used %"],
                        numeric_cols=["Run Hrs", "Used %"],
                        center_cols=["Unit"]
                    )

            with tab4:
                if not review:
                    st.success("No flagged issues found.")
                else:
                    df_review = pd.DataFrame(review).drop_duplicates()
                    render_html_table(
                        df_review,
                        ["Section", "Item", "Issue", "Raw Value"]
                    )

        except Exception as e:
            st.error(f"Execution Failed: {e}")
