import os
import re
import shutil
import tempfile
import subprocess
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from docx import Document
from dateutil import parser as dt_parser

st.set_page_config(
    page_title="Fleet Running Hours",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=Inter:wght@400;500;600&display=swap');

:root{
  --bg:#071019;
  --bg2:#0d1725;
  --bg3:#122132;
  --line:#20344a;
  --gold:#c99818;
  --text:#eef5ff;
  --muted:#a8bdd2;
  --soft:#7d92a8;
  --ok:#1f9d68;
  --warn:#d5911e;
  --bad:#cf5548;
  --note:#567da4;
}

html, body, [class*="css"]{
  background:var(--bg)!important;
  color:var(--muted)!important;
  font-family:'Inter', sans-serif!important;
}

.main, .block-container{
  background:radial-gradient(circle at top right, rgba(201,152,24,.05), transparent 24%), var(--bg)!important;
}

[data-testid="stSidebar"], [data-testid="collapsedControl"]{
  display:none!important;
}

.block-container{
  padding-top:1rem!important;
}

.hero-k{
  font-size:.68rem;
  letter-spacing:.24em;
  text-transform:uppercase;
  color:var(--gold);
  font-weight:700;
}

.hero-h{
  font-family:'Space Grotesk', sans-serif;
  font-size:1.95rem;
  font-weight:700;
  color:var(--text);
  line-height:1.05;
  margin-top:.25rem;
}

.hero-rule{
  height:1px;
  margin:1rem 0 1.2rem 0;
  background:linear-gradient(90deg,var(--gold),var(--line),transparent);
}

.metric-grid{
  display:grid;
  grid-template-columns:repeat(6,1fr);
  gap:1rem;
  margin:1rem 0 1.2rem 0;
}

.metric{
  background:linear-gradient(180deg,var(--bg2),var(--bg3));
  border:1px solid var(--line);
  border-top:2px solid var(--gold);
  border-radius:14px;
  padding:1rem;
}

.metric-v{
  font-family:'Space Grotesk', sans-serif;
  font-size:1.32rem;
  font-weight:700;
  color:var(--text);
  line-height:1.05;
}

.metric-l{
  font-size:.64rem;
  text-transform:uppercase;
  letter-spacing:.15em;
  color:var(--soft);
  margin-top:6px;
}

.stTabs [data-baseweb="tab-list"]{
  background:transparent;
  gap:.45rem;
  border-bottom:1px solid var(--line);
  padding:0 0 .65rem 0;
}

.stTabs [data-baseweb="tab"]{
  background:linear-gradient(180deg,var(--bg2),#112132);
  border:1px solid var(--line);
  border-radius:12px;
  color:var(--muted);
  padding:.6rem 1rem;
  font-weight:600;
}

.stTabs [aria-selected="true"]{
  background:linear-gradient(180deg,#17304a,#102132)!important;
  border-color:#3e607d!important;
  color:var(--text)!important;
}

[data-testid="stFileUploadDropzone"]{
  background:rgba(201,152,24,.04)!important;
  border:1.25px dashed rgba(201,152,24,.34)!important;
  border-radius:14px!important;
}

.banner{
  border:1px solid var(--line);
  background:var(--bg2);
  border-radius:12px;
  padding:.95rem 1rem;
  margin:.8rem 0 1rem 0;
}

.banner strong{ color:var(--text); }

.banner-ok{ border-left:3px solid var(--ok); }
.banner-warn{ border-left:3px solid var(--warn); }
.banner-bad{ border-left:3px solid var(--bad); }
.banner-note{ border-left:3px solid var(--note); }

.table-wrap{
  overflow-x:auto;
  border:1px solid var(--line);
  border-radius:14px;
  background:var(--bg2);
}

.report-table{
  width:100%;
  border-collapse:separate;
  border-spacing:0;
  min-width:940px;
}

.report-table thead th{
  background:#102030;
  color:var(--text);
  font-size:.76rem;
  text-transform:uppercase;
  letter-spacing:.08em;
  padding:.85rem .75rem;
  border-bottom:1px solid var(--line);
  white-space:nowrap;
}

.report-table tbody td{
  padding:.82rem .75rem;
  border-bottom:1px solid rgba(32,52,74,.75);
  color:var(--muted);
  font-size:.93rem;
  vertical-align:middle;
}

.report-table tbody tr:nth-child(even){
  background:rgba(255,255,255,.016);
}

.report-table td.num{
  text-align:right;
  font-variant-numeric:tabular-nums;
}

.report-table td.center{
  text-align:center;
}

.status-chip{
  display:inline-flex;
  align-items:center;
  gap:.38rem;
  padding:.28rem .56rem;
  border-radius:999px;
  font-size:.78rem;
  font-weight:700;
  white-space:nowrap;
}

.status-overdue{
  color:#ffe0db;
  background:rgba(207,85,72,.16);
  border:1px solid rgba(207,85,72,.35);
}

.status-high{
  color:#ffe8bf;
  background:rgba(213,145,30,.16);
  border:1px solid rgba(213,145,30,.35);
}

.status-ok{
  color:#d9fff0;
  background:rgba(31,157,104,.16);
  border:1px solid rgba(31,157,104,.35);
}

.status-nodata{
  color:#dde7f2;
  background:rgba(86,125,164,.16);
  border:1px solid rgba(86,125,164,.35);
}
</style>
""", unsafe_allow_html=True)

ME_COMPONENTS = [
    "CYLINDER COVER", "PISTON ASSEMBLY", "STUFFING BOX", "PISTON CROWN",
    "CYLINDER LINER", "EXHAUST VALVE", "STARTING VALVE", "SAFETY VALVE",
    "FUEL VALVES", "FUEL PUMP", "PLUNGER AND BARREL(RENEWAL)",
    "PLUNGER AND BARREL", "FUEL PUMP SUCTION VALVE",
    "FUEL PUMP PUNCTURE VALVE", "CROSSHEAD BEARINGS",
    "BOTTOM END BEARINGS", "MAIN BEARINGS"
]

AUX_COMPONENTS = [
    "CYLINDER HEAD", "PISTON", "CONNECTING ROD", "CYLINDER LINERS",
    "FUEL VALVES (1)", "FUEL PUMPS", "CRANK PIN BEARING",
    "MAIN BEARING", "ADJUST VALVE HEAD CLEARANCE"
]

DG_COMPONENTS = [
    "TURBOCHARGER (2)", "TURBOCHARGER (3)", "COOLING WATER PUMP",
    "COOL WATER THERMOSTAT VALVE", "L.O. THERMOSTAT VALVE",
    "THRUST BEARING", "AIR COOLER", "L.O. COOLER CLEAN",
    "F.W. COOLER CLEAN", "L.O. RENEWAL", "ALTERNATOR CLEANING"
]

OTHER_SIMPLE_COMPONENTS = [
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
    "FUEL VALVES **(1)**": "FUEL VALVES (1)",
    "TURBOCHARGER(2)": "TURBOCHARGER (2)",
    "TURBOCHARGER(3)": "TURBOCHARGER (3)",
    "PLUNGER AND BARREL (RENEWAL)": "PLUNGER AND BARREL(RENEWAL)",
    "JACKET FW NO.1": "JACKET FW"
}

TEXTUAL_VALUES = {"N/A", "NO RECORD", "NOT WORKING", "CENTRAL", "COOLER"}

def fl(txt: Any) -> str:
    if txt is None:
        return ""
    raw = str(txt).replace("\x07", "").replace("\xa0", " ").replace("\t", " ")
    lines = [line.strip() for line in raw.split("\n") if line.strip()]
    return " ".join(lines) if lines else ""

def normalize_token(txt: Any) -> str:
    s = fl(txt).upper()
    s = s.replace("**", " ")
    s = s.replace("[", "").replace("]", "")
    s = re.sub(r"\s+", " ", s).strip(" :-#*")
    s = s.replace("SEPT", "SEP")
    return ALIASES.get(s, s)

def parse_num(txt: Any) -> Optional[float]:
    s = normalize_token(txt)
    if not s or s in TEXTUAL_VALUES or "OBSERVATION" in s:
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
    except Exception:
        return None

def parse_date(txt: Any) -> Tuple[Optional[str], Optional[str], bool]:
    raw = fl(txt).replace("[", "").replace("]", "").strip()
    if not raw or raw in {"-", ""}:
        return None, None, False
    norm = normalize_token(raw)
    if norm in {"1", "2"}:
        return None, None, False
    if norm in TEXTUAL_VALUES:
        return None, norm, False
    try:
        dt = dt_parser.parse(raw, dayfirst=True, fuzzy=False)
        return dt.date().isoformat(), None, False
    except Exception:
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

def component_sort_key(name: str, order: List[str]) -> int:
    try:
        return order.index(name)
    except ValueError:
        return 9999

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
            except Exception:
                pass
        shutil.rmtree(outdir, ignore_errors=True)

def read_docx_structure(docx_bytes: bytes) -> Tuple[List[str], List[List[List[str]]]]:
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
                vals = [fl(c.text) for c in row.cells]
                if any(v.strip() for v in vals):
                    rows.append(vals)
            if rows:
                tables.append(rows)
        return paragraphs, tables
    finally:
        try:
            os.unlink(tp)
        except Exception:
            pass

def all_rows(paragraphs: List[str], tables: List[List[List[str]]]) -> List[List[str]]:
    rows = []
    for p in paragraphs:
        rows.append([p])
    for tbl in tables:
        rows.extend(tbl)
    return rows

def row_text(row: List[str]) -> str:
    return " | ".join([fl(x) for x in row if fl(x)])

def norm_row(row: List[str]) -> List[str]:
    return [normalize_token(x) for x in row]

def find_first_row(rows: List[List[str]], predicate) -> Optional[int]:
    for i, row in enumerate(rows):
        if predicate(row):
            return i
    return None

def extract_header(rows: List[List[str]], review: List[Dict[str, str]]) -> Dict[str, Any]:
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
            review.append({
                "Section": "Header",
                "Item": "Report Date",
                "Issue": "Invalid date format",
                "Raw Value": raw or ""
            })

    m = re.search(r"TOTAL RUNNING HOURS\s*:?\s*([\d,\.]+)", text, re.I)
    if m:
        out["me_total_hours"] = parse_num(m.group(1))

    m = re.search(r"THIS MONTH\s*:?\s*([\d,\.]+)", text, re.I)
    if m:
        out["me_this_month"] = parse_num(m.group(1))

    return out

def find_me_start(rows: List[List[str]]) -> Optional[int]:
    return find_first_row(rows, lambda r: "MAIN ENGINE" in normalize_token(row_text(r)) or "CYL. NO." in normalize_token(row_text(r)))

def find_other_start(rows: List[List[str]]) -> Optional[int]:
    return find_first_row(rows, lambda r: "TURBOCHARGER" in normalize_token(row_text(r)) and "COOLERS" in normalize_token(row_text(r)))

def find_aux_start(rows: List[List[str]]) -> Optional[int]:
    return find_first_row(rows, lambda r: "AUX. ENGINE MAKER / TYPE" in normalize_token(row_text(r)))

def find_dg_start(rows: List[List[str]]) -> Optional[int]:
    return find_first_row(rows, lambda r: "D/G NO1" in normalize_token(row_text(r)) or "D/G NO.1" in normalize_token(row_text(r)))

def extract_me(rows: List[List[str]], me_start: Optional[int], other_start: Optional[int], review: List[Dict[str, str]]) -> List[Dict[str, Any]]:
    if me_start is None:
        return []
    end = other_start if other_start is not None and other_start > me_start else len(rows)
    zone = rows[me_start:end]
    out = []

    for i in range(len(zone) - 1):
        r1, r2 = zone[i], zone[i + 1]
        if len(r1) < 4 or len(r2) < 4:
            continue

        comp = normalize_token(r1[0])
        if comp not in ME_COMPONENTS:
            continue

        marker1 = normalize_token(r1[2]) if len(r1) > 2 else ""
        marker2 = normalize_token(r2[2]) if len(r2) > 2 else ""
        if marker1 != "1" or marker2 != "2":
            continue

        raw_period = r1[1] if len(r1) > 1 else ""
        period = parse_num(raw_period)
        observation_based = "OBSERVATION" in normalize_token(raw_period)
        if observation_based:
            review.append({
                "Section": "Main Engine",
                "Item": comp,
                "Issue": "Observation-based periodicity",
                "Raw Value": fl(raw_period)
            })

        max_cols = min(len(r1), len(r2))
        cyl_count = min(max_cols - 3, 7)

        for j in range(cyl_count):
            raw_date = r1[3 + j] if 3 + j < len(r1) else ""
            raw_hrs = r2[3 + j] if 3 + j < len(r2) else ""

            iso, raw_text, invalid = parse_date(raw_date)
            hrs = parse_num(raw_hrs)

            if invalid:
                review.append({
                    "Section": "Main Engine",
                    "Item": f"{comp} / Cyl {j+1}",
                    "Issue": "Invalid date",
                    "Raw Value": raw_text or ""
                })

            if raw_text in TEXTUAL_VALUES:
                review.append({
                    "Section": "Main Engine",
                    "Item": f"{comp} / Cyl {j+1}",
                    "Issue": "Textual date value kept",
                    "Raw Value": raw_text
                })

            if iso or raw_text or hrs is not None:
                used = (hrs / period * 100) if (hrs is not None and period and period > 0) else None
                out.append({
                    "Status": get_status(hrs, period),
                    "Component": comp,
                    "Engine": "ME",
                    "Unit": f"Cyl {j+1}",
                    "Periodicity": "OBSERVATION" if observation_based else (int(period) if period and float(period).is_integer() else (period if period is not None else "—")),
                    "Last O/H": iso or raw_text or "—",
                    "Hrs Since": format_hours(hrs),
                    "Used %": format_percent(used)
                })

    out.sort(key=lambda x: (component_sort_key(x["Component"], ME_COMPONENTS), int(re.search(r"\d+", x["Unit"]).group())))
    return out

def extract_aux_meta(zone: List[List[str]]) -> Dict[str, Any]:
    text = " ".join(row_text(r) for r in zone)
    out = {"aux_total_hours": None, "aux_this_month": None}

    m = re.search(r"TOTAL HOURS\s*:?\s*([\d,\.]+)", text, re.I)
    if m:
        out["aux_total_hours"] = parse_num(m.group(1))

    m = re.search(r"HOURS THIS MONTH\s*([\d,\.]+)", text, re.I)
    if m:
        out["aux_this_month"] = parse_num(m.group(1))

    return out

def find_aux_header_in_zone(zone: List[List[str]]) -> Optional[int]:
    for i, row in enumerate(zone):
        joined = " | ".join(norm_row(row))
        if "DESCRIPTION" in joined and "1" in joined and "2" in joined and ("PERIODICITY" in joined or "PERIODICTLY" in joined):
            return i
    return None

def extract_aux(rows: List[List[str]], aux_start: Optional[int], dg_start: Optional[int], review: List[Dict[str, str]]) -> Tuple[List[Dict[str, Any]], Dict[str, Any], List[str]]:
    if aux_start is None:
        return [], {"aux_total_hours": None, "aux_this_month": None}, AUX_COMPONENTS.copy()

    end = dg_start if dg_start is not None and dg_start > aux_start else len(rows)
    zone = rows[aux_start:end]
    meta = extract_aux_meta(zone)
    header_idx = find_aux_header_in_zone(zone)
    if header_idx is None:
        review.append({
            "Section": "Aux Engine",
            "Item": "Header",
            "Issue": "AUX detail header not found",
            "Raw Value": "DESCRIPTION | PERIODICITY | 1 | 2"
        })
        return [], meta, AUX_COMPONENTS.copy()

    body = zone[header_idx + 1:]
    out = []
    i = 0

    while i < len(body) - 1:
        r1, r2 = body[i], body[i + 1]
        comp = normalize_token(r1[0] if len(r1) > 0 else "")
        marker1 = normalize_token(r1[2] if len(r1) > 2 else "")
        marker2 = normalize_token(r2[2] if len(r2) > 2 else "")

        if comp in AUX_COMPONENTS and marker1 == "1" and marker2 == "2":
            raw_period = r1[1] if len(r1) > 1 else ""
            period = parse_num(raw_period)
            raw_date = r1[3] if len(r1) > 3 else ""
            raw_hrs = r2[3] if len(r2) > 3 else ""

            iso, raw_text, invalid = parse_date(raw_date)
            hrs = parse_num(raw_hrs)

            if invalid:
                review.append({
                    "Section": "Aux Engine",
                    "Item": comp,
                    "Issue": "Invalid date",
                    "Raw Value": raw_text or ""
                })

            if raw_text in TEXTUAL_VALUES:
                review.append({
                    "Section": "Aux Engine",
                    "Item": comp,
                    "Issue": "Textual date value kept",
                    "Raw Value": raw_text
                })

            if raw_hrs and parse_num(raw_hrs) is None and normalize_token(raw_hrs) not in {"", "2"}:
                review.append({
                    "Section": "Aux Engine",
                    "Item": comp,
                    "Issue": "Non-numeric running hours",
                    "Raw Value": fl(raw_hrs)
                })

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
    missing = [x for x in AUX_COMPONENTS if x not in found]
    for m in missing:
        review.append({
            "Section": "Aux Engine",
            "Item": m,
            "Issue": "Expected AUX component not found",
            "Raw Value": ""
        })

    out.sort(key=lambda x: component_sort_key(x["Component"], AUX_COMPONENTS))
    return out, meta, missing

def extract_other_simple_zone(rows: List[List[str]], review: List[Dict[str, str]]) -> List[Dict[str, Any]]:
    out = []
    for row in rows:
        cells = [fl(c) for c in row]
        for idx, cell in enumerate(cells):
            comp = normalize_token(cell)
            if comp not in OTHER_SIMPLE_COMPONENTS:
                continue

            period = parse_num(cells[idx + 1]) if idx + 1 < len(cells) else None
            raw_date = cells[idx + 2] if idx + 2 < len(cells) else ""
            raw_hrs = cells[idx + 3] if idx + 3 < len(cells) else ""

            iso, raw_text, invalid = parse_date(raw_date)
            hrs = parse_num(raw_hrs)

            if invalid:
                review.append({
                    "Section": "Other Equipment",
                    "Item": comp,
                    "Issue": "Invalid date",
                    "Raw Value": raw_text or ""
                })

            if raw_text in TEXTUAL_VALUES:
                review.append({
                    "Section": "Other Equipment",
                    "Item": comp,
                    "Issue": "Textual date value kept",
                    "Raw Value": raw_text
                })

            if raw_hrs and parse_num(raw_hrs) is None and normalize_token(raw_hrs) not in {"", "1", "2"}:
                review.append({
                    "Section": "Other Equipment",
                    "Item": comp,
                    "Issue": "Non-numeric running hours",
                    "Raw Value": fl(raw_hrs)
                })

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

def extract_dg_zone(rows: List[List[str]], review: List[Dict[str, str]]) -> List[Dict[str, Any]]:
    out = []
    for i in range(len(rows) - 1):
        r1, r2 = rows[i], rows[i + 1]
        comp = normalize_token(r1[0] if len(r1) > 0 else "")
        if comp not in DG_COMPONENTS:
            continue

        marker1 = normalize_token(r1[2] if len(r1) > 2 else "")
        marker2 = normalize_token(r2[2] if len(r2) > 2 else "")
        if marker1 != "1" or marker2 != "2":
            continue

        period = parse_num(r1[1] if len(r1) > 1 else "")

        for unit_idx in range(3):
            c = 3 + unit_idx
            raw_date = r1[c] if c < len(r1) else ""
            raw_hrs = r2[c] if c < len(r2) else ""

            iso, raw_text, invalid = parse_date(raw_date)
            hrs = parse_num(raw_hrs)

            if invalid:
                review.append({
                    "Section": "Other Equipment",
                    "Item": f"{comp} / DG {unit_idx+1}",
                    "Issue": "Invalid date",
                    "Raw Value": raw_text or ""
                })

            if raw_text in TEXTUAL_VALUES:
                review.append({
                    "Section": "Other Equipment",
                    "Item": f"{comp} / DG {unit_idx+1}",
                    "Issue": "Textual date value kept",
                    "Raw Value": raw_text
                })

            if raw_hrs and parse_num(raw_hrs) is None and normalize_token(raw_hrs) not in {"", "1", "2"}:
                review.append({
                    "Section": "Other Equipment",
                    "Item": f"{comp} / DG {unit_idx+1}",
                    "Issue": "Non-numeric running hours",
                    "Raw Value": fl(raw_hrs)
                })

            if iso or raw_text or hrs is not None:
                used = (hrs / period * 100) if (hrs is not None and period and period > 0) else None
                out.append({
                    "Status": get_status(hrs, period),
                    "Description": comp,
                    "Unit": f"DG {unit_idx+1}",
                    "Periodicity": int(period) if period and float(period).is_integer() else (period if period is not None else "—"),
                    "Last Date": iso or raw_text or "—",
                    "Run Hrs": format_hours(hrs),
                    "Used %": format_percent(used)
                })
    return out

def dedupe_records(records: List[Dict[str, Any]], keys: List[str]) -> List[Dict[str, Any]]:
    seen = set()
    out = []
    for r in records:
        k = tuple(r.get(x) for x in keys)
        if k not in seen:
            seen.add(k)
            out.append(r)
    return out

def extract_other(rows: List[List[str]], other_start: Optional[int], aux_start: Optional[int], dg_start: Optional[int], review: List[Dict[str, str]]) -> List[Dict[str, Any]]:
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
        component_sort_key(x["Description"], OTHER_SIMPLE_COMPONENTS + DG_COMPONENTS),
        999 if x["Unit"] == "—" else int(re.search(r"\d+", x["Unit"]).group())
    ))
    return parts

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

def render_html_table(df: pd.DataFrame, cols: List[str], numeric_cols: Optional[List[str]] = None, center_cols: Optional[List[str]] = None):
    numeric_cols = numeric_cols or []
    center_cols = center_cols or []

    html = ['<div class="table-wrap"><table class="report-table"><thead><tr>']
    for c in cols:
        html.append(f"<th>{c}</th>")
    html.append("</tr></thead><tbody>")

    for _, row in df[cols].iterrows():
        html.append("<tr>")
        for c in cols:
            val = row[c]
            cls = ""
            if c in numeric_cols:
                cls = ' class="num"'
            elif c in center_cols:
                cls = ' class="center"'
            safe = "—" if pd.isna(val) else str(val)
            html.append(f"<td{cls}>{safe}</td>")
        html.append("</tr>")
    html.append("</tbody></table></div>")

    st.markdown("".join(html), unsafe_allow_html=True)

st.markdown("""
<div class="hero-k">Running Hours Management System</div>
<div class="hero-h">TEC04 Extraction Matrix</div>
<div class="hero-rule"></div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("Upload TEC04 / TEC-004 report (.doc)", type=["doc"])

if uploaded:
    with st.spinner("Executing extraction with validation..."):
        try:
            docx_data = convert_doc_to_docx(uploaded.read())
            paragraphs, tables = read_docx_structure(docx_data)
            rows = all_rows(paragraphs, tables)

            review: List[Dict[str, str]] = []

            header = extract_header(rows, review)
            me_start = find_me_start(rows)
            other_start = find_other_start(rows)
            aux_start = find_aux_start(rows)
            dg_start = find_dg_start(rows)

            me_data = extract_me(rows, me_start, other_start, review)
            aux_data, aux_meta, aux_missing = extract_aux(rows, aux_start, dg_start, review)
            other_data = extract_other(rows, other_start, aux_start, dg_start, review)

            n_overdue = sum(1 for r in (me_data + aux_data + other_data) if r["Status"] == "OVERDUE")
            n_high = sum(1 for r in (me_data + aux_data + other_data) if r["Status"] == "HIGH PRIORITY")
            n_review = len(review)

            st.markdown(f"""
            <div class="metric-grid">
              <div class="metric"><div class="metric-v">{header.get('vessel') or 'UNKNOWN'}</div><div class="metric-l">Vessel</div></div>
              <div class="metric"><div class="metric-v">{header.get('report_date') or '—'}</div><div class="metric-l">Report Date</div></div>
              <div class="metric"><div class="metric-v">{format_hours(header.get('me_total_hours'))}</div><div class="metric-l">ME Total Hrs</div></div>
              <div class="metric"><div class="metric-v">{format_hours(header.get('me_this_month'))}</div><div class="metric-l">ME This Month</div></div>
              <div class="metric"><div class="metric-v">{n_overdue}</div><div class="metric-l">Overdue</div></div>
              <div class="metric"><div class="metric-v">{n_review}</div><div class="metric-l">Review Items</div></div>
            </div>
            """, unsafe_allow_html=True)

            if aux_missing:
                st.markdown(
                    f'<div class="banner banner-warn"><strong>Aux parsed as the visible 9-component block.</strong> '
                    f'Missing expected items: {", ".join(aux_missing)}</div>',
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    '<div class="banner banner-ok"><strong>Aux extracted successfully.</strong> '
                    'The app treats AUX as the single 9-component list shown in the report.</div>',
                    unsafe_allow_html=True
                )

            if review:
                st.markdown(
                    f'<div class="banner banner-note"><strong>Validation active.</strong> '
                    f'{len(review)} item(s) were flagged for operator review instead of being silently discarded.</div>',
                    unsafe_allow_html=True
                )

            tab1, tab2, tab3, tab4 = st.tabs([
                f"Main Engine ({len(me_data)})",
                f"Aux Engine ({len(aux_data)})",
                f"Other Equipment ({len(other_data)})",
                f"Review Queue ({len(review)})"
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
                        f'<div class="banner banner-note"><strong>Aux Meta:</strong> '
                        f'Total Hours = {format_hours(aux_meta.get("aux_total_hours"))}, '
                        f'Hours This Month = {format_hours(aux_meta.get("aux_this_month"))}</div>',
                        unsafe_allow_html=True
                    )

            with tab3:
                if not other_data:
                    st.info("No Other Equipment records found.")
                else:
                    df_other = pd.DataFrame(other_data)
                    df_other["Status"] = df_other["Status"].apply(render_status_chip)
                    render_html_table(
                        df_other,
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
