import os
import re
import shutil
import tempfile
import subprocess
import difflib
from pathlib import Path
from dataclasses import dataclass
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
  --bg2:#0c1623;
  --bg3:#122031;
  --line:#1f3349;
  --gold:#c99818;
  --text:#edf4ff;
  --muted:#9db3c7;
  --soft:#71879b;

  --ok:#1f9d68;
  --ok-bg:rgba(31,157,104,.14);
  --warn:#d48a10;
  --warn-bg:rgba(212,138,16,.16);
  --bad:#d14b3f;
  --bad-bg:rgba(209,75,63,.16);
  --nodata:#6d8297;
  --nodata-bg:rgba(109,130,151,.16);
}

html, body, [class*="css"]{
  background:var(--bg)!important;
  color:var(--muted)!important;
  font-family:'Inter', sans-serif!important;
}

.main, .block-container{
  background:radial-gradient(circle at top right, rgba(201,152,24,.05), transparent 28%), var(--bg)!important;
}

[data-testid="stSidebar"], [data-testid="collapsedControl"]{
  display:none!important;
}

.block-container{
  padding-top:1.2rem!important;
  animation:fadeUp .55s ease-out;
}

.hero-k{
  font-size:.68rem;
  letter-spacing:.24em;
  text-transform:uppercase;
  color:var(--gold);
  font-weight:700;
  animation:fadeIn .55s ease-out;
}

.hero-h{
  font-family:'Space Grotesk', sans-serif;
  font-size:2rem;
  font-weight:700;
  color:var(--text);
  line-height:1.05;
  margin-top:.2rem;
  animation:fadeUp .65s ease-out;
}

.hero-rule{
  height:1px;
  margin:1rem 0 1.2rem 0;
  background:linear-gradient(90deg,var(--gold),var(--line),transparent);
  animation:glowLine 1.2s ease-out;
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
  box-shadow:0 10px 28px rgba(0,0,0,.18);
  transform:translateY(0);
  transition:transform .22s ease, box-shadow .22s ease, border-color .22s ease;
  animation:cardIn .45s ease-out both;
}

.metric:hover{
  transform:translateY(-2px);
  box-shadow:0 14px 34px rgba(0,0,0,.24);
  border-color:#32506d;
}

.metric-v{
  font-family:'Space Grotesk', sans-serif;
  font-size:1.38rem;
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
  background:linear-gradient(180deg,var(--bg2),#0f1b2a);
  border:1px solid var(--line);
  border-radius:12px;
  color:var(--muted);
  padding:.6rem 1rem;
  font-weight:600;
  transition:all .2s ease;
}

.stTabs [data-baseweb="tab"]:hover{
  border-color:#35516a;
  color:var(--text);
}

.stTabs [aria-selected="true"]{
  background:linear-gradient(180deg,#14263a,#102131)!important;
  border-color:#36546e!important;
  color:var(--text)!important;
  box-shadow:inset 0 0 0 1px rgba(201,152,24,.22), 0 6px 20px rgba(0,0,0,.18);
}

.stTabs [data-baseweb="tab-panel"]{
  padding-top:1rem;
  animation:fadeUp .35s ease-out;
}

[data-testid="stFileUploadDropzone"]{
  background:rgba(201,152,24,.04)!important;
  border:1.25px dashed rgba(201,152,24,.32)!important;
  border-radius:14px!important;
  transition:all .2s ease;
}

[data-testid="stFileUploadDropzone"]:hover{
  border-color:rgba(201,152,24,.55)!important;
  background:rgba(201,152,24,.07)!important;
}

.banner{
  border:1px solid var(--line);
  background:var(--bg2);
  border-radius:12px;
  padding:.9rem 1rem;
  margin:.7rem 0 1rem 0;
  animation:fadeUp .4s ease-out;
}

.banner-ok{border-left:3px solid var(--ok);}
.banner-warn{border-left:3px solid var(--warn);}
.banner-bad{border-left:3px solid var(--bad);}

.kicker{
  color:var(--text);
  font-weight:600;
}

.small{
  color:var(--soft);
  font-size:.9rem;
}

.table-wrap{
  overflow-x:auto;
  border:1px solid var(--line);
  border-radius:14px;
  background:linear-gradient(180deg, rgba(255,255,255,.01), rgba(255,255,255,.00)), var(--bg2);
  animation:fadeUp .35s ease-out;
}

.report-table{
  width:100%;
  border-collapse:separate;
  border-spacing:0;
  min-width:920px;
}

.report-table thead th{
  position:sticky;
  top:0;
  z-index:1;
  background:#102030;
  color:var(--text);
  font-size:.76rem;
  text-transform:uppercase;
  letter-spacing:.08em;
  padding:.9rem .8rem;
  border-bottom:1px solid var(--line);
  white-space:nowrap;
}

.report-table tbody td{
  padding:.82rem .8rem;
  border-bottom:1px solid rgba(31,51,73,.75);
  color:var(--muted);
  font-size:.94rem;
  vertical-align:middle;
}

.report-table tbody tr:nth-child(even){
  background:rgba(255,255,255,.015);
}

.report-table tbody tr:hover{
  background:rgba(201,152,24,.045);
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
  gap:.4rem;
  padding:.28rem .58rem;
  border-radius:999px;
  font-size:.78rem;
  font-weight:700;
  letter-spacing:.02em;
  white-space:nowrap;
}

.status-overdue{
  color:#ffd8d3;
  background:var(--bad-bg);
  border:1px solid rgba(209,75,63,.35);
}

.status-high{
  color:#ffe7b8;
  background:var(--warn-bg);
  border:1px solid rgba(212,138,16,.35);
}

.status-ok{
  color:#d9ffef;
  background:var(--ok-bg);
  border:1px solid rgba(31,157,104,.35);
}

.status-nodata{
  color:#d8e5f1;
  background:var(--nodata-bg);
  border:1px solid rgba(109,130,151,.35);
}

@keyframes fadeIn{
  from{opacity:0}
  to{opacity:1}
}

@keyframes fadeUp{
  from{opacity:0; transform:translateY(10px)}
  to{opacity:1; transform:translateY(0)}
}

@keyframes cardIn{
  from{opacity:0; transform:translateY(14px) scale(.985)}
  to{opacity:1; transform:translateY(0) scale(1)}
}

@keyframes glowLine{
  from{opacity:0; transform:scaleX(.85); transform-origin:left}
  to{opacity:1; transform:scaleX(1); transform-origin:left}
}
</style>
""", unsafe_allow_html=True)

ME_COMPONENTS = {
    "CYLINDER COVER", "PISTON ASSEMBLY", "STUFFING BOX", "PISTON CROWN",
    "CYLINDER LINER", "EXHAUST VALVE", "STARTING VALVE", "SAFETY VALVE",
    "FUEL VALVES", "FUEL PUMP", "PLUNGER AND BARREL(RENEWAL)",
    "PLUNGER AND BARREL", "FUEL PUMP SUCTION VALVE",
    "FUEL PUMP PUNCTURE VALVE", "CROSSHEAD BEARINGS",
    "BOTTOM END BEARINGS", "MAIN BEARINGS"
}

AUX_COMPONENTS = {
    "CYLINDER HEAD", "PISTON", "CONNECTING ROD", "CYLINDER LINERS",
    "FUEL VALVES (1)", "FUEL PUMPS", "CRANK PIN BEARING",
    "MAIN BEARING", "ADJUST VALVE HEAD CLEARANCE"
}

DG_COMPONENTS = {
    "TURBOCHARGER (2)", "TURBOCHARGER (3)", "COOLING WATER PUMP",
    "COOL WATER THERMOSTAT VALVE", "L.O. THERMOSTAT VALVE",
    "THRUST BEARING", "AIR COOLER", "L.O. COOLER CLEAN",
    "F.W. COOLER CLEAN", "L.O. RENEWAL", "ALTERNATOR CLEANING"
}

OE_COMPONENTS = {
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

ME_ORDER = [
    "CYLINDER COVER", "PISTON ASSEMBLY", "STUFFING BOX", "PISTON CROWN",
    "CYLINDER LINER", "EXHAUST VALVE", "STARTING VALVE", "SAFETY VALVE",
    "FUEL VALVES", "FUEL PUMP", "PLUNGER AND BARREL(RENEWAL)",
    "PLUNGER AND BARREL", "FUEL PUMP SUCTION VALVE",
    "FUEL PUMP PUNCTURE VALVE", "CROSSHEAD BEARINGS",
    "BOTTOM END BEARINGS", "MAIN BEARINGS"
]

AUX_ORDER = [
    "CYLINDER HEAD", "PISTON", "CONNECTING ROD", "CYLINDER LINERS",
    "FUEL VALVES (1)", "FUEL PUMPS", "CRANK PIN BEARING",
    "MAIN BEARING", "ADJUST VALVE HEAD CLEARANCE"
]

DG_ORDER = [
    "TURBOCHARGER (2)", "TURBOCHARGER (3)", "COOLING WATER PUMP",
    "COOL WATER THERMOSTAT VALVE", "L.O. THERMOSTAT VALVE",
    "THRUST BEARING", "AIR COOLER", "L.O. COOLER CLEAN",
    "F.W. COOLER CLEAN", "L.O. RENEWAL", "ALTERNATOR CLEANING"
]

ALIASES = {
    "PERIODICTLY": "PERIODICITY",
    "PLUNGER AND BARREL (RENEWAL)": "PLUNGER AND BARREL(RENEWAL)",
    "PLUNGER AND BARREL  (RENEWAL)": "PLUNGER AND BARREL(RENEWAL)",
    "FUEL VALVES(1)": "FUEL VALVES (1)",
    "TURBOCHARGER(2)": "TURBOCHARGER (2)",
    "TURBOCHARGER(3)": "TURBOCHARGER (3)",
    "EXAUST VALVE": "EXHAUST VALVE",
    "JACKET FW NO.1": "JACKET FW",
    "PERIODICITY": "PERIODICITY",
}

KNOWN_TEXT_STATES = {
    "N/A", "NO RECORD", "NOT WORKING", "CENTRAL", "COOLER"
}

SECTION_COMPONENTS = {
    "ME": ME_COMPONENTS,
    "AUX": AUX_COMPONENTS,
    "DG": DG_COMPONENTS,
    "OE": OE_COMPONENTS,
}

@dataclass
class WarningItem:
    section: str
    severity: str
    message: str
    source: str = ""

def fl(txt: Any) -> str:
    if txt is None:
        return ""
    raw = str(txt).replace("\x07", "").replace("\xa0", " ").replace("\t", " ")
    lines = [line.strip() for line in raw.split("\n") if line.strip()]
    return " ".join(lines) if lines else ""

def normalize_token(txt: Any) -> str:
    s = re.sub(r"\s+", " ", fl(txt).upper()).strip(" :-#*[]")
    s = s.replace("SEPT", "SEP")
    return ALIASES.get(s, s)

def add_warning(warnings: List[WarningItem], section: str, severity: str, message: str, source: str = ""):
    warnings.append(WarningItem(section=section, severity=severity, message=message, source=source))

def parse_num(txt: Any) -> Optional[float]:
    s = fl(txt).upper().replace("[", "").replace("]", "").strip()
    if not s or s in {"-", ""}:
        return None
    if s in KNOWN_TEXT_STATES:
        return None
    if "OBSERVATION" in s:
        return None

    m = re.search(r"\d[\d,\.]*", s)
    if not m:
        return None

    token = m.group()

    if re.fullmatch(r"\d{1,3}(\.\d{3})+", token) or re.fullmatch(r"\d{1,3}(,\d{3})+", token):
        token = token.replace(".", "").replace(",", "")
    elif token.count(".") == 1 and token.count(",") == 0:
        pass
    elif token.count(",") == 1 and token.count(".") == 0:
        parts = token.split(",")
        if len(parts[-1]) == 3:
            token = token.replace(",", "")
        else:
            token = token.replace(",", ".")
    else:
        token = token.replace(",", "")

    try:
        return float(token)
    except Exception:
        return None

def parse_date(txt: Any) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    raw = fl(txt).replace("[", "").replace("]", "").strip()
    if not raw or raw in {"-", ""}:
        return None, None, None

    norm = normalize_token(raw)

    if norm in KNOWN_TEXT_STATES:
        return None, norm, None
    if norm in {"1", "2"}:
        return None, None, None
    if len(raw) > 40:
        return None, None, raw

    try:
        dt = dt_parser.parse(raw, dayfirst=True, fuzzy=False)
        return dt.date().isoformat(), None, None
    except Exception:
        return None, None, raw

def format_hours(value: Optional[float]) -> str:
    if value is None:
        return "—"
    if float(value).is_integer():
        return f"{int(value):,}"
    return f"{value:,.1f}"

def format_percent(value: Optional[float]) -> str:
    if value is None:
        return "—"
    return f"{value:.1f}%"

def get_status(hrs: Optional[float], period: Optional[float]) -> str:
    if hrs is None or period is None or period <= 0:
        return "NO DATA"
    ratio = hrs / period
    if ratio >= 1.0:
        return "OVERDUE"
    if ratio >= 0.8:
        return "HIGH PRIORITY"
    return "OK"

def guarded_component_match(token: str, section: str) -> Tuple[str, bool, bool]:
    valid_components = SECTION_COMPONENTS.get(section, set())
    if token in valid_components:
        return token, False, False

    matches = difflib.get_close_matches(token, valid_components, n=2, cutoff=0.90)
    if not matches:
        return token, False, False
    if len(matches) == 1:
        return matches[0], True, False

    r1 = difflib.SequenceMatcher(None, token, matches[0]).ratio()
    r2 = difflib.SequenceMatcher(None, token, matches[1]).ratio()

    if (r1 - r2) < 0.04:
        return token, False, True

    return matches[0], True, False

def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError(
            "LibreOffice not found in runtime environment. "
            "Install soffice for .doc support or upload .docx files."
        )

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
                "--convert-to",
                "docx",
                src,
                "--outdir",
                outdir,
            ],
            capture_output=True,
            timeout=120,
            text=True,
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

def doc_to_grid(docx_bytes: bytes) -> Dict[str, Any]:
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as t:
        t.write(docx_bytes)
        temp_path = t.name

    try:
        doc = Document(temp_path)
        paragraphs = [fl(p.text) for p in doc.paragraphs if fl(p.text)]
        tables = []
        for ti, table in enumerate(doc.tables):
            rows = []
            for row in table.rows:
                vals = [fl(c.text) for c in row.cells]
                if any(vals):
                    rows.append(vals)
            if rows:
                tables.append({"table_index": ti, "rows": rows})
        return {"paragraphs": paragraphs, "tables": tables}
    finally:
        try:
            os.unlink(temp_path)
        except Exception:
            pass

def classify_table(rows: List[List[str]]) -> str:
    blob = " ".join(" ".join(r) for r in rows).upper()
    row_blob = [" ".join(r).upper() for r in rows]

    scores = {"ME": 0, "AUX": 0, "DG": 0, "OE": 0}

    if "MAIN ENGINE" in blob:
        scores["ME"] += 3
    if "CYL. NO." in blob or "CYL NO." in blob:
        scores["ME"] += 2
    if "DATE OF LAST O/H" in blob and "RUNNING HOURS SINCE LAST O/H" in blob:
        scores["ME"] += 2
    if any(comp in blob for comp in ["CYLINDER COVER", "PISTON ASSEMBLY", "EXHAUST VALVE", "EXAUST VALVE"]):
        scores["ME"] += 2

    if "AUX. ENGINE MAKER / TYPE" in blob or "AUXILIARY ENGINE" in blob:
        scores["AUX"] += 4
    if "TOTAL HOURS" in blob and "HOURS THIS MONTH" in blob:
        scores["AUX"] += 2
    if any("DESCRIPTION PERIODICITY 1 2" in re.sub(r"\s+", " ", rb) for rb in row_blob):
        scores["AUX"] += 2

    if "D/G NO1" in blob or "D/G NO.1" in blob or "DIESEL GENERATOR" in blob:
        scores["DG"] += 4
    if "ALTERNATOR CLEANING" in blob:
        scores["DG"] += 2
    if any(comp in blob for comp in ["TURBOCHARGER (2)", "TURBOCHARGER (3)", "COOLING WATER PUMP"]):
        scores["DG"] += 2

    if any(h in blob for h in ["TURBOCHARGER", "COOLERS", "A/C & REFR. COMPRESSORS", "AUXILIARY BOILER", "MAIN AIR COMPRESSORS"]):
        scores["OE"] += 3
    if any(comp in blob for comp in ["GENERAL O/H", "BALANCING OF ROTOR SHAFT", "SERVICE AIR COMPRESSOR"]):
        scores["OE"] += 2

    best = max(scores, key=scores.get)

    if best == "DG" and scores["DG"] < 4:
        return "OTHER"
    if best == "AUX" and scores["AUX"] < 4:
        return "OTHER"
    if scores[best] == 0:
        return "OTHER"

    return best

def extract_header(model: Dict[str, Any], warnings: List[WarningItem]) -> Dict[str, Any]:
    text = " ".join(
        model["paragraphs"] +
        [" ".join(" ".join(r) for r in t["rows"]) for t in model["tables"]]
    )

    out = {
        "vessel": "UNKNOWN",
        "report_date": None,
        "me_total_hours": None,
        "me_this_month": None,
        "me_type": None,
        "aux_total_hours": None,
        "aux_this_month": None,
    }

    m = re.search(
        r"VESSEL['’]S NAME\s*:\s*(?:MV\s+)?(.+?)\s+DATE\s*:?\s*([A-Z0-9/ .-]+)",
        text,
        re.I,
    )
    if m:
        out["vessel"] = re.sub(r"(?i)^MV\s+", "", fl(m.group(1))).strip()
        iso, state, bad = parse_date(m.group(2))
        out["report_date"] = iso or state or fl(m.group(2))
        if bad:
            add_warning(warnings, "Header", "warning", f"Could not normalize report date: {bad}", bad)
    else:
        add_warning(warnings, "Header", "error", "Failed to extract vessel/date header")

    m = re.search(r"TYPE\s*:\s*([A-Z0-9&/\-– ]+?)\s+TOTAL RUNNING HOURS", text, re.I)
    if m:
        candidate = fl(m.group(1))
        if candidate and candidate != ":":
            out["me_type"] = candidate

    m = re.search(r"TOTAL RUNNING HOURS\s*:?\s*([\d,\.]+)", text, re.I)
    if m:
        out["me_total_hours"] = parse_num(m.group(1))

    m = re.search(r"THIS MONTH\s*:?\s*([\d,\.]+)", text, re.I)
    if m:
        out["me_this_month"] = parse_num(m.group(1))

    return out

def extract_me(table_rows: List[List[str]], warnings: List[WarningItem]) -> List[Dict[str, Any]]:
    records = []

    for i in range(len(table_rows) - 1):
        row1 = table_rows[i]
        row2 = table_rows[i + 1]

        if not row1 or not row2:
            continue

        comp_raw = normalize_token(row1[0]) if len(row1) > 0 else ""
        comp, was_fuzzy, ambiguous = guarded_component_match(comp_raw, "ME")

        if ambiguous:
            add_warning(warnings, "Main Engine", "warning", f"Ambiguous component label: {comp_raw}", comp_raw)
            continue

        marker_1 = normalize_token(row1[2] if len(row1) > 2 else "")
        marker_2 = normalize_token(row2[2] if len(row2) > 2 else "")

        if comp not in ME_COMPONENTS or marker_1 != "1" or marker_2 != "2":
            continue

        if was_fuzzy:
            add_warning(warnings, "Main Engine", "warning", f"Fuzzy-matched component '{comp_raw}' -> '{comp}'", comp_raw)

        periodicity_cell = row1[1] if len(row1) > 1 else ""
        periodicity = parse_num(periodicity_cell)
        observation_based = "OBSERVATION" in normalize_token(periodicity_cell)

        max_cols = min(len(row1), len(row2))
        cyl_count = min(max(0, max_cols - 3), 7)

        for j in range(cyl_count):
            raw_date = row1[3 + j] if 3 + j < len(row1) else ""
            raw_hrs = row2[3 + j] if 3 + j < len(row2) else ""

            iso, state, bad_date = parse_date(raw_date)
            hrs = parse_num(raw_hrs)

            if bad_date:
                add_warning(warnings, "Main Engine", "warning", f"Invalid date for {comp} cyl {j+1}: {bad_date}", raw_date)

            if fl(raw_hrs) and hrs is None and normalize_token(raw_hrs) not in KNOWN_TEXT_STATES:
                add_warning(warnings, "Main Engine", "warning", f"Non-numeric hours for {comp} cyl {j+1}", raw_hrs)

            if iso or state or hrs is not None or fl(raw_date) or fl(raw_hrs):
                ratio = (hrs / periodicity) if (hrs is not None and periodicity and periodicity > 0) else None
                records.append({
                    "Status": get_status(hrs, periodicity),
                    "Component": comp,
                    "Engine": "ME",
                    "Unit": f"Cyl {j+1}",
                    "Periodicity": "OBSERVATION" if observation_based else (int(periodicity) if periodicity and float(periodicity).is_integer() else periodicity or "—"),
                    "Last O/H": iso or state or (raw_date if fl(raw_date) else "—"),
                    "Hrs Since": hrs,
                    "Hrs Since Display": format_hours(hrs),
                    "Used Ratio": ratio if ratio is not None else 0.0,
                    "Used %": round(ratio * 100, 1) if ratio is not None else None,
                })

    if not records:
        add_warning(warnings, "Main Engine", "error", "No main engine records extracted")

    return records

def extract_aux(table_rows: List[List[str]], warnings: List[WarningItem]) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    records = []
    meta = {"aux_total_hours": None, "aux_this_month": None}

    blob = " ".join(" ".join(r) for r in table_rows)
    totals = re.findall(r"TOTAL HOURS:?\s*([\d,\.]+)", blob, re.I)
    months = re.findall(r"HOURS THIS MONTH\s*([\d,\.]+)", blob, re.I)

    if totals:
        meta["aux_total_hours"] = parse_num(totals[0])
    if months:
        meta["aux_this_month"] = parse_num(months[0])

    start = None
    stop = None

    for idx, row in enumerate(table_rows):
        normed = [normalize_token(c) for c in row]
        row_join = " ".join(normed)

        has_desc = any(c == "DESCRIPTION" for c in normed)
        has_period = any(c in {"PERIODICITY", "PERIODICTLY"} for c in normed) or "PERIODICITY" in row_join
        has_1 = "1" in normed
        has_2 = "2" in normed

        if has_desc and has_period and has_1 and has_2:
            start = idx + 1
            continue

        if any("D/G NO1" in c or "D/G NO.1" in c for c in normed):
            stop = idx
            break

    if start is None:
        add_warning(warnings, "Aux Engine", "error", "Aux description block not found")
        return records, meta

    end_idx = stop if stop is not None else len(table_rows)

    for i in range(start, end_idx - 1):
        row1 = table_rows[i]
        row2 = table_rows[i + 1]

        if not row1 or not row2:
            continue

        comp_raw = normalize_token(row1[0]) if len(row1) > 0 else ""
        comp, was_fuzzy, ambiguous = guarded_component_match(comp_raw, "AUX")

        if ambiguous:
            add_warning(warnings, "Aux Engine", "warning", f"Ambiguous component label: {comp_raw}", comp_raw)
            continue

        marker_1 = normalize_token(row1[2] if len(row1) > 2 else "")
        marker_2 = normalize_token(row2[2] if len(row2) > 2 else "")

        if comp not in AUX_COMPONENTS or marker_1 != "1" or marker_2 != "2":
            continue

        if was_fuzzy:
            add_warning(warnings, "Aux Engine", "warning", f"Fuzzy-matched component '{comp_raw}' -> '{comp}'", comp_raw)

        periodicity = parse_num(row1[1] if len(row1) > 1 else "")
        raw_date = row1[3] if len(row1) > 3 else ""
        raw_hrs = row2[3] if len(row2) > 3 else ""

        iso, state, bad_date = parse_date(raw_date)
        hrs = parse_num(raw_hrs)

        if bad_date:
            add_warning(warnings, "Aux Engine", "warning", f"Invalid date for {comp}: {bad_date}", raw_date)
        if fl(raw_hrs) and hrs is None and normalize_token(raw_hrs) not in KNOWN_TEXT_STATES:
            add_warning(warnings, "Aux Engine", "warning", f"Non-numeric hours for {comp}", raw_hrs)

        ratio = (hrs / periodicity) if (hrs is not None and periodicity and periodicity > 0) else None

        records.append({
            "Status": get_status(hrs, periodicity),
            "Component": comp,
            "Engine": "AUX-1",
            "Unit": "Engine",
            "Periodicity": int(periodicity) if periodicity and float(periodicity).is_integer() else periodicity or "—",
            "Last O/H": iso or state or (raw_date if fl(raw_date) else "—"),
            "Hrs Since": hrs,
            "Hrs Since Display": format_hours(hrs),
            "Used Ratio": ratio if ratio is not None else 0.0,
            "Used %": round(ratio * 100, 1) if ratio is not None else None,
        })

    if not records:
        add_warning(warnings, "Aux Engine", "error", "No auxiliary engine rows extracted after header detection")

    return records, meta

def extract_dg(table_rows: List[List[str]], warnings: List[WarningItem]) -> List[Dict[str, Any]]:
    records = []
    start = None

    for idx, row in enumerate(table_rows):
        joined = re.sub(r"\s+", " ", " ".join(normalize_token(c) for c in row)).strip()
        if "DESCRIPTION PERIODICITY D/G NO1" in joined or "DESCRIPTION PERIODICITY D/G NO.1" in joined:
            start = idx + 1
            break

    if start is None:
        return records

    for i in range(start, len(table_rows) - 1):
        row1 = table_rows[i]
        row2 = table_rows[i + 1]

        if not row1 or not row2:
            continue

        comp_raw = normalize_token(row1[0]) if len(row1) > 0 else ""
        comp, was_fuzzy, ambiguous = guarded_component_match(comp_raw, "DG")

        if ambiguous:
            add_warning(warnings, "D/G Equipment", "warning", f"Ambiguous component label: {comp_raw}", comp_raw)
            continue

        marker_1 = normalize_token(row1[2] if len(row1) > 2 else "")
        marker_2 = normalize_token(row2[2] if len(row2) > 2 else "")

        if comp not in DG_COMPONENTS or marker_1 != "1" or marker_2 != "2":
            continue

        if was_fuzzy:
            add_warning(warnings, "D/G Equipment", "warning", f"Fuzzy-matched component '{comp_raw}' -> '{comp}'", comp_raw)

        periodicity = parse_num(row1[1] if len(row1) > 1 else "")

        for gen_idx in range(3):
            col_idx = 3 + gen_idx
            raw_date = row1[col_idx] if col_idx < len(row1) else ""
            raw_hrs = row2[col_idx] if col_idx < len(row2) else ""

            iso, state, bad_date = parse_date(raw_date)
            hrs = parse_num(raw_hrs)

            if bad_date:
                add_warning(warnings, "D/G Equipment", "warning", f"Invalid date for {comp} generator {gen_idx+1}: {bad_date}", raw_date)
            if fl(raw_hrs) and hrs is None and normalize_token(raw_hrs) not in KNOWN_TEXT_STATES:
                add_warning(warnings, "D/G Equipment", "warning", f"Non-numeric hours for {comp} generator {gen_idx+1}", raw_hrs)

            if iso or state or hrs is not None or fl(raw_date) or fl(raw_hrs):
                ratio = (hrs / periodicity) if (hrs is not None and periodicity and periodicity > 0) else None
                records.append({
                    "Status": get_status(hrs, periodicity),
                    "Component": comp,
                    "Engine": f"D/G {gen_idx+1}",
                    "Unit": "Engine",
                    "Periodicity": int(periodicity) if periodicity and float(periodicity).is_integer() else periodicity or "—",
                    "Last O/H": iso or state or (raw_date if fl(raw_date) else "—"),
                    "Hrs Since": hrs,
                    "Hrs Since Display": format_hours(hrs),
                    "Used Ratio": ratio if ratio is not None else 0.0,
                    "Used %": round(ratio * 100, 1) if ratio is not None else None,
                })

    return records

def extract_oe(table_rows: List[List[str]], warnings: List[WarningItem]) -> List[Dict[str, Any]]:
    records = []

    for row in table_rows:
        cells = [fl(c) for c in row]

        for idx, cell in enumerate(cells):
            comp_raw = normalize_token(cell)
            comp, was_fuzzy, ambiguous = guarded_component_match(comp_raw, "OE")

            if ambiguous:
                continue
            if comp not in OE_COMPONENTS:
                continue

            if was_fuzzy:
                add_warning(warnings, "Other Equipment", "warning", f"Fuzzy-matched component '{comp_raw}' -> '{comp}'", comp_raw)

            periodicity = parse_num(cells[idx + 1]) if idx + 1 < len(cells) else None
            raw_date = cells[idx + 2] if idx + 2 < len(cells) else ""
            raw_hrs = cells[idx + 3] if idx + 3 < len(cells) else ""

            iso, state, bad_date = parse_date(raw_date)
            hrs = parse_num(raw_hrs)

            if bad_date:
                add_warning(warnings, "Other Equipment", "warning", f"Invalid date for {comp}: {bad_date}", raw_date)
            if fl(raw_hrs) and hrs is None and normalize_token(raw_hrs) not in KNOWN_TEXT_STATES:
                add_warning(warnings, "Other Equipment", "warning", f"Non-numeric hours for {comp}", raw_hrs)

            if iso or state or hrs is not None or fl(raw_date) or fl(raw_hrs):
                records.append({
                    "Section": "Other Equipment",
                    "Description": comp,
                    "Periodicity": int(periodicity) if periodicity and float(periodicity).is_integer() else periodicity or "—",
                    "Last Date": iso or state or (raw_date if fl(raw_date) else "—"),
                    "Run Hrs": hrs,
                    "Run Hrs Display": format_hours(hrs),
                })

    dedup = []
    seen = set()
    for r in records:
        key = tuple(r.items())
        if key not in seen:
            seen.add(key)
            dedup.append(r)

    return dedup

def extract_telemetry(docx_bytes: bytes) -> Dict[str, Any]:
    model = doc_to_grid(docx_bytes)
    warnings: List[WarningItem] = []

    header = extract_header(model, warnings)

    me_rows = []
    aux_rows = []
    dg_rows = []
    oe_rows = []

    for table in model["tables"]:
        rows = table["rows"]
        table_type = classify_table(rows)

        if table_type == "ME":
            me_rows.extend(extract_me(rows, warnings))
        elif table_type == "AUX":
            aux_part, aux_meta = extract_aux(rows, warnings)
            aux_rows.extend(aux_part)
            if header.get("aux_total_hours") is None:
                header["aux_total_hours"] = aux_meta.get("aux_total_hours")
            if header.get("aux_this_month") is None:
                header["aux_this_month"] = aux_meta.get("aux_this_month")
        elif table_type == "DG":
            dg_rows.extend(extract_dg(rows, warnings))
        elif table_type == "OE":
            oe_rows.extend(extract_oe(rows, warnings))

    return {
        "header": header,
        "me_rows": me_rows,
        "aux_rows": aux_rows,
        "dg_rows": dg_rows,
        "oe_rows": oe_rows,
        "warnings": warnings,
    }

def apply_report_sort(df: pd.DataFrame, order_list: List[str], unit_col: str = "Unit", engine_col: str = "Engine") -> pd.DataFrame:
    if df.empty:
        return df.copy()

    order_map = {name: idx for idx, name in enumerate(order_list)}
    out = df.copy()
    out["_component_order"] = out["Component"].map(order_map).fillna(9999)

    if unit_col in out.columns:
        out["_unit_num"] = out[unit_col].astype(str).str.extract(r"(\d+)").fillna(999).astype(int)
    else:
        out["_unit_num"] = 999

    if engine_col in out.columns:
        out["_engine_num"] = out[engine_col].astype(str).str.extract(r"(\d+)").fillna(1).astype(int)
    else:
        out["_engine_num"] = 1

    return out.sort_values(by=["_component_order", "_engine_num", "_unit_num"]).drop(
        columns=["_component_order", "_engine_num", "_unit_num"]
    )

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

def render_warning_summary(warnings: List[WarningItem]):
    if not warnings:
        st.markdown(
            '<div class="banner banner-ok"><div class="kicker">Extraction completed with no validation warnings.</div></div>',
            unsafe_allow_html=True
        )
        return

    n_err = sum(1 for w in warnings if w.severity == "error")
    n_warn = sum(1 for w in warnings if w.severity == "warning")

    cls = "banner-bad" if n_err else "banner-warn"
    st.markdown(
        f'<div class="banner {cls}"><div class="kicker">Validation notes: {n_err} errors, {n_warn} warnings.</div>'
        f'<div class="small">The parser kept extraction conservative where labels or values were ambiguous.</div></div>',
        unsafe_allow_html=True
    )

    with st.expander("View validation notes"):
        df_warn = pd.DataFrame([w.__dict__ for w in warnings])
        st.dataframe(df_warn, use_container_width=True, hide_index=True)

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

            if c == "Status":
                html.append(f"<td{cls}>{val}</td>")
            else:
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
    with st.spinner("Executing resilient extraction..."):
        try:
            docx_data = convert_doc_to_docx(uploaded.read())
            payload = extract_telemetry(docx_data)

            header = payload["header"]
            me_data = payload["me_rows"]
            aux_data = payload["aux_rows"]
            dg_data = payload["dg_rows"]
            oe_data = payload["oe_rows"]
            warnings = payload["warnings"]

            combined_critical = me_data + aux_data + dg_data
            n_od = sum(1 for c in combined_critical if c["Status"] == "OVERDUE")
            n_hp = sum(1 for c in combined_critical if c["Status"] == "HIGH PRIORITY")

            st.markdown(f"""
            <div class="metric-grid">
              <div class="metric"><div class="metric-v">{header.get('vessel') or 'UNKNOWN'}</div><div class="metric-l">Vessel</div></div>
              <div class="metric"><div class="metric-v">{header.get('report_date') or '—'}</div><div class="metric-l">Report Date</div></div>
              <div class="metric"><div class="metric-v">{format_hours(header.get('me_total_hours'))}</div><div class="metric-l">ME Total Hrs</div></div>
              <div class="metric"><div class="metric-v">{format_hours(header.get('me_this_month'))}</div><div class="metric-l">ME This Month</div></div>
              <div class="metric"><div class="metric-v">{n_od}</div><div class="metric-l">Overdue</div></div>
              <div class="metric"><div class="metric-v">{n_hp}</div><div class="metric-l">High Priority</div></div>
            </div>
            """, unsafe_allow_html=True)

            render_warning_summary(warnings)

            tab1, tab2, tab3, tab4 = st.tabs([
                f"Main Engine ({len(me_data)})",
                f"Aux Engines ({len(aux_data)})",
                f"D/G Equipment ({len(dg_data)})",
                f"Other Equipment ({len(oe_data)})"
            ])

            with tab1:
                if not me_data:
                    st.info("No Main Engine records found.")
                else:
                    df_me = pd.DataFrame(me_data)
                    df_me = apply_report_sort(df_me, ME_ORDER)
                    df_me["Status"] = df_me["Status"].apply(render_status_chip)
                    df_me["Used %"] = df_me["Used %"].apply(format_percent)
                    render_html_table(
                        df_me,
                        ["Status", "Component", "Engine", "Unit", "Periodicity", "Last O/H", "Hrs Since Display", "Used %"],
                        numeric_cols=["Hrs Since Display", "Used %"],
                        center_cols=["Engine", "Unit"]
                    )

            with tab2:
                if not aux_data:
                    st.info("No Auxiliary Engine records found.")
                else:
                    df_aux = pd.DataFrame(aux_data)
                    df_aux = apply_report_sort(df_aux, AUX_ORDER)
                    df_aux["Status"] = df_aux["Status"].apply(render_status_chip)
                    df_aux["Used %"] = df_aux["Used %"].apply(format_percent)
                    render_html_table(
                        df_aux,
                        ["Status", "Component", "Engine", "Unit", "Periodicity", "Last O/H", "Hrs Since Display", "Used %"],
                        numeric_cols=["Hrs Since Display", "Used %"],
                        center_cols=["Engine", "Unit"]
                    )

            with tab3:
                if not dg_data:
                    st.info("No D/G equipment records found.")
                else:
                    df_dg = pd.DataFrame(dg_data)
                    df_dg = apply_report_sort(df_dg, DG_ORDER)
                    df_dg["Status"] = df_dg["Status"].apply(render_status_chip)
                    df_dg["Used %"] = df_dg["Used %"].apply(format_percent)
                    render_html_table(
                        df_dg,
                        ["Status", "Component", "Engine", "Unit", "Periodicity", "Last O/H", "Hrs Since Display", "Used %"],
                        numeric_cols=["Hrs Since Display", "Used %"],
                        center_cols=["Engine", "Unit"]
                    )

            with tab4:
                if not oe_data:
                    st.info("No Other Equipment records found.")
                else:
                    df_oe = pd.DataFrame(oe_data).sort_values(by=["Description", "Last Date"], na_position="last")
                    render_html_table(
                        df_oe,
                        ["Description", "Periodicity", "Last Date", "Run Hrs Display"],
                        numeric_cols=["Run Hrs Display"]
                    )

        except Exception as e:
            st.error(f"Execution Failed: {e}")
