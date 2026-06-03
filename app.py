import os
import re
import shutil
import tempfile
import subprocess
from pathlib import Path
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
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
  --bg:#06111a;
  --bg2:#0d1825;
  --bg3:#132234;
  --line:#203246;
  --gold:#c99818;
  --text:#edf4ff;
  --muted:#9fb4c8;
  --soft:#6f869a;
  --ok:#2e8b57;
  --warn:#c27a00;
  --bad:#b94034;
}

html, body, [class*="css"]{
  background:var(--bg)!important;
  color:var(--muted)!important;
  font-family:'Inter', sans-serif!important;
}

.main, .block-container{background:var(--bg)!important;}
[data-testid="stSidebar"], [data-testid="collapsedControl"]{display:none!important;}

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
  margin-top:.2rem;
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
  border-radius:12px;
  padding:1rem;
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
  gap:.35rem;
  border-bottom:1px solid var(--line);
  padding:0 0 .5rem 0;
}

.stTabs [data-baseweb="tab"]{
  background:var(--bg2);
  border:1px solid var(--line);
  border-radius:10px;
  color:var(--muted);
  padding:.55rem .95rem;
  font-weight:600;
}

.stTabs [aria-selected="true"]{
  background:#132437!important;
  border-color:#2b4761!important;
  color:var(--text)!important;
}

.stTabs [data-baseweb="tab-panel"]{
  padding-top:1rem;
}

[data-testid="stFileUploadDropzone"]{
  background:rgba(201,152,24,.04)!important;
  border:1.25px dashed rgba(201,152,24,.32)!important;
  border-radius:12px!important;
}

.banner{
  border:1px solid var(--line);
  background:var(--bg2);
  border-radius:12px;
  padding:.9rem 1rem;
  margin:.7rem 0 1rem 0;
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
    "M/E L.O.", "JACKET FW", "JACKET FW NO.1", "PISTON L.O.", "ATMOSPHERIC CONDENSER",
    "AIR COND. COMPRESSOR NO.1", "AIR COND. COMPRESSOR NO.2",
    "AIR. COND. COOLER CLEANING", "REFRIGERATION COMPRESSOR NO.1",
    "REFRIGERATION COMPRESSOR NO.2", "FURNACE INSPECTION", "BURNER ATOMIZER",
    "FORCED DRAFT FAN", "FEED PUMPS NO.1", "FEED PUMPS NO.2",
    "WASHING THE TUBES", "O/H CIRC. PUMP NO.1", "O/H CIRC. PUMP NO.2",
    "STARTING MAIN AIR COMPRESSOR NO.1", "STARTING MAIN AIR COMPRESSOR NO.2",
    "SERVICE AIR COMPRESSOR", "EMERGENCY AIR COMPRESSOR NO."
}

ALIASES = {
    "PERIODICTLY": "PERIODICITY",
    "PLUNGER AND BARREL (RENEWAL)": "PLUNGER AND BARREL(RENEWAL)",
    "PLUNGER AND BARREL  (RENEWAL)": "PLUNGER AND BARREL(RENEWAL)",
    "FUEL VALVES(1)": "FUEL VALVES (1)",
    "TURBOCHARGER(2)": "TURBOCHARGER (2)",
    "TURBOCHARGER(3)": "TURBOCHARGER (3)",
    "EXAUST VALVE": "EXHAUST VALVE",
    "JACKET FW NO.1": "JACKET FW",
    "JACKET FW": "JACKET FW",
    "PERIODICITY": "PERIODICITY",
}

KNOWN_TEXT_STATES = {
    "N/A", "NO RECORD", "NOT WORKING", "CENTRAL", "COOLER"
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
    if len(raw) > 30:
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

def get_status(hrs: Optional[float], period: Optional[float]) -> str:
    if hrs is None or period is None or period <= 0:
        return "NO DATA"
    ratio = hrs / period
    if ratio >= 1.0:
        return "OVERDUE"
    if ratio >= 0.8:
        return "HIGH PRIORITY"
    return "OK"

def add_warning(warnings: List[WarningItem], section: str, severity: str, message: str, source: str = ""):
    warnings.append(WarningItem(section=section, severity=severity, message=message, source=source))

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
    from docx import Document

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
    if "MAIN ENGINE" in blob and "DATE OF LAST O/H" in blob:
        return "ME"
    if "AUX. ENGINE MAKER / TYPE" in blob:
        return "AUX"
    if "D/G NO1" in blob:
        return "DG"
    if "TURBOCHARGER" in blob or "MAIN AIR COMPRESSORS" in blob or "A/C & REFR. COMPRESSORS" in blob:
        return "OE"
    return "OTHER"

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
    i = 0

    while i < len(table_rows) - 1:
        row1 = table_rows[i]
        row2 = table_rows[i + 1]

        comp = normalize_token(row1[0]) if row1 else ""
        marker_1 = normalize_token(row1[2] if len(row1) > 2 else "")
        marker_2 = normalize_token(row2[2] if len(row2) > 2 else "")

        if comp in ME_COMPONENTS and marker_1 == "1" and marker_2 == "2":
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
                        "Hrs Since": hrs if hrs is not None else None,
                        "Hrs Since Display": format_hours(hrs),
                        "Used Ratio": ratio if ratio is not None else 0.0,
                        "Used %": round(ratio * 100, 1) if ratio is not None else None,
                    })
            i += 2
            continue

        i += 1

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
        if len(normed) >= 4 and normed[0] == "DESCRIPTION" and "1" in normed and "2" in normed:
            start = idx + 1
            continue
        if len(normed) >= 4 and normed[0] == "DESCRIPTION" and any("D/G NO1" in c for c in normed):
            stop = idx
            break

    if start is None:
        add_warning(warnings, "Aux Engine", "error", "Aux description block not found")
        return records, meta

    end_idx = stop if stop is not None else len(table_rows)
    i = start

    while i < end_idx - 1:
        row1 = table_rows[i]
        row2 = table_rows[i + 1]

        comp = normalize_token(row1[0]) if row1 else ""
        marker_1 = normalize_token(row1[2] if len(row1) > 2 else "")
        marker_2 = normalize_token(row2[2] if len(row2) > 2 else "")

        if comp in AUX_COMPONENTS and marker_1 == "1" and marker_2 == "2":
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
                "Hrs Since": hrs if hrs is not None else None,
                "Hrs Since Display": format_hours(hrs),
                "Used Ratio": ratio if ratio is not None else 0.0,
                "Used %": round(ratio * 100, 1) if ratio is not None else None,
            })
            i += 2
            continue

        i += 1

    if not records:
        add_warning(warnings, "Aux Engine", "warning", "No auxiliary engine component rows extracted")

    return records, meta

def extract_dg(table_rows: List[List[str]], warnings: List[WarningItem]) -> List[Dict[str, Any]]:
    records = []
    start = None

    for idx, row in enumerate(table_rows):
        normed = [normalize_token(c) for c in row]
        if len(normed) >= 4 and normed[0] == "DESCRIPTION" and any("D/G NO1" in c for c in normed):
            start = idx + 1
            break

    if start is None:
        return records

    i = start
    while i < len(table_rows) - 1:
        row1 = table_rows[i]
        row2 = table_rows[i + 1]

        comp = normalize_token(row1[0]) if row1 else ""
        marker_1 = normalize_token(row1[2] if len(row1) > 2 else "")
        marker_2 = normalize_token(row2[2] if len(row2) > 2 else "")

        if comp in DG_COMPONENTS and marker_1 == "1" and marker_2 == "2":
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
                        "Hrs Since": hrs if hrs is not None else None,
                        "Hrs Since Display": format_hours(hrs),
                        "Used Ratio": ratio if ratio is not None else 0.0,
                        "Used %": round(ratio * 100, 1) if ratio is not None else None,
                    })
            i += 2
            continue

        i += 1

    return records

def extract_oe(table_rows: List[List[str]], warnings: List[WarningItem]) -> List[Dict[str, Any]]:
    records = []

    for row in table_rows:
        cells = [fl(c) for c in row]

        for idx, cell in enumerate(cells):
            comp = normalize_token(cell)
            if comp in OE_COMPONENTS:
                periodicity = parse_num(cells[idx + 1]) if idx + 1 < len(cells) else None
                raw_date = cells[idx + 2] if idx + 2 < len(cells) else ""
                raw_hrs = cells[idx + 3] if idx + 3 < len(cells) else ""

                iso, state, bad_date = parse_date(raw_date)
                hrs = parse_num(raw_hrs)

                if bad_date:
                    add_warning(warnings, "Other Equipment", "warning", f"Invalid date for {comp}: {bad_date}", raw_date)
                if fl(raw_hrs) and hrs is None and normalize_token(raw_hrs) not in KNOWN_TEXT_STATES:
                    add_warning(warnings, "Other Equipment", "warning", f"Non-numeric hours for {comp}", raw_hrs)

                if iso or state or hrs is not None or periodicity is not None or fl(raw_date) or fl(raw_hrs):
                    records.append({
                        "Description": comp,
                        "Periodicity": int(periodicity) if periodicity and float(periodicity).is_integer() else periodicity or "—",
                        "Last Date": iso or state or (raw_date if fl(raw_date) else "—"),
                        "Run Hrs": hrs if hrs is not None else None,
                        "Run Hrs Display": format_hours(hrs),
                    })

    dedup = [dict(t) for t in {tuple(d.items()) for d in records}]
    return dedup

def build_payload(docx_bytes: bytes) -> Dict[str, Any]:
    warnings: List[WarningItem] = []
    model = doc_to_grid(docx_bytes)
    header = extract_header(model, warnings)

    me_rows: List[Dict[str, Any]] = []
    aux_rows: List[Dict[str, Any]] = []
    dg_rows: List[Dict[str, Any]] = []
    oe_rows: List[Dict[str, Any]] = []
    aux_meta = {"aux_total_hours": None, "aux_this_month": None}

    for table in model["tables"]:
        kind = classify_table(table["rows"])
        if kind == "ME":
            me_rows.extend(extract_me(table["rows"], warnings))
        elif kind == "AUX":
            rows, meta = extract_aux(table["rows"], warnings)
            aux_rows.extend(rows)
            for k, v in meta.items():
                if v is not None:
                    aux_meta[k] = v
            dg_rows.extend(extract_dg(table["rows"], warnings))
        elif kind == "DG":
            dg_rows.extend(extract_dg(table["rows"], warnings))
        elif kind == "OE":
            oe_rows.extend(extract_oe(table["rows"], warnings))

    err_count = sum(1 for w in warnings if w.severity == "error")
    warn_count = sum(1 for w in warnings if w.severity == "warning")
    quality_score = max(0, 100 - err_count * 18 - warn_count * 4)

    return {
        "header": header,
        "aux_meta": aux_meta,
        "me_rows": me_rows,
        "aux_rows": aux_rows,
        "dg_rows": dg_rows,
        "oe_rows": oe_rows,
        "warnings": warnings,
        "quality_score": quality_score,
    }

def render_quality_banner(payload: Dict[str, Any]):
    warnings = payload["warnings"]
    score = payload["quality_score"]
    critical = [w for w in warnings if w.severity == "error"]
    advisory = [w for w in warnings if w.severity == "warning"]

    if critical:
        st.markdown(
            f'<div class="banner banner-bad"><span class="kicker">Integrity status:</span> '
            f'critical extraction issues detected. Quality score: {score}/100.</div>',
            unsafe_allow_html=True,
        )
    elif advisory:
        st.markdown(
            f'<div class="banner banner-warn"><span class="kicker">Integrity status:</span> '
            f'extraction completed with advisory exceptions. Quality score: {score}/100.</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f'<div class="banner banner-ok"><span class="kicker">Integrity status:</span> '
            f'no extraction anomalies detected. Quality score: {score}/100.</div>',
            unsafe_allow_html=True,
        )

def style_dataframe(df: pd.DataFrame):
    def highlight(row):
        status = str(row.get("Status", ""))
        if status == "OVERDUE":
            return ["background-color: rgba(185,64,52,.18); color:#edf4ff;"] * len(row)
        if status == "HIGH PRIORITY":
            return ["background-color: rgba(194,122,0,.16); color:#edf4ff;"] * len(row)
        return [""] * len(row)
    return df.style.apply(highlight, axis=1)

OPER_COL_CONFIG = {
    "Periodicity": st.column_config.TextColumn("Limit"),
    "Last O/H": st.column_config.TextColumn("Last O/H"),
    "Hrs Since Display": st.column_config.TextColumn("Current Hrs"),
    "Used Ratio": st.column_config.ProgressColumn("Wear Level", format="%.1f%%", min_value=0, max_value=1),
    "Used %": st.column_config.NumberColumn("Used %", format="%.1f"),
}

st.markdown("""
<div class="hero-k">Running Hours Management System</div>
<div class="hero-h">TEC-004 Fleet Command Console</div>
<div class="hero-rule"></div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("Upload TEC04 / TEC-004 Report (.doc or .docx)", type=["doc", "docx"])

if uploaded:
    with st.spinner("Extracting and validating telemetry..."):
        try:
            raw = uploaded.read()
            docx_data = raw if uploaded.name.lower().endswith(".docx") else convert_doc_to_docx(raw)
            payload = build_payload(docx_data)

            header = payload["header"]
            aux_meta = payload["aux_meta"]
            me_data = payload["me_rows"]
            aux_data = payload["aux_rows"]
            dg_data = payload["dg_rows"]
            oe_data = payload["oe_rows"]

            operational_rows = me_data + aux_data + dg_data
            n_overdue = sum(1 for r in operational_rows if r.get("Status") == "OVERDUE")
            n_priority = sum(1 for r in operational_rows if r.get("Status") == "HIGH PRIORITY")
            n_notes = len(payload["warnings"])

            st.markdown(f"""
            <div class="metric-grid">
              <div class="metric"><div class="metric-v">{header['vessel']}</div><div class="metric-l">Vessel</div></div>
              <div class="metric"><div class="metric-v">{header['report_date'] or '—'}</div><div class="metric-l">Report Date</div></div>
              <div class="metric"><div class="metric-v">{format_hours(header['me_total_hours'])}</div><div class="metric-l">ME Total Hrs</div></div>
              <div class="metric"><div class="metric-v">{format_hours(header['me_this_month'])}</div><div class="metric-l">ME This Month</div></div>
              <div class="metric"><div class="metric-v">{format_hours(aux_meta['aux_total_hours'])}</div><div class="metric-l">Aux Total Hrs</div></div>
              <div class="metric"><div class="metric-v">{n_overdue}</div><div class="metric-l">Overdue Items</div></div>
            </div>
            """, unsafe_allow_html=True)

            render_quality_banner(payload)

            if n_overdue > 0:
                st.error(f"{n_overdue} components are overdue and require immediate maintenance attention.")
            elif n_priority > 0:
                st.warning(f"{n_priority} components are approaching maintenance threshold.")
            else:
                st.success("No overdue components detected in the extracted operational dataset.")

            tabs = st.tabs([
                f"Main Engine ({len(me_data)})",
                f"Aux Engine ({len(aux_data)})",
                f"D-G Equipment ({len(dg_data)})",
                f"Other Equipment ({len(oe_data)})",
            ])

            with tabs[0]:
                if not me_data:
                    st.info("No Main Engine records found.")
                else:
                    df_me = pd.DataFrame(me_data)
                    df_me["_cyl"] = df_me["Unit"].str.extract(r"(\d+)").astype(float)
                    df_me = df_me.sort_values(by=["Component", "_cyl"]).drop(columns=["_cyl", "Hrs Since"])
                    st.dataframe(style_dataframe(df_me), use_container_width=True, hide_index=True, column_config=OPER_COL_CONFIG)

            with tabs[1]:
                if not aux_data:
                    st.info("No Auxiliary Engine records found.")
                else:
                    df_aux = pd.DataFrame(aux_data).drop(columns=["Hrs Since"])
                    st.dataframe(style_dataframe(df_aux), use_container_width=True, hide_index=True, column_config=OPER_COL_CONFIG)

            with tabs[2]:
                if not dg_data:
                    st.info("No D-G Equipment records found.")
                else:
                    df_dg = pd.DataFrame(dg_data)
                    sub1, sub2, sub3 = st.tabs(["Generator 1", "Generator 2", "Generator 3"])
                    for idx, subtab in enumerate([sub1, sub2, sub3], start=1):
                        with subtab:
                            df_sub = df_dg[df_dg["Engine"] == f"D/G {idx}"].drop(columns=["Hrs Since"])
                            if df_sub.empty:
                                st.info(f"No records found for Generator {idx}.")
                            else:
                                st.dataframe(style_dataframe(df_sub), use_container_width=True, hide_index=True, column_config=OPER_COL_CONFIG)

            with tabs[3]:
                if not oe_data:
                    st.info("No Other Equipment records found.")
                else:
                    df_oe = pd.DataFrame(oe_data).drop(columns=["Run Hrs"])
                    st.dataframe(
                        df_oe,
                        use_container_width=True,
                        hide_index=True,
                        column_config={"Run Hrs Display": st.column_config.TextColumn("Run Hrs")}
                    )

            if n_notes:
                with st.expander("Data Quality Notes", expanded=False):
                    critical = [w for w in payload["warnings"] if w.severity == "error"]
                    advisory = [w for w in payload["warnings"] if w.severity == "warning"]

                    if critical:
                        st.markdown("**Blocking issues**")
                        for w in critical:
                            st.write(f"- [{w.section}] {w.message}")

                    if advisory:
                        st.markdown("**Advisory issues**")
                        for w in advisory:
                            st.write(f"- [{w.section}] {w.message}")

            st.caption("This console is optimized for operational readability while preserving extraction integrity.")

        except Exception as e:
            st.error(f"Execution failed: {e}")
            st.caption("The application stops on structural/runtime failure instead of presenting unverified telemetry.")
else:
    st.markdown(
        '<div class="banner"><span class="kicker">Awaiting input.</span> '
        'Upload a TEC04 / TEC-004 running-hours report in .doc or .docx format.</div>',
        unsafe_allow_html=True,
    )
