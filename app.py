import os
import re
import shutil
import tempfile
import subprocess
from pathlib import Path
from dataclasses import dataclass, asdict
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
:root {
  --bg: #071019;
  --bg2: #0c1623;
  --line: #1b2d44;
  --gold: #c99818;
  --t0: #ebf3ff;
  --t1: #a9bdd4;
  --ok: #1f9d55;
  --warn: #c27a00;
  --bad: #c0392b;
  --info: #2d7ff9;
}
html, body, [class*="css"] {
  background: var(--bg)!important;
  color: var(--t1)!important;
  font-family: 'Inter', sans-serif!important;
}
.main, .block-container { background: var(--bg)!important; }
[data-testid="stSidebar"], [data-testid="collapsedControl"] { display:none!important; }
.hero-k {
  font-size:.66rem;
  letter-spacing:.24em;
  text-transform:uppercase;
  color:var(--gold);
  font-weight:700;
}
.hero-h {
  font-family:'Space Grotesk', sans-serif;
  font-size:1.8rem;
  font-weight:700;
  color:var(--t0);
  line-height:1.1;
  margin-top:.2rem;
}
.hero-rule {
  height:1px;
  margin:1rem 0;
  background:linear-gradient(90deg,var(--gold),var(--line),transparent);
}
.metric-grid {
  display:grid;
  grid-template-columns:repeat(5,1fr);
  gap:1rem;
  margin:1.25rem 0 1rem;
}
.metric {
  background:var(--bg2);
  border:1px solid var(--line);
  border-radius:12px;
  padding:1rem;
  border-top:2px solid var(--gold);
}
.metric-v {
  font-family:'Space Grotesk', sans-serif;
  font-size:1.35rem;
  font-weight:700;
  color:var(--t0);
}
.metric-l {
  font-size:.64rem;
  text-transform:uppercase;
  letter-spacing:.14em;
  color:#71879f;
  margin-top:5px;
}
.panel {
  background:var(--bg2);
  border:1px solid var(--line);
  border-radius:12px;
  padding:1rem;
  margin:.75rem 0;
}
.badge-ok,.badge-warn,.badge-bad,.badge-info {
  display:inline-block;
  padding:.25rem .55rem;
  border-radius:999px;
  font-size:.75rem;
  font-weight:600;
  margin-right:.4rem;
}
.badge-ok { background:rgba(31,157,85,.16); color:#87d7a8; }
.badge-warn { background:rgba(194,122,0,.18); color:#f2c46d; }
.badge-bad { background:rgba(192,57,43,.18); color:#f2a6a0; }
.badge-info { background:rgba(45,127,249,.18); color:#9ec4ff; }
[data-testid="stFileUploadDropzone"] {
  background:rgba(201,152,24,.05)!important;
  border:1.5px dashed rgba(201,152,24,.4)!important;
  border-radius:12px!important;
}
.small { color:#93a8bf; font-size:.88rem; }
</style>
""", unsafe_allow_html=True)

ME_COMPONENTS = {
    "CYLINDER COVER",
    "PISTON ASSEMBLY",
    "STUFFING BOX",
    "PISTON CROWN",
    "CYLINDER LINER",
    "EXHAUST VALVE",
    "STARTING VALVE",
    "SAFETY VALVE",
    "FUEL VALVES",
    "FUEL PUMP",
    "PLUNGER AND BARREL(RENEWAL)",
    "PLUNGER AND BARREL",
    "FUEL PUMP SUCTION VALVE",
    "FUEL PUMP PUNCTURE VALVE",
    "CROSSHEAD BEARINGS",
    "BOTTOM END BEARINGS",
    "MAIN BEARINGS",
}

AUX_COMPONENTS = {
    "CYLINDER HEAD",
    "PISTON",
    "CONNECTING ROD",
    "CYLINDER LINERS",
    "FUEL VALVES (1)",
    "FUEL PUMPS",
    "CRANK PIN BEARING",
    "MAIN BEARING",
    "ADJUST VALVE HEAD CLEARANCE",
}

OE_COMPONENTS = {
    "TURBOCHARGER (2)",
    "TURBOCHARGER (3)",
    "AIR COOLER",
    "L.O. COOLER CLEAN",
    "F.W. COOLER CLEAN",
    "COOL WATER THERMOSTAT VALVE",
    "L.O. RENEWAL",
    "L.O. THERMOSTAT VALVE",
    "ALTERNATOR CLEANING",
    "THRUST BEARING",
    "TURBOCHARGER",
    "COOLERS",
    "A/C & REFR. COMPRESSORS",
    "GENERAL O/H",
    "M/E L.O.",
    "AIR COND. COMPRESSOR NO.1",
    "BALANCING OF ROTOR SHAFT",
    "JACKET FW NO.1",
    "AIR COND. COMPRESSOR NO.2",
    "AIR COOLER CLEANING",
    "PISTON L.O.",
    "AIR. COND. COOLER CLEANING",
    "ATMOSPHERIC CONDENSER",
    "REFRIGERATION COMPRESSOR NO.1",
    "REFRIGERATION COMPRESSOR NO.2",
    "AUXILIARY BOILER",
    "EXH GAS BOILER",
    "MAIN AIR COMPRESSORS",
    "FURNACE INSPECTION",
    "WASHING THE TUBES",
    "STARTING MAIN AIR COMPRESSOR NO.1",
    "BURNER ATOMIZER",
    "O/H CIRC. PUMP NO.1",
    "STARTING MAIN AIR COMPRESSOR NO.2",
    "FORCED DRAFT FAN",
    "O/H CIRC. PUMP NO.2",
    "SERVICE AIR COMPRESSOR",
    "FEED PUMPS NO.1",
    "EMERGENCY AIR COMPRESSOR NO.",
    "FEED PUMPS NO.2",
    "COOLING WATER PUMP",
}

ALIASES = {
    "PERIODICTLY": "PERIODICITY",
    "PLUNGER AND BARREL (RENEWAL)": "PLUNGER AND BARREL(RENEWAL)",
    "COOL WATER THERMOSTAT VALVE": "COOL WATER THERMOSTAT VALVE",
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
    return ALIASES.get(s, s)


def parse_num(txt: Any) -> Optional[int]:
    s = fl(txt).upper().replace("[", "").replace("]", "").strip()
    if not s or s in {"N/A", "-"}:
        return None
    if any(w in s for w in ["MONTH", "YEAR", "DAY", "OBSERVATION", "CENTRAL", "COOLER"]):
        return None
    m = re.search(r"\d[\d.,]*", s)
    if not m:
        return None
    token = m.group()
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", token) or re.fullmatch(r"\d{1,3}(,\d{3})+", token):
        token = token.replace(".", "").replace(",", "")
    elif "." in token and "," in token:
        token = token.replace(",", "")
    try:
        return int(float(token))
    except Exception:
        return None


def parse_date(txt: Any) -> Tuple[Optional[str], Optional[str]]:
    s = fl(txt).replace("[", "").replace("]", "").strip()
    if not s or s in {"-", "N/A", "1", "2"} or re.fullmatch(r"\d+", s):
        return None, None
    try:
        dt = dt_parser.parse(s, dayfirst=True, fuzzy=False)
        return dt.date().isoformat(), None
    except Exception:
        return None, s


def get_status(hrs: Optional[int], period: Optional[int]) -> str:
    if not hrs or not period or period <= 0:
        return "🔵 NO DATA"
    ratio = hrs / period
    if ratio >= 1.0:
        return "🔴 OVERDUE"
    if ratio >= 0.8:
        return "🟠 HIGH PRIORITY"
    return "🟢 OK"


def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError(
            "LibreOffice not found in runtime environment. "
            "For .doc support install soffice, or upload .docx files only."
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
        tp = t.name

    try:
        doc = Document(tp)
        paragraphs = [fl(p.text) for p in doc.paragraphs if fl(p.text)]
        tables = []

        for ti, table in enumerate(doc.tables):
            rows = []
            for row in table.rows:
                row_vals = [fl(c.text) for c in row.cells]
                if any(row_vals):
                    rows.append(row_vals)
            if rows:
                tables.append({"table_index": ti, "rows": rows})

        return {"paragraphs": paragraphs, "tables": tables}
    finally:
        try:
            os.unlink(tp)
        except Exception:
            pass


def classify_table(rows: List[List[str]]) -> str:
    blob = " ".join(" ".join(r) for r in rows).upper()

    if "MAIN ENGINE" in blob and "TOTAL RUNNING HOURS" in blob:
        return "ME"
    if "AUX. ENGINE MAKER / TYPE" in blob or ("HOURS THIS MONTH" in blob and "SERIAL NR" in blob):
        return "AUX"
    if "D/G NO1" in blob or "TURBOCHARGER" in blob or "MAIN AIR COMPRESSORS" in blob:
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
    }

    m = re.search(
        r"VESSEL['’]S NAME\s*:\s*(?:MV\s+)?(.+?)\s+DATE\s*:?\s*([A-Z0-9/ .-]+)",
        text,
        re.I,
    )
    if m:
        out["vessel"] = re.sub(r"(?i)^MV\s+", "", fl(m.group(1))).strip()
        iso, bad = parse_date(m.group(2))
        out["report_date"] = iso or m.group(2).strip()
        if bad:
            warnings.append(WarningItem("Header", "warning", f"Unparsed report date: {bad}", bad))
    else:
        warnings.append(WarningItem("Header", "error", "Failed to extract vessel/date header"))

    m = re.search(r"TOTAL RUNNING HOURS\s*:?\s*([\d,\.]+)", text, re.I)
    if m:
        out["me_total_hours"] = parse_num(m.group(1))
    else:
        warnings.append(WarningItem("Header", "warning", "Main engine total running hours not found"))

    m = re.search(r"THIS MONTH\s*:?\s*([\d,\.]+)", text, re.I)
    if m:
        out["me_this_month"] = parse_num(m.group(1))
    else:
        warnings.append(WarningItem("Header", "warning", "Main engine this-month hours not found"))

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
            period_cell = row1[1] if len(row1) > 1 else ""
            period = parse_num(period_cell)

            if period is None and "OBSERVATION" not in normalize_token(period_cell):
                warnings.append(
                    WarningItem("Main Engine", "warning", f"Missing periodicity for {comp}", " | ".join(row1))
                )

            max_cols = min(len(row1), len(row2))
            cyl_count = max(0, max_cols - 3)

            for j in range(cyl_count):
                raw_date = row1[3 + j] if 3 + j < len(row1) else ""
                raw_hrs = row2[3 + j] if 3 + j < len(row2) else ""

                iso, bad_date = parse_date(raw_date)
                hrs = parse_num(raw_hrs)

                if bad_date:
                    warnings.append(
                        WarningItem(
                            "Main Engine",
                            "warning",
                            f"Invalid date for {comp} cyl {j+1}: {bad_date}",
                            raw_date,
                        )
                    )

                if iso or hrs is not None:
                    records.append({
                        "Status": get_status(hrs, period),
                        "Component": comp,
                        "Engine": "ME",
                        "Unit": f"Cyl {j+1}",
                        "Periodicity": period if period is not None else "OBS",
                        "Last O/H": iso or (raw_date if fl(raw_date) else "—"),
                        "Hrs Since": hrs if hrs is not None else 0,
                        "Used %": f"{round((hrs / period) * 100, 1)}%" if hrs is not None and period else "0.0%",
                    })
            i += 2
            continue

        i += 1

    if not records:
        warnings.append(WarningItem("Main Engine", "error", "No main engine records extracted"))

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
    for idx, row in enumerate(table_rows):
        normed = [normalize_token(c) for c in row]
        if len(normed) >= 4 and normed[0] == "DESCRIPTION" and "1" in normed and "2" in normed:
            start = idx + 1
            break

    if start is None:
        warnings.append(WarningItem("Aux Engine", "error", "Aux description table header not found"))
        return records, meta

    i = start
    while i < len(table_rows) - 1:
        row1 = table_rows[i]
        row2 = table_rows[i + 1]
        comp = normalize_token(row1[0]) if row1 else ""

        if comp in AUX_COMPONENTS:
            period = parse_num(row1[1] if len(row1) > 1 else "")
            date_cell = row1[3] if len(row1) > 3 else ""
            hrs_cell = row2[3] if len(row2) > 3 else ""

            iso, bad_date = parse_date(date_cell)
            hrs = parse_num(hrs_cell)

            if bad_date:
                warnings.append(
                    WarningItem("Aux Engine", "warning", f"Invalid date for {comp}: {bad_date}", date_cell)
                )

            records.append({
                "Status": get_status(hrs, period),
                "Component": comp,
                "Engine": "AUX-1",
                "Unit": "Engine",
                "Periodicity": period,
                "Last O/H": iso or (date_cell if fl(date_cell) else "—"),
                "Hrs Since": hrs if hrs is not None else 0,
                "Used %": f"{round((hrs / period) * 100, 1)}%" if hrs is not None and period else "0.0%",
            })
            i += 2
            continue

        if comp == "DESCRIPTION" and any("D/G NO1" in normalize_token(c) for c in row1):
            break

        i += 1

    if not records:
        warnings.append(WarningItem("Aux Engine", "error", "No auxiliary engine records extracted"))

    return records, meta


def extract_oe(table_rows: List[List[str]], warnings: List[WarningItem]) -> List[Dict[str, Any]]:
    records = []

    for row in table_rows:
        cells = [fl(c) for c in row]

        for idx, cell in enumerate(cells):
            comp = normalize_token(cell)

            if comp in OE_COMPONENTS:
                period = parse_num(cells[idx + 1]) if idx + 1 < len(cells) else None
                raw_date = cells[idx + 2] if idx + 2 < len(cells) else ""
                raw_hrs = cells[idx + 3] if idx + 3 < len(cells) else ""

                iso, bad_date = parse_date(raw_date)
                hrs = parse_num(raw_hrs)

                if bad_date:
                    warnings.append(
                        WarningItem("Other Equipment", "warning", f"Invalid date for {comp}: {bad_date}", raw_date)
                    )

                if iso or hrs is not None or period is not None:
                    records.append({
                        "Section": "Other Equipment",
                        "Description": comp,
                        "Periodicity": period,
                        "Last Date": iso or (raw_date if fl(raw_date) else "—"),
                        "Run Hrs": hrs if hrs is not None else 0,
                    })

    dedup = [dict(t) for t in {tuple(d.items()) for d in records}]

    if not dedup:
        warnings.append(WarningItem("Other Equipment", "warning", "No other-equipment records extracted"))

    return dedup


def build_payload(docx_bytes: bytes) -> Dict[str, Any]:
    warnings: List[WarningItem] = []
    model = doc_to_grid(docx_bytes)
    header = extract_header(model, warnings)

    me_rows: List[Dict[str, Any]] = []
    aux_rows: List[Dict[str, Any]] = []
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
        elif kind == "OE":
            oe_rows.extend(extract_oe(table["rows"], warnings))

    err_count = sum(1 for w in warnings if w.severity == "error")
    warn_count = sum(1 for w in warnings if w.severity == "warning")
    score = max(0, 100 - err_count * 18 - warn_count * 4)

    return {
        "header": header,
        "aux_meta": aux_meta,
        "me_rows": me_rows,
        "aux_rows": aux_rows,
        "oe_rows": oe_rows,
        "warnings": [asdict(w) for w in warnings],
        "quality_score": score,
    }


def badge_html(score: int) -> str:
    if score >= 90:
        return '<span class="badge-ok">Validated</span>'
    if score >= 70:
        return '<span class="badge-warn">Review advised</span>'
    return '<span class="badge-bad">Manual review required</span>'


st.markdown("""
<div class="hero-k">Running Hours Management System</div>
<div class="hero-h">TEC-004 Extraction Matrix · Verified Build</div>
<div class="hero-rule"></div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("Upload TEC-004 Report (.doc or .docx)", type=["doc", "docx"])

if uploaded:
    with st.spinner("Parsing TEC-004 with validation and confidence scoring..."):
        try:
            raw = uploaded.read()
            docx_data = raw if uploaded.name.lower().endswith(".docx") else convert_doc_to_docx(raw)
            payload = build_payload(docx_data)

            header = payload["header"]
            me_data = payload["me_rows"]
            aux_data = payload["aux_rows"]
            oe_data = payload["oe_rows"]
            warnings = payload["warnings"]
            score = payload["quality_score"]

            n_od = sum(1 for r in me_data + aux_data if "OVERDUE" in r["Status"])
            n_hp = sum(1 for r in me_data + aux_data if "HIGH PRIORITY" in r["Status"])

            st.markdown(f"""
            <div class="metric-grid">
              <div class="metric"><div class="metric-v">{header['vessel']}</div><div class="metric-l">Vessel</div></div>
              <div class="metric"><div class="metric-v">{header['report_date'] or '—'}</div><div class="metric-l">Report Date</div></div>
              <div class="metric"><div class="metric-v">{header['me_total_hours'] or 0:,}</div><div class="metric-l">ME Total Hrs</div></div>
              <div class="metric"><div class="metric-v">{header['me_this_month'] or 0:,}</div><div class="metric-l">ME This Month</div></div>
              <div class="metric"><div class="metric-v">{score}/100</div><div class="metric-l">Extraction Quality</div></div>
            </div>
            """, unsafe_allow_html=True)

            st.markdown(badge_html(score), unsafe_allow_html=True)
            st.markdown(
                f'<span class="badge-bad">Overdue: {n_od}</span>'
                f'<span class="badge-warn">High priority: {n_hp}</span>'
                f'<span class="badge-info">Warnings: {len(warnings)}</span>',
                unsafe_allow_html=True
            )

            review, tab1, tab2, tab3, rawtab = st.tabs([
                f"🧭 Review ({len(warnings)})",
                f"⚙ Main Engine ({len(me_data)})",
                f"🔩 Aux Engines ({len(aux_data)})",
                f"🛠 Other Equipment ({len(oe_data)})",
                "📦 JSON",
            ])

            with review:
                if not warnings:
                    st.success("No extraction issues detected.")
                else:
                    df_warn = pd.DataFrame(warnings)
                    st.dataframe(df_warn, use_container_width=True, hide_index=True)
                    st.caption("Every warning is meant to stop silent bad data from reaching the engineer.")

            with tab1:
                if not me_data:
                    st.info("No Main Engine records found.")
                else:
                    df_me = pd.DataFrame(me_data)
                    df_me["_cyl"] = df_me["Unit"].str.extract(r"(\d+)").astype(float)
                    df_me = df_me.sort_values(by=["Component", "_cyl"]).drop(columns=["_cyl"])
                    st.dataframe(df_me.astype(str), use_container_width=True, hide_index=True)

            with tab2:
                if not aux_data:
                    st.info("No Auxiliary Engine records found.")
                else:
                    df_aux = pd.DataFrame(aux_data)
                    st.dataframe(df_aux.astype(str), use_container_width=True, hide_index=True)

            with tab3:
                if not oe_data:
                    st.info("No Other Equipment records found.")
                else:
                    df_oe = pd.DataFrame(oe_data)
                    st.dataframe(df_oe.astype(str), use_container_width=True, hide_index=True)

            with rawtab:
                st.json(payload)

        except Exception as e:
            st.error(f"Execution Failed: {e}")
            st.caption("The app now fails loudly instead of hiding structural parsing problems.")
else:
    st.markdown(
        '<div class="panel"><div class="small">'
        'Upload a TEC-004 monthly running-hours report in .doc or .docx format. '
        'This build prioritizes structural parsing, validation, and review visibility '
        'over silent best-effort extraction.'
        '</div></div>',
        unsafe_allow_html=True
    )
