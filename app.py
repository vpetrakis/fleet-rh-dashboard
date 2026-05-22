import streamlit as st
st.set_page_config(
    page_title="Fleet Running Hours Command",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

import os
import re
import json
import shutil
import sqlite3
import hashlib
import tempfile
import subprocess
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Tuple

import pandas as pd

# ============================================================
# THEME
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Manrope:wght@500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

:root{
  --bg:#060b12;
  --bg2:#0b111a;
  --bg3:#0f1722;
  --bg4:#131d2a;
  --line:#1f2d3f;
  --line2:#2c4058;
  --text:#e6eef7;
  --muted:#9cb0c7;
  --soft:#627a92;
  --gold:#b88a1b;
  --gold2:#ddb23b;
  --gold3:#f0c85d;
  --red:#ef5350;
  --amber:#ffb74d;
  --green:#66bb6a;
  --blue:#64b5f6;
  --cyan:#4dd0e1;
  --radius:14px;
  --shadow:0 18px 60px rgba(0,0,0,.42);
}
html, body, [class*="css"] {
  background: var(--bg) !important;
  color: var(--text) !important;
  font-family: 'Inter', sans-serif !important;
}
.main, .main > div, .block-container { background: var(--bg) !important; }
.block-container {
  max-width: 100% !important;
  padding: 1.2rem 1.8rem 3rem !important;
}
[data-testid="collapsedControl"], [data-testid="stSidebar"] { display:none !important; }

.main::before{
  content:"";
  position:fixed;
  inset:0;
  pointer-events:none;
  z-index:0;
  background:
    radial-gradient(ellipse 70% 50% at 10% 0%, rgba(221,178,59,.08), transparent 60%),
    radial-gradient(ellipse 70% 55% at 100% 100%, rgba(77,208,225,.05), transparent 55%);
}
.block-container > * { position: relative; z-index: 1; }

.hero {
  display:flex;
  align-items:flex-end;
  justify-content:space-between;
  gap:1rem;
  margin-bottom:1rem;
}
.hero-kicker{
  font-size:.7rem;
  letter-spacing:.22em;
  text-transform:uppercase;
  color:var(--gold3);
  margin-bottom:.35rem;
  font-weight:700;
}
.hero-title{
  font-family:'Manrope',sans-serif;
  font-size:2rem;
  font-weight:800;
  line-height:1.05;
  letter-spacing:-.04em;
  color:var(--text);
}
.hero-sub{
  color:var(--muted);
  font-size:.95rem;
  max-width:900px;
  margin-top:.45rem;
  line-height:1.6;
}
.hero-rule{
  height:1px;
  background:linear-gradient(90deg,var(--gold2) 0%, var(--line) 35%, transparent 100%);
  margin: .85rem 0 1.35rem;
}

.upload-panel{
  background:
    linear-gradient(180deg, rgba(221,178,59,.06), rgba(100,181,246,.03)),
    linear-gradient(180deg, rgba(19,29,42,.96), rgba(11,17,26,.96));
  border:1px solid rgba(221,178,59,.28);
  border-radius:var(--radius);
  padding:1rem 1rem 1rem;
  box-shadow:var(--shadow);
}
[data-testid="stFileUploadDropzone"]{
  background:rgba(221,178,59,.04) !important;
  border:1.5px dashed rgba(221,178,59,.55) !important;
  border-radius:16px !important;
  padding:2.1rem 1.4rem !important;
}
[data-testid="stFileUploadDropzone"]:hover{
  border-color:var(--gold3) !important;
  background:rgba(221,178,59,.07) !important;
}

.kpi-grid{
  display:grid;
  grid-template-columns:repeat(6,1fr);
  gap:.8rem;
  margin-bottom:1.2rem;
}
.kpi{
  background:linear-gradient(180deg, rgba(15,23,34,.95), rgba(10,16,24,.95));
  border:1px solid var(--line);
  border-radius:var(--radius);
  padding:.9rem 1rem 1rem;
  position:relative;
  overflow:hidden;
  box-shadow:var(--shadow);
}
.kpi::before{
  content:"";
  position:absolute;
  left:0; right:0; top:0;
  height:2px;
  background:linear-gradient(90deg, var(--gold2), transparent 75%);
}
.kpi-v{
  font-family:'Manrope',sans-serif;
  font-size:1.55rem;
  font-weight:800;
  line-height:1.1;
  letter-spacing:-.04em;
  color:var(--text);
}
.kpi-l{
  color:var(--soft);
  font-size:.62rem;
  text-transform:uppercase;
  letter-spacing:.18em;
  margin-top:.38rem;
}
.kpi.g::before{ background:linear-gradient(90deg, var(--green), transparent 75%); }
.kpi.r::before{ background:linear-gradient(90deg, var(--red), transparent 75%); }
.kpi.a::before{ background:linear-gradient(90deg, var(--amber), transparent 75%); }
.kpi.b::before{ background:linear-gradient(90deg, var(--blue), transparent 75%); }

.badge{
  display:inline-flex;
  align-items:center;
  gap:.35rem;
  border-radius:999px;
  padding:.34rem .62rem;
  font-size:.72rem;
  font-weight:700;
  border:1px solid var(--line2);
  background:rgba(15,23,34,.9);
}
.badge.red{ color:#ffd8d8; border-color:rgba(239,83,80,.35); background:rgba(239,83,80,.11);}
.badge.amber{ color:#ffe7c5; border-color:rgba(255,183,77,.35); background:rgba(255,183,77,.10);}
.badge.green{ color:#daf6dc; border-color:rgba(102,187,106,.35); background:rgba(102,187,106,.10);}
.badge.blue{ color:#d7ebff; border-color:rgba(100,181,246,.35); background:rgba(100,181,246,.10);}
.badge.cyan{ color:#d7fbff; border-color:rgba(77,208,225,.35); background:rgba(77,208,225,.08);}

.banner{
  border-radius:var(--radius);
  padding:.9rem 1rem;
  border:1px solid var(--line);
  margin-bottom:.95rem;
  line-height:1.5;
  box-shadow:var(--shadow);
}
.banner.warn{
  border-color:rgba(255,183,77,.35);
  background:rgba(255,183,77,.09);
  color:#ffe8c9;
}
.banner.err{
  border-color:rgba(239,83,80,.35);
  background:rgba(239,83,80,.10);
  color:#ffd8d8;
}
.banner.ok{
  border-color:rgba(102,187,106,.35);
  background:rgba(102,187,106,.10);
  color:#daf6dc;
}

[data-testid="stDataFrame"]{
  border:1px solid var(--line) !important;
  border-radius:var(--radius) !important;
  overflow:hidden !important;
  box-shadow:var(--shadow);
}
.dvn-scroller{ background:var(--bg2) !important; }

.stButton>button{
  background:linear-gradient(135deg,var(--gold3),var(--gold2)) !important;
  color:#091017 !important;
  border:none !important;
  border-radius:11px !important;
  font-weight:800 !important;
  letter-spacing:.05em !important;
  text-transform:uppercase !important;
  padding:.72rem 1.15rem !important;
}
.stDownloadButton>button{
  background:var(--bg4) !important;
  color:var(--text) !important;
  border:1px solid var(--line2) !important;
  border-radius:11px !important;
  font-weight:700 !important;
  padding:.7rem 1rem !important;
}
div[data-baseweb="select"] > div,
.stTextInput > div > div > input {
  background:var(--bg4) !important;
  color:var(--text) !important;
  border:1px solid var(--line2) !important;
  border-radius:10px !important;
}
.stCheckbox label, .stRadio label, .stSelectbox label, .stMultiSelect label {
  color:var(--soft) !important;
  font-size:.68rem !important;
  text-transform:uppercase !important;
  letter-spacing:.12em !important;
}
.streamlit-expanderHeader{
  background:var(--bg3) !important;
  border:1px solid var(--line) !important;
  border-radius:12px !important;
  color:var(--text) !important;
}
.streamlit-expanderContent{
  background:var(--bg2) !important;
  border:1px solid var(--line) !important;
  border-top:none !important;
  border-radius:0 0 12px 12px !important;
}
.muted{ color:var(--muted); font-size:.88rem; line-height:1.6; }
.micro{ font-family:'JetBrains Mono', monospace; color:var(--soft); font-size:.72rem; }

@media (max-width:1200px){
  .kpi-grid{ grid-template-columns:repeat(3,1fr); }
}
@media (max-width:800px){
  .kpi-grid{ grid-template-columns:repeat(2,1fr); }
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# CONFIG
# ============================================================
DB_PATH = "tec004_reports.db"
PARSER_VERSION = "tec004_enterprise_v2_hybrid"
STATUS_ORDER = {"OVERDUE": 0, "HIGH PRIORITY": 1, "OK": 2, "NO DATA": 3}
ENGINE_ORDER = {"ME": 0, "AUX-1": 1, "AUX-2": 2, "AUX-3": 3}
SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}

# ============================================================
# DB
# ============================================================
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS vessels (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        created_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS reports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_id INTEGER,
        vessel_name TEXT NOT NULL,
        report_date TEXT,
        filename TEXT NOT NULL,
        file_hash TEXT NOT NULL,
        parser_version TEXT NOT NULL,
        me_total_hrs REAL,
        me_this_month REAL,
        raw_json TEXT,
        created_at TEXT NOT NULL,
        UNIQUE(vessel_name, report_date, file_hash),
        FOREIGN KEY (vessel_id) REFERENCES vessels(id)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS parsed_rows (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        report_id INTEGER NOT NULL,
        category TEXT NOT NULL,
        engine_label TEXT,
        unit TEXT,
        description TEXT NOT NULL,
        periodicity_raw TEXT,
        periodicity_hours REAL,
        last_oh_date TEXT,
        hrs_since REAL,
        pct_used REAL,
        status TEXT,
        source_table_index INTEGER,
        section_order INTEGER,
        component_order INTEGER,
        engine_order INTEGER,
        unit_order INTEGER,
        source_row_start INTEGER,
        source_row_end INTEGER,
        source_col_date INTEGER,
        source_col_hours INTEGER,
        raw_date_text TEXT,
        raw_hours_text TEXT,
        normalized_date_text TEXT,
        normalized_hours_text TEXT,
        confidence REAL,
        issue_count INTEGER DEFAULT 0,
        was_repaired INTEGER DEFAULT 0,
        repair_notes TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY (report_id) REFERENCES reports(id)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS parse_issues (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        report_id INTEGER,
        row_key TEXT,
        severity TEXT,
        issue_code TEXT,
        message TEXT,
        table_index INTEGER,
        row_index INTEGER,
        created_at TEXT NOT NULL,
        FOREIGN KEY (report_id) REFERENCES reports(id)
    )
    """)

    conn.commit()
    conn.close()

init_db()

# ============================================================
# HELPERS
# ============================================================
def now_iso():
    return datetime.utcnow().isoformat()

def md5_bytes(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()

def safe_float(x):
    try:
        v = float(x)
        return 0.0 if pd.isna(v) else v
    except Exception:
        return 0.0

def safe_int(x):
    try:
        return int(float(x))
    except Exception:
        return 0

def cyl_num(unit: str) -> int:
    m = re.search(r'(\d+)', str(unit or ""))
    return int(m.group(1)) if m else 999

def normalize_whitespace(text: Any) -> str:
    return re.sub(r'\s+', ' ', str(text or '')).strip()

def first_line(text: Any) -> str:
    raw = str(text or "").replace("\x07", "").replace("\xa0", " ").replace("\t", " ")
    parts = re.split(r'[\r\n\x0b]+', raw)
    for p in parts:
        p = normalize_whitespace(p)
        if p:
            return p
    return ""

def clean_name(text: Any) -> str:
    t = first_line(text)
    t = re.sub(r'(?i)ALEXIS\s*Date?', '', t)
    t = re.sub(r'(?i)Page\s*\d+\s*of\s*\d+', '', t)
    t = re.sub(r'(?i)MTS Marine Ltd.*', '', t)
    t = re.sub(r'(?i)^MV\s+', '', t)
    return normalize_whitespace(t)

def make_issue(severity: str, code: str, message: str, table_idx: int = None, row_idx: int = None, row_key: str = "") -> Dict[str, Any]:
    return {
        "severity": severity,
        "issue_code": code,
        "message": message,
        "table_index": table_idx,
        "row_index": row_idx,
        "row_key": row_key,
    }

def is_component_name(name: str) -> bool:
    u = normalize_whitespace(str(name or "")).upper()
    if not u or len(u) < 2:
        return False
    bad = [
        "DESCRIPTION","REMARKS","COMPONENT","PERIODICITY","PERIODICTLY","DATE OF LAST",
        "RUNNING HOURS","TYPE","MAIN ENGINE","AUX. ENGINE","AUX ENGINE","TURBOCHARGER",
        "CYL. NO","TOTAL HOURS","HOURS THIS MONTH","SERIAL NR","VESSEL","TITLE",
        "COPY","CHIEF ENGINEER","NOTE 1","NOTE 2"
    ]
    if any(x in u for x in bad):
        return False
    if re.fullmatch(r'[\d\s,./:\-\[\]]+', u):
        return False
    if len(u) > 90:
        return False
    return bool(re.search(r'[A-Z]', u))

def repair_token(text: Any) -> Dict[str, Any]:
    raw = first_line(text)
    token = raw.strip()
    notes = []
    repaired = False

    if not token:
        return {"raw": raw, "clean": "", "repaired": False, "notes": []}

    new_token = token
    if "[" in new_token or "]" in new_token:
        new_token = new_token.replace("[", "").replace("]", "")
        repaired = True
        notes.append("Removed bracket artifacts")

    new_token2 = re.sub(r'\s+', ' ', new_token).strip()
    if new_token2 != new_token:
        new_token = new_token2
        repaired = True
        notes.append("Collapsed repeated whitespace")

    if re.search(r'^\d{4,6}\s+\d{2,3}$', new_token):
        left, right = new_token.split()
        new_token = left
        repaired = True
        notes.append(f"Trimmed suspicious trailing numeric fragment '{right}'")

    return {"raw": raw, "clean": normalize_whitespace(new_token), "repaired": repaired, "notes": notes}

def normalize_date_string(text: Any) -> Dict[str, Any]:
    rep = repair_token(text)
    s = rep["clean"].upper()

    if not s or s in {"N/A", "NA", "-", "CENTRAL", "COOLER"}:
        return {"raw": rep["raw"], "value": "", "repaired": rep["repaired"], "notes": rep["notes"]}

    if re.fullmatch(r'\d+', s):
        return {"raw": rep["raw"], "value": "", "repaired": rep["repaired"], "notes": rep["notes"]}

    s = s.replace("DEC.", "DEC").replace("NOV.", "NOV").replace("OCT.", "OCT").replace("JAN.", "JAN")
    s = re.sub(r'(?i)\bSEPT\b', 'SEP', s)
    s = re.sub(r'(?i)\bMARCH\b', 'MARCH', s)
    s = re.sub(r'\s+', ' ', s).strip()

    if len(s) > 24:
        return {"raw": rep["raw"], "value": "", "repaired": True, "notes": rep["notes"] + ["Rejected overlong date token"]}

    return {"raw": rep["raw"], "value": s, "repaired": rep["repaired"], "notes": rep["notes"]}

def normalize_number_string(text: Any) -> Dict[str, Any]:
    rep = repair_token(text)
    s = rep["clean"].upper()

    if not s or s in {"", "-", "N/A", "NA", "CENTRAL", "COOLER"}:
        return {"raw": rep["raw"], "value": 0.0, "normalized_text": "", "repaired": rep["repaired"], "notes": rep["notes"]}

    if any(k in s for k in ["MONTH","YEAR","WEEK","DAY","OBSERVATION","OBS"]):
        return {"raw": rep["raw"], "value": 0.0, "normalized_text": s, "repaired": rep["repaired"], "notes": rep["notes"]}

    t = re.sub(r'([,.])\s+', r'\1', s)
    m = re.search(r'\d[\d,.\s]*', t)
    if not m:
        return {"raw": rep["raw"], "value": 0.0, "normalized_text": s, "repaired": rep["repaired"], "notes": rep["notes"]}

    block = m.group().strip()
    if re.search(r'^\d{4,6}\s+\d{2,3}$', block):
        left, right = block.split()
        block = left
        rep["repaired"] = True
        rep["notes"] = rep["notes"] + [f"Removed suspicious trailing fragment '{right}' from numeric token"]

    block = block.replace(" ", "")
    sep = max(block.rfind("."), block.rfind(","))
    if sep > 0 and len(block) - sep == 4:
        block = re.sub(r'[,.]', '', block)
    elif sep > 0:
        block = re.sub(r'[,.]', '', block[:sep])
    else:
        block = re.sub(r'[,.]', '', block)

    try:
        val = float(block)
    except Exception:
        val = 0.0

    return {
        "raw": rep["raw"],
        "value": val,
        "normalized_text": block,
        "repaired": rep["repaired"],
        "notes": rep["notes"]
    }

def periodicity_value(raw: Any) -> Dict[str, Any]:
    return normalize_number_string(raw)

def compute_status(hrs: float, period: float) -> str:
    if hrs <= 0 or period <= 0:
        return "NO DATA"
    ratio = hrs / period
    if ratio >= 1.0:
        return "OVERDUE"
    if ratio >= 0.8:
        return "HIGH PRIORITY"
    return "OK"

def compute_pct(hrs: float, period: float) -> float:
    if hrs <= 0 or period <= 0:
        return 0.0
    return round(hrs / period, 4)

def build_confidence(issue_count: int, repaired: bool, date_ok: bool, hrs_ok: bool, period_ok: bool) -> float:
    score = 1.0
    if repaired: score -= 0.08
    if not date_ok: score -= 0.12
    if not hrs_ok: score -= 0.12
    if not period_ok: score -= 0.10
    score -= min(0.45, issue_count * 0.07)
    return max(0.0, round(score, 2))

def rect_grid(table) -> List[List[str]]:
    grid = []
    if not table.rows:
        return grid
    max_cols = max(len(r.cells) for r in table.rows)
    for row in table.rows:
        vals = []
        for cell in row.cells:
            raw = re.sub(r'[\x0b\r]', '\n', cell.text).replace('\x07', '')
            lines = [normalize_whitespace(x) for x in raw.split('\n') if normalize_whitespace(x)]
            vals.append(lines[0] if lines else "")
        while len(vals) < max_cols:
            vals.append("")
        grid.append(vals)
    return grid

def normalize_row_record(rec: Dict[str, Any]) -> Dict[str, Any]:
    rec = rec or {}
    out = dict(rec)
    defaults = {
        "category": "",
        "engine_label": "",
        "unit": "",
        "description": "",
        "periodicity_raw": "",
        "periodicity_hours": 0.0,
        "last_oh_date": "",
        "hrs_since": 0.0,
        "pct_used": 0.0,
        "status": "NO DATA",
        "source_table_index": "",
        "section_order": 0,
        "component_order": 0,
        "engine_order": 0,
        "unit_order": 0,
        "source_row_start": "",
        "source_row_end": "",
        "source_col_date": "",
        "source_col_hours": "",
        "raw_date_text": "",
        "raw_hours_text": "",
        "normalized_date_text": "",
        "normalized_hours_text": "",
        "confidence": 0.0,
        "issue_count": 0,
        "was_repaired": 0,
        "repair_notes": "",
    }
    for k, v in defaults.items():
        out.setdefault(k, v)
    return out

def normalize_rows_payload(rows) -> List[Dict[str, Any]]:
    if rows is None:
        return []
    if isinstance(rows, pd.DataFrame):
        rows = rows.to_dict("records")
    if not isinstance(rows, list):
        return []
    out = []
    for r in rows:
        if isinstance(r, dict):
            out.append(normalize_row_record(r))
        else:
            out.append(normalize_row_record({}))
    return out

def normalize_other_equipment(rows) -> List[Dict[str, Any]]:
    if rows is None:
        return []
    if isinstance(rows, pd.DataFrame):
        rows = rows.to_dict("records")
    if not isinstance(rows, list):
        return []
    out = []
    for r in rows:
        if not isinstance(r, dict):
            r = {}
        r.setdefault("section", "")
        r.setdefault("description", "")
        r.setdefault("last_date", "")
        r.setdefault("run_hrs", "")
        r.setdefault("source_table_index", "")
        r.setdefault("source_row", "")
        r.setdefault("section_order", 0)
        r.setdefault("item_order", 0)
        out.append(r)
    return out

def normalize_parsed_payload(parsed: Dict[str, Any]) -> Dict[str, Any]:
    parsed = parsed or {}
    parsed.setdefault("vessel_name", "UNKNOWN")
    parsed.setdefault("report_date", "")
    parsed.setdefault("me_total_hrs", 0)
    parsed.setdefault("me_this_month", 0)
    parsed.setdefault("filename", "")
    parsed.setdefault("file_hash", "")
    parsed.setdefault("uploaded_at", now_iso())
    parsed["me_comps"] = normalize_rows_payload(parsed.get("me_comps", []))
    parsed["aux_comps"] = normalize_rows_payload(parsed.get("aux_comps", []))
    parsed["components"] = normalize_rows_payload(parsed.get("components", []))
    parsed["other_equipment"] = normalize_other_equipment(parsed.get("other_equipment", []))
    parsed.setdefault("issues", [])
    if not isinstance(parsed["issues"], list):
        parsed["issues"] = []
    parsed.setdefault("debug", {})
    return parsed

def row_key(rec: Dict[str, Any]) -> str:
    return " | ".join([
        str(rec.get("category", "")),
        str(rec.get("description", "")),
        str(rec.get("engine_label", "")),
        str(rec.get("unit", "")),
        str(rec.get("source_table_index", "")),
        str(rec.get("source_row_start", "")),
        str(rec.get("source_col_date", "")),
    ])

# ============================================================
# CONVERSION
# ============================================================
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found. Add libreoffice to your environment.")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as tf:
        tf.write(raw)
        src = tf.name
    outdir = tempfile.mkdtemp(prefix="lo_")
    out = os.path.join(outdir, Path(src).stem + ".docx")
    profile = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        r = subprocess.run(
            [soffice, "--headless", "--norestore", "--nofirststartwizard", f"-env:UserInstallation={profile}",
             "--convert-to", "docx", src, "--outdir", outdir],
            capture_output=True, timeout=120
        )
        if not os.path.exists(out):
            raise RuntimeError(r.stderr.decode("utf-8", "ignore")[:400] or r.stdout.decode("utf-8", "ignore")[:400])
        with open(out, "rb") as f:
            return f.read()
    finally:
        for p in [src, out]:
            try:
                if os.path.exists(p):
                    os.unlink(p)
            except Exception:
                pass
        shutil.rmtree(outdir, ignore_errors=True)

# ============================================================
# PARSER
# ============================================================
def detect_vessel_and_date(doc) -> Tuple[str, str, List[Dict[str, Any]]]:
    vessel = "UNKNOWN"
    report_date = ""
    issues = []

    for p in doc.paragraphs:
        txt = normalize_whitespace(p.text)
        if not txt:
            continue
        mv = re.search(r"Vessel[’'`s\s]*Name\s*:?\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+)", txt, re.I)
        if mv and vessel == "UNKNOWN":
            vessel = clean_name(mv.group(1))
        md = re.search(r"Date\s*:?\s*(.+)", txt, re.I)
        if md and not report_date:
            report_date = normalize_date_string(md.group(1))["value"]
        if vessel != "UNKNOWN" and report_date:
            break

    if vessel == "UNKNOWN":
        issues.append(make_issue("warning", "VESSEL_NOT_FOUND", "Could not extract vessel name."))
    if not report_date:
        issues.append(make_issue("warning", "REPORT_DATE_NOT_FOUND", "Could not extract report date."))
    return vessel, report_date, issues

def detect_me_totals(grid: List[List[str]]) -> Tuple[float, float]:
    me_total = 0
    me_month = 0
    flat = " | ".join([" | ".join(row) for row in grid[:4]])
    m1 = re.search(r"Total Running Hours[\s:ǀ|]+([\d,]+)", flat, re.I)
    m2 = re.search(r"This Month[\s:]+([\d,]+)", flat, re.I)
    if m1:
        me_total = normalize_number_string(m1.group(1))["value"]
    if m2:
        me_month = normalize_number_string(m2.group(1))["value"]
    return me_total, me_month

def parse_me_table(grid: List[List[str]], table_idx: int, section_base: int = 1000) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, Any]]:
    records, issues = [], []
    debug = {"table_idx": table_idx, "detected": False, "rows_scanned": 0, "blocks_found": 0, "records_emitted": 0}

    if len(grid) < 3:
        return records, issues, debug

    header = " ".join(" ".join(r) for r in grid[:5]).upper()
    if "MAIN ENGINE" not in header:
        return records, issues, debug

    debug["detected"] = True

    period_col = 1
    marker_col = 2
    first_cyl_col = 3
    remarks_col = len(grid[0]) - 1

    for c in range(len(grid[0])):
        head_text = " ".join(str(grid[r][c]).upper() for r in range(min(4, len(grid))))
        if "REMARK" in head_text:
            remarks_col = c
            break

    actual_cyls = max(1, min(7, remarks_col - first_cyl_col))
    if actual_cyls <= 0:
        actual_cyls = 6

    end = len(grid)
    for r, row in enumerate(grid):
        joined = " ".join(str(x).upper() for x in row)
        if any(stop in joined for stop in ["NOTE 1", "AUX. ENGINE", "AUX ENGINE", "TURBOCHARGER"]):
            end = r
            break

    component_order = 0
    r = 1
    while r < end - 1:
        debug["rows_scanned"] += 1
        row1 = grid[r]
        row2 = grid[r + 1] if r + 1 < end else []
        name = clean_name(row1[0] if len(row1) > 0 else "")
        period_raw = row1[period_col] if len(row1) > period_col else ""
        marker = normalize_whitespace(row1[marker_col] if len(row1) > marker_col else "")

        marker_ok = ("1" in marker) if marker else False

        if is_component_name(name) and marker_ok:
            component_order += 1
            debug["blocks_found"] += 1

            pair_marker = normalize_whitespace(row2[marker_col] if len(row2) > marker_col else "")
            block_issues = []
            if pair_marker and "2" not in pair_marker:
                block_issues.append(make_issue(
                    "warning", "ME_PAIR_MARKER_ODD",
                    f"Expected marker 2 under '{name}', found '{pair_marker}'.",
                    table_idx, r
                ))

            period_norm = periodicity_value(period_raw)
            period_hours = period_norm["value"]

            for cyl in range(1, actual_cyls + 1):
                ci = first_cyl_col + cyl - 1
                date_raw = row1[ci] if ci < len(row1) else ""
                hrs_raw = row2[ci] if ci < len(row2) else ""

                date_norm = normalize_date_string(date_raw)
                hrs_norm = normalize_number_string(hrs_raw)

                local_issues = list(block_issues)
                if date_raw and not date_norm["value"]:
                    local_issues.append(make_issue("warning", "ME_BAD_DATE", f"Unclear date '{first_line(date_raw)}' for {name} Cyl {cyl}.", table_idx, r))
                if hrs_raw and hrs_norm["value"] <= 0:
                    local_issues.append(make_issue("warning", "ME_BAD_HOURS", f"Unclear hours '{first_line(hrs_raw)}' for {name} Cyl {cyl}.", table_idx, r + 1))
                if period_raw and period_hours <= 0:
                    local_issues.append(make_issue("info", "ME_PERIOD_NON_NUMERIC", f"Non-numeric periodicity '{first_line(period_raw)}' for {name}.", table_idx, r))

                # permissive row admission restored
                if date_norm["value"] or hrs_norm["value"] > 0 or first_line(date_raw) or first_line(hrs_raw):
                    repair_notes = period_norm["notes"] + date_norm["notes"] + hrs_norm["notes"]
                    was_repaired = int(period_norm["repaired"] or date_norm["repaired"] or hrs_norm["repaired"])
                    rec = normalize_row_record({
                        "category": "MAIN_ENGINE",
                        "engine_label": "ME",
                        "unit": f"Cyl {cyl}",
                        "description": name,
                        "periodicity_raw": first_line(period_raw),
                        "periodicity_hours": period_hours,
                        "last_oh_date": date_norm["value"],
                        "hrs_since": hrs_norm["value"],
                        "pct_used": compute_pct(hrs_norm["value"], period_hours),
                        "status": compute_status(hrs_norm["value"], period_hours),
                        "source_table_index": table_idx,
                        "section_order": section_base,
                        "component_order": component_order,
                        "engine_order": 0,
                        "unit_order": cyl,
                        "source_row_start": r,
                        "source_row_end": r + 1,
                        "source_col_date": ci,
                        "source_col_hours": ci,
                        "raw_date_text": date_norm["raw"],
                        "raw_hours_text": hrs_norm["raw"],
                        "normalized_date_text": date_norm["value"],
                        "normalized_hours_text": hrs_norm["normalized_text"],
                        "was_repaired": was_repaired,
                        "repair_notes": " | ".join(dict.fromkeys(repair_notes)),
                    })
                    rec["confidence"] = build_confidence(
                        issue_count=len(local_issues),
                        repaired=bool(was_repaired),
                        date_ok=bool(date_norm["value"]),
                        hrs_ok=(hrs_norm["value"] > 0),
                        period_ok=(period_hours > 0)
                    )
                    rec["issue_count"] = len(local_issues)
                    rk = row_key(rec)
                    for it in local_issues:
                        it["row_key"] = rk
                        issues.append(it)
                    records.append(rec)
                    debug["records_emitted"] += 1
            r += 2
        else:
            r += 1

    return records, issues, debug

def parse_aux_table(grid: List[List[str]], table_idx: int, section_base: int = 2000) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, Any]]:
    records, issues = [], []
    debug = {"table_idx": table_idx, "detected": False, "rows_scanned": 0, "blocks_found": 0, "records_emitted": 0}

    joined = " ".join(" ".join(r) for r in grid[:14]).upper()
    if "AUX. ENGINE" not in joined and "AUX ENGINE" not in joined:
        return records, issues, debug

    debug["detected"] = True

    desc_row = None
    for ri, row in enumerate(grid):
        if any("DESCRIPTION" == normalize_whitespace(c).upper() for c in row):
            desc_row = ri
            break
    if desc_row is None:
        issues.append(make_issue("warning", "AUX_DESC_ROW_NOT_FOUND", "Auxiliary engine description row not found.", table_idx))
        return records, issues, debug

    start_col = 3
    cyl_count = 6
    seen_numeric_headers = []
    for ci in range(2, len(grid[desc_row])):
        t = normalize_whitespace(grid[desc_row][ci])
        if re.fullmatch(r'\d+', t):
            seen_numeric_headers.append((ci, int(t)))
    if seen_numeric_headers:
        start_col = min(ci for ci, _ in seen_numeric_headers)
        cyl_count = max(1, min(7, max(n for _, n in seen_numeric_headers)))

    engines_found = ["AUX-1", "AUX-2", "AUX-3"]
    component_order = 0
    r = desc_row + 1

    while r < len(grid) - 1:
        debug["rows_scanned"] += 1
        row1 = grid[r]
        row2 = grid[r + 1] if r + 1 < len(grid) else []

        name = clean_name(row1[0] if len(row1) > 0 else "")
        period_raw = row1[1] if len(row1) > 1 else ""
        marker = normalize_whitespace(row1[2] if len(row1) > 2 else "")
        marker_ok = ("1" in marker) if marker else False

        if is_component_name(name) and marker_ok:
            component_order += 1
            debug["blocks_found"] += 1

            pair_marker = normalize_whitespace(row2[2] if len(row2) > 2 else "")
            block_issues = []
            if pair_marker and "2" not in pair_marker:
                block_issues.append(make_issue(
                    "warning", "AUX_PAIR_MARKER_ODD",
                    f"Expected marker 2 under AUX component '{name}', found '{pair_marker}'.",
                    table_idx, r
                ))

            period_norm = periodicity_value(period_raw)
            period_hours = period_norm["value"]

            for eng_idx, eng_label in enumerate(engines_found):
                grp_start = start_col + eng_idx * cyl_count
                for cyl in range(1, cyl_count + 1):
                    ci = grp_start + cyl - 1
                    date_raw = row1[ci] if ci < len(row1) else ""
                    hrs_raw = row2[ci] if ci < len(row2) else ""

                    date_norm = normalize_date_string(date_raw)
                    hrs_norm = normalize_number_string(hrs_raw)

                    local_issues = list(block_issues)
                    if date_raw and not date_norm["value"]:
                        local_issues.append(make_issue("warning", "AUX_BAD_DATE", f"Unclear date '{first_line(date_raw)}' for {name} {eng_label} Cyl {cyl}.", table_idx, r))
                    if hrs_raw and hrs_norm["value"] <= 0:
                        local_issues.append(make_issue("warning", "AUX_BAD_HOURS", f"Unclear hours '{first_line(hrs_raw)}' for {name} {eng_label} Cyl {cyl}.", table_idx, r + 1))
                    if period_raw and period_hours <= 0:
                        local_issues.append(make_issue("info", "AUX_PERIOD_NON_NUMERIC", f"Non-numeric periodicity '{first_line(period_raw)}' for {name}.", table_idx, r))

                    # permissive row admission restored
                    if date_norm["value"] or hrs_norm["value"] > 0 or first_line(date_raw) or first_line(hrs_raw):
                        repair_notes = period_norm["notes"] + date_norm["notes"] + hrs_norm["notes"]
                        was_repaired = int(period_norm["repaired"] or date_norm["repaired"] or hrs_norm["repaired"])
                        rec = normalize_row_record({
                            "category": "AUX_ENGINE",
                            "engine_label": eng_label,
                            "unit": f"Cyl {cyl}",
                            "description": name,
                            "periodicity_raw": first_line(period_raw),
                            "periodicity_hours": period_hours,
                            "last_oh_date": date_norm["value"],
                            "hrs_since": hrs_norm["value"],
                            "pct_used": compute_pct(hrs_norm["value"], period_hours),
                            "status": compute_status(hrs_norm["value"], period_hours),
                            "source_table_index": table_idx,
                            "section_order": section_base,
                            "component_order": component_order,
                            "engine_order": eng_idx + 1,
                            "unit_order": cyl,
                            "source_row_start": r,
                            "source_row_end": r + 1,
                            "source_col_date": ci,
                            "source_col_hours": ci,
                            "raw_date_text": date_norm["raw"],
                            "raw_hours_text": hrs_norm["raw"],
                            "normalized_date_text": date_norm["value"],
                            "normalized_hours_text": hrs_norm["normalized_text"],
                            "was_repaired": was_repaired,
                            "repair_notes": " | ".join(dict.fromkeys(repair_notes)),
                        })
                        rec["confidence"] = build_confidence(
                            issue_count=len(local_issues),
                            repaired=bool(was_repaired),
                            date_ok=bool(date_norm["value"]),
                            hrs_ok=(hrs_norm["value"] > 0),
                            period_ok=(period_hours > 0)
                        )
                        rec["issue_count"] = len(local_issues)
                        rk = row_key(rec)
                        for it in local_issues:
                            it["row_key"] = rk
                            issues.append(it)
                        records.append(rec)
                        debug["records_emitted"] += 1
            r += 2
        else:
            r += 1

    return records, issues, debug

def parse_other_equipment(grid: List[List[str]], table_idx: int, section_base: int = 3000) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    records, issues = [], []

    joined = " ".join(" ".join(r) for r in grid[:10]).upper()
    if not any(k in joined for k in ["TURBOCHARGER", "COOLERS", "A/C", "AUXILIARY BOILER", "MAIN AIR COMPRESSORS", "DG NO"]):
        return records, issues

    skip = {"", "TURBOCHARGER", "COOLERS", "A/C & REFR. COMPRESSORS", "AUXILIARY BOILER", "EXH GAS BOILER",
            "MAIN AIR COMPRESSORS", "PERIODICTLY", "DATE OF LAST O/H", "RUN HRS", "DESCRIPTION"}

    item_order = 0
    for r, row in enumerate(grid):
        for sec_name, desc_col, date_col, hrs_col, sec_ord in [
            ("Turbocharger / Aux Boiler", 0, 1, 2, section_base + 10),
            ("Coolers / Exh Gas Boiler", 5, 6, 7, section_base + 20),
            ("A/C & Compressors", 10, 11, 12, section_base + 30),
        ]:
            desc = clean_name(row[desc_col] if desc_col < len(row) else "")
            if desc.upper() in skip or not is_component_name(desc):
                continue
            item_order += 1
            d = normalize_date_string(row[date_col] if date_col < len(row) else "")
            h = normalize_number_string(row[hrs_col] if hrs_col < len(row) else "")
            if d["value"] or h["value"] > 0 or d["raw"] or h["raw"]:
                records.append({
                    "section": sec_name,
                    "description": desc,
                    "last_date": d["value"] or d["raw"],
                    "run_hrs": safe_int(h["value"]) if h["value"] > 0 else (h["raw"] or ""),
                    "source_table_index": table_idx,
                    "source_row": r,
                    "section_order": sec_ord,
                    "item_order": item_order,
                })

    return records, issues

def dedupe_rows(records: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    out, issues = [], []
    seen = set()
    for rec in normalize_rows_payload(records):
        key = (
            rec.get("category"), rec.get("engine_label"), rec.get("unit"),
            rec.get("description"), rec.get("source_table_index"),
            rec.get("component_order"), rec.get("engine_order"), rec.get("unit_order")
        )
        if key in seen:
            issues.append(make_issue(
                "warning", "DUPLICATE_ROW",
                f"Dropped duplicate row {rec.get('description')} / {rec.get('engine_label')} / {rec.get('unit')}.",
                rec.get("source_table_index"), rec.get("source_row_start"), row_key(rec)
            ))
            continue
        seen.add(key)
        out.append(rec)
    return out, issues

def parse_docx_bytes(docx_bytes: bytes, filename: str = "") -> Dict[str, Any]:
    from docx import Document

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
        tf.write(docx_bytes)
        path = tf.name
    try:
        doc = Document(path)
    except Exception as e:
        raise ValueError(f"Cannot open DOCX: {e}")
    finally:
        try:
            os.unlink(path)
        except Exception:
            pass

    if not doc.tables:
        raise ValueError("No tables found. File does not appear to be a valid TEC-004 report.")

    vessel_name, report_date, issues = detect_vessel_and_date(doc)
    me_total_hrs = 0
    me_this_month = 0
    me_rows, aux_rows, oe_rows = [], [], []
    debug_tables = []

    for ti, table in enumerate(doc.tables):
        grid = rect_grid(table)
        if not grid:
            continue

        if ti == 0:
            me_total_hrs, me_this_month = detect_me_totals(grid)

        a, ia, dbg_me = parse_me_table(grid, ti, 1000)
        b, ib, dbg_ax = parse_aux_table(grid, ti, 2000)
        c, ic = parse_other_equipment(grid, ti, 3000)

        me_rows.extend(a)
        aux_rows.extend(b)
        oe_rows.extend(c)
        issues.extend(ia)
        issues.extend(ib)
        issues.extend(ic)

        debug_tables.append({
            "table_index": ti,
            "me_detected": dbg_me["detected"],
            "me_blocks": dbg_me["blocks_found"],
            "me_rows": dbg_me["records_emitted"],
            "aux_detected": dbg_ax["detected"],
            "aux_blocks": dbg_ax["blocks_found"],
            "aux_rows": dbg_ax["records_emitted"],
        })

    all_rows = me_rows + aux_rows
    all_rows, dd_issues = dedupe_rows(all_rows)
    issues.extend(dd_issues)

    if not all_rows:
        issues.append(make_issue("error", "NO_COMPONENTS", "No ME or AUX components extracted."))

    me_rows = [r for r in all_rows if r.get("category") == "MAIN_ENGINE"]
    aux_rows = [r for r in all_rows if r.get("category") == "AUX_ENGINE"]

    debug = {
        "tables_scanned": len(doc.tables),
        "me_rows_total": len(me_rows),
        "aux_rows_total": len(aux_rows),
        "oe_rows_total": len(oe_rows),
        "issues_total": len(issues),
        "tables": debug_tables,
    }

    return normalize_parsed_payload({
        "vessel_name": vessel_name,
        "report_date": report_date,
        "me_total_hrs": me_total_hrs,
        "me_this_month": me_this_month,
        "me_comps": me_rows,
        "aux_comps": aux_rows,
        "components": all_rows,
        "other_equipment": oe_rows,
        "issues": issues,
        "filename": filename,
        "uploaded_at": now_iso(),
        "debug": debug,
    })

# ============================================================
# SAVE / HISTORY
# ============================================================
def ensure_vessel(conn, vessel_name: str) -> int:
    cur = conn.cursor()
    cur.execute("SELECT id FROM vessels WHERE name = ?", (vessel_name,))
    row = cur.fetchone()
    if row:
        return row["id"]
    cur.execute("INSERT INTO vessels (name, created_at) VALUES (?, ?)", (vessel_name, now_iso()))
    conn.commit()
    return cur.lastrowid

def report_exists(conn, vessel_name: str, report_date: str, file_hash: str) -> bool:
    cur = conn.cursor()
    cur.execute(
        "SELECT id FROM reports WHERE vessel_name = ? AND report_date = ? AND file_hash = ?",
        (vessel_name, report_date, file_hash)
    )
    return cur.fetchone() is not None

def save_report(parsed: Dict[str, Any]) -> Tuple[bool, str]:
    parsed = normalize_parsed_payload(parsed)
    conn = get_conn()
    try:
        vessel_id = ensure_vessel(conn, parsed["vessel_name"])
        if report_exists(conn, parsed["vessel_name"], parsed["report_date"], parsed["file_hash"]):
            return False, "This report already exists in the database."

        cur = conn.cursor()
        cur.execute("""
            INSERT INTO reports (
                vessel_id, vessel_name, report_date, filename, file_hash,
                parser_version, me_total_hrs, me_this_month, raw_json, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            vessel_id, parsed["vessel_name"], parsed["report_date"], parsed["filename"],
            parsed["file_hash"], PARSER_VERSION, parsed["me_total_hrs"], parsed["me_this_month"],
            json.dumps(parsed, ensure_ascii=False, default=str), now_iso()
        ))
        report_id = cur.lastrowid

        for rec in normalize_rows_payload(parsed["components"]):
            cur.execute("""
                INSERT INTO parsed_rows (
                    report_id, category, engine_label, unit, description,
                    periodicity_raw, periodicity_hours, last_oh_date, hrs_since, pct_used, status,
                    source_table_index, section_order, component_order, engine_order, unit_order,
                    source_row_start, source_row_end, source_col_date, source_col_hours,
                    raw_date_text, raw_hours_text, normalized_date_text, normalized_hours_text,
                    confidence, issue_count, was_repaired, repair_notes, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                report_id,
                rec["category"], rec["engine_label"], rec["unit"], rec["description"],
                rec["periodicity_raw"], rec["periodicity_hours"], rec["last_oh_date"], rec["hrs_since"],
                rec["pct_used"], rec["status"], rec["source_table_index"], rec["section_order"],
                rec["component_order"], rec["engine_order"], rec["unit_order"],
                rec["source_row_start"], rec["source_row_end"], rec["source_col_date"], rec["source_col_hours"],
                rec["raw_date_text"], rec["raw_hours_text"], rec["normalized_date_text"], rec["normalized_hours_text"],
                rec["confidence"], rec["issue_count"], rec["was_repaired"], rec["repair_notes"], now_iso()
            ))

        for issue in parsed.get("issues", []):
            if not isinstance(issue, dict):
                continue
            cur.execute("""
                INSERT INTO parse_issues (
                    report_id, row_key, severity, issue_code, message,
                    table_index, row_index, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                report_id, issue.get("row_key",""), issue.get("severity",""), issue.get("issue_code",""),
                issue.get("message",""), issue.get("table_index"), issue.get("row_index"), now_iso()
            ))

        conn.commit()
        return True, f"Saved successfully. Report ID: {report_id}"
    except Exception as e:
        conn.rollback()
        return False, f"Save failed: {e}"
    finally:
        conn.close()

def load_recent_reports(limit: int = 30) -> pd.DataFrame:
    conn = get_conn()
    try:
        df = pd.read_sql_query(f"""
            SELECT
                r.id, r.vessel_name, r.report_date, r.filename, r.parser_version,
                r.me_total_hrs, r.me_this_month, r.created_at,
                COUNT(DISTINCT pr.id) AS parsed_rows,
                COALESCE(SUM(CASE WHEN pi.severity='error' THEN 1 ELSE 0 END),0) AS errors,
                COALESCE(SUM(CASE WHEN pi.severity='warning' THEN 1 ELSE 0 END),0) AS warnings
            FROM reports r
            LEFT JOIN parsed_rows pr ON pr.report_id = r.id
            LEFT JOIN parse_issues pi ON pi.report_id = r.id
            GROUP BY r.id
            ORDER BY r.id DESC
            LIMIT {int(limit)}
        """, conn)
        return df
    finally:
        conn.close()

# ============================================================
# TABLE RENDERING
# ============================================================
def build_preview_df(records: List[Dict[str, Any]], include_trace: bool = False, mode: str = "source") -> pd.DataFrame:
    base_cols = ["Status","Component","Engine","Unit","Periodicity","Last O/H","Hrs Since","Used %","Confidence","Issues","Repaired"]
    if include_trace:
        base_cols += ["Table","Rows","Raw Date","Raw Hrs","Repair Notes"]

    if not records:
        return pd.DataFrame(columns=base_cols)

    if isinstance(records, pd.DataFrame):
        df = records.copy()
    else:
        df = pd.DataFrame([normalize_row_record(r) if isinstance(r, dict) else normalize_row_record({}) for r in records])

    defaults = {
        "status":"NO DATA","description":"","engine_label":"","unit":"",
        "periodicity_hours":0.0,"periodicity_raw":"","last_oh_date":"","hrs_since":0.0,
        "pct_used":0.0,"confidence":0.0,"issue_count":0,"was_repaired":0,"repair_notes":"",
        "source_table_index":"","section_order":0,"component_order":0,"engine_order":0,"unit_order":0,
        "source_row_start":"","source_row_end":"","raw_date_text":"","raw_hours_text":""
    }
    for col, val in defaults.items():
        if col not in df.columns:
            df[col] = val

    if mode == "priority":
        df["_s"] = df["status"].map(lambda x: STATUS_ORDER.get(str(x), 9))
        df["_pct"] = df["pct_used"].map(safe_float)
        df = df.sort_values(["_s", "_pct"], ascending=[True, False])
    elif mode == "component":
        df["_d"] = df["description"].astype(str).str.upper()
        df["_e"] = df["engine_label"].map(lambda x: ENGINE_ORDER.get(str(x), 999))
        df["_u"] = df["unit"].astype(str).map(cyl_num)
        df = df.sort_values(["_d", "_e", "_u"])
    else:
        df = df.sort_values(["section_order", "component_order", "engine_order", "unit_order", "source_row_start"])

    out = pd.DataFrame()
    out["Status"] = df["status"].map(lambda s: {
        "OVERDUE":"🔴 OVERDUE",
        "HIGH PRIORITY":"🟠 HIGH PRIORITY",
        "OK":"🟢 OK",
        "NO DATA":"🔵 NO DATA"
    }.get(str(s), "🔵 NO DATA"))
    out["Component"] = df["description"].astype(str)
    out["Engine"] = df["engine_label"].astype(str)
    out["Unit"] = df["unit"].astype(str)
    out["Periodicity"] = [
        f"{safe_int(x)}" if safe_float(x) > 0 else (str(raw).strip() if str(raw).strip() else "—")
        for x, raw in zip(df["periodicity_hours"], df["periodicity_raw"])
    ]
    out["Last O/H"] = [str(x).strip() if str(x).strip() else "—" for x in df["last_oh_date"]]
    out["Hrs Since"] = [f"{safe_int(x)}" if safe_float(x) > 0 else "—" for x in df["hrs_since"]]
    out["Used %"] = [round(safe_float(x)*100,1) if safe_float(x) > 0 else 0.0 for x in df["pct_used"]]
    out["Confidence"] = [round(safe_float(x)*100,0) for x in df["confidence"]]
    out["Issues"] = [safe_int(x) for x in df["issue_count"]]
    out["Repaired"] = ["Yes" if safe_int(x)==1 else "—" for x in df["was_repaired"]]

    if include_trace:
        out["Table"] = df["source_table_index"].astype(str)
        out["Rows"] = [f"{a}-{b}" if str(a).strip() or str(b).strip() else "—" for a, b in zip(df["source_row_start"], df["source_row_end"])]
        out["Raw Date"] = df["raw_date_text"].astype(str)
        out["Raw Hrs"] = df["raw_hours_text"].astype(str)
        out["Repair Notes"] = df["repair_notes"].astype(str)

    return out

def issues_df(issues: List[Dict[str, Any]]) -> pd.DataFrame:
    if not issues:
        return pd.DataFrame(columns=["Severity","Code","Message","Table","Row","Row Key"])
    df = pd.DataFrame([x for x in issues if isinstance(x, dict)])
    if df.empty:
        return pd.DataFrame(columns=["Severity","Code","Message","Table","Row","Row Key"])
    for c in ["severity","issue_code","message","table_index","row_index","row_key"]:
        if c not in df.columns:
            df[c] = ""
    df["_sev"] = df["severity"].map(lambda x: SEVERITY_ORDER.get(str(x), 9))
    df = df.sort_values(["_sev", "table_index", "row_index"])
    out = pd.DataFrame()
    out["Severity"] = df["severity"]
    out["Code"] = df["issue_code"]
    out["Message"] = df["message"]
    out["Table"] = df["table_index"]
    out["Row"] = df["row_index"]
    out["Row Key"] = df["row_key"]
    return out

TABLE_CONFIG = {
    "Status": st.column_config.TextColumn("Status", width=145),
    "Component": st.column_config.TextColumn("Component", width=270),
    "Engine": st.column_config.TextColumn("Engine", width=82),
    "Unit": st.column_config.TextColumn("Unit", width=72),
    "Periodicity": st.column_config.TextColumn("Periodicity", width=110),
    "Last O/H": st.column_config.TextColumn("Last O/H", width=118),
    "Hrs Since": st.column_config.TextColumn("Hrs Since", width=96),
    "Used %": st.column_config.ProgressColumn("Used %", min_value=0, max_value=150, format="%.1f%%", width=130),
    "Confidence": st.column_config.ProgressColumn("Confidence", min_value=0, max_value=100, format="%d%%", width=120),
    "Issues": st.column_config.NumberColumn("Issues", width=70),
    "Repaired": st.column_config.TextColumn("Repaired", width=80),
    "Table": st.column_config.TextColumn("Table", width=62),
    "Rows": st.column_config.TextColumn("Rows", width=75),
    "Raw Date": st.column_config.TextColumn("Raw Date", width=120),
    "Raw Hrs": st.column_config.TextColumn("Raw Hrs", width=110),
    "Repair Notes": st.column_config.TextColumn("Repair Notes", width=340),
}

def show_df(df: pd.DataFrame, height: int = None):
    h = height or min(880, 38*len(df) + 44)
    cfg = {k: v for k, v in TABLE_CONFIG.items() if k in df.columns}
    st.dataframe(df, use_container_width=True, hide_index=True, height=h, column_config=cfg)

# ============================================================
# STATE
# ============================================================
if "parsed_reports" not in st.session_state:
    st.session_state.parsed_reports = []
if "active_report_hash" not in st.session_state:
    st.session_state.active_report_hash = None
if "save_feedback" not in st.session_state:
    st.session_state.save_feedback = {}

# ============================================================
# HEADER
# ============================================================
st.markdown("""
<div class="hero">
  <div>
    <div class="hero-kicker">Fleet Running Hours Command</div>
    <div class="hero-title">TEC‑004 Parser & Review Console</div>
    <div class="hero-sub">
      Multi-report ingestion, exact source-order matrices, tolerant crew-input normalization,
      traceable repairs, and SQLite-backed persistence in a premium review-first workspace.
    </div>
  </div>
</div>
<div class="hero-rule"></div>
""", unsafe_allow_html=True)

# ============================================================
# UPLOAD
# ============================================================
st.markdown('<div class="upload-panel">', unsafe_allow_html=True)
u1, u2 = st.columns([2.2, 1.1], gap="large")
with u1:
    files = st.file_uploader("Upload TEC-004 reports", type=["doc"], accept_multiple_files=True, label_visibility="collapsed")
    st.markdown('<div class="muted">Upload one or many TEC‑004 legacy <b>.doc</b> reports. Every file is parsed into a staged review object before save.</div>', unsafe_allow_html=True)
with u2:
    st.markdown(
        '<div class="muted"><b>Default matrix order</b><br>Exact Word-file order<br><br>'
        '<b>Parser mode</b><br>Hybrid permissive extraction + repair-aware normalization<br><br>'
        '<b>Goal</b><br>Show rows first, then warn precisely</div>',
        unsafe_allow_html=True
    )
st.markdown('</div>', unsafe_allow_html=True)

if files:
    existing_hashes = {r.get("file_hash") for r in st.session_state.parsed_reports}
    new_items = [f for f in files if md5_bytes(f.getvalue()) not in existing_hashes]
    if new_items:
        prog = st.progress(0)
        for i, uploaded in enumerate(new_items, start=1):
            raw = uploaded.getvalue()
            fh = md5_bytes(raw)
            with st.spinner(f"Parsing {uploaded.name}..."):
                try:
                    docx_bytes = convert_doc_to_docx(raw)
                    parsed = parse_docx_bytes(docx_bytes, filename=uploaded.name)
                    parsed["file_hash"] = fh
                    parsed["filename"] = uploaded.name
                    parsed = normalize_parsed_payload(parsed)
                    st.session_state.parsed_reports.append(parsed)
                    if st.session_state.active_report_hash is None:
                        st.session_state.active_report_hash = fh
                except Exception as e:
                    fail = normalize_parsed_payload({
                        "vessel_name": "UNKNOWN",
                        "report_date": "",
                        "filename": uploaded.name,
                        "file_hash": fh,
                        "components": [],
                        "me_comps": [],
                        "aux_comps": [],
                        "other_equipment": [],
                        "issues": [make_issue("error", "PARSE_FAILURE", f"{uploaded.name}: {e}")]
                    })
                    st.session_state.parsed_reports.append(fail)
                    if st.session_state.active_report_hash is None:
                        st.session_state.active_report_hash = fh
            prog.progress(i / len(new_items))
        prog.empty()

reports = [normalize_parsed_payload(r) for r in st.session_state.parsed_reports]

st.markdown("## Report Queue")
if not reports:
    st.info("Upload one or more TEC-004 .doc reports to begin.")
    st.stop()

queue_rows = []
for r in reports:
    errs = sum(1 for x in r["issues"] if isinstance(x, dict) and x.get("severity") == "error")
    warns = sum(1 for x in r["issues"] if isinstance(x, dict) and x.get("severity") == "warning")
    repaired = sum(1 for x in r["components"] if safe_int(x.get("was_repaired")) == 1)
    queue_rows.append({
        "Active": "●" if r["file_hash"] == st.session_state.active_report_hash else "",
        "Filename": r["filename"],
        "Vessel": r["vessel_name"],
        "Report Date": r["report_date"] or "—",
        "Rows": len(r["components"]),
        "Warnings": warns,
        "Errors": errs,
        "Repaired Rows": repaired,
        "Hash": r["file_hash"],
    })

queue_df = pd.DataFrame(queue_rows)
st.dataframe(
    queue_df,
    use_container_width=True,
    hide_index=True,
    height=min(360, 38*len(queue_df)+44),
    column_config={
        "Active": st.column_config.TextColumn("Active", width=55),
        "Filename": st.column_config.TextColumn("Filename", width=260),
        "Vessel": st.column_config.TextColumn("Vessel", width=150),
        "Report Date": st.column_config.TextColumn("Report Date", width=120),
        "Rows": st.column_config.NumberColumn("Rows", width=70),
        "Warnings": st.column_config.NumberColumn("Warnings", width=90),
        "Errors": st.column_config.NumberColumn("Errors", width=70),
        "Repaired Rows": st.column_config.NumberColumn("Repaired", width=85),
        "Hash": st.column_config.TextColumn("Hash", width=240),
    }
)

sel_options = {f"{r['filename']}  |  {r['vessel_name']}  |  {r['report_date'] or '—'}": r["file_hash"] for r in reports}
selected_label = st.selectbox(
    "Active report",
    list(sel_options.keys()),
    index=list(sel_options.values()).index(st.session_state.active_report_hash) if st.session_state.active_report_hash in sel_options.values() else 0
)
st.session_state.active_report_hash = sel_options[selected_label]

active = next((r for r in reports if r["file_hash"] == st.session_state.active_report_hash), reports[0])
active = normalize_parsed_payload(active)

all_rows = normalize_rows_payload(active["components"])
me_rows = normalize_rows_payload(active["me_comps"])
aux_rows = normalize_rows_payload(active["aux_comps"])
oe_rows = normalize_other_equipment(active["other_equipment"])
all_issues = active["issues"]
debug = active.get("debug", {})

n_err = sum(1 for x in all_issues if isinstance(x, dict) and x.get("severity") == "error")
n_warn = sum(1 for x in all_issues if isinstance(x, dict) and x.get("severity") == "warning")
n_rep = sum(1 for x in all_rows if safe_int(x.get("was_repaired")) == 1)
n_od = sum(1 for x in all_rows if x.get("status") == "OVERDUE")
n_hp = sum(1 for x in all_rows if x.get("status") == "HIGH PRIORITY")

st.markdown(f"""
<div class="kpi-grid">
  <div class="kpi b"><div class="kpi-v">{active["vessel_name"]}</div><div class="kpi-l">Vessel</div></div>
  <div class="kpi"><div class="kpi-v">{active["report_date"] or "—"}</div><div class="kpi-l">Report Date</div></div>
  <div class="kpi"><div class="kpi-v">{safe_int(active["me_total_hrs"]):,}</div><div class="kpi-l">ME Total Hours</div></div>
  <div class="kpi"><div class="kpi-v">{safe_int(active["me_this_month"]):,}</div><div class="kpi-l">ME This Month</div></div>
  <div class="kpi a"><div class="kpi-v">{len(all_rows)}</div><div class="kpi-l">Parsed Rows</div></div>
  <div class="kpi r"><div class="kpi-v">{n_warn} / {n_err}</div><div class="kpi-l">Warnings / Errors</div></div>
</div>
""", unsafe_allow_html=True)

b1, b2, b3, b4 = st.columns(4)
with b1:
    st.markdown(f'<span class="badge red">Overdue: {n_od}</span>', unsafe_allow_html=True)
with b2:
    st.markdown(f'<span class="badge amber">High Priority: {n_hp}</span>', unsafe_allow_html=True)
with b3:
    st.markdown(f'<span class="badge blue">Repaired Rows: {n_rep}</span>', unsafe_allow_html=True)
with b4:
    st.markdown(f'<span class="badge cyan">Other Equipment: {len(oe_rows)}</span>', unsafe_allow_html=True)

if n_err:
    st.markdown(f'<div class="banner err"><b>{n_err}</b> parser errors detected. Review the issue log and parser diagnostics before save.</div>', unsafe_allow_html=True)
elif n_warn:
    st.markdown(f'<div class="banner warn"><b>{n_warn}</b> warnings detected. Rows were still emitted using permissive extraction where possible.</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="banner ok">No parser issues detected for the active report.</div>', unsafe_allow_html=True)

# ============================================================
# COMMAND BAR
# ============================================================
c1, c2, c3, c4 = st.columns([1.2, 1.4, 1.4, 2.8])
with c1:
    reviewed = st.checkbox("Review complete", value=False)
with c2:
    if st.button("Save active report", disabled=not reviewed):
        ok, msg = save_report(active)
        st.session_state.save_feedback[active["file_hash"]] = (ok, msg)
with c3:
    preview_export = build_preview_df(all_rows, include_trace=True, mode="source")
    st.download_button(
        "Download active CSV",
        data=preview_export.to_csv(index=False).encode("utf-8"),
        file_name=f"{Path(active['filename']).stem or 'tec004'}_preview.csv",
        mime="text/csv"
    )
with c4:
    st.markdown(f'<div class="micro">File: {active["filename"]} &nbsp;•&nbsp; Hash: {active["file_hash"][:18]}...</div>', unsafe_allow_html=True)

fb = st.session_state.save_feedback.get(active["file_hash"])
if fb:
    ok, msg = fb
    if ok:
        st.success(msg)
    else:
        st.warning(msg)

# ============================================================
# TABS
# ============================================================
tab_me, tab_aux, tab_oe, tab_issues, tab_debug, tab_history = st.tabs([
    "Main Engine", "Auxiliary Engines", "Other Equipment", "Parse Issues", "Parser Debug", "Saved Reports"
])

with tab_me:
    f1, f2, f3, f4 = st.columns([2.0, 2.0, 1.6, 2.4])
    with f1:
        comp_opt = ["All"] + [x for x in sorted({r["description"] for r in me_rows}) if x]
        me_comp = st.selectbox("Component", comp_opt, key="me_comp_h")
    with f2:
        me_status = st.selectbox("Status", ["All", "Overdue only", "High Priority +", "Issue rows only", "Repaired only"], key="me_status_h")
    with f3:
        me_trace = st.checkbox("Show trace", value=True, key="me_trace_h")
    with f4:
        me_sort = st.radio("Sort", ["Source order", "Priority", "Alphabetical"], horizontal=True, key="me_sort_h")

    me_view = me_rows[:]
    if me_comp != "All":
        me_view = [r for r in me_view if r["description"] == me_comp]
    if me_status == "Overdue only":
        me_view = [r for r in me_view if r["status"] == "OVERDUE"]
    elif me_status == "High Priority +":
        me_view = [r for r in me_view if r["status"] in ("OVERDUE", "HIGH PRIORITY")]
    elif me_status == "Issue rows only":
        me_view = [r for r in me_view if safe_int(r["issue_count"]) > 0]
    elif me_status == "Repaired only":
        me_view = [r for r in me_view if safe_int(r["was_repaired"]) == 1]

    me_mode = "source" if me_sort == "Source order" else ("priority" if me_sort == "Priority" else "component")
    me_df = build_preview_df(me_view, include_trace=me_trace, mode=me_mode)
    show_df(me_df)

with tab_aux:
    f1, f2, f3, f4, f5 = st.columns([1.4, 2.0, 2.0, 1.6, 2.4])
    with f1:
        eng_opt = ["All"] + [x for x in sorted({r["engine_label"] for r in aux_rows}) if x]
        ax_eng = st.selectbox("Engine", eng_opt, key="ax_eng_h")
    with f2:
        comp_opt = ["All"] + [x for x in sorted({r["description"] for r in aux_rows}) if x]
        ax_comp = st.selectbox("Component", comp_opt, key="ax_comp_h")
    with f3:
        ax_status = st.selectbox("Status", ["All", "Overdue only", "High Priority +", "Issue rows only", "Repaired only"], key="ax_status_h")
    with f4:
        ax_trace = st.checkbox("Show trace", value=True, key="ax_trace_h")
    with f5:
        ax_sort = st.radio("Sort", ["Source order", "Priority", "Alphabetical"], horizontal=True, key="ax_sort_h")

    ax_view = aux_rows[:]
    if ax_eng != "All":
        ax_view = [r for r in ax_view if r["engine_label"] == ax_eng]
    if ax_comp != "All":
        ax_view = [r for r in ax_view if r["description"] == ax_comp]
    if ax_status == "Overdue only":
        ax_view = [r for r in ax_view if r["status"] == "OVERDUE"]
    elif ax_status == "High Priority +":
        ax_view = [r for r in ax_view if r["status"] in ("OVERDUE", "HIGH PRIORITY")]
    elif ax_status == "Issue rows only":
        ax_view = [r for r in ax_view if safe_int(r["issue_count"]) > 0]
    elif ax_status == "Repaired only":
        ax_view = [r for r in ax_view if safe_int(r["was_repaired"]) == 1]

    ax_mode = "source" if ax_sort == "Source order" else ("priority" if ax_sort == "Priority" else "component")
    ax_df = build_preview_df(ax_view, include_trace=ax_trace, mode=ax_mode)
    show_df(ax_df)

with tab_oe:
    if not oe_rows:
        st.info("No other equipment rows found in the active report.")
    else:
        oe_df = pd.DataFrame(oe_rows).sort_values(["section_order", "item_order"])
        st.dataframe(
            oe_df[["section", "description", "last_date", "run_hrs", "source_table_index", "source_row"]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "section": st.column_config.TextColumn("Section", width=220),
                "description": st.column_config.TextColumn("Description", width=300),
                "last_date": st.column_config.TextColumn("Last Date", width=130),
                "run_hrs": st.column_config.TextColumn("Run Hrs", width=120),
                "source_table_index": st.column_config.TextColumn("Table", width=70),
                "source_row": st.column_config.TextColumn("Row", width=70),
            },
            height=min(780, 38*len(oe_df)+44)
        )

with tab_issues:
    idf = issues_df(all_issues)
    if idf.empty:
        st.success("No parse issues logged for the active report.")
    else:
        sev = st.multiselect("Severity", ["error","warning","info"], default=["error","warning","info"])
        idf2 = idf[idf["Severity"].isin(sev)] if sev else idf.copy()
        st.dataframe(
            idf2, use_container_width=True, hide_index=True,
            height=min(650, 38*len(idf2)+44),
            column_config={
                "Severity": st.column_config.TextColumn("Severity", width=90),
                "Code": st.column_config.TextColumn("Code", width=170),
                "Message": st.column_config.TextColumn("Message", width=540),
                "Table": st.column_config.TextColumn("Table", width=70),
                "Row": st.column_config.TextColumn("Row", width=70),
                "Row Key": st.column_config.TextColumn("Row Key", width=320),
            }
        )

with tab_debug:
    st.markdown("### Parser diagnostics")
    dbg = active.get("debug", {})
    d1, d2, d3, d4, d5 = st.columns(5)
    with d1:
        st.metric("Tables scanned", dbg.get("tables_scanned", 0))
    with d2:
        st.metric("ME rows", dbg.get("me_rows_total", 0))
    with d3:
        st.metric("AUX rows", dbg.get("aux_rows_total", 0))
    with d4:
        st.metric("OE rows", dbg.get("oe_rows_total", 0))
    with d5:
        st.metric("Issues", dbg.get("issues_total", 0))

    dbg_tables = dbg.get("tables", [])
    if dbg_tables:
        st.dataframe(pd.DataFrame(dbg_tables), use_container_width=True, hide_index=True)
    else:
        st.info("No parser diagnostics available.")

with tab_history:
    hist = load_recent_reports(30)
    if hist.empty:
        st.info("No saved reports in SQLite yet.")
    else:
        st.dataframe(
            hist, use_container_width=True, hide_index=True,
            height=min(700, 38*len(hist)+44),
            column_config={
                "id": st.column_config.NumberColumn("ID", width=65),
                "vessel_name": st.column_config.TextColumn("Vessel", width=150),
                "report_date": st.column_config.TextColumn("Report Date", width=120),
                "filename": st.column_config.TextColumn("Filename", width=260),
                "parser_version": st.column_config.TextColumn("Parser", width=220),
                "me_total_hrs": st.column_config.NumberColumn("ME Total", width=100, format="%d"),
                "me_this_month": st.column_config.NumberColumn("ME Month", width=100, format="%d"),
                "created_at": st.column_config.TextColumn("Saved At", width=180),
                "parsed_rows": st.column_config.NumberColumn("Rows", width=70),
                "errors": st.column_config.NumberColumn("Errors", width=70),
                "warnings": st.column_config.NumberColumn("Warnings", width=85),
            }
        )
