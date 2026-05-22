import streamlit as st
st.set_page_config(
    page_title="Fleet Running Hours v10.10",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

import os
import re
import json
import math
import sqlite3
import hashlib
import shutil
import tempfile
import subprocess
from pathlib import Path
from datetime import datetime
from typing import Any, Dict, List, Tuple

import pandas as pd

# ============================================================
# STYLE
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Space+Grotesk:wght@500;700&family=JetBrains+Mono:wght@400;500&display=swap');

:root{
  --bg:#070b12;
  --bg2:#0d131c;
  --bg3:#111926;
  --bg4:#162131;
  --line:#1e2b3d;
  --line2:#2a3c54;
  --text:#e6eef8;
  --muted:#9bb0c7;
  --soft:#6d839a;
  --gold:#b98a1d;
  --gold2:#ddb140;
  --gold3:#f1ca66;
  --red:#f0625d;
  --amber:#ffb65c;
  --green:#63c174;
  --blue:#5aa6f7;
  --cyan:#45d1d6;
  --radius:14px;
  --shadow:0 18px 60px rgba(0,0,0,.35);
}
html, body, [class*="css"]{
  background:var(--bg)!important;
  color:var(--text)!important;
  font-family:'Inter',sans-serif!important;
}
.main,.main>div,.block-container{background:var(--bg)!important;}
.block-container{
  max-width:100%!important;
  padding:1.1rem 1.6rem 3rem!important;
}
[data-testid="collapsedControl"],[data-testid="stSidebar"]{display:none!important;}
.main::before{
  content:"";
  position:fixed;
  inset:0;
  pointer-events:none;
  background:
    radial-gradient(ellipse 60% 35% at 0% 0%, rgba(221,177,64,.08), transparent 60%),
    radial-gradient(ellipse 55% 35% at 100% 100%, rgba(69,209,214,.05), transparent 55%);
}
.block-container>*{position:relative;z-index:1;}
.hero-k{font-size:.68rem;letter-spacing:.25em;text-transform:uppercase;color:var(--gold3);font-weight:700;}
.hero-h{font-family:'Space Grotesk',sans-serif;font-size:2.0rem;font-weight:700;letter-spacing:-.05em;color:var(--text);line-height:1.05;margin-top:.25rem;}
.hero-s{font-size:.95rem;color:var(--muted);line-height:1.6;max-width:1050px;margin-top:.5rem;}
.hero-rule{height:1px;background:linear-gradient(90deg,var(--gold2),var(--line),transparent);margin:.9rem 0 1.15rem;}
.panel{
  background:linear-gradient(180deg, rgba(255,255,255,.02), rgba(255,255,255,.01));
  border:1px solid var(--line);
  border-radius:var(--radius);
  padding:1rem 1rem 1rem;
  box-shadow:var(--shadow);
}
.upload-panel{
  background:
    linear-gradient(180deg, rgba(221,177,64,.05), rgba(90,166,247,.03)),
    linear-gradient(180deg, rgba(17,25,38,.98), rgba(13,19,28,.98));
  border:1px solid rgba(221,177,64,.24);
}
[data-testid="stFileUploadDropzone"]{
  background:rgba(221,177,64,.04)!important;
  border:1.5px dashed rgba(221,177,64,.52)!important;
  border-radius:16px!important;
  padding:2rem 1.2rem!important;
}
[data-testid="stFileUploadDropzone"]:hover{
  background:rgba(221,177,64,.06)!important;
  border-color:var(--gold3)!important;
}
.metric-grid{
  display:grid;
  grid-template-columns:repeat(7,1fr);
  gap:.8rem;
  margin:.95rem 0 1rem;
}
.metric-card{
  background:linear-gradient(180deg,var(--bg3),var(--bg2));
  border:1px solid var(--line);
  border-radius:var(--radius);
  padding:.85rem .95rem .95rem;
  box-shadow:var(--shadow);
  position:relative;
  overflow:hidden;
}
.metric-card::before{
  content:"";
  position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,var(--gold2),transparent 75%);
}
.metric-card.g::before{background:linear-gradient(90deg,var(--green),transparent 75%);}
.metric-card.r::before{background:linear-gradient(90deg,var(--red),transparent 75%);}
.metric-card.a::before{background:linear-gradient(90deg,var(--amber),transparent 75%);}
.metric-card.b::before{background:linear-gradient(90deg,var(--blue),transparent 75%);}
.metric-v{font-family:'Space Grotesk',sans-serif;font-size:1.55rem;font-weight:700;color:var(--text);letter-spacing:-.04em;line-height:1.06;}
.metric-l{font-size:.6rem;color:var(--soft);text-transform:uppercase;letter-spacing:.18em;margin-top:.36rem;}
.section-bar{display:flex;align-items:center;gap:.8rem;margin:1.5rem 0 .8rem;}
.section-icon{
  width:34px;height:34px;border-radius:10px;display:flex;align-items:center;justify-content:center;
  background:rgba(221,177,64,.08);border:1px solid rgba(221,177,64,.18);
}
.section-title{font-family:'Space Grotesk',sans-serif;font-size:1rem;font-weight:700;color:var(--text);}
.section-sub{font-family:'JetBrains Mono',monospace;font-size:.62rem;color:var(--soft);margin-top:2px;}
.section-rule{flex:1;height:1px;background:linear-gradient(90deg,var(--line),transparent);}
.section-badge{
  font-family:'JetBrains Mono',monospace;font-size:.62rem;padding:4px 9px;border-radius:999px;
  background:var(--bg3);border:1px solid var(--line2);color:var(--muted);
}
.banner{
  border-radius:var(--radius);
  padding:.85rem 1rem;
  border:1px solid var(--line);
  margin-bottom:.85rem;
  line-height:1.5;
  box-shadow:var(--shadow);
}
.banner.ok{border-color:rgba(99,193,116,.35);background:rgba(99,193,116,.09);color:#ddf7e4;}
.banner.warn{border-color:rgba(255,182,92,.35);background:rgba(255,182,92,.10);color:#ffe8cb;}
.banner.err{border-color:rgba(240,98,93,.35);background:rgba(240,98,93,.10);color:#ffd8d6;}
.muted{color:var(--muted);font-size:.88rem;line-height:1.65;}
.micro{font-family:'JetBrains Mono',monospace;color:var(--soft);font-size:.72rem;}
[data-testid="stDataFrame"]{
  border:1px solid var(--line)!important;
  border-radius:var(--radius)!important;
  overflow:hidden!important;
  box-shadow:var(--shadow)!important;
}
.dvn-scroller{background:var(--bg2)!important;}
.streamlit-expanderHeader{
  background:var(--bg3)!important;
  border:1px solid var(--line)!important;
  border-radius:12px!important;
  color:var(--text)!important;
}
.streamlit-expanderContent{
  background:var(--bg2)!important;
  border:1px solid var(--line)!important;
  border-top:none!important;
  border-radius:0 0 12px 12px!important;
}
.stButton>button{
  background:linear-gradient(135deg,var(--gold3),var(--gold2))!important;
  color:#061018!important;
  border:none!important;
  border-radius:11px!important;
  padding:.72rem 1.1rem!important;
  font-weight:800!important;
  text-transform:uppercase!important;
  letter-spacing:.05em!important;
}
.stDownloadButton>button{
  background:var(--bg3)!important;
  color:var(--text)!important;
  border:1px solid var(--line2)!important;
  border-radius:11px!important;
  padding:.72rem 1rem!important;
  font-weight:700!important;
}
div[data-baseweb="select"] > div,
.stTextInput > div > div > input{
  background:var(--bg4)!important;
  color:var(--text)!important;
  border:1px solid var(--line2)!important;
  border-radius:10px!important;
}
.stSelectbox label,.stMultiSelect label,.stCheckbox label,.stRadio label{
  color:var(--soft)!important;
  font-size:.66rem!important;
  text-transform:uppercase!important;
  letter-spacing:.12em!important;
}
@media (max-width:1280px){.metric-grid{grid-template-columns:repeat(4,1fr);}}
@media (max-width:840px){.metric-grid{grid-template-columns:repeat(2,1fr);}}
</style>
""", unsafe_allow_html=True)

# ============================================================
# CONSTANTS
# ============================================================
DB_PATH = "tec004_v1010.db"
PARSER_VERSION = "v10.10-operational"
STATUS_ORDER = {"OVERDUE": 0, "HIGH PRIORITY": 1, "OK": 2, "NO DATA": 3}
SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}
ENGINE_ORDER = {"ME": 0, "AUX-1": 1, "AUX-2": 2, "AUX-3": 3}

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
    CREATE TABLE IF NOT EXISTS reports(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL,
        report_date TEXT,
        filename TEXT NOT NULL,
        file_hash TEXT NOT NULL,
        parser_version TEXT NOT NULL,
        me_total_hrs REAL,
        me_this_month REAL,
        raw_json TEXT,
        created_at TEXT NOT NULL,
        UNIQUE(vessel_name, report_date, file_hash)
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS parsed_rows(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        report_id INTEGER NOT NULL,
        category TEXT,
        engine_label TEXT,
        unit TEXT,
        description TEXT,
        periodicity_raw TEXT,
        periodicity_hours REAL,
        last_oh_date TEXT,
        hrs_since REAL,
        pct_used REAL,
        status TEXT,
        source_table_index INTEGER,
        source_row_start INTEGER,
        source_row_end INTEGER,
        source_col_date INTEGER,
        source_col_hours INTEGER,
        raw_date_text TEXT,
        raw_hours_text TEXT,
        confidence REAL,
        issue_count INTEGER,
        was_repaired INTEGER,
        repair_notes TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(report_id) REFERENCES reports(id)
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS parse_issues(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        report_id INTEGER,
        severity TEXT,
        issue_code TEXT,
        message TEXT,
        row_key TEXT,
        table_index INTEGER,
        row_index INTEGER,
        created_at TEXT NOT NULL,
        FOREIGN KEY(report_id) REFERENCES reports(id)
    )
    """)
    conn.commit()
    conn.close()

init_db()

# ============================================================
# GENERIC HELPERS
# ============================================================
def now_iso() -> str:
    return datetime.utcnow().isoformat()

def md5_bytes(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()

def fl(text: Any) -> str:
    raw = str(text or "").replace("\x07", "").replace("\xa0", " ").replace("\t", " ")
    parts = re.split(r"[\r\n\x0b]+", raw)
    for p in parts:
        s = re.sub(r"\s+", " ", p).strip()
        if s:
            return s
    return ""

def norm_space(text: Any) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()

def safe_float(x) -> float:
    try:
        v = float(x)
        if math.isnan(v):
            return 0.0
        return v
    except Exception:
        return 0.0

def safe_int(x) -> int:
    try:
        return int(float(x))
    except Exception:
        return 0

def cyl_num(unit: str) -> int:
    m = re.search(r"(\d+)", str(unit or ""))
    return int(m.group(1)) if m else 999

def row_key(rec: Dict[str, Any]) -> str:
    return " | ".join([
        str(rec.get("category","")),
        str(rec.get("description","")),
        str(rec.get("engine_label","")),
        str(rec.get("unit","")),
        str(rec.get("source_table_index","")),
        str(rec.get("source_row_start",""))
    ])

def make_issue(severity: str, code: str, message: str, table_index=None, row_index=None, rk="") -> Dict[str, Any]:
    return {
        "severity": severity,
        "issue_code": code,
        "message": message,
        "table_index": table_index,
        "row_index": row_index,
        "row_key": rk,
    }

def clean_name(text: Any) -> str:
    t = fl(text)
    t = re.sub(r"(?i)^MV\s+", "", t)
    t = re.sub(r"(?i)ALEXIS\s*Date?.*", "", t)
    t = re.sub(r"(?i)Page\s*\d+\s*of\s*\d+", "", t)
    t = re.sub(r"\*+", "", t)
    t = re.sub(r"#+", "", t)
    t = re.sub(r"\s{2,}", " ", t).strip(" -:")
    return t.strip()

def is_component_name(name: str) -> bool:
    u = norm_space(name).upper()
    if not u or len(u) < 2:
        return False
    bad = [
        "DESCRIPTION","REMARKS","PERIODICITY","PERIODICTLY","DATE OF LAST","RUN HRS","RUNNING HOURS",
        "MAIN ENGINE","AUX. ENGINE","AUX ENGINE","CYL. NO","TYPE:","TOTAL RUNNING HOURS","THIS MONTH",
        "HOURS THIS MONTH","TOTAL HOURS","SERIAL NR","VESSEL","CHIEF ENGINEER","COPY TO","NOTE 1","NOTE 2"
    ]
    if any(x in u for x in bad):
        return False
    if re.fullmatch(r"[\d\s,./:\-\[\]()]+", u):
        return False
    if len(u) > 90:
        return False
    return bool(re.search(r"[A-Za-z]", u))

# ============================================================
# NORMALIZATION
# ============================================================
MONTHS = {
    "JAN":"JAN","JANUARY":"JAN",
    "FEB":"FEB","FEBRUARY":"FEB",
    "MAR":"MAR","MARCH":"MAR",
    "APR":"APR","APRIL":"APR",
    "MAY":"MAY",
    "JUN":"JUN","JUNE":"JUN",
    "JUL":"JUL","JULY":"JUL",
    "AUG":"AUG","AUGUST":"AUG",
    "SEP":"SEP","SEPT":"SEP","SEPTEMBER":"SEP",
    "OCT":"OCT","OCTOBER":"OCT",
    "NOV":"NOV","NOVEMBER":"NOV",
    "DEC":"DEC","DECEMBER":"DEC",
}

def clean_token(text: Any) -> Tuple[str, bool, List[str]]:
    raw = fl(text)
    notes = []
    repaired = False
    s = raw.strip()
    if not s:
        return "", False, []
    if "[" in s or "]" in s:
        s = s.replace("[", "").replace("]", "")
        repaired = True
        notes.append("Removed bracket artifacts")
    s2 = re.sub(r"\s+", " ", s).strip()
    if s2 != s:
        s = s2
        repaired = True
        notes.append("Collapsed repeated whitespace")
    return s, repaired, notes

def normalize_date(text: Any) -> Dict[str, Any]:
    raw = fl(text)
    token, repaired, notes = clean_token(text)
    u = token.upper().replace(".", "")
    if not u or u in {"N/A","NA","-","CENTRAL","COOLER"}:
        return {"raw": raw, "value": "", "repaired": repaired, "notes": notes}

    if re.fullmatch(r"\d+", u):
        return {"raw": raw, "value": "", "repaired": repaired, "notes": notes}

    m = re.match(r"^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$", u)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= mo <= 12 and 1 <= d <= 31:
            yy = str(y)[-2:]
            return {"raw": raw, "value": f"{d:02d}/{mo:02d}/{yy}", "repaired": repaired, "notes": notes}
        return {"raw": raw, "value": "", "repaired": True, "notes": notes + ["Rejected invalid numeric date"]}

    m = re.match(r"^(\d{1,2})\s+([A-Z]+)\s+(\d{2,4})$", u)
    if m:
        d = int(m.group(1))
        mon = MONTHS.get(m.group(2), "")
        y = str(m.group(3))[-2:]
        if mon and 1 <= d <= 31:
            return {"raw": raw, "value": f"{d:02d} {mon} {y}", "repaired": repaired, "notes": notes}

    if len(u) > 24:
        return {"raw": raw, "value": "", "repaired": True, "notes": notes + ["Rejected overlong date token"]}

    return {"raw": raw, "value": token, "repaired": repaired, "notes": notes}

def normalize_number(text: Any) -> Dict[str, Any]:
    raw = fl(text)
    token, repaired, notes = clean_token(text)
    u = token.upper()
    if not u or u in {"N/A","NA","-","CENTRAL","COOLER"}:
        return {"raw": raw, "value": 0.0, "text": "", "repaired": repaired, "notes": notes}
    if any(k in u for k in ["MONTH","YEAR","WEEK","DAY","OBSERVATION","OBS"]):
        return {"raw": raw, "value": 0.0, "text": token, "repaired": repaired, "notes": notes}

    m = re.search(r"\d[\d,.\s]*", u)
    if not m:
        return {"raw": raw, "value": 0.0, "text": token, "repaired": repaired, "notes": notes}

    block = m.group().strip()
    if re.search(r"^\d{4,6}\s+\d{2,3}$", block):
        left, right = block.split()
        block = left
        repaired = True
        notes = notes + [f"Trimmed suspicious trailing numeric fragment '{right}'"]

    block = block.replace(" ", "")
    sep = max(block.rfind("."), block.rfind(","))
    if sep > 0 and len(block) - sep == 4:
        block = re.sub(r"[,.]", "", block)
    elif sep > 0:
        block = re.sub(r"[,.]", "", block[:sep])
    else:
        block = re.sub(r"[,.]", "", block)

    try:
        val = float(block)
    except Exception:
        val = 0.0

    return {"raw": raw, "value": val, "text": block, "repaired": repaired, "notes": notes}

def compute_status(hrs: float, periodicity: float) -> str:
    if hrs <= 0 or periodicity <= 0:
        return "NO DATA"
    ratio = hrs / periodicity
    if ratio >= 1.0:
        return "OVERDUE"
    if ratio >= 0.8:
        return "HIGH PRIORITY"
    return "OK"

def compute_pct(hrs: float, periodicity: float) -> float:
    if hrs <= 0 or periodicity <= 0:
        return 0.0
    return round(hrs / periodicity, 4)

def confidence(issue_count: int, repaired: bool, has_date: bool, has_hrs: bool, has_period: bool) -> float:
    score = 1.0
    if repaired:
        score -= 0.08
    if not has_period:
        score -= 0.12
    if not has_date:
        score -= 0.08
    if not has_hrs:
        score -= 0.08
    score -= min(0.42, issue_count * 0.07)
    return max(0.0, round(score, 2))

# ============================================================
# DOC CONVERSION
# ============================================================
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found. Add libreoffice to the environment.")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as tf:
        tf.write(raw)
        src = tf.name
    outdir = tempfile.mkdtemp(prefix="lo_")
    out = os.path.join(outdir, Path(src).stem + ".docx")
    profile = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        r = subprocess.run(
            [soffice, "--headless", "--norestore", "--nofirststartwizard",
             f"-env:UserInstallation={profile}",
             "--convert-to", "docx", src, "--outdir", outdir],
            capture_output=True, timeout=120
        )
        if not os.path.exists(out):
            raise RuntimeError(r.stderr.decode("utf-8","ignore")[:400] or r.stdout.decode("utf-8","ignore")[:400])
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
# GRID EXTRACTION
# ============================================================
def rect_grid(table) -> List[List[str]]:
    """
    Operational compromise:
    - keep row.cells ordering
    - normalize to max column count
    - preserve first visible line
    This is intentionally simple because the parser below is pattern-driven on paired rows.
    """
    rows = []
    if not table.rows:
        return rows
    max_cols = max(len(r.cells) for r in table.rows)
    for row in table.rows:
        vals = []
        for cell in row.cells:
            raw = cell.text.replace("\x07", "")
            raw = re.sub(r"[\r\x0b]+", "\n", raw)
            lines = [norm_space(x) for x in raw.split("\n") if norm_space(x)]
            vals.append(lines[0] if lines else "")
        while len(vals) < max_cols:
            vals.append("")
        rows.append(vals)
    return rows

# ============================================================
# SECTION DETECTION
# ============================================================
def detect_vessel_and_report_date(doc) -> Tuple[str, str, List[Dict[str, Any]]]:
    vessel = "UNKNOWN"
    report_date = ""
    issues = []
    for p in doc.paragraphs:
        txt = norm_space(p.text)
        if not txt:
            continue
        mv = re.search(r"Vessel[’'`s]*\s*Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+)", txt, re.I)
        if mv and vessel == "UNKNOWN":
            vessel = clean_name(mv.group(1))
        md = re.search(r"Date\s*:?\s*(.+)", txt, re.I)
        if md and not report_date:
            report_date = normalize_date(md.group(1))["value"]
        if vessel != "UNKNOWN" and report_date:
            break
    if vessel == "UNKNOWN":
        issues.append(make_issue("warning", "VESSEL_NOT_FOUND", "Could not extract vessel name."))
    if not report_date:
        issues.append(make_issue("warning", "REPORT_DATE_NOT_FOUND", "Could not extract report date."))
    return vessel, report_date, issues

def detect_me_totals(grid: List[List[str]]) -> Tuple[float, float]:
    blob = " | ".join(" | ".join(r) for r in grid[:5])
    tot = 0.0
    month = 0.0
    m1 = re.search(r"Total Running Hours[\s:ǀ|]+([\d,]+)", blob, re.I)
    m2 = re.search(r"This Month[\s:]+([\d,]+)", blob, re.I)
    if m1:
        tot = normalize_number(m1.group(1))["value"]
    if m2:
        month = normalize_number(m2.group(1))["value"]
    return tot, month

# ============================================================
# MAIN ENGINE PARSER
# ============================================================
def parse_me_table(grid: List[List[str]], table_idx: int) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, Any]]:
    records, issues = [], []
    dbg = {"table_index": table_idx, "section": "ME", "detected": False, "blocks": 0, "rows_emitted": 0}

    if len(grid) < 4:
        return records, issues, dbg

    head_blob = " ".join(" ".join(r) for r in grid[:6]).upper()
    if "MAIN ENGINE" not in head_blob:
        return records, issues, dbg

    dbg["detected"] = True

    # Determine cylinder columns from the header row containing "CYL. No."
    cyl_header_row = None
    cyl_cols = []
    for ri in range(min(8, len(grid))):
        row_u = [norm_space(c).upper() for c in grid[ri]]
        if any("CYL. NO." in c or "CYL NO." in c for c in row_u):
            cyl_header_row = ri
            for ci, c in enumerate(row_u):
                if "CYL. NO." in c or "CYL NO." in c:
                    m = re.search(r"(\d+)", c)
                    cyl_no = int(m.group(1)) if m else len(cyl_cols) + 1
                    cyl_cols.append((ci, cyl_no))
            break

    # Fallback for compact extracted layout like the attached sample summary
    # In that layout row looks like: description | periodicity | marker | cyl1 | cyl2 ...
    if not cyl_cols:
        # try row 1 or 2
        for ri in range(min(5, len(grid))):
            row_u = [norm_space(c).upper() for c in grid[ri]]
            marker_cols = [i for i, c in enumerate(row_u) if "CYL" in c and "NO" in c]
            if marker_cols:
                cyl_cols = [(ci, idx + 1) for idx, ci in enumerate(marker_cols)]
                cyl_header_row = ri
                break

    if not cyl_cols:
        # conservative fallback based on sample structure
        # desc=0, periodicity=1, marker=2, cylinder dates/hours start at 3
        max_guess = min(7, max(1, len(grid[0]) - 3))
        cyl_cols = [(3 + i, i + 1) for i in range(max_guess)]

    start_row = (cyl_header_row + 1) if cyl_header_row is not None else 2

    stop_row = len(grid)
    for ri in range(start_row, len(grid)):
        blob = " ".join(grid[ri]).upper()
        if "NOTE 1" in blob or "AUX. ENGINE" in blob or "AUX ENGINE" in blob or "TURBOCHARGER" in blob:
            stop_row = ri
            break

    r = start_row
    component_order = 0
    while r < stop_row - 1:
        row1 = grid[r]
        row2 = grid[r + 1] if r + 1 < stop_row else []

        desc = clean_name(row1[0] if len(row1) > 0 else "")
        periodicity_raw = row1[1] if len(row1) > 1 else ""
        marker1 = fl(row1[2] if len(row1) > 2 else "")
        marker2 = fl(row2[2] if len(row2) > 2 else "")

        if is_component_name(desc) and ("1" in marker1):
            component_order += 1
            dbg["blocks"] += 1
            if marker2 and "2" not in marker2:
                issues.append(make_issue("warning", "ME_PAIR_MARKER", f"Unexpected second marker '{marker2}' under '{desc}'.", table_idx, r))

            p = normalize_number(periodicity_raw)
            periodicity = p["value"]

            for col_idx, cyl in cyl_cols:
                raw_date = row1[col_idx] if col_idx < len(row1) else ""
                raw_hrs = row2[col_idx] if col_idx < len(row2) else ""

                d = normalize_date(raw_date)
                h = normalize_number(raw_hrs)

                local_issues = []
                if raw_date and not d["value"]:
                    local_issues.append(make_issue("warning", "ME_BAD_DATE", f"Invalid/unclear date '{fl(raw_date)}' for {desc} Cyl {cyl}.", table_idx, r))
                if raw_hrs and h["value"] <= 0:
                    local_issues.append(make_issue("warning", "ME_BAD_HOURS", f"Invalid/unclear hours '{fl(raw_hrs)}' for {desc} Cyl {cyl}.", table_idx, r + 1))
                if periodicity_raw and periodicity <= 0:
                    local_issues.append(make_issue("info", "ME_NON_NUM_PERIOD", f"Non-numeric periodicity '{fl(periodicity_raw)}' for '{desc}'.", table_idx, r))

                # key operational rule:
                # emit if there is any meaningful raw payload in the paired cells
                if d["value"] or h["value"] > 0 or fl(raw_date) or fl(raw_hrs):
                    rec = {
                        "category": "MAIN_ENGINE",
                        "engine_label": "ME",
                        "unit": f"Cyl {cyl}",
                        "description": desc,
                        "periodicity_raw": fl(periodicity_raw),
                        "periodicity_hours": periodicity,
                        "last_oh_date": d["value"],
                        "hrs_since": h["value"],
                        "pct_used": compute_pct(h["value"], periodicity),
                        "status": compute_status(h["value"], periodicity),
                        "source_table_index": table_idx,
                        "source_row_start": r,
                        "source_row_end": r + 1,
                        "source_col_date": col_idx,
                        "source_col_hours": col_idx,
                        "raw_date_text": d["raw"],
                        "raw_hours_text": h["raw"],
                        "confidence": 0.0,
                        "issue_count": len(local_issues),
                        "was_repaired": int(d["repaired"] or h["repaired"] or p["repaired"]),
                        "repair_notes": " | ".join(dict.fromkeys(p["notes"] + d["notes"] + h["notes"])),
                        "component_order": component_order,
                        "engine_order": 0,
                        "unit_order": cyl,
                    }
                    rk = row_key(rec)
                    for it in local_issues:
                        it["row_key"] = rk
                        issues.append(it)
                    rec["confidence"] = confidence(
                        issue_count=len(local_issues),
                        repaired=bool(rec["was_repaired"]),
                        has_date=bool(d["value"]),
                        has_hrs=(h["value"] > 0),
                        has_period=(periodicity > 0)
                    )
                    records.append(rec)
                    dbg["rows_emitted"] += 1
            r += 2
        else:
            r += 1

    return records, issues, dbg

# ============================================================
# AUXILIARY ENGINE PARSER
# ============================================================
def parse_aux_table(grid: List[List[str]], table_idx: int) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, Any]]:
    records, issues = [], []
    dbg = {"table_index": table_idx, "section": "AUX", "detected": False, "blocks": 0, "rows_emitted": 0}

    joined = " ".join(" ".join(r) for r in grid[:20]).upper()
    if "AUX. ENGINE" not in joined and "AUX ENGINE" not in joined:
        return records, issues, dbg

    dbg["detected"] = True

    desc_row = None
    for ri, row in enumerate(grid):
        row_u = [norm_space(c).upper() for c in row]
        if len(row_u) >= 4 and row_u[0] == "DESCRIPTION" and "PERIODICTLY" in row_u[1]:
            desc_row = ri
            break
        if len(row_u) >= 4 and row_u[0] == "DESCRIPTION" and "PERIODICITY" in row_u[1]:
            desc_row = ri
            break

    if desc_row is None:
        issues.append(make_issue("warning", "AUX_DESC_NOT_FOUND", "Could not locate AUX DESCRIPTION row.", table_idx))
        return records, issues, dbg

    # Detect compact pattern from sample:
    # DESCRIPTION | PERIODICTLY | 1 | 2
    compact_mode = False
    aux_engine_labels = ["AUX-1"]
    row_u = [norm_space(c).upper() for c in grid[desc_row]]
    if len(row_u) >= 4 and row_u[2] == "1" and row_u[3] == "2":
        compact_mode = True

    component_order = 0
    r = desc_row + 1

    if compact_mode:
        while r < len(grid) - 1:
            row1 = grid[r]
            row2 = grid[r + 1] if r + 1 < len(grid) else []

            desc = clean_name(row1[0] if len(row1) > 0 else "")
            periodicity_raw = row1[1] if len(row1) > 1 else ""
            marker1 = fl(row1[2] if len(row1) > 2 else "")
            marker2 = fl(row2[2] if len(row2) > 2 else "")
            date_raw = row1[3] if len(row1) > 3 else ""
            hrs_raw = row2[3] if len(row2) > 3 else ""

            if is_component_name(desc) and ("1" in marker1):
                component_order += 1
                dbg["blocks"] += 1
                if marker2 and "2" not in marker2:
                    issues.append(make_issue("warning", "AUX_PAIR_MARKER", f"Unexpected AUX second marker '{marker2}' under '{desc}'.", table_idx, r))

                p = normalize_number(periodicity_raw)
                d = normalize_date(date_raw)
                h = normalize_number(hrs_raw)

                local_issues = []
                if date_raw and not d["value"]:
                    local_issues.append(make_issue("warning", "AUX_BAD_DATE", f"Invalid/unclear date '{fl(date_raw)}' for {desc}.", table_idx, r))
                if hrs_raw and h["value"] <= 0:
                    local_issues.append(make_issue("warning", "AUX_BAD_HOURS", f"Invalid/unclear hours '{fl(hrs_raw)}' for {desc}.", table_idx, r + 1))
                if periodicity_raw and p["value"] <= 0:
                    local_issues.append(make_issue("info", "AUX_NON_NUM_PERIOD", f"Non-numeric periodicity '{fl(periodicity_raw)}' for '{desc}'.", table_idx, r))

                if d["value"] or h["value"] > 0 or fl(date_raw) or fl(hrs_raw):
                    rec = {
                        "category": "AUX_ENGINE",
                        "engine_label": "AUX-1",
                        "unit": "General",
                        "description": desc,
                        "periodicity_raw": fl(periodicity_raw),
                        "periodicity_hours": p["value"],
                        "last_oh_date": d["value"],
                        "hrs_since": h["value"],
                        "pct_used": compute_pct(h["value"], p["value"]),
                        "status": compute_status(h["value"], p["value"]),
                        "source_table_index": table_idx,
                        "source_row_start": r,
                        "source_row_end": r + 1,
                        "source_col_date": 3,
                        "source_col_hours": 3,
                        "raw_date_text": d["raw"],
                        "raw_hours_text": h["raw"],
                        "confidence": 0.0,
                        "issue_count": len(local_issues),
                        "was_repaired": int(d["repaired"] or h["repaired"] or p["repaired"]),
                        "repair_notes": " | ".join(dict.fromkeys(p["notes"] + d["notes"] + h["notes"])),
                        "component_order": component_order,
                        "engine_order": 1,
                        "unit_order": 1,
                    }
                    rk = row_key(rec)
                    for it in local_issues:
                        it["row_key"] = rk
                        issues.append(it)
                    rec["confidence"] = confidence(
                        issue_count=len(local_issues),
                        repaired=bool(rec["was_repaired"]),
                        has_date=bool(d["value"]),
                        has_hrs=(h["value"] > 0),
                        has_period=(p["value"] > 0)
                    )
                    records.append(rec)
                    dbg["rows_emitted"] += 1
                r += 2
            else:
                if "DESCRIPTION" in " ".join(row1).upper() and "D/G" in " ".join(row1).upper():
                    break
                r += 1
        return records, issues, dbg

    # fallback grouped mode
    # DESCRIPTION | PERIODICTLY | AUX1 cyl1..n | AUX2 cyl1..n | AUX3 cyl1..n
    start_col = 3
    cyl_count = 6
    numeric_headers = []
    for ci in range(2, len(grid[desc_row])):
        txt = norm_space(grid[desc_row][ci])
        if re.fullmatch(r"\d+", txt):
            numeric_headers.append((ci, int(txt)))
    if numeric_headers:
        start_col = min(ci for ci, _ in numeric_headers)
        cyl_count = max(1, min(7, max(n for _, n in numeric_headers)))

    r = desc_row + 1
    while r < len(grid) - 1:
        row1 = grid[r]
        row2 = grid[r + 1] if r + 1 < len(grid) else []

        desc = clean_name(row1[0] if len(row1) > 0 else "")
        periodicity_raw = row1[1] if len(row1) > 1 else ""
        marker1 = fl(row1[2] if len(row1) > 2 else "")
        marker2 = fl(row2[2] if len(row2) > 2 else "")

        if is_component_name(desc) and ("1" in marker1):
            component_order += 1
            dbg["blocks"] += 1
            if marker2 and "2" not in marker2:
                issues.append(make_issue("warning", "AUX_PAIR_MARKER", f"Unexpected AUX second marker '{marker2}' under '{desc}'.", table_idx, r))

            p = normalize_number(periodicity_raw)
            for eng_idx, eng in enumerate(["AUX-1", "AUX-2", "AUX-3"]):
                grp_start = start_col + eng_idx * cyl_count
                for cyl in range(1, cyl_count + 1):
                    ci = grp_start + cyl - 1
                    raw_date = row1[ci] if ci < len(row1) else ""
                    raw_hrs = row2[ci] if ci < len(row2) else ""

                    d = normalize_date(raw_date)
                    h = normalize_number(raw_hrs)

                    local_issues = []
                    if raw_date and not d["value"]:
                        local_issues.append(make_issue("warning", "AUX_BAD_DATE", f"Invalid/unclear date '{fl(raw_date)}' for {desc} {eng} Cyl {cyl}.", table_idx, r))
                    if raw_hrs and h["value"] <= 0:
                        local_issues.append(make_issue("warning", "AUX_BAD_HOURS", f"Invalid/unclear hours '{fl(raw_hrs)}' for {desc} {eng} Cyl {cyl}.", table_idx, r + 1))

                    if d["value"] or h["value"] > 0 or fl(raw_date) or fl(raw_hrs):
                        rec = {
                            "category": "AUX_ENGINE",
                            "engine_label": eng,
                            "unit": f"Cyl {cyl}",
                            "description": desc,
                            "periodicity_raw": fl(periodicity_raw),
                            "periodicity_hours": p["value"],
                            "last_oh_date": d["value"],
                            "hrs_since": h["value"],
                            "pct_used": compute_pct(h["value"], p["value"]),
                            "status": compute_status(h["value"], p["value"]),
                            "source_table_index": table_idx,
                            "source_row_start": r,
                            "source_row_end": r + 1,
                            "source_col_date": ci,
                            "source_col_hours": ci,
                            "raw_date_text": d["raw"],
                            "raw_hours_text": h["raw"],
                            "confidence": 0.0,
                            "issue_count": len(local_issues),
                            "was_repaired": int(d["repaired"] or h["repaired"] or p["repaired"]),
                            "repair_notes": " | ".join(dict.fromkeys(p["notes"] + d["notes"] + h["notes"])),
                            "component_order": component_order,
                            "engine_order": eng_idx + 1,
                            "unit_order": cyl,
                        }
                        rk = row_key(rec)
                        for it in local_issues:
                            it["row_key"] = rk
                            issues.append(it)
                        rec["confidence"] = confidence(
                            issue_count=len(local_issues),
                            repaired=bool(rec["was_repaired"]),
                            has_date=bool(d["value"]),
                            has_hrs=(h["value"] > 0),
                            has_period=(p["value"] > 0)
                        )
                        records.append(rec)
                        dbg["rows_emitted"] += 1
            r += 2
        else:
            r += 1

    return records, issues, dbg

# ============================================================
# OTHER EQUIPMENT / DG
# ============================================================
def parse_other_equipment(grid: List[List[str]], table_idx: int) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, Any]]:
    records, issues = [], []
    dbg = {"table_index": table_idx, "section": "OE", "detected": False, "rows_emitted": 0}

    joined = " ".join(" ".join(r) for r in grid[:30]).upper()
    if not any(k in joined for k in ["TURBOCHARGER", "COOLERS", "A/C", "MAIN AIR COMPRESSORS", "AUXILIARY BOILER", "D/G NO"]):
        return records, issues, dbg

    dbg["detected"] = True
    item_order = 0

    # 3-column equipment layout from sample
    skip = {
        "", "TURBOCHARGER", "COOLERS", "A/C & REFR. COMPRESSORS", "AUXILIARY BOILER",
        "EXH GAS BOILER", "MAIN AIR COMPRESSORS", "DESCRIPTION", "PERIODICTLY",
        "DATE OF LAST O/H", "RUN HRS", "DATE OF LAST CLEANING", "DATE OF LAST INSPECTION", "DATE"
    }

    for r, row in enumerate(grid):
        for sec, dc, dtc, hc in [
            ("Turbocharger / Aux Boiler", 0, 2, 3),
            ("Coolers / Exh Gas Boiler", 5, 6, 7),
            ("A/C & Compressors", 10, 11, 12),
        ]:
            desc = clean_name(row[dc] if dc < len(row) else "")
            if not desc or desc.upper() in skip or not is_component_name(desc):
                continue

            item_order += 1
            d = normalize_date(row[dtc] if dtc < len(row) else "")
            h = normalize_number(row[hc] if hc < len(row) else "")
            if d["value"] or h["value"] > 0 or d["raw"] or h["raw"]:
                records.append({
                    "section": sec,
                    "description": desc,
                    "last_date": d["value"] or d["raw"],
                    "run_hrs": safe_int(h["value"]) if h["value"] > 0 else (h["raw"] or ""),
                    "source_table_index": table_idx,
                    "source_row": r,
                    "item_order": item_order,
                })
                dbg["rows_emitted"] += 1

    # DG equipment paired rows
    dg_header_row = None
    for ri, row in enumerate(grid):
        blob = " ".join(row).upper()
        if "D/G NO1" in blob or "D/G No1" in blob or ("DESCRIPTION" in blob and "D/G" in blob):
            dg_header_row = ri
            break

    if dg_header_row is not None:
        r = dg_header_row + 1
        while r < len(grid) - 1:
            row1 = grid[r]
            row2 = grid[r + 1] if r + 1 < len(grid) else []

            # left block
            for base in [(0, 2), (8, 10)]:
                desc_idx, data_start = base
                desc = clean_name(row1[desc_idx] if desc_idx < len(row1) else "")
                period = fl(row1[desc_idx + 1] if desc_idx + 1 < len(row1) else "")
                marker1 = fl(row1[data_start] if data_start < len(row1) else "")
                marker2 = fl(row2[data_start] if data_start < len(row2) else "")

                if is_component_name(desc) and marker1 == "1":
                    vals1 = [row1[data_start + i] if data_start + i < len(row1) else "" for i in range(1, 4)]
                    vals2 = [row2[data_start + i] if data_start + i < len(row2) else "" for i in range(1, 4)]

                    for dg_i, dg_label in enumerate(["D/G No1", "D/G No2", "D/G No3"]):
                        d = normalize_date(vals1[dg_i])
                        h = normalize_number(vals2[dg_i])
                        if d["value"] or h["value"] > 0 or d["raw"] or h["raw"]:
                            item_order += 1
                            records.append({
                                "section": "D/G Equipment",
                                "description": f"{desc} — {dg_label}",
                                "last_date": d["value"] or d["raw"],
                                "run_hrs": safe_int(h["value"]) if h["value"] > 0 else (h["raw"] or ""),
                                "source_table_index": table_idx,
                                "source_row": r,
                                "item_order": item_order,
                            })
                            dbg["rows_emitted"] += 1
                    if marker2 and marker2 != "2":
                        issues.append(make_issue("warning", "DG_PAIR_MARKER", f"Unexpected D/G second marker '{marker2}' under '{desc}'.", table_idx, r + 1))
            r += 2

    return records, issues, dbg

# ============================================================
# MASTER PARSER
# ============================================================
def dedupe_component_rows(rows: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    out, issues = [], []
    seen = set()
    for rec in rows:
        key = (
            rec.get("category"),
            rec.get("engine_label"),
            rec.get("unit"),
            rec.get("description"),
            rec.get("source_table_index"),
            rec.get("source_row_start"),
            rec.get("source_col_date"),
        )
        if key in seen:
            issues.append(make_issue("warning", "DUPLICATE_ROW", f"Dropped duplicate row {rec.get('description')} / {rec.get('engine_label')} / {rec.get('unit')}.", rec.get("source_table_index"), rec.get("source_row_start"), row_key(rec)))
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

    vessel_name, report_date, issues = detect_vessel_and_report_date(doc)

    me_total_hrs, me_this_month = 0.0, 0.0
    me_rows, aux_rows, oe_rows = [], [], []
    debug_tables = []
    raw_tables_preview = []

    for ti, table in enumerate(doc.tables):
        grid = rect_grid(table)
        if not grid:
            continue

        if ti == 0:
            me_total_hrs, me_this_month = detect_me_totals(grid)

        # store raw preview for first few tables
        raw_tables_preview.append({
            "table_index": ti,
            "rows": len(grid),
            "cols": max(len(r) for r in grid) if grid else 0,
            "sample": grid[:12]
        })

        me_r, me_i, me_dbg = parse_me_table(grid, ti)
        ax_r, ax_i, ax_dbg = parse_aux_table(grid, ti)
        oe_r, oe_i, oe_dbg = parse_other_equipment(grid, ti)

        me_rows.extend(me_r)
        aux_rows.extend(ax_r)
        oe_rows.extend(oe_r)
        issues.extend(me_i)
        issues.extend(ax_i)
        issues.extend(oe_i)

        debug_tables.extend([me_dbg, ax_dbg, oe_dbg])

    all_components = me_rows + aux_rows
    all_components, dd_issues = dedupe_component_rows(all_components)
    issues.extend(dd_issues)

    me_rows = [r for r in all_components if r["category"] == "MAIN_ENGINE"]
    aux_rows = [r for r in all_components if r["category"] == "AUX_ENGINE"]

    if not all_components:
        issues.append(make_issue("error", "NO_COMPONENTS", "No ME/AUX component rows were extracted."))

    debug = {
        "tables_scanned": len(doc.tables),
        "me_rows_total": len(me_rows),
        "aux_rows_total": len(aux_rows),
        "oe_rows_total": len(oe_rows),
        "issues_total": len(issues),
        "table_debug": debug_tables,
        "raw_tables_preview": raw_tables_preview[:5],
    }

    return {
        "vessel_name": vessel_name,
        "report_date": report_date,
        "me_total_hrs": me_total_hrs,
        "me_this_month": me_this_month,
        "filename": filename,
        "uploaded_at": now_iso(),
        "components": all_components,
        "me_comps": me_rows,
        "aux_comps": aux_rows,
        "other_equipment": oe_rows,
        "issues": issues,
        "debug": debug,
    }

# ============================================================
# SAVE
# ============================================================
def report_exists(vessel_name: str, report_date: str, file_hash: str) -> bool:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM reports WHERE vessel_name=? AND report_date=? AND file_hash=?", (vessel_name, report_date, file_hash))
    found = cur.fetchone() is not None
    conn.close()
    return found

def save_report(parsed: Dict[str, Any]) -> Tuple[bool, str]:
    conn = get_conn()
    try:
        if report_exists(parsed["vessel_name"], parsed["report_date"], parsed["file_hash"]):
            return False, "This report already exists in the database."

        cur = conn.cursor()
        cur.execute("""
            INSERT INTO reports(
                vessel_name, report_date, filename, file_hash, parser_version,
                me_total_hrs, me_this_month, raw_json, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            parsed["vessel_name"], parsed["report_date"], parsed["filename"], parsed["file_hash"],
            PARSER_VERSION, parsed["me_total_hrs"], parsed["me_this_month"],
            json.dumps(parsed, ensure_ascii=False, default=str), now_iso()
        ))
        report_id = cur.lastrowid

        for r in parsed["components"]:
            cur.execute("""
                INSERT INTO parsed_rows(
                    report_id, category, engine_label, unit, description,
                    periodicity_raw, periodicity_hours, last_oh_date, hrs_since, pct_used, status,
                    source_table_index, source_row_start, source_row_end, source_col_date, source_col_hours,
                    raw_date_text, raw_hours_text, confidence, issue_count, was_repaired, repair_notes, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                report_id, r["category"], r["engine_label"], r["unit"], r["description"],
                r["periodicity_raw"], r["periodicity_hours"], r["last_oh_date"], r["hrs_since"], r["pct_used"], r["status"],
                r["source_table_index"], r["source_row_start"], r["source_row_end"], r["source_col_date"], r["source_col_hours"],
                r["raw_date_text"], r["raw_hours_text"], r["confidence"], r["issue_count"], r["was_repaired"], r["repair_notes"], now_iso()
            ))

        for i in parsed["issues"]:
            cur.execute("""
                INSERT INTO parse_issues(
                    report_id, severity, issue_code, message, row_key, table_index, row_index, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                report_id, i.get("severity",""), i.get("issue_code",""), i.get("message",""),
                i.get("row_key",""), i.get("table_index"), i.get("row_index"), now_iso()
            ))

        conn.commit()
        return True, f"Saved successfully. Report ID {report_id}."
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
# DATAFRAME BUILDERS
# ============================================================
def build_component_df(records: List[Dict[str, Any]], include_trace=False, sort_mode="source") -> pd.DataFrame:
    base_cols = ["Status","Component","Engine","Unit","Periodicity","Last O/H","Hrs Since","Used %","Confidence","Issues","Repaired"]
    if include_trace:
        base_cols += ["Table","Rows","Raw Date","Raw Hrs","Repair Notes"]

    if not records:
        return pd.DataFrame(columns=base_cols)

    df = pd.DataFrame(records)

    for c, default in {
        "status":"NO DATA","description":"","engine_label":"","unit":"",
        "periodicity_hours":0.0,"periodicity_raw":"","last_oh_date":"",
        "hrs_since":0.0,"pct_used":0.0,"confidence":0.0,"issue_count":0,
        "was_repaired":0,"repair_notes":"","source_table_index":"",
        "source_row_start":"","source_row_end":"","raw_date_text":"","raw_hours_text":"",
        "component_order":0,"engine_order":0,"unit_order":0
    }.items():
        if c not in df.columns:
            df[c] = default

    if sort_mode == "priority":
        df["_s"] = df["status"].map(lambda x: STATUS_ORDER.get(str(x), 9))
        df["_p"] = df["pct_used"].map(safe_float)
        df = df.sort_values(["_s", "_p"], ascending=[True, False])
    elif sort_mode == "alphabetical":
        df["_d"] = df["description"].astype(str).str.upper()
        df["_e"] = df["engine_label"].map(lambda x: ENGINE_ORDER.get(str(x), 999))
        df["_u"] = df["unit"].astype(str).map(cyl_num)
        df = df.sort_values(["_d", "_e", "_u"])
    else:
        df = df.sort_values(["source_table_index", "component_order", "engine_order", "unit_order", "source_row_start"])

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
    out["Used %"] = [round(safe_float(x) * 100, 1) if safe_float(x) > 0 else 0.0 for x in df["pct_used"]]
    out["Confidence"] = [round(safe_float(x) * 100, 0) for x in df["confidence"]]
    out["Issues"] = [safe_int(x) for x in df["issue_count"]]
    out["Repaired"] = ["Yes" if safe_int(x) == 1 else "—" for x in df["was_repaired"]]

    if include_trace:
        out["Table"] = df["source_table_index"].astype(str)
        out["Rows"] = [f"{a}-{b}" for a, b in zip(df["source_row_start"], df["source_row_end"])]
        out["Raw Date"] = df["raw_date_text"].astype(str)
        out["Raw Hrs"] = df["raw_hours_text"].astype(str)
        out["Repair Notes"] = df["repair_notes"].astype(str)

    return out

def build_issues_df(issues: List[Dict[str, Any]]) -> pd.DataFrame:
    if not issues:
        return pd.DataFrame(columns=["Severity","Code","Message","Table","Row","Row Key"])
    df = pd.DataFrame(issues)
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
    "Component": st.column_config.TextColumn("Component", width=260),
    "Engine": st.column_config.TextColumn("Engine", width=82),
    "Unit": st.column_config.TextColumn("Unit", width=78),
    "Periodicity": st.column_config.TextColumn("Periodicity", width=105),
    "Last O/H": st.column_config.TextColumn("Last O/H", width=120),
    "Hrs Since": st.column_config.TextColumn("Hrs Since", width=105),
    "Used %": st.column_config.ProgressColumn("Used %", min_value=0, max_value=150, format="%.1f%%", width=130),
    "Confidence": st.column_config.ProgressColumn("Confidence", min_value=0, max_value=100, format="%d%%", width=120),
    "Issues": st.column_config.NumberColumn("Issues", width=70),
    "Repaired": st.column_config.TextColumn("Repaired", width=80),
    "Table": st.column_config.TextColumn("Table", width=60),
    "Rows": st.column_config.TextColumn("Rows", width=74),
    "Raw Date": st.column_config.TextColumn("Raw Date", width=125),
    "Raw Hrs": st.column_config.TextColumn("Raw Hrs", width=110),
    "Repair Notes": st.column_config.TextColumn("Repair Notes", width=340),
}

def show_df(df: pd.DataFrame, height=None):
    h = height or min(850, 38 * len(df) + 44)
    cfg = {k: v for k, v in TABLE_CONFIG.items() if k in df.columns}
    st.dataframe(df, use_container_width=True, hide_index=True, height=h, column_config=cfg)

# ============================================================
# SESSION STATE
# ============================================================
if "parsed_reports" not in st.session_state:
    st.session_state.parsed_reports = []
if "active_hash" not in st.session_state:
    st.session_state.active_hash = None
if "save_feedback" not in st.session_state:
    st.session_state.save_feedback = {}

# ============================================================
# HEADER
# ============================================================
st.markdown("""
<div class="hero-k">Fleet Running Hours Monitor</div>
<div class="hero-h">TEC‑004 Operational Console v10.10</div>
<div class="hero-s">
Built for legacy TEC‑004 .doc uploads with paired-row Main Engine parsing, compact Auxiliary Engine parsing,
traceable normalization, matrix preview before save, and SQLite-backed report history.
</div>
<div class="hero-rule"></div>
""", unsafe_allow_html=True)

# ============================================================
# UPLOAD
# ============================================================
with st.container():
    st.markdown('<div class="panel upload-panel">', unsafe_allow_html=True)
    c1, c2 = st.columns([2.1, 1.0], gap="large")
    with c1:
        uploaded_files = st.file_uploader(
            "Upload TEC-004 reports",
            type=["doc"],
            accept_multiple_files=True,
            label_visibility="collapsed"
        )
        st.markdown('<div class="muted">Accepted input: legacy <b>.doc</b> TEC‑004 running-hours reports. Files are converted to DOCX, parsed, previewed, and only then saved.</div>', unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class="muted">
        <b>Parser focus</b><br>Main Engine paired rows · AUX compact paired rows<br><br>
        <b>Preview priority</b><br>Matrix visibility before persistence<br><br>
        <b>Storage</b><br>SQLite report ledger
        </div>
        """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

if uploaded_files:
    existing_hashes = {r.get("file_hash") for r in st.session_state.parsed_reports}
    new_files = [f for f in uploaded_files if md5_bytes(f.getvalue()) not in existing_hashes]
    if new_files:
        progress = st.progress(0)
        for i, up in enumerate(new_files, start=1):
            raw = up.getvalue()
            file_hash = md5_bytes(raw)
            with st.spinner(f"Converting and parsing {up.name}..."):
                try:
                    docx_bytes = convert_doc_to_docx(raw)
                    parsed = parse_docx_bytes(docx_bytes, filename=up.name)
                    parsed["file_hash"] = file_hash
                    st.session_state.parsed_reports.append(parsed)
                    if st.session_state.active_hash is None:
                        st.session_state.active_hash = file_hash
                except Exception as e:
                    st.session_state.parsed_reports.append({
                        "vessel_name": "UNKNOWN",
                        "report_date": "",
                        "me_total_hrs": 0,
                        "me_this_month": 0,
                        "filename": up.name,
                        "file_hash": file_hash,
                        "uploaded_at": now_iso(),
                        "components": [],
                        "me_comps": [],
                        "aux_comps": [],
                        "other_equipment": [],
                        "issues": [make_issue("error", "PARSE_FAILURE", f"{up.name}: {e}")],
                        "debug": {}
                    })
                    if st.session_state.active_hash is None:
                        st.session_state.active_hash = file_hash
            progress.progress(i / len(new_files))
        progress.empty()

reports = st.session_state.parsed_reports

st.markdown("## Report Queue")
if not reports:
    st.info("Upload one or more TEC‑004 .doc reports to begin.")
    st.stop()

queue = []
for r in reports:
    errs = sum(1 for x in r["issues"] if x.get("severity") == "error")
    warns = sum(1 for x in r["issues"] if x.get("severity") == "warning")
    repaired = sum(1 for x in r["components"] if safe_int(x.get("was_repaired")) == 1)
    queue.append({
        "Active": "●" if r["file_hash"] == st.session_state.active_hash else "",
        "Filename": r["filename"],
        "Vessel": r["vessel_name"],
        "Report Date": r["report_date"] or "—",
        "Rows": len(r["components"]),
        "Warnings": warns,
        "Errors": errs,
        "Repaired": repaired,
        "Hash": r["file_hash"][:18] + "…",
    })

queue_df = pd.DataFrame(queue)
st.dataframe(
    queue_df,
    use_container_width=True,
    hide_index=True,
    height=min(340, 38 * len(queue_df) + 44),
    column_config={
        "Active": st.column_config.TextColumn("Active", width=55),
        "Filename": st.column_config.TextColumn("Filename", width=240),
        "Vessel": st.column_config.TextColumn("Vessel", width=150),
        "Report Date": st.column_config.TextColumn("Report Date", width=115),
        "Rows": st.column_config.NumberColumn("Rows", width=70),
        "Warnings": st.column_config.NumberColumn("Warnings", width=80),
        "Errors": st.column_config.NumberColumn("Errors", width=70),
        "Repaired": st.column_config.NumberColumn("Repaired", width=80),
        "Hash": st.column_config.TextColumn("Hash", width=180),
    }
)

selector_map = {f"{r['filename']} | {r['vessel_name']} | {r['report_date'] or '—'}": r["file_hash"] for r in reports}
selected_label = st.selectbox(
    "Active report",
    list(selector_map.keys()),
    index=list(selector_map.values()).index(st.session_state.active_hash) if st.session_state.active_hash in selector_map.values() else 0
)
st.session_state.active_hash = selector_map[selected_label]
active = next(r for r in reports if r["file_hash"] == st.session_state.active_hash)

all_rows = active["components"]
me_rows = active["me_comps"]
aux_rows = active["aux_comps"]
oe_rows = active["other_equipment"]
issues = active["issues"]
debug = active["debug"]

n_od = sum(1 for x in all_rows if x.get("status") == "OVERDUE")
n_hp = sum(1 for x in all_rows if x.get("status") == "HIGH PRIORITY")
n_ok = sum(1 for x in all_rows if x.get("status") == "OK")
n_warn = sum(1 for x in issues if x.get("severity") == "warning")
n_err = sum(1 for x in issues if x.get("severity") == "error")

st.markdown(f"""
<div class="metric-grid">
  <div class="metric-card b"><div class="metric-v">{active["vessel_name"]}</div><div class="metric-l">Vessel</div></div>
  <div class="metric-card"><div class="metric-v">{active["report_date"] or "—"}</div><div class="metric-l">Report Date</div></div>
  <div class="metric-card"><div class="metric-v">{safe_int(active["me_total_hrs"]):,}</div><div class="metric-l">ME Total Hrs</div></div>
  <div class="metric-card"><div class="metric-v">{safe_int(active["me_this_month"]):,}</div><div class="metric-l">ME This Month</div></div>
  <div class="metric-card a"><div class="metric-v">{len(all_rows)}</div><div class="metric-l">Parsed Rows</div></div>
  <div class="metric-card r"><div class="metric-v">{n_od}</div><div class="metric-l">Overdue</div></div>
  <div class="metric-card g"><div class="metric-v">{n_ok}</div><div class="metric-l">OK</div></div>
</div>
""", unsafe_allow_html=True)

if n_err:
    st.markdown(f'<div class="banner err"><b>{n_err}</b> parser errors detected. Review the matrix and issue log before saving. [file:1]</div>', unsafe_allow_html=True)
elif n_warn:
    st.markdown(f'<div class="banner warn"><b>{n_warn}</b> parser warnings detected. Rows were still emitted when meaningful raw payload existed, which is important for noisy values like bracketed hours or invalid date tokens in the sample. [file:1]</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="banner ok">No parser issues were logged for this report. [file:1]</div>', unsafe_allow_html=True)

# ============================================================
# ACTIONS
# ============================================================
a1, a2, a3, a4 = st.columns([1.2, 1.3, 1.6, 2.7])
with a1:
    ready_to_save = st.checkbox("Review complete", value=False)
with a2:
    if st.button("Save active report", disabled=not ready_to_save):
        ok, msg = save_report(active)
        st.session_state.save_feedback[active["file_hash"]] = (ok, msg)
with a3:
    export_df = build_component_df(all_rows, include_trace=True, sort_mode="source")
    st.download_button(
        "Download active CSV",
        data=export_df.to_csv(index=False).encode("utf-8"),
        file_name=f"{Path(active['filename']).stem}_v1010_preview.csv",
        mime="text/csv"
    )
with a4:
    st.markdown(f'<div class="micro">Parser {PARSER_VERSION} • File {active["filename"]} • Hash {active["file_hash"][:18]}…</div>', unsafe_allow_html=True)

fb = st.session_state.save_feedback.get(active["file_hash"])
if fb:
    if fb[0]:
        st.success(fb[1])
    else:
        st.warning(fb[1])

# ============================================================
# TABS
# ============================================================
tab_me, tab_aux, tab_oe, tab_issues, tab_debug, tab_history = st.tabs([
    "Main Engine", "Auxiliary Engines", "Other Equipment", "Parse Issues", "Parser Debug", "Saved Reports"
])

with tab_me:
    st.markdown('<div class="section-bar"><div class="section-icon">⚙</div><div><div class="section-title">Main Engine Matrix</div><div class="section-sub">paired date/hour rows under cylinder headers</div></div><div class="section-rule"></div><div class="section-badge">rows: %d</div></div>' % len(me_rows), unsafe_allow_html=True)
    f1, f2, f3, f4 = st.columns([2.1, 1.9, 1.4, 2.2])
    with f1:
        me_comp = st.selectbox("Component", ["All"] + sorted({r["description"] for r in me_rows}), key="me_comp")
    with f2:
        me_status = st.selectbox("Status", ["All","Overdue only","High Priority +","Issue rows only","Repaired only"], key="me_status")
    with f3:
        me_trace = st.checkbox("Show trace", value=True, key="me_trace")
    with f4:
        me_sort = st.radio("Sort", ["Source order","Priority","Alphabetical"], horizontal=True, key="me_sort")

    me_view = me_rows[:]
    if me_comp != "All":
        me_view = [r for r in me_view if r["description"] == me_comp]
    if me_status == "Overdue only":
        me_view = [r for r in me_view if r["status"] == "OVERDUE"]
    elif me_status == "High Priority +":
        me_view = [r for r in me_view if r["status"] in ("OVERDUE","HIGH PRIORITY")]
    elif me_status == "Issue rows only":
        me_view = [r for r in me_view if safe_int(r["issue_count"]) > 0]
    elif me_status == "Repaired only":
        me_view = [r for r in me_view if safe_int(r["was_repaired"]) == 1]

    sort_mode = {"Source order":"source","Priority":"priority","Alphabetical":"alphabetical"}[me_sort]
    show_df(build_component_df(me_view, include_trace=me_trace, sort_mode=sort_mode))

with tab_aux:
    st.markdown('<div class="section-bar"><div class="section-icon">🔩</div><div><div class="section-title">Auxiliary Engine Matrix</div><div class="section-sub">compact paired-row parser with grouped fallback</div></div><div class="section-rule"></div><div class="section-badge">rows: %d</div></div>' % len(aux_rows), unsafe_allow_html=True)
    f1, f2, f3, f4, f5 = st.columns([1.4, 2.0, 1.9, 1.4, 2.2])
    with f1:
        ax_eng = st.selectbox("Engine", ["All"] + sorted({r["engine_label"] for r in aux_rows}), key="ax_eng")
    with f2:
        ax_comp = st.selectbox("Component", ["All"] + sorted({r["description"] for r in aux_rows}), key="ax_comp")
    with f3:
        ax_status = st.selectbox("Status", ["All","Overdue only","High Priority +","Issue rows only","Repaired only"], key="ax_status")
    with f4:
        ax_trace = st.checkbox("Show trace", value=True, key="ax_trace")
    with f5:
        ax_sort = st.radio("Sort", ["Source order","Priority","Alphabetical"], horizontal=True, key="ax_sort")

    ax_view = aux_rows[:]
    if ax_eng != "All":
        ax_view = [r for r in ax_view if r["engine_label"] == ax_eng]
    if ax_comp != "All":
        ax_view = [r for r in ax_view if r["description"] == ax_comp]
    if ax_status == "Overdue only":
        ax_view = [r for r in ax_view if r["status"] == "OVERDUE"]
    elif ax_status == "High Priority +":
        ax_view = [r for r in ax_view if r["status"] in ("OVERDUE","HIGH PRIORITY")]
    elif ax_status == "Issue rows only":
        ax_view = [r for r in ax_view if safe_int(r["issue_count"]) > 0]
    elif ax_status == "Repaired only":
        ax_view = [r for r in ax_view if safe_int(r["was_repaired"]) == 1]

    sort_mode = {"Source order":"source","Priority":"priority","Alphabetical":"alphabetical"}[ax_sort]
    show_df(build_component_df(ax_view, include_trace=ax_trace, sort_mode=sort_mode))

with tab_oe:
    st.markdown('<div class="section-bar"><div class="section-icon">🛠</div><div><div class="section-title">Other Equipment</div><div class="section-sub">turbocharger · coolers · A/C · D/G equipment</div></div><div class="section-rule"></div><div class="section-badge">rows: %d</div></div>' % len(oe_rows), unsafe_allow_html=True)
    if not oe_rows:
        st.info("No other equipment rows found.")
    else:
        oe_df = pd.DataFrame(oe_rows).sort_values(["section","item_order"])
        st.dataframe(
            oe_df[["section","description","last_date","run_hrs","source_table_index","source_row"]],
            use_container_width=True,
            hide_index=True,
            height=min(780, 38 * len(oe_df) + 44),
            column_config={
                "section": st.column_config.TextColumn("Section", width=220),
                "description": st.column_config.TextColumn("Description", width=320),
                "last_date": st.column_config.TextColumn("Last Date / O/H", width=140),
                "run_hrs": st.column_config.TextColumn("Run Hrs", width=120),
                "source_table_index": st.column_config.TextColumn("Table", width=70),
                "source_row": st.column_config.TextColumn("Row", width=70),
            }
        )

with tab_issues:
    st.markdown('<div class="section-bar"><div class="section-icon">⚠</div><div><div class="section-title">Parse Issues</div><div class="section-sub">validation warnings and extraction anomalies</div></div><div class="section-rule"></div><div class="section-badge">issues: %d</div></div>' % len(issues), unsafe_allow_html=True)
    idf = build_issues_df(issues)
    if idf.empty:
        st.success("No issues found.")
    else:
        sev = st.multiselect("Severity", ["error","warning","info"], default=["error","warning","info"])
        if sev:
            idf = idf[idf["Severity"].isin(sev)]
        st.dataframe(
            idf, use_container_width=True, hide_index=True,
            height=min(700, 38 * len(idf) + 44),
            column_config={
                "Severity": st.column_config.TextColumn("Severity", width=90),
                "Code": st.column_config.TextColumn("Code", width=160),
                "Message": st.column_config.TextColumn("Message", width=520),
                "Table": st.column_config.TextColumn("Table", width=65),
                "Row": st.column_config.TextColumn("Row", width=65),
                "Row Key": st.column_config.TextColumn("Row Key", width=320),
            }
        )

with tab_debug:
    st.markdown('<div class="section-bar"><div class="section-icon">🧪</div><div><div class="section-title">Parser Debug</div><div class="section-sub">raw section detection and table previews</div></div><div class="section-rule"></div><div class="section-badge">diagnostics</div></div>', unsafe_allow_html=True)
    d1, d2, d3, d4 = st.columns(4)
    with d1:
        st.metric("Tables scanned", debug.get("tables_scanned", 0))
    with d2:
        st.metric("ME rows", debug.get("me_rows_total", 0))
    with d3:
        st.metric("AUX rows", debug.get("aux_rows_total", 0))
    with d4:
        st.metric("OE rows", debug.get("oe_rows_total", 0))

    tdbg = pd.DataFrame(debug.get("table_debug", []))
    if not tdbg.empty:
        st.markdown("### Section detection")
        st.dataframe(tdbg, use_container_width=True, hide_index=True)

    previews = debug.get("raw_tables_preview", [])
    if previews:
        st.markdown("### Raw table preview")
        for item in previews:
            with st.expander(f"Table {item['table_index']} · {item['rows']} rows · {item['cols']} cols"):
                sample_df = pd.DataFrame(item["sample"])
                st.dataframe(sample_df, use_container_width=True, hide_index=True)

with tab_history:
    st.markdown('<div class="section-bar"><div class="section-icon">🗄</div><div><div class="section-title">Saved Reports</div><div class="section-sub">sqlite-backed report history</div></div><div class="section-rule"></div><div class="section-badge">history</div></div>', unsafe_allow_html=True)
    hist = load_recent_reports(30)
    if hist.empty:
        st.info("No saved reports in the database yet.")
    else:
        st.dataframe(
            hist, use_container_width=True, hide_index=True,
            height=min(700, 38 * len(hist) + 44),
            column_config={
                "id": st.column_config.NumberColumn("ID", width=60),
                "vessel_name": st.column_config.TextColumn("Vessel", width=150),
                "report_date": st.column_config.TextColumn("Report Date", width=120),
                "filename": st.column_config.TextColumn("Filename", width=260),
                "parser_version": st.column_config.TextColumn("Parser", width=160),
                "me_total_hrs": st.column_config.NumberColumn("ME Total", width=100, format="%d"),
                "me_this_month": st.column_config.NumberColumn("ME Month", width=100, format="%d"),
                "created_at": st.column_config.TextColumn("Saved At", width=180),
                "parsed_rows": st.column_config.NumberColumn("Rows", width=70),
                "errors": st.column_config.NumberColumn("Errors", width=70),
                "warnings": st.column_config.NumberColumn("Warnings", width=85),
            }
        )
