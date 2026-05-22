import streamlit as st
st.set_page_config(
    page_title="TEC-004 Running Hours Parser",
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
from typing import List, Dict, Any, Tuple, Optional

import pandas as pd

# ============================================================
# DESIGN
# ============================================================
st.markdown("""
<style>
:root {
  --bg: #071018;
  --bg2: #0b1724;
  --bg3: #102030;
  --line: #1f3347;
  --line2: #294764;
  --text: #d9e8f5;
  --muted: #93abc2;
  --soft: #5e768e;
  --gold: #d2a532;
  --gold2: #f0c24f;
  --red: #ef5350;
  --amber: #ffb74d;
  --green: #66bb6a;
  --blue: #64b5f6;
  --radius: 12px;
}

html, body, [class*="css"] {
  background: var(--bg) !important;
  color: var(--text) !important;
}

.main, .main > div, .block-container {
  background: var(--bg) !important;
}

.block-container {
  max-width: 100% !important;
  padding-top: 1.25rem !important;
  padding-bottom: 3rem !important;
  padding-left: 2rem !important;
  padding-right: 2rem !important;
}

[data-testid="collapsedControl"] { display: none !important; }

h1, h2, h3 {
  color: var(--text) !important;
}

[data-testid="stFileUploadDropzone"] {
  background: linear-gradient(180deg, rgba(210,165,50,.09), rgba(100,181,246,.05)) !important;
  border: 1.5px dashed var(--gold) !important;
  border-radius: 16px !important;
}

[data-testid="stMetric"] {
  background: var(--bg3) !important;
  border: 1px solid var(--line) !important;
  border-radius: var(--radius) !important;
}

[data-testid="stDataFrame"] {
  border: 1px solid var(--line) !important;
  border-radius: var(--radius) !important;
  overflow: hidden !important;
}

.stButton > button {
  background: linear-gradient(135deg, var(--gold2), var(--gold)) !important;
  color: #0b1117 !important;
  border: none !important;
  border-radius: 10px !important;
  font-weight: 700 !important;
  padding: 0.6rem 1.2rem !important;
}

.stDownloadButton > button {
  background: var(--bg3) !important;
  color: var(--text) !important;
  border: 1px solid var(--line2) !important;
  border-radius: 10px !important;
  font-weight: 600 !important;
}

div[data-baseweb="select"] > div {
  background: var(--bg3) !important;
  border-color: var(--line2) !important;
}

section[data-testid="stSidebar"] {
  display: none !important;
}

.hr-title {
  font-size: 0.75rem;
  text-transform: uppercase;
  letter-spacing: .14em;
  color: var(--gold2);
  margin-bottom: .2rem;
}

.kpi-wrap {
  display: grid;
  grid-template-columns: repeat(6, 1fr);
  gap: .8rem;
  margin: .8rem 0 1rem;
}

.kpi {
  background: var(--bg3);
  border: 1px solid var(--line);
  border-radius: var(--radius);
  padding: .9rem 1rem;
}

.kpi-v {
  font-size: 1.6rem;
  font-weight: 800;
  color: var(--text);
  line-height: 1.1;
}

.kpi-l {
  color: var(--soft);
  font-size: .62rem;
  text-transform: uppercase;
  letter-spacing: .16em;
  margin-top: .35rem;
}

.tag {
  display: inline-block;
  padding: .2rem .55rem;
  border-radius: 999px;
  font-size: .72rem;
  font-weight: 700;
  border: 1px solid var(--line2);
}

.tag-red { color: #ffd4d4; background: rgba(239,83,80,.14); border-color: rgba(239,83,80,.35); }
.tag-amber { color: #ffe4bf; background: rgba(255,183,77,.14); border-color: rgba(255,183,77,.35); }
.tag-green { color: #d4f5d6; background: rgba(102,187,106,.14); border-color: rgba(102,187,106,.35); }
.tag-blue { color: #d7ebff; background: rgba(100,181,246,.14); border-color: rgba(100,181,246,.35); }

.small-note {
  color: var(--muted);
  font-size: .85rem;
  line-height: 1.55;
}

.section-rule {
  height: 1px;
  background: linear-gradient(90deg, var(--gold) 0%, var(--line) 30%, transparent 100%);
  margin: .55rem 0 1.4rem;
}

.warn-box {
  background: rgba(255,183,77,.10);
  border: 1px solid rgba(255,183,77,.35);
  border-radius: var(--radius);
  padding: .8rem 1rem;
  margin-bottom: .7rem;
}

.err-box {
  background: rgba(239,83,80,.10);
  border: 1px solid rgba(239,83,80,.35);
  border-radius: var(--radius);
  padding: .8rem 1rem;
  margin-bottom: .7rem;
}

.ok-box {
  background: rgba(102,187,106,.10);
  border: 1px solid rgba(102,187,106,.35);
  border-radius: var(--radius);
  padding: .8rem 1rem;
  margin-bottom: .7rem;
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# DB
# ============================================================
DB_PATH = "tec004_app.db"

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
        source_row_start INTEGER,
        source_row_end INTEGER,
        source_col_date INTEGER,
        source_col_hours INTEGER,
        raw_date_text TEXT,
        raw_hours_text TEXT,
        confidence REAL,
        issue_count INTEGER DEFAULT 0,
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
PARSER_VERSION = "tec004_v1.2_schema_safe"

STATUS_ORDER = {
    "OVERDUE": 0,
    "HIGH PRIORITY": 1,
    "OK": 2,
    "NO DATA": 3
}

def now_iso():
    return datetime.utcnow().isoformat()

def md5_bytes(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()

def first_line(text: str) -> str:
    if text is None:
        return ""
    txt = str(text).replace("\x07", "").replace("\xa0", " ").replace("\t", " ")
    parts = re.split(r'[\r\n\x0b]+', txt)
    for p in parts:
        p = re.sub(r'\s+', ' ', p).strip()
        if p:
            return p
    return ""

def normalize_whitespace(text: str) -> str:
    return re.sub(r'\s+', ' ', str(text or '')).strip()

def clean_name(text: str) -> str:
    t = first_line(text)
    t = re.sub(r'(?i)ALEXIS\s*Date?', '', t)
    t = re.sub(r'(?i)Page\s*\d+\s*of\s*\d+', '', t)
    t = re.sub(r'(?i)MTS Marine Ltd.*', '', t)
    return normalize_whitespace(t)

def is_component_name(name: str) -> bool:
    u = normalize_whitespace(str(name or "")).upper()
    if not u or len(u) < 2:
        return False
    bad_terms = [
        "DESCRIPTION", "REMARKS", "COMPONENT", "PERIODICITY", "PERIODICTLY",
        "DATE OF LAST", "RUNNING HOURS", "TYPE:", "MAIN ENGINE", "AUX. ENGINE",
        "AUX ENGINE", "TURBOCHARGER", "CYL. NO", "TOTAL HOURS", "HOURS THIS MONTH",
        "SERIAL NR", "VESSEL", "COPY", "CHIEF ENGINEER", "NOTE 1", "NOTE 2"
    ]
    if any(bt in u for bt in bad_terms):
        return False
    if re.fullmatch(r'[\d\s,./:-]+', u):
        return False
    if len(u) > 80:
        return False
    return bool(re.search(r'[A-Z]', u))

def parse_date_value(text: str) -> str:
    s = first_line(text).replace("[", "").replace("]", "").strip()
    if not s:
        return ""
    if s.upper() in {"N/A", "NA", "-", "CENTRAL", "COOLER"}:
        return ""
    if len(s) > 24:
        return ""
    if re.fullmatch(r'\d+', s):
        return ""
    return s

def parse_number(text: str) -> float:
    s = first_line(text).upper().replace("[", "").replace("]", "").strip()
    if not s or s in {"", "-", "N/A", "NA", "CENTRAL", "COOLER"}:
        return 0.0
    if any(k in s for k in ["MONTH", "YEAR", "WEEK", "DAY", "OBSERVATION", "OBS"]):
        return 0.0
    s = re.sub(r'([,.])\s+', r'\1', s)
    m = re.search(r'\d[\d,.]*', s)
    if not m:
        return 0.0
    block = m.group()
    sep = max(block.rfind("."), block.rfind(","))
    if sep > 0 and len(block) - sep == 4:
        block = re.sub(r'[,.]', '', block)
    elif sep > 0:
        block = re.sub(r'[,.]', '', block[:sep])
    else:
        block = re.sub(r'[,.]', '', block)
    try:
        return float(block)
    except Exception:
        return 0.0

def periodicity_value(raw: str) -> float:
    return parse_number(raw)

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

def confidence_score(date_txt: str, hrs_txt: str, period_raw: str, issues: List[Dict[str, Any]]) -> float:
    score = 1.0
    if not parse_date_value(date_txt):
        score -= 0.15
    if parse_number(hrs_txt) <= 0:
        score -= 0.15
    if periodicity_value(period_raw) <= 0:
        score -= 0.10
    score -= min(0.50, 0.08 * len(issues))
    return max(0.0, round(score, 2))

def row_key(rec: Dict[str, Any]) -> str:
    return " | ".join([
        str(rec.get("category", "")),
        str(rec.get("engine_label", "")),
        str(rec.get("unit", "")),
        str(rec.get("description", "")),
        str(rec.get("source_table_index", "")),
        str(rec.get("source_row_start", "")),
        str(rec.get("source_col_date", "")),
    ])

def make_issue(severity: str, code: str, message: str, table_idx: int = None, row_idx: int = None, rkey: str = "") -> Dict[str, Any]:
    return {
        "severity": severity,
        "issue_code": code,
        "message": message,
        "table_index": table_idx,
        "row_index": row_idx,
        "row_key": rkey,
    }

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
        "source_row_start": "",
        "source_row_end": "",
        "source_col_date": "",
        "source_col_hours": "",
        "raw_date_text": "",
        "raw_hours_text": "",
        "confidence": 0.0,
        "issue_count": 0,
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

def normalize_parsed_payload(parsed: dict) -> dict:
    parsed = parsed or {}
    parsed.setdefault("vessel_name", "UNKNOWN")
    parsed.setdefault("report_date", "")
    parsed.setdefault("me_total_hrs", 0)
    parsed.setdefault("me_this_month", 0)
    parsed["me_comps"] = normalize_rows_payload(parsed.get("me_comps", []))
    parsed["aux_comps"] = normalize_rows_payload(parsed.get("aux_comps", []))
    parsed["components"] = normalize_rows_payload(parsed.get("components", []))
    parsed.setdefault("other_equipment", [])
    if isinstance(parsed.get("other_equipment"), pd.DataFrame):
        parsed["other_equipment"] = parsed["other_equipment"].to_dict("records")
    if not isinstance(parsed.get("other_equipment"), list):
        parsed["other_equipment"] = []
    parsed.setdefault("issues", [])
    if not isinstance(parsed.get("issues"), list):
        parsed["issues"] = []
    parsed.setdefault("uploaded_at", now_iso())
    return parsed

def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found. Add libreoffice to your environment/packages.")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as tf:
        tf.write(raw)
        src = tf.name
    outdir = tempfile.mkdtemp(prefix="lo_")
    out = os.path.join(outdir, Path(src).stem + ".docx")
    profile = f"file:///tmp/lo_profile_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        r = subprocess.run(
            [
                soffice, "--headless", "--norestore", "--nofirststartwizard",
                f"-env:UserInstallation={profile}",
                "--convert-to", "docx",
                src,
                "--outdir", outdir
            ],
            capture_output=True,
            timeout=120
        )
        if not os.path.exists(out):
            stderr = r.stderr.decode("utf-8", "ignore")[:500]
            stdout = r.stdout.decode("utf-8", "ignore")[:500]
            raise RuntimeError(f"LibreOffice conversion failed. stdout={stdout} stderr={stderr}")
        with open(out, "rb") as f:
            return f.read()
    finally:
        try:
            if os.path.exists(src):
                os.unlink(src)
        except Exception:
            pass
        try:
            if os.path.exists(out):
                os.unlink(out)
        except Exception:
            pass
        shutil.rmtree(outdir, ignore_errors=True)

def rect_grid(table) -> List[List[str]]:
    grid = []
    if not table.rows:
        return grid
    max_cols = max(len(r.cells) for r in table.rows)
    for row in table.rows:
        cells = []
        for cell in row.cells:
            raw = re.sub(r'[\x0b\r]', '\n', cell.text).replace('\x07', '')
            lines = [normalize_whitespace(x) for x in raw.split("\n") if normalize_whitespace(x)]
            cells.append(lines[0] if lines else "")
        while len(cells) < max_cols:
            cells.append("")
        grid.append(cells)
    return grid

# ============================================================
# PARSER
# ============================================================
def detect_vessel_and_date(doc) -> Tuple[str, str, List[Dict[str, Any]]]:
    vessel = "UNKNOWN"
    report_date = ""
    issues = []

    joined = "\n".join([normalize_whitespace(p.text) for p in doc.paragraphs if normalize_whitespace(p.text)])
    if joined:
        m_v = re.search(r"Vessel[’'`s\s]*Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+)", joined, re.I)
        if m_v:
            vessel = clean_name(m_v.group(1))
        m_d = re.search(r"Date\s*:\s*([^\n]+)", joined, re.I)
        if m_d:
            report_date = parse_date_value(m_d.group(1))

    if vessel == "UNKNOWN":
        for p in doc.paragraphs:
            txt = normalize_whitespace(p.text)
            m_v = re.search(r"Vessel[’'`s\s]*Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+)", txt, re.I)
            if m_v:
                vessel = clean_name(m_v.group(1))
                break

    if not report_date:
        for p in doc.paragraphs:
            txt = normalize_whitespace(p.text)
            m_d = re.search(r"Date\s*:\s*(.+)", txt, re.I)
            if m_d:
                report_date = parse_date_value(m_d.group(1))
                if report_date:
                    break

    if vessel == "UNKNOWN":
        issues.append(make_issue("warning", "VESSEL_NOT_FOUND", "Could not extract vessel name from paragraphs."))
    if not report_date:
        issues.append(make_issue("warning", "REPORT_DATE_NOT_FOUND", "Could not extract report date from paragraphs."))

    return vessel, report_date, issues

def detect_me_totals(grid: List[List[str]]) -> Tuple[float, float]:
    me_total = 0
    me_month = 0
    flat = " | ".join([" | ".join(row) for row in grid[:3]])
    m1 = re.search(r"Total Running Hours[\s:ǀ|]+([\d,]+)", flat, re.I)
    m2 = re.search(r"This Month[\s:]+([\d,]+)", flat, re.I)
    if m1:
        me_total = parse_number(m1.group(1))
    if m2:
        me_month = parse_number(m2.group(1))
    return me_total, me_month

def parse_me_table(grid: List[List[str]], table_idx: int) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    records = []
    issues = []

    header_text = " ".join(" ".join(r) for r in grid[:3]).upper()
    if "MAIN ENGINE" not in header_text:
        return records, issues

    if len(grid) < 3:
        issues.append(make_issue("warning", "ME_TABLE_TOO_SHORT", "Main Engine table too short to parse.", table_idx))
        return records, issues

    period_col = 1
    remarks_col = len(grid[0]) - 1

    for c in range(len(grid[0])):
        h = " ".join(str(grid[r][c]).upper() for r in range(min(4, len(grid))))
        if "REMARK" in h:
            remarks_col = c

    marker_col = period_col + 1
    first_cyl_col = period_col + 2
    actual_cyls = max(1, min(7, remarks_col - first_cyl_col))
    if actual_cyls <= 0:
        actual_cyls = max(1, min(7, len(grid[0]) - first_cyl_col - 1))

    end = len(grid)
    for r, row in enumerate(grid):
        joined = " ".join(str(x).upper() for x in row)
        if any(stop in joined for stop in ["NOTE 1", "TURBOCHARGER", "AUX. ENGINE", "AUX ENGINE"]):
            end = r
            break

    r = 1
    while r < end - 1:
        current = grid[r]
        nxt = grid[r + 1] if r + 1 < end else []
        name = clean_name(current[0] if len(current) > 0 else "")
        period_raw = current[period_col] if len(current) > period_col else ""
        period = periodicity_value(period_raw)
        marker = normalize_whitespace(current[marker_col] if len(current) > marker_col else "")

        if is_component_name(name) and marker == "1":
            pair_marker = normalize_whitespace(nxt[marker_col] if len(nxt) > marker_col else "")
            pair_issue_list = []
            if pair_marker != "2":
                pair_issue_list.append(make_issue(
                    "warning", "ME_PAIR_MARKER_MISSING",
                    f"Expected paired marker '2' below component '{name}', found '{pair_marker or '(blank)'}'.",
                    table_idx, r
                ))

            for cyl in range(1, actual_cyls + 1):
                ci = first_cyl_col + cyl - 1
                date_txt = current[ci] if ci < len(current) else ""
                hrs_txt = nxt[ci] if ci < len(nxt) else ""
                date_val = parse_date_value(date_txt)
                hrs_val = parse_number(hrs_txt)

                local_issues = list(pair_issue_list)
                if date_txt and not date_val:
                    local_issues.append(make_issue("warning", "ME_BAD_DATE", f"Suspicious ME date '{date_txt}' for {name} Cyl {cyl}.", table_idx, r))
                if hrs_txt and hrs_val <= 0:
                    local_issues.append(make_issue("warning", "ME_BAD_HOURS", f"Suspicious ME hours '{hrs_txt}' for {name} Cyl {cyl}.", table_idx, r + 1))
                if period_raw and period <= 0:
                    local_issues.append(make_issue("info", "ME_NON_NUMERIC_PERIOD", f"Non-numeric or observational periodicity '{period_raw}' for {name}.", table_idx, r))

                if date_val or hrs_val > 0:
                    rec = normalize_row_record({
                        "category": "MAIN_ENGINE",
                        "engine_label": "ME",
                        "unit": f"Cyl {cyl}",
                        "description": name,
                        "periodicity_raw": first_line(period_raw),
                        "periodicity_hours": period,
                        "last_oh_date": date_val,
                        "hrs_since": hrs_val,
                        "pct_used": compute_pct(hrs_val, period),
                        "status": compute_status(hrs_val, period),
                        "source_table_index": table_idx,
                        "source_row_start": r,
                        "source_row_end": r + 1,
                        "source_col_date": ci,
                        "source_col_hours": ci,
                        "raw_date_text": first_line(date_txt),
                        "raw_hours_text": first_line(hrs_txt),
                        "confidence": confidence_score(date_txt, hrs_txt, period_raw, local_issues),
                        "issue_count": len(local_issues),
                    })
                    records.append(rec)
                    rk = row_key(rec)
                    for it in local_issues:
                        it["row_key"] = rk
                        issues.append(it)
            r += 2
        else:
            r += 1

    return records, issues

def parse_aux_table(grid: List[List[str]], table_idx: int) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    records = []
    issues = []

    joined = " ".join(" ".join(r) for r in grid[:10]).upper()
    if "AUX. ENGINE" not in joined and "AUX ENGINE" not in joined:
        return records, issues

    desc_row = None
    for ri, row in enumerate(grid):
        if any("DESCRIPTION" == normalize_whitespace(c).upper() for c in row):
            desc_row = ri
            break
    if desc_row is None:
        issues.append(make_issue("warning", "AUX_DESC_ROW_NOT_FOUND", "AUX description row not found.", table_idx))
        return records, issues

    header_band = [" ".join(grid[i]).upper() for i in range(max(0, desc_row - 6), desc_row + 1)]
    header_join = " ".join(header_band)
    present_engines = []
    for n in [1, 2, 3]:
        if f"AUX. ENGINE NO.{n}" in header_join or f"AUX ENGINE NO.{n}" in header_join or f"NO.{n}" in header_join:
            present_engines.append(f"AUX-{n}")
    if not present_engines:
        present_engines = ["AUX-1", "AUX-2", "AUX-3"]

    col_nums = []
    for ci in range(2, len(grid[desc_row])):
        token = normalize_whitespace(grid[desc_row][ci])
        if re.fullmatch(r'\d+', token):
            col_nums.append((ci, int(token)))

    if not col_nums:
        issues.append(make_issue("warning", "AUX_CYL_HEADERS_MISSING", "No AUX cylinder headers found; using fallback grouping.", table_idx, desc_row))
        cyl_count = 6
        start_col = 3
    else:
        max_seen = max(n for _, n in col_nums)
        cyl_count = max(1, min(7, max_seen))
        start_col = min(ci for ci, _ in col_nums)

    group_count = len(present_engines)
    if group_count < 1:
        group_count = 1

    r = desc_row + 1
    while r < len(grid) - 1:
        current = grid[r]
        nxt = grid[r + 1] if r + 1 < len(grid) else []

        name = clean_name(current[0] if len(current) > 0 else "")
        period_raw = current[1] if len(current) > 1 else ""
        period = periodicity_value(period_raw)
        marker = normalize_whitespace(current[2] if len(current) > 2 else "")

        if is_component_name(name) and marker == "1":
            pair_marker = normalize_whitespace(nxt[2] if len(nxt) > 2 else "")
            pair_issue_list = []
            if pair_marker != "2":
                pair_issue_list.append(make_issue(
                    "warning", "AUX_PAIR_MARKER_MISSING",
                    f"Expected paired marker '2' below AUX component '{name}', found '{pair_marker or '(blank)'}'.",
                    table_idx, r
                ))

            for engine_idx in range(group_count):
                eng_label = present_engines[engine_idx] if engine_idx < len(present_engines) else f"AUX-{engine_idx+1}"
                group_start = start_col + (engine_idx * cyl_count)
                for cyl in range(1, cyl_count + 1):
                    ci = group_start + cyl - 1
                    date_txt = current[ci] if ci < len(current) else ""
                    hrs_txt = nxt[ci] if ci < len(nxt) else ""
                    date_val = parse_date_value(date_txt)
                    hrs_val = parse_number(hrs_txt)

                    local_issues = list(pair_issue_list)
                    if date_txt and not date_val:
                        local_issues.append(make_issue("warning", "AUX_BAD_DATE", f"Suspicious AUX date '{date_txt}' for {name} {eng_label} Cyl {cyl}.", table_idx, r))
                    if hrs_txt and hrs_val <= 0:
                        local_issues.append(make_issue("warning", "AUX_BAD_HOURS", f"Suspicious AUX hours '{hrs_txt}' for {name} {eng_label} Cyl {cyl}.", table_idx, r + 1))
                    if period_raw and period <= 0:
                        local_issues.append(make_issue("info", "AUX_NON_NUMERIC_PERIOD", f"Non-numeric or observational periodicity '{period_raw}' for {name}.", table_idx, r))

                    if date_val or hrs_val > 0:
                        rec = normalize_row_record({
                            "category": "AUX_ENGINE",
                            "engine_label": eng_label,
                            "unit": f"Cyl {cyl}",
                            "description": name,
                            "periodicity_raw": first_line(period_raw),
                            "periodicity_hours": period,
                            "last_oh_date": date_val,
                            "hrs_since": hrs_val,
                            "pct_used": compute_pct(hrs_val, period),
                            "status": compute_status(hrs_val, period),
                            "source_table_index": table_idx,
                            "source_row_start": r,
                            "source_row_end": r + 1,
                            "source_col_date": ci,
                            "source_col_hours": ci,
                            "raw_date_text": first_line(date_txt),
                            "raw_hours_text": first_line(hrs_txt),
                            "confidence": confidence_score(date_txt, hrs_txt, period_raw, local_issues),
                            "issue_count": len(local_issues),
                        })
                        records.append(rec)
                        rk = row_key(rec)
                        for it in local_issues:
                            it["row_key"] = rk
                            issues.append(it)
            r += 2
        else:
            r += 1

    return records, issues

def parse_other_equipment(grid: List[List[str]], table_idx: int) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    records = []
    issues = []

    joined = " ".join(" ".join(r) for r in grid[:8]).upper()
    if not any(x in joined for x in ["TURBOCHARGER", "COOLERS", "A/C", "AUXILIARY BOILER", "D/G NO1", "D/G NO.1", "DG NO1"]):
        return records, issues

    skip = {
        "", "TURBOCHARGER", "COOLERS", "A/C & REFR. COMPRESSORS",
        "AUXILIARY BOILER", "EXH GAS BOILER", "MAIN AIR COMPRESSORS",
        "DESCRIPTION", "PERIODICTLY", "DATE OF LAST O/H", "RUN HRS"
    }

    for r, row in enumerate(grid):
        for sec, dc, dtc, hrc in [
            ("Turbocharger / Aux Boiler", 0, 2, 3),
            ("Coolers / Exh Gas Boiler", 5, 6, 7),
            ("A/C & Compressors", 10, 11, 12),
        ]:
            desc = clean_name(row[dc] if dc < len(row) else "")
            if desc.upper() in skip or not is_component_name(desc):
                continue
            dt = first_line(row[dtc] if dtc < len(row) else "")
            hr = first_line(row[hrc] if hrc < len(row) else "")
            if dt or hr:
                records.append({
                    "section": sec,
                    "description": desc,
                    "last_date": dt,
                    "run_hrs": hr,
                    "source_table_index": table_idx,
                    "source_row": r,
                })

    return records, issues

def dedupe_rows(records: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    out = []
    issues = []
    seen = set()

    for rec in normalize_rows_payload(records):
        k = (
            rec.get("category"),
            rec.get("engine_label"),
            rec.get("unit"),
            rec.get("description"),
            rec.get("last_oh_date"),
            float(rec.get("hrs_since") or 0),
            rec.get("source_table_index"),
            rec.get("source_row_start"),
            rec.get("source_col_date"),
        )
        if k in seen:
            issues.append(make_issue(
                "warning", "DUPLICATE_ROW",
                f"Duplicate parsed row dropped: {rec.get('description')} / {rec.get('engine_label')} / {rec.get('unit')}.",
                rec.get("source_table_index"),
                rec.get("source_row_start"),
                row_key(rec)
            ))
            continue
        seen.add(k)
        out.append(rec)

    return out, issues

def parse_docx_bytes(docx_bytes: bytes) -> Dict[str, Any]:
    from docx import Document

    issues = []
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
        tf.write(docx_bytes)
        temp_path = tf.name
    try:
        doc = Document(temp_path)
    except Exception as e:
        raise ValueError(f"Cannot open DOCX: {e}")
    finally:
        try:
            os.unlink(temp_path)
        except Exception:
            pass

    if not doc.tables:
        raise ValueError("No tables found. This does not look like a valid TEC-004 report.")

    vessel_name, report_date, meta_issues = detect_vessel_and_date(doc)
    issues.extend(meta_issues)

    me_total_hrs = 0
    me_this_month = 0
    all_me = []
    all_aux = []
    all_oe = []

    for idx, table in enumerate(doc.tables):
        grid = rect_grid(table)
        if not grid:
            continue

        if idx == 0:
            me_total_hrs, me_this_month = detect_me_totals(grid)

        me_rows, me_issues = parse_me_table(grid, idx)
        aux_rows, aux_issues = parse_aux_table(grid, idx)
        oe_rows, oe_issues = parse_other_equipment(grid, idx)

        all_me.extend(me_rows)
        all_aux.extend(aux_rows)
        all_oe.extend(oe_rows)
        issues.extend(me_issues)
        issues.extend(aux_issues)
        issues.extend(oe_issues)

    combined = all_me + all_aux
    combined, dedupe_issues = dedupe_rows(combined)
    issues.extend(dedupe_issues)

    if not combined:
        issues.append(make_issue("error", "NO_COMPONENTS", "No Main Engine or Auxiliary Engine component rows were extracted."))

    return normalize_parsed_payload({
        "vessel_name": vessel_name,
        "report_date": report_date,
        "me_total_hrs": me_total_hrs,
        "me_this_month": me_this_month,
        "me_comps": [x for x in combined if x.get("category") == "MAIN_ENGINE"],
        "aux_comps": [x for x in combined if x.get("category") == "AUX_ENGINE"],
        "components": combined,
        "other_equipment": all_oe,
        "issues": issues,
        "uploaded_at": now_iso(),
    })

# ============================================================
# DB PERSISTENCE
# ============================================================
def ensure_vessel(conn, vessel_name: str) -> int:
    cur = conn.cursor()
    cur.execute("SELECT id FROM vessels WHERE name = ?", (vessel_name,))
    row = cur.fetchone()
    if row:
        return row["id"]
    cur.execute(
        "INSERT INTO vessels (name, created_at) VALUES (?, ?)",
        (vessel_name, now_iso())
    )
    conn.commit()
    return cur.lastrowid

def report_exists(conn, vessel_name: str, report_date: str, file_hash: str) -> bool:
    cur = conn.cursor()
    cur.execute(
        "SELECT id FROM reports WHERE vessel_name = ? AND report_date = ? AND file_hash = ?",
        (vessel_name, report_date, file_hash)
    )
    return cur.fetchone() is not None

def save_parsed_report(parsed: Dict[str, Any]) -> Tuple[bool, str]:
    parsed = normalize_parsed_payload(parsed)
    conn = get_conn()
    try:
        vessel_id = ensure_vessel(conn, parsed["vessel_name"])
        if report_exists(conn, parsed["vessel_name"], parsed["report_date"], parsed.get("file_hash", "")):
            return False, "This exact report already exists in the database."

        cur = conn.cursor()
        cur.execute("""
            INSERT INTO reports (
                vessel_id, vessel_name, report_date, filename, file_hash,
                parser_version, me_total_hrs, me_this_month, raw_json, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            vessel_id,
            parsed.get("vessel_name", "UNKNOWN"),
            parsed.get("report_date", ""),
            parsed.get("filename", ""),
            parsed.get("file_hash", ""),
            PARSER_VERSION,
            parsed.get("me_total_hrs", 0),
            parsed.get("me_this_month", 0),
            json.dumps(parsed, ensure_ascii=False, default=str),
            now_iso()
        ))
        report_id = cur.lastrowid

        for rec in normalize_rows_payload(parsed.get("components", [])):
            cur.execute("""
                INSERT INTO parsed_rows (
                    report_id, category, engine_label, unit, description,
                    periodicity_raw, periodicity_hours, last_oh_date, hrs_since,
                    pct_used, status, source_table_index, source_row_start, source_row_end,
                    source_col_date, source_col_hours, raw_date_text, raw_hours_text,
                    confidence, issue_count, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                report_id,
                rec.get("category"),
                rec.get("engine_label"),
                rec.get("unit"),
                rec.get("description"),
                rec.get("periodicity_raw"),
                rec.get("periodicity_hours"),
                rec.get("last_oh_date"),
                rec.get("hrs_since"),
                rec.get("pct_used"),
                rec.get("status"),
                rec.get("source_table_index"),
                rec.get("source_row_start"),
                rec.get("source_row_end"),
                rec.get("source_col_date"),
                rec.get("source_col_hours"),
                rec.get("raw_date_text"),
                rec.get("raw_hours_text"),
                rec.get("confidence"),
                rec.get("issue_count", 0),
                now_iso()
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
                report_id,
                issue.get("row_key", ""),
                issue.get("severity"),
                issue.get("issue_code"),
                issue.get("message"),
                issue.get("table_index"),
                issue.get("row_index"),
                now_iso()
            ))

        conn.commit()
        return True, f"Saved successfully. Report ID: {report_id}"
    except Exception as e:
        conn.rollback()
        return False, f"Database save failed: {e}"
    finally:
        conn.close()

def load_recent_reports(limit: int = 20) -> pd.DataFrame:
    conn = get_conn()
    try:
        df = pd.read_sql_query(f"""
            SELECT
                r.id,
                r.vessel_name,
                r.report_date,
                r.filename,
                r.parser_version,
                r.me_total_hrs,
                r.me_this_month,
                r.created_at,
                COUNT(pr.id) AS parsed_rows,
                COALESCE(SUM(CASE WHEN pi.severity = 'error' THEN 1 ELSE 0 END), 0) AS errors,
                COALESCE(SUM(CASE WHEN pi.severity = 'warning' THEN 1 ELSE 0 END), 0) AS warnings
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
def safe_float(x):
    try:
        v = float(x)
        return 0.0 if pd.isna(v) else v
    except Exception:
        return 0.0

def cyl_num(unit: str) -> int:
    m = re.search(r'(\d+)', str(unit))
    return int(m.group(1)) if m else 999

def build_preview_df(records: List[Dict[str, Any]], include_trace: bool = False, mode: str = "matrix") -> pd.DataFrame:
    base_cols = [
        "Status", "Component", "Engine", "Unit",
        "Periodicity", "Last O/H", "Hrs Since", "Used %",
        "Confidence", "Issues"
    ]
    if include_trace:
        base_cols += ["Table", "Rows", "Raw Date", "Raw Hrs"]

    if not records:
        return pd.DataFrame(columns=base_cols)

    if isinstance(records, pd.DataFrame):
        df = records.copy()
    else:
        safe_records = []
        for r in records:
            if isinstance(r, dict):
                safe_records.append(normalize_row_record(r))
            else:
                safe_records.append(normalize_row_record({}))
        df = pd.DataFrame(safe_records)

    required_defaults = {
        "status": "NO DATA",
        "description": "",
        "engine_label": "",
        "unit": "",
        "periodicity_hours": 0.0,
        "periodicity_raw": "",
        "last_oh_date": "",
        "hrs_since": 0.0,
        "pct_used": 0.0,
        "confidence": 0.0,
        "issue_count": 0,
        "source_table_index": "",
        "source_row_start": "",
        "source_row_end": "",
        "raw_date_text": "",
        "raw_hours_text": "",
    }

    for col, default in required_defaults.items():
        if col not in df.columns:
            df[col] = default

    df["_status_order"] = df["status"].map(lambda x: STATUS_ORDER.get(str(x), 9))
    df["_pct"] = df["pct_used"].map(safe_float)
    df["_desc"] = df["description"].astype(str).str.upper()
    df["_cyl"] = df["unit"].astype(str).map(cyl_num)

    if mode == "priority":
        df = df.sort_values(["_status_order", "_pct"], ascending=[True, False])
    else:
        df = df.sort_values(["engine_label", "_desc", "_cyl"])

    out = pd.DataFrame()
    out["Status"] = df["status"].map(lambda s: {
        "OVERDUE": "🔴 OVERDUE",
        "HIGH PRIORITY": "🟠 HIGH PRIORITY",
        "OK": "🟢 OK",
        "NO DATA": "🔵 NO DATA",
    }.get(str(s), "🔵 NO DATA"))

    out["Component"] = df["description"].astype(str)
    out["Engine"] = df["engine_label"].astype(str)
    out["Unit"] = df["unit"].astype(str)

    out["Periodicity"] = [
        f"{int(float(x))}" if safe_float(x) > 0 else (str(raw).strip() if str(raw).strip() else "—")
        for x, raw in zip(df["periodicity_hours"], df["periodicity_raw"])
    ]

    out["Last O/H"] = [
        str(x).strip() if str(x).strip() and str(x).strip().lower() not in {"nan", "none"} else "—"
        for x in df["last_oh_date"]
    ]

    out["Hrs Since"] = [
        f"{int(float(x))}" if safe_float(x) > 0 else "—"
        for x in df["hrs_since"]
    ]

    out["Used %"] = [
        round(safe_float(x) * 100, 1) if safe_float(x) > 0 else 0.0
        for x in df["pct_used"]
    ]

    out["Confidence"] = [
        round(safe_float(x) * 100, 0)
        for x in df["confidence"]
    ]

    out["Issues"] = [
        int(safe_float(x))
        for x in df["issue_count"]
    ]

    if include_trace:
        out["Table"] = df["source_table_index"].astype(str)
        out["Rows"] = [
            f"{a}-{b}" if str(a).strip() or str(b).strip() else "—"
            for a, b in zip(df["source_row_start"], df["source_row_end"])
        ]
        out["Raw Date"] = df["raw_date_text"].astype(str)
        out["Raw Hrs"] = df["raw_hours_text"].astype(str)

    return out

PREVIEW_CONFIG = {
    "Status": st.column_config.TextColumn("Status", width=150),
    "Component": st.column_config.TextColumn("Component", width=250),
    "Engine": st.column_config.TextColumn("Engine", width=85),
    "Unit": st.column_config.TextColumn("Unit", width=75),
    "Periodicity": st.column_config.TextColumn("Periodicity", width=120),
    "Last O/H": st.column_config.TextColumn("Last O/H", width=115),
    "Hrs Since": st.column_config.TextColumn("Hrs Since", width=100),
    "Used %": st.column_config.ProgressColumn("Used %", min_value=0, max_value=150, format="%.1f%%", width=130),
    "Confidence": st.column_config.ProgressColumn("Confidence", min_value=0, max_value=100, format="%d%%", width=110),
    "Issues": st.column_config.NumberColumn("Issues", width=75),
    "Table": st.column_config.TextColumn("Table", width=70),
    "Rows": st.column_config.TextColumn("Rows", width=80),
    "Raw Date": st.column_config.TextColumn("Raw Date", width=120),
    "Raw Hrs": st.column_config.TextColumn("Raw Hrs", width=100),
}

def show_df(df: pd.DataFrame, height: int = None):
    height = height or min(860, 38 * len(df) + 44)
    cfg = {k: v for k, v in PREVIEW_CONFIG.items() if k in df.columns}
    st.dataframe(df, use_container_width=True, hide_index=True, height=height, column_config=cfg)

def issues_df(issues: List[Dict[str, Any]]) -> pd.DataFrame:
    if not issues:
        return pd.DataFrame(columns=["Severity", "Code", "Message", "Table", "Row", "Row Key"])
    safe_issues = [x for x in issues if isinstance(x, dict)]
    if not safe_issues:
        return pd.DataFrame(columns=["Severity", "Code", "Message", "Table", "Row", "Row Key"])
    df = pd.DataFrame(safe_issues)
    for col in ["severity", "issue_code", "message", "table_index", "row_index", "row_key"]:
        if col not in df.columns:
            df[col] = ""
    out = pd.DataFrame()
    out["Severity"] = df["severity"]
    out["Code"] = df["issue_code"]
    out["Message"] = df["message"]
    out["Table"] = df["table_index"]
    out["Row"] = df["row_index"]
    out["Row Key"] = df["row_key"]
    return out

# ============================================================
# SESSION
# ============================================================
if "parsed" not in st.session_state:
    st.session_state.parsed = None

if "save_result" not in st.session_state:
    st.session_state.save_result = None

# ============================================================
# UI
# ============================================================
st.markdown('<div class="hr-title">TEC-004 Running Hours</div>', unsafe_allow_html=True)
st.title("Upload → Inspect → Save")
st.markdown('<div class="section-rule"></div>', unsafe_allow_html=True)

st.markdown(
    '<div class="small-note">This app is built to make parser mistakes visible before persistence. '
    'Main Engine and Auxiliary Engine results are shown as preview matrices with row-level provenance, '
    'confidence, and warnings so misaligned cylinder parsing is easier to catch.</div>',
    unsafe_allow_html=True
)

with st.expander("Upload TEC-004 .doc report", expanded=(st.session_state.parsed is None)):
    c1, c2 = st.columns([2.2, 1.3], gap="large")
    with c1:
        uploaded = st.file_uploader("Upload TEC-004 .doc", type=["doc"], label_visibility="collapsed")
        if uploaded:
            st.caption(f"{uploaded.name} · {uploaded.size/1024:.1f} kB")
    with c2:
        st.markdown(
            '<div class="small-note">'
            '<b>Accepted:</b> legacy .doc TEC-004 monthly report<br>'
            '<b>Extracted:</b> vessel, report date, M/E totals, ME rows, AUX rows<br>'
            '<b>Protection:</b> preview before save, duplicate guard, row-level issue log'
            '</div>',
            unsafe_allow_html=True
        )

    if uploaded:
        raw = uploaded.read()
        fhash = md5_bytes(raw)

        current_parsed = normalize_parsed_payload(st.session_state.parsed or {})
        current_hash = current_parsed.get("file_hash")

        if current_hash != fhash:
            with st.spinner("Converting .doc → .docx..."):
                try:
                    docx_bytes = convert_doc_to_docx(raw)
                except Exception as e:
                    st.error(f"Conversion failed: {e}")
                    st.stop()

            with st.spinner("Parsing TEC-004 tables..."):
                try:
                    parsed = parse_docx_bytes(docx_bytes)
                    parsed["filename"] = uploaded.name
                    parsed["file_hash"] = fhash
                    parsed = normalize_parsed_payload(parsed)
                except Exception as e:
                    st.error(f"Parse failed: {e}")
                    st.stop()

            st.session_state.parsed = parsed
            st.session_state.save_result = None

if st.session_state.parsed is None:
    st.info("Upload a TEC-004 .doc report to begin.")
    st.stop()

p = normalize_parsed_payload(st.session_state.parsed or {})

all_rows = normalize_rows_payload(p.get("components") or [])
me_rows = normalize_rows_payload(p.get("me_comps") or [])
aux_rows = normalize_rows_payload(p.get("aux_comps") or [])
oe_rows = p.get("other_equipment") or []
all_issues = p.get("issues") or []

err_count = sum(1 for i in all_issues if isinstance(i, dict) and i.get("severity") == "error")
warn_count = sum(1 for i in all_issues if isinstance(i, dict) and i.get("severity") == "warning")
od_count = sum(1 for r in all_rows if r.get("status") == "OVERDUE")
hp_count = sum(1 for r in all_rows if r.get("status") == "HIGH PRIORITY")

me_total_display = int(p.get("me_total_hrs", 0) or 0)
me_month_display = int(p.get("me_this_month", 0) or 0)

st.markdown(f"""
<div class="kpi-wrap">
  <div class="kpi"><div class="kpi-v">{p.get('vessel_name', 'UNKNOWN')}</div><div class="kpi-l">Vessel</div></div>
  <div class="kpi"><div class="kpi-v">{p.get('report_date', '') or '—'}</div><div class="kpi-l">Report Date</div></div>
  <div class="kpi"><div class="kpi-v">{me_total_display:,}</div><div class="kpi-l">ME Total Hrs</div></div>
  <div class="kpi"><div class="kpi-v">{me_month_display:,}</div><div class="kpi-l">ME This Month</div></div>
  <div class="kpi"><div class="kpi-v">{len(all_rows)}</div><div class="kpi-l">Parsed Rows</div></div>
  <div class="kpi"><div class="kpi-v">{warn_count} / {err_count}</div><div class="kpi-l">Warnings / Errors</div></div>
</div>
""", unsafe_allow_html=True)

status_col1, status_col2, status_col3, status_col4 = st.columns(4)
with status_col1:
    st.markdown(f'<span class="tag tag-red">Overdue: {od_count}</span>', unsafe_allow_html=True)
with status_col2:
    st.markdown(f'<span class="tag tag-amber">High Priority: {hp_count}</span>', unsafe_allow_html=True)
with status_col3:
    st.markdown(f'<span class="tag tag-blue">ME Rows: {len(me_rows)}</span>', unsafe_allow_html=True)
with status_col4:
    st.markdown(f'<span class="tag tag-green">AUX Rows: {len(aux_rows)}</span>', unsafe_allow_html=True)

if err_count:
    st.markdown(f'<div class="err-box"><b>{err_count}</b> error-level parse issues detected. Review preview before save.</div>', unsafe_allow_html=True)
elif warn_count:
    st.markdown(f'<div class="warn-box"><b>{warn_count}</b> warning-level parse issues detected. Review trace columns before save.</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="ok-box">No parse issues detected for this file.</div>', unsafe_allow_html=True)

# ============================================================
# SAVE PANEL
# ============================================================
save_c1, save_c2, save_c3 = st.columns([1.2, 1.2, 4])

with save_c1:
    allow_save = st.checkbox("I reviewed the preview", value=False)

with save_c2:
    if st.button("Save to SQLite", disabled=not allow_save):
        ok, msg = save_parsed_report(p)
        st.session_state.save_result = (ok, msg)

with save_c3:
    export_df = build_preview_df(all_rows, include_trace=True, mode="matrix")
    csv_bytes = export_df.to_csv(index=False).encode("utf-8")
    st.download_button("Download preview CSV", data=csv_bytes, file_name="tec004_preview.csv", mime="text/csv")

if st.session_state.save_result:
    ok, msg = st.session_state.save_result
    if ok:
        st.success(msg)
    else:
        st.warning(msg)

# ============================================================
# MAIN ENGINE
# ============================================================
st.markdown("## Main Engine")

mef1, mef2, mef3, mef4 = st.columns([2, 2, 2.2, 2.4])
with mef1:
    me_comp_opt = ["All"] + sorted({r.get("description", "") for r in me_rows if r.get("description", "")})
    me_comp = st.selectbox("Component", me_comp_opt, key="me_comp")
with mef2:
    me_status = st.selectbox("Status", ["All", "Overdue only", "High Priority +", "Issue rows only"], key="me_status")
with mef3:
    me_trace = st.checkbox("Show trace columns", value=True, key="me_trace")
with mef4:
    me_sort = st.radio("Sort", ["Component → Cylinder", "Priority → % Used"], horizontal=True, key="me_sort")

me_view = me_rows[:]
if me_comp != "All":
    me_view = [r for r in me_view if r.get("description") == me_comp]
if me_status == "Overdue only":
    me_view = [r for r in me_view if r.get("status") == "OVERDUE"]
elif me_status == "High Priority +":
    me_view = [r for r in me_view if r.get("status") in ("OVERDUE", "HIGH PRIORITY")]
elif me_status == "Issue rows only":
    me_view = [r for r in me_view if int(r.get("issue_count", 0)) > 0]

me_df = build_preview_df(
    me_view,
    include_trace=me_trace,
    mode="priority" if "Priority" in me_sort else "matrix"
)
show_df(me_df)

# ============================================================
# AUX ENGINE
# ============================================================
st.markdown("## Auxiliary Engines")

axf1, axf2, axf3, axf4, axf5 = st.columns([1.4, 2, 2, 2.2, 2.4])
with axf1:
    eng_opt = ["All"] + sorted({r.get("engine_label", "") for r in aux_rows if r.get("engine_label", "")})
    ax_eng = st.selectbox("Engine", eng_opt, key="ax_eng")
with axf2:
    ax_comp_opt = ["All"] + sorted({r.get("description", "") for r in aux_rows if r.get("description", "")})
    ax_comp = st.selectbox("Component", ax_comp_opt, key="ax_comp")
with axf3:
    ax_status = st.selectbox("Status", ["All", "Overdue only", "High Priority +", "Issue rows only"], key="ax_status")
with axf4:
    ax_trace = st.checkbox("Show trace columns", value=True, key="ax_trace")
with axf5:
    ax_sort = st.radio("Sort", ["Component → Cylinder", "Priority → % Used"], horizontal=True, key="ax_sort")

ax_view = aux_rows[:]
if ax_eng != "All":
    ax_view = [r for r in ax_view if r.get("engine_label") == ax_eng]
if ax_comp != "All":
    ax_view = [r for r in ax_view if r.get("description") == ax_comp]
if ax_status == "Overdue only":
    ax_view = [r for r in ax_view if r.get("status") == "OVERDUE"]
elif ax_status == "High Priority +":
    ax_view = [r for r in ax_view if r.get("status") in ("OVERDUE", "HIGH PRIORITY")]
elif ax_status == "Issue rows only":
    ax_view = [r for r in ax_view if int(r.get("issue_count", 0)) > 0]

ax_df = build_preview_df(
    ax_view,
    include_trace=ax_trace,
    mode="priority" if "Priority" in ax_sort else "matrix"
)
show_df(ax_df)

# ============================================================
# OTHER EQUIPMENT
# ============================================================
st.markdown("## Other Equipment")

if not oe_rows:
    st.info("No other equipment preview rows extracted.")
else:
    oe_df = pd.DataFrame(oe_rows)
    for col in ["section", "description", "last_date", "run_hrs", "source_table_index", "source_row"]:
        if col not in oe_df.columns:
            oe_df[col] = ""
    st.dataframe(
        oe_df[["section", "description", "last_date", "run_hrs", "source_table_index", "source_row"]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "section": st.column_config.TextColumn("Section", width=220),
            "description": st.column_config.TextColumn("Description", width=300),
            "last_date": st.column_config.TextColumn("Last Date", width=140),
            "run_hrs": st.column_config.TextColumn("Run Hrs", width=120),
            "source_table_index": st.column_config.TextColumn("Table", width=80),
            "source_row": st.column_config.TextColumn("Row", width=80),
        }
    )

# ============================================================
# ISSUES
# ============================================================
st.markdown("## Parse Issues")

idf = issues_df(all_issues)
if idf.empty:
    st.success("No parse issues logged for this upload.")
else:
    sev_filter = st.multiselect("Severity filter", ["error", "warning", "info"], default=["error", "warning", "info"])
    idf2 = idf[idf["Severity"].isin(sev_filter)] if sev_filter else idf.copy()
    st.dataframe(
        idf2,
        use_container_width=True,
        hide_index=True,
        height=min(500, 38 * len(idf2) + 44),
        column_config={
            "Severity": st.column_config.TextColumn("Severity", width=90),
            "Code": st.column_config.TextColumn("Code", width=160),
            "Message": st.column_config.TextColumn("Message", width=500),
            "Table": st.column_config.TextColumn("Table", width=70),
            "Row": st.column_config.TextColumn("Row", width=70),
            "Row Key": st.column_config.TextColumn("Row Key", width=320),
        }
    )

# ============================================================
# DATABASE HISTORY
# ============================================================
st.markdown("## Recent Saved Reports")

hist = load_recent_reports(20)
if hist.empty:
    st.info("No saved reports yet.")
else:
    st.dataframe(
        hist,
        use_container_width=True,
        hide_index=True,
        column_config={
            "id": st.column_config.NumberColumn("ID", width=65),
            "vessel_name": st.column_config.TextColumn("Vessel", width=150),
            "report_date": st.column_config.TextColumn("Report Date", width=120),
            "filename": st.column_config.TextColumn("Filename", width=260),
            "parser_version": st.column_config.TextColumn("Parser", width=180),
            "me_total_hrs": st.column_config.NumberColumn("ME Total", width=100, format="%d"),
            "me_this_month": st.column_config.NumberColumn("ME Month", width=100, format="%d"),
            "created_at": st.column_config.TextColumn("Saved At", width=190),
            "parsed_rows": st.column_config.NumberColumn("Rows", width=70),
            "errors": st.column_config.NumberColumn("Errors", width=70),
            "warnings": st.column_config.NumberColumn("Warnings", width=85),
        }
    )
