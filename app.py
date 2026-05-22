import streamlit as st
st.set_page_config(
    page_title="Fleet Running Hours Monitor 10.10 Debug-Operational",
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
from typing import Any, Dict, List, Tuple

import pandas as pd

# ══════════════════════════════════════════════════════════════════
#  DESIGN
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
  --bg:#060b12; --bg2:#0a111b; --bg3:#111a29; --bg4:#162131;
  --line:#18283d; --line2:#27405e;
  --gold:#b8870c; --gold2:#d5a223; --gold3:#edc55d;
  --red:#eb5b5b; --ora:#f6ad4e; --grn:#60be74; --blu:#67aef5; --cyn:#4ad3d7;
  --t0:#e8f0fb; --t1:#a8bed6; --t2:#6e86a0; --t3:#3e556f;
  --ff:'Space Grotesk',sans-serif; --fi:'Inter',sans-serif; --fm:'JetBrains Mono',monospace;
  --r:14px;
}
html, body, [class*="css"] {
  background:var(--bg)!important; color:var(--t1)!important; font-family:var(--fi)!important;
}
.main, .main > div, .block-container { background:var(--bg)!important; }
.block-container { max-width:100%!important; padding:1.05rem 1.65rem 3rem!important; }
[data-testid="collapsedControl"], [data-testid="stSidebar"] { display:none!important; }

.main::before{
  content:""; position:fixed; inset:0; pointer-events:none; z-index:0;
  background:
    radial-gradient(ellipse 75% 45% at 0% 0%, rgba(184,135,12,.08), transparent 60%),
    radial-gradient(ellipse 60% 40% at 100% 100%, rgba(74,211,215,.05), transparent 55%);
}
.block-container > * { position:relative; z-index:1; }

.hero-k { font-size:.66rem; letter-spacing:.24em; text-transform:uppercase; color:var(--gold3); font-weight:700; }
.hero-h { font-family:var(--ff); font-size:2rem; font-weight:700; color:var(--t0); letter-spacing:-.04em; line-height:1.05; margin-top:.22rem; }
.hero-s { color:var(--t1); font-size:.92rem; line-height:1.65; margin-top:.45rem; max-width:1100px; }
.hero-rule { height:1px; margin:.85rem 0 1.1rem; background:linear-gradient(90deg,var(--gold2),var(--line),transparent); }

.panel{
  background:linear-gradient(180deg, rgba(255,255,255,.02), rgba(255,255,255,.01));
  border:1px solid var(--line);
  border-radius:var(--r);
  padding:1rem;
  box-shadow:0 18px 52px rgba(0,0,0,.32);
}
.upload-panel{
  background:
    linear-gradient(180deg, rgba(184,135,12,.05), rgba(103,174,245,.03)),
    linear-gradient(180deg, rgba(17,26,41,.97), rgba(10,17,27,.97));
  border:1px solid rgba(184,135,12,.22);
}
[data-testid="stFileUploadDropzone"]{
  background:rgba(184,135,12,.04)!important;
  border:1.5px dashed rgba(184,135,12,.55)!important;
  border-radius:16px!important;
  padding:2rem 1.3rem!important;
}
[data-testid="stFileUploadDropzone"]:hover{
  border-color:var(--gold3)!important;
  background:rgba(184,135,12,.07)!important;
}
.metric-grid{
  display:grid;
  grid-template-columns:repeat(8,1fr);
  gap:.75rem;
  margin:1rem 0 1rem;
}
.metric{
  background:linear-gradient(180deg,var(--bg3),var(--bg2));
  border:1px solid var(--line);
  border-radius:var(--r);
  padding:.82rem .95rem .95rem;
  position:relative;
  overflow:hidden;
  box-shadow:0 16px 48px rgba(0,0,0,.25);
}
.metric::before{
  content:""; position:absolute; top:0; left:0; right:0; height:2px;
  background:linear-gradient(90deg,var(--gold2),transparent 75%);
}
.metric.g::before{ background:linear-gradient(90deg,var(--grn),transparent 75%); }
.metric.r::before{ background:linear-gradient(90deg,var(--red),transparent 75%); }
.metric.o::before{ background:linear-gradient(90deg,var(--ora),transparent 75%); }
.metric.b::before{ background:linear-gradient(90deg,var(--blu),transparent 75%); }
.metric.c::before{ background:linear-gradient(90deg,var(--cyn),transparent 75%); }
.metric-v{
  font-family:var(--ff); font-size:1.48rem; font-weight:700; color:var(--t0); line-height:1.05; letter-spacing:-.04em;
}
.metric-l{
  color:var(--t3); font-size:.58rem; text-transform:uppercase; letter-spacing:.17em; margin-top:.34rem;
}
.bar{
  display:flex; align-items:center; gap:.8rem; margin:1.45rem 0 .8rem;
}
.bar-i{
  width:34px; height:34px; border-radius:10px; display:flex; align-items:center; justify-content:center;
  background:rgba(184,135,12,.09); border:1px solid rgba(184,135,12,.2);
}
.bar-t{ font-family:var(--ff); font-size:1rem; font-weight:700; color:var(--t0); }
.bar-s{ font-family:var(--fm); font-size:.6rem; color:var(--t3); margin-top:2px; }
.bar-r{ flex:1; height:1px; background:linear-gradient(90deg,var(--line),transparent); }
.bar-b{
  font-family:var(--fm); font-size:.6rem; color:var(--t2);
  border:1px solid var(--line2); background:var(--bg3); border-radius:999px; padding:4px 9px;
}
.banner{
  border-radius:var(--r); padding:.85rem 1rem; border:1px solid var(--line); margin-bottom:.85rem; line-height:1.55;
  box-shadow:0 18px 50px rgba(0,0,0,.22);
}
.banner.ok{ background:rgba(96,190,116,.10); border-color:rgba(96,190,116,.28); color:#dff7e4; }
.banner.warn{ background:rgba(246,173,78,.10); border-color:rgba(246,173,78,.28); color:#ffe7c8; }
.banner.err{ background:rgba(235,91,91,.10); border-color:rgba(235,91,91,.28); color:#ffdada; }

.muted{ color:var(--t1); font-size:.88rem; line-height:1.65; }
.micro{ font-family:var(--fm); color:var(--t3); font-size:.72rem; }

[data-testid="stDataFrame"]{
  border:1px solid var(--line)!important;
  border-radius:var(--r)!important;
  overflow:hidden!important;
  box-shadow:0 18px 50px rgba(0,0,0,.26)!important;
}
.dvn-scroller{ background:var(--bg2)!important; }

.streamlit-expanderHeader{
  background:var(--bg3)!important; border:1px solid var(--line)!important; border-radius:12px!important; color:var(--t0)!important;
}
.streamlit-expanderContent{
  background:var(--bg2)!important; border:1px solid var(--line)!important; border-top:none!important; border-radius:0 0 12px 12px!important;
}
.stButton>button{
  background:linear-gradient(135deg,var(--gold3),var(--gold2))!important;
  color:#06101a!important; border:none!important; border-radius:11px!important;
  padding:.72rem 1.05rem!important; font-weight:800!important; text-transform:uppercase!important; letter-spacing:.05em!important;
}
.stDownloadButton>button{
  background:var(--bg3)!important; color:var(--t0)!important; border:1px solid var(--line2)!important;
  border-radius:11px!important; padding:.72rem 1rem!important; font-weight:700!important;
}
div[data-baseweb="select"] > div,
.stTextInput > div > div > input {
  background:var(--bg4)!important; color:var(--t0)!important; border:1px solid var(--line2)!important; border-radius:10px!important;
}
.stSelectbox label,.stMultiSelect label,.stCheckbox label,.stRadio label{
  color:var(--t3)!important; font-size:.66rem!important; text-transform:uppercase!important; letter-spacing:.12em!important;
}
@media (max-width:1280px){ .metric-grid{ grid-template-columns:repeat(4,1fr);} }
@media (max-width:860px){ .metric-grid{ grid-template-columns:repeat(2,1fr);} }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
#  CONSTANTS / DB
# ══════════════════════════════════════════════════════════════════
DB_PATH = "tec004_debug_operational_v1010.db"
PARSER_VERSION = "10.10_debug_first_dual_parser"

STATUS_ORDER = {'OVERDUE': 0, 'HIGH PRIORITY': 1, 'OK': 2, 'NO DATA': 3}
STATUS_MAP = {
    'OVERDUE': '🔴 OVERDUE',
    'HIGH PRIORITY': '🟠 HIGH PRIORITY',
    'OK': '🟢 OK',
    'NO DATA': '🔵 NO DATA'
}
ISSUE_ORDER = {'error': 0, 'warning': 1, 'info': 2}

# ══════════════════════════════════════════════════════════════════
#  SQLITE
# ══════════════════════════════════════════════════════════════════
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
        periodicity REAL,
        last_oh_date TEXT,
        hrs_since REAL,
        pct_used REAL,
        status TEXT,
        source_table_index INTEGER,
        source_row_start INTEGER,
        source_row_end INTEGER,
        source_col_date INTEGER,
        source_col_hours INTEGER,
        parser_source TEXT,
        created_at TEXT NOT NULL
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
        created_at TEXT NOT NULL
    )
    """)
    conn.commit()
    conn.close()

init_db()

# ══════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════
def now_iso():
    return datetime.utcnow().isoformat()

def md5_bytes(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()

def fl(txt: Any) -> str:
    raw = str(txt or "").replace('\x07', '').replace('\xa0', ' ').replace('\t', ' ')
    for part in re.split(r'[\r\n\x0b]+', raw):
        s = re.sub(r'\s+', ' ', part).strip()
        if s:
            return s
    return ''

def clean_name(txt: Any) -> str:
    t = fl(txt)
    t = re.sub(r'(?i)^MV\s+', '', t)
    t = re.sub(r'(?i)ALEXIS\s*Date?', '', t)
    t = re.sub(r'(?i)Page\s*\d+\s*of\s*\d+', '', t)
    t = re.sub(r'  +', ' ', t)
    return t.strip(" -:")

def norm_upper(txt: Any) -> str:
    return fl(txt).upper()

def sf(x):
    try:
        v = float(x)
        return 0.0 if pd.isna(v) else v
    except Exception:
        return 0.0

def cyl_n(u: str) -> int:
    m = re.search(r'\d+', str(u))
    return int(m.group()) if m else 999

def status_of(hrs: float, period: float) -> str:
    if hrs <= 0 or period <= 0:
        return 'NO DATA'
    r = hrs / period
    if r >= 1.0:
        return 'OVERDUE'
    if r >= 0.8:
        return 'HIGH PRIORITY'
    return 'OK'

def pct_used(hrs: float, period: float) -> float:
    return round(hrs / period, 4) if hrs and period else 0.0

def make_issue(severity: str, code: str, message: str, table_index=None, row_index=None, row_key='') -> Dict[str, Any]:
    return {
        "severity": severity,
        "issue_code": code,
        "message": message,
        "table_index": table_index,
        "row_index": row_index,
        "row_key": row_key
    }

def is_component_name(name: str) -> bool:
    u = fl(name).upper()
    if not u or len(u) < 2:
        return False
    bad = (
        'DESCRIPTION','REMARKS','COMPONENT','PERIODICITY','PERIODICTLY',
        'DATE OF LAST','RUNNING HOURS','MAIN ENGINE','AUX. ENGINE','TYPE:',
        'TOTAL RUNNING HOURS','THIS MONTH','CYL. NO.','SERIAL NR',
        'HOURS THIS MONTH','TOTAL HOURS','AUX. ENGINE MAKER / TYPE',
        'TURBOCHARGER','COOLERS','A/C & REFR. COMPRESSORS','MAIN AIR COMPRESSORS'
    )
    if any(b in u for b in bad):
        return False
    if re.fullmatch(r'[\d./ ,:\-\[\]()]+', u):
        return False
    return bool(re.search(r'[A-Z]', u))

def parse_number(txt: Any) -> float:
    s = fl(txt).strip().upper()
    if not s or s in ('', '-', 'N/A', 'NA', 'CENTRAL', 'COOLER'):
        return 0.0
    s = s.replace('[', '').replace(']', '')
    if any(w in s for w in ('MONTH','YEAR','WEEK','DAY','OBS')):
        return 0.0
    m = re.search(r'\d[\d,\.]*', s)
    if not m:
        return 0.0
    block = m.group()
    sep = max(block.rfind('.'), block.rfind(','))
    if sep > 0 and len(block) - sep == 4:
        block = re.sub(r'[,\.]', '', block)
    elif sep > 0:
        block = re.sub(r'[,\.]', '', block[:sep])
    else:
        block = re.sub(r'[,\.]', '', block)
    try:
        return float(block)
    except Exception:
        return 0.0

def parse_date(txt: Any) -> str:
    s = fl(txt).strip()
    if not s or s in ('-', '1', '2', 'N/A', 'n/a', 'NA', 'Central', 'cooler', 'CENTRAL', 'COOLER'):
        return ''
    if re.fullmatch(r'^\d+$', s):
        return ''
    if len(s) > 28:
        return ''
    s = s.replace('[', '').replace(']', '').strip()
    if re.match(r'^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$', s):
        dd, mm, yy = re.split(r'[/-]', s)
        try:
            di, mi = int(dd), int(mm)
            if not (1 <= mi <= 12 and 1 <= di <= 31):
                return ''
        except Exception:
            return ''
        return f"{int(dd):02d}/{int(mm):02d}/{str(yy)[-2:]}"
    m = re.match(r'^(\d{1,2})\s+([A-Za-z\.]+)\s+(\d{2,4})$', s)
    if m:
        day = int(m.group(1))
        mon = m.group(2).replace('.', '').upper()
        year = str(m.group(3))[-2:]
        months = {
            'JAN':'JAN','JANUARY':'JAN',
            'FEB':'FEB','FEBRUARY':'FEB',
            'MAR':'MAR','MARCH':'MAR',
            'APR':'APR','APRIL':'APR',
            'MAY':'MAY',
            'JUN':'JUN','JUNE':'JUN',
            'JUL':'JUL','JULY':'JUL',
            'AUG':'AUG','AUGUST':'AUG',
            'SEP':'SEP','SEPT':'SEP','SEPTEMBER':'SEP',
            'OCT':'OCT','OCTOBER':'OCT',
            'NOV':'NOV','NOVEMBER':'NOV',
            'DEC':'DEC','DECEMBER':'DEC',
        }
        if 1 <= day <= 31 and mon in months:
            return f"{day:02d} {months[mon]} {year}"
    return s if re.search(r'[A-Za-z/]', s) else ''

def make_row(cat, eng, unit, name, period, date, hrs, table_idx=None, row_start=None, row_end=None, col_date=None, col_hours=None, parser_source='grid'):
    return {
        'category': cat,
        'engine_label': eng,
        'unit': unit,
        'description': name,
        'periodicity': period,
        'last_oh_date': date,
        'hrs_since': hrs,
        'pct_used': pct_used(hrs, period),
        'status': status_of(hrs, period),
        'source_table_index': table_idx,
        'source_row_start': row_start,
        'source_row_end': row_end,
        'source_col_date': col_date,
        'source_col_hours': col_hours,
        'parser_source': parser_source,
    }

# ══════════════════════════════════════════════════════════════════
#  DOC CONVERSION
# ══════════════════════════════════════════════════════════════════
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found. Install libreoffice in the runtime.")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as t:
        t.write(raw)
        src = t.name
    outdir = tempfile.mkdtemp(prefix="lo_")
    out = os.path.join(outdir, Path(src).stem + ".docx")
    profile = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        r = subprocess.run(
            [soffice, "--headless", "--norestore", "--nofirststartwizard",
             f"-env:UserInstallation={profile}", "--convert-to", "docx", src, "--outdir", outdir],
            capture_output=True, timeout=120
        )
        if not os.path.exists(out):
            raise RuntimeError(r.stderr.decode("utf-8", "ignore")[:500] or r.stdout.decode("utf-8", "ignore")[:500])
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

# ══════════════════════════════════════════════════════════════════
#  TABLE NORMALIZATION
# ══════════════════════════════════════════════════════════════════
def raw_grid(table) -> List[List[str]]:
    grid = []
    max_cols = 0
    for row in table.rows:
        cells = []
        for cell in row.cells:
            raw = re.sub(r'[\x0b\r]', '\n', cell.text).replace('\x07', '')
            lines = [ln.replace('\xa0', ' ').replace('\t', ' ').strip()
                     for ln in raw.split('\n') if ln.strip()]
            cells.append(lines[0] if lines else '')
        max_cols = max(max_cols, len(cells))
        grid.append(cells)
    for row in grid:
        while len(row) < max_cols:
            row.append('')
    return grid

def dedup_grid(table) -> List[List[str]]:
    """
    De-duplicate obvious repeated merged cells inside the same row.
    """
    grid = []
    max_cols = 0
    for row in table.rows:
        cells = []
        prior_tc = None
        for cell in row.cells:
            tc = cell._tc
            if tc is prior_tc:
                continue
            prior_tc = tc
            raw = re.sub(r'[\x0b\r]', '\n', cell.text).replace('\x07', '')
            lines = [ln.replace('\xa0', ' ').replace('\t', ' ').strip()
                     for ln in raw.split('\n') if ln.strip()]
            cells.append(lines[0] if lines else '')
        max_cols = max(max_cols, len(cells))
        grid.append(cells)
    for row in grid:
        while len(row) < max_cols:
            row.append('')
    return grid

def table_text(grid: List[List[str]]) -> str:
    return "\n".join(" | ".join(fl(c) for c in row) for row in grid)

def classify_table(grid: List[List[str]]) -> str:
    txt = table_text(grid).upper()
    if 'MAIN ENGINE' in txt or ('CYL. NO.1' in txt and 'REMARKS' in txt):
        return 'ME'
    if 'AUX. ENGINE MAKER / TYPE' in txt or ('AUX. ENGINE NO.1' in txt and 'DESCRIPTION' in txt and 'PERIODICTLY' in txt):
        return 'AUX'
    if 'D/G NO1' in txt or 'D/G NO.1' in txt or 'D/G NO2' in txt:
        return 'DG'
    if 'TURBOCHARGER' in txt or 'A/C & REFR. COMPRESSORS' in txt or 'MAIN AIR COMPRESSORS' in txt or 'COOLERS' in txt:
        return 'OE'
    return 'UNKNOWN'

# ══════════════════════════════════════════════════════════════════
#  GRID PARSERS
# ══════════════════════════════════════════════════════════════════
def parse_me_from_grid(grid: List[List[str]], table_idx: int, issues: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    rows = []
    if not grid:
        return rows

    rem_col = None
    cyl_map = []
    header_band = grid[:5]

    for c in range(len(grid[0])):
        coltxt = " ".join(norm_upper(header_band[r][c]) if c < len(header_band[r]) else '' for r in range(len(header_band)))
        if 'REMARK' in coltxt and rem_col is None:
            rem_col = c
        m = re.search(r'CYL\.\s*NO\.\s*(\d+)', coltxt)
        if m:
            cyl_map.append((c, int(m.group(1))))

    if rem_col is None:
        rem_col = len(grid[0]) - 1

    if not cyl_map:
        # fallback: detect the first run of columns before remarks after marker column
        for r, row in enumerate(header_band):
            for c, cell in enumerate(row):
                if re.search(r'CYL\.\s*NO\.\s*1', norm_upper(cell)):
                    cyl_map = [(x, i + 1) for i, x in enumerate(range(c, rem_col))]
                    break
            if cyl_map:
                break

    if not cyl_map:
        return rows

    period_col = 1
    marker_col = cyl_map[0][0] - 1 if cyl_map and cyl_map[0][0] >= 2 else 2

    end = len(grid)
    for r, row in enumerate(grid):
        joined = " ".join(norm_upper(x) for x in row)
        if 'NOTE 1' in joined or 'TURBOCHARGER' in joined or 'AUX. ENGINE MAKER / TYPE' in joined:
            end = r
            break

    r = 0
    while r < end - 1:
        name = clean_name(grid[r][0] if len(grid[r]) > 0 else '')
        period = parse_number(grid[r][period_col] if period_col < len(grid[r]) else '')
        marker = fl(grid[r][marker_col] if marker_col < len(grid[r]) else '')
        if is_component_name(name) and marker == '1':
            nxt = grid[r + 1] if r + 1 < len(grid) else []
            for col_idx, cyl_no in cyl_map:
                date = parse_date(grid[r][col_idx] if col_idx < len(grid[r]) else '')
                hrs = parse_number(nxt[col_idx] if col_idx < len(nxt) else '')
                if date or hrs > 0:
                    rows.append(make_row(
                        'MAIN_ENGINE', 'ME', f'Cyl {cyl_no}', name, period, date, hrs,
                        table_idx, r, min(r + 1, len(grid) - 1), col_idx, col_idx, 'grid'
                    ))
            r += 2
        else:
            r += 1
    return rows

def find_aux_groups_from_grid(grid: List[List[str]]) -> Tuple[int, List[Tuple[str, int, int]]]:
    desc_row = None
    for i, row in enumerate(grid):
        rowtxt = " | ".join(norm_upper(c) for c in row)
        if 'DESCRIPTION' in rowtxt and 'PERIODICTLY' in rowtxt:
            desc_row = i
            break
    if desc_row is None:
        return -1, []

    engine_groups = []
    for i, row in enumerate(grid[:desc_row + 1]):
        labels = []
        for c, cell in enumerate(row):
            u = norm_upper(cell)
            if 'AUX. ENGINE NO.1' in u:
                labels.append(('AUX-1', c))
            elif 'AUX. ENGINE NO.2' in u:
                labels.append(('AUX-2', c))
            elif 'AUX. ENGINE NO.3' in u:
                labels.append(('AUX-3', c))
        if labels:
            labels = sorted(labels, key=lambda x: x[1])
            for j, (lbl, start) in enumerate(labels):
                end = labels[j + 1][1] if j + 1 < len(labels) else len(grid[desc_row])
                engine_groups.append((lbl, start, end))
            return desc_row, engine_groups

    # fallback from repeated 1..6 headers
    nums = []
    for c in range(2, len(grid[desc_row])):
        txt = fl(grid[desc_row][c])
        if re.fullmatch(r'\d+', txt):
            nums.append((c, int(txt)))
    if nums:
        starts = [c for c, n in nums if n == 1]
        labels = ['AUX-1', 'AUX-2', 'AUX-3']
        for i, s in enumerate(starts[:3]):
            e = starts[i + 1] if i + 1 < len(starts) else len(grid[desc_row])
            engine_groups.append((labels[i], s, e))
    return desc_row, engine_groups

def parse_aux_from_grid(grid: List[List[str]], table_idx: int, issues: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    rows = []
    if not grid:
        return rows
    desc_row, groups = find_aux_groups_from_grid(grid)
    if desc_row < 0 or not groups:
        return rows

    r = desc_row + 1
    while r < len(grid) - 1:
        name = clean_name(grid[r][0] if len(grid[r]) > 0 else '')
        period = parse_number(grid[r][1] if len(grid[r]) > 1 else '')
        marker = fl(grid[r][2] if len(grid[r]) > 2 else '')
        if is_component_name(name) and marker == '1':
            nxt = grid[r + 1] if r + 1 < len(grid) else []
            for eng_label, start, end in groups:
                for ci in range(start, min(end, len(grid[r]))):
                    hdr = fl(grid[desc_row][ci] if ci < len(grid[desc_row]) else '')
                    if not re.fullmatch(r'\d+', hdr):
                        continue
                    unit_no = int(hdr)
                    date = parse_date(grid[r][ci] if ci < len(grid[r]) else '')
                    hrs = parse_number(nxt[ci] if ci < len(nxt) else '')
                    if date or hrs > 0:
                        rows.append(make_row(
                            'AUX_ENGINE', eng_label, f'Cyl {unit_no}', name, period, date, hrs,
                            table_idx, r, min(r + 1, len(grid) - 1), ci, ci, 'grid'
                        ))
            r += 2
        else:
            r += 1
    return rows

def parse_oe_from_grid(grid: List[List[str]], table_idx: int) -> List[Dict[str, Any]]:
    rows = []
    if not grid:
        return rows

    section_maps = [
        ('Turbocharger / Aux Boiler', 0, 1, 2, 3),
        ('Coolers / Exh Gas Boiler', 5, 6, 7, 8),
        ('A/C & Compressors', 10, 11, 11, 12),
    ]
    skip = {
        'TURBOCHARGER','AUXILIARY BOILER','COOLERS','EXH GAS BOILER','EXH GAS  BOILER',
        'A/C & REFR. COMPRESSORS','MAIN AIR COMPRESSORS','PERIODICTLY','DATE OF LAST INSPECTION',
        'RUN HRS','DATE OF LAST CLEANING','DATE','PERIODICITY',''
    }

    for r, row in enumerate(grid):
        def gc(i): return row[i] if i < len(row) else ''
        for sec, dc, pc, dtc, hrc in section_maps:
            desc = clean_name(gc(dc))
            if not desc or desc.upper() in skip or not is_component_name(desc):
                continue
            periodicity = parse_number(gc(pc))
            dt = parse_date(gc(dtc))
            hrs = parse_number(gc(hrc))
            if dt or hrs > 0 or periodicity > 0:
                rows.append({
                    'section': sec,
                    'description': desc,
                    'periodicity': periodicity,
                    'last_date': dt,
                    'run_hrs': int(hrs) if hrs > 0 else '',
                    'source_table_index': table_idx,
                    'source_row': r
                })
    return rows

def parse_dg_from_grid(grid: List[List[str]], table_idx: int) -> List[Dict[str, Any]]:
    rows = []
    if not grid:
        return rows
    for base in [0, 8]:
        r = 0
        while r < len(grid) - 1:
            desc = clean_name(grid[r][base] if base < len(grid[r]) else '')
            period = parse_number(grid[r][base + 1] if base + 1 < len(grid[r]) else '')
            marker = fl(grid[r][base + 2] if base + 2 < len(grid[r]) else '')
            if is_component_name(desc) and marker == '1':
                nxt = grid[r + 1] if r + 1 < len(grid) else []
                for i, eng_label in enumerate(['D/G 1', 'D/G 2', 'D/G 3']):
                    ci = base + 2 + i
                    date = parse_date(grid[r][ci] if ci < len(grid[r]) else '')
                    hrs = parse_number(nxt[ci] if ci < len(nxt) else '')
                    if date or hrs > 0:
                        rows.append({
                            'section': 'D/G Equipment',
                            'description': desc,
                            'engine_label': eng_label,
                            'periodicity': period,
                            'last_date': date,
                            'run_hrs': int(hrs) if hrs > 0 else '',
                            'source_table_index': table_idx,
                            'source_row': r
                        })
                r += 2
            else:
                r += 1
    return rows

# ══════════════════════════════════════════════════════════════════
#  TEXT FALLBACK PARSERS
# ══════════════════════════════════════════════════════════════════
DATE_RX = r'(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{1,2}\s+[A-Za-z\.]+\s+\d{2,4})'
NUM_RX = r'\[?\d[\d,\.]*\]?'

def lines_from_doc(doc) -> List[str]:
    lines = []
    for p in doc.paragraphs:
        txt = fl(p.text)
        if txt:
            lines.append(txt)
    for table in doc.tables:
        rg = raw_grid(table)
        for row in rg:
            for cell in row:
                txt = fl(cell)
                if txt:
                    lines.append(txt)
    return lines

def extract_stream_between(lines: List[str], start_markers: List[str], end_markers: List[str]) -> List[str]:
    start = None
    end = len(lines)
    for i, line in enumerate(lines):
        up = line.upper()
        if start is None and any(m in up for m in start_markers):
            start = i
            continue
        if start is not None and any(m in up for m in end_markers):
            end = i
            break
    return lines[start:end] if start is not None else []

def parse_me_from_text(lines: List[str], table_idx_hint: int = -1) -> List[Dict[str, Any]]:
    rows = []
    seg = extract_stream_between(
        lines,
        ['MAIN ENGINE', 'CYL. NO.1'],
        ['NOTE 1', 'TURBOCHARGER', 'AUX. ENGINE MAKER / TYPE']
    )
    if not seg:
        return rows

    cyl_headers = [f'CYL. NO.{i}' for i in range(1, 8)]
    data = [x for x in seg if fl(x)]

    i = 0
    while i < len(data):
        name = clean_name(data[i])
        if is_component_name(name):
            period = parse_number(data[i + 1]) if i + 1 < len(data) else 0.0
            marker1 = fl(data[i + 2]) if i + 2 < len(data) else ''
            if marker1 == '1':
                date_vals = []
                j = i + 3
                while j < len(data) and len(date_vals) < 7 and fl(data[j]) != '2':
                    tok = fl(data[j])
                    if tok.upper() not in ('REMARKS',):
                        date_vals.append(tok)
                    j += 1
                if j < len(data) and fl(data[j]) == '2':
                    j += 1
                    hrs_vals = []
                    while j < len(data) and len(hrs_vals) < 7:
                        tok = fl(data[j])
                        if is_component_name(tok):
                            break
                        hrs_vals.append(tok)
                        j += 1

                    for idx in range(7):
                        d = parse_date(date_vals[idx]) if idx < len(date_vals) else ''
                        h = parse_number(hrs_vals[idx]) if idx < len(hrs_vals) else 0.0
                        if d or h > 0:
                            rows.append(make_row(
                                'MAIN_ENGINE', 'ME', f'Cyl {idx+1}', name, period, d, h,
                                table_idx_hint, None, None, None, None, 'text'
                            ))
                    i = j
                    continue
        i += 1
    return rows

def parse_aux_from_text(lines: List[str], table_idx_hint: int = -1) -> List[Dict[str, Any]]:
    rows = []
    seg = extract_stream_between(
        lines,
        ['AUX. ENGINE MAKER / TYPE', 'AUX. ENGINE NO.1'],
        ['DESCRIPTION', 'D/G NO1', 'TURBOCHARGER (2)', '1ST COPY']
    )
    if not seg:
        # broader fallback
        seg = extract_stream_between(
            lines,
            ['AUX. ENGINE NO.1', 'SERIAL NR:'],
            ['D/G NO1', 'TURBOCHARGER (2)', '1ST COPY']
        )
    all_lines = seg if seg else lines

    # look for repeated 1..6 x 3 header then parse component blocks
    start_idx = None
    for i, line in enumerate(all_lines):
        if 'DESCRIPTION' in line.upper() and 'PERIODICTLY' in line.upper():
            start_idx = i
            break
    if start_idx is None:
        return rows

    data = [fl(x) for x in all_lines[start_idx + 1:] if fl(x)]
    i = 0
    while i < len(data):
        name = clean_name(data[i])
        if is_component_name(name):
            period = parse_number(data[i + 1]) if i + 1 < len(data) else 0.0
            marker1 = fl(data[i + 2]) if i + 2 < len(data) else ''
            if marker1 == '1':
                date_vals = []
                j = i + 3
                while j < len(data) and len(date_vals) < 18 and fl(data[j]) != '2':
                    tok = fl(data[j])
                    if is_component_name(tok) and len(date_vals) < 18:
                        break
                    date_vals.append(tok)
                    j += 1
                if j < len(data) and fl(data[j]) == '2':
                    j += 1
                    hrs_vals = []
                    while j < len(data) and len(hrs_vals) < 18:
                        tok = fl(data[j])
                        if is_component_name(tok):
                            break
                        hrs_vals.append(tok)
                        j += 1

                    for eng_i, eng_label in enumerate(['AUX-1', 'AUX-2', 'AUX-3']):
                        base = eng_i * 6
                        for cyl in range(6):
                            idx = base + cyl
                            d = parse_date(date_vals[idx]) if idx < len(date_vals) else ''
                            h = parse_number(hrs_vals[idx]) if idx < len(hrs_vals) else 0.0
                            if d or h > 0:
                                rows.append(make_row(
                                    'AUX_ENGINE', eng_label, f'Cyl {cyl+1}', name, period, d, h,
                                    table_idx_hint, None, None, None, None, 'text'
                                ))
                    i = j
                    continue
        i += 1
    return rows

# ══════════════════════════════════════════════════════════════════
#  MERGE / DEDUPE
# ══════════════════════════════════════════════════════════════════
def dedupe_component_rows(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    best = {}
    for r in records:
        key = (r.get('category'), r.get('description'), r.get('engine_label'), r.get('unit'))
        score = 0
        if r.get('last_oh_date'):
            score += 1
        if sf(r.get('hrs_since')) > 0:
            score += 1
        if r.get('parser_source') == 'grid':
            score += 1
        prev = best.get(key)
        if prev is None:
            best[key] = (score, r)
        else:
            if score > prev[0]:
                best[key] = (score, r)
    return [v[1] for v in best.values()]

# ══════════════════════════════════════════════════════════════════
#  DOCUMENT PARSER
# ══════════════════════════════════════════════════════════════════
def parse_doc_bytes(docx_bytes: bytes, filename: str = '') -> Dict[str, Any]:
    from docx import Document

    issues = []
    warnings = []

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as t:
        t.write(docx_bytes)
        tp = t.name
    try:
        doc = Document(tp)
    except Exception as e:
        raise ValueError(f"Cannot open converted DOCX: {e}")
    finally:
        try:
            os.unlink(tp)
        except Exception:
            pass

    if not doc.tables:
        raise ValueError("No tables found — verify this is a TEC-004 report.")

    vessel_name = 'UNKNOWN'
    report_date = None
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt:
            continue
        if m := re.search(r"Vessel[\u2019\u2018']?s?\s+Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+?)(?:\s{2,}|\t|Date:|Date\s*:|$)", txt, re.I):
            vessel_name = clean_name(m.group(1))
        if m := re.search(r"Date\s*:\s*(.+)", txt, re.I):
            report_date = parse_date(m.group(1).strip())
        if vessel_name != 'UNKNOWN' and report_date:
            break

    if vessel_name == 'UNKNOWN':
        issues.append(make_issue('warning', 'VESSEL_NOT_FOUND', 'Could not extract vessel name.'))
        warnings.append('Could not extract vessel name.')
    if not report_date:
        issues.append(make_issue('warning', 'REPORT_DATE_NOT_FOUND', 'Could not extract report date.'))

    me_total_hrs = None
    me_this_month = None

    me_grid_rows, aux_grid_rows = [], []
    oe_rows, dg_rows = [], []
    debug_tables = []

    for ti, table in enumerate(doc.tables):
        rg = raw_grid(table)
        dg = dedup_grid(table)

        if me_total_hrs is None or me_this_month is None:
            for row in rg[:4]:
                line = " ".join(str(x) for x in row)
                if me_total_hrs is None:
                    if m := re.search(r'Total Running Hours[\s:ǀ|]+([\d,]+)', line, re.I):
                        me_total_hrs = int(parse_number(m.group(1)))
                if me_this_month is None:
                    if m := re.search(r'This Month[\s:]+([\d,]+)', line, re.I):
                        me_this_month = int(parse_number(m.group(1)))

        raw_type = classify_table(rg)
        dedup_type = classify_table(dg)

        me_chunk = parse_me_from_grid(rg, ti, issues)
        if not me_chunk:
            me_chunk = parse_me_from_grid(dg, ti, issues)
        if me_chunk:
            me_grid_rows.extend(me_chunk)

        aux_chunk = parse_aux_from_grid(rg, ti, issues)
        if not aux_chunk:
            aux_chunk = parse_aux_from_grid(dg, ti, issues)
        if aux_chunk:
            aux_grid_rows.extend(aux_chunk)

        oe_chunk = parse_oe_from_grid(rg, ti)
        if not oe_chunk:
            oe_chunk = parse_oe_from_grid(dg, ti)
        if oe_chunk:
            oe_rows.extend(oe_chunk)

        dg_chunk = parse_dg_from_grid(rg, ti)
        if not dg_chunk:
            dg_chunk = parse_dg_from_grid(dg, ti)
        if dg_chunk:
            dg_rows.extend(dg_chunk)

        debug_tables.append({
            'table_index': ti,
            'raw_type': raw_type,
            'dedup_type': dedup_type,
            'raw_rows': len(rg),
            'raw_cols': max(len(r) for r in rg) if rg else 0,
            'dedup_rows': len(dg),
            'dedup_cols': max(len(r) for r in dg) if dg else 0,
            'me_rows': len(me_chunk),
            'aux_rows': len(aux_chunk),
            'oe_rows': len(oe_chunk),
            'dg_rows': len(dg_chunk),
            'raw_sample': rg[:10],
            'dedup_sample': dg[:10],
        })

    # Text fallback
    all_lines = lines_from_doc(doc)
    me_text_rows = parse_me_from_text(all_lines)
    aux_text_rows = parse_aux_from_text(all_lines)

    if me_text_rows and not me_grid_rows:
        issues.append(make_issue('info', 'ME_TEXT_FALLBACK_USED', 'Main Engine rows were extracted using text fallback.'))
    if aux_text_rows and not aux_grid_rows:
        issues.append(make_issue('info', 'AUX_TEXT_FALLBACK_USED', 'Auxiliary Engine rows were extracted using text fallback.'))

    me_rows = dedupe_component_rows(me_grid_rows + me_text_rows)
    aux_rows = dedupe_component_rows(aux_grid_rows + aux_text_rows)
    components = me_rows + aux_rows

    if not components:
        issues.append(make_issue('error', 'NO_COMPONENTS', 'No ME/AUX components were extracted. Review raw and dedup grids in Parser Debug.'))
        warnings.append('No ME/AUX components extracted.')

    return {
        'vessel_name': vessel_name,
        'report_date': report_date,
        'me_total_hrs': me_total_hrs,
        'me_this_month': me_this_month,
        'components': components,
        'me_comps': me_rows,
        'aux_comps': aux_rows,
        'other_equipment': oe_rows,
        'dg_equipment': dg_rows,
        'warnings': warnings,
        'issues': issues,
        'debug': {
            'tables_scanned': len(doc.tables),
            'table_debug': debug_tables,
            'me_grid_rows_total': len(me_grid_rows),
            'me_text_rows_total': len(me_text_rows),
            'aux_grid_rows_total': len(aux_grid_rows),
            'aux_text_rows_total': len(aux_text_rows),
            'me_rows_total': len(me_rows),
            'aux_rows_total': len(aux_rows),
            'oe_rows_total': len(oe_rows),
            'dg_rows_total': len(dg_rows),
            'all_lines_preview': all_lines[:220],
        },
        'uploaded_at': now_iso(),
        'filename': filename,
    }

# ══════════════════════════════════════════════════════════════════
#  DATABASE IO
# ══════════════════════════════════════════════════════════════════
def report_exists(vessel_name: str, report_date: str, file_hash: str) -> bool:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM reports WHERE vessel_name=? AND report_date=? AND file_hash=?", (vessel_name, report_date, file_hash))
    ok = cur.fetchone() is not None
    conn.close()
    return ok

def save_report(parsed: Dict[str, Any]) -> Tuple[bool, str]:
    conn = get_conn()
    try:
        if report_exists(parsed['vessel_name'], parsed['report_date'], parsed['file_hash']):
            return False, "This report is already saved."

        cur = conn.cursor()
        cur.execute("""
            INSERT INTO reports(
                vessel_name, report_date, filename, file_hash, parser_version,
                me_total_hrs, me_this_month, raw_json, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            parsed['vessel_name'], parsed['report_date'], parsed['filename'], parsed['file_hash'],
            PARSER_VERSION, parsed['me_total_hrs'], parsed['me_this_month'],
            json.dumps(parsed, ensure_ascii=False, default=str), now_iso()
        ))
        report_id = cur.lastrowid

        for r in parsed['components']:
            cur.execute("""
                INSERT INTO parsed_rows(
                    report_id, category, engine_label, unit, description,
                    periodicity, last_oh_date, hrs_since, pct_used, status,
                    source_table_index, source_row_start, source_row_end,
                    source_col_date, source_col_hours, parser_source, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                report_id, r.get('category'), r.get('engine_label'), r.get('unit'), r.get('description'),
                r.get('periodicity'), r.get('last_oh_date'), r.get('hrs_since'), r.get('pct_used'), r.get('status'),
                r.get('source_table_index'), r.get('source_row_start'), r.get('source_row_end'),
                r.get('source_col_date'), r.get('source_col_hours'), r.get('parser_source'), now_iso()
            ))

        for i in parsed.get('issues', []):
            cur.execute("""
                INSERT INTO parse_issues(
                    report_id, severity, issue_code, message, row_key,
                    table_index, row_index, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                report_id, i.get('severity'), i.get('issue_code'), i.get('message'),
                i.get('row_key'), i.get('table_index'), i.get('row_index'), now_iso()
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
          COUNT(DISTINCT pr.id) AS parsed_rows
        FROM reports r
        LEFT JOIN parsed_rows pr ON pr.report_id = r.id
        GROUP BY r.id
        ORDER BY r.id DESC
        LIMIT {int(limit)}
        """, conn)
        return df
    finally:
        conn.close()

# ══════════════════════════════════════════════════════════════════
#  DATAFRAME RENDERING
# ══════════════════════════════════════════════════════════════════
def build_component_df(records: List[Dict[str, Any]], include_trace: bool = False, mode: str = 'matrix') -> pd.DataFrame:
    cols = ['Status','Component','Engine','Unit','Periodicity','Last O/H','Hrs Since','Used %','Parser']
    if include_trace:
        cols += ['Table','Rows','Cols']
    if not records:
        return pd.DataFrame(columns=cols)

    df = pd.DataFrame(records)
    for c, d in {
        'status':'NO DATA','description':'','engine_label':'','unit':'',
        'periodicity':0.0,'last_oh_date':'','hrs_since':0.0,'pct_used':0.0,
        'source_table_index':'','source_row_start':'','source_row_end':'',
        'source_col_date':'','source_col_hours':'','parser_source':''
    }.items():
        if c not in df.columns:
            df[c] = d

    df['_s'] = df['status'].map(lambda x: STATUS_ORDER.get(str(x), 9))
    df['_p'] = df['pct_used'].apply(sf)

    if mode == 'matrix':
        df['_k1'] = df['description'].astype(str).str.upper()
        df['_k2'] = df['engine_label'].astype(str)
        df['_k3'] = df['unit'].apply(cyl_n)
        df = df.sort_values(['_k1','_k2','_k3']).drop(columns=['_s','_p','_k1','_k2','_k3'])
    elif mode == 'priority':
        df = df.sort_values(['_s','_p'], ascending=[True, False]).drop(columns=['_s','_p'])
    else:
        df = df.sort_values(['parser_source','source_table_index','source_row_start','source_col_date']).drop(columns=['_s','_p'])

    out = pd.DataFrame(index=range(len(df)))
    out['Status'] = df['status'].map(lambda x: STATUS_MAP.get(str(x), '🔵 NO DATA'))
    out['Component'] = df['description'].values
    out['Engine'] = df['engine_label'].values
    out['Unit'] = df['unit'].values
    out['Periodicity'] = [f"{int(float(x))}" if sf(x) > 0 else "—" for x in df['periodicity'].values]
    out['Last O/H'] = [str(x) if x and str(x) not in ('nan','None','') else '—' for x in df['last_oh_date'].values]
    out['Hrs Since'] = [f"{int(float(x))}" if sf(x) > 0 else "—" for x in df['hrs_since'].values]
    out['Used %'] = [round(float(x) * 100, 1) if sf(x) > 0 else 0.0 for x in df['pct_used'].values]
    out['Parser'] = df['parser_source'].values
    if include_trace:
        out['Table'] = [str(x) if str(x) != 'nan' else '—' for x in df['source_table_index'].values]
        out['Rows'] = [f"{a}-{b}" if str(a) != 'nan' and str(b) != 'nan' else '—'
                       for a, b in zip(df['source_row_start'].values, df['source_row_end'].values)]
        out['Cols'] = [f"{a}/{b}" if str(a) != 'nan' and str(b) != 'nan' else '—'
                       for a, b in zip(df['source_col_date'].values, df['source_col_hours'].values)]
    return out

def build_issue_df(issues: List[Dict[str, Any]]) -> pd.DataFrame:
    if not issues:
        return pd.DataFrame(columns=['Severity','Code','Message','Table','Row','Row Key'])
    df = pd.DataFrame(issues)
    for c in ['severity','issue_code','message','table_index','row_index','row_key']:
        if c not in df.columns:
            df[c] = ''
    df['_ord'] = df['severity'].map(lambda x: ISSUE_ORDER.get(str(x), 9))
    df = df.sort_values(['_ord','table_index','row_index'])
    out = pd.DataFrame()
    out['Severity'] = df['severity']
    out['Code'] = df['issue_code']
    out['Message'] = df['message']
    out['Table'] = df['table_index']
    out['Row'] = df['row_index']
    out['Row Key'] = df['row_key']
    return out

def build_oe_df(records: List[Dict[str, Any]]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame(columns=['Section','Description','Engine','Periodicity','Last Date / O/H','Run Hrs','Table','Row'])
    df = pd.DataFrame(records)
    out = pd.DataFrame()
    out['Section'] = df.get('section', '')
    out['Description'] = df.get('description', '')
    out['Engine'] = df.get('engine_label', '')
    out['Periodicity'] = [f"{int(float(x))}" if sf(x) > 0 else "—" for x in df.get('periodicity', pd.Series([0]*len(df)))]
    out['Last Date / O/H'] = [str(x) if x else '—' for x in df.get('last_date', pd.Series(['']*len(df)))]
    out['Run Hrs'] = [str(int(x)) if sf(x) > 0 else "—" for x in df.get('run_hrs', pd.Series([0]*len(df)))]
    out['Table'] = df.get('source_table_index', '')
    out['Row'] = df.get('source_row', '')
    return out

CC = {
    'Status': st.column_config.TextColumn('Status', width=145),
    'Component': st.column_config.TextColumn('Component', width=270),
    'Engine': st.column_config.TextColumn('Engine', width=88),
    'Unit': st.column_config.TextColumn('Unit', width=74),
    'Periodicity': st.column_config.TextColumn('Periodicity', width=100),
    'Last O/H': st.column_config.TextColumn('Last O/H', width=115),
    'Hrs Since': st.column_config.TextColumn('Hrs Since', width=100),
    'Used %': st.column_config.ProgressColumn('Used %', min_value=0, max_value=150, format='%.1f%%', width=134),
    'Parser': st.column_config.TextColumn('Parser', width=70),
    'Table': st.column_config.TextColumn('Table', width=60),
    'Rows': st.column_config.TextColumn('Rows', width=75),
    'Cols': st.column_config.TextColumn('Cols', width=80),
}

def show_component_matrix(records: List[Dict[str, Any]], include_trace=False, mode='matrix', height=None):
    df = build_component_df(records, include_trace=include_trace, mode=mode)
    if df.empty:
        st.info("No data.")
        return
    cfg = {k: v for k, v in CC.items() if k in df.columns}
    h = height or min(860, 38 * len(df) + 44)
    st.dataframe(df, use_container_width=True, hide_index=True, height=h, column_config=cfg)

# ══════════════════════════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════════════════════════
if 'parsed_reports' not in st.session_state:
    st.session_state.parsed_reports = []
if 'active_report_hash' not in st.session_state:
    st.session_state.active_report_hash = None
if 'save_feedback' not in st.session_state:
    st.session_state.save_feedback = {}

# ══════════════════════════════════════════════════════════════════
#  HEADER
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero-k">Fleet Running Hours Monitor</div>
<div class="hero-h">TEC‑004 10.10 Debug-Operational Build</div>
<div class="hero-s">
Debug-first parser with raw-grid and dedup-grid visibility, dual extraction paths for Main Engine and Auxiliary Engines,
matrix preview before save, and SQLite-backed report history.
</div>
<div class="hero-rule"></div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
#  UPLOAD
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="panel upload-panel">', unsafe_allow_html=True)
uc, ic = st.columns([2.1, 1.0], gap='large')
with uc:
    uploaded_files = st.file_uploader(
        'Upload TEC-004 .doc files',
        type=['doc'],
        accept_multiple_files=True,
        label_visibility='collapsed'
    )
    st.markdown('<div class="muted">This build never leaves you blind with just “No data.” It always stores raw extraction previews, deduplicated table previews, parser classification, and fallback results so you can see exactly what the runtime produced.</div>', unsafe_allow_html=True)
with ic:
    st.markdown("""
    <div class="muted">
    <b>Extraction paths</b><br>
    Grid parser + text fallback<br><br>
    <b>Debug</b><br>
    Raw grid · Dedup grid · Classification<br><br>
    <b>Persistence</b><br>
    SQLite report ledger
    </div>
    """, unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_files:
    existing_hashes = {r.get('file_hash') for r in st.session_state.parsed_reports}
    new_files = [f for f in uploaded_files if md5_bytes(f.getvalue()) not in existing_hashes]
    if new_files:
        prog = st.progress(0)
        for i, uploaded in enumerate(new_files, start=1):
            raw = uploaded.getvalue()
            fh = md5_bytes(raw)
            with st.spinner(f"Parsing {uploaded.name}..."):
                try:
                    docx = convert_doc_to_docx(raw)
                    parsed = parse_doc_bytes(docx, filename=uploaded.name)
                    parsed['file_hash'] = fh
                    parsed['filename'] = uploaded.name
                    st.session_state.parsed_reports.append(parsed)
                    if st.session_state.active_report_hash is None:
                        st.session_state.active_report_hash = fh
                except Exception as e:
                    st.session_state.parsed_reports.append({
                        'vessel_name': 'UNKNOWN',
                        'report_date': '',
                        'me_total_hrs': 0,
                        'me_this_month': 0,
                        'components': [],
                        'me_comps': [],
                        'aux_comps': [],
                        'other_equipment': [],
                        'dg_equipment': [],
                        'warnings': [],
                        'issues': [make_issue('error', 'PARSE_FAILURE', f'{uploaded.name}: {e}')],
                        'debug': {},
                        'uploaded_at': now_iso(),
                        'filename': uploaded.name,
                        'file_hash': fh,
                    })
                    if st.session_state.active_report_hash is None:
                        st.session_state.active_report_hash = fh
            prog.progress(i / len(new_files))
        prog.empty()

reports = st.session_state.parsed_reports
st.markdown("## Report Queue")
if not reports:
    st.info("Upload one or more TEC‑004 reports to begin.")
    st.stop()

queue_rows = []
for r in reports:
    errs = sum(1 for x in r.get('issues', []) if x.get('severity') == 'error')
    warns = sum(1 for x in r.get('issues', []) if x.get('severity') == 'warning')
    queue_rows.append({
        'Active': '●' if r.get('file_hash') == st.session_state.active_report_hash else '',
        'Filename': r.get('filename', ''),
        'Vessel': r.get('vessel_name', 'UNKNOWN'),
        'Report Date': r.get('report_date') or '—',
        'ME+AUX Rows': len(r.get('components', [])),
        'Warnings': warns,
        'Errors': errs,
        'Hash': (r.get('file_hash', '')[:18] + '…') if r.get('file_hash') else ''
    })
queue_df = pd.DataFrame(queue_rows)
st.dataframe(
    queue_df,
    use_container_width=True,
    hide_index=True,
    height=min(340, 38 * len(queue_df) + 44),
    column_config={
        'Active': st.column_config.TextColumn('Active', width=55),
        'Filename': st.column_config.TextColumn('Filename', width=260),
        'Vessel': st.column_config.TextColumn('Vessel', width=160),
        'Report Date': st.column_config.TextColumn('Report Date', width=120),
        'ME+AUX Rows': st.column_config.NumberColumn('ME+AUX Rows', width=100),
        'Warnings': st.column_config.NumberColumn('Warnings', width=80),
        'Errors': st.column_config.NumberColumn('Errors', width=70),
        'Hash': st.column_config.TextColumn('Hash', width=180),
    }
)

selector = {f"{r.get('filename','')} | {r.get('vessel_name','UNKNOWN')} | {r.get('report_date') or '—'}": r.get('file_hash') for r in reports}
selected_label = st.selectbox(
    "Active report",
    list(selector.keys()),
    index=list(selector.values()).index(st.session_state.active_report_hash) if st.session_state.active_report_hash in selector.values() else 0
)
st.session_state.active_report_hash = selector[selected_label]
active = next(r for r in reports if r.get('file_hash') == st.session_state.active_report_hash)

me = active.get('me_comps', [])
aux = active.get('aux_comps', [])
oe = active.get('other_equipment', [])
dg = active.get('dg_equipment', [])
all_ = active.get('components', [])
issues = active.get('issues', [])
debug = active.get('debug', {})

n_od = sum(1 for c in all_ if c.get('status') == 'OVERDUE')
n_hp = sum(1 for c in all_ if c.get('status') == 'HIGH PRIORITY')
n_ok = sum(1 for c in all_ if c.get('status') == 'OK')
n_err = sum(1 for x in issues if x.get('severity') == 'error')
n_warn = sum(1 for x in issues if x.get('severity') == 'warning')

st.markdown(f"""
<div class="metric-grid">
  <div class="metric b"><div class="metric-v">{active.get('vessel_name','UNKNOWN')}</div><div class="metric-l">Vessel</div></div>
  <div class="metric"><div class="metric-v">{active.get('report_date') or '—'}</div><div class="metric-l">Report Date</div></div>
  <div class="metric c"><div class="metric-v">{active.get('me_total_hrs') or 0:,}</div><div class="metric-l">ME Total Hrs</div></div>
  <div class="metric c"><div class="metric-v">{active.get('me_this_month') or 0:,}</div><div class="metric-l">ME This Month</div></div>
  <div class="metric"><div class="metric-v">{len(me)}</div><div class="metric-l">ME Rows</div></div>
  <div class="metric"><div class="metric-v">{len(aux)}</div><div class="metric-l">AUX Rows</div></div>
  <div class="metric o"><div class="metric-v">{n_hp}</div><div class="metric-l">High Priority</div></div>
  <div class="metric r"><div class="metric-v">{n_od}</div><div class="metric-l">Overdue</div></div>
</div>
""", unsafe_allow_html=True)

if n_err:
    st.markdown(f'<div class="banner err"><b>{n_err}</b> parser errors detected. Open Parser Debug to inspect raw tables and fallback results.</div>', unsafe_allow_html=True)
elif n_warn:
    st.markdown(f'<div class="banner warn"><b>{n_warn}</b> parser warnings detected. Parsed rows are still shown where either extraction path succeeded.</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="banner ok">No parser issues detected for the active report.</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
#  ACTIONS
# ══════════════════════════════════════════════════════════════════
a1, a2, a3, a4 = st.columns([1.2, 1.3, 1.6, 2.7])
with a1:
    review_complete = st.checkbox("Review complete", value=False)
with a2:
    if st.button("Save active report", disabled=not review_complete):
        ok, msg = save_report(active)
        st.session_state.save_feedback[active['file_hash']] = (ok, msg)
with a3:
    export_df = build_component_df(all_, include_trace=True, mode='source')
    st.download_button(
        "Download active CSV",
        data=export_df.to_csv(index=False).encode('utf-8'),
        file_name=f"{Path(active.get('filename','report')).stem}_preview.csv",
        mime="text/csv"
    )
with a4:
    st.markdown(f'<div class="micro">Parser {PARSER_VERSION} • File {active.get("filename","")} • Hash {active.get("file_hash","")[:18]}…</div>', unsafe_allow_html=True)

fb = st.session_state.save_feedback.get(active.get('file_hash'))
if fb:
    if fb[0]:
        st.success(fb[1])
    else:
        st.warning(fb[1])

# ══════════════════════════════════════════════════════════════════
#  TABS
# ══════════════════════════════════════════════════════════════════
tab_me, tab_aux, tab_oe, tab_dg, tab_issues, tab_debug, tab_history = st.tabs([
    "Main Engine", "Auxiliary Engines", "Other Equipment", "D/G Equipment", "Parse Issues", "Parser Debug", "Saved Reports"
])

with tab_me:
    st.markdown(f'<div class="bar"><div class="bar-i">⚙</div><div><div class="bar-t">Main Engine Matrix</div><div class="bar-s">grid parser + text fallback</div></div><div class="bar-r"></div><div class="bar-b">rows: {len(me)}</div></div>', unsafe_allow_html=True)
    mf1, mf2, mf3, mf4 = st.columns([2.0, 1.8, 1.4, 2.2])
    with mf1:
        me_comp = st.selectbox('Component', ['All'] + sorted({c['description'] for c in me}), key='me_comp')
    with mf2:
        me_status = st.selectbox('Status', ['All','Overdue only','High Priority +','OK only'], key='me_status')
    with mf3:
        me_trace = st.checkbox('Show trace', value=True, key='me_trace')
    with mf4:
        me_sort = st.radio('Sort', ['Component → Cylinder', 'Priority → % Used', 'Source order'], horizontal=True, key='me_sort')

    me_view = me[:]
    if me_comp != 'All':
        me_view = [c for c in me_view if c['description'] == me_comp]
    if me_status == 'Overdue only':
        me_view = [c for c in me_view if c['status'] == 'OVERDUE']
    elif me_status == 'High Priority +':
        me_view = [c for c in me_view if c['status'] in ('OVERDUE', 'HIGH PRIORITY')]
    elif me_status == 'OK only':
        me_view = [c for c in me_view if c['status'] == 'OK']

    mode = 'matrix' if 'Component' in me_sort else ('priority' if 'Priority' in me_sort else 'source')
    show_component_matrix(me_view, include_trace=me_trace, mode=mode)

with tab_aux:
    st.markdown(f'<div class="bar"><div class="bar-i">🔩</div><div><div class="bar-t">Auxiliary Engines Matrix</div><div class="bar-s">grid parser + text fallback</div></div><div class="bar-r"></div><div class="bar-b">rows: {len(aux)}</div></div>', unsafe_allow_html=True)
    af1, af2, af3, af4, af5 = st.columns([1.4, 2.0, 1.8, 1.4, 2.2])
    with af1:
        ax_eng = st.selectbox('Engine', ['All'] + sorted({c['engine_label'] for c in aux}), key='ax_eng')
    with af2:
        ax_comp = st.selectbox('Component', ['All'] + sorted({c['description'] for c in aux}), key='ax_comp')
    with af3:
        ax_status = st.selectbox('Status', ['All','Overdue only','High Priority +','OK only'], key='ax_status')
    with af4:
        ax_trace = st.checkbox('Show trace', value=True, key='ax_trace')
    with af5:
        ax_sort = st.radio('Sort', ['Component → Cylinder', 'Priority → % Used', 'Source order'], horizontal=True, key='ax_sort')

    ax_view = aux[:]
    if ax_eng != 'All':
        ax_view = [c for c in ax_view if c['engine_label'] == ax_eng]
    if ax_comp != 'All':
        ax_view = [c for c in ax_view if c['description'] == ax_comp]
    if ax_status == 'Overdue only':
        ax_view = [c for c in ax_view if c['status'] == 'OVERDUE']
    elif ax_status == 'High Priority +':
        ax_view = [c for c in ax_view if c['status'] in ('OVERDUE', 'HIGH PRIORITY')]
    elif ax_status == 'OK only':
        ax_view = [c for c in ax_view if c['status'] == 'OK']

    mode = 'matrix' if 'Component' in ax_sort else ('priority' if 'Priority' in ax_sort else 'source')
    show_component_matrix(ax_view, include_trace=ax_trace, mode=mode)

with tab_oe:
    st.markdown(f'<div class="bar"><div class="bar-i">🛠</div><div><div class="bar-t">Other Equipment</div><div class="bar-s">grid extraction</div></div><div class="bar-r"></div><div class="bar-b">rows: {len(oe)}</div></div>', unsafe_allow_html=True)
    oe_df = build_oe_df(oe)
    if oe_df.empty:
        st.info("No other equipment data found.")
    else:
        st.dataframe(
            oe_df, use_container_width=True, hide_index=True,
            height=min(760, 38 * len(oe_df) + 44)
        )

with tab_dg:
    st.markdown(f'<div class="bar"><div class="bar-i">⚡</div><div><div class="bar-t">D/G Equipment</div><div class="bar-s">grid extraction</div></div><div class="bar-r"></div><div class="bar-b">rows: {len(dg)}</div></div>', unsafe_allow_html=True)
    dg_df = build_oe_df(dg)
    if dg_df.empty:
        st.info("No D/G equipment data found.")
    else:
        st.dataframe(
            dg_df, use_container_width=True, hide_index=True,
            height=min(760, 38 * len(dg_df) + 44)
        )

with tab_issues:
    st.markdown(f'<div class="bar"><div class="bar-i">⚠</div><div><div class="bar-t">Parse Issues</div><div class="bar-s">warnings and extraction failures</div></div><div class="bar-r"></div><div class="bar-b">issues: {len(issues)}</div></div>', unsafe_allow_html=True)
    idf = build_issue_df(issues)
    if idf.empty:
        st.success("No issues found.")
    else:
        sev = st.multiselect('Severity', ['error','warning','info'], default=['error','warning','info'])
        if sev:
            idf = idf[idf['Severity'].isin(sev)]
        st.dataframe(
            idf, use_container_width=True, hide_index=True,
            height=min(700, 38 * len(idf) + 44)
        )

with tab_debug:
    st.markdown('<div class="bar"><div class="bar-i">🧪</div><div><div class="bar-t">Parser Debug</div><div class="bar-s">raw grid · dedup grid · text fallback preview</div></div><div class="bar-r"></div><div class="bar-b">debug</div></div>', unsafe_allow_html=True)
    d1, d2, d3, d4, d5, d6 = st.columns(6)
    with d1:
        st.metric("Tables scanned", debug.get('tables_scanned', 0))
    with d2:
        st.metric("ME grid", debug.get('me_grid_rows_total', 0))
    with d3:
        st.metric("ME text", debug.get('me_text_rows_total', 0))
    with d4:
        st.metric("AUX grid", debug.get('aux_grid_rows_total', 0))
    with d5:
        st.metric("AUX text", debug.get('aux_text_rows_total', 0))
    with d6:
        st.metric("Final ME+AUX", len(all_))

    tdbg = pd.DataFrame(debug.get('table_debug', []))
    if not tdbg.empty:
        show_cols = [c for c in ['table_index','raw_type','dedup_type','raw_rows','raw_cols','dedup_rows','dedup_cols','me_rows','aux_rows','oe_rows','dg_rows'] if c in tdbg.columns]
        st.markdown("### Table classification")
        st.dataframe(tdbg[show_cols], use_container_width=True, hide_index=True)

        st.markdown("### Raw / dedup previews")
        for item in debug.get('table_debug', []):
            with st.expander(f"Table {item['table_index']} · raw {item['raw_type']} · dedup {item['dedup_type']}"):
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("**Raw grid**")
                    st.dataframe(pd.DataFrame(item['raw_sample']), use_container_width=True, hide_index=True)
                with c2:
                    st.markdown("**Dedup grid**")
                    st.dataframe(pd.DataFrame(item['dedup_sample']), use_container_width=True, hide_index=True)

    lines = debug.get('all_lines_preview', [])
    if lines:
        st.markdown("### Text stream preview")
        st.dataframe(pd.DataFrame({"line": lines}), use_container_width=True, hide_index=True, height=420)

with tab_history:
    st.markdown('<div class="bar"><div class="bar-i">🗄</div><div><div class="bar-t">Saved Reports</div><div class="bar-s">SQLite-backed report history</div></div><div class="bar-r"></div><div class="bar-b">history</div></div>', unsafe_allow_html=True)
    hist = load_recent_reports(30)
    if hist.empty:
        st.info("No saved reports in the database yet.")
    else:
        st.dataframe(
            hist, use_container_width=True, hide_index=True,
            height=min(700, 38 * len(hist) + 44)
        )
