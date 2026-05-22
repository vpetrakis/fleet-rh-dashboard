import streamlit as st
st.set_page_config(
    page_title="Fleet Running Hours Monitor v10.10 Hybrid",
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
from typing import Any, Dict, List

import pandas as pd

# ══════════════════════════════════════════════════════════════════
#  PAGE STYLE
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
  --ink:#060b12; --ink2:#0b111b; --ink3:#101826; --ink4:#152031;
  --line:#1b2a3f; --line2:#2b415e;
  --gold:#b8870c; --gold2:#d7a727; --gold3:#edc252;
  --red:#ef5350; --ora:#ffb74d; --grn:#66bb6a; --blu:#64b5f6; --cyn:#4dd0e1;
  --t0:#e8f0fb; --t1:#a8bfd9; --t2:#6f87a1; --t3:#41566f;
  --ff:'Space Grotesk', sans-serif; --fi:'Inter', sans-serif; --fm:'JetBrains Mono', monospace;
  --r:14px;
}
html, body, [class*="css"] {
  background: var(--ink) !important;
  color: var(--t1) !important;
  font-family: var(--fi) !important;
}
.main, .main > div, .block-container { background: var(--ink) !important; }
.block-container { max-width: 100% !important; padding: 1.1rem 1.7rem 3rem !important; }
[data-testid="collapsedControl"], [data-testid="stSidebar"] { display:none !important; }

.main::before{
  content:"";
  position:fixed; inset:0; pointer-events:none; z-index:0;
  background:
    radial-gradient(ellipse 70% 45% at 0% 0%, rgba(184,135,12,.08), transparent 60%),
    radial-gradient(ellipse 60% 40% at 100% 100%, rgba(77,208,225,.05), transparent 55%);
}
.block-container > * { position:relative; z-index:1; }

.hero-k {
  font-size:.67rem; letter-spacing:.24em; text-transform:uppercase; color:var(--gold3); font-weight:700;
}
.hero-h {
  font-family:var(--ff); font-size:2rem; font-weight:700; color:var(--t0); letter-spacing:-.04em; line-height:1.05; margin-top:.24rem;
}
.hero-s {
  color:var(--t1); font-size:.93rem; line-height:1.65; margin-top:.5rem; max-width:1100px;
}
.hero-rule {
  height:1px; margin:.85rem 0 1.15rem; background:linear-gradient(90deg,var(--gold2),var(--line),transparent);
}
.panel {
  background:linear-gradient(180deg, rgba(255,255,255,.02), rgba(255,255,255,.01));
  border:1px solid var(--line);
  border-radius:var(--r);
  padding:1rem;
  box-shadow:0 18px 50px rgba(0,0,0,.35);
}
.upload-panel {
  background:
    linear-gradient(180deg, rgba(184,135,12,.05), rgba(100,181,246,.03)),
    linear-gradient(180deg, rgba(16,24,38,.96), rgba(11,17,27,.96));
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
  grid-template-columns:repeat(7,1fr);
  gap:.8rem;
  margin:1rem 0 1rem;
}
.metric{
  background:linear-gradient(180deg,var(--ink3),var(--ink2));
  border:1px solid var(--line);
  border-radius:var(--r);
  padding:.85rem .95rem .95rem;
  box-shadow:0 18px 50px rgba(0,0,0,.28);
  position:relative;
  overflow:hidden;
}
.metric::before{
  content:"";
  position:absolute; top:0; left:0; right:0; height:2px;
  background:linear-gradient(90deg,var(--gold2),transparent 75%);
}
.metric.r::before{ background:linear-gradient(90deg,var(--red),transparent 75%); }
.metric.o::before{ background:linear-gradient(90deg,var(--ora),transparent 75%); }
.metric.g::before{ background:linear-gradient(90deg,var(--grn),transparent 75%); }
.metric.b::before{ background:linear-gradient(90deg,var(--blu),transparent 75%); }
.metric-v{
  font-family:var(--ff); font-size:1.52rem; font-weight:700; color:var(--t0); letter-spacing:-.04em; line-height:1.05;
}
.metric-l{
  color:var(--t3); font-size:.6rem; text-transform:uppercase; letter-spacing:.18em; margin-top:.34rem;
}
.bar{
  display:flex; align-items:center; gap:.8rem; margin:1.45rem 0 .8rem;
}
.bar-i{
  width:34px; height:34px; border-radius:10px; display:flex; align-items:center; justify-content:center;
  background:rgba(184,135,12,.1); border:1px solid rgba(184,135,12,.2);
}
.bar-t{ font-family:var(--ff); font-size:1rem; font-weight:700; color:var(--t0); }
.bar-s{ font-family:var(--fm); font-size:.61rem; color:var(--t3); margin-top:2px; }
.bar-r{ flex:1; height:1px; background:linear-gradient(90deg,var(--line),transparent); }
.bar-b{
  font-family:var(--fm); font-size:.62rem; color:var(--t2);
  border:1px solid var(--line2); background:var(--ink3); border-radius:999px; padding:4px 9px;
}
.banner{
  border-radius:var(--r); padding:.85rem 1rem; border:1px solid var(--line); margin-bottom:.85rem; line-height:1.55;
  box-shadow:0 18px 50px rgba(0,0,0,.22);
}
.banner.ok{ background:rgba(102,187,106,.10); border-color:rgba(102,187,106,.28); color:#dff7e2; }
.banner.warn{ background:rgba(255,183,77,.10); border-color:rgba(255,183,77,.28); color:#ffe8c7; }
.banner.err{ background:rgba(239,83,80,.10); border-color:rgba(239,83,80,.28); color:#ffd8d8; }

.muted{ color:var(--t1); font-size:.88rem; line-height:1.65; }
.micro{ font-family:var(--fm); color:var(--t3); font-size:.72rem; }

[data-testid="stDataFrame"]{
  border:1px solid var(--line)!important;
  border-radius:var(--r)!important;
  overflow:hidden!important;
  box-shadow:0 18px 50px rgba(0,0,0,.26)!important;
}
.dvn-scroller{ background:var(--ink2)!important; }

.streamlit-expanderHeader{
  background:var(--ink3)!important; border:1px solid var(--line)!important; border-radius:12px!important; color:var(--t0)!important;
}
.streamlit-expanderContent{
  background:var(--ink2)!important; border:1px solid var(--line)!important; border-top:none!important; border-radius:0 0 12px 12px!important;
}
.stButton>button{
  background:linear-gradient(135deg,var(--gold3),var(--gold2))!important;
  color:#081018!important; border:none!important; border-radius:11px!important;
  padding:.72rem 1.1rem!important; font-weight:800!important; text-transform:uppercase!important; letter-spacing:.05em!important;
}
.stDownloadButton>button{
  background:var(--ink3)!important; color:var(--t0)!important; border:1px solid var(--line2)!important; border-radius:11px!important;
  padding:.72rem 1rem!important; font-weight:700!important;
}
div[data-baseweb="select"] > div,
.stTextInput > div > div > input{
  background:var(--ink4)!important; color:var(--t0)!important; border:1px solid var(--line2)!important; border-radius:10px!important;
}
.stSelectbox label,.stMultiSelect label,.stCheckbox label,.stRadio label{
  color:var(--t3)!important; font-size:.66rem!important; text-transform:uppercase!important; letter-spacing:.12em!important;
}
@media (max-width:1280px){ .metric-grid{ grid-template-columns:repeat(4,1fr);} }
@media (max-width:840px){ .metric-grid{ grid-template-columns:repeat(2,1fr);} }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
#  CONSTANTS / DB
# ══════════════════════════════════════════════════════════════════
DB_PATH = "tec004_hybrid_v1010.db"
PARSER_VERSION = "hybrid_v10_10_original_parser"

_STATUS_ORD = {'OVERDUE': 0, 'HIGH PRIORITY': 1, 'OK': 2, 'NO DATA': 3}
_STATUS_MAP = {
    'OVERDUE': '🔴 OVERDUE',
    'HIGH PRIORITY': '🟠 HIGH PRIORITY',
    'OK': '🟢 OK',
    'NO DATA': '🔵 NO DATA'
}
_ISSUE_ORD = {'error': 0, 'warning': 1, 'info': 2}

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
        created_at TEXT NOT NULL,
        FOREIGN KEY(report_id) REFERENCES reports(id)
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

def _sf(x):
    try:
        v = float(x)
        return 0.0 if pd.isna(v) else v
    except Exception:
        return 0.0

def _cyl_n(u: str) -> int:
    m = re.search(r'\d+', str(u))
    return int(m.group()) if m else 999

def _fl(txt: str) -> str:
    for part in re.split(r'[\r\n\x0b]+', str(txt or '')):
        s = re.sub(r'[\x07\xa0\t]+', ' ', part).strip()
        if s:
            return re.sub(r'  +', ' ', s)
    return ''

def _clean_name(txt: str) -> str:
    t = _fl(txt)
    t = re.sub(r'(?i)ALEXIS\s*Date?', '', t)
    t = re.sub(r'(?i)Page\s*\d+\s*of\s*\d+', '', t)
    return re.sub(r'  +', ' ', t).strip()

def _is_comp(name: str) -> bool:
    u = name.upper().strip()
    if not u or len(u) < 2:
        return False
    for bad in ('DESCRIPTION','REMARKS','COMPONENT','-','PERIODICITY',
                'PERIODICTLY','DATE OF LAST','RUNNING HOURS','TYPE:',
                'AUX. ENGINE','TURBOCHARGER','MAIN ENGINE'):
        if bad in u:
            return False
    if re.fullmatch(r'[\d./ ,:\-\[\]]+', u):
        return False
    if len(u) > 55:
        return False
    return bool(re.search(r'[A-Za-z]', u))

def _date(txt: str) -> str:
    s = _fl(txt).strip()
    if not s or s in ('-', '1', '2', '1/', '/ 2', 'N/A', 'n/a'):
        return ''
    if len(s) > 20:
        return ''
    if re.match(r'^\d+$', s):
        return ''
    return s

def _num(txt: str) -> float:
    s = _fl(txt).strip().upper()
    if not s or s in ('', '-', 'N/A', 'CENTRAL'):
        return 0.0
    s = s.replace('[', '').replace(']', '')
    if any(w in s for w in ('MONTH','YEAR','WEEK','DAY','OBS')):
        return 0.0
    s = re.sub(r'([,\.])\s+', r'\1', s)
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

def _status(hrs: float, period: float) -> str:
    if hrs <= 0 or period <= 0:
        return 'NO DATA'
    r = hrs / period
    if r >= 1.0:
        return 'OVERDUE'
    if r >= 0.8:
        return 'HIGH PRIORITY'
    return 'OK'

def _pct(hrs: float, period: float) -> float:
    return round(hrs / period, 4) if period and hrs else 0.0

def _make(cat, eng, unit, name, period, date, hrs, table_idx=None, row_start=None, row_end=None) -> dict:
    return {
        'category': cat,
        'engine_label': eng,
        'unit': unit,
        'description': name,
        'periodicity': period,
        'last_oh_date': date,
        'hrs_since': hrs,
        'pct_used': _pct(hrs, period),
        'status': _status(hrs, period),
        'source_table_index': table_idx,
        'source_row_start': row_start,
        'source_row_end': row_end,
    }

# ══════════════════════════════════════════════════════════════════
#  CONVERSION
# ══════════════════════════════════════════════════════════════════
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found — packages.txt/environment must contain libreoffice.")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as t:
        t.write(raw)
        src = t.name
    outdir = tempfile.mkdtemp(prefix="lo_")
    out = os.path.join(outdir, Path(src).stem + ".docx")
    pf = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        r = subprocess.run(
            [soffice, "--headless", "--norestore", "--nofirststartwizard",
             f"-env:UserInstallation={pf}", "--convert-to", "docx",
             src, "--outdir", outdir],
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

# ══════════════════════════════════════════════════════════════════
#  ORIGINAL PARSER CORE — PRESERVED
# ══════════════════════════════════════════════════════════════════
def _rect_grid(table) -> list:
    """Creates a rectangular grid exactly like the original baseline parser."""
    grid = []
    if not table.rows:
        return grid
    max_cols = max(len(row.cells) for row in table.rows)
    for row in table.rows:
        cells = []
        for cell in row.cells:
            raw = re.sub(r'[\x0b\r]', '\n', cell.text).replace('\x07', '')
            lines = [ln.replace('\xa0', ' ').replace('\t', ' ').strip()
                     for ln in raw.split('\n') if ln.strip()]
            cells.append(lines[0] if lines else '')
        while len(cells) < max_cols:
            cells.append("")
        grid.append(cells)
    return grid

def _parse_me(grid: list, table_idx: int = None) -> list:
    if len(grid) < 3:
        return []
    is_me = False
    for r in range(min(3, len(grid))):
        if any('MAIN ENGINE' in str(c).upper() for c in grid[r]):
            is_me = True
    if not is_me:
        return []

    per_col, rem_col = 1, len(grid[0]) - 1
    for c in range(len(grid[0])):
        h_txt = "".join(str(grid[r][c]).upper() for r in range(min(3, len(grid))))
        if 'REMARK' in h_txt:
            rem_col = c

    actual_cyls = max(1, min(7, rem_col - per_col - 1))
    FIRST_CYL  = per_col + 2
    MARKER_COL = per_col + 1

    end = len(grid)
    for r, row in enumerate(grid):
        f_row = ' '.join(str(x) for x in row).upper()
        if any(x in f_row for x in ('NOTE 1','TURBOCHARGER','AUX. ENGINE')):
            end = r
            break

    result, r = [], 1
    while r < end - 1:
        name   = _clean_name(grid[r][0] if grid[r] else '')
        period = _num(grid[r][per_col] if per_col < len(grid[r]) else '')
        marker = str(grid[r][MARKER_COL] if MARKER_COL < len(grid[r]) else '').strip()

        if _is_comp(name) and '1' in marker:
            nxt = grid[r + 1] if r + 1 < len(grid) else []
            for cyl in range(1, actual_cyls + 1):
                ci   = FIRST_CYL + cyl - 1
                date = _date(grid[r][ci] if ci < len(grid[r]) else '')
                hrs  = _num(nxt[ci] if ci < len(nxt) else '')
                if date or hrs > 0:
                    result.append(_make('MAIN_ENGINE','ME',f'Cyl {cyl}',name,period,date,hrs,table_idx,r,r+1))
            r += 2
        else:
            r += 1
    return result

def _parse_aux(grid: list, table_idx: int = None) -> list:
    desc_row = None
    for ri, row in enumerate(grid):
        if 'DESCRIPTION' in str(row[0] if row else '').upper():
            desc_row = ri
            break
    if desc_row is None:
        return []

    actual_cyls = 0
    seen_nums = set()
    for ci in range(3, len(grid[desc_row])):
        txt = str(grid[desc_row][ci]).strip()
        if txt and re.match(r'^\d+$', txt):
            n = int(txt)
            if n not in seen_nums:
                if n == actual_cyls + 1:
                    actual_cyls = n
                    seen_nums.add(n)
                elif actual_cyls > 0:
                    break
        elif actual_cyls > 0:
            break
    actual_cyls = max(1, min(6, actual_cyls)) if actual_cyls else 6

    AUX1 = 3
    AUX2 = AUX1 + actual_cyls
    AUX3 = AUX2 + actual_cyls

    result, r = [], desc_row + 1
    while r < len(grid) - 1:
        name   = _clean_name(grid[r][0] if grid[r] else '')
        period = _num(grid[r][1] if len(grid[r]) > 1 else '')
        marker = str(grid[r][2] if len(grid[r]) > 2 else '').strip()

        if _is_comp(name) and '1' in marker:
            nxt = grid[r + 1] if r + 1 < len(grid) else []
            for cyl in range(1, actual_cyls + 1):
                for eng, start in (('AUX-1', AUX1), ('AUX-2', AUX2), ('AUX-3', AUX3)):
                    ci   = start + cyl - 1
                    date = _date(grid[r][ci] if ci < len(grid[r]) else '')
                    hrs  = _num(nxt[ci] if ci < len(nxt) else '')
                    if date or hrs > 0:
                        result.append(_make('AUX_ENGINE',eng,f'Cyl {cyl}',name,period,date,hrs,table_idx,r,r+1))
            r += 2
        else:
            r += 1
    return result

def _parse_oe_grid(grid: list, table_idx: int = None) -> list:
    oe = []
    SKIP = {'TURBOCHARGER','AUXILIARY BOILER','COOLERS','EXH GAS  BOILER',
            'A/C & REFR. COMPRESSORS','MAIN AIR COMPRESSORS',
            'PERIODICTLY','DATE OF LAST INSPECTION','RUN HRS',
            'DATE OF LAST CLEANING','DATE','PERIODICITY',''}

    for r, row in enumerate(grid):
        def gc(ci): return str(row[ci]) if ci < len(row) else ''
        for sec, dc, dtc, hrc in [
            ('Turbocharger / Aux Boiler', 0, 1, 3),
            ('Coolers / Exh Gas Boiler',  5, 6, 8),
            ('A/C & Compressors',         10,11,12),
        ]:
            desc = gc(dc).strip()
            if not desc or desc.upper() in SKIP:
                continue
            desc_clean = _clean_name(desc)
            if not _is_comp(desc_clean):
                continue

            dv = gc(dtc)
            hv = gc(hrc)
            if dv or hv:
                oe.append({
                    'section': sec,
                    'description': desc_clean,
                    'last_date': dv,
                    'run_hrs': hv,
                    'source_table_index': table_idx,
                    'source_row': r
                })

    DG = ['D/G 1','D/G 2','D/G 3']
    for ri, row in enumerate(grid):
        if ri == 0:
            continue
        def gc2(ci): return str(row[ci]) if ci < len(row) else ''
        for dc, tc, ds in [(0,2,3),(9,11,12)]:
            desc = _clean_name(gc2(dc))
            rt = gc2(tc).strip()
            if not desc or rt not in ('1','2') or not _is_comp(desc):
                continue

            for gi, gl in enumerate(DG):
                val = gc2(ds + gi)
                if not val:
                    continue
                key = f"{desc} — {gl}"

                if rt == '1':
                    oe.append({
                        'section':'D/G Equipment',
                        'description':key,
                        'last_date':val,
                        'run_hrs':'',
                        'source_table_index': table_idx,
                        'source_row': ri
                    })
                else:
                    for e in reversed(oe):
                        if e['description'] == key and e['run_hrs'] == '':
                            e['run_hrs'] = val
                            break
                    else:
                        oe.append({
                            'section':'D/G Equipment',
                            'description':key,
                            'last_date':'',
                            'run_hrs':val,
                            'source_table_index': table_idx,
                            'source_row': ri
                        })
    return oe

def parse_doc_bytes(docx_bytes: bytes, filename: str = "") -> dict:
    from docx import Document

    warns = []
    issues = []
    debug_tables = []

    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as t:
        t.write(docx_bytes)
        tp = t.name
    try:
        doc = Document(tp)
    except Exception as e:
        raise ValueError(f"Cannot open document: {e}")
    finally:
        try:
            os.unlink(tp)
        except Exception:
            pass

    if not doc.tables:
        raise ValueError("No tables found — is this a TEC-004 report?")

    vn = 'UNKNOWN'
    rd = None
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt:
            continue
        if m := re.search(r"Vessel[\u2019\u2018']?s?\s+Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+?)(?:\s{2,}|\t|Date:|Date\s*:|$)", txt, re.I):
            vn = _clean_name(m.group(1))
        if m := re.search(r"Date\s*:\s*(.+)", txt, re.I):
            rd = _date(m.group(1).strip())
        if vn != 'UNKNOWN' and rd:
            break
    if vn == 'UNKNOWN':
        warns.append("Could not extract vessel name.")
        issues.append({"severity":"warning","issue_code":"VESSEL_NOT_FOUND","message":"Could not extract vessel name.","table_index":None,"row_index":None,"row_key":""})
    if not rd:
        issues.append({"severity":"warning","issue_code":"REPORT_DATE_NOT_FOUND","message":"Could not extract report date.","table_index":None,"row_index":None,"row_key":""})

    me_tot = me_mo = None
    g0 = _rect_grid(doc.tables[0]) if len(doc.tables) > 0 else []
    for cell_txt in (g0[0] if g0 else []):
        if m := re.search(r'Total Running Hours[\s:ǀ|]+([\d,]+)', str(cell_txt), re.I):
            me_tot = int(_num(m.group(1)))
        if m := re.search(r'This Month[\s:]+([\d,]+)', str(cell_txt), re.I):
            me_mo  = int(_num(m.group(1)))

    me_comps, aux_comps, oe = [], [], []
    raw_tables_preview = []

    for ti, table in enumerate(doc.tables):
        grid = _rect_grid(table)
        if not grid:
            continue

        raw_tables_preview.append({
            "table_index": ti,
            "rows": len(grid),
            "cols": max(len(r) for r in grid) if grid else 0,
            "sample": grid[:12]
        })

        mc = _parse_me(grid, ti)
        ac = _parse_aux(grid, ti)
        oc = _parse_oe_grid(grid, ti)

        if mc:
            me_comps.extend(mc)
        if ac:
            aux_comps.extend(ac)
        if oc:
            oe.extend(oc)

        debug_tables.append({
            "table_index": ti,
            "me_rows": len(mc),
            "aux_rows": len(ac),
            "oe_rows": len(oc)
        })

    comps = me_comps + aux_comps
    if not comps:
        warns.append("No components extracted.")
        issues.append({"severity":"error","issue_code":"NO_COMPONENTS","message":"No components extracted. Verify this is a TEC-004 report.","table_index":None,"row_index":None,"row_key":""})

    return {
        'vessel_name': vn,
        'report_date': rd,
        'me_total_hrs': me_tot,
        'me_this_month': me_mo,
        'components': comps,
        'me_comps': me_comps,
        'aux_comps': aux_comps,
        'other_equipment': oe,
        'warnings': warns,
        'issues': issues,
        'debug': {
            'tables_scanned': len(doc.tables),
            'me_rows_total': len(me_comps),
            'aux_rows_total': len(aux_comps),
            'oe_rows_total': len(oe),
            'table_debug': debug_tables,
            'raw_tables_preview': raw_tables_preview[:5],
        },
        'uploaded_at': now_iso(),
        'filename': filename,
    }

# ══════════════════════════════════════════════════════════════════
#  STORAGE / HISTORY
# ══════════════════════════════════════════════════════════════════
def report_exists(vessel_name: str, report_date: str, file_hash: str) -> bool:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM reports WHERE vessel_name=? AND report_date=? AND file_hash=?", (vessel_name, report_date, file_hash))
    found = cur.fetchone() is not None
    conn.close()
    return found

def save_report(parsed: Dict[str, Any]) -> tuple[bool, str]:
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
                    periodicity, last_oh_date, hrs_since, pct_used, status,
                    source_table_index, source_row_start, source_row_end, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                report_id, r.get("category"), r.get("engine_label"), r.get("unit"), r.get("description"),
                r.get("periodicity"), r.get("last_oh_date"), r.get("hrs_since"), r.get("pct_used"), r.get("status"),
                r.get("source_table_index"), r.get("source_row_start"), r.get("source_row_end"), now_iso()
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
#  DATAFRAME BUILDERS
# ══════════════════════════════════════════════════════════════════
def _build(records: list, include_trace: bool = False, mode: str = 'matrix') -> pd.DataFrame:
    cols = ['Status','Component','Engine','Unit', 'Periodicity','Last O/H','Hrs Since','Used %']
    if include_trace:
        cols += ['Table','Rows']
    if not records:
        return pd.DataFrame(columns=cols)

    df = pd.DataFrame(records)

    for c, default in {
        'status':'NO DATA','description':'','engine_label':'','unit':'',
        'periodicity':0.0,'last_oh_date':'','hrs_since':0.0,'pct_used':0.0,
        'source_table_index':'','source_row_start':'','source_row_end':''
    }.items():
        if c not in df.columns:
            df[c] = default

    df['_s'] = df['status'].map(lambda x: _STATUS_ORD.get(str(x), 4))
    df['_p'] = df['pct_used'].apply(lambda x: _sf(x))

    if mode == 'matrix':
        df['_k1'] = df['description'].astype(str).str.upper()
        df['_k2'] = df['unit'].apply(_cyl_n)
        df = df.sort_values(['_k1','_k2']).drop(columns=['_k1','_k2','_s','_p'])
    elif mode == 'priority':
        df = df.sort_values(['_s','_p'], ascending=[True,False]).drop(columns=['_s','_p'])
    else:
        df = df.sort_values(['source_table_index','source_row_start']).drop(columns=['_s','_p'])

    out = pd.DataFrame(index=range(len(df)))
    out['Status']      = df['status'].map(lambda x: _STATUS_MAP.get(str(x), '🔵 NO DATA'))
    out['Component']   = df['description'].values
    out['Engine']      = df['engine_label'].values
    out['Unit']        = df['unit'].values
    out['Periodicity'] = [f"{int(float(x))}" if _sf(x) > 0 else "—" for x in df['periodicity'].values]
    out['Last O/H']    = [str(x) if x and str(x) not in ('nan','None','') else '—' for x in df['last_oh_date'].values]
    out['Hrs Since']   = [f"{int(float(x))}" if _sf(x) > 0 else "—" for x in df['hrs_since'].values]
    out['Used %']      = [round(float(x)*100, 1) if _sf(x) > 0 else 0.0 for x in df['pct_used'].values]
    if include_trace:
        out['Table'] = [str(x) if str(x) != 'nan' else '—' for x in df['source_table_index'].values]
        out['Rows']  = [f"{a}-{b}" if str(a) != 'nan' and str(b) != 'nan' else '—'
                        for a, b in zip(df['source_row_start'].values, df['source_row_end'].values)]
    return out

def build_issues_df(issues: List[Dict[str, Any]]) -> pd.DataFrame:
    if not issues:
        return pd.DataFrame(columns=["Severity","Code","Message","Table","Row","Row Key"])
    df = pd.DataFrame(issues)
    for c in ["severity","issue_code","message","table_index","row_index","row_key"]:
        if c not in df.columns:
            df[c] = ""
    df["_sev"] = df["severity"].map(lambda x: _ISSUE_ORD.get(str(x), 9))
    df = df.sort_values(["_sev", "table_index", "row_index"])
    out = pd.DataFrame()
    out["Severity"] = df["severity"]
    out["Code"] = df["issue_code"]
    out["Message"] = df["message"]
    out["Table"] = df["table_index"]
    out["Row"] = df["row_index"]
    out["Row Key"] = df["row_key"]
    return out

_CC = {
    'Status':      st.column_config.TextColumn('Status', width=145),
    'Component':   st.column_config.TextColumn('Component', width=260),
    'Engine':      st.column_config.TextColumn('Engine', width=82),
    'Unit':        st.column_config.TextColumn('Unit', width=72),
    'Periodicity': st.column_config.TextColumn('Periodicity', width=100),
    'Last O/H':    st.column_config.TextColumn('Last O/H', width=115),
    'Hrs Since':   st.column_config.TextColumn('Hrs Since', width=100),
    'Used %':      st.column_config.ProgressColumn('Used %', min_value=0, max_value=150, format='%.1f%%', width=132),
    'Table':       st.column_config.TextColumn('Table', width=60),
    'Rows':        st.column_config.TextColumn('Rows', width=75),
}

def show_matrix(records: list, include_trace: bool = False, mode: str = 'matrix', height: int = None):
    tbl = _build(records, include_trace=include_trace, mode=mode)
    if tbl.empty:
        st.info("No data.")
        return
    cfg = {k: v for k, v in _CC.items() if k in tbl.columns}
    h = height or min(860, 38 * len(tbl) + 44)
    st.dataframe(tbl, use_container_width=True, hide_index=True, height=h, column_config=cfg)

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
<div class="hero-h">TEC‑004 Hybrid Operational Build</div>
<div class="hero-s">
Original parser preserved for matrix recall, wrapped with multi-report staging, preview-before-save workflow,
SQLite report history, and optional trace columns for table/row visibility.
</div>
<div class="hero-rule"></div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
#  UPLOAD PANEL
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="panel upload-panel">', unsafe_allow_html=True)
uc, ic = st.columns([2.1, 1.0], gap='large')
with uc:
    uploaded_files = st.file_uploader(
        'Upload TEC-004 .doc reports',
        type=['doc'],
        accept_multiple_files=True,
        label_visibility='collapsed'
    )
    st.markdown('<div class="muted">Upload one or many legacy TEC‑004 <b>.doc</b> reports. Files are converted to DOCX, parsed with the original working extraction core, reviewed as matrices, and only then saved.</div>', unsafe_allow_html=True)
with ic:
    st.markdown("""
    <div class="muted">
    <b>Parser mode</b><br>Original extraction core preserved<br><br>
    <b>Workflow</b><br>Upload → preview → review → save<br><br>
    <b>Persistence</b><br>SQLite report ledger
    </div>
    """, unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_files:
    existing_hashes = {r.get("file_hash") for r in st.session_state.parsed_reports}
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
                    parsed["file_hash"] = fh
                    parsed["filename"] = uploaded.name
                    st.session_state.parsed_reports.append(parsed)
                    if st.session_state.active_report_hash is None:
                        st.session_state.active_report_hash = fh
                except Exception as e:
                    st.session_state.parsed_reports.append({
                        "vessel_name": "UNKNOWN",
                        "report_date": "",
                        "me_total_hrs": 0,
                        "me_this_month": 0,
                        "components": [],
                        "me_comps": [],
                        "aux_comps": [],
                        "other_equipment": [],
                        "warnings": [],
                        "issues": [{"severity":"error","issue_code":"PARSE_FAILURE","message":f"{uploaded.name}: {e}","table_index":None,"row_index":None,"row_key":""}],
                        "debug": {},
                        "uploaded_at": now_iso(),
                        "filename": uploaded.name,
                        "file_hash": fh,
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
    errs = sum(1 for x in r.get("issues", []) if x.get("severity") == "error")
    warns = sum(1 for x in r.get("issues", []) if x.get("severity") == "warning")
    queue_rows.append({
        "Active": "●" if r.get("file_hash") == st.session_state.active_report_hash else "",
        "Filename": r.get("filename", ""),
        "Vessel": r.get("vessel_name", "UNKNOWN"),
        "Report Date": r.get("report_date") or "—",
        "Rows": len(r.get("components", [])),
        "Warnings": warns,
        "Errors": errs,
        "Hash": (r.get("file_hash", "")[:18] + "…") if r.get("file_hash") else ""
    })
queue_df = pd.DataFrame(queue_rows)
st.dataframe(
    queue_df,
    use_container_width=True,
    hide_index=True,
    height=min(340, 38 * len(queue_df) + 44),
    column_config={
        "Active": st.column_config.TextColumn("Active", width=55),
        "Filename": st.column_config.TextColumn("Filename", width=250),
        "Vessel": st.column_config.TextColumn("Vessel", width=150),
        "Report Date": st.column_config.TextColumn("Report Date", width=120),
        "Rows": st.column_config.NumberColumn("Rows", width=70),
        "Warnings": st.column_config.NumberColumn("Warnings", width=80),
        "Errors": st.column_config.NumberColumn("Errors", width=70),
        "Hash": st.column_config.TextColumn("Hash", width=180),
    }
)

selector = {f"{r.get('filename','')} | {r.get('vessel_name','UNKNOWN')} | {r.get('report_date') or '—'}": r.get("file_hash") for r in reports}
selected_label = st.selectbox(
    "Active report",
    list(selector.keys()),
    index=list(selector.values()).index(st.session_state.active_report_hash) if st.session_state.active_report_hash in selector.values() else 0
)
st.session_state.active_report_hash = selector[selected_label]
active = next(r for r in reports if r.get("file_hash") == st.session_state.active_report_hash)

me = active.get("me_comps", [])
aux = active.get("aux_comps", [])
oe = active.get("other_equipment", [])
all_ = active.get("components", [])
issues = active.get("issues", [])
debug = active.get("debug", {})

n_od = sum(1 for c in all_ if c.get('status') == 'OVERDUE')
n_hp = sum(1 for c in all_ if c.get('status') == 'HIGH PRIORITY')
n_ok = sum(1 for c in all_ if c.get('status') == 'OK')
n_err = sum(1 for x in issues if x.get("severity") == "error")
n_warn = sum(1 for x in issues if x.get("severity") == "warning")

st.markdown(f"""
<div class="metric-grid">
  <div class="metric b"><div class="metric-v">{active.get('vessel_name','UNKNOWN')}</div><div class="metric-l">Vessel</div></div>
  <div class="metric"><div class="metric-v">{active.get('report_date') or '—'}</div><div class="metric-l">Report Date</div></div>
  <div class="metric"><div class="metric-v">{active.get('me_total_hrs') or 0:,}</div><div class="metric-l">ME Total Hrs</div></div>
  <div class="metric"><div class="metric-v">{active.get('me_this_month') or 0:,}</div><div class="metric-l">ME This Month</div></div>
  <div class="metric o"><div class="metric-v">{len(all_)}</div><div class="metric-l">Parsed Rows</div></div>
  <div class="metric r"><div class="metric-v">{n_od}</div><div class="metric-l">Overdue</div></div>
  <div class="metric g"><div class="metric-v">{n_ok}</div><div class="metric-l">OK</div></div>
</div>
""", unsafe_allow_html=True)

if n_err:
    st.markdown(f'<div class="banner err"><b>{n_err}</b> parser errors detected. Review matrices and debug tabs before save.</div>', unsafe_allow_html=True)
elif n_warn:
    st.markdown(f'<div class="banner warn"><b>{n_warn}</b> warnings detected. Matrix rows were still emitted wherever the original parser found paired date/hour data.</div>', unsafe_allow_html=True)
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
        st.session_state.save_feedback[active["file_hash"]] = (ok, msg)
with a3:
    export_df = _build(all_, include_trace=True, mode='source')
    st.download_button(
        "Download active CSV",
        data=export_df.to_csv(index=False).encode("utf-8"),
        file_name=f"{Path(active.get('filename','report')).stem}_preview.csv",
        mime="text/csv"
    )
with a4:
    st.markdown(f'<div class="micro">Parser {PARSER_VERSION} • File {active.get("filename","")} • Hash {active.get("file_hash","")[:18]}…</div>', unsafe_allow_html=True)

fb = st.session_state.save_feedback.get(active.get("file_hash"))
if fb:
    if fb[0]:
        st.success(fb[1])
    else:
        st.warning(fb[1])

# ══════════════════════════════════════════════════════════════════
#  TABS
# ══════════════════════════════════════════════════════════════════
tab_me, tab_aux, tab_oe, tab_issues, tab_debug, tab_history = st.tabs([
    "Main Engine", "Auxiliary Engines", "Other Equipment", "Parse Issues", "Parser Debug", "Saved Reports"
])

with tab_me:
    st.markdown(f'<div class="bar"><div class="bar-i">⚙</div><div><div class="bar-t">Main Engine Matrix</div><div class="bar-s">original paired-row extraction preserved</div></div><div class="bar-r"></div><div class="bar-b">rows: {len(me)}</div></div>', unsafe_allow_html=True)
    mf1, mf2, mf3, mf4 = st.columns([2.0, 1.8, 1.4, 2.2])
    with mf1:
        me_comp = st.selectbox("Component", ["All"] + sorted({c['description'] for c in me}), key="me_comp")
    with mf2:
        me_status = st.selectbox("Status", ["All","Overdue only","High Priority +","OK only"], key="me_status")
    with mf3:
        me_trace = st.checkbox("Show trace", value=True, key="me_trace")
    with mf4:
        me_sort = st.radio("Sort", ["Component → Cylinder", "Priority → % Used", "Source order"], horizontal=True, key="me_sort")

    me_view = me[:]
    if me_comp != 'All':
        me_view = [c for c in me_view if c['description'] == me_comp]
    if me_status == 'Overdue only':
        me_view = [c for c in me_view if c['status'] == 'OVERDUE']
    elif me_status == 'High Priority +':
        me_view = [c for c in me_view if c['status'] in ('OVERDUE','HIGH PRIORITY')]
    elif me_status == 'OK only':
        me_view = [c for c in me_view if c['status'] == 'OK']

    mode = 'matrix' if 'Component' in me_sort else ('priority' if 'Priority' in me_sort else 'source')
    show_matrix(me_view, include_trace=me_trace, mode=mode)

with tab_aux:
    st.markdown(f'<div class="bar"><div class="bar-i">🔩</div><div><div class="bar-t">Auxiliary Engines Matrix</div><div class="bar-s">original AUX extraction preserved</div></div><div class="bar-r"></div><div class="bar-b">rows: {len(aux)}</div></div>', unsafe_allow_html=True)
    af1, af2, af3, af4, af5 = st.columns([1.4, 2.0, 1.8, 1.4, 2.2])
    with af1:
        ax_eng = st.selectbox("Engine", ["All"] + sorted({c['engine_label'] for c in aux}), key="ax_eng")
    with af2:
        ax_comp = st.selectbox("Component", ["All"] + sorted({c['description'] for c in aux}), key="ax_comp")
    with af3:
        ax_status = st.selectbox("Status", ["All","Overdue only","High Priority +","OK only"], key="ax_status")
    with af4:
        ax_trace = st.checkbox("Show trace", value=True, key="ax_trace")
    with af5:
        ax_sort = st.radio("Sort", ["Component → Cylinder", "Priority → % Used", "Source order"], horizontal=True, key="ax_sort")

    ax_view = aux[:]
    if ax_eng != 'All':
        ax_view = [c for c in ax_view if c['engine_label'] == ax_eng]
    if ax_comp != 'All':
        ax_view = [c for c in ax_view if c['description'] == ax_comp]
    if ax_status == 'Overdue only':
        ax_view = [c for c in ax_view if c['status'] == 'OVERDUE']
    elif ax_status == 'High Priority +':
        ax_view = [c for c in ax_view if c['status'] in ('OVERDUE','HIGH PRIORITY')]
    elif ax_status == 'OK only':
        ax_view = [c for c in ax_view if c['status'] == 'OK']

    mode = 'matrix' if 'Component' in ax_sort else ('priority' if 'Priority' in ax_sort else 'source')
    show_matrix(ax_view, include_trace=ax_trace, mode=mode)

with tab_oe:
    st.markdown(f'<div class="bar"><div class="bar-i">🛠</div><div><div class="bar-t">Other Equipment</div><div class="bar-s">turbocharger · coolers · A/C · D/G equipment</div></div><div class="bar-r"></div><div class="bar-b">rows: {len(oe)}</div></div>', unsafe_allow_html=True)
    if not oe:
        st.info("No other equipment data found in this report.")
    else:
        oe_df = pd.DataFrame(oe)
        st.dataframe(
            oe_df[['section','description','last_date','run_hrs','source_table_index','source_row']],
            use_container_width=True,
            hide_index=True,
            column_config={
                'section': st.column_config.TextColumn('Section', width=220),
                'description': st.column_config.TextColumn('Description', width=300),
                'last_date': st.column_config.TextColumn('Last Date / O/H', width=160),
                'run_hrs': st.column_config.TextColumn('Run Hrs', width=120),
                'source_table_index': st.column_config.TextColumn('Table', width=70),
                'source_row': st.column_config.TextColumn('Row', width=70),
            },
            height=min(780, 38 * len(oe_df) + 44)
        )

with tab_issues:
    st.markdown(f'<div class="bar"><div class="bar-i">⚠</div><div><div class="bar-t">Parse Issues</div><div class="bar-s">warnings and extraction failures</div></div><div class="bar-r"></div><div class="bar-b">issues: {len(issues)}</div></div>', unsafe_allow_html=True)
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
    st.markdown('<div class="bar"><div class="bar-i">🧪</div><div><div class="bar-t">Parser Debug</div><div class="bar-s">raw table previews and section counts</div></div><div class="bar-r"></div><div class="bar-b">debug</div></div>', unsafe_allow_html=True)
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
        st.markdown("### Section counts")
        st.dataframe(tdbg, use_container_width=True, hide_index=True)

    previews = debug.get("raw_tables_preview", [])
    if previews:
        st.markdown("### Raw table preview")
        for item in previews:
            with st.expander(f"Table {item['table_index']} · {item['rows']} rows · {item['cols']} cols"):
                st.dataframe(pd.DataFrame(item["sample"]), use_container_width=True, hide_index=True)

with tab_history:
    st.markdown('<div class="bar"><div class="bar-i">🗄</div><div><div class="bar-t">Saved Reports</div><div class="bar-s">SQLite-backed report history</div></div><div class="bar-r"></div><div class="bar-b">history</div></div>', unsafe_allow_html=True)
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
                "parser_version": st.column_config.TextColumn("Parser", width=200),
                "me_total_hrs": st.column_config.NumberColumn("ME Total", width=100, format="%d"),
                "me_this_month": st.column_config.NumberColumn("ME Month", width=100, format="%d"),
                "created_at": st.column_config.TextColumn("Saved At", width=180),
                "parsed_rows": st.column_config.NumberColumn("Rows", width=70),
            }
        )
