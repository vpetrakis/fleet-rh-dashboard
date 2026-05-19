"""
╔══════════════════════════════════════════════════════════════════════╗
║   FLEET RUNNING HOURS MONITORING SYSTEM  v6.1                        ║
║   hardened parser · dual aux extraction · safe commit               ║
╚══════════════════════════════════════════════════════════════════════╝
"""

import streamlit as st

st.set_page_config(
    page_title="Fleet Running Hours",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="expanded",
)

import os
import re
import sqlite3
import tempfile
import hashlib
import subprocess
import shutil
from datetime import datetime
from pathlib import Path
import pandas as pd
from docx import Document

# ═══════════════════════════════════════════════════════════════════
#  GLOBAL UI
# ═══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
  --bg:#03060d; --bg1:#06091a; --bg2:#080e20; --bg3:#0b1228; --bg4:#0f1830;
  --b1:#0f1c35; --b2:#182840; --b3:#223350;
  --gold:#c89a14; --gold2:#e0b422; --gold3:#f5cc44;
  --red:#cc2828; --red2:#ff5c5c; --red3:#ff8a8a;
  --orange:#b85518; --ora2:#ff8833; --ora3:#ffb366;
  --green:#0d8a4a; --grn2:#22c55e; --grn3:#6ee7b7;
  --blue:#1444a8; --blu2:#3b82f6; --blu3:#93c5fd;
  --t0:#f2f7ff; --t1:#c0d0e8; --t2:#6a84a8; --t3:#304060;
  --ff:'Space Grotesk', sans-serif;
  --fi:'Inter', sans-serif;
  --fm:'JetBrains Mono', monospace;
}
*, *::before, *::after { box-sizing: border-box; }
html, body, [class*="css"] {
  font-family: var(--fi) !important;
  background: var(--bg) !important;
  color: var(--t1) !important;
  -webkit-font-smoothing: antialiased;
}
.main, .main > div { background: var(--bg) !important; }
.block-container { padding: 2rem 2.5rem 5rem !important; max-width: 100% !important; }

.main::before {
  content: ''; position: fixed; inset: 0; pointer-events: none; z-index: 0;
  background:
    radial-gradient(ellipse 90% 50% at -10% -5%, rgba(200,154,20,0.06) 0%, transparent 55%),
    radial-gradient(ellipse 70% 45% at 110% 105%, rgba(20,68,168,0.05) 0%, transparent 55%);
}

[data-testid="stSidebar"] { background: var(--bg1) !important; border-right: 1px solid var(--b2) !important; }
[data-testid="stSidebar"] * { color: var(--t1) !important; }
[data-testid="stSidebarContent"] { padding: 1.5rem 1.25rem !important; }
[data-testid="stSidebar"] .stSelectbox > div > div {
  background: var(--bg3) !important; border: 1px solid var(--b2) !important; border-radius: 6px !important;
}

h1 { font-family: var(--ff) !important; font-size: 1.8rem !important; font-weight: 700 !important; color: var(--t0) !important; letter-spacing: -0.02em !important; line-height: 1.2 !important; }
h2 { font-family: var(--ff) !important; font-size: 1.2rem !important; font-weight: 600 !important; color: var(--t0) !important; }
h3 { font-family: var(--ff) !important; font-size: 1rem !important; font-weight: 500 !important; color: var(--t1) !important; }

[data-testid="stMetric"] {
  background: var(--bg3) !important; border: 1px solid var(--b2) !important;
  border-radius: 10px !important; padding: 1rem 1.2rem 1.1rem !important;
  position: relative !important; overflow: hidden !important;
  transition: border-color .25s, transform .2s !important;
}
[data-testid="stMetric"]:hover { border-color: var(--b3) !important; transform: translateY(-2px) !important; }
[data-testid="stMetricValue"] {
  font-family: var(--ff) !important; font-size: 2rem !important; font-weight: 700 !important;
  color: var(--t0) !important; letter-spacing: -0.03em !important;
}
[data-testid="stMetricLabel"] {
  font-family: var(--fi) !important; color: var(--t3) !important; font-size: 0.62rem !important;
  text-transform: uppercase !important; letter-spacing: 0.15em !important;
}

[data-testid="stDataFrame"] {
  border: 1px solid var(--b2) !important; border-radius: 10px !important; overflow: hidden !important;
  box-shadow: 0 4px 24px rgba(0,0,0,0.35) !important;
}
.dvn-scroller { background: var(--bg2) !important; }

.stButton > button {
  background: linear-gradient(135deg, var(--gold) 0%, #8a6a08 100%) !important;
  color: #000 !important; border: none !important; font-family: var(--ff) !important; font-weight: 600 !important;
  font-size: 0.82rem !important; letter-spacing: 0.06em !important; text-transform: uppercase !important;
  border-radius: 7px !important; padding: .6rem 1.8rem !important;
  box-shadow: 0 2px 14px rgba(200,154,20,.2),inset 0 1px 0 rgba(255,255,255,.1) !important;
  transition: all .18s !important;
}
.stButton > button:hover {
  background: linear-gradient(135deg, var(--gold2) 0%, var(--gold) 100%) !important;
  box-shadow: 0 5px 22px rgba(200,154,20,.38) !important;
  transform: translateY(-2px) !important;
}
.stButton > button:active { transform: translateY(0) !important; }

[data-testid="stFileUploadDropzone"] {
  background: linear-gradient(160deg,rgba(200,154,20,.04) 0%,rgba(20,68,168,.03) 100%) !important;
  border: 1.5px dashed var(--gold) !important; border-radius: 14px !important;
  transition: all .3s !important; padding: 3rem 2rem !important;
}
[data-testid="stFileUploadDropzone"]:hover {
  background: rgba(200,154,20,.07) !important; border-color: var(--gold2) !important;
  box-shadow: 0 0 40px rgba(200,154,20,.07) !important;
}
[data-testid="stFileUploadDropzone"] p,
[data-testid="stFileUploadDropzone"] span {
  color: var(--gold2) !important; font-family: var(--ff) !important; font-size: .95rem !important; font-weight: 500 !important;
}
[data-testid="stFileUploadDropzone"] small { color: var(--t2) !important; }

.stTabs [data-baseweb="tab-list"] {
  background: var(--bg2) !important; border-radius: 10px 10px 0 0 !important;
  border-bottom: 1px solid var(--b2) !important; gap: 0 !important; padding: 0 1rem !important;
}
.stTabs [data-baseweb="tab"] {
  background: transparent !important; color: var(--t3) !important; font-family: var(--ff) !important;
  font-weight: 500 !important; letter-spacing: .04em !important; font-size: .75rem !important;
  text-transform: uppercase !important; padding: .85rem 1.3rem !important;
  border-bottom: 2px solid transparent !important; margin-bottom: -1px !important; transition: color .2s !important;
}
.stTabs [data-baseweb="tab"]:hover { color: var(--t2) !important; }
.stTabs [aria-selected="true"] { color: var(--gold2) !important; border-bottom: 2px solid var(--gold) !important; }
.stTabs [data-baseweb="tab-panel"] {
  background: var(--bg2) !important; border: 1px solid var(--b2) !important; border-top: none !important;
  border-radius: 0 0 10px 10px !important; padding: 1.5rem !important;
}

.stSelectbox > div > div, .stMultiSelect > div > div {
  background: var(--bg3) !important; border: 1px solid var(--b2) !important; border-radius: 7px !important;
  color: var(--t1) !important; font-family: var(--fi) !important;
}
.stSelectbox label, .stMultiSelect label {
  font-family: var(--fi) !important; color: var(--t3) !important; font-size: .7rem !important;
  text-transform: uppercase !important; letter-spacing: .1em !important;
}

.stAlert { border-radius: 8px !important; border-left-width: 3px !important; }
hr { border-color: var(--b2) !important; opacity:1 !important; margin:1.5rem 0 !important; }
a { color: var(--gold2) !important; text-decoration: none !important; }
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: var(--bg1); }
::-webkit-scrollbar-thumb { background: var(--b3); border-radius: 3px; }

@keyframes slideDown  { from{opacity:0;transform:translateY(-18px)} to{opacity:1;transform:translateY(0)} }
@keyframes slideUp    { from{opacity:0;transform:translateY(14px)}  to{opacity:1;transform:translateY(0)} }
@keyframes slideRight { from{opacity:0;transform:translateX(-14px)} to{opacity:1;transform:translateX(0)} }
@keyframes numIn      { from{opacity:0;transform:translateY(8px)}   to{opacity:1;transform:translateY(0)} }
@keyframes popIn      { from{opacity:0;transform:scale(.9)}          to{opacity:1;transform:scale(1)} }
@keyframes successPop { 0%{transform:scale(.85);opacity:0} 55%{transform:scale(1.02)} 100%{transform:scale(1);opacity:1} }
@keyframes goldLine   { from{width:0;opacity:0} to{width:100%;opacity:1} }

.ph { animation: slideDown .45s cubic-bezier(.22,1,.36,1) both; }
.ph h1 { margin: 0; }
.ph-eye {
  font-family: var(--fi); font-size: .6rem; font-weight: 500; letter-spacing: .22em;
  text-transform: uppercase; color: var(--gold); margin-bottom: .3rem;
  animation: slideDown .4s .05s ease both;
}
.ph-line {
  height: 1px; margin: .35rem 0 1.75rem;
  background: linear-gradient(90deg, var(--gold) 0%, var(--b2) 30%, transparent 100%);
  animation: goldLine .7s .1s ease both;
}

.kc {
  background: var(--bg3); border: 1px solid var(--b2); border-radius: 10px;
  padding: 1rem 1.2rem 1.1rem; position: relative; overflow: hidden;
  animation: slideUp .4s ease both; transition: border-color .25s, transform .2s, box-shadow .25s; cursor: default;
}
.kc:hover { transform: translateY(-4px); box-shadow: 0 14px 40px rgba(0,0,0,.5); }
.kc::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; border-radius: 10px 10px 0 0; }
.kc.gold   { border-color: rgba(200,154,20,.3); } .kc.gold::before   { background: linear-gradient(90deg, var(--gold),  transparent 70%); }
.kc.red    { border-color: rgba(204,40,40,.25); } .kc.red::before    { background: linear-gradient(90deg, var(--red),   transparent 70%); }
.kc.orange { border-color: rgba(184,85,24,.25); } .kc.orange::before { background: linear-gradient(90deg, var(--orange),transparent 70%); }
.kc.green  { border-color: rgba(13,138,74,.25); } .kc.green::before  { background: linear-gradient(90deg, var(--green), transparent 70%); }
.kc.blue   { border-color: rgba(20,68,168,.25); } .kc.blue::before   { background: linear-gradient(90deg, var(--blue),  transparent 70%); }

.kc-val {
  font-family: var(--ff); font-size: 2.2rem; font-weight: 700; line-height: 1.1;
  letter-spacing: -.04em; position: relative; z-index: 1; animation: numIn .4s .1s ease both;
}
.kc.gold .kc-val { color: var(--gold3); }
.kc.red .kc-val { color: var(--red2); }
.kc.orange .kc-val { color: var(--ora2); }
.kc.green .kc-val { color: var(--grn2); }
.kc.blue .kc-val { color: var(--blu2); }

.kc-lbl {
  font-family: var(--fi); font-size: .6rem; font-weight: 500; text-transform: uppercase; letter-spacing: .16em;
  color: var(--t3); margin-top: 5px; position: relative; z-index: 1;
}

.sl {
  font-family: var(--fi); font-size: .58rem; font-weight: 600; letter-spacing: .22em; text-transform: uppercase;
  color: var(--t3); display: flex; align-items: center; gap: .75rem; margin: 1.75rem 0 1rem;
}
.sl::after { content: ''; flex: 1; height: 1px; background: linear-gradient(90deg, var(--b2), transparent); }

.ps-row { display: flex; gap: .65rem; flex-wrap: wrap; margin: 1rem 0 1.5rem; }
.ps {
  background: var(--bg3); border: 1px solid var(--b2); border-radius: 9px; padding: .65rem 1.1rem .7rem; min-width: 86px;
  animation: popIn .35s ease both; transition: border-color .2s, transform .15s; position: relative; overflow: hidden;
}
.ps:hover { transform: translateY(-2px); border-color: var(--b3); }
.ps::before { content:''; position:absolute; top:0; left:0; right:0; height:2px; }
.ps.red::before    { background: var(--red); }
.ps.orange::before { background: var(--orange); }
.ps.green::before  { background: var(--green); }
.ps.blue::before   { background: var(--blue); }
.ps-val { font-family: var(--ff); font-size: 1.7rem; font-weight: 700; line-height: 1; letter-spacing: -.03em; }
.ps-lbl { font-family: var(--fi); font-size: .58rem; text-transform: uppercase; letter-spacing: .14em; color: var(--t3); margin-top: 4px; }
.ps.red .ps-val { color: var(--red2); }
.ps.orange .ps-val { color: var(--ora2); }
.ps.green .ps-val { color: var(--grn2); }
.ps.blue .ps-val { color: var(--blu2); }

.ic {
  background: var(--bg3); border: 1px solid var(--b2); border-radius: 12px; padding: 1.4rem 1.6rem;
  font-family: var(--fi); font-size: .82rem; color: var(--t2); line-height: 1.9; animation: slideRight .4s ease both;
}
.ic-title {
  font-family: var(--ff); font-size: .58rem; font-weight: 600; letter-spacing: .2em; text-transform: uppercase;
  color: var(--gold2); margin-bottom: .4rem;
}

.sb {
  background: linear-gradient(135deg,rgba(13,138,74,.12),rgba(13,138,74,.04));
  border: 1px solid rgba(13,138,74,.3); border-radius: 10px; padding: 1rem 1.5rem; color: var(--grn3);
  font-family: var(--ff); font-size: .92rem; font-weight: 500; animation: successPop .5s cubic-bezier(.34,1.56,.64,1) both;
  display: flex; align-items: center; gap: .75rem;
}
.ac {
  background: rgba(13,138,74,.04); border: 1px solid rgba(13,138,74,.12); border-radius: 10px;
  padding: 1.75rem; text-align: center; color: var(--grn3); font-family: var(--ff); font-size: .95rem; font-weight: 500;
}

.logo {
  font-family: var(--ff); font-size: 1.15rem; font-weight: 700; letter-spacing: .04em; color: var(--gold2);
  display:flex; align-items:center; gap:.5rem;
}
.logo-tag { font-family: var(--fi); font-size: .57rem; text-transform: uppercase; letter-spacing: .2em; color: var(--t3); margin-top: 3px; }
.logo-rule { height: 1px; margin: 1.2rem 0; background: linear-gradient(90deg, var(--gold), transparent); }

.vc {
  display: flex; align-items: center; justify-content: space-between; background: var(--bg3); border: 1px solid var(--b2);
  border-radius: 7px; padding: .5rem .8rem; margin-bottom: .28rem; transition: border-color .2s, background .2s; animation: slideRight .28s ease both;
}
.vc:hover { border-color: var(--b3); background: var(--bg4); }
.vc.crit { border-left: 2px solid var(--red); }
.vc.warn { border-left: 2px solid var(--orange); }
.vc.safe { border-left: 2px solid var(--green); }
.vc-name {
  font-family: var(--ff); font-size: .73rem; font-weight: 600; color: var(--t1); white-space: nowrap;
  overflow: hidden; text-overflow: ellipsis; max-width: 115px;
}
.vc-tags { display: flex; gap: 4px; align-items: center; flex-shrink: 0; }
.vt { font-family: var(--fm); font-size: .56rem; font-weight: 500; padding: 1px 6px; border-radius: 3px; }
.vt.od { background: rgba(204,40,40,.15); color: var(--red2); }
.vt.hp { background: rgba(184,85,24,.15); color: var(--ora2); }
.vt.ok { background: rgba(13,138,74,.15); color: var(--grn2); }

.ml {
  display: flex; gap: 1.4rem; flex-wrap: wrap; font-family: var(--fm); font-size: .66rem; color: var(--t3); margin: .7rem 0 0;
}
.ml b { color: var(--t2); font-weight: 500; }

.filter-count {
  margin: .55rem 0 1rem; font-family: var(--fm); font-size: .68rem; color: var(--t2);
}
.filter-count b { color: var(--t1); }
.filter-count .od { color: var(--red2); }
.filter-count .hp { color: var(--ora2); }
.filter-count .ok { color: var(--grn2); }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════
#  DATABASE
# ═══════════════════════════════════════════════════════════════════
DB_PATH = Path("running_hours.db")

def get_db():
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    c = get_db()
    c.execute("PRAGMA journal_mode=WAL;")
    c.executescript("""
    CREATE TABLE IF NOT EXISTS vessels(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        created_at TEXT NOT NULL DEFAULT(datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS upload_log(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL,
        filename TEXT NOT NULL,
        file_hash TEXT NOT NULL,
        report_date TEXT,
        me_total_hrs INTEGER,
        me_this_month INTEGER,
        parse_confidence REAL DEFAULT 0,
        component_count INTEGER DEFAULT 0,
        warning_count INTEGER DEFAULT 0,
        uploaded_at TEXT NOT NULL DEFAULT(datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS components(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL,
        category TEXT NOT NULL,
        engine_label TEXT NOT NULL,
        unit TEXT NOT NULL,
        description TEXT NOT NULL,
        periodicity REAL,
        last_oh_date TEXT,
        last_oh_hrs REAL,
        hrs_since REAL,
        pct_used REAL,
        status TEXT NOT NULL,
        updated_at TEXT NOT NULL DEFAULT(datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS other_equipment(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL,
        section TEXT NOT NULL,
        description TEXT NOT NULL,
        periodicity TEXT,
        last_date TEXT,
        run_hrs TEXT,
        updated_at TEXT NOT NULL DEFAULT(datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_cv ON components(vessel_name);
    CREATE INDEX IF NOT EXISTS idx_cs ON components(status);
    CREATE INDEX IF NOT EXISTS idx_ul_v ON upload_log(vessel_name);
    """)
    c.commit()
    c.close()

init_db()


# ═══════════════════════════════════════════════════════════════════
#  CONVERSION
# ═══════════════════════════════════════════════════════════════════
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found. packages.txt must contain: libreoffice")

    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as t:
        t.write(raw)
        src = t.name

    outdir = tempfile.mkdtemp(prefix="tec004_")
    target = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".docx")
    profile = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"

    try:
        r = subprocess.run(
            [
                soffice,
                "--headless",
                "--norestore",
                "--nofirststartwizard",
                f"-env:UserInstallation={profile}",
                "--convert-to", "docx",
                src,
                "--outdir", outdir,
            ],
            capture_output=True,
            timeout=120,
        )
        if not os.path.exists(target):
            raise RuntimeError(f"exit {r.returncode}: {r.stderr.decode('utf-8','ignore')[:400]}")
        with open(target, "rb") as f:
            return f.read()
    finally:
        for p in [src, target]:
            try:
                if os.path.exists(p):
                    os.unlink(p)
            except Exception:
                pass
        shutil.rmtree(outdir, ignore_errors=True)


# ═══════════════════════════════════════════════════════════════════
#  PARSER HELPERS
# ═══════════════════════════════════════════════════════════════════
def clean_text(x):
    if x is None:
        return ""
    return re.sub(r"\s+", " ", str(x).replace("\xa0", " ")).strip()

def norm(x):
    x = clean_text(x).upper()
    x = x.replace("O/H", "OH")
    x = x.replace("PERIODICTLY", "PERIODICITY")
    x = x.replace("D/G", "DG")
    x = x.replace("AUX.", "AUX")
    x = x.replace("NO.", "NO")
    x = x.replace("’", "'")
    x = re.sub(r"[^A-Z0-9 /().:\-]", "", x)
    return x

def row_text(row):
    return " | ".join([clean_text(c) for c in row if clean_text(c)])

def table_grid(table):
    return [[clean_text(c.text) for c in row.cells] for row in table.rows]

def all_lines(doc):
    out = []
    for p in doc.paragraphs:
        t = clean_text(p.text)
        if t:
            out.append(t)
    for table in doc.tables:
        for row in table_grid(table):
            t = row_text(row)
            if t:
                out.append(t)
    return out

def parse_number(raw):
    if raw is None:
        return None
    s = clean_text(raw)
    if not s or s.upper() in {"N/A", "NA", "-", "--"}:
        return None
    s = s.replace(",", "")
    m = re.search(r"\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except Exception:
        return None

def parse_periodicity(raw):
    if raw is None:
        return None
    u = norm(raw)
    if not u or "OBSERVATION" in u or u in {"NA", "N/A", "-", "--"}:
        return None
    m = re.search(r"\d[\d.,]*", clean_text(raw))
    if not m:
        return None
    v = m.group(0).replace(",", "")
    if "." in v and len(v.split(".")[-1]) == 3:
        v = v.replace(".", "")
    try:
        return float(v)
    except Exception:
        return None

def parse_date(raw):
    if raw is None:
        return None
    raw = clean_text(raw)
    if not raw or raw.upper() in {"N/A", "NA", "-", "--"}:
        return None
    raw = raw.strip("[]")
    raw = raw.replace(".", " ")
    raw = re.sub(r"\bSEPT\b", "SEP", raw, flags=re.I)
    raw = re.sub(r"\bJUNE\b", "JUN", raw, flags=re.I)
    raw = re.sub(r"\bJULY\b", "JUL", raw, flags=re.I)
    raw = re.sub(r"\s+", " ", raw).strip()

    fmts = [
        "%d %b %y", "%d %B %y", "%d %b %Y", "%d %B %Y",
        "%d/%m/%y", "%d/%m/%Y", "%d-%m-%y", "%d-%m-%Y",
        "%b %Y", "%B %Y", "%Y-%m-%d"
    ]
    for fmt in fmts:
        for v in (raw, raw.upper(), raw.title()):
            try:
                return datetime.strptime(v, fmt).strftime("%Y-%m-%d")
            except Exception:
                pass
    return None

def parse_compact_date(tok):
    tok = clean_text(tok).upper()
    tok = re.sub(r"[^0-9A-Z]", "", tok)
    if not tok or tok in {"NA", "NANANA"}:
        return None

    if re.fullmatch(r"\d{6}", tok):
        dd = int(tok[:2]); mm = int(tok[2:4]); yy = int(tok[4:6])
        yy = 2000 + yy if yy < 70 else 1900 + yy
        try:
            return datetime(yy, mm, dd).strftime("%Y-%m-%d")
        except Exception:
            return None

    for fmt in ("%d%b%y", "%d%b%Y"):
        try:
            return datetime.strptime(tok.title(), fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    return None

def parse_hours(raw):
    if raw is None:
        return None
    s = clean_text(raw)
    if not s or s.upper() in {"N/A", "NA", "-", "--"}:
        return None
    nums = re.findall(r"\d[\d,]*", s)
    for n in nums:
        try:
            v = float(n.replace(",", ""))
            if v >= 0:
                return v
        except Exception:
            pass
    return None

def pct_used(h, p):
    if h is None or p is None or p <= 0:
        return 0.0
    return round(h / p, 4)

def status_from(h, p):
    if h is None or p is None or p <= 0:
        return "NO DATA"
    r = h / p
    if r >= 1.0:
        return "OVERDUE"
    if r >= 0.80:
        return "HIGH PRIORITY"
    return "OK"

def dedup_components(rows):
    out = []
    seen = set()
    for x in rows:
        key = (
            x.get("category"),
            x.get("engine_label"),
            x.get("unit"),
            x.get("description"),
            x.get("periodicity"),
            x.get("last_oh_date"),
            x.get("hrs_since"),
        )
        if key not in seen:
            seen.add(key)
            out.append(x)
    return out

def safe_int(v):
    try:
        if v is None:
            return None
        return int(float(v))
    except Exception:
        return None


# ═══════════════════════════════════════════════════════════════════
#  DOCUMENT DETECTION
# ═══════════════════════════════════════════════════════════════════
def detect_vessel_name(doc):
    texts = all_lines(doc)
    patterns = [
        r"VESSEL['’]S NAME[: ]+(?:MV )?([A-Z][A-Z0-9 \-]+)",
        r"VESSELS NAME[: ]+(?:MV )?([A-Z][A-Z0-9 \-]+)",
        r"TITLE VESSELS NAME MV ([A-Z][A-Z0-9 \-]+)",
        r"\bMV ([A-Z][A-Z0-9 \-]+)\b"
    ]
    for txt in texts:
        u = txt.upper()
        for pat in patterns:
            m = re.search(pat, u)
            if m:
                return clean_text(m.group(1))
    return "UNKNOWN"

def detect_report_date(doc):
    texts = all_lines(doc)
    for txt in texts:
        m = re.search(r"DATE[: ]*([A-Z0-9/.\- ]{5,30})", txt, re.I)
        if m:
            d = parse_date(m.group(1))
            if d:
                return d
    return None

def detect_me_totals(doc):
    text = "\n".join(all_lines(doc))
    mt = None
    mm = None
    m1 = re.search(r"TOTAL RUNNING HOURS[: ǀ|]*([\d,]+)", text, re.I)
    m2 = re.search(r"THIS MONTH[: ]*([\d,]+)", text, re.I)
    if m1:
        try:
            mt = int(m1.group(1).replace(",", ""))
        except Exception:
            pass
    if m2:
        try:
            mm = int(m2.group(1).replace(",", ""))
        except Exception:
            pass
    return mt, mm


# ═══════════════════════════════════════════════════════════════════
#  MAIN ENGINE PARSER
# ═══════════════════════════════════════════════════════════════════
def find_me_table(doc):
    for table in doc.tables:
        grid = table_grid(table)
        head = " ".join(row_text(r) for r in grid[:8]).upper()
        if "MAIN ENGINE" in head and "CYL" in head:
            return grid
    return None

def detect_me_cylinders(grid):
    cols = []
    seen = set()
    for ri, row in enumerate(grid[:5]):
        for ci, cell in enumerate(row):
            m = re.search(r"CYL\.?\s*NO\.?\s*(\d+)", norm(cell))
            if m:
                lbl = f"Cyl {int(m.group(1))}"
                if lbl not in seen:
                    seen.add(lbl)
                    cols.append((ci, lbl))
    return cols

def parse_main_engine(grid):
    warnings = []
    comps = []
    if not grid:
        return [], ["Main engine table not found."]

    cyl_cols = detect_me_cylinders(grid)
    if not cyl_cols:
        return [], ["Main engine cylinder columns not detected."]

    i = 2
    while i < len(grid) - 1:
        r1 = grid[i]
        r2 = grid[i + 1] if i + 1 < len(grid) else []
        nm = clean_text(r1[0]) if r1 else ""

        if not nm:
            i += 1
            continue

        nu = norm(nm)
        if any(x in nu for x in [
            "TITLE", "VESSELS NAME", "NOTE 1", "NOTE 2", "MAIN ENGINE",
            "DATE OF LAST", "RUNNING HOURS SINCE LAST", "TYPE", "TOTAL RUNNING HOURS"
        ]):
            i += 1
            continue

        marker1 = clean_text(r1[2]) if len(r1) > 2 else ""
        marker2 = clean_text(r2[2]) if len(r2) > 2 else ""
        same_name = nm == (clean_text(r2[0]) if r2 else nm)

        if marker1 == "1" and marker2 == "2":
            p = parse_periodicity(r1[1] if len(r1) > 1 else None)
            for ci, lbl in cyl_cols:
                d = parse_date(r1[ci] if ci < len(r1) else None)
                h = parse_hours(r2[ci] if ci < len(r2) else None)
                if d is None and h is None:
                    continue
                comps.append({
                    "category": "MAINENGINE",
                    "engine_label": "ME",
                    "unit": lbl,
                    "description": nm,
                    "periodicity": p,
                    "last_oh_date": d,
                    "last_oh_hrs": h,
                    "hrs_since": h,
                    "pct_used": pct_used(h, p),
                    "status": status_from(h, p),
                })
            i += 2
        elif same_name and marker1 == "1" and marker2 == "2":
            i += 2
        else:
            i += 1

    if not comps:
        warnings.append("No main-engine components extracted.")
    return dedup_components(comps), warnings


# ═══════════════════════════════════════════════════════════════════
#  OTHER EQUIPMENT PARSER
# ═══════════════════════════════════════════════════════════════════
def parse_other_equipment(doc):
    oe = []
    skip_terms = {
        "TURBOCHARGER", "AUXILIARY BOILER", "COOLERS", "EXH GAS BOILER",
        "A/C & REFR. COMPRESSORS", "AC REFR. COMPRESSORS", "MAIN AIR COMPRESSORS",
        "PERIODICTLY", "DATE OF LAST INSPECTION", "DATE OF LAST CLEANING",
        "DATE OF LAST O/H", "RUN HRS", "DATE", "PERIODICITY"
    }

    for table in doc.tables:
        grid = table_grid(table)
        for row in grid:
            cells = [clean_text(c) for c in row]
            if not any(cells):
                continue

            for sec, dc, pc, datec, hrsc in [
                ("TURBOCHARGER / AUX BOILER", 0, 1, 2, 3),
                ("COOLERS / EXH GAS BOILER", 5, 5, 6, 7),
                ("A/C & COMPRESSORS", 10, 10, 11, 12),
            ]:
                if dc >= len(cells):
                    continue
                desc = cells[dc]
                if not desc or norm(desc) in skip_terms:
                    continue

                periodicity = cells[pc] if pc < len(cells) else ""
                last_date = cells[datec] if datec < len(cells) else ""
                run_hrs = cells[hrsc] if hrsc < len(cells) else ""

                if clean_text(last_date) or clean_text(run_hrs):
                    oe.append({
                        "section": sec,
                        "description": desc,
                        "periodicity": periodicity,
                        "last_date": parse_date(last_date) or last_date,
                        "run_hrs": run_hrs
                    })

    dedup = []
    seen = set()
    for x in oe:
        k = (x["section"], x["description"], x["periodicity"], x["last_date"], x["run_hrs"])
        if k not in seen:
            seen.add(k)
            dedup.append(x)
    return dedup


# ═══════════════════════════════════════════════════════════════════
#  AUX PARSER — BLOCK 1 (AUX ENGINE No.1/2/3 row-pair section)
# ═══════════════════════════════════════════════════════════════════
def looks_like_aux_rowpair_table(grid):
    txt = " ".join(row_text(r) for r in grid[:22]).upper()
    return "AUX ENGINE MAKER / TYPE" in txt or "AUX. ENGINE MAKER / TYPE" in txt

def find_aux_rowpair_tables(doc):
    out = []
    for table in doc.tables:
        grid = table_grid(table)
        if looks_like_aux_rowpair_table(grid):
            out.append(grid)
    return out

def parse_aux_rowpair_tables(doc):
    warnings = []
    comps = []

    tables = find_aux_rowpair_tables(doc)
    if not tables:
        return [], ["Aux engine row-pair block not found."]

    for grid in tables:
        # This section in the converted sample is often collapsed and only reliably yields one aux engine block.
        # We parse it as a single aux unit when only one 1/2 column pair is present.
        desc_header_idx = None
        for i, row in enumerate(grid):
            if len(row) >= 4 and norm(row[0]) == "DESCRIPTION" and "PERIODICITY" in norm(row[1]):
                desc_header_idx = i
                break

        if desc_header_idx is None:
            continue

        i = desc_header_idx + 1
        while i < len(grid) - 1:
            r1 = grid[i]
            r2 = grid[i + 1] if i + 1 < len(grid) else []
            nm = clean_text(r1[0]) if r1 else ""
            if not nm:
                i += 1
                continue

            nu = norm(nm)
            if any(x in nu for x in [
                "DESCRIPTION", "PERIODICITY", "TITLE", "TABLE", "VESSELS NAME", "DATE"
            ]):
                i += 1
                continue

            marker1 = clean_text(r1[2]) if len(r1) > 2 else ""
            marker2 = clean_text(r2[2]) if len(r2) > 2 else ""

            if marker1 == "1" and marker2 == "2":
                p = parse_periodicity(r1[1] if len(r1) > 1 else None)
                d = parse_date(r1[3] if len(r1) > 3 else None)
                h = parse_hours(r2[3] if len(r2) > 3 else None)

                if d is not None or h is not None:
                    comps.append({
                        "category": "AUXENGINE",
                        "engine_label": "AUX-1",
                        "unit": "AUX-1",
                        "description": nm,
                        "periodicity": p,
                        "last_oh_date": d,
                        "last_oh_hrs": h,
                        "hrs_since": h,
                        "pct_used": pct_used(h, p),
                        "status": status_from(h, p),
                    })
                i += 2
            else:
                i += 1

    if not comps:
        warnings.append("No aux-engine row-pair components extracted.")
    return dedup_components(comps), warnings


# ═══════════════════════════════════════════════════════════════════
#  AUX PARSER — BLOCK 2 (DG No1/No2/No3 matrix)
# ═══════════════════════════════════════════════════════════════════
def looks_like_dg_table(grid):
    txt = " ".join(row_text(r) for r in grid[:30]).upper()
    return ("DG NO1" in txt or "D/G NO1" in txt) and ("DG NO2" in txt or "D/G NO2" in txt)

def find_dg_tables(doc):
    out = []
    for table in doc.tables:
        grid = table_grid(table)
        if looks_like_dg_table(grid):
            out.append(grid)
    return out

def split_tokens(cell_text):
    s = clean_text(cell_text).upper()
    if not s:
        return []
    s = s.replace("NANANA", " NA NA NA ")
    s = s.replace("N/A", " NA ")
    s = s.replace("[", " ").replace("]", " ")
    s = re.sub(r"[^A-Z0-9/.-]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s.split(" ") if s else []

def decode_dg_payload(cell_text):
    tokens = split_tokens(cell_text)
    date_val = None
    hrs_val = None

    for tok in tokens:
        d = parse_date(tok) or parse_compact_date(tok)
        if d and date_val is None:
            date_val = d
            continue

        if tok not in {"NA"}:
            n = parse_number(tok)
            if n is not None:
                # Prefer realistic running hours vs tiny tokens
                if n >= 50:
                    if hrs_val is None or n > hrs_val:
                        hrs_val = n

    return date_val, hrs_val

def parse_dg_tables(doc):
    warnings = []
    comps = []
    tables = find_dg_tables(doc)

    if not tables:
        return [], ["DG matrix block not found."]

    for grid in tables:
        header_idx = None
        for i, row in enumerate(grid):
            rt = norm(" ".join(row))
            if "DESCRIPTION" in rt and "PERIODICITY" in rt and ("DG NO1" in rt or "D/G NO1" in rt):
                header_idx = i
                break

        if header_idx is None:
            # fallback: locate first row where DG tokens appear and description/periodicity likely occupy cols 0,1
            for i, row in enumerate(grid):
                joined = " ".join(row).upper()
                if ("DG NO1" in joined or "D/G NO1" in joined) and ("DG NO2" in joined or "D/G NO2" in joined):
                    header_idx = i
                    break

        if header_idx is None:
            continue

        i = header_idx + 1
        while i < len(grid) - 1:
            r1 = grid[i]
            r2 = grid[i + 1] if i + 1 < len(grid) else []
            desc = clean_text(r1[0]) if len(r1) > 0 else ""
            desc_u = norm(desc)

            if not desc:
                i += 1
                continue

            if any(x in desc_u for x in [
                "DESCRIPTION", "TITLE", "TABLE", "VESSELS NAME", "DATE",
                "CHIEF ENGINEER", "REMARKS"
            ]):
                i += 1
                continue

            p = parse_periodicity(r1[1] if len(r1) > 1 else None)

            row_has = False
            dg_map = {"DG-1": 2, "DG-2": 3, "DG-3": 4}
            for eng, col in dg_map.items():
                d = None
                h = None

                if col < len(r1):
                    d1, h1 = decode_dg_payload(r1[col])
                    d = d or d1
                    h = h or h1

                if col < len(r2):
                    d2, h2 = decode_dg_payload(r2[col])
                    d = d or d2
                    h = h or h2

                # fallback: some conversions spill payload into adjacent cell
                if (d is None and h is None) and col + 1 < len(r1):
                    d3, h3 = decode_dg_payload((r1[col] if col < len(r1) else "") + " " + r1[col + 1])
                    d = d or d3
                    h = h or h3
                if (d is None and h is None) and col + 1 < len(r2):
                    d4, h4 = decode_dg_payload((r2[col] if col < len(r2) else "") + " " + r2[col + 1])
                    d = d or d4
                    h = h or h4

                if d is None and h is None:
                    continue

                row_has = True
                comps.append({
                    "category": "AUXENGINE",
                    "engine_label": eng,
                    "unit": eng,
                    "description": desc,
                    "periodicity": p,
                    "last_oh_date": d,
                    "last_oh_hrs": h,
                    "hrs_since": h,
                    "pct_used": pct_used(h, p),
                    "status": status_from(h, p),
                })

            if row_has:
                i += 2
            else:
                i += 1

    if not comps:
        warnings.append("No aux-engine components extracted from DG matrix.")
    return dedup_components(comps), warnings


# ═══════════════════════════════════════════════════════════════════
#  FULL PARSER
# ═══════════════════════════════════════════════════════════════════
def parse_doc_bytes(docx: bytes) -> dict:
    warnings = []

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as t:
        t.write(docx)
        tp = t.name

    try:
        doc = Document(tp)
    except Exception as e:
        raise ValueError(f"Cannot open DOCX: {e}")
    finally:
        try:
            os.unlink(tp)
        except Exception:
            pass

    if not doc.tables:
        raise ValueError("No tables detected — is this TEC-004?")

    vessel_name = detect_vessel_name(doc)
    report_date = detect_report_date(doc)
    me_total_hrs, me_this_month = detect_me_totals(doc)

    if vessel_name == "UNKNOWN":
        warnings.append("Could not extract vessel name.")

    me_grid = find_me_table(doc)
    me_rows, me_warns = parse_main_engine(me_grid)
    warnings.extend(me_warns)

    aux_rowpair_rows, aux_rowpair_warns = parse_aux_rowpair_tables(doc)
    dg_rows, dg_warns = parse_dg_tables(doc)
    warnings.extend(aux_rowpair_warns)
    warnings.extend(dg_warns)

    other_equipment = parse_other_equipment(doc)

    components = dedup_components(me_rows + aux_rowpair_rows + dg_rows)

    confidence = 0.0
    if vessel_name != "UNKNOWN":
        confidence += 0.15
    if report_date:
        confidence += 0.10
    if me_total_hrs is not None:
        confidence += 0.05
    if any(c["category"] == "MAINENGINE" for c in components):
        confidence += 0.25
    if any(c["category"] == "AUXENGINE" for c in components):
        confidence += 0.25
    if len(components) >= 10:
        confidence += 0.10
    if len(components) >= 20:
        confidence += 0.10
    confidence = round(min(confidence, 1.0), 2)

    return {
        "vessel_name": vessel_name,
        "report_date": report_date,
        "me_total_hrs": me_total_hrs,
        "me_this_month": me_this_month,
        "components": components,
        "other_equipment": other_equipment,
        "warnings": warnings,
        "parse_confidence": confidence,
    }


# ═══════════════════════════════════════════════════════════════════
#  PERSISTENCE
# ═══════════════════════════════════════════════════════════════════
def save_parsed(parsed, filename, fhash):
    conn = get_db()
    c = conn.cursor()
    now = datetime.utcnow().isoformat() + "Z"
    v = parsed["vessel_name"]

    try:
        c.execute("INSERT OR IGNORE INTO vessels(name,created_at) VALUES (?,?)", (v, now))
        c.execute("""
            INSERT INTO upload_log(
                vessel_name, filename, file_hash, report_date,
                me_total_hrs, me_this_month, parse_confidence,
                component_count, warning_count, uploaded_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            v, filename, fhash, parsed["report_date"],
            parsed["me_total_hrs"], parsed["me_this_month"],
            parsed.get("parse_confidence", 0.0),
            len(parsed["components"]),
            len(parsed.get("warnings", [])),
            now
        ))

        c.execute("DELETE FROM components WHERE vessel_name=?", (v,))
        for x in parsed["components"]:
            c.execute("""
                INSERT INTO components(
                    vessel_name, category, engine_label, unit, description,
                    periodicity, last_oh_date, last_oh_hrs, hrs_since,
                    pct_used, status, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                v, x["category"], x["engine_label"], x["unit"], x["description"],
                x["periodicity"], x["last_oh_date"], x["last_oh_hrs"],
                x["hrs_since"], x["pct_used"], x["status"], now
            ))

        c.execute("DELETE FROM other_equipment WHERE vessel_name=?", (v,))
        for x in parsed["other_equipment"]:
            c.execute("""
                INSERT INTO other_equipment(
                    vessel_name, section, description, periodicity,
                    last_date, run_hrs, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                v, x["section"], x["description"], x.get("periodicity", ""),
                x.get("last_date", ""), x.get("run_hrs", ""), now
            ))

        conn.commit()
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()


# ═══════════════════════════════════════════════════════════════════
#  READERS
# ═══════════════════════════════════════════════════════════════════
@st.cache_data(ttl=10)
def get_vessels():
    c = get_db()
    r = c.execute("SELECT name FROM vessels ORDER BY name").fetchall()
    c.close()
    return [x["name"] for x in r]

@st.cache_data(ttl=10)
def get_comps(vessel):
    c = get_db()
    df = pd.read_sql_query("SELECT * FROM components WHERE vessel_name=?", c, params=(vessel,))
    c.close()
    return df

@st.cache_data(ttl=10)
def get_oe(vessel):
    c = get_db()
    df = pd.read_sql_query("SELECT * FROM other_equipment WHERE vessel_name=? ORDER BY section, description", c, params=(vessel,))
    c.close()
    return df

@st.cache_data(ttl=10)
def get_history(vessel):
    c = get_db()
    df = pd.read_sql_query("""
        SELECT filename, report_date, me_total_hrs, me_this_month,
               parse_confidence, component_count, warning_count, uploaded_at
        FROM upload_log
        WHERE vessel_name=?
        ORDER BY uploaded_at DESC
        LIMIT 20
    """, c, params=(vessel,))
    c.close()
    return df

@st.cache_data(ttl=10)
def get_summary():
    c = get_db()
    df = pd.read_sql_query("""
        SELECT c.vessel_name,
               COUNT(CASE WHEN c.status='OVERDUE' THEN 1 END) AS overdue,
               COUNT(CASE WHEN c.status='HIGH PRIORITY' THEN 1 END) AS high_priority,
               COUNT(CASE WHEN c.status='OK' THEN 1 END) AS ok,
               COUNT(CASE WHEN c.status='NO DATA' THEN 1 END) AS no_data,
               COUNT(*) AS total,
               MAX(u.uploaded_at) AS last_upload,
               MAX(u.me_total_hrs) AS me_total_hrs,
               MAX(u.report_date) AS report_date
        FROM components c
        LEFT JOIN upload_log u ON u.vessel_name=c.vessel_name
        GROUP BY c.vessel_name
        ORDER BY overdue DESC, high_priority DESC, c.vessel_name
    """, c)
    c.close()
    return df

@st.cache_data(ttl=10)
def get_all_fleet_comps():
    c = get_db()
    df = pd.read_sql_query("SELECT * FROM components", c)
    c.close()
    return df


# ═══════════════════════════════════════════════════════════════════
#  RENDER ENGINE
# ═══════════════════════════════════════════════════════════════════
def _sf(x):
    try:
        v = float(x)
        return None if pd.isna(v) else v
    except Exception:
        return None

def _cyl(u):
    m = re.search(r"\d+", str(u))
    return int(m.group()) if m else 999

_S = {
    "OVERDUE": {"bg":"#1f0505","bgs":"#2d0707","ts":"#ff6b6b","tm":"#ff8080","tn":"#ff3333","td":"#773333"},
    "HIGH PRIORITY": {"bg":"#1e0d02","bgs":"#2d1503","ts":"#ffaa44","tm":"#ff9933","tn":"#ffcc00","td":"#774422"},
    "OK": {"bg":"#021208","bgs":"#042010","ts":"#4ade80","tm":"#22c55e","tn":"#4ade80","td":"#0f4023"},
    "NO DATA": {"bg":"#090e18","bgs":"#0c1422","ts":"#7da3d8","tm":"#5f7fa6","tn":"#7da3d8","td":"#2a3950"},
    "_": {"bg":"#090e18","bgs":"#0c1422","ts":"#4a6688","tm":"#334d66","tn":"#334d66","td":"#1a2a38"},
}

_COL_CFG = {
    "Status": st.column_config.TextColumn("Status", width=130),
    "Vessel": st.column_config.TextColumn("Vessel", width=140),
    "Component": st.column_config.TextColumn("Component", width=205),
    "Engine": st.column_config.TextColumn("Engine", width=90),
    "Unit": st.column_config.TextColumn("Unit", width=80),
    "Periodicity": st.column_config.NumberColumn("Periodicity", format="%d", width=100),
    "Last O/H": st.column_config.TextColumn("Last O/H", width=100),
    "Hrs Since": st.column_config.NumberColumn("Hrs Since", format="%d hrs", width=100),
    "% Used": st.column_config.ProgressColumn("% Used", min_value=0, max_value=160, format="%.1f%%", width=120),
}

def _build_display(df: pd.DataFrame, sort_priority: bool = False) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Status","Vessel","Component","Engine","Unit","Periodicity","Last O/H","Hrs Since","% Used"])

    d = df.copy()
    if sort_priority:
        ORD = {"OVERDUE":0,"HIGH PRIORITY":1,"OK":2,"NO DATA":3}
        d["_s"] = d["status"].map(lambda s: ORD.get(str(s), 4))
        d["_p"] = d["pct_used"].apply(lambda x: _sf(x) or 0.0)
        if "vessel_name" in d.columns:
            d = d.sort_values(["_s","_p","vessel_name"], ascending=[True,False,True]).drop(columns=["_s","_p"])
        else:
            d = d.sort_values(["_s","_p"], ascending=[True,False]).drop(columns=["_s","_p"])
    else:
        d["_k1"] = d["description"].astype(str).str.upper()
        d["_k2"] = d["unit"].apply(_cyl)
        if "vessel_name" in d.columns:
            d = d.sort_values(["vessel_name","_k1","_k2"]).drop(columns=["_k1","_k2"])
        else:
            d = d.sort_values(["_k1","_k2"]).drop(columns=["_k1","_k2"])

    out = pd.DataFrame(index=range(len(d)))
    out["Status"] = d["status"].values
    out["Vessel"] = d.get("vessel_name", pd.Series(["—"] * len(d))).values
    out["Component"] = d["description"].values
    out["Engine"] = d["engine_label"].values
    out["Unit"] = d["unit"].values
    out["Periodicity"] = [safe_int(x) for x in d["periodicity"].values]
    out["Last O/H"] = [str(x) if x and str(x) not in ("nan","None","") else "—" for x in d["last_oh_date"].values]
    out["Hrs Since"] = [safe_int(x) for x in d["hrs_since"].values]
    out["% Used"] = [round(float(x)*100,1) if _sf(x) else 0.0 for x in d["pct_used"].values]
    return out

def _apply_style(df: pd.DataFrame):
    def rs(row):
        c = _S.get(str(row.get("Status","")), _S["_"])
        return [
            f"background-color:{c['bgs']};color:{c['ts']};font-weight:700",
            f"background-color:{c['bg']};color:#c0d0e8;font-weight:700",
            f"background-color:{c['bg']};color:{c['tm']};font-weight:600",
            f"background-color:{c['bg']};color:{c['td']}",
            f"background-color:{c['bg']};color:{c['td']}",
            f"background-color:{c['bg']};color:{c['td']}",
            f"background-color:{c['bg']};color:{c['td']}",
            f"background-color:{c['bg']};color:{c['tm']};font-weight:600",
            f"background-color:{c['bg']};color:{c['tn']};font-weight:700",
        ]
    return df.style.apply(rs, axis=1)

def render_table(df: pd.DataFrame, height: int = None, priority: bool = False):
    if isinstance(df, list):
        df = pd.DataFrame(df)
    if df.empty:
        st.info("No data to display.")
        return
    tbl = _build_display(df, sort_priority=priority)
    h = height or min(720, 38 * (len(tbl) + 1) + 4)
    st.dataframe(_apply_style(tbl), use_container_width=True, hide_index=True, height=h, column_config=_COL_CFG)


# ═══════════════════════════════════════════════════════════════════
#  UI HELPERS
# ═══════════════════════════════════════════════════════════════════
def kpi(val, lbl, color="gold", delay=0):
    return f'<div class="kc {color}" style="animation-delay:{delay}s"><div class="kc-val">{val}</div><div class="kc-lbl">{lbl}</div></div>'

def ph(icon, title, eye=""):
    e = f'<div class="ph-eye">{eye}</div>' if eye else ""
    return f'<div class="ph">{e}<h1>{icon}&nbsp;&nbsp;{title}</h1></div><div class="ph-line"></div>'

def sl(txt):
    return f'<div class="sl">{txt}</div>'

def filter_count(total, od, hp, ok):
    return (
        f'<div class="filter-count">Showing <b>{total}</b> records — '
        f'<span class="od">{od} overdue</span> · '
        f'<span class="hp">{hp} high priority</span> · '
        f'<span class="ok">{ok} OK</span></div>'
    )


# ═══════════════════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown('<div class="logo">FLEET MONITOR</div><div class="logo-tag">Running Hours Management System</div><div class="logo-rule"></div>', unsafe_allow_html=True)
    page = st.selectbox("nav", ["Fleet Overview", "Vessel Detail", "Upload Report", "Upload History"], label_visibility="collapsed")
    st.markdown("<br>", unsafe_allow_html=True)

    vessels = get_vessels()
    selv = st.selectbox("Active Vessel", vessels if vessels else [None])

    if not vessels:
        st.info("No data loaded. Upload a report to begin.")

    if vessels:
        smry = get_summary()
        if not smry.empty:
            st.markdown('<div class="logo-rule"></div>', unsafe_allow_html=True)
            st.markdown('<div style="font-family:var(--fi);font-size:.56rem;text-transform:uppercase;letter-spacing:.2em;color:var(--t3);margin-bottom:.5rem">Vessel Status</div>', unsafe_allow_html=True)
            for idx, (_, row) in enumerate(smry.iterrows()):
                od = int(row["overdue"])
                hp = int(row["high_priority"])
                cls = "crit" if od > 0 else "warn" if hp > 0 else "safe"
                tags = ""
                if od > 0:
                    tags += f'<span class="vt od">{od} OD</span>'
                if hp > 0:
                    tags += f'<span class="vt hp">{hp} HP</span>'
                if od == 0 and hp == 0:
                    tags += '<span class="vt ok">OK</span>'
                st.markdown(
                    f'<div class="vc {cls}" style="animation-delay:{idx*0.04}s"><div class="vc-name">{row["vessel_name"]}</div><div class="vc-tags">{tags}</div></div>',
                    unsafe_allow_html=True
                )

    st.markdown('<div class="logo-rule"></div>', unsafe_allow_html=True)
    dbkb = DB_PATH.stat().st_size / 1024 if DB_PATH.exists() else 0
    st.markdown(f'<div style="font-family:var(--fm);font-size:.58rem;color:var(--t3)">db {dbkb:.0f} kb · {len(vessels)} vessels · v6.1 dual-aux parser</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════
#  PAGES
# ═══════════════════════════════════════════════════════════════════
if page == "Fleet Overview":
    st.markdown(ph("⚓", "Fleet Master Matrix", "Universal Fleet Telemetry"), unsafe_allow_html=True)

    smry = get_summary()
    allcomps = get_all_fleet_comps()

    if smry.empty or allcomps.empty:
        st.info("No data loaded. Upload a report to begin.")
        st.stop()

    tv = len(smry)
    tc = len(allcomps)
    tod = int((allcomps["status"] == "OVERDUE").sum())
    thp = int((allcomps["status"] == "HIGH PRIORITY").sum())
    tok = int((allcomps["status"] == "OK").sum())

    k1, k2, k3, k4, k5 = st.columns(5)
    for col, val, lbl, clr, dly in zip(
        [k1, k2, k3, k4, k5],
        [tv, tc, tod, thp, tok],
        ["Vessels", "Components", "Overdue", "High Priority", "OK"],
        ["blue", "gold", "red", "orange", "green"],
        [0, .07, .14, .21, .28]
    ):
        with col:
            st.markdown(kpi(val, lbl, clr, dly), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(sl("Universal Component Control Grid"), unsafe_allow_html=True)

    f1, f2, f3, f4 = st.columns([1.5, 1.5, 2, 2])
    with f1:
        vesself = st.selectbox("Filter Vessel Context", ["All Fleet"] + sorted(allcomps["vessel_name"].unique().tolist()), key="mstv")
    with f2:
        catf = st.selectbox("Filter Machinery Type", ["All", "Main Engine", "Aux Engines"], key="mstcat")
    with f3:
        stf = st.selectbox("Filter Component Urgency", ["Critical Focus", "All Statuses", "Overdue Only", "High Priority Only", "OK Only", "No Data Only"], key="mstst")
    with f4:
        compf = st.selectbox("Search Component Definition", ["All"] + sorted(allcomps["description"].unique().tolist()), key="mstcomp")

    filt = allcomps.copy()
    if vesself != "All Fleet":
        filt = filt[filt["vessel_name"] == vesself]
    if catf == "Main Engine":
        filt = filt[filt["category"] == "MAINENGINE"]
    elif catf == "Aux Engines":
        filt = filt[filt["category"] == "AUXENGINE"]

    if stf == "Critical Focus":
        filt = filt[filt["status"].isin(["OVERDUE", "HIGH PRIORITY"])]
    elif stf == "Overdue Only":
        filt = filt[filt["status"] == "OVERDUE"]
    elif stf == "High Priority Only":
        filt = filt[filt["status"] == "HIGH PRIORITY"]
    elif stf == "OK Only":
        filt = filt[filt["status"] == "OK"]
    elif stf == "No Data Only":
        filt = filt[filt["status"] == "NO DATA"]

    if compf != "All":
        filt = filt[filt["description"] == compf]

    ns = len(filt)
    no = int((filt["status"] == "OVERDUE").sum())
    nh = int((filt["status"] == "HIGH PRIORITY").sum())
    nk = int((filt["status"] == "OK").sum())

    st.markdown(filter_count(ns, no, nh, nk), unsafe_allow_html=True)

    if filt.empty:
        st.markdown('<div class="ac">No records match the current filter matrix</div>', unsafe_allow_html=True)
    else:
        render_table(filt, height=min(900, 38 * (ns + 1) + 4), priority=True)

elif page == "Vessel Detail":
    if not selv:
        st.info("Select a vessel from the sidebar.")
        st.stop()

    st.markdown(ph("🚢", selv, "Component Analysis"), unsafe_allow_html=True)
    df = get_comps(selv)
    oe = get_oe(selv)

    if df.empty:
        st.info("No data for this vessel.")
        st.stop()

    ntot = len(df)
    nod = int((df["status"] == "OVERDUE").sum())
    nhp = int((df["status"] == "HIGH PRIORITY").sum())
    nok = int((df["status"] == "OK").sum())
    nnd = int((df["status"] == "NO DATA").sum())

    k1, k2, k3, k4, k5 = st.columns(5)
    for col, val, lbl, clr, dly in zip(
        [k1, k2, k3, k4, k5],
        [ntot, nod, nhp, nok, nnd],
        ["Total", "Overdue", "High Priority", "OK", "No Data"],
        ["gold", "red", "orange", "green", "blue"],
        [0, .07, .14, .21, .28]
    ):
        with col:
            st.markdown(kpi(val, lbl, clr, dly), unsafe_allow_html=True)

    hist = get_history(selv)
    if not hist.empty:
        last = hist.iloc[0]
        mtf = f"{int(last['me_total_hrs']):,}" if pd.notna(last["me_total_hrs"]) else "—"
        mmf = f"{int(last['me_this_month']):,}" if pd.notna(last["me_this_month"]) else "—"
        conf = f"{float(last['parse_confidence']):.2f}" if pd.notna(last["parse_confidence"]) else "—"
        compc = int(last["component_count"]) if pd.notna(last["component_count"]) else 0
        warnc = int(last["warning_count"]) if pd.notna(last["warning_count"]) else 0

        st.markdown(
            f'<div class="ml"><span>File <b>{last["filename"]}</b></span><span>Report <b>{last["report_date"] or "—"}</b></span><span>ME <b>{mtf}</b> total / <b>{mmf}</b> month</span><span>Confidence <b>{conf}</b></span><span>Rows <b>{compc}</b></span><span>Warnings <b>{warnc}</b></span><span>Uploaded <b>{str(last["uploaded_at"])[:16]}</b></span></div>',
            unsafe_allow_html=True
        )

    st.markdown("---")
    tabs = st.tabs(["Alerts", "Main Engine", "Aux Engines", "Other Equipment"])

    with tabs[0]:
        st.markdown(sl("Urgent Interrupt Diagnostics"), unsafe_allow_html=True)
        alerts = df[df["status"].isin(["OVERDUE", "HIGH PRIORITY"])]
        if alerts.empty:
            st.markdown('<div class="ac">All machinery components within acceptable operational bounds</div>', unsafe_allow_html=True)
        else:
            no = int((alerts["status"] == "OVERDUE").sum())
            nh = int((alerts["status"] == "HIGH PRIORITY").sum())
            st.markdown(filter_count(len(alerts), no, nh, 0), unsafe_allow_html=True)
            render_table(alerts, priority=True)

    with tabs[1]:
        me = df[df["category"] == "MAINENGINE"]
        if me.empty:
            st.info("No Main Engine data available.")
        else:
            st.markdown(sl("Main Engine Telemetry Matrix"), unsafe_allow_html=True)
            fa, fb = st.columns(2)
            with fa:
                selmc = st.selectbox("Machinery Element", ["All"] + sorted(me["description"].unique().tolist()), key="mec")
            with fb:
                selms = st.selectbox("Machinery Status", ["All", "Overdue only", "High Priority only", "OK only", "No Data only"], key="mes")

            v = me.copy()
            if selmc != "All":
                v = v[v["description"] == selmc]
            if selms == "Overdue only":
                v = v[v["status"] == "OVERDUE"]
            elif selms == "High Priority only":
                v = v[v["status"] == "HIGH PRIORITY"]
            elif selms == "OK only":
                v = v[v["status"] == "OK"]
            elif selms == "No Data only":
                v = v[v["status"] == "NO DATA"]

            no = int((v["status"] == "OVERDUE").sum())
            nh = int((v["status"] == "HIGH PRIORITY").sum())
            nk = int((v["status"] == "OK").sum())
            st.markdown(filter_count(len(v), no, nh, nk), unsafe_allow_html=True)
            render_table(v, priority=False)

    with tabs[2]:
        aux = df[df["category"] == "AUXENGINE"]
        if aux.empty:
            st.info("No Auxiliary Engine data available.")
        else:
            st.markdown(sl("Auxiliary Prime Movers Telemetry Matrix"), unsafe_allow_html=True)
            fa, fb = st.columns(2)
            with fa:
                selae = st.selectbox("Aux Generator Node", ["All"] + sorted(aux["engine_label"].unique().tolist()), key="auxe")
            with fb:
                selas = st.selectbox("Node Condition", ["All", "Overdue only", "High Priority only", "OK only", "No Data only"], key="auxs")

            v = aux.copy()
            if selae != "All":
                v = v[v["engine_label"] == selae]
            if selas == "Overdue only":
                v = v[v["status"] == "OVERDUE"]
            elif selas == "High Priority only":
                v = v[v["status"] == "HIGH PRIORITY"]
            elif selas == "OK only":
                v = v[v["status"] == "OK"]
            elif selas == "No Data only":
                v = v[v["status"] == "NO DATA"]

            no = int((v["status"] == "OVERDUE").sum())
            nh = int((v["status"] == "HIGH PRIORITY").sum())
            nk = int((v["status"] == "OK").sum())
            st.markdown(filter_count(len(v), no, nh, nk), unsafe_allow_html=True)
            render_table(v, priority=False)

    with tabs[3]:
        if oe.empty:
            st.info("No auxiliary plant or extension machinery metrics located.")
        else:
            for sec in sorted(oe["section"].unique()):
                st.markdown(sl(sec), unsafe_allow_html=True)
                sd = oe[oe["section"] == sec][["description", "periodicity", "last_date", "run_hrs"]].copy()
                sd.columns = ["Machinery Description", "Maintenance Periodicity", "Inspection Date", "Logged Hours"]
                st.dataframe(sd, use_container_width=True, hide_index=True)

elif page == "Upload Report":
    st.markdown(ph("📤", "Upload Report", "TEC-004 Log Processing"), unsafe_allow_html=True)

    colup, colinfo = st.columns([3, 2], gap="large")
    with colup:
        uploaded = st.file_uploader("file", type=["doc"], label_visibility="collapsed")
    with colinfo:
        st.markdown("""
        <div class='ic'>
            <div class='ic-title'>Accepted Specification</div>
            TEC-004 Running Hours Monthly Log Report<br>
            Supported format: <b>native .doc</b><br><br>
            <div class='ic-title'>Extraction Coverage</div>
            Vessel identity and report date<br>
            Main engine matrix<br>
            Aux engine row-pair block<br>
            DG No1 / No2 / No3 maintenance matrix<br>
            Other equipment sections<br>
            Priority classification and dashboard persistence
        </div>
        """, unsafe_allow_html=True)

    if uploaded:
        raw = uploaded.read()
        fh = hashlib.md5(raw).hexdigest()

        with st.spinner("Executing secure LibreOffice conversion and structural extraction loops..."):
            try:
                docx = convert_doc_to_docx(raw)
            except Exception as e:
                st.error(f"Headless conversion failed: {e}")
                st.stop()

            try:
                parsed = parse_doc_bytes(docx)
            except ValueError as e:
                st.error(f"Telemetry stream interpretation broke down: {e}")
                st.stop()

        comps = parsed["components"]
        nc = len(comps)
        nod = sum(1 for c in comps if c["status"] == "OVERDUE")
        nhp = sum(1 for c in comps if c["status"] == "HIGH PRIORITY")
        nok = sum(1 for c in comps if c["status"] == "OK")
        noe = len(parsed["other_equipment"])
        conf = parsed["parse_confidence"]

        st.markdown(sl("Extracted Telemetry Stream Preview"), unsafe_allow_html=True)

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Asset", parsed["vessel_name"])
        c2.metric("Report Window", parsed["report_date"] or "—")
        c3.metric("ME Accumulated", f"{parsed['me_total_hrs']:,}" if parsed["me_total_hrs"] else "—")
        c4.metric("Monthly Increment", f"{parsed['me_this_month']:,}" if parsed["me_this_month"] else "—")
        c5.metric("Data Channels", nc)
        c6.metric("Confidence", f"{conf:.2f}")

        st.markdown(f"""
        <div class='ps-row'>
          <div class='ps red'><div class='ps-val'>{nod}</div><div class='ps-lbl'>Overdue</div></div>
          <div class='ps orange'><div class='ps-val'>{nhp}</div><div class='ps-lbl'>High Priority</div></div>
          <div class='ps green'><div class='ps-val'>{nok}</div><div class='ps-lbl'>OK</div></div>
          <div class='ps blue'><div class='ps-val'>{noe}</div><div class='ps-lbl'>External Systems</div></div>
        </div>
        """, unsafe_allow_html=True)

        for w in parsed["warnings"]:
            st.warning(f"Structural warning: {w}")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(sl("Extracted Telemetry Matrices · Pre-Commit Review"), unsafe_allow_html=True)

        dfpreview = pd.DataFrame(parsed["components"]) if parsed["components"] else pd.DataFrame()
        prevtabs = st.tabs(["Main Engine Matrix", "Aux Engines Matrix", "Other Equipment"])

        with prevtabs[0]:
            if not dfpreview.empty:
                meprev = dfpreview[dfpreview["category"] == "MAINENGINE"]
                if not meprev.empty:
                    render_table(meprev, height=400, priority=True)
                else:
                    st.info("No Main Engine telemetry extracted from this report.")
            else:
                st.info("No component data available.")

        with prevtabs[1]:
            if not dfpreview.empty:
                auxprev = dfpreview[dfpreview["category"] == "AUXENGINE"]
                if not auxprev.empty:
                    render_table(auxprev, height=400, priority=True)
                else:
                    st.info("No Auxiliary Engine telemetry extracted from this report.")
            else:
                st.info("No component data available.")

        with prevtabs[2]:
            if parsed["other_equipment"]:
                oeprev = pd.DataFrame(parsed["other_equipment"])
                oeprev.columns = ["Machinery Category", "Description", "Periodicity", "Last Inspected", "Logged Hours"]
                st.dataframe(oeprev, use_container_width=True, hide_index=True, height=400)
            else:
                st.info("No Auxiliary Plant or Other Equipment data extracted from this report.")

        st.markdown("---")
        colbtn, _ = st.columns([1, 4])
        with colbtn:
            if st.button("COMMIT STREAM TO DATABASE", use_container_width=True):
                save_parsed(parsed, uploaded.name, fh)
                for fn in [get_vessels, get_comps, get_oe, get_summary, get_all_fleet_comps, get_history]:
                    fn.clear()
                st.markdown(
                    f'<div class="sb"><span style="font-size:1.4rem">✓</span><span>System Telemetry Confirmed · <b>{parsed["vessel_name"]}</b> committed to database · {nc} lines mapped.</span></div>',
                    unsafe_allow_html=True
                )
                st.balloons()

elif page == "Upload History":
    st.markdown(ph("🕘", "Upload History", "System Audit Trails"), unsafe_allow_html=True)

    if not selv:
        st.info("Select a tracking asset from the sidebar selector.")
        st.stop()

    st.markdown(sl(f"Chronological Logs · {selv}"), unsafe_allow_html=True)
    hist = get_history(selv)
    if hist.empty:
        st.info("No transaction trail entries recorded for this hull identification.")
    else:
        d = hist.copy()
        d.columns = [
            "Logged Filename", "Extracted Target Date", "ME Combined Total",
            "ME Monthly Increment", "Parse Confidence", "Component Count",
            "Warning Count", "Transaction Timestamp"
        ]
        d["ME Combined Total"] = d["ME Combined Total"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
        d["ME Monthly Increment"] = d["ME Monthly Increment"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
        st.dataframe(d, use_container_width=True, hide_index=True)
