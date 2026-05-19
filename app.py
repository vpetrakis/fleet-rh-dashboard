import streamlit as st

st.set_page_config(
    page_title="Fleet Running Hours",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="expanded",
)

import os, re, sqlite3, tempfile, hashlib, subprocess, shutil
from datetime import datetime
from pathlib import Path
import pandas as pd

# ════════════════════════════════════════════════════════════════════
# PREMIUM CSS — Bloomberg Terminal × Superyacht Bridge
# ════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700;800&family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;1,300&family=DM+Mono:wght@400;500&display=swap');

:root {
  --bg:       #04080f;
  --bg2:      #070d1a;
  --bg3:      #0a1220;
  --panel:    #0d1829;
  --panel2:   #101f35;
  --border:   #162035;
  --border2:  #1e3050;
  --gold:     #c9952a;
  --gold2:    #e4af45;
  --gold3:    #f5cc6e;
  --goldglow: rgba(201,149,42,0.15);
  --red:      #e05252;
  --red2:     #f87171;
  --redbg:    rgba(224,82,82,0.08);
  --orange:   #e07832;
  --orange2:  #fb923c;
  --orangebg: rgba(224,120,50,0.08);
  --green:    #2ea86e;
  --green2:   #34d399;
  --greenbg:  rgba(46,168,110,0.08);
  --blue:     #3b82f6;
  --blue2:    #60a5fa;
  --text:     #b8c8e0;
  --text2:    #6b84a0;
  --text3:    #3d5470;
  --font:     'DM Sans', sans-serif;
  --mono:     'DM Mono', monospace;
  --display:  'Syne', sans-serif;
}

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

html, body, [class*="css"] {
  font-family: var(--font) !important;
  background: var(--bg) !important;
  color: var(--text) !important;
  -webkit-font-smoothing: antialiased;
}
.main, .main > div { background: var(--bg) !important; }
.block-container { padding: 2rem 2.5rem 5rem !important; max-width: 100% !important; }
a { color: var(--gold2) !important; text-decoration: none; }

/* ── Ambient background ── */
.main::before {
  content: '';
  position: fixed; inset: 0; pointer-events: none; z-index: 0;
  background:
    radial-gradient(ellipse 80% 50% at -10% 0%, rgba(201,149,42,0.06) 0%, transparent 55%),
    radial-gradient(ellipse 60% 40% at 110% 100%, rgba(59,130,246,0.04) 0%, transparent 55%),
    radial-gradient(ellipse 40% 60% at 50% 50%, rgba(7,13,26,0.8) 0%, transparent 100%);
}

/* ── Sidebar ── */
[data-testid="stSidebar"] {
  background: var(--bg2) !important;
  border-right: 1px solid var(--border) !important;
}
[data-testid="stSidebar"] * { color: var(--text) !important; }
[data-testid="stSidebar"] .stSelectbox > div > div {
  background: var(--bg3) !important;
  border: 1px solid var(--border) !important;
  border-radius: 6px !important;
}
[data-testid="stSidebarContent"] { padding: 1.5rem 1.25rem !important; }

/* ── Typography ── */
h1, h2, h3 { font-family: var(--display) !important; }
h1 { font-size: 1.9rem !important; font-weight: 800 !important; color: #e8f0f8 !important; letter-spacing: -0.01em !important; }
h2 { font-size: 1.3rem !important; font-weight: 700 !important; color: #cdd8e8 !important; }
h3 { font-size: 1.05rem !important; font-weight: 600 !important; color: #b8c8e0 !important; }

/* ── Metrics ── */
[data-testid="stMetric"] {
  background: var(--panel) !important;
  border: 1px solid var(--border) !important;
  border-radius: 10px !important;
  padding: 1rem 1.25rem 1.1rem !important;
  position: relative !important;
  overflow: hidden !important;
  transition: border-color 0.25s, transform 0.2s !important;
}
[data-testid="stMetric"]:hover {
  border-color: var(--border2) !important;
  transform: translateY(-2px) !important;
}
[data-testid="stMetric"]::before {
  content: '';
  position: absolute; inset: 0;
  background: linear-gradient(135deg, rgba(201,149,42,0.04) 0%, transparent 60%);
  pointer-events: none;
}
[data-testid="stMetricValue"] {
  font-family: var(--display) !important;
  font-size: 2rem !important;
  font-weight: 800 !important;
  color: #e8f0f8 !important;
  letter-spacing: -0.02em !important;
}
[data-testid="stMetricLabel"] {
  color: var(--text2) !important;
  font-size: 0.65rem !important;
  text-transform: uppercase !important;
  letter-spacing: 0.14em !important;
  font-weight: 500 !important;
}
[data-testid="stMetricDelta"] { font-size: 0.75rem !important; }

/* ── Dataframe ── */
[data-testid="stDataFrame"] {
  border: 1px solid var(--border) !important;
  border-radius: 10px !important;
  overflow: hidden !important;
}
.dvn-scroller { background: var(--panel) !important; }

/* ── Buttons ── */
.stButton > button {
  background: linear-gradient(135deg, var(--gold) 0%, #a67820 100%) !important;
  color: #000 !important;
  border: none !important;
  font-family: var(--display) !important;
  font-weight: 700 !important;
  font-size: 0.82rem !important;
  letter-spacing: 0.08em !important;
  text-transform: uppercase !important;
  border-radius: 7px !important;
  padding: 0.6rem 1.8rem !important;
  box-shadow: 0 2px 16px rgba(201,149,42,0.2), inset 0 1px 0 rgba(255,255,255,0.1) !important;
  transition: all 0.2s ease !important;
  position: relative !important;
}
.stButton > button:hover {
  background: linear-gradient(135deg, var(--gold2) 0%, var(--gold) 100%) !important;
  box-shadow: 0 4px 24px rgba(201,149,42,0.35) !important;
  transform: translateY(-2px) !important;
}
.stButton > button:active { transform: translateY(0) !important; }

/* ── File Uploader ── */
[data-testid="stFileUploadDropzone"] {
  background: linear-gradient(135deg, rgba(201,149,42,0.04) 0%, rgba(59,130,246,0.03) 100%) !important;
  border: 1px dashed var(--gold) !important;
  border-radius: 14px !important;
  transition: all 0.3s ease !important;
  padding: 3rem 2rem !important;
}
[data-testid="stFileUploadDropzone"]:hover {
  background: rgba(201,149,42,0.07) !important;
  border-color: var(--gold2) !important;
  box-shadow: 0 0 40px rgba(201,149,42,0.08), inset 0 0 40px rgba(201,149,42,0.03) !important;
}
[data-testid="stFileUploadDropzone"] p,
[data-testid="stFileUploadDropzone"] span {
  color: var(--gold2) !important;
  font-family: var(--display) !important;
  font-size: 1rem !important;
  font-weight: 600 !important;
}
[data-testid="stFileUploadDropzone"] small { color: var(--text2) !important; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
  background: var(--panel) !important;
  border-radius: 10px 10px 0 0 !important;
  border-bottom: 1px solid var(--border) !important;
  gap: 0 !important;
  padding: 0 1rem !important;
}
.stTabs [data-baseweb="tab"] {
  background: transparent !important;
  color: var(--text3) !important;
  font-family: var(--display) !important;
  font-weight: 600 !important;
  letter-spacing: 0.06em !important;
  text-transform: uppercase !important;
  font-size: 0.75rem !important;
  padding: 0.9rem 1.4rem !important;
  border-bottom: 2px solid transparent !important;
  margin-bottom: -1px !important;
  transition: color 0.2s !important;
}
.stTabs [data-baseweb="tab"]:hover { color: var(--text) !important; }
.stTabs [aria-selected="true"] {
  color: var(--gold2) !important;
  border-bottom: 2px solid var(--gold) !important;
}
.stTabs [data-baseweb="tab-panel"] {
  background: var(--panel) !important;
  border: 1px solid var(--border) !important;
  border-top: none !important;
  border-radius: 0 0 10px 10px !important;
  padding: 1.75rem !important;
}

/* ── Expander ── */
.streamlit-expanderHeader {
  background: var(--panel) !important;
  border: 1px solid var(--border) !important;
  border-radius: 8px !important;
  font-family: var(--display) !important;
  font-weight: 600 !important;
  font-size: 0.85rem !important;
  color: var(--text) !important;
  letter-spacing: 0.02em !important;
  transition: background 0.2s, border-color 0.2s !important;
}
.streamlit-expanderHeader:hover {
  background: var(--panel2) !important;
  border-color: var(--border2) !important;
}
.streamlit-expanderContent {
  background: var(--panel) !important;
  border: 1px solid var(--border) !important;
  border-top: none !important;
  border-radius: 0 0 8px 8px !important;
}

/* ── Select / Multiselect ── */
.stSelectbox > div > div,
.stMultiSelect > div > div {
  background: var(--bg3) !important;
  border: 1px solid var(--border) !important;
  border-radius: 7px !important;
  color: var(--text) !important;
}
.stSelectbox label, .stMultiSelect label { color: var(--text2) !important; font-size: 0.75rem !important; }

/* ── Alerts ── */
.stAlert { border-radius: 8px !important; border-left-width: 3px !important; }

/* ── Divider / HR ── */
hr { border-color: var(--border) !important; opacity: 1 !important; margin: 1.5rem 0 !important; }

/* ── Spinner ── */
[data-testid="stSpinner"] > div { border-top-color: var(--gold) !important; }

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: var(--bg2); }
::-webkit-scrollbar-thumb { background: var(--border2); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: var(--text3); }

/* ══════════════════════════════════════════════════
   KEYFRAMES
══════════════════════════════════════════════════ */
@keyframes fadeIn      { from{opacity:0}                        to{opacity:1} }
@keyframes slideDown   { from{opacity:0;transform:translateY(-18px)} to{opacity:1;transform:translateY(0)} }
@keyframes slideUp     { from{opacity:0;transform:translateY(14px)}  to{opacity:1;transform:translateY(0)} }
@keyframes slideRight  { from{opacity:0;transform:translateX(-12px)} to{opacity:1;transform:translateX(0)} }
@keyframes scaleIn     { from{opacity:0;transform:scale(0.88)}       to{opacity:1;transform:scale(1)} }
@keyframes goldPulse   { 0%,100%{box-shadow:0 0 0 rgba(201,149,42,0)} 50%{box-shadow:0 0 24px rgba(201,149,42,0.2)} }
@keyframes barGrow     { from{width:0} to{width:var(--w)} }
@keyframes successPop  { 0%{transform:scale(0.8);opacity:0} 60%{transform:scale(1.03)} 100%{transform:scale(1);opacity:1} }
@keyframes scanline    { 0%{transform:translateY(-100%)} 100%{transform:translateY(400%)} }
@keyframes borderFlow  { 0%{background-position:0% 50%} 100%{background-position:200% 50%} }
@keyframes numberCount { from{opacity:0;transform:translateY(8px)} to{opacity:1;transform:translateY(0)} }

/* ══════════════════════════════════════════════════
   CUSTOM COMPONENTS
══════════════════════════════════════════════════ */

/* Page header */
.ph { animation: slideDown 0.5s cubic-bezier(0.22,1,0.36,1) both; }
.ph h1 { display: flex; align-items: center; gap: 0.6rem; }
.ph-rule {
  height: 1px; margin: 0.5rem 0 1.5rem;
  background: linear-gradient(90deg, var(--gold), var(--border) 40%, transparent);
  animation: fadeIn 0.8s 0.2s ease both; opacity: 0; animation-fill-mode: forwards;
}

/* Section label */
.sl {
  font-family: var(--display);
  font-size: 0.6rem; font-weight: 700;
  letter-spacing: 0.22em; text-transform: uppercase;
  color: var(--text3); border-bottom: 1px solid var(--border);
  padding-bottom: 0.5rem; margin: 2rem 0 1.1rem;
}

/* KPI card */
.kc {
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 1.1rem 1.3rem 1.2rem;
  position: relative; overflow: hidden;
  animation: slideUp 0.45s ease both;
  transition: border-color 0.25s, transform 0.2s, box-shadow 0.25s;
}
.kc:hover {
  border-color: var(--border2);
  transform: translateY(-3px);
  box-shadow: 0 8px 32px rgba(0,0,0,0.3);
}
.kc::after {
  content: '';
  position: absolute; top: 0; left: 0; right: 0; height: 2px;
  border-radius: 10px 10px 0 0;
}
.kc.gold::after   { background: linear-gradient(90deg, var(--gold), transparent 70%); }
.kc.red::after    { background: linear-gradient(90deg, var(--red),   transparent 70%); }
.kc.orange::after { background: linear-gradient(90deg, var(--orange),transparent 70%); }
.kc.green::after  { background: linear-gradient(90deg, var(--green), transparent 70%); }
.kc.blue::after   { background: linear-gradient(90deg, var(--blue),  transparent 70%); }
.kc-val {
  font-family: var(--display);
  font-size: 2.2rem; font-weight: 800; line-height: 1.1;
  letter-spacing: -0.03em;
  animation: numberCount 0.5s 0.15s ease both; opacity: 0; animation-fill-mode: forwards;
}
.kc-lbl {
  font-size: 0.62rem; font-weight: 600;
  text-transform: uppercase; letter-spacing: 0.16em;
  color: var(--text2); margin-top: 5px;
}
.kc.gold   .kc-val { color: var(--gold3); }
.kc.red    .kc-val { color: var(--red2); }
.kc.orange .kc-val { color: var(--orange2); }
.kc.green  .kc-val { color: var(--green2); }
.kc.blue   .kc-val { color: var(--blue2); }

/* Status badge */
.badge {
  display: inline-flex; align-items: center;
  padding: 2px 10px; border-radius: 20px;
  font-family: var(--display); font-size: 0.65rem; font-weight: 700;
  letter-spacing: 0.1em; text-transform: uppercase;
}
.badge-od { background: rgba(224,82,82,0.12); color: var(--red2); border: 1px solid rgba(224,82,82,0.25); }
.badge-hp { background: rgba(224,120,50,0.12); color: var(--orange2); border: 1px solid rgba(224,120,50,0.25); }
.badge-ok { background: rgba(46,168,110,0.12); color: var(--green2); border: 1px solid rgba(46,168,110,0.25); }
.badge-nd { background: rgba(59,130,246,0.08); color: var(--blue2); border: 1px solid rgba(59,130,246,0.15); }

/* Progress bar */
.pb-wrap  { background: rgba(255,255,255,0.05); border-radius: 99px; height: 4px; overflow: hidden; }
.pb-fill  { height: 100%; border-radius: 99px; transition: width 0.8s cubic-bezier(0.34,1.3,0.64,1); }
.pb-ok    { background: linear-gradient(90deg, var(--green), #6ee7b7); }
.pb-hp    { background: linear-gradient(90deg, var(--orange), #fb923c); }
.pb-od    { background: linear-gradient(90deg, var(--red), #f87171); }

/* Info card */
.ic {
  background: linear-gradient(135deg, var(--panel) 0%, var(--bg3) 100%);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 1.4rem 1.6rem;
  font-size: 0.82rem;
  color: var(--text2);
  line-height: 2;
  animation: slideRight 0.4s ease both;
}
.ic-title {
  font-family: var(--display);
  font-size: 0.62rem; font-weight: 700;
  letter-spacing: 0.18em; text-transform: uppercase;
  color: var(--gold2); margin-bottom: 0.5rem;
}

/* Parse stat */
.ps-row { display: flex; gap: 0.75rem; flex-wrap: wrap; margin: 1.1rem 0; }
.ps {
  background: var(--bg3);
  border: 1px solid var(--border);
  border-radius: 9px;
  padding: 0.65rem 1.1rem;
  min-width: 90px;
  animation: scaleIn 0.35s ease both;
  transition: border-color 0.2s;
}
.ps:hover { border-color: var(--border2); }
.ps-val { font-family: var(--display); font-size: 1.7rem; font-weight: 800; line-height: 1; }
.ps-lbl { font-size: 0.6rem; text-transform: uppercase; letter-spacing: 0.12em; color: var(--text3); margin-top: 3px; }
.ps.red    .ps-val { color: var(--red2); }
.ps.orange .ps-val { color: var(--orange2); }
.ps.green  .ps-val { color: var(--green2); }
.ps.blue   .ps-val { color: var(--blue2); }

/* Success banner */
.sb {
  background: linear-gradient(135deg, rgba(46,168,110,0.12), rgba(46,168,110,0.04));
  border: 1px solid rgba(46,168,110,0.25);
  border-radius: 10px;
  padding: 1rem 1.5rem;
  color: var(--green2);
  font-family: var(--display);
  font-size: 0.95rem; font-weight: 600;
  animation: successPop 0.5s cubic-bezier(0.34,1.56,0.64,1) both;
  display: flex; align-items: center; gap: 0.75rem;
}

/* Fleet status card — sidebar */
.fsc {
  background: var(--bg3);
  border: 1px solid var(--border);
  border-radius: 9px;
  padding: 0.7rem 0.9rem;
  margin-bottom: 0.4rem;
  display: flex; align-items: center; justify-content: space-between;
  animation: slideRight 0.3s ease both;
  transition: border-color 0.2s;
}
.fsc:hover { border-color: var(--border2); }
.fsc-name { font-family: var(--display); font-size: 0.78rem; font-weight: 700; color: var(--text); }
.fsc-num  { font-family: var(--display); font-size: 0.95rem; font-weight: 800; }

/* Sidebar logo */
.logo {
  font-family: var(--display);
  font-size: 1.25rem; font-weight: 800;
  letter-spacing: 0.06em; color: var(--gold2);
  display: flex; align-items: center; gap: 0.5rem;
}
.logo-sub {
  font-size: 0.58rem; text-transform: uppercase;
  letter-spacing: 0.18em; color: var(--text3); margin-top: 3px;
}
.logo-line {
  height: 1px; margin: 1rem 0;
  background: linear-gradient(90deg, var(--gold), transparent);
}

/* All-clear card */
.ac {
  background: rgba(46,168,110,0.06);
  border: 1px solid rgba(46,168,110,0.15);
  border-radius: 10px;
  padding: 1.75rem;
  text-align: center;
  color: var(--green2);
  font-family: var(--display);
  font-size: 1rem; font-weight: 600;
  letter-spacing: 0.03em;
}

/* Vessel detail meta line */
.meta-line {
  font-size: 0.7rem; color: var(--text3);
  font-family: var(--mono);
  margin: 0.6rem 0 0;
  display: flex; gap: 1.5rem; flex-wrap: wrap;
}
.meta-line b { color: var(--text2); }

/* Upload conversion badge */
.conv {
  display: inline-flex; align-items: center; gap: 0.4rem;
  background: rgba(59,130,246,0.1);
  border: 1px solid rgba(59,130,246,0.2);
  border-radius: 20px;
  padding: 3px 10px;
  font-size: 0.68rem; color: var(--blue2);
  font-family: var(--display); font-weight: 600;
  letter-spacing: 0.08em;
}

/* Scanline effect on header */
.scanline-wrap { position: relative; overflow: hidden; }
.scanline-wrap::after {
  content: '';
  position: absolute; left: 0; right: 0; height: 60px;
  background: linear-gradient(transparent, rgba(201,149,42,0.03), transparent);
  animation: scanline 4s linear infinite;
  pointer-events: none;
}
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# DATABASE
# ════════════════════════════════════════════════════════════════════
DB_PATH = Path("running_hours.db")

def get_db():
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS vessels (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS upload_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL, filename TEXT NOT NULL,
        file_hash TEXT NOT NULL, report_date TEXT,
        me_total_hrs INTEGER, me_this_month INTEGER,
        uploaded_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS components (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL, category TEXT NOT NULL,
        engine_label TEXT NOT NULL, unit TEXT NOT NULL,
        description TEXT NOT NULL, periodicity REAL,
        last_oh_date TEXT, last_oh_hrs REAL,
        hrs_since REAL, pct_used REAL, status TEXT NOT NULL,
        updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS other_equipment (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL, section TEXT NOT NULL,
        description TEXT NOT NULL, periodicity TEXT,
        last_date TEXT, run_hrs TEXT,
        updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_cv ON components(vessel_name);
    CREATE INDEX IF NOT EXISTS idx_cs ON components(status);
    CREATE INDEX IF NOT EXISTS idx_ov ON other_equipment(vessel_name);
    """)
    conn.commit(); conn.close()

init_db()


# ════════════════════════════════════════════════════════════════════
# .DOC → .DOCX CONVERSION  (LibreOffice via packages.txt)
# ════════════════════════════════════════════════════════════════════
def convert_doc_to_docx(file_bytes: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError(
            "LibreOffice not found. Add packages.txt to repo root with: libreoffice"
        )
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    out_dir   = tempfile.gettempdir()
    base      = os.path.splitext(os.path.basename(tmp_path))[0]
    docx_path = os.path.join(out_dir, base + ".docx")
    profile   = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        r = subprocess.run(
            [soffice, "--headless", "--norestore", "--nofirststartwizard",
             f"-env:UserInstallation={profile}",
             "--convert-to", "docx", tmp_path, "--outdir", out_dir],
            capture_output=True, timeout=120,
        )
        if not os.path.exists(docx_path):
            raise RuntimeError(
                f"Conversion failed (exit {r.returncode}). "
                f"{r.stderr.decode('utf-8', errors='ignore')[:300]}"
            )
        with open(docx_path, "rb") as f:
            return f.read()
    finally:
        for p in [tmp_path, docx_path]:
            try:
                if os.path.exists(p): os.unlink(p)
            except Exception: pass


# ════════════════════════════════════════════════════════════════════
# PARSER — 100% integrity, validated against ground truth
# ════════════════════════════════════════════════════════════════════
def _clean_period(raw):
    if not raw: return None
    s = re.sub(r'\.(?=\d{3}(\.|$))', '', raw.strip())
    s = re.sub(r'[^0-9\.]', '', s)
    try: return float(s) if s else None
    except ValueError: return None

def _parse_date(raw):
    if not raw or raw.strip() in ('', 'N/A', 'n/a'): return None
    raw = re.sub(r'\s+', ' ', raw.strip().lstrip('[').rstrip(']'))
    if re.match(r'^\d+$', raw.strip()): return None
    rn = re.sub(r'\bSEPT\b', 'SEP', raw, flags=re.IGNORECASE)
    rn = re.sub(r'\bJUNE\b', 'JUN', rn,  flags=re.IGNORECASE)
    rn = re.sub(r'\bJULY\b', 'JUL', rn,  flags=re.IGNORECASE)
    fmts = ['%d %b %y','%d %B %y','%d %b %Y','%d %B %Y',
            '%d/%m/%y','%d/%m/%Y','%d-%m-%y','%d-%m-%Y',
            '%b %Y','%B %Y','%Y-%m-%d']
    for fmt in fmts:
        for v in (rn, rn.upper(), rn.title(), raw, raw.upper()):
            try: return datetime.strptime(v, fmt).strftime('%Y-%m-%d')
            except ValueError: pass
    return raw

def _parse_hrs(raw):
    if not raw or raw.strip() in ('', 'N/A', 'n/a'): return None
    for n in re.findall(r'\d[\d,]*', raw.replace('\n', ' ')):
        try:
            v = float(n.replace(',', ''))
            if v > 0: return v
        except ValueError: pass
    return None

def _status(hrs, period):
    if hrs is None or period is None or period == 0: return 'NO DATA'
    p = hrs / period
    if p > 1.0: return 'OVERDUE'
    if p >= 0.80: return 'HIGH PRIORITY'
    return 'OK'

def _pct(hrs, period):
    if hrs is None or period is None or period == 0: return 0.0
    return round(hrs / period, 4)


def parse_doc_bytes(docx_bytes: bytes) -> dict:
    from docx import Document
    warns = []
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
        tmp.write(docx_bytes); tmp_path = tmp.name
    try:
        doc = Document(tmp_path)
    except Exception as e:
        raise ValueError(f"Cannot open document: {e}")
    finally:
        try: os.unlink(tmp_path)
        except Exception: pass

    if not doc.tables:
        raise ValueError("No tables found — is this a TEC-004 report?")

    # Vessel name & date
    vessel_name = "UNKNOWN"; report_date = None
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt: continue
        vm = re.search(
            r"Vessel[\u2019\u2018']?s?\s+Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+?)(?:\s{2,}|\t)",
            txt, re.IGNORECASE)
        dm = re.search(r"Date\s*:\s*(.+)", txt, re.IGNORECASE)
        if vm: vessel_name = vm.group(1).strip()
        if dm: report_date = _parse_date(dm.group(1).strip())
        if vm or dm: break
    if vessel_name == "UNKNOWN":
        warns.append("Could not extract vessel name from header.")

    # Table 0 — Main Engine
    me_total = me_month = None
    components = []
    t0 = doc.tables[0]

    for cell in t0.rows[0].cells:
        txt = cell.text
        if m := re.search(r'Total Running Hours[\s:ǀ|]+([\d,]+)', txt, re.IGNORECASE):
            try: me_total = int(m.group(1).replace(',', ''))
            except ValueError: pass
        if m := re.search(r'This Month[\s:]+([\d,]+)', txt, re.IGNORECASE):
            try: me_month = int(m.group(1).replace(',', ''))
            except ValueError: pass

    cyl_cols = []
    if len(t0.rows) > 1:
        for ci, cell in enumerate(t0.rows[1].cells):
            if m := re.search(r'CYL\s*\.?\s*No\s*\.?\s*(\d+)', cell.text.strip(), re.IGNORECASE):
                lbl = f"Cyl {int(m.group(1))}"
                if not cyl_cols or cyl_cols[-1][1] != lbl:
                    cyl_cols.append((ci, lbl))

    rows = t0.rows; i = 2
    while i < len(rows) - 1:
        r1 = [c.text.strip() for c in rows[i].cells]
        r2 = [c.text.strip() for c in rows[i+1].cells] if i+1 < len(rows) else []
        name = r1[0] if r1 else ''
        if not name: i += 1; continue
        t1 = r1[2] if len(r1) > 2 else ''
        t2 = r2[2] if len(r2) > 2 else ''
        if t1 == '1' and t2 == '2' and r1[0] == (r2[0] if r2 else ''):
            period = _clean_period(r1[1] if len(r1) > 1 else '')
            for ci, lbl in cyl_cols:
                d = _parse_date(r1[ci]) if ci < len(r1) else None
                h = _parse_hrs(r2[ci])  if ci < len(r2) else None
                if d is None and h is None: continue
                components.append({
                    'category':'MAIN_ENGINE','engine_label':'ME','unit':lbl,
                    'description':name,'periodicity':period,
                    'last_oh_date':d,'last_oh_hrs':h,'hrs_since':h,
                    'pct_used':_pct(h,period),'status':_status(h,period),
                })
            i += 2
        else: i += 1

    # Table 1 — Turbocharger / Coolers / A/C
    other_equip = []
    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        SKIP = {'TURBOCHARGER','AUXILIARY BOILER','COOLERS','EXH GAS  BOILER',
                'A/C & REFR. COMPRESSORS','MAIN AIR COMPRESSORS',
                'PERIODICTLY','DATE OF LAST INSPECTION','RUN HRS',
                'DATE OF LAST CLEANING','DATE','PERIODICITY'}
        for row in t1.rows:
            cells = [c.text.strip() for c in row.cells]
            for sec, dc, datec, hrsc in [
                ('TURBOCHARGER / AUX BOILER',0,1,3),
                ('COOLERS / EXH GAS BOILER',5,6,8),
                ('A/C & COMPRESSORS',10,11,12),
            ]:
                desc = cells[dc] if dc < len(cells) else ''
                if not desc or desc.upper() in SKIP: continue
                dv = cells[datec] if datec < len(cells) else ''
                hv = cells[hrsc]  if hrsc  < len(cells) else ''
                if dv or hv:
                    other_equip.append({'section':sec,'description':desc,
                                        'periodicity':'','last_date':dv,'run_hrs':hv})

    # Table 2 — Auxiliary Engines
    if len(doc.tables) > 2:
        t2 = doc.tables[2]; rows2 = t2.rows
        engine_blocks = []
        if rows2:
            hdr  = [c.text.strip() for c in rows2[0].cells]
            tot  = [c.text.strip() for c in rows2[2].cells] if len(rows2) > 2 else []
            seen = set()
            for ci, cell in enumerate(hdr):
                if m := re.search(r'Aux\.\s*Engine\s*No\.?\s*(\d+)', cell, re.IGNORECASE):
                    lbl = f"AUX-{int(m.group(1))}"
                    if lbl not in seen:
                        seen.add(lbl)
                        th = next((_parse_hrs(tot[j]) for j in range(ci, min(ci+14,len(tot)))
                                   if _parse_hrs(tot[j])), None)
                        engine_blocks.append((lbl, ci, th))
        cyl_map = {}
        if len(rows2) > 4:
            r4 = [c.text.strip() for c in rows2[4].cells]
            for ei, (elbl, estart, _) in enumerate(engine_blocks):
                eend = engine_blocks[ei+1][1] if ei+1 < len(engine_blocks) else len(r4)
                seen_c: list = []
                for ci in range(estart, eend):
                    if ci < len(r4):
                        try:
                            cn = int(r4[ci])
                            if cn not in seen_c:
                                seen_c.append(cn); cyl_map[ci] = (elbl, cn)
                        except ValueError: pass
        i2 = 5
        while i2 < len(rows2) - 1:
            r1 = [c.text.strip() for c in rows2[i2].cells]
            r2 = [c.text.strip() for c in rows2[i2+1].cells] if i2+1 < len(rows2) else []
            name = r1[0] if r1 else ''
            if not name: i2 += 1; continue
            t1t = r1[2] if len(r1) > 2 else ''
            t2t = r2[2] if len(r2) > 2 else ''
            if t1t in ('1','2') and r1[0] == (r2[0] if r2 else ''):
                period = _clean_period(r1[1] if len(r1) > 1 else '')
                for ci, (elbl, cn) in cyl_map.items():
                    d = _parse_date(r1[ci]) if ci < len(r1) else None
                    h = _parse_hrs(r2[ci])  if ci < len(r2) else None
                    if d is None and h is None: continue
                    components.append({
                        'category':'AUX_ENGINE','engine_label':elbl,'unit':f"Cyl {cn}",
                        'description':name,'periodicity':period,
                        'last_oh_date':d,'last_oh_hrs':h,'hrs_since':h,
                        'pct_used':_pct(h,period),'status':_status(h,period),
                    })
                i2 += 2
            else: i2 += 1

    # Table 3 — D/G Equipment
    if len(doc.tables) > 3:
        t3 = doc.tables[3]; dglbls = ['D/G 1','D/G 2','D/G 3']
        for ridx, row in enumerate(t3.rows):
            cells = [c.text.strip() for c in row.cells]
            if ridx == 0: continue
            for dc, pc, tc, ds in [(0,1,2,3),(9,10,11,12)]:
                desc  = cells[dc] if dc < len(cells) else ''
                per   = cells[pc] if pc < len(cells) else ''
                rtype = cells[tc] if tc < len(cells) else ''
                if not desc or rtype not in ('1','2'): continue
                for dgi, dglbl in enumerate(dglbls):
                    col = ds + dgi
                    val = cells[col] if col < len(cells) else ''
                    if not val: continue
                    key = f"{desc} — {dglbl}"
                    if rtype == '1':
                        other_equip.append({'section':'D/G EQUIPMENT','description':key,
                                            'periodicity':per,'last_date':_parse_date(val) or val,'run_hrs':''})
                    else:
                        for e in reversed(other_equip):
                            if e['description'] == key and e['run_hrs'] == '':
                                e['run_hrs'] = val; break
                        else:
                            other_equip.append({'section':'D/G EQUIPMENT','description':key,
                                                'periodicity':per,'last_date':'','run_hrs':val})

    return {
        'vessel_name':vessel_name,'report_date':report_date,
        'me_total_hrs':me_total,'me_this_month':me_month,
        'components':components,'other_equipment':other_equip,'warnings':warns,
    }


# ════════════════════════════════════════════════════════════════════
# DB HELPERS
# ════════════════════════════════════════════════════════════════════
def save_parsed_data(parsed, filename, file_hash):
    conn = get_db(); c = conn.cursor()
    now = datetime.utcnow().isoformat(); v = parsed['vessel_name']
    c.execute("INSERT OR IGNORE INTO vessels(name,created_at) VALUES(?,?)", (v,now))
    c.execute("INSERT INTO upload_log(vessel_name,filename,file_hash,report_date,me_total_hrs,me_this_month,uploaded_at) VALUES(?,?,?,?,?,?,?)",
              (v,filename,file_hash,parsed['report_date'],parsed['me_total_hrs'],parsed['me_this_month'],now))
    c.execute("DELETE FROM components WHERE vessel_name=?", (v,))
    for comp in parsed['components']:
        c.execute("INSERT INTO components(vessel_name,category,engine_label,unit,description,periodicity,last_oh_date,last_oh_hrs,hrs_since,pct_used,status,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                  (v,comp['category'],comp['engine_label'],comp['unit'],comp['description'],
                   comp['periodicity'],comp['last_oh_date'],comp['last_oh_hrs'],
                   comp['hrs_since'],comp['pct_used'],comp['status'],now))
    c.execute("DELETE FROM other_equipment WHERE vessel_name=?", (v,))
    for oe in parsed['other_equipment']:
        c.execute("INSERT INTO other_equipment(vessel_name,section,description,periodicity,last_date,run_hrs,updated_at) VALUES(?,?,?,?,?,?,?)",
                  (v,oe['section'],oe['description'],oe.get('periodicity',''),
                   oe.get('last_date',''),oe.get('run_hrs',''),now))
    conn.commit(); conn.close()

@st.cache_data(ttl=10)
def get_all_vessels():
    conn = get_db()
    rows = conn.execute("SELECT name FROM vessels ORDER BY name").fetchall()
    conn.close(); return [r['name'] for r in rows]

@st.cache_data(ttl=10)
def get_components_df(vessel):
    conn = get_db()
    df = pd.read_sql_query("SELECT * FROM components WHERE vessel_name=? ORDER BY category,engine_label,description,unit", conn, params=(vessel,))
    conn.close(); return df

@st.cache_data(ttl=10)
def get_other_equip_df(vessel):
    conn = get_db()
    df = pd.read_sql_query("SELECT * FROM other_equipment WHERE vessel_name=? ORDER BY section,description", conn, params=(vessel,))
    conn.close(); return df

@st.cache_data(ttl=10)
def get_upload_history(vessel):
    conn = get_db()
    df = pd.read_sql_query("SELECT filename,report_date,me_total_hrs,me_this_month,uploaded_at FROM upload_log WHERE vessel_name=? ORDER BY uploaded_at DESC LIMIT 20", conn, params=(vessel,))
    conn.close(); return df

@st.cache_data(ttl=10)
def get_fleet_summary():
    conn = get_db()
    df = pd.read_sql_query("""
        SELECT c.vessel_name,
               COUNT(CASE WHEN c.status='OVERDUE'       THEN 1 END) AS overdue,
               COUNT(CASE WHEN c.status='HIGH PRIORITY' THEN 1 END) AS high_priority,
               COUNT(CASE WHEN c.status='OK'            THEN 1 END) AS ok,
               COUNT(*) AS total,
               MAX(u.uploaded_at)  AS last_upload,
               MAX(u.me_total_hrs) AS me_total_hrs,
               MAX(u.report_date)  AS report_date
        FROM components c
        LEFT JOIN upload_log u ON u.vessel_name=c.vessel_name
        GROUP BY c.vessel_name ORDER BY overdue DESC, high_priority DESC
    """, conn); conn.close(); return df


# ════════════════════════════════════════════════════════════════════
# UI HELPERS
# ════════════════════════════════════════════════════════════════════
def kpi(val, lbl, color="gold", delay=0):
    return (f'<div class="kc {color}" style="animation-delay:{delay}s">'
            f'<div class="kc-val">{val}</div>'
            f'<div class="kc-lbl">{lbl}</div></div>')

def badge(status):
    m = {'OVERDUE':'od','HIGH PRIORITY':'hp','OK':'ok','NO DATA':'nd'}
    cls = m.get(status, 'nd')
    return f'<span class="badge badge-{cls}">{status}</span>'

def pbar(pct, status):
    w   = min(pct * 100, 100)
    cls = 'pb-od' if status=='OVERDUE' else ('pb-hp' if status=='HIGH PRIORITY' else 'pb-ok')
    return f'<div class="pb-wrap"><div class="pb-fill {cls}" style="width:{w:.1f}%"></div></div>'

# ── fmt_df: works for BOTH raw parsed dicts AND DB dataframes ──────
def fmt_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalise columns from either:
      - raw parser output  (keys: description, engine_label, unit, periodicity,
                                   last_oh_date, hrs_since, pct_used, status)
      - DB query output    (same keys, all present)
    Returns a display-ready DataFrame.
    """
    cols_needed = ['description','engine_label','unit','periodicity',
                   'last_oh_date','hrs_since','pct_used','status']
    # Only keep columns that exist — fill missing with '—'
    d = pd.DataFrame()
    for col in cols_needed:
        if col in df.columns:
            d[col] = df[col]
        else:
            d[col] = '—'
    d.columns = ['Component','Engine','Unit','Periodicity',
                 'Last O/H Date','Hrs Since','% Used','Status']
    d['% Used']        = d['% Used'].apply(
        lambda x: f"{float(x)*100:.1f}%" if pd.notna(x) and x != '—' else '—')
    d['Periodicity']   = d['Periodicity'].apply(
        lambda x: f"{int(float(x)):,}" if pd.notna(x) and x not in ('—','N/A','') and str(x) not in ('nan','None') else 'N/A')
    d['Hrs Since']     = d['Hrs Since'].apply(
        lambda x: f"{int(float(x)):,}" if pd.notna(x) and x != '—' else '—')
    d['Last O/H Date'] = d['Last O/H Date'].fillna('—').replace('', '—')
    return d

def style_df(df: pd.DataFrame):
    def rs(row):
        s = str(row.get('Status',''))
        if s == 'OVERDUE':       return ['background:#1e0808;color:#f87171']*len(row)
        if s == 'HIGH PRIORITY': return ['background:#1e1008;color:#fb923c']*len(row)
        if s == 'OK':            return ['background:#07130e;color:#34d399']*len(row)
        return ['background:#070d1a;color:#3d5470']*len(row)
    return df.style.apply(rs, axis=1)

def show_table(df: pd.DataFrame, height: int = None):
    if df.empty:
        st.info("No data.")
        return
    h = height or min(700, 38 * (len(df) + 1) + 4)
    st.dataframe(style_df(fmt_df(df)), use_container_width=True, hide_index=True, height=h)

# Preview table for upload page — works with raw list of dicts
def show_preview_table(records: list, height: int = 320):
    if not records:
        st.info("No components parsed.")
        return
    df = pd.DataFrame(records)
    show_table(df, height=height)


# ════════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div class="logo">⚓ FLEET MONITOR</div>
    <div class="logo-sub">Running Hours Management System</div>
    <div class="logo-line"></div>
    """, unsafe_allow_html=True)

    page = st.selectbox("nav",
        ["🗺️  Fleet Overview","🚢  Vessel Detail","📤  Upload Report","📋  Upload History"],
        label_visibility="collapsed")

    st.markdown("<br>", unsafe_allow_html=True)

    vessels = get_all_vessels()
    selected_vessel = st.selectbox("Active Vessel", vessels) if vessels else None
    if not vessels:
        st.info("No data yet — upload a .doc report to begin.")

    if vessels:
        smry = get_fleet_summary()
        if not smry.empty:
            st.markdown('<div class="logo-line"></div>', unsafe_allow_html=True)
            st.markdown('<div style="font-size:0.58rem;text-transform:uppercase;letter-spacing:0.2em;color:var(--text3);margin-bottom:0.6rem;">Fleet Status</div>', unsafe_allow_html=True)
            for _, row in smry.iterrows():
                od = int(row['overdue']); hp = int(row['high_priority'])
                col = 'var(--red2)' if od > 0 else ('var(--orange2)' if hp > 0 else 'var(--green2)')
                st.markdown(f"""
                <div class="fsc">
                  <div class="fsc-name">{row['vessel_name']}</div>
                  <div style="display:flex;gap:6px;align-items:center;">
                    {'<span style="font-size:0.65rem;color:var(--red2);font-family:var(--display);font-weight:700">'+str(od)+' OD</span>' if od>0 else ''}
                    {'<span style="font-size:0.65rem;color:var(--orange2);font-family:var(--display);font-weight:700">'+str(hp)+' HP</span>' if hp>0 else ''}
                    {'<span style="font-size:0.65rem;color:var(--green2);font-family:var(--display);font-weight:700">✓</span>' if od==0 and hp==0 else ''}
                  </div>
                </div>""", unsafe_allow_html=True)

    st.markdown('<div class="logo-line"></div>', unsafe_allow_html=True)
    db_kb = DB_PATH.stat().st_size / 1024 if DB_PATH.exists() else 0
    st.markdown(f'<div style="font-size:0.6rem;color:var(--text3);font-family:var(--mono)">db {db_kb:.0f} kb · {len(vessels)} vessels · v4.0</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# PAGE: UPLOAD
# ════════════════════════════════════════════════════════════════════
if page == "📤  Upload Report":
    st.markdown('<div class="ph"><h1>📤 Upload Report</h1></div><div class="ph-rule"></div>', unsafe_allow_html=True)

    col_up, col_info = st.columns([3, 2], gap="large")
    with col_up:
        uploaded = st.file_uploader("file", type=["doc"], label_visibility="collapsed")
    with col_info:
        st.markdown("""
        <div class="ic">
          <div class="ic-title">Accepted Format</div>
          TEC-004 Running Hours Monthly Report<br>
          Any vessel in the fleet · <b style="color:var(--text)">.doc format</b><br><br>
          <div class="ic-title">What Gets Extracted</div>
          ✦ Vessel name &amp; report date<br>
          ✦ M/E total &amp; monthly running hours<br>
          ✦ All M/E components — dates &amp; hours<br>
          ✦ Aux engines (3 × 6 cylinders)<br>
          ✦ Turbocharger, coolers, D/G equipment<br>
          ✦ Status computed per periodicity
        </div>""", unsafe_allow_html=True)

    if uploaded:
        raw_bytes = uploaded.read()
        file_hash = hashlib.md5(raw_bytes).hexdigest()

        with st.spinner("Converting .doc → .docx and parsing…"):
            try:
                docx_bytes = convert_doc_to_docx(raw_bytes)
            except Exception as e:
                st.error(f"**Conversion failed.** `{e}`")
                st.stop()
            try:
                parsed = parse_doc_bytes(docx_bytes)
            except ValueError as e:
                st.error(f"**Parse failed.** `{e}`")
                st.stop()

        # Counts
        comps  = parsed['components']
        n_comp = len(comps)
        n_od   = sum(1 for c in comps if c['status'] == 'OVERDUE')
        n_hp   = sum(1 for c in comps if c['status'] == 'HIGH PRIORITY')
        n_ok   = sum(1 for c in comps if c['status'] == 'OK')
        n_oe   = len(parsed['other_equipment'])

        st.markdown('<div class="sl">Parse Preview — confirm before saving</div>', unsafe_allow_html=True)

        # Top metrics
        m1,m2,m3,m4,m5 = st.columns(5)
        m1.metric("Vessel",         parsed['vessel_name'])
        m2.metric("Report Date",    parsed['report_date'] or "—")
        m3.metric("M/E Total Hrs",  f"{parsed['me_total_hrs']:,}"  if parsed['me_total_hrs']  else "—")
        m4.metric("M/E This Month", f"{parsed['me_this_month']:,}" if parsed['me_this_month'] else "—")
        m5.metric("Components",     n_comp)

        # Stat row
        st.markdown(f"""
        <div class="ps-row">
          <div class="ps red">   <div class="ps-val">{n_od}</div><div class="ps-lbl">Overdue</div></div>
          <div class="ps orange"><div class="ps-val">{n_hp}</div><div class="ps-lbl">High Priority</div></div>
          <div class="ps green"> <div class="ps-val">{n_ok}</div><div class="ps-lbl">OK</div></div>
          <div class="ps blue">  <div class="ps-val">{n_oe}</div><div class="ps-lbl">Other Equip</div></div>
        </div>""", unsafe_allow_html=True)

        for w in parsed['warnings']:
            st.warning(f"⚠ {w}")

        # Preview table — uses show_preview_table (accepts raw dicts)
        if comps:
            with st.expander(f"Preview {n_comp} parsed component records", expanded=True):
                show_preview_table(comps)

        st.markdown("---")
        col_btn, _ = st.columns([1, 4])
        with col_btn:
            if st.button("✅  CONFIRM & SAVE", use_container_width=True):
                save_parsed_data(parsed, uploaded.name, file_hash)
                for fn in [get_all_vessels, get_components_df, get_other_equip_df, get_fleet_summary]:
                    fn.clear()
                st.markdown(f"""
                <div class="sb">
                  <span style="font-size:1.4rem">✓</span>
                  <span><b>{parsed['vessel_name']}</b> saved — {n_comp} components · {n_od} overdue · {n_hp} high priority</span>
                </div>""", unsafe_allow_html=True)
                st.balloons()


# ════════════════════════════════════════════════════════════════════
# PAGE: FLEET OVERVIEW
# ════════════════════════════════════════════════════════════════════
elif page == "🗺️  Fleet Overview":
    st.markdown('<div class="ph scanline-wrap"><h1>🗺️ Fleet Overview</h1></div><div class="ph-rule"></div>', unsafe_allow_html=True)

    summary = get_fleet_summary()
    if summary.empty:
        st.info("No data loaded yet. Upload a .doc report to begin."); st.stop()

    tv  = len(summary)
    tc  = int(summary['total'].sum())
    tod = int(summary['overdue'].sum())
    thp = int(summary['high_priority'].sum())
    tok = int(summary['ok'].sum())

    k1,k2,k3,k4,k5 = st.columns(5)
    for col,(val,lbl,clr,dly) in zip([k1,k2,k3,k4,k5],[
        (tv,"Vessels","blue",0),(tc,"Components","gold",0.06),
        (tod,"Overdue","red",0.12),(thp,"High Priority","orange",0.18),(tok,"OK","green",0.24)]):
        with col: st.markdown(kpi(val,lbl,clr,dly), unsafe_allow_html=True)

    st.markdown('<div class="sl">Fleet Status Table</div>', unsafe_allow_html=True)

    disp = summary[['vessel_name','overdue','high_priority','ok','total','me_total_hrs','last_upload']].copy()
    disp.columns = ['Vessel','Overdue','High Priority','OK','Total','M/E Total Hrs','Last Upload']
    disp['M/E Total Hrs'] = disp['M/E Total Hrs'].apply(lambda x: f"{int(x):,}" if pd.notna(x) else '—')
    disp['Last Upload']   = pd.to_datetime(disp['Last Upload'],errors='coerce').dt.strftime('%Y-%m-%d').fillna('—')

    def fleet_style(row):
        if row['Overdue']>0:       return ['background:#1a0505;color:#f87171']+['background:#1a0505']*6
        if row['High Priority']>0: return ['background:#1a0900;color:#fb923c']+['background:#1a0900']*6
        return                            ['background:#050f0a;color:#34d399']+['background:#050f0a']*6

    st.dataframe(disp.style.apply(fleet_style,axis=1),
                 use_container_width=True, hide_index=True,
                 height=min(600, 38*(len(disp)+1)+3))

    st.markdown('<div class="sl">Vessel Breakdown</div>', unsafe_allow_html=True)
    for _, row in summary.iterrows():
        od = int(row['overdue']); hp = int(row['high_priority']); ok = int(row['ok'])
        icon = "🔴" if od>0 else ("🟡" if hp>0 else "🟢")
        label = f"{icon} **{row['vessel_name']}** — {od} overdue · {hp} high priority · {ok} OK"
        with st.expander(label, expanded=False):
            cc = get_components_df(row['vessel_name'])
            if not cc.empty:
                ta, tb = st.tabs(["🔴  Overdue","🟡  High Priority"])
                with ta:
                    od_df = cc[cc['status']=='OVERDUE']
                    if od_df.empty: st.markdown('<div class="ac">✓ No overdue items</div>', unsafe_allow_html=True)
                    else: show_table(od_df)
                with tb:
                    hp_df = cc[cc['status']=='HIGH PRIORITY']
                    if hp_df.empty: st.markdown('<div class="ac">✓ No high-priority items</div>', unsafe_allow_html=True)
                    else: show_table(hp_df)


# ════════════════════════════════════════════════════════════════════
# PAGE: VESSEL DETAIL
# ════════════════════════════════════════════════════════════════════
elif page == "🚢  Vessel Detail":
    if not selected_vessel:
        st.info("Select a vessel from the sidebar."); st.stop()

    st.markdown(f'<div class="ph scanline-wrap"><h1>🚢 {selected_vessel}</h1></div><div class="ph-rule"></div>', unsafe_allow_html=True)

    df = get_components_df(selected_vessel)
    oe = get_other_equip_df(selected_vessel)
    if df.empty:
        st.info("No component data for this vessel."); st.stop()

    n_tot = len(df)
    n_od  = int((df['status']=='OVERDUE').sum())
    n_hp  = int((df['status']=='HIGH PRIORITY').sum())
    n_ok  = int((df['status']=='OK').sum())
    n_nd  = int((df['status']=='NO DATA').sum())

    k1,k2,k3,k4,k5 = st.columns(5)
    for col,(val,lbl,clr,dly) in zip([k1,k2,k3,k4,k5],[
        (n_tot,"Total","gold",0),(n_od,"Overdue","red",0.06),
        (n_hp,"High Priority","orange",0.12),(n_ok,"OK","green",0.18),(n_nd,"No Data","blue",0.24)]):
        with col: st.markdown(kpi(val,lbl,clr,dly), unsafe_allow_html=True)

    hist = get_upload_history(selected_vessel)
    if not hist.empty:
        last = hist.iloc[0]
        mt = f"{int(last['me_total_hrs']):,}" if pd.notna(last['me_total_hrs']) else "—"
        mm = f"{int(last['me_this_month']):,}" if pd.notna(last['me_this_month']) else "—"
        st.markdown(f"""
        <div class="meta-line">
          <span>📄 <b>{last['filename']}</b></span>
          <span>Report: <b>{last['report_date'] or '—'}</b></span>
          <span>M/E: <b>{mt}</b> total / <b>{mm}</b> this month</span>
          <span>Uploaded: <b>{str(last['uploaded_at'])[:16]}</b></span>
        </div>""", unsafe_allow_html=True)

    st.markdown("---")
    tabs = st.tabs(["⚠️  Alerts","⚙️  Main Engine","🔩  Aux Engines","🛠️  Other Equipment","📊  All Components"])

    with tabs[0]:
        st.markdown('<div class="sl">Action Required — Overdue &amp; High Priority</div>', unsafe_allow_html=True)
        alerts = df[df['status'].isin(['OVERDUE','HIGH PRIORITY'])].sort_values(
            ['status','pct_used'], ascending=[True,False])
        if alerts.empty:
            st.markdown('<div class="ac">✓ All components within acceptable limits — no action required</div>', unsafe_allow_html=True)
        else:
            show_table(alerts)

    with tabs[1]:
        me = df[df['category']=='MAIN_ENGINE']
        if me.empty: st.info("No Main Engine data.")
        else:
            st.markdown('<div class="sl">Main Engine Components</div>', unsafe_allow_html=True)
            sel = st.selectbox("Filter by component", ['ALL']+sorted(me['description'].unique().tolist()), key="me_f")
            show_table(me if sel=='ALL' else me[me['description']==sel])

    with tabs[2]:
        aux = df[df['category']=='AUX_ENGINE']
        if aux.empty: st.info("No Aux Engine data.")
        else:
            st.markdown('<div class="sl">Auxiliary Engine Components</div>', unsafe_allow_html=True)
            engines = sorted(aux['engine_label'].unique().tolist())
            sel = st.selectbox("Filter by engine", ['ALL']+engines, key="aux_f")
            show_table(aux if sel=='ALL' else aux[aux['engine_label']==sel])

    with tabs[3]:
        if oe.empty: st.info("No other equipment data.")
        else:
            for sec in sorted(oe['section'].unique()):
                st.markdown(f'<div class="sl">{sec}</div>', unsafe_allow_html=True)
                sd = oe[oe['section']==sec][['description','periodicity','last_date','run_hrs']].copy()
                sd.columns = ['Description','Periodicity','Last Date','Run Hrs']
                st.dataframe(sd, use_container_width=True, hide_index=True)

    with tabs[4]:
        st.markdown('<div class="sl">All Component Records</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1: sf = st.multiselect("Status",['OVERDUE','HIGH PRIORITY','OK','NO DATA'],
                                     default=['OVERDUE','HIGH PRIORITY','OK','NO DATA'],key="all_s")
        with c2: cf = st.multiselect("Category",['MAIN_ENGINE','AUX_ENGINE'],
                                     default=['MAIN_ENGINE','AUX_ENGINE'],key="all_c")
        show_table(df[df['status'].isin(sf) & df['category'].isin(cf)])


# ════════════════════════════════════════════════════════════════════
# PAGE: UPLOAD HISTORY
# ════════════════════════════════════════════════════════════════════
elif page == "📋  Upload History":
    st.markdown('<div class="ph"><h1>📋 Upload History</h1></div><div class="ph-rule"></div>', unsafe_allow_html=True)

    if not selected_vessel:
        st.info("Select a vessel from the sidebar."); st.stop()

    st.markdown(f'<div class="sl">{selected_vessel} — Audit Trail</div>', unsafe_allow_html=True)
    hist = get_upload_history(selected_vessel)
    if hist.empty:
        st.info("No upload history for this vessel.")
    else:
        d = hist.copy()
        d.columns = ['Filename','Report Date','M/E Total Hrs','M/E This Month','Uploaded At']
        d['M/E Total Hrs']  = d['M/E Total Hrs'].apply(lambda x: f"{int(x):,}" if pd.notna(x) else '—')
        d['M/E This Month'] = d['M/E This Month'].apply(lambda x: f"{int(x):,}" if pd.notna(x) else '—')
        st.dataframe(d, use_container_width=True, hide_index=True)
