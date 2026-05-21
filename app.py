"""
╔══════════════════════════════════════════════════════════════════════╗
║   FLEET RUNNING HOURS MONITORING SYSTEM  v5.2                        ║
║   100% data integrity · .doc native · Streamlit Cloud production     ║
╚══════════════════════════════════════════════════════════════════════╝
"""

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

# ═══════════════════════════════════════════════════════════════════
#  GLOBAL UI STEALTH DESIGN SYSTEM
# ═══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
  --bg:      #03060d;  --bg1: #06091a;  --bg2: #080e20;
  --bg3:     #0b1228;  --bg4: #0f1830;
  --b1:      #0f1c35;  --b2:  #182840;  --b3:  #223350;
  --gold:    #c89a14;  --gold2:#e0b422; --gold3:#f5cc44;
  --red:     #cc2828;  --red2: #ff5c5c; --red3: #ff8a8a;
  --orange:  #b85518;  --ora2: #ff8833; --ora3: #ffb366;
  --green:   #0d8a4a;  --grn2: #22c55e; --grn3: #6ee7b7;
  --blue:    #1444a8;  --blu2: #3b82f6; --blu3: #93c5fd;
  --t0: #f2f7ff; --t1: #c0d0e8; --t2: #6a84a8; --t3: #304060;
  --ff: 'Space Grotesk', sans-serif;
  --fi: 'Inter', sans-serif;
  --fm: 'JetBrains Mono', monospace;
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

/* Ambient space background glow */
.main::before {
  content: ''; position: fixed; inset: 0; pointer-events: none; z-index: 0;
  background:
    radial-gradient(ellipse 90% 50% at -10% -5%,  rgba(200,154,20,0.06) 0%, transparent 55%),
    radial-gradient(ellipse 70% 45% at 110% 105%,  rgba(20,68,168,0.05) 0%, transparent 55%);
}

/* Sidebar structure */
[data-testid="stSidebar"] {
  background: var(--bg1) !important;
  border-right: 1px solid var(--b2) !important;
}
[data-testid="stSidebar"] * { color: var(--t1) !important; }
[data-testid="stSidebarContent"] { padding: 1.5rem 1.25rem !important; }
[data-testid="stSidebar"] .stSelectbox > div > div {
  background: var(--bg3) !important; border: 1px solid var(--b2) !important;
  border-radius: 6px !important;
}

/* Operational headings */
h1 { font-family: var(--ff) !important; font-size: 1.8rem !important;
     font-weight: 700 !important; color: var(--t0) !important;
     letter-spacing: -0.02em !important; line-height: 1.2 !important; }
h2 { font-family: var(--ff) !important; font-size: 1.2rem !important;
     font-weight: 600 !important; color: var(--t0) !important; }
h3 { font-family: var(--ff) !important; font-size: 1rem !important;
     font-weight: 500 !important; color: var(--t1) !important; }

/* Unified dashboard metric blocks */
[data-testid="stMetric"] {
  background: var(--bg3) !important; border: 1px solid var(--b2) !important;
  border-radius: 10px !important; padding: 1rem 1.2rem 1.1rem !important;
  position: relative !important; overflow: hidden !important;
  transition: border-color .25s, transform .2s !important;
}
[data-testid="stMetric"]:hover {
  border-color: var(--b3) !important; transform: translateY(-2px) !important;
}
[data-testid="stMetricValue"] {
  font-family: var(--ff) !important; font-size: 2rem !important;
  font-weight: 700 !important; color: var(--t0) !important;
  letter-spacing: -0.03em !important;
}
[data-testid="stMetricLabel"] {
  font-family: var(--fi) !important; color: var(--t3) !important;
  font-size: 0.62rem !important; text-transform: uppercase !important;
  letter-spacing: 0.15em !important;
}

/* Enhanced telemetry layout grids */
[data-testid="stDataFrame"] {
  border: 1px solid var(--b2) !important;
  border-radius: 10px !important; overflow: hidden !important;
  box-shadow: 0 4px 24px rgba(0,0,0,0.35) !important;
}
.dvn-scroller { background: var(--bg2) !important; }

/* Control buttons styling */
.stButton > button {
  background: linear-gradient(135deg, var(--gold) 0%, #8a6a08 100%) !important;
  color: #000 !important; border: none !important;
  font-family: var(--ff) !important; font-weight: 600 !important;
  font-size: 0.82rem !important; letter-spacing: 0.06em !important;
  text-transform: uppercase !important; border-radius: 7px !important;
  padding: .6rem 1.8rem !important;
  box-shadow: 0 2px 14px rgba(200,154,20,.2),inset 0 1px 0 rgba(255,255,255,.1) !important;
  transition: all .18s !important;
}
.stButton > button:hover {
  background: linear-gradient(135deg, var(--gold2) 0%, var(--gold) 100%) !important;
  box-shadow: 0 5px 22px rgba(200,154,20,.38) !important;
  transform: translateY(-2px) !important;
}
.stButton > button:active { transform: translateY(0) !important; }

/* File drop interface layout */
[data-testid="stFileUploadDropzone"] {
  background: linear-gradient(160deg,rgba(200,154,20,.04) 0%,rgba(20,68,168,.03) 100%) !important;
  border: 1.5px dashed var(--gold) !important; border-radius: 14px !important;
  transition: all .3s !important; padding: 3rem 2rem !important;
}
[data-testid="stFileUploadDropzone"]:hover {
  background: rgba(200,154,20,.07) !important;
  border-color: var(--gold2) !important;
  box-shadow: 0 0 40px rgba(200,154,20,.07) !important;
}
[data-testid="stFileUploadDropzone"] p,
[data-testid="stFileUploadDropzone"] span {
  color: var(--gold2) !important; font-family: var(--ff) !important;
  font-size: .95rem !important; font-weight: 500 !important;
}
[data-testid="stFileUploadDropzone"] small { color: var(--t2) !important; }

/* Telemetry display navigation tabs */
.stTabs [data-baseweb="tab-list"] {
  background: var(--bg2) !important; border-radius: 10px 10px 0 0 !important;
  border-bottom: 1px solid var(--b2) !important; gap: 0 !important; padding: 0 1rem !important;
}
.stTabs [data-baseweb="tab"] {
  background: transparent !important; color: var(--t3) !important;
  font-family: var(--ff) !important; font-weight: 500 !important;
  letter-spacing: .04em !important; font-size: .75rem !important;
  text-transform: uppercase !important;
  padding: .85rem 1.3rem !important;
  border-bottom: 2px solid transparent !important;
  margin-bottom: -1px !important; transition: color .2s !important;
}
.stTabs [data-baseweb="tab"]:hover { color: var(--t2) !important; }
.stTabs [aria-selected="true"] {
  color: var(--gold2) !important; border-bottom: 2px solid var(--gold) !important;
}
.stTabs [data-baseweb="tab-panel"] {
  background: var(--bg2) !important; border: 1px solid var(--b2) !important;
  border-top: none !important; border-radius: 0 0 10px 10px !important;
  padding: 1.5rem !important;
}

/* Context filters selectors */
.stSelectbox > div > div, .stMultiSelect > div > div {
  background: var(--bg3) !important; border: 1px solid var(--b2) !important;
  border-radius: 7px !important; color: var(--t1) !important;
  font-family: var(--fi) !important;
}
.stSelectbox label, .stMultiSelect label {
  font-family: var(--fi) !important; color: var(--t3) !important;
  font-size: .7rem !important; text-transform: uppercase !important;
  letter-spacing: .1em !important;
}

/* Interface components modifications */
.stAlert { border-radius: 8px !important; border-left-width: 3px !important; }
hr { border-color: var(--b2) !important; opacity:1 !important; margin:1.5rem 0 !important; }
a { color: var(--gold2) !important; text-decoration: none !important; }
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: var(--bg1); }
::-webkit-scrollbar-thumb { background: var(--b3); border-radius: 3px; }

/* CSS View Presentation Keyframes */
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
  font-family: var(--fi); font-size: .6rem; font-weight: 500;
  letter-spacing: .22em; text-transform: uppercase; color: var(--gold);
  margin-bottom: .3rem;
  animation: slideDown .4s .05s ease both; animation-fill-mode: both;
}
.ph-line {
  height: 1px; margin: .35rem 0 1.75rem;
  background: linear-gradient(90deg, var(--gold),  transparent 70%);
  animation: goldLine .7s .1s ease both; animation-fill-mode: both;
}

/* Premium KPI Interface layouts */
.kc {
  background: var(--bg3); border: 1px solid var(--b2); border-radius: 10px;
  padding: 1rem 1.2rem 1.1rem; position: relative; overflow: hidden;
  animation: slideUp .4s ease both; animation-fill-mode: both;
  transition: border-color .25s, transform .2s, box-shadow .25s; cursor: default;
}
.kc:hover { transform: translateY(-4px); box-shadow: 0 14px 40px rgba(0,0,0,.5); }
.kc::before {
  content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px;
  border-radius: 10px 10px 0 0;
}
.kc.gold   { border-color: rgba(200,154,20,.3); }
.kc.gold::before   { background: linear-gradient(90deg, var(--gold),  transparent 70%); }
.kc.gold::after    { content:''; position:absolute; inset:0; border-radius:10px; background:linear-gradient(160deg,rgba(200,154,20,.05) 0%,transparent 60%); pointer-events:none; }
.kc.gold:hover     { border-color: rgba(224,180,34,.5); box-shadow: 0 14px 40px rgba(0,0,0,.5),0 0 24px rgba(200,154,20,.07); }
kc.red    { border-color: rgba(204,40,40,.25); }
.kc.red::before    { background: linear-gradient(90deg, var(--red),   transparent 70%); }
.kc.red::after     { content:''; position:absolute; inset:0; border-radius:10px; background:linear-gradient(160deg,rgba(204,40,40,.05) 0%,transparent 60%); pointer-events:none; }
.kc.red:hover      { border-color: rgba(255,92,92,.4); box-shadow:  0 14px 40px rgba(0,0,0,.5),0 0 24px rgba(204,40,40,.07); }
.kc.orange { border-color: rgba(184,85,24,.25); }
.kc.orange::before { background: linear-gradient(90deg, var(--orange),transparent 70%); }
.kc.orange::after  { content:''; position:absolute; inset:0; border-radius:10px; background:linear-gradient(160deg,rgba(184,85,24,.05) 0%,transparent 60%); pointer-events:none; }
.kc.orange:hover   { border-color: rgba(255,136,51,.4); box-shadow: 0 14px 40px rgba(0,0,0,.5),0 0 24px rgba(184,85,24,.07); }
.kc.green  { border-color: rgba(13,138,74,.25); }
.kc.green::before  { background: linear-gradient(90deg, var(--green), transparent 70%); }
.kc.green::after   { content:''; position:absolute; inset:0; border-radius:10px; background:linear-gradient(160deg,rgba(13,138,74,.05) 0%,transparent 60%); pointer-events:none; }
.kc.green:hover    { border-color: rgba(34,197,94,.4); box-shadow:  0 14px 40px rgba(0,0,0,.5),0 0 24px rgba(13,138,74,.07); }
.kc.blue   { border-color: rgba(20,68,168,.25); }
.kc.blue::before   { background: linear-gradient(90deg, var(--blue),  transparent 70%); }
.kc.blue::after    { content:''; position:absolute; inset:0; border-radius:10px; background:linear-gradient(160deg,rgba(20,68,168,.05) 0%,transparent 60%); pointer-events:none; }
.kc.blue:hover     { border-color: rgba(59,130,246,.4); box-shadow:  0 14px 40px rgba(0,0,0,.5),0 0 24px rgba(20,68,168,.07); }
.kc-val {
  font-family: var(--ff); font-size: 2.2rem; font-weight: 700; line-height: 1.1;
  letter-spacing: -.04em; position: relative; z-index: 1;
  animation: numIn .4s .1s ease both; animation-fill-mode: both;
}
.kc.gold   .kc-val { color: var(--gold3); }
.kc.red    .kc-val { color: var(--red2);  }
.kc.orange .kc-val { color: var(--ora2);  }
.kc.green  .kc-val { color: var(--grn2);  }
.kc.blue   .kc-val { color: var(--blu2);  }

.kc-lbl {
  font-family: var(--fi); font-size: .6rem; font-weight: 500;
  text-transform: uppercase; letter-spacing: .16em;
  color: var(--t3); margin-top: 5px; position: relative; z-index: 1;
}

/* Section dividers labels styling */
.sl {
  font-family: var(--fi); font-size: .58rem; font-weight: 600;
  letter-spacing: .22em; text-transform: uppercase; color: var(--t3);
  display: flex; align-items: center; gap: .75rem;
  margin: 1.75rem 0 1rem;
}
.sl::after { content: ''; flex: 1; height: 1px; background: linear-gradient(90deg, var(--b2), transparent); }

/* Analytical summary metrics grid */
.ps-row { display: flex; gap: .65rem; flex-wrap: wrap; margin: 1rem 0 1.5rem; }
.ps {
  background: var(--bg3); border: 1px solid var(--b2); border-radius: 9px;
  padding: .65rem 1.1rem .7rem; min-width: 86px;
  animation: popIn .35s ease both; animation-fill-mode: both;
  transition: border-color .2s, transform .15s; position: relative; overflow: hidden;
}
.ps:hover { transform: translateY(-2px); border-color: var(--b3); }
.ps::before { content:''; position:absolute; top:0; left:0; right:0; height:2px; }
.ps.red::before    { background: var(--red); }
.ps.orange::before { background: var(--orange); }
.ps.green::before  { background: var(--green); }
.ps.blue::before   { background: var(--blue); }
.ps-val { font-family: var(--ff); font-size: 1.7rem; font-weight: 700; line-height: 1; letter-spacing: -.03em; }
.ps-lbl { font-family: var(--fi); font-size: .58rem; text-transform: uppercase; letter-spacing: .14em; color: var(--t3); margin-top: 4px; }
.ps.red    .ps-val { color: var(--red2); }
.ps.orange .ps-val { color: var(--ora2); }
.ps.green  .ps-val { color: var(--grn2); }
.ps.blue   .ps-val { color: var(--blu2); }

/* Core context informative panels */
.ic {
  background: var(--bg3); border: 1px solid var(--b2); border-radius: 12px;
  padding: 1.4rem 1.6rem; font-family: var(--fi); font-size: .82rem;
  color: var(--t2); line-height: 1.9;
  animation: slideRight .4s ease both; animation-fill-mode: both;
}
.ic-title { font-family: var(--ff); font-size: .58rem; font-weight: 600;
  letter-spacing: .2em; text-transform: uppercase; color: var(--gold2); margin-bottom: .4rem; }

/* Visual validation component banners */
.sb {
  background: linear-gradient(135deg,rgba(13,138,74,.12),rgba(13,138,74,.04));
  border: 1px solid rgba(13,138,74,.3); border-radius: 10px;
  padding: 1rem 1.5rem; color: var(--grn3); font-family: var(--ff);
  font-size: .92rem; font-weight: 500;
  animation: successPop .5s cubic-bezier(.34,1.56,.64,1) both;
  display: flex; align-items: center; gap: .75rem;
}
.ac {
  background: rgba(13,138,74,.04); border: 1px solid rgba(13,138,74,.12);
  border-radius: 10px; padding: 1.75rem; text-align: center;
  color: var(--grn3); font-family: var(--ff); font-size: .95rem; font-weight: 500;
}

/* Core sidebar presentation design elements */
.logo { font-family: var(--ff); font-size: 1.15rem; font-weight: 700;
  letter-spacing: .04em; color: var(--gold2); display:flex; align-items:center; gap:.5rem; }
.logo-tag { font-family: var(--fi); font-size: .57rem; text-transform: uppercase;
  letter-spacing: .2em; color: var(--t3); margin-top: 3px; }
.logo-rule { height: 1px; margin: 1.2rem 0;
  background: linear-gradient(90deg, var(--gold), transparent); }
.vc {
  display: flex; align-items: center; justify-content: space-between;
  background: var(--bg3); border: 1px solid var(--b2); border-radius: 7px;
  padding: .5rem .8rem; margin-bottom: .28rem;
  transition: border-color .2s, background .2s;
  animation: slideRight .28s ease both; animation-fill-mode: both;
}
.vc:hover { border-color: var(--b3); background: var(--bg4); }
.vc.crit  { border-left: 2px solid var(--red); }
.vc.warn  { border-left: 2px solid var(--orange); }
.vc.safe  { border-left: 2px solid var(--green); }
.vc-name  { font-family: var(--ff); font-size: .73rem; font-weight: 600;
  color: var(--t1); white-space: nowrap; overflow: hidden;
  text-overflow: ellipsis; max-width: 115px; }
.vc-tags  { display: flex; gap: 4px; align-items: center; flex-shrink: 0; }
.vt       { font-family: var(--fm); font-size: .56rem; font-weight: 500;
  padding: 1px 6px; border-radius: 3px; }
.vt.od    { background: rgba(204,40,40,.15);  color: var(--red2); }
.vt.hp    { background: rgba(184,85,24,.15);  color: var(--ora2); }
.vt.ok    { background: rgba(13,138,74,.15);  color: var(--grn2); }

/* Premium Asset overview telemetry dashboard panels */
.vc-hero {
  background: var(--bg3); border: 1px solid var(--b2);
  border-radius: 12px; padding: 1.25rem 1.6rem;
  display: flex; align-items: center; justify-content: space-between;
  flex-wrap: wrap; gap: 1rem; margin: .5rem 0 1.5rem;
  animation: slideDown .4s ease both;
}
.vc-hero.crit { border-left: 4px solid var(--red); }
.vc-hero.warn { border-left: 4px solid var(--orange); }
.vc-hero.safe { border-left: 4px solid var(--green); }
.vc-hero-name { font-family: var(--ff); font-size: 1.25rem; font-weight: 700; color: var(--t0); }
.vc-hero-meta { font-family: var(--fm); font-size: .66rem; color: var(--t3); margin-top: 3px; }
.vc-hero-stats { display: flex; gap: .9rem; align-items: center; flex-wrap: wrap; }
.vc-stat { text-align: center; }
.vc-stat-val { font-family: var(--ff); font-size: 1.5rem; font-weight: 800; line-height: 1; }
.vc-stat-lbl { font-family: var(--fi); font-size: .58rem; text-transform: uppercase;
  letter-spacing: .14em; color: var(--t3); margin-top: 2px; }
.sev-badge { border-radius: 6px; padding: 5px 13px;
  font-family: var(--fm); font-size: .68rem; font-weight: 700;
  letter-spacing: .07em; }
.sev-badge.crit { background: rgba(204,40,40,.14); color: var(--red2); border: 1px solid rgba(204,40,40,.4); }
.sev-badge.warn { background: rgba(184,85,24,.14); color: var(--ora2); border: 1px solid rgba(184,85,24,.4); }
.sev-badge.safe { background: rgba(13,138,74,.12); color: var(--grn2); border: 1px solid rgba(13,138,74,.35); }

/* Inline meta line */
.ml { display: flex; gap: 1.4rem; flex-wrap: wrap;
  font-family: var(--fm); font-size: .66rem; color: var(--t3); margin: .7rem 0 0; }
.ml b { color: var(--t2); font-weight: 500; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════
#  DATABASE LAYER (INTEGRATED WAL LOCK PROTECTION)
# ═══════════════════════════════════════════════════════════════════
DB_PATH = Path("running_hours.db")

def get_db():
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    c = get_db()
    # Forces Write-Ahead Logging to guarantee synchronous multi-user pipelines
    c.execute("PRAGMA journal_mode=WAL;")
    c.executescript("""
    CREATE TABLE IF NOT EXISTS vessels(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        created_at TEXT NOT NULL DEFAULT(datetime('now')));
    CREATE TABLE IF NOT EXISTS upload_log(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL, filename TEXT NOT NULL,
        file_hash TEXT NOT NULL, report_date TEXT,
        me_total_hrs INTEGER, me_this_month INTEGER,
        uploaded_at TEXT NOT NULL DEFAULT(datetime('now')));
    CREATE TABLE IF NOT EXISTS components(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL, category TEXT NOT NULL,
        engine_label TEXT NOT NULL, unit TEXT NOT NULL,
        description TEXT NOT NULL, periodicity REAL,
        last_oh_date TEXT, last_oh_hrs REAL,
        hrs_since REAL, pct_used REAL, status TEXT NOT NULL,
        updated_at TEXT NOT NULL DEFAULT(datetime('now')));
    CREATE TABLE IF NOT EXISTS other_equipment(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL, section TEXT NOT NULL,
        description TEXT NOT NULL, periodicity TEXT,
        last_date TEXT, run_hrs TEXT,
        updated_at TEXT NOT NULL DEFAULT(datetime('now')));
    CREATE INDEX IF NOT EXISTS idx_cv ON components(vessel_name);
    CREATE INDEX IF NOT EXISTS idx_cs ON components(status);
    """)
    c.commit(); c.close()

init_db()


# ═══════════════════════════════════════════════════════════════════
#  CONVERSION  — Headless File Decryption Engine
# ═══════════════════════════════════════════════════════════════════
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found. packages.txt must contain: libreoffice")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as t:
        t.write(raw); tp = t.name
    od = tempfile.gettempdir()
    base = os.path.splitext(os.path.basename(tp))[0]
    dp = os.path.join(od, base + ".docx")
    pf = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        r = subprocess.run(
            [soffice,"--headless","--norestore","--nofirststartwizard",
             f"-env:UserInstallation={pf}","--convert-to","docx",tp,"--outdir",od],
            capture_output=True, timeout=120)
        if not os.path.exists(dp):
            raise RuntimeError(f"exit {r.returncode}: {r.stderr.decode('utf-8','ignore')[:300]}")
        with open(dp,"rb") as f: return f.read()
    finally:
        for p in [tp,dp]:
            try:
                if os.path.exists(p): os.unlink(p)
            except Exception: pass


# ═══════════════════════════════════════════════════════════════════
#  DETERMINISTIC TELEMETRY STREAM PARSER
# ═══════════════════════════════════════════════════════════════════
def _cp(raw):
    if not raw: return None
    s = re.sub(r'\.(?=\d{3}(\.|$))','',raw.strip())
    s = re.sub(r'[^0-9\.]','',s)
    try: return float(s) if s else None
    except ValueError: return None

def _pd(raw):
    if not raw or raw.strip() in ('','N/A','n/a'): return None
    raw = re.sub(r'\s+',' ',raw.strip().lstrip('[').rstrip(']'))
    if re.match(r'^\d+$',raw.strip()): return None
    rn = re.sub(r'\bSEPT\b','SEP',raw,flags=re.IGNORECASE)
    rn = re.sub(r'\bJUNE\b','JUN',rn, flags=re.IGNORECASE)
    rn = re.sub(r'\bJULY\b','JUL',rn, flags=re.IGNORECASE)
    fmts=['%d %b %y','%d %B %y','%d %b %Y','%d %B %Y',
          '%d/%m/%y','%d/%m/%Y','%d-%m-%y','%d-%m-%Y',
          '%b %Y','%B %Y','%Y-%m-%d']
    for fmt in fmts:
        for v in (rn,rn.upper(),rn.title(),raw,raw.upper()):
            try: return datetime.strptime(v,fmt).strftime('%Y-%m-%d')
            except ValueError: pass
    return raw

def _ph(raw):
    if not raw or raw.strip() in ('','N/A','n/a'): return None
    for n in re.findall(r'\d[\d,]*',raw.replace('\n',' ')):
        try:
            v=float(n.replace(',',''))
            if v>0: return v
        except ValueError: pass
    return None

def _st(h,p):
    if h is None or p is None or p==0: return 'NO DATA'
    r=h/p
    if r>1.0: return 'OVERDUE'
    if r>=0.80: return 'HIGH PRIORITY'
    return 'OK'

def _pt(h,p):
    if h is None or p is None or p==0: return 0.0
    return round(h/p,4)

def parse_doc_bytes(docx: bytes) -> dict:
    from docx import Document
    warns=[]
    with tempfile.NamedTemporaryFile(suffix='.docx',delete=False) as t:
        t.write(docx); tp=t.name
    try: doc=Document(tp)
    except Exception as e: raise ValueError(f"Cannot open: {e}")
    finally:
        try: os.unlink(tp)
        except Exception: pass
    if not doc.tables: raise ValueError("No tables — is this TEC-004?")

    vn='UNKNOWN'; rd=None
    for para in doc.paragraphs:
        txt=para.text.strip()
        if not txt: continue
        vm=re.search(r"Vessel[\u2019\u2018']?s?\s+Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+?)(?:\s{2,}|\t)",txt,re.IGNORECASE)
        dm=re.search(r"Date\s*:\s*(.+)",txt,re.IGNORECASE)
        if vm: vn=vm.group(1).strip()
        if dm: rd=_pd(dm.group(1).strip())
        if vm or dm: break
    if vn=='UNKNOWN': warns.append("Could not extract vessel name.")

    mt=mm=None; comps=[]; t0=doc.tables[0]
    for cell in t0.rows[0].cells:
        x=cell.text
        if m:=re.search(r'Total Running Hours[\s:ǀ|]+([\d,]+)',x,re.IGNORECASE):
            try: mt=int(m.group(1).replace(',',''))
            except ValueError: pass
        if m:=re.search(r'This Month[\s:]+([\d,]+)',x,re.IGNORECASE):
            try: mm=int(m.group(1).replace(',',''))
            except ValueError: pass

    cc=[]
    if len(t0.rows)>1:
        for ci,cell in enumerate(t0.rows[1].cells):
            if m:=re.search(r'CYL\s*\.?\s*No\s*\.?\s*(\d+)',cell.text.strip(),re.IGNORECASE):
                lbl=f"Cyl {int(m.group(1))}"
                if not cc or cc[-1][1]!=lbl: cc.append((ci,lbl))

    rows=t0.rows; i=2
    while i<len(rows)-1:
        r1=[c.text.strip() for c in rows[i].cells]
        r2=[c.text.strip() for c in rows[i+1].cells] if i+1<len(rows) else []
        nm=r1[0] if r1 else ''
        if not nm: i+=1; continue
        if (r1[2] if len(r1)>2 else '')=='1' and (r2[2] if r2 else '')=='2' and r1[0]==(r2[0] if r2 else ''):
            p=_cp(r1[1] if len(r1)>1 else '')
            for ci,lbl in cc:
                d=_pd(r1[ci]) if ci<len(r1) else None
                h=_ph(r2[ci]) if ci<len(r2) else None
                if d is None and h is None: continue
                comps.append({'category':'MAIN_ENGINE','engine_label':'ME','unit':lbl,
                    'description':nm,'periodicity':p,'last_oh_date':d,'last_oh_hrs':h,
                    'hrs_since':h,'pct_used':_pt(h,p),'status':_st(h,p)})
            i+=2
        else: i+=1

    oe=[]
    if len(doc.tables)>1:
        t1=doc.tables[1]
        SK={'TURBOCHARGER','AUXILIARY BOILER','COOLERS','EXH GAS  BOILER',
            'A/C & REFR. COMPRESSORS','MAIN AIR COMPRESSORS',
            'PERIODICTLY','DATE OF LAST INSPECTION','RUN HRS',
            'DATE OF LAST CLEANING','DATE','PERIODICITY'}
        for row in t1.rows:
            cells=[c.text.strip() for c in row.cells]
            for sec,dc,datec,hrsc in [('TURBOCHARGER / AUX BOILER',0,1,3),
                                       ('COOLERS / EXH GAS BOILER',5,6,8),
                                       ('A/C & COMPRESSORS',10,11,12)]:
                desc=cells[dc] if dc<len(cells) else ''
                if not desc or desc.upper() in SK: continue
                dv=cells[datec] if datec<len(cells) else ''
                hv=cells[hrsc]  if hrsc <len(cells) else ''
                if dv or hv:
                    oe.append({'section':sec,'description':desc,'periodicity':'','last_date':dv,'run_hrs':hv})

    if len(doc.tables)>2:
        t2=doc.tables[2]; rows2=t2.rows; eblocks=[]
        if rows2:
            hdr=[c.text.strip() for c in rows2[0].cells]
            tot=[c.text.strip() for c in rows2[2].cells] if len(rows2)>2 else []
            seen=set()
            for ci,cell in enumerate(hdr):
                if m:=re.search(r'Aux\.\s*Engine\s*No\.?\s*(\d+)',cell,re.IGNORECASE):
                    lbl=f"AUX-{int(m.group(1))}"
                    if lbl not in seen:
                        seen.add(lbl)
                        th=next((_ph(tot[j]) for j in range(ci,min(ci+14,len(tot))) if _ph(tot[j])),None)
                        eblocks.append((lbl,ci,th))
        cm={}
        if len(rows2)>4:
            r4=[c.text.strip() for c in rows2[4].cells]
            for ei,(elbl,es,_) in enumerate(eblocks):
                ee=eblocks[ei+1][1] if ei+1<len(eblocks) else len(r4)
                sc2: list=[]
                for ci in range(es,ee):
                    if ci<len(r4):
                        try:
                            cn=int(r4[ci])
                            if cn not in sc2: sc2.append(cn); cm[ci]=(elbl,cn)
                        except ValueError: pass
        i2=5
        while i2<len(rows2)-1:
            r1=[c.text.strip() for c in rows2[i2].cells]
            r2=[c.text.strip() for c in rows2[i2+1].cells] if i2+1<len(rows2) else []
            nm=r1[0] if r1 else ''
            if not nm: i2+=1; continue
            if (r1[2] if len(r1)>2 else '') in ('1','2') and r1[0]==(r2[0] if r2 else ''):
                p=_cp(r1[1] if len(r1)>1 else '')
                for ci,(elbl,cn) in cm.items():
                    d=_pd(r1[ci]) if ci<len(r1) else None
                    h=_ph(r2[ci]) if ci<len(r2) else None
                    if d is None and h is None: continue
                    comps.append({'category':'AUX_ENGINE','engine_label':elbl,'unit':f"Cyl {cn}",
                        'description':nm,'periodicity':p,'last_oh_date':d,'last_oh_hrs':h,
                        'hrs_since':h,'pct_used':_pt(h,p),'status':_st(h,p)})
                i2+=2
            else: i2+=1

    if len(doc.tables)>3:
        t3=doc.tables[3]; dg=['D/G 1','D/G 2','D/G 3']
        for ri,row in enumerate(t3.rows):
            cells=[c.text.strip() for c in row.cells]
            if ri==0: continue
            for dc,pc,tc,ds in [(0,1,2,3),(9,10,11,12)]:
                desc=cells[dc] if dc<len(cells) else ''
                per=cells[pc]  if pc<len(cells) else ''
                rt=cells[tc]   if tc<len(cells) else ''
                if not desc or rt not in ('1','2'): continue
                for gi,gl in enumerate(dg):
                    col=ds+gi; val=cells[col] if col<len(cells) else ''
                    if not val: continue
                    key=f"{desc} — {gl}"
                    if rt=='1':
                        oe.append({'section':'D/G EQUIPMENT','description':key,
                            'periodicity':per,'last_date':_pd(val) or val,'run_hrs':''})
                    else:
                        for e in reversed(oe):
                            if e['description']==key and e['run_hrs']=='':
                                e['run_hrs']=val; break
                        else:
                            oe.append({'section':'D/G EQUIPMENT','description':key,
                                'periodicity':per,'last_date':'','run_hrs':val})

    return {'vessel_name':vn,'report_date':rd,'me_total_hrs':mt,'me_this_month':mm,
            'components':comps,'other_equipment':oe,'warnings':warns}


# ═══════════════════════════════════════════════════════════════════
#  PERSISTENT DATA MANAGEMENT LAYER (TRANSACTION ROLLBACK SECURED)
# ═══════════════════════════════════════════════════════════════════
def save_parsed(parsed, filename, fhash):
    conn = get_db()
    c = conn.cursor()
    now = datetime.utcnow().isoformat() + "Z"
    v = parsed['vessel_name']
    
    # Fully secured context execution loop protecting table updates against failures
    try:
        c.execute("INSERT OR IGNORE INTO vessels(name,created_at) VALUES(?,?)",(v,now))
        c.execute("INSERT INTO upload_log(vessel_name,filename,file_hash,report_date,me_total_hrs,me_this_month,uploaded_at) VALUES(?,?,?,?,?,?,?)",
            (v,filename,fhash,parsed['report_date'],parsed['me_total_hrs'],parsed['me_this_month'],now))
        c.execute("DELETE FROM components WHERE vessel_name=?",(v,))
        for x in parsed['components']:
            c.execute("INSERT INTO components(vessel_name,category,engine_label,unit,description,periodicity,last_oh_date,last_oh_hrs,hrs_since,pct_used,status,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                (v,x['category'],x['engine_label'],x['unit'],x['description'],
                 x['periodicity'],x['last_oh_date'],x['last_oh_hrs'],
                 x['hrs_since'],x['pct_used'],x['status'],now))
        c.execute("DELETE FROM other_equipment WHERE vessel_name=?",(v,))
        for x in parsed['other_equipment']:
            c.execute("INSERT INTO other_equipment(vessel_name,section,description,periodicity,last_date,run_hrs,updated_at) VALUES(?,?,?,?,?,?,?)",
                (v,x['section'],x['description'],x.get('periodicity',''),
                 x.get('last_date',''),x.get('run_hrs',''),now))
        conn.commit()
    except Exception as e:
        conn.rollback() # Safeguards database files from half-written executions
        raise e
    finally:
        conn.close()

@st.cache_data(ttl=10)
def get_vessels():
    c=get_db(); r=c.execute("SELECT name FROM vessels ORDER BY name").fetchall(); c.close()
    return [x['name'] for x in r]

@st.cache_data(ttl=10)
def get_comps(vessel):
    c=get_db()
    df=pd.read_sql_query("SELECT * FROM components WHERE vessel_name=?",c,params=(vessel,))
    c.close(); return df

@st.cache_data(ttl=10)
def get_oe(vessel):
    c=get_db()
    df=pd.read_sql_query("SELECT * FROM other_equipment WHERE vessel_name=? ORDER BY section,description",c,params=(vessel,))
    c.close(); return df

@st.cache_data(ttl=10)
def get_history(vessel):
    c=get_db()
    df=pd.read_sql_query("SELECT filename,report_date,me_total_hrs,me_this_month,uploaded_at FROM upload_log WHERE vessel_name=? ORDER BY uploaded_at DESC LIMIT 20",c,params=(vessel,))
    c.close(); return df

@st.cache_data(ttl=10)
def get_summary():
    c=get_db()
    df=pd.read_sql_query("""
        SELECT c.vessel_name,
            COUNT(CASE WHEN c.status='OVERDUE'       THEN 1 END) AS overdue,
            COUNT(CASE WHEN c.status='HIGH PRIORITY' THEN 1 END) AS high_priority,
            COUNT(CASE WHEN c.status='OK'            THEN 1 END) AS ok,
            COUNT(*) AS total,
            MAX(u.uploaded_at) AS last_upload,
            MAX(u.me_total_hrs) AS me_total_hrs,
            MAX(u.report_date) AS report_date
        FROM components c
        LEFT JOIN upload_log u ON u.vessel_name=c.vessel_name
        GROUP BY c.vessel_name ORDER BY overdue DESC, high_priority DESC
    """,c); c.close(); return df

@st.cache_data(ttl=10)
def get_all_fleet_comps():
    c = get_db()
    df = pd.read_sql_query("SELECT * FROM components", c)
    c.close(); return df


# ═══════════════════════════════════════════════════════════════════
#  ULTRA PREMIUM MASTER RENDERING ENGINE (Failsafe Integrated)
# ═══════════════════════════════════════════════════════════════════
def _sf(x):
    try:
        v=float(x); return None if pd.isna(v) else v
    except: return None

def _cyl(u):
    m=re.search(r'\d+',str(u)); return int(m.group()) if m else 999

_S = {
    'OVERDUE':       {'bg':'#1f0505','bgs':'#2d0707','ts':'#ff6b6b','tm':'#ff8080','tn':'#ff3333','td':'#773333'},
    'HIGH PRIORITY': {'bg':'#1e0d02','bgs':'#2d1503','ts':'#ffaa44','tm':'#ff9933','tn':'#ffcc00','td':'#774422'},
    'OK':            {'bg':'#021208','bgs':'#042010','ts':'#4ade80','tm':'#22c55e','tn':'#4ade80','td':'#0f4023'},
    '_':             {'bg':'#090e18','bgs':'#0c1422','ts':'#4a6688','tm':'#334d66','tn':'#334d66','td':'#1a2a38'},
}

_COL_CFG = {
    "Status":      st.column_config.TextColumn("Status",      width=130),
    "Vessel":      st.column_config.TextColumn("Vessel",      width=140),
    "Component":   st.column_config.TextColumn("Component",   width=205),
    "Engine":      st.column_config.TextColumn("Engine",      width=80),
    "Unit":        st.column_config.TextColumn("Unit",        width=68),
    "Periodicity": st.column_config.NumberColumn("Periodicity", format="%d", width=100),
    "Last O/H":    st.column_config.TextColumn("Last O/H",    width=100),
    "Hrs Since":   st.column_config.NumberColumn("Hrs Since",  format="%d hrs", width=100),
    "% Used":      st.column_config.ProgressColumn(
                       "% Used", min_value=0, max_value=160, format="%.1f%%", width=120),
}

def _build_display(df: pd.DataFrame, sort_priority: bool = False) -> pd.DataFrame:
    """Build the display table with smart conditional sorting based on data context."""
    if df.empty:
        return pd.DataFrame(columns=['Status','Vessel','Component','Engine','Unit',
                                      'Periodicity','Last O/H','Hrs Since','% Used'])
    d = df.copy()
    if sort_priority:
        _ORD = {'OVERDUE':0,'HIGH PRIORITY':1,'OK':2,'NO DATA':3}
        d['_s'] = d['status'].map(lambda s: _ORD.get(str(s),4))
        d['_p'] = d['pct_used'].apply(lambda x: _sf(x) or 0.0)
        # Condition Check: Are we in the Master Matrix (has vessel) or Pre-Commit (no vessel yet)?
        if 'vessel_name' in d.columns:
            d = d.sort_values(['_s','_p','vessel_name'], ascending=[True,False,True]).drop(columns=['_s','_p'])
        else:
            d = d.sort_values(['_s','_p'], ascending=[True,False]).drop(columns=['_s','_p'])
    else:
        d['_k1'] = d['description'].str.upper()
        d['_k2'] = d['unit'].apply(_cyl)
        if 'vessel_name' in d.columns:
            d = d.sort_values(['vessel_name','_k1','_k2']).drop(columns=['_k1','_k2'])
        else:
            d = d.sort_values(['_k1','_k2']).drop(columns=['_k1','_k2'])

    out = pd.DataFrame(index=range(len(d)))
    out['Status']      = d['status'].values
    out['Vessel']      = d.get('vessel_name', pd.Series(['—']*len(d))).values
    out['Component']   = d['description'].values
    out['Engine']      = d['engine_label'].values
    out['Unit']        = d['unit'].values
    out['Periodicity'] = [int(float(x)) if _sf(x) else None for x in d['periodicity'].values]
    out['Last O/H']    = [str(x) if x and str(x) not in ('nan','None','') else '—' for x in d['last_oh_date'].values]
    out['Hrs Since']   = [int(float(x)) if _sf(x) else None for x in d['hrs_since'].values]
    out['% Used']      = [round(float(x)*100,1) if _sf(x) else 0.0 for x in d['pct_used'].values]
    return out

def _apply_style(df: pd.DataFrame):
    def rs(row):
        c = _S.get(str(row.get('Status','')), _S['_'])
        return [
            f"background-color:{c['bgs']};color:{c['ts']};font-weight:700",   # Status
            f"background-color:{c['bg']};color:#c0d0e8;font-weight:700",      # Vessel
            f"background-color:{c['bg']};color:{c['tm']};font-weight:600",     # Component
            f"background-color:{c['bg']};color:{c['td']}",                     # Engine
            f"background-color:{c['bg']};color:{c['td']}",                     # Unit
            f"background-color:{c['bg']};color:{c['td']}",                     # Periodicity
            f"background-color:{c['bg']};color:{c['td']}",                     # Last O/H
            f"background-color:{c['bg']};color:{c['tm']};font-weight:600",     # Hrs Since
            f"background-color:{c['bg']};color:{c['tn']};font-weight:700",     # % Used
        ]
    return df.style.apply(rs, axis=1)

def render_table(df: pd.DataFrame, height: int = None, priority: bool = False):
    if isinstance(df, list): df = pd.DataFrame(df)
    if df.empty: st.info("No data to display."); return
    tbl = _build_display(df, sort_priority=priority)
    h   = height or min(720, 38*(len(tbl)+1)+4)
    st.dataframe(_apply_style(tbl), use_container_width=True,
                 hide_index=True, height=h, column_config=_COL_CFG)


# ═══════════════════════════════════════════════════════════════════
#  UI LAYOUT INJECTION ENGINE
# ═══════════════════════════════════════════════════════════════════
def kpi(val, lbl, color="gold", delay=0):
    return (f'<div class="kc {color}" style="animation-delay:{delay}s">'
            f'<div class="kc-val">{val}</div><div class="kc-lbl">{lbl}</div></div>')

def ph(icon, title, eye=""):
    e = f'<div class="ph-eye">{eye}</div>' if eye else ''
    return f'<div class="ph">{e}<h1>{icon}&nbsp;&nbsp;{title}</h1></div><div class="ph-line"></div>'

def sl(txt):
    return f'<div class="sl">{txt}</div>'

def filter_count(total, od, hp, ok):
    return (f'<div class="filter-count">Showing <b>{total}</b> records — '
            f'<span class="od">{od} overdue</span> · '
            f'<span class="hp">{hp} high priority</span> · '
            f'<span class="ok">{ok} OK</span></div>')


# ═══════════════════════════════════════════════════════════════════
#  SIDEBAR PRESENTATION LAYER
# ═══════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown('<div class="logo">⚓ FLEET MONITOR</div>'
                '<div class="logo-tag">Running Hours Management System</div>'
                '<div class="logo-rule"></div>', unsafe_allow_html=True)

    page = st.selectbox("nav",
        ["🗺️  Fleet Overview","🚢  Vessel Detail",
         "📤  Upload Report","📋  Upload History"],
        label_visibility="collapsed")

    st.markdown("<br>", unsafe_allow_html=True)
    vessels = get_vessels()
    sel_v   = st.selectbox("Active Vessel", vessels) if vessels else None
    if not vessels: st.info("No data — upload a report to begin.")

    if vessels:
        smry = get_summary()
        if not smry.empty:
            st.markdown('<div class="logo-rule"></div>', unsafe_allow_html=True)
            st.markdown('<div style="font-family:var(--fi);font-size:.56rem;text-transform:uppercase;'
                        'letter-spacing:.2em;color:var(--t3);margin-bottom:.5rem">Vessel Status</div>',
                        unsafe_allow_html=True)
            for idx,(_, row) in enumerate(smry.iterrows()):
                od=int(row['overdue']); hp=int(row['high_priority'])
                cls='crit' if od>0 else ('warn' if hp>0 else 'safe')
                tags=''
                if od>0: tags+=f'<span class="vt od">{od} OD</span>'
                if hp>0: tags+=f'<span class="vt hp">{hp} HP</span>'
                if od==0 and hp==0: tags+='<span class="vt ok">✓ OK</span>'
                st.markdown(
                    f'<div class="vc {cls}" style="animation-delay:{idx*.04}s">'
                    f'<div class="vc-name">{row["vessel_name"]}</div>'
                    f'<div class="vc-tags">{tags}</div></div>',
                    unsafe_allow_html=True)

    st.markdown('<div class="logo-rule"></div>', unsafe_allow_html=True)
    db_kb = DB_PATH.stat().st_size/1024 if DB_PATH.exists() else 0
    st.markdown(f'<div style="font-family:var(--fm);font-size:.58rem;color:var(--t3)">'
                f'db {db_kb:.0f} kb · {len(vessels)} vessels · v5.2</div>',
                unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════
#  PAGE LAYOUT CONTROLLERS
# ═══════════════════════════════════════════════════════════════════

# ── PAGE: FLEET MASTER MATRIX (ULTRA-PREMIUM CONTROL DECK) ─────────
if page == "🗺️  Fleet Overview":
    st.markdown(ph("🗺️", "Fleet Master Matrix", "Universal Fleet Telemetry"), unsafe_allow_html=True)

    smry = get_summary()
    all_comps = get_all_fleet_comps()
    
    if smry.empty or all_comps.empty:
        st.info("No data loaded. Upload a report to begin."); st.stop()

    tv = len(smry); tc = len(all_comps)
    tod = int((all_comps['status'] == 'OVERDUE').sum())
    thp = int((all_comps['status'] == 'HIGH PRIORITY').sum())
    tok = int((all_comps['status'] == 'OK').sum())
    
    k1,k2,k3,k4,k5 = st.columns(5)
    for col,(val,lbl,clr,dly) in zip([k1,k2,k3,k4,k5],[
        (tv,"Vessels","blue",0),(tc,"Components","gold",.07),
        (tod,"Overdue","red",.14),(thp,"High Priority","orange",.21),(tok,"OK","green",.28)]):
        with col: st.markdown(kpi(val,lbl,clr,dly), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(sl("Universal Component Control Grid"), unsafe_allow_html=True)

    # Universal Matrix Layout Filters
    f1, f2, f3, f4 = st.columns([1.5, 1.5, 2, 2])
    with f1:
        vessel_f = st.selectbox("Filter Vessel Context", ["All Fleet"] + sorted(all_comps['vessel_name'].unique().tolist()), key="mst_v")
    with f2:
        cat_f = st.selectbox("Filter Machinery Type", ["All", "Main Engine", "Aux Engines"], key="mst_cat")
    with f3:
        st_f = st.selectbox("Filter Component Urgency", ["🔴 Critical Focus (Overdue + High Pri)", "All Statuses", "🔴 Overdue Only", "🟡 High Priority Only", "🟢 OK Only"], key="mst_st")
    with f4:
        comp_f = st.selectbox("Search Component Definition", ["All"] + sorted(all_comps['description'].unique().tolist()), key="mst_comp")

    filt = all_comps.copy()
    if vessel_f != "All Fleet": filt = filt[filt['vessel_name'] == vessel_f]
    if cat_f == "Main Engine":    filt = filt[filt['category'] == 'MAIN_ENGINE']
    elif cat_f == "Aux Engines":  filt = filt[filt['category'] == 'AUX_ENGINE']
    
    if st_f == "🔴 Critical Focus (Overdue + High Pri)": filt = filt[filt['status'].isin(['OVERDUE', 'HIGH PRIORITY'])]
    elif st_f == "🔴 Overdue Only":      filt = filt[filt['status'] == 'OVERDUE']
    elif st_f == "🟡 High Priority Only": filt = filt[filt['status'] == 'HIGH PRIORITY']
    elif st_f == "🟢 OK Only":           filt = filt[filt['status'] == 'OK']
        
    if comp_f != "All": filt = filt[filt['description'] == comp_f]

    ns = len(filt); no = int((filt['status'] == 'OVERDUE').sum())
    nh = int((filt['status'] == 'HIGH PRIORITY').sum()); nk = int((filt['status'] == 'OK').sum())
    st.markdown(filter_count(ns, no, nh, nk), unsafe_allow_html=True)

    if filt.empty:
        st.markdown('<div class="ac">✓ No records match the current filter matrix</div>', unsafe_allow_html=True)
    else:
        # Priority mapping pins worst lifecycle risks across all ships to the top row
        render_table(filt, height=min(900, 38 * (ns + 1) + 4), priority=True)

# ── PAGE: VESSEL DETAIL ANALYSIS ─────────────────────────────────────
# ── PAGE: VESSEL DETAIL ANALYSIS ─────────────────────────────────────
elif page == "🚢  Vessel Detail":
    if not sel_v:
        st.info("Select a vessel from the sidebar.")
        st.stop()

    st.markdown(ph("🚢", sel_v, "Component Analysis"), unsafe_allow_html=True)

    df = get_comps(sel_v)
    oe = get_oe(sel_v)

    if df.empty:
        st.info("No data for this vessel.")
        st.stop()

    # Topline status counts
    n_tot = len(df)
    n_od  = int((df["status"] == "OVERDUE").sum())
    n_hp  = int((df["status"] == "HIGH PRIORITY").sum())
    n_ok  = int((df["status"] == "OK").sum())
    n_nd  = int((df["status"] == "NO DATA").sum())

    k1, k2, k3, k4, k5 = st.columns(5)
    for col, (val, lbl, clr, dly) in zip(
        [k1, k2, k3, k4, k5],
        [
            (n_tot, "Total",        "gold",   0.00),
            (n_od,  "Overdue",      "red",    0.07),
            (n_hp,  "High Priority","orange", 0.14),
            (n_ok,  "OK",           "green",  0.21),
            (n_nd,  "No Data",      "blue",   0.28),
        ],
    ):
        with col:
            st.markdown(kpi(val, lbl, clr, dly), unsafe_allow_html=True)

    # Last upload meta
    hist = get_history(sel_v)
    if not hist.empty:
        last = hist.iloc[0]
        mt = f"{int(last['me_total_hrs']):,}"   if pd.notna(last["me_total_hrs"])   else "—"
        mm = f"{int(last['me_this_month']):,}" if pd.notna(last["me_this_month"]) else "—"
        st.markdown(
            f"""
            <div class="ml">
              <span>📄 <b>{last['filename']}</b></span>
              <span>Report: <b>{last['report_date'] or '—'}</b></span>
              <span>M/E: <b>{mt}</b> total · <b>{mm}</b> this month</span>
              <span>Uploaded: <b>{str(last['uploaded_at'])[:16]}</b></span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("---")
    tabs = st.tabs(["⚠️  Alerts", "⚙️  Main Engine", "🔩  Aux Engines", "🛠️  Other Equipment"])

    # Alerts tab
    with tabs[0]:
        st.markdown(sl("Urgent Interrupt Diagnostics — Attention Required"), unsafe_allow_html=True)
        alerts = df[df["status"].isin(["OVERDUE", "HIGH PRIORITY"])]
        if alerts.empty:
            st.markdown(
                '<div class="ac">✓ All machinery components within acceptable operational bounds</div>',
                unsafe_allow_html=True,
            )
        else:
            no = int((alerts["status"] == "OVERDUE").sum())
            nh = int((alerts["status"] == "HIGH PRIORITY").sum())
            st.markdown(filter_count(len(alerts), no, nh, 0), unsafe_allow_html=True)
            render_table(alerts, priority=True)

    # Main Engine tab
    with tabs[1]:
        me = df[df["category"] == "MAIN_ENGINE"]
        if me.empty:
            st.info("No Main Engine data available.")
        else:
            st.markdown(sl("Main Engine Telemetry Matrix"), unsafe_allow_html=True)
            fa, fb = st.columns(2)
            with fa:
                sel_mc = st.selectbox(
                    "Machinery Element",
                    ["All"] + sorted(me["description"].unique().tolist()),
                    key="me_c",
                )
            with fb:
                sel_ms = st.selectbox(
                    "Machinery Status",
                    ["All", "🔴 Overdue only", "🟡 High Priority +", "🟢 OK only"],
                    key="me_s",
                )

            v = me.copy()
            if sel_mc != "All":
                v = v[v["description"] == sel_mc]
            if sel_ms == "🔴 Overdue only":
                v = v[v["status"] == "OVERDUE"]
            elif sel_ms == "🟡 High Priority +":
                v = v[v["status"].isin(["OVERDUE", "HIGH PRIORITY"])]
            elif sel_ms == "🟢 OK only":
                v = v[v["status"] == "OK"]

            no = int((v["status"] == "OVERDUE").sum())
            nh = int((v["status"] == "HIGH PRIORITY").sum())
            nk = int((v["status"] == "OK").sum())
            st.markdown(filter_count(len(v), no, nh, nk), unsafe_allow_html=True)
            render_table(v, priority=False)

    # Aux Engines tab
    with tabs[2]:
        aux = df[df["category"] == "AUX_ENGINE"]
        if aux.empty:
            st.info("No Auxiliary Engine data available.")
        else:
            st.markdown(sl("Auxiliary Prime Movers Telemetry Matrix"), unsafe_allow_html=True)
            fa, fb = st.columns(2)
            with fa:
                sel_ae = st.selectbox(
                    "Aux Generator Node",
                    ["All"] + sorted(aux["engine_label"].unique().tolist()),
                    key="aux_e",
                )
            with fb:
                sel_as = st.selectbox(
                    "Node Condition",
                    ["All", "🔴 Overdue only", "🟡 High Priority +", "🟢 OK only"],
                    key="aux_s",
                )

            v = aux.copy()
            if sel_ae != "All":
                v = v[v["engine_label"] == sel_ae]
            if sel_as == "🔴 Overdue only":
                v = v[v["status"] == "OVERDUE"]
            elif sel_as == "🟡 High Priority +":
                v = v[v["status"].isin(["OVERDUE", "HIGH PRIORITY"])]
            elif sel_as == "🟢 OK only":
                v = v[v["status"] == "OK"]

            no = int((v["status"] == "OVERDUE").sum())
            nh = int((v["status"] == "HIGH PRIORITY").sum())
            nk = int((v["status"] == "OK").sum())
            st.markdown(filter_count(len(v), no, nh, nk), unsafe_allow_html=True)
            render_table(v, priority=False)

    # Other Equipment tab
    with tabs[3]:
        if oe.empty:
            st.info("No auxiliary plant or extension machinery metrics located.")
        else:
            for sec in sorted(oe["section"].unique()):
                st.markdown(sl(sec), unsafe_allow_html=True)
                sd = oe[oe["section"] == sec][
                    ["description", "periodicity", "last_date", "run_hrs"]
                ].copy()
                sd.columns = [
                    "Machinery Description",
                    "Maintenance Periodicity",
                    "Inspection Date",
                    "Logged Hours",
                ]
                st.dataframe(sd, use_container_width=True, hide_index=True)
# ── PAGE: REPORT INGESTION MANAGEMENT ────────────────────────────────
elif page == "📤  Upload Report":
    st.markdown(ph("📤","Upload Report","TEC-004 Log Processing"), unsafe_allow_html=True)

    col_up, col_info = st.columns([3,2], gap="large")
    with col_up:
        uploaded = st.file_uploader("file", type=["doc"], label_visibility="collapsed")
    with col_info:
        st.markdown("""
        <div class="ic">
          <div class="ic-title">Accepted Specification</div>
          TEC-004 Running Hours Monthly Log Report<br>
          Supported Frameworks &nbsp;·&nbsp; <b>Native .doc binary streams only</b><br><br>
          <div class="ic-title">Automated Extraction Vectors</div>
          ✦ Vessel Identifier &amp; Structural Report Date<br>
          ✦ M/E Absolute Combined &amp; Monthly Deltas<br>
          ✦ Complete Main Engine Structural Wear Sub-matrices<br>
          ✦ Auxiliary Generating Assemblies (3 Units × up to 7 Cylinders)<br>
          ✦ Turbocharging Plants, Compressors, &amp; Plant Systems<br>
          ✦ Fatigue Threshold Computations calculated synchronously
        </div>""", unsafe_allow_html=True)

    if uploaded:
        raw = uploaded.read()
        fh  = hashlib.md5(raw).hexdigest()
        with st.spinner("Executing secure LibreOffice conversion and structural extraction loops..."):
            try:
                docx = convert_doc_to_docx(raw)
            except Exception as e:
                st.error(f"Headless server file decryption faulted: `{e}`")
                st.stop()
            try:
                parsed = parse_doc_bytes(docx)
            except ValueError as e:
                st.error(f"Telemetry stream interpretation broke down: `{e}`")
                st.stop()

        # 🔁 Auto‑commit immediately after successful parse (no manual commit button)
        save_parsed(parsed, uploaded.name, fh)
        for fn in [get_vessels, get_comps, get_oe, get_summary, get_all_fleet_comps]:
            fn.clear()

        comps  = parsed['components']
        nc  = len(comps)
        nod = sum(1 for c in comps if c['status'] == 'OVERDUE')
        nhp = sum(1 for c in comps if c['status'] == 'HIGH PRIORITY')
        nok = sum(1 for c in comps if c['status'] == 'OK')
        noe = len(parsed['other_equipment'])

        st.markdown(sl("Extracted Telemetry Stream Preview"), unsafe_allow_html=True)
        c1,c2,c3,c4,c5 = st.columns(5)
        c1.metric("Asset",            parsed['vessel_name'])
        c2.metric("Report Window",    parsed['report_date'] or "—")
        c3.metric("M/E Accumulated",  f"{parsed['me_total_hrs']:,}"  if parsed['me_total_hrs']  else "—")
        c4.metric("Monthly Increment", f"{parsed['me_this_month']:,}" if parsed['me_this_month'] else "—")
        c5.metric("Data Channels",    nc)
        st.markdown(f"""
        <div class="ps-row">
          <div class="ps red"    style="animation-delay:0s">   <div class="ps-val">{nod}</div><div class="ps-lbl">Overdue</div></div>
          <div class="ps orange" style="animation-delay:.07s"> <div class="ps-val">{nhp}</div><div class="ps-lbl">High Priority</div></div>
          <div class="ps green"  style="animation-delay:.14s"> <div class="ps-val">{nok}</div><div class="ps-lbl">OK</div></div>
          <div class="ps blue"   style="animation-delay:.21s"> <div class="ps-val">{noe}</div><div class="ps-lbl">External Systems</div></div>
        </div>""", unsafe_allow_html=True)

        # Always show a success banner (commit already done)
        st.markdown(
            f'<div class="sb"><span style="font-size:1.4rem">✓</span>'
            f'<span>System Telemetry Confirmed — <b>{parsed["vessel_name"]}</b> '
            f'committed to database structure ({nc} lines mapped).</span>'
            '</div>',
            unsafe_allow_html=True
        )

        for w in parsed['warnings']:
            st.warning(f"⚠ Core structural indicator anomaly: {w}")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(sl("Extracted Telemetry Matrices — Live View"), unsafe_allow_html=True)

        df_preview = pd.DataFrame(parsed['components']) if parsed['components'] else pd.DataFrame()

# ── PAGE: AUDIT TRAIL LOG SYSTEM ─────────────────────────────────────
elif page == "📋  Upload History":
    st.markdown(ph("📋","Upload History","System Audit Trails"), unsafe_allow_html=True)
    if not sel_v: st.info("Select a tracking asset from the sidebar selector."); st.stop()
    st.markdown(sl(f"Chronological Logs: {sel_v}"), unsafe_allow_html=True)
    hist = get_history(sel_v)
    if hist.empty:
        st.info("No transaction trail entries recorded for this hull identification.")
    else:
        d=hist.copy()
        d.columns=['Logged Filename','Extracted Target Date','M/E Combined Total','M/E Monthly Increment','Transaction Timestamp']
        d['M/E Combined Total']  = d['M/E Combined Total'].apply(lambda x:f"{int(x):,}" if pd.notna(x) else '—')
        d['M/E Monthly Increment'] = d['M/E Monthly Increment'].apply(lambda x:f"{int(x):,}" if pd.notna(x) else '—')
        st.dataframe(d, use_container_width=True, hide_index=True)
