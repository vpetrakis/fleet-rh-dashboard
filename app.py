"""
Fleet Running Hours  v16  ·  FRONTEND REVISED / BACKEND PRESERVED
══════════════════════════════════════════════════════════════════════════════
BACKEND — preserved as-is in logic and schema:
  ME : raw-grid parser
  AUX: dedup-grid parser
  OE : exact column-indexed parser
  DG : paired-row dedup parser

FRONTEND — fully revised:
  Refined dark UI · cleaner hierarchy · responsive KPI grid
  Card-based sections · improved filters · polished HTML tables
  Better upload panel · better empty states · better matrix readability
══════════════════════════════════════════════════════════════════════════════
"""
import streamlit as st
st.set_page_config(
    page_title="Fleet Running Hours",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

import os, re, shutil, tempfile, subprocess, hashlib
from pathlib import Path
from datetime import datetime
from typing import Any, Dict, List, Tuple
import pandas as pd


# ══════════════════════════════════════════════════════════════════════════
#  DESIGN SYSTEM  — REVISED FRONTEND
# ══════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap');

:root{
  --bg:#050814;
  --bg2:#08101d;
  --bg3:#0c1526;
  --bg4:#101b30;
  --bg5:#13203a;
  --bg6:#182743;

  --line:#1b2b44;
  --line2:#243754;
  --line3:#30486a;

  --gold:#c08818;
  --gold2:#d8a93a;
  --gold3:#f2c85c;

  --red:#d83939;
  --amber:#dd8426;
  --green:#11a85b;
  --blue:#3e73da;

  --txt:#e5eefc;
  --txt2:#a8bdd8;
  --txt3:#6f88a7;
  --txt4:#445a77;
  --txt5:#1d2f47;

  --ff:'Space Grotesk',sans-serif;
  --fi:'Inter',sans-serif;
  --fm:'JetBrains Mono',monospace;

  --shadow-lg:0 28px 80px rgba(0,0,0,.42);
  --shadow-md:0 14px 42px rgba(0,0,0,.28);
  --shadow-sm:0 8px 24px rgba(0,0,0,.18);
  --radius-xl:20px;
  --radius-lg:16px;
  --radius-md:12px;
  --radius-sm:10px;
}

*,*::before,*::after{box-sizing:border-box}
html,body,[class*="css"]{
  background:var(--bg)!important;
  color:var(--txt2)!important;
  font-family:var(--fi)!important;
  -webkit-font-smoothing:antialiased;
  text-rendering:optimizeLegibility;
}
body{margin:0}
.main,.main>div{background:transparent!important}
.block-container{
  max-width:1500px!important;
  padding:1.2rem 1.4rem 5rem!important;
}
.block-container > *{position:relative;z-index:1}
.main::before{
  content:'';
  position:fixed; inset:0; pointer-events:none; z-index:0;
  background:
    radial-gradient(900px 500px at -10% -15%, rgba(192,136,24,.10), transparent 58%),
    radial-gradient(900px 600px at 120% 110%, rgba(62,115,218,.08), transparent 58%),
    linear-gradient(180deg, rgba(255,255,255,.01), transparent 18%);
}
[data-testid="stSidebar"],[data-testid="collapsedControl"]{display:none!important}
[data-testid="stAppViewContainer"]{background:transparent!important}
section.main > div{padding-top:0!important}

a,button{transition:all .18s ease}
hr{border-color:var(--line)!important;opacity:1!important}

::-webkit-scrollbar{width:8px;height:8px}
::-webkit-scrollbar-track{background:var(--bg2)}
::-webkit-scrollbar-thumb{background:var(--line3);border-radius:999px}
::-webkit-scrollbar-thumb:hover{background:#3b5880}

/* Upload */
[data-testid="stFileUploader"] > div:first-child{width:100%}
[data-testid="stFileUploadDropzone"]{
  background:
    radial-gradient(circle at top right, rgba(192,136,24,.09), transparent 40%),
    linear-gradient(180deg, rgba(255,255,255,.015), rgba(255,255,255,.01))!important;
  border:1.5px dashed rgba(216,169,58,.55)!important;
  border-radius:18px!important;
  padding:2.6rem 2rem!important;
  transition:all .22s ease!important;
  min-height:180px!important;
}
[data-testid="stFileUploadDropzone"]:hover{
  border-color:rgba(242,200,92,.95)!important;
  background:
    radial-gradient(circle at top right, rgba(192,136,24,.13), transparent 40%),
    linear-gradient(180deg, rgba(255,255,255,.025), rgba(255,255,255,.012))!important;
  box-shadow:0 0 0 1px rgba(216,169,58,.14), 0 20px 60px rgba(0,0,0,.18)!important;
}
[data-testid="stFileUploadDropzone"] p,
[data-testid="stFileUploadDropzone"] span{
  color:var(--txt)!important;
  font-family:var(--ff)!important;
  font-size:.96rem!important;
  font-weight:600!important;
}
[data-testid="stFileUploadDropzone"] small{
  color:var(--txt3)!important;
  font-size:.78rem!important;
}

/* Inputs */
.stSelectbox label,.stRadio label{
  color:var(--txt4)!important;
  font-size:.60rem!important;
  font-weight:600!important;
  text-transform:uppercase!important;
  letter-spacing:.16em!important;
  margin-bottom:.35rem!important;
}
div[data-baseweb="select"] > div{
  background:linear-gradient(180deg,var(--bg4),var(--bg3))!important;
  border:1px solid var(--line2)!important;
  border-radius:12px!important;
  color:var(--txt)!important;
  min-height:46px!important;
  box-shadow:none!important;
}
div[data-baseweb="select"] *{
  color:var(--txt)!important;
  font-family:var(--fi)!important;
}
[data-baseweb="radio"]{
  background:transparent!important;
}
[data-baseweb="radio"] label{
  background:linear-gradient(180deg,var(--bg4),var(--bg3));
  border:1px solid var(--line2);
  border-radius:12px;
  padding:.55rem .9rem;
  margin-right:.45rem;
}
[data-baseweb="radio"] label:hover{
  border-color:var(--line3);
}

/* Buttons */
.stButton>button{
  background:linear-gradient(135deg,var(--gold2),var(--gold))!important;
  color:#06111d!important;
  border:none!important;
  border-radius:12px!important;
  padding:.72rem 1.2rem!important;
  font-family:var(--ff)!important;
  font-weight:700!important;
  font-size:.76rem!important;
  letter-spacing:.08em!important;
  text-transform:uppercase!important;
  box-shadow:0 10px 24px rgba(192,136,24,.22)!important;
}
.stButton>button:hover{
  transform:translateY(-1px)!important;
  box-shadow:0 16px 34px rgba(192,136,24,.28)!important;
  background:linear-gradient(135deg,var(--gold3),var(--gold2))!important;
}

/* Expander */
.streamlit-expanderHeader{
  background:linear-gradient(180deg,rgba(255,255,255,.02),rgba(255,255,255,.01))!important;
  border:1px solid var(--line2)!important;
  border-radius:16px!important;
  color:var(--txt)!important;
  font-family:var(--ff)!important;
  font-size:.90rem!important;
  font-weight:600!important;
}
.streamlit-expanderHeader:hover{
  border-color:var(--line3)!important;
  background:linear-gradient(180deg,rgba(255,255,255,.03),rgba(255,255,255,.015))!important;
}
.streamlit-expanderContent{
  background:transparent!important;
  border:none!important;
  padding:1.2rem 0 .35rem!important;
}

.stAlert{
  border-radius:12px!important;
  border-left-width:3px!important;
}

@keyframes fadeUp{
  from{opacity:0;transform:translateY(14px)}
  to{opacity:1;transform:translateY(0)}
}
@keyframes fadeDown{
  from{opacity:0;transform:translateY(-10px)}
  to{opacity:1;transform:translateY(0)}
}
@keyframes glowLine{
  from{width:0;opacity:0}
  to{width:100%;opacity:1}
}
@keyframes popIn{
  0%{opacity:0;transform:scale(.97)}
  100%{opacity:1;transform:scale(1)}
}

/* Utility classes used by HTML */
.app-hero{
  position:relative;
  overflow:hidden;
  margin:.2rem 0 1.25rem;
  border:1px solid rgba(255,255,255,.05);
  border-radius:24px;
  background:
    radial-gradient(700px 260px at 0% 0%, rgba(192,136,24,.12), transparent 60%),
    radial-gradient(600px 240px at 100% 100%, rgba(62,115,218,.09), transparent 58%),
    linear-gradient(180deg, rgba(255,255,255,.02), rgba(255,255,255,.01));
  box-shadow:var(--shadow-lg);
  padding:1.35rem 1.4rem 1.35rem;
  animation:fadeDown .45s cubic-bezier(.22,1,.36,1) both;
}
.app-hero::after{
  content:'';
  position:absolute; inset:0; pointer-events:none;
  background:linear-gradient(90deg, rgba(255,255,255,.05), transparent 18%, transparent 82%, rgba(255,255,255,.03));
  opacity:.35;
}
.hero-grid{
  display:grid;
  grid-template-columns: 1.65fr .95fr;
  gap:1rem;
  align-items:stretch;
}
.hero-kicker{
  font-family:var(--fm);
  font-size:.64rem;
  letter-spacing:.24em;
  text-transform:uppercase;
  color:var(--gold2);
}
.hero-title{
  font-family:var(--ff);
  font-size:clamp(1.7rem,3vw,2.55rem);
  line-height:1.02;
  letter-spacing:-.045em;
  color:var(--txt);
  font-weight:700;
  margin-top:.48rem;
}
.hero-text{
  color:var(--txt3);
  font-size:.92rem;
  max-width:72ch;
  line-height:1.65;
  margin-top:.8rem;
}
.hero-meta{
  display:flex;
  flex-wrap:wrap;
  gap:.6rem;
  margin-top:1rem;
}
.pill{
  display:inline-flex;
  align-items:center;
  gap:.45rem;
  padding:.44rem .72rem;
  border-radius:999px;
  background:rgba(255,255,255,.03);
  border:1px solid rgba(255,255,255,.06);
  color:var(--txt2);
  font-family:var(--fm);
  font-size:.66rem;
  white-space:nowrap;
}
.hero-panel{
  border-radius:20px;
  border:1px solid rgba(255,255,255,.06);
  background:linear-gradient(180deg, rgba(255,255,255,.03), rgba(255,255,255,.012));
  padding:1rem 1rem .95rem;
  display:flex;
  flex-direction:column;
  justify-content:space-between;
  min-height:100%;
}
.hero-panel-title{
  color:var(--txt);
  font-family:var(--ff);
  font-size:.95rem;
  font-weight:600;
  margin-bottom:.7rem;
}
.hero-stat{
  display:grid;
  grid-template-columns:1fr auto;
  gap:.7rem;
  padding:.58rem 0;
  border-top:1px solid rgba(255,255,255,.05);
}
.hero-stat:first-of-type{border-top:none;padding-top:0}
.hero-stat-label{
  color:var(--txt3);
  font-size:.72rem;
  text-transform:uppercase;
  letter-spacing:.13em;
}
.hero-stat-val{
  color:var(--txt);
  font-family:var(--fm);
  font-size:.78rem;
  font-weight:600;
  text-align:right;
}

/* KPI cards */
.kpi-card{
  position:relative;
  overflow:hidden;
  border-radius:18px;
  border:1px solid rgba(255,255,255,.05);
  background:
    linear-gradient(180deg, rgba(255,255,255,.022), rgba(255,255,255,.008)),
    linear-gradient(160deg,var(--bg4),var(--bg3));
  padding:1rem 1rem .95rem;
  min-height:118px;
  box-shadow:var(--shadow-sm);
  animation:fadeUp .38s ease both;
}
.kpi-card::before{
  content:'';
  position:absolute; inset:0;
  background:linear-gradient(145deg,var(--accent-fade),transparent 58%);
  pointer-events:none;
}
.kpi-top{
  display:flex;
  align-items:flex-start;
  justify-content:space-between;
  gap:.8rem;
}
.kpi-label{
  color:var(--txt4);
  font-size:.60rem;
  line-height:1.2;
  font-weight:600;
  text-transform:uppercase;
  letter-spacing:.16em;
}
.kpi-icon{
  width:34px;height:34px;border-radius:12px;
  display:flex;align-items:center;justify-content:center;
  background:rgba(255,255,255,.04);
  border:1px solid rgba(255,255,255,.06);
  color:var(--accent);
  font-size:.9rem;
}
.kpi-value{
  margin-top:.8rem;
  color:var(--txt);
  font-family:var(--ff);
  font-size:clamp(1.05rem,2vw,1.55rem);
  line-height:1.02;
  font-weight:700;
  letter-spacing:-.04em;
  word-break:break-word;
}
.kpi-sub{
  margin-top:.52rem;
  color:var(--accent);
  font-family:var(--fm);
  font-size:.66rem;
  letter-spacing:.05em;
}

/* Section card */
.section-shell{
  margin-top:1.45rem;
  border-radius:22px;
  border:1px solid rgba(255,255,255,.05);
  background:
    linear-gradient(180deg, rgba(255,255,255,.018), rgba(255,255,255,.008));
  box-shadow:var(--shadow-md);
  overflow:hidden;
}
.section-head{
  display:grid;
  grid-template-columns:auto 1fr auto;
  gap:.9rem;
  align-items:center;
  padding:1rem 1.05rem .95rem;
  border-bottom:1px solid rgba(255,255,255,.05);
  background:linear-gradient(180deg, rgba(255,255,255,.012), transparent);
}
.section-icon{
  width:42px;height:42px;border-radius:14px;
  display:flex;align-items:center;justify-content:center;
  background:var(--accent-box);
  border:1px solid var(--accent-line);
  font-size:1rem;
}
.section-title{
  color:var(--txt);
  font-family:var(--ff);
  font-size:1.02rem;
  font-weight:600;
  line-height:1.05;
}
.section-sub{
  color:var(--txt3);
  font-family:var(--fm);
  font-size:.62rem;
  letter-spacing:.06em;
  margin-top:.28rem;
}
.section-badge{
  white-space:nowrap;
  padding:.46rem .8rem;
  border-radius:999px;
  border:1px solid rgba(255,255,255,.07);
  background:rgba(255,255,255,.03);
  color:var(--txt2);
  font-family:var(--fm);
  font-size:.64rem;
}

/* Micro summary */
.microline{
  margin:.15rem 0 .95rem;
  color:var(--txt4);
  font-family:var(--fm);
  font-size:.64rem;
  line-height:1.9;
}
.microline b{color:var(--txt2)}

/* Empty blocks */
.empty-block{
  border-radius:16px;
  padding:1.2rem 1rem;
  text-align:center;
  border:1px dashed rgba(255,255,255,.08);
  background:linear-gradient(180deg, rgba(255,255,255,.018), rgba(255,255,255,.008));
  color:var(--txt3);
  font-family:var(--ff);
  font-size:.88rem;
  font-weight:500;
}
.upload-note{
  height:100%;
  border-radius:18px;
  border:1px solid rgba(255,255,255,.06);
  background:linear-gradient(180deg, rgba(255,255,255,.02), rgba(255,255,255,.008));
  padding:1rem 1rem .9rem;
}
.upload-note h4{
  color:var(--txt);
  font-family:var(--ff);
  font-size:.92rem;
  font-weight:600;
  margin:0 0 .75rem;
}
.upload-note ul{
  list-style:none;
  margin:0;padding:0;
  display:grid;gap:.68rem;
}
.upload-note li{
  display:grid;
  grid-template-columns:auto 1fr;
  gap:.58rem;
  align-items:start;
  color:var(--txt2);
  font-size:.79rem;
  line-height:1.55;
}
.upload-note .dot{
  width:9px;height:9px;border-radius:999px;
  background:linear-gradient(180deg,var(--gold3),var(--gold));
  margin-top:.35rem;
  box-shadow:0 0 0 4px rgba(192,136,24,.10);
}
.upload-note b{color:var(--txt)}

.status-ok{color:#6ee48f}
.status-hp{color:#ef9f49}
.status-od{color:#f16666}
.status-nd{color:#79a0ef}

.success-banner{
  margin-top:1rem;
  border-radius:16px;
  padding:.95rem 1rem;
  border:1px solid rgba(17,168,91,.25);
  background:linear-gradient(135deg, rgba(17,168,91,.14), rgba(17,168,91,.05));
  color:#8df0b3;
  font-family:var(--ff);
  font-size:.88rem;
  font-weight:500;
  box-shadow:var(--shadow-sm);
  animation:popIn .35s ease both;
}
.success-banner strong{color:#c7ffd8}

/* Table shells */
.table-shell{
  overflow-x:auto;
  border-radius:18px;
  border:1px solid rgba(255,255,255,.05);
  background:linear-gradient(180deg, rgba(255,255,255,.012), rgba(255,255,255,.004));
  box-shadow:var(--shadow-sm);
}
.matrix-table{
  width:100%;
  border-collapse:separate;
  border-spacing:0;
  min-width:980px;
}
.matrix-table thead th{
  position:sticky;
  top:0;
  z-index:2;
}
.filters-caption{
  color:var(--txt4);
  font-family:var(--fm);
  font-size:.60rem;
  text-transform:uppercase;
  letter-spacing:.18em;
  margin-bottom:.65rem;
}

/* Small helpers */
.footer-note{
  margin-top:1rem;
  text-align:right;
  color:var(--txt5);
  font-family:var(--fm);
  font-size:.58rem;
  letter-spacing:.08em;
}

@media (max-width: 1200px){
  .hero-grid{grid-template-columns:1fr}
}
@media (max-width: 900px){
  .block-container{padding:1rem .75rem 4rem!important}
  .section-head{grid-template-columns:auto 1fr}
  .section-badge{grid-column:1 / -1; justify-self:start}
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  STATUS COLOUR PALETTE  — BACKEND/OUTPUT CONTRACT PRESERVED
# ══════════════════════════════════════════════════════════════════════════
_SC = {
    "OVERDUE":       {"row":"#180607","tag_bg":"#2a0b0d","tag_fg":"#ff6d74","comp":"#ff8b92",
                      "dim":"#9b5b60","num":"#ff8d94","bar_f":"#ea4c55","bar_e":"#351013","bord":"#381215"},
    "HIGH PRIORITY": {"row":"#1a1106","tag_bg":"#2c1b08","tag_fg":"#ffb054","comp":"#ffc070",
                      "dim":"#a07d54","num":"#ffc070","bar_f":"#dd8426","bar_e":"#34210e","bord":"#3b250f"},
    "OK":            {"row":"#07140d","tag_bg":"#0d2417","tag_fg":"#56d68a","comp":"#73e5a0",
                      "dim":"#5b8f70","num":"#73e5a0","bar_f":"#11a85b","bar_e":"#0d2617","bord":"#11301d"},
    "NO DATA":       {"row":"#09101a","tag_bg":"#0d1726","tag_fg":"#7ea2ef","comp":"#92b2f7",
                      "dim":"#6d83a5","num":"#92b2f7","bar_f":"#3e73da","bar_e":"#101b2d","bord":"#16243a"},
}
_ORD = {"OVERDUE":0,"HIGH PRIORITY":1,"OK":2,"NO DATA":3}


# ══════════════════════════════════════════════════════════════════════════
#  CONVERSION  (.doc → .docx via LibreOffice)  — BACKEND UNCHANGED
# ══════════════════════════════════════════════════════════════════════════
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found — packages.txt must contain: libreoffice")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as t:
        t.write(raw); src = t.name
    outdir = tempfile.mkdtemp(prefix="lo_")
    out = os.path.join(outdir, Path(src).stem + ".docx")
    pf  = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        r = subprocess.run(
            [soffice, "--headless", "--norestore", "--nofirststartwizard",
             f"-env:UserInstallation={pf}", "--convert-to", "docx", src, "--outdir", outdir],
            capture_output=True, timeout=120)
        if not os.path.exists(out):
            raise RuntimeError(r.stderr.decode("utf-8","ignore")[:400])
        with open(out,"rb") as f: return f.read()
    finally:
        for p in [src,out]:
            try:
                if os.path.exists(p): os.unlink(p)
            except Exception: pass
        shutil.rmtree(outdir, ignore_errors=True)


# ══════════════════════════════════════════════════════════════════════════
#  GRID BUILDERS  — BACKEND UNCHANGED
# ══════════════════════════════════════════════════════════════════════════
def _raw_grid(table) -> List[List[str]]:
    grid=[]; mc=0
    for row in table.rows:
        cells=[]
        for cell in row.cells:
            raw=re.sub(r'[\x0b\r]','\n',cell.text).replace('\x07','')
            lines=[ln.replace('\xa0',' ').replace('\t',' ').strip() for ln in raw.split('\n') if ln.strip()]
            cells.append(lines[0] if lines else '')
        mc=max(mc,len(cells)); grid.append(cells)
    for row in grid:
        while len(row)<mc: row.append('')
    return grid

def _dedup_grid(table) -> List[List[str]]:
    grid=[]; mc=0
    for row in table.rows:
        cells=[]; prior=None
        for cell in row.cells:
            if cell._tc is prior: continue
            prior=cell._tc
            raw=re.sub(r'[\x0b\r]','\n',cell.text).replace('\x07','')
            lines=[ln.replace('\xa0',' ').replace('\t',' ').strip() for ln in raw.split('\n') if ln.strip()]
            cells.append(lines[0] if lines else '')
        mc=max(mc,len(cells)); grid.append(cells)
    for row in grid:
        while len(row)<mc: row.append('')
    return grid


# ══════════════════════════════════════════════════════════════════════════
#  TEXT HELPERS  — BACKEND UNCHANGED
# ══════════════════════════════════════════════════════════════════════════
def _fl(t: Any) -> str:
    raw=str(t or "").replace('\x07','').replace('\xa0',' ').replace('\t',' ')
    for p in re.split(r'[\r\n\x0b]+',raw):
        s=re.sub(r'\s+',' ',p).strip()
        if s: return s
    return ''

def _clean_name(t: Any) -> str:
    s=_fl(t)
    s=re.sub(r'(?i)^MV\s+','',s)
    s=re.sub(r'(?i)ALEXIS\s*Date?','',s)
    s=re.sub(r'(?i)Page\s*\d+\s*of\s*\d+','',s)
    return re.sub(r'  +',' ',s).strip(" -:")

def _parse_number(t: Any) -> float:
    s=_fl(t).strip().upper()
    if not s or s in ('','-','N/A','CENTRAL','COOLER'): return 0.0
    s=s.replace('[','').replace(']','')
    if any(w in s for w in ('MONTH','YEAR','WEEK','DAY','OBS','NO RECORD','BASED ON','OBSERVATION')): return 0.0
    m=re.search(r'\d[\d,\.]*',s)
    if not m: return 0.0
    b=m.group(); sep=max(b.rfind('.'),b.rfind(','))
    if sep>0 and len(b)-sep==4: b=re.sub(r'[,.]','',b)
    elif sep>0: b=re.sub(r'[,.]','',b[:sep])
    else: b=re.sub(r'[,.]','',b)
    try: return float(b)
    except: return 0.0

def _parse_date(t: Any) -> str:
    s=_fl(t).strip()
    if not s: return ''
    s=s.replace('[','').replace(']','').strip()
    if s in ('-','1','2','N/A','n/a','NA','Central','CENTRAL','COOLER',
             'NO RECORD','NOT WORKING','N.A.'): return ''
    if re.fullmatch(r'^\d+$',s): return ''
    if len(s)>32: return ''
    return s if re.search(r'[A-Za-z0-9/]',s) else ''

def _status(hrs: float, period: float) -> str:
    if hrs<=0 or period<=0: return 'NO DATA'
    r=hrs/period
    if r>=1.0: return 'OVERDUE'
    if r>=0.8: return 'HIGH PRIORITY'
    return 'OK'

def _pct(hrs: float, period: float) -> float:
    return round(hrs/period,4) if hrs and period else 0.0

def _is_comp(n: str) -> bool:
    u=_fl(n).upper()
    if not u or len(u)<2: return False
    BAD=('DESCRIPTION','REMARKS','COMPONENT','PERIODICITY','PERIODICTLY','DATE OF LAST',
         'RUNNING HOURS','MAIN ENGINE','AUX. ENGINE','TYPE:','1-DATE OF LAST',
         'TOTAL RUNNING','THIS MONTH','CYL. NO','NOTE 1','BASED ON','SERIAL NR',
         'HOURS THIS MONTH','AUX. ENGINE MAKER')
    if any(b in u for b in BAD): return False
    if re.fullmatch(r'[\d./ ,:\-\[\]()]+',u): return False
    return bool(re.search(r'[A-Za-z]',u))

def _mk(cat,eng,unit,nm,per,dt,hrs) -> Dict:
    return {'category':cat,'engine_label':eng,'unit':unit,'description':nm,
            'periodicity':per,'last_oh_date':dt,'hrs_since':hrs,
            'pct_used':_pct(hrs,per),'status':_status(hrs,per)}


# ══════════════════════════════════════════════════════════════════════════
#  ME PARSER  — BACKEND UNCHANGED
# ══════════════════════════════════════════════════════════════════════════
def _parse_me(grid: List[List[str]]) -> List[Dict]:
    if not grid: return []
    if not any('MAIN ENGINE' in (_fl(grid[r][c]).upper() if c<len(grid[r]) else '')
               for r in range(min(3,len(grid))) for c in range(min(15,len(grid[r])))): return []

    rem_col=None
    for r in range(min(5,len(grid))):
        for ci,txt in enumerate(grid[r]):
            if 'REMARK' in _fl(txt).upper() and ci>3: rem_col=ci; break
        if rem_col: break
    if rem_col is None: rem_col=12

    MARKER=2; PERIOD=1; FIRST=3
    actual_cyls=max(1,min(8,rem_col-FIRST))

    end=len(grid)
    for r,row in enumerate(grid):
        j=' '.join(_fl(x) for x in row).upper()
        if 'NOTE 1' in j or 'TURBOCHARGER' in j or 'AUX. ENGINE MAKER' in j: end=r; break

    result=[]; r=0
    while r<end-1:
        nm    =_clean_name(grid[r][0] if grid[r] else '')
        period=_parse_number(grid[r][PERIOD] if PERIOD<len(grid[r]) else '')
        marker=_fl(grid[r][MARKER] if MARKER<len(grid[r]) else '').strip()
        if _is_comp(nm) and marker=='1':
            nxt=grid[r+1] if r+1<len(grid) else []
            for cyl in range(1,actual_cyls+1):
                ci=FIRST+cyl-1
                d=_parse_date(grid[r][ci] if ci<len(grid[r]) else '')
                h=_parse_number(nxt[ci] if ci<len(nxt) else '')
                if d or h>0: result.append(_mk('MAIN_ENGINE','ME',f'Cyl {cyl}',nm,period,d,h))
            r+=2
        else: r+=1
    return result


# ══════════════════════════════════════════════════════════════════════════
#  AUX PARSER  — BACKEND UNCHANGED
# ══════════════════════════════════════════════════════════════════════════
def _find_aux_groups(grid: List[List[str]]) -> Tuple[int, List[Tuple]]:
    dr=None
    for i,row in enumerate(grid):
        rt=' | '.join(_fl(c) for c in row).upper()
        if 'DESCRIPTION' in rt and ('PERIODICTLY' in rt or 'PERIODICITY' in rt): dr=i; break
    if dr is None: return -1,[]
    nums=[(c,int(_fl(grid[dr][c])))
          for c in range(2,len(grid[dr]))
          if re.fullmatch(r'\d+',_fl(grid[dr][c]))]
    if nums:
        starts=[c for c,n in nums if n==1]
        groups=[]
        for i,s in enumerate(starts[:3]):
            ds=s+1
            de=(starts[i+1]+1) if i+1<len(starts) else len(grid[dr])
            groups.append((['AUX-1','AUX-2','AUX-3'][i],ds,de))
        if groups: return dr,groups
    return dr,[]

def _parse_aux(grid: List[List[str]]) -> List[Dict]:
    if not grid: return []
    dr,groups=_find_aux_groups(grid)
    if dr<0 or not groups: return []
    result=[]; r=dr+1
    while r<len(grid)-1:
        nm    =_clean_name(grid[r][0] if grid[r] else '')
        period=_parse_number(grid[r][1] if len(grid[r])>1 else '')
        marker=_fl(grid[r][2] if len(grid[r])>2 else '').strip()
        if _is_comp(nm) and marker=='1':
            nxt=grid[r+1] if r+1<len(grid) else []
            for eng,start,end in groups:
                for ci_idx,ci in enumerate(range(start,min(end,len(grid[r])))):
                    d=_parse_date(grid[r][ci] if ci<len(grid[r]) else '')
                    h=_parse_number(nxt[ci] if ci<len(nxt) else '')
                    if d or h>0:
                        result.append(_mk('AUX_ENGINE',eng,f'Cyl {ci_idx+1}',nm,period,d,h))
            r+=2
        else: r+=1
    return result


# ══════════════════════════════════════════════════════════════════════════
#  OE PARSER  — BACKEND UNCHANGED
# ══════════════════════════════════════════════════════════════════════════
_OE_HEADER = {
    'TURBOCHARGER','AUXILIARY BOILER','COOLERS','EXH GAS BOILER','EXH GAS  BOILER',
    'A/C & REFR. COMPRESSORS','MAIN AIR COMPRESSORS','PERIODICTLY','PERIODICITY',
    'DATE OF LAST INSPECTION','DATE OF LAST O/H','RUN HRS','DATE OF LAST CLEANING',
    'DATE','DESCRIPTION','',
}

def _is_oe_comp(n: str) -> bool:
    u=_fl(n).upper().strip()
    if not u or len(u)<2 or u in _OE_HEADER: return False
    if re.fullmatch(r'[\d./ ,:\-\[\]()]+',u): return False
    return bool(re.search(r'[A-Za-z]',u))

def _parse_oe(grid: List[List[str]]) -> List[Dict]:
    rows=[]
    for row in grid:
        r=row
        def gc(i,_r=r): return _fl(_r[i]) if i<len(_r) else ''

        da=_clean_name(gc(0))
        if _is_oe_comp(da):
            per=_parse_number(gc(1)); dt=_parse_date(gc(2)); hrs=_parse_number(gc(3))
            if dt or hrs>0 or per>0:
                rows.append({'section':'Turbocharger / Aux Boiler','description':da,
                             'periodicity':per,'last_date':dt,'run_hrs':hrs})

        db=_clean_name(gc(5))
        if _is_oe_comp(db):
            dt=_parse_date(gc(6)); hrs=_parse_number(gc(7))
            if dt or hrs>0:
                rows.append({'section':'Coolers / Exh Gas Boiler','description':db,
                             'periodicity':0,'last_date':dt,'run_hrs':hrs})

        dc=_clean_name(gc(10))
        if _is_oe_comp(dc):
            dt=_parse_date(gc(11)); hrs=_parse_number(gc(12))
            if dt or hrs>0:
                rows.append({'section':'A/C & Compressors','description':dc,
                             'periodicity':0,'last_date':dt,'run_hrs':hrs})
    return rows


# ══════════════════════════════════════════════════════════════════════════
#  DG PARSER  — BACKEND UNCHANGED
# ══════════════════════════════════════════════════════════════════════════
_DG_SKIP = {'DESCRIPTION','PERIODICTLY','PERIODICITY','D/G NO1','D/G NO2','D/G NO3',''}

def _parse_dg(grid: List[List[str]]) -> List[Dict]:
    rows=[]; r=0
    while r<len(grid)-1:
        r1=grid[r]; r2=grid[r+1] if r+1<len(grid) else []

        def gc1(i,_row=r1): return _fl(_row[i]) if i<len(_row) else ''
        def gc2(i,_row=r2): return _fl(_row[i]) if i<len(_row) else ''

        dl=_clean_name(gc1(0))
        if _is_oe_comp(dl) and dl.upper() not in _DG_SKIP and gc1(2)=='1':
            per=_parse_number(gc1(1))
            for gi,gl in enumerate(['D/G 1','D/G 2','D/G 3']):
                dt=_parse_date(gc1(3+gi)); hrs=_parse_number(gc2(3+gi))
                if dt or hrs>0:
                    rows.append({'section':'D/G Equipment','description':dl,'engine_label':gl,
                                 'periodicity':per,'last_date':dt,'run_hrs':hrs,
                                 'status':_status(hrs,per)})

        dr=_clean_name(gc1(9))
        if _is_oe_comp(dr) and dr.upper() not in _DG_SKIP and gc1(11)=='1':
            per=_parse_number(gc1(10))
            for gi,gl in enumerate(['D/G 1','D/G 2','D/G 3']):
                dt=_parse_date(gc1(12+gi)); hrs=_parse_number(gc2(12+gi))
                if dt or hrs>0:
                    rows.append({'section':'D/G Equipment','description':dr,'engine_label':gl,
                                 'periodicity':per,'last_date':dt,'run_hrs':hrs,
                                 'status':_status(hrs,per)})
        r+=1
    return rows


# ══════════════════════════════════════════════════════════════════════════
#  TEXT FALLBACK  — BACKEND UNCHANGED
# ══════════════════════════════════════════════════════════════════════════
def _lines_from_doc(doc) -> List[str]:
    lines=[]
    for p in doc.paragraphs:
        t=_fl(p.text)
        if t: lines.append(t)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t=_fl(cell.text)
                if t: lines.append(t)
    return lines

def _between(lines,starts,ends):
    s=e=None
    for i,ln in enumerate(lines):
        u=ln.upper()
        if s is None and any(m in u for m in starts): s=i; continue
        if s is not None and any(m in u for m in ends): e=i; break
    return lines[s:e] if s is not None else []

def _parse_me_text(lines: List[str]) -> List[Dict]:
    rows=[]
    seg=_between(lines,['MAIN ENGINE','CYL. NO.1'],['NOTE 1','TURBOCHARGER','AUX. ENGINE MAKER / TYPE'])
    if not seg: return rows
    data=[x for x in seg if _fl(x)]; i=0
    while i<len(data):
        nm=_clean_name(data[i])
        if _is_comp(nm):
            period=_parse_number(data[i+1]) if i+1<len(data) else 0.0
            if _fl(data[i+2] if i+2<len(data) else '')=='1':
                dates=[]; j=i+3
                while j<len(data) and len(dates)<8 and _fl(data[j])!='2':
                    dates.append(_fl(data[j])); j+=1
                if j<len(data) and _fl(data[j])=='2':
                    j+=1; hv=[]
                    while j<len(data) and len(hv)<8:
                        if _is_comp(_fl(data[j])): break
                        hv.append(_fl(data[j])); j+=1
                    for k in range(8):
                        d=_parse_date(dates[k]) if k<len(dates) else ''
                        h=_parse_number(hv[k]) if k<len(hv) else 0.0
                        if d or h>0: rows.append(_mk('MAIN_ENGINE','ME',f'Cyl {k+1}',nm,period,d,h))
                    i=j; continue
        i+=1
    return rows

def _parse_aux_text(lines: List[str]) -> List[Dict]:
    rows=[]
    seg=_between(lines,['AUX. ENGINE MAKER / TYPE','AUX. ENGINE NO.1'],
                 ['D/G NO1','TURBOCHARGER (2)','1ST COPY'])
    if not seg:
        seg=_between(lines,['AUX. ENGINE NO.1'],['D/G NO1','TURBOCHARGER (2)','1ST COPY'])
    si=None
    for i,ln in enumerate(seg):
        if 'DESCRIPTION' in ln.upper() and ('PERIODICTLY' in ln.upper() or 'PERIODICITY' in ln.upper()):
            si=i; break
    if si is None: return rows
    data=[_fl(x) for x in seg[si+1:] if _fl(x)]; i=0
    while i<len(data):
        nm=_clean_name(data[i])
        if _is_comp(nm):
            period=_parse_number(data[i+1]) if i+1<len(data) else 0.0
            if _fl(data[i+2] if i+2<len(data) else '')=='1':
                dates=[]; j=i+3
                while j<len(data) and len(dates)<18 and _fl(data[j])!='2':
                    if _is_comp(_fl(data[j])): break
                    dates.append(_fl(data[j])); j+=1
                if j<len(data) and _fl(data[j])=='2':
                    j+=1; hv=[]
                    while j<len(data) and len(hv)<18:
                        if _is_comp(_fl(data[j])): break
                        hv.append(_fl(data[j])); j+=1
                    for ei,elbl in enumerate(['AUX-1','AUX-2','AUX-3']):
                        for cyl in range(6):
                            idx=ei*6+cyl
                            d=_parse_date(dates[idx]) if idx<len(dates) else ''
                            h=_parse_number(hv[idx]) if idx<len(hv) else 0.0
                            if d or h>0: rows.append(_mk('AUX_ENGINE',elbl,f'Cyl {cyl+1}',nm,period,d,h))
                    i=j; continue
        i+=1
    return rows


# ══════════════════════════════════════════════════════════════════════════
#  DEDUPE + MASTER PARSE  — BACKEND UNCHANGED
# ══════════════════════════════════════════════════════════════════════════
def _dedupe(records: List[Dict]) -> List[Dict]:
    best={}
    for r in records:
        key=(r.get('category'),r.get('description'),r.get('engine_label'),r.get('unit'))
        score=(1 if r.get('last_oh_date') else 0)+(1 if r.get('hrs_since',0)>0 else 0)
        prev=best.get(key)
        if prev is None or score>prev[0]: best[key]=(score,r)
    return [v[1] for v in best.values()]

def parse_docx(docx_bytes: bytes) -> Dict:
    from docx import Document
    warns=[]
    with tempfile.NamedTemporaryFile(suffix='.docx',delete=False) as t:
        t.write(docx_bytes); tp=t.name
    try:    doc=Document(tp)
    except Exception as e: raise ValueError(f"Cannot open docx: {e}")
    finally:
        try: os.unlink(tp)
        except Exception: pass

    if not doc.tables: raise ValueError("No tables found — is this a TEC-004 report?")

    vn='UNKNOWN'; rd=None
    for para in doc.paragraphs:
        txt=para.text.strip()
        if not txt: continue
        if m:=re.search(r"Vessel[\u2019\u2018']?s?\s+Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+?)(?:\s{2,}|\t|Date:|$)",txt,re.I):
            vn=_clean_name(m.group(1))
        if m:=re.search(r"Date\s*:\s*(.+)",txt,re.I):
            rd=_parse_date(m.group(1).strip())
        if vn!='UNKNOWN' and rd: break
    if vn=='UNKNOWN': warns.append("Could not extract vessel name.")

    mt=mo=None
    me_g:List[Dict]=[]; aux_g:List[Dict]=[]
    oe_rows:List[Dict]=[]; dg_rows:List[Dict]=[]

    for table in doc.tables:
        rg=_raw_grid(table); dg=_dedup_grid(table)
        full=' '.join(_fl(c) for row in rg[:3] for c in row).upper()

        if mt is None:
            for row in rg[:4]:
                line=' '.join(str(x) for x in row)
                if mt is None:
                    if m:=re.search(r'Total Running Hours[\s:ǀ|]+([\d,]+)',line,re.I):
                        mt=int(_parse_number(m.group(1)))
                if mo is None:
                    if m:=re.search(r'This Month[\s:]+([\d,]+)',line,re.I):
                        mo=int(_parse_number(m.group(1)))

        me_g.extend(_parse_me(rg))
        aux_g.extend(_parse_aux(dg))
        if 'TURBOCHARGER' in full and 'A/C & REFR' in full and 'COOLERS' in full:
            oe_rows.extend(_parse_oe(dg))
        if 'D/G NO' in full.replace(' ','').replace('.',''):
            dg_rows.extend(_parse_dg(dg))

    all_lines=_lines_from_doc(doc)
    me=_dedupe(me_g+_parse_me_text(all_lines))
    aux=_dedupe(aux_g+_parse_aux_text(all_lines))
    oe=_dedupe(oe_rows)

    if not me and not aux: warns.append("No components extracted.")

    return {
        'vessel_name':vn,'report_date':rd,
        'me_total_hrs':mt,'me_this_month':mo,
        'me':me,'aux':aux,'oe':oe,'dg':dg_rows,
        'warnings':warns,
        'parsed_at':datetime.utcnow().isoformat(),
    }


# ══════════════════════════════════════════════════════════════════════════
#  FRONTEND HELPERS  — REVISED
# ══════════════════════════════════════════════════════════════════════════
def _html_escape(v: Any) -> str:
    s = str(v if v is not None else "")
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace('"', "&quot;"))

def _cyl_n(u: str) -> int:
    m=re.search(r'\d+',str(u)); return int(m.group()) if m else 999

def _sort_recs(records: List[Dict], mode: str) -> List[Dict]:
    if mode=='matrix':
        return sorted(records,key=lambda c:(c['description'].upper(),c.get('engine_label',''),_cyl_n(c.get('unit',''))))
    return sorted(records,key=lambda c:(_ORD.get(c['status'],4),-(c.get('pct_used') or 0)))

def _empty_html(msg: str, accent: str = "#6f88a7") -> str:
    return (
        f'<div class="empty-block" style="border-color:rgba(255,255,255,.08);">'
        f'<div style="font-size:1.05rem;margin-bottom:.35rem;color:{accent}">●</div>'
        f'{_html_escape(msg)}'
        f'</div>'
    )

def _kpi(val: Any, lbl: str, accent: str, icon: str='•', sub: str='') -> str:
    r,g,b=int(accent[1:3],16),int(accent[3:5],16),int(accent[5:7],16)
    val = _html_escape(val)
    lbl = _html_escape(lbl)
    sub = _html_escape(sub)
    return (
        f'<div class="kpi-card" style="--accent:{accent};--accent-fade:rgba({r},{g},{b},.12)">'
        f'  <div class="kpi-top">'
        f'    <div class="kpi-label">{lbl}</div>'
        f'    <div class="kpi-icon">{icon}</div>'
        f'  </div>'
        f'  <div class="kpi-value">{val}</div>'
        f'  <div class="kpi-sub">{sub}</div>'
        f'</div>'
    )

def _section(icon: str, title: str, sub: str, badge: str, accent: str) -> str:
    r,g,b=int(accent[1:3],16),int(accent[3:5],16),int(accent[5:7],16)
    return (
        f'<div class="section-head" style="--accent-box:rgba({r},{g},{b},.10);--accent-line:rgba({r},{g},{b},.28)">'
        f'  <div class="section-icon">{icon}</div>'
        f'  <div>'
        f'    <div class="section-title">{_html_escape(title)}</div>'
        f'    <div class="section-sub">{_html_escape(sub)}</div>'
        f'  </div>'
        f'  <div class="section-badge">{_html_escape(badge)}</div>'
        f'</div>'
    )

def _rec_line(recs: List[Dict]) -> str:
    n=len(recs)
    od=sum(1 for c in recs if c.get('status')=='OVERDUE')
    hp=sum(1 for c in recs if c.get('status')=='HIGH PRIORITY')
    ok=sum(1 for c in recs if c.get('status')=='OK')
    nd=sum(1 for c in recs if c.get('status')=='NO DATA')
    parts=[f'<b>{n}</b> records']
    if od: parts.append(f'<span class="status-od">{od} overdue</span>')
    if hp: parts.append(f'<span class="status-hp">{hp} high priority</span>')
    if ok: parts.append(f'<span class="status-ok">{ok} OK</span>')
    if nd: parts.append(f'<span class="status-nd">{nd} no data</span>')
    return f'<div class="microline">{" · ".join(parts)}</div>'

def _matrix_html(records: List[Dict], mode: str='matrix') -> str:
    if not records: return ''
    HEADS=['Status','Component','Engine','Unit','Periodicity','Last O/H','Hrs Since','Usage']
    th = (
        "padding:12px 14px;background:#0c1526;color:#647d9a;"
        "font-family:Inter,sans-serif;font-size:.60rem;font-weight:700;"
        "text-transform:uppercase;letter-spacing:.16em;text-align:left;"
        "border-bottom:1px solid rgba(255,255,255,.06);white-space:nowrap"
    )
    header='<tr>' + ''.join(f'<th style="{th}">{h}</th>' for h in HEADS) + '</tr>'

    body=''
    for rec in _sort_recs(records,mode):
        s=str(rec.get('status','NO DATA'))
        c=_SC.get(s,_SC['NO DATA'])
        pct=float(rec.get('pct_used') or 0)
        hrs=float(rec.get('hrs_since') or 0)
        per=float(rec.get('periodicity') or 0)
        pct_s=f"{pct*100:.1f}%" if pct>0 else '—'
        hrs_s=f"{int(hrs):,}" if hrs>0 else '—'
        per_s=f"{int(per):,}" if per>0 else '—'
        dt_s=_html_escape(str(rec.get('last_oh_date') or '—') or '—')
        bw=max(0,min(100,pct*100))
        desc=_html_escape(rec.get('description',''))
        eng=_html_escape(rec.get('engine_label',''))
        unit=_html_escape(rec.get('unit',''))

        tag=(f'<span style="display:inline-flex;align-items:center;justify-content:center;'
             f'padding:4px 10px;border-radius:999px;background:{c["tag_bg"]};color:{c["tag_fg"]};'
             f'font-family:JetBrains Mono,monospace;font-size:.62rem;font-weight:700;'
             f'letter-spacing:.04em;white-space:nowrap">{_html_escape(s)}</span>')

        bar=(
            f'<div style="display:grid;grid-template-columns:1fr auto;gap:10px;align-items:center;min-width:140px">'
            f'  <div style="height:6px;background:{c["bar_e"]};border-radius:999px;overflow:hidden">'
            f'    <div style="width:{bw:.1f}%;height:100%;background:{c["bar_f"]};border-radius:999px"></div>'
            f'  </div>'
            f'  <span style="font-family:JetBrains Mono,monospace;font-size:.69rem;font-weight:700;'
            f'         color:{c["bar_f"]};min-width:46px;text-align:right">{pct_s}</span>'
            f'</div>'
        )

        def td(v, fg, fw='400', align='left', ff="'Inter',sans-serif", fs='.80rem', mw=''):
            mw_s=f'max-width:{mw};overflow:hidden;text-overflow:ellipsis;' if mw else ''
            return (f'<td style="padding:11px 14px;color:{fg};font-family:{ff};font-size:{fs};'
                    f'font-weight:{fw};text-align:{align};white-space:nowrap;{mw_s}">{v}</td>')

        body += (
            f'<tr style="background:{c["row"]};border-bottom:1px solid {c["bord"]};">'
            f'<td style="padding:11px 14px">{tag}</td>'
            + td(desc,c['comp'],'600','left',"'Space Grotesk',sans-serif",'.83rem','340px')
            + td(eng,c['dim'],'500')
            + td(unit,c['dim'],'500','left',"'JetBrains Mono',monospace",'.74rem')
            + td(per_s,c['dim'],'500','right',"'JetBrains Mono',monospace",'.74rem')
            + td(dt_s,c['dim'],'500','left',"'JetBrains Mono',monospace",'.74rem')
            + td(hrs_s,c['num'],'700','right',"'JetBrains Mono',monospace",'.76rem')
            + f'<td style="padding:11px 14px">{bar}</td>'
            f'</tr>'
        )

    return (
        f'<div class="table-shell">'
        f'  <table class="matrix-table">'
        f'    <thead>{header}</thead>'
        f'    <tbody>{body}</tbody>'
        f'  </table>'
        f'</div>'
    )

def _oe_html(rows: List[Dict], show_status: bool=False) -> str:
    if not rows: return ''
    HEADS=(['Component','Engine','Period','Last Date','Run Hrs','Status']
           if show_status else ['Component','Section','Period','Last Date','Run Hrs'])
    th = (
        "padding:12px 14px;background:#0c1526;color:#647d9a;"
        "font-family:Inter,sans-serif;font-size:.60rem;font-weight:700;"
        "text-transform:uppercase;letter-spacing:.16em;text-align:left;"
        "border-bottom:1px solid rgba(255,255,255,.06);white-space:nowrap"
    )
    header='<tr>' + ''.join(f'<th style="{th}">{h}</th>' for h in HEADS) + '</tr>'
    body=''

    for row in sorted(rows,key=lambda r:(r.get('description',''),r.get('engine_label',''))):
        bg='#0a1321'; fm='#b7caf0'; fd='#7f96b4'
        if show_status:
            s=str(row.get('status','NO DATA'))
            cc=_SC.get(s,_SC['NO DATA'])
            bg=cc['row']; fm=cc['comp']; fd=cc['dim']

        per=float(row.get('periodicity',0) or 0)
        per_s=f"{int(per):,}" if per>0 else '—'
        dt_s=_html_escape(str(row.get('last_date') or '—') or '—')
        hrs=float(row.get('run_hrs',0) or 0)
        hrs_s=f"{int(hrs):,}" if hrs>0 else '—'
        desc=_html_escape(row.get('description',''))

        st_cell=''
        if show_status:
            s=str(row.get('status','NO DATA')); cc=_SC.get(s,_SC['NO DATA'])
            st_cell=(f'<td style="padding:11px 14px">'
                     f'<span style="display:inline-flex;padding:4px 10px;border-radius:999px;'
                     f'background:{cc["tag_bg"]};color:{cc["tag_fg"]};font-family:JetBrains Mono,monospace;'
                     f'font-size:.62rem;font-weight:700;letter-spacing:.04em">{_html_escape(s)}</span></td>')

        def td(v,fg,ff="'Inter',sans-serif",fw='500',align='left',fs='.80rem'):
            return (f'<td style="padding:11px 14px;color:{fg};font-family:{ff};font-size:{fs};'
                    f'font-weight:{fw};text-align:{align};white-space:nowrap">{v}</td>')

        body += (
            f'<tr style="background:{bg};border-bottom:1px solid rgba(255,255,255,.05)">'
            + td(desc,fm,"'Space Grotesk',sans-serif",'600','left','.82rem')
            + (td(_html_escape(row.get('engine_label','')),fd,"'JetBrains Mono',monospace",'500','left','.75rem') if show_status
               else td(_html_escape(row.get('section','')),fd,"'Inter',sans-serif",'500','left','.79rem'))
            + td(per_s,fd,"'JetBrains Mono',monospace",'500','right','.75rem')
            + td(dt_s,fd,"'JetBrains Mono',monospace",'500','left','.75rem')
            + td(hrs_s,'#8db3ff',"'JetBrains Mono',monospace",'700','right','.76rem')
            + st_cell
            + '</tr>'
        )

    return (
        f'<div class="table-shell">'
        f'  <table class="matrix-table" style="min-width:860px">'
        f'    <thead>{header}</thead>'
        f'    <tbody>{body}</tbody>'
        f'  </table>'
        f'</div>'
    )

def _show(records: List[Dict], mode: str='matrix'):
    if not records:
        st.markdown(_empty_html("No records match the current filter."), unsafe_allow_html=True)
    else:
        st.markdown(_matrix_html(records,mode), unsafe_allow_html=True)

def _hero_html() -> str:
    return """
    <div class="app-hero">
      <div class="hero-grid">
        <div>
          <div class="hero-kicker">Running Hours Management System</div>
          <div class="hero-title">Fleet Running Hours Extraction Matrix</div>
          <div class="hero-text">
            Upload a TEC-004 report to extract Main Engine, Auxiliary Engines,
            Other Equipment, and D/G Equipment into structured running-hours matrices
            with priority classification and cleaner operational review.
          </div>
          <div class="hero-meta">
            <span class="pill">⚓ TEC-004 Parser</span>
            <span class="pill">ME · AUX · OE · D/G</span>
            <span class="pill">Priority Logic: 80% / 100%</span>
          </div>
        </div>
        <div class="hero-panel">
          <div>
            <div class="hero-panel-title">System Profile</div>
            <div class="hero-stat">
              <div class="hero-stat-label">Input</div>
              <div class="hero-stat-val">Legacy .doc report</div>
            </div>
            <div class="hero-stat">
              <div class="hero-stat-label">Conversion</div>
              <div class="hero-stat-val">LibreOffice headless</div>
            </div>
            <div class="hero-stat">
              <div class="hero-stat-label">Backend</div>
              <div class="hero-stat-val">Section-specific parsers</div>
            </div>
            <div class="hero-stat">
              <div class="hero-stat-label">UI</div>
              <div class="hero-stat-val">Pure Streamlit + HTML matrices</div>
            </div>
          </div>
        </div>
      </div>
    </div>
    """

def _section_open():
    st.markdown('<div class="section-shell">', unsafe_allow_html=True)

def _section_close():
    st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════════════════════════════════
if 'data' not in st.session_state:
    st.session_state.data=None


# ══════════════════════════════════════════════════════════════════════════
#  PAGE HEADER  — REVISED
# ══════════════════════════════════════════════════════════════════════════
st.markdown(_hero_html(), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  UPLOAD PANEL  — REVISED
# ══════════════════════════════════════════════════════════════════════════
with st.expander("Upload TEC-004 Report", expanded=(st.session_state.data is None)):
    uc, ic = st.columns([1.7, 1.0], gap='large')

    with uc:
        uploaded=st.file_uploader('Drop TEC-004 .doc',type=['doc'],label_visibility='collapsed')

    with ic:
        st.markdown("""
        <div class="upload-note">
          <h4>Upload Notes</h4>
          <ul>
            <li><span class="dot"></span><span><b>Format:</b> legacy TEC-004 running-hours report in <b>.doc</b> format.</span></li>
            <li><span class="dot"></span><span><b>Parsing:</b> raw-grid Main Engine, dedup Auxiliary, exact-index OE and paired D/G logic.</span></li>
            <li><span class="dot"></span><span><b>Output:</b> structured matrices for ME, AUX-1/2/3, Other Equipment, and D/G Equipment.</span></li>
            <li><span class="dot"></span><span><b>Priority bands:</b> over 100% = Overdue, 80%+ = High Priority.</span></li>
          </ul>
        </div>
        """, unsafe_allow_html=True)

    if uploaded:
        raw=uploaded.read(); fh=hashlib.md5(raw).hexdigest()
        if st.session_state.data is None or st.session_state.data.get('_hash')!=fh:
            with st.spinner('Converting .doc → .docx via LibreOffice…'):
                try: docx=convert_doc_to_docx(raw)
                except Exception as e:
                    st.error(f'Conversion failed: {e}')
                    st.stop()

            with st.spinner('Parsing TEC-004 tables…'):
                try: result=parse_docx(docx)
                except ValueError as e:
                    st.error(f'Parse failed: {e}')
                    st.stop()

            for w in result['warnings']:
                st.warning(f'⚠ {w}')

            if not result['me'] and not result['aux']:
                st.error('No components extracted. Verify this is a TEC-004 report.')
                st.stop()

            result['_hash']=fh
            result['_filename']=uploaded.name
            st.session_state.data=result

            ac=result['me']+result['aux']
            od=sum(1 for c in ac if c['status']=='OVERDUE')
            hp=sum(1 for c in ac if c['status']=='HIGH PRIORITY')

            st.markdown(
                f'<div class="success-banner">✓ <strong>{_html_escape(result["vessel_name"])}</strong> loaded successfully — '
                f'{len(ac)} tracked records · {od} overdue · {hp} high priority</div>',
                unsafe_allow_html=True
            )


# ══════════════════════════════════════════════════════════════════════════
#  EMPTY STATE  — REVISED
# ══════════════════════════════════════════════════════════════════════════
if st.session_state.data is None:
    st.markdown("""
    <div style="display:flex;align-items:center;justify-content:center;min-height:34vh;">
      <div style="width:min(760px,100%);border-radius:24px;padding:2rem 1.4rem;text-align:center;
                  border:1px solid rgba(255,255,255,.06);
                  background:linear-gradient(180deg,rgba(255,255,255,.02),rgba(255,255,255,.008));
                  box-shadow:var(--shadow-md);">
        <div style="font-size:2rem;color:var(--gold2);margin-bottom:.65rem">⚓</div>
        <div style="font-family:var(--ff);font-size:1.15rem;font-weight:600;color:var(--txt);">
          No report loaded
        </div>
        <div style="margin-top:.6rem;color:var(--txt3);font-size:.92rem;line-height:1.7;max-width:58ch;margin-inline:auto;">
          Upload a TEC-004 running-hours report to generate structured matrices for the
          main engine, auxiliary engines, other equipment, and diesel generator equipment.
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════
#  DATA READY
# ══════════════════════════════════════════════════════════════════════════
d=st.session_state.data
me=d['me']; aux=d['aux']; oe=d['oe']; dg=d['dg']
ac=me+aux
n_od=sum(1 for c in ac if c['status']=='OVERDUE')
n_hp=sum(1 for c in ac if c['status']=='HIGH PRIORITY')
n_ok=sum(1 for c in ac if c['status']=='OK')
mt=d.get('me_total_hrs'); mo=d.get('me_this_month')


# ══════════════════════════════════════════════════════════════════════════
#  KPI GRID  — REVISED
# ══════════════════════════════════════════════════════════════════════════
k1,k2,k3,k4 = st.columns(4, gap='large')
for col, payload in zip(
    [k1,k2,k3,k4],
    [
        (d['vessel_name'], 'Vessel', '#3e73da', '⚓', _html_escape(d.get('_filename','report.doc'))),
        (d['report_date'] or '—', 'Report Date', '#d8a93a', '🗓', 'Report metadata'),
        (f"{mt:,}" if mt else '—', 'M/E Total Hours', '#11a85b', 'Σ', 'Main engine cumulative'),
        (f"{mo:,}" if mo else '—', 'M/E This Month', '#11a85b', '△', 'Monthly increment'),
    ]
):
    with col:
        st.markdown(_kpi(*payload), unsafe_allow_html=True)

k5,k6,k7,k8 = st.columns(4, gap='large')
for col, payload in zip(
    [k5,k6,k7,k8],
    [
        (len(me), 'ME Records', '#c08818', '⚙', 'Main engine matrix'),
        (len(aux), 'AUX Records', '#3e73da', '🔩', 'Auxiliary engine matrix'),
        (n_od, 'Overdue', '#d83939', '▲', 'Immediate attention'),
        (n_hp, 'High Priority', '#dd8426', '◆', f'{n_ok} records currently OK'),
    ]
):
    with col:
        st.markdown(_kpi(*payload), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 1 — MAIN ENGINE  — REVISED FRONTEND
# ══════════════════════════════════════════════════════════════════════════
me_od=sum(1 for c in me if c['status']=='OVERDUE')
me_hp=sum(1 for c in me if c['status']=='HIGH PRIORITY')

_section_open()
st.markdown(_section('⚙','Main Engine',f'{len(me)} component records',f'{me_od} OD · {me_hp} HP','#c08818'),unsafe_allow_html=True)
st.markdown('<div style="padding:1rem 1.05rem 1.15rem">', unsafe_allow_html=True)
st.markdown('<div class="filters-caption">Filters and ordering</div>', unsafe_allow_html=True)

f1,f2,f3=st.columns([2.2,2.1,3.2], gap='large')
with f1:
    mc_sel=st.selectbox('Component',['All']+sorted({c['description'] for c in me}),key='mc')
with f2:
    ms_sel=st.selectbox('Status',['All','Overdue only','High Priority +','OK only'],key='ms')
with f3:
    mr_sel=st.radio('Sort',['Component → Cylinder','Priority → % Used'],horizontal=True,key='mr')

v=me[:]
if mc_sel!='All':              v=[c for c in v if c['description']==mc_sel]
if ms_sel=='Overdue only':     v=[c for c in v if c['status']=='OVERDUE']
elif ms_sel=='High Priority +':v=[c for c in v if c['status'] in ('OVERDUE','HIGH PRIORITY')]
elif ms_sel=='OK only':        v=[c for c in v if c['status']=='OK']

st.markdown(_rec_line(v),unsafe_allow_html=True)
_show(v,mode='matrix' if 'Component' in mr_sel else 'priority')
st.markdown('</div>', unsafe_allow_html=True)
_section_close()


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 2 — AUX ENGINES  — REVISED FRONTEND
# ══════════════════════════════════════════════════════════════════════════
ax_od=sum(1 for c in aux if c['status']=='OVERDUE')
ax_hp=sum(1 for c in aux if c['status']=='HIGH PRIORITY')

_section_open()
st.markdown(_section('🔩','Auxiliary Engines',
                     f'{len(aux)} component records · AUX-1 · AUX-2 · AUX-3',
                     f'{ax_od} OD · {ax_hp} HP','#3e73da'),unsafe_allow_html=True)
st.markdown('<div style="padding:1rem 1.05rem 1.15rem">', unsafe_allow_html=True)
st.markdown('<div class="filters-caption">Filters and ordering</div>', unsafe_allow_html=True)

g1,g2,g3,g4=st.columns([1.4,2.0,2.0,3.0],gap='large')
with g1:
    ae_sel=st.selectbox('Engine',['All']+sorted({c['engine_label'] for c in aux}),key='ae')
with g2:
    ac_sel=st.selectbox('Component',['All']+sorted({c['description'] for c in aux}),key='ac')
with g3:
    as_sel=st.selectbox('Status',['All','Overdue only','High Priority +','OK only'],key='axs')
with g4:
    ar_sel=st.radio('Sort',['Component → Cylinder','Priority → % Used'],horizontal=True,key='ar')

v=aux[:]
if ae_sel!='All':              v=[c for c in v if c['engine_label']==ae_sel]
if ac_sel!='All':              v=[c for c in v if c['description']==ac_sel]
if as_sel=='Overdue only':     v=[c for c in v if c['status']=='OVERDUE']
elif as_sel=='High Priority +':v=[c for c in v if c['status'] in ('OVERDUE','HIGH PRIORITY')]
elif as_sel=='OK only':        v=[c for c in v if c['status']=='OK']

st.markdown(_rec_line(v),unsafe_allow_html=True)
_show(v,mode='matrix' if 'Component' in ar_sel else 'priority')
st.markdown('</div>', unsafe_allow_html=True)
_section_close()


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 3 — OTHER EQUIPMENT  — REVISED FRONTEND
# ══════════════════════════════════════════════════════════════════════════
_section_open()
st.markdown(_section('🛠','Other Equipment',
                     'Turbocharger · Coolers · A/C & Compressors',
                     f'{len(oe)} records','#11a85b'),unsafe_allow_html=True)
st.markdown('<div style="padding:1rem 1.05rem 1.15rem">', unsafe_allow_html=True)

if not oe:
    st.markdown(_empty_html("No other equipment data found.", "#56d68a"), unsafe_allow_html=True)
else:
    secs=sorted({r['section'] for r in oe})
    st.markdown('<div class="filters-caption">Section filter</div>', unsafe_allow_html=True)
    os_sel=st.selectbox('Section',['All']+secs,key='oe_s')
    ov=oe if os_sel=='All' else [r for r in oe if r['section']==os_sel]
    st.markdown(f'<div class="microline"><b>{len(ov)}</b> records</div>',unsafe_allow_html=True)
    st.markdown(_oe_html(ov,show_status=False),unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)
_section_close()


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 4 — D/G EQUIPMENT  — REVISED FRONTEND
# ══════════════════════════════════════════════════════════════════════════
_section_open()
st.markdown(_section('⚡','D/G Equipment',
                     'Diesel Generator components · D/G 1 · D/G 2 · D/G 3',
                     f'{len(dg)} records','#3e73da'),unsafe_allow_html=True)
st.markdown('<div style="padding:1rem 1.05rem 1.15rem">', unsafe_allow_html=True)

if not dg:
    st.markdown(_empty_html("No D/G equipment data found.", "#7ea2ef"), unsafe_allow_html=True)
else:
    st.markdown('<div class="filters-caption">Unit filter</div>', unsafe_allow_html=True)
    dg_sel=st.selectbox('D/G Unit',['All']+sorted({r.get('engine_label','') for r in dg}),key='dge')
    dv=dg if dg_sel=='All' else [r for r in dg if r.get('engine_label')==dg_sel]
    st.markdown(f'<div class="microline"><b>{len(dv)}</b> records</div>',unsafe_allow_html=True)
    st.markdown(_oe_html(dv,show_status=True),unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)
_section_close()


# ══════════════════════════════════════════════════════════════════════════
#  FOOTER
# ══════════════════════════════════════════════════════════════════════════
st.markdown(
    f'<div class="footer-note">Parsed at { _html_escape(d.get("parsed_at","")) }</div>',
    unsafe_allow_html=True
)
