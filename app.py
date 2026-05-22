"""
Fleet Running Hours Monitor  v11
Single page — upload → KPIs → 3 matrices (ME · AUX · Other Equipment)
Parser: validated VBA port, 13/13 ground truth, zero ghost cylinders
"""
import streamlit as st
st.set_page_config(
    page_title="Fleet Running Hours",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

import os, re, subprocess, shutil, tempfile, hashlib
from pathlib import Path
import pandas as pd

# ══════════════════════════════════════════════════════════════════
#  DESIGN SYSTEM
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=Inter:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

/* ── tokens ── */
:root {
  --ink:   #010408;
  --ink1:  #03070e;
  --ink2:  #050b16;
  --ink3:  #080f1e;
  --ink4:  #0b1428;
  --line:  #0f1e36;
  --line2: #182840;
  --line3: #223555;

  --gold:  #b8870c;
  --gold2: #d4a020;
  --gold3: #edbb30;
  --gold4: #f7d060;

  --red:   #b81818;
  --red2:  #e03030;
  --red3:  #ff5858;
  --red4:  #ffa0a0;

  --ora:   #9a3e08;
  --ora2:  #c85c18;
  --ora3:  #f07830;
  --ora4:  #ffb870;

  --grn:   #066030;
  --grn2:  #0c9848;
  --grn3:  #18d868;
  --grn4:  #80f0b0;

  --blu:   #0a2880;
  --blu2:  #1448c0;
  --blu3:  #3878f0;
  --blu4:  #90c0ff;

  --t0: #ddeeff;
  --t1: #9ab8d8;
  --t2: #506880;
  --t3: #243850;
  --t4: #0e1e30;

  --ff: 'Space Grotesk', sans-serif;
  --fi: 'Inter', sans-serif;
  --fm: 'JetBrains Mono', monospace;
  --r:  10px;
}

/* ── reset ── */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
html, body, [class*="css"] {
  font-family: var(--fi) !important;
  background: var(--ink) !important;
  color: var(--t1) !important;
  -webkit-font-smoothing: antialiased;
  text-rendering: optimizeLegibility;
}
.main, .main > div { background: var(--ink) !important; }
.block-container { max-width: 100% !important; padding: 0 2.5rem 5rem !important; }

/* ── atmospheric layer ── */
.main::before {
  content: ''; position: fixed; inset: 0;
  pointer-events: none; z-index: 0;
  background:
    radial-gradient(ellipse 120% 70% at -20% -15%, rgba(184,135,12,.05) 0%, transparent 50%),
    radial-gradient(ellipse 100% 60% at 120% 115%, rgba(10,40,128,.04) 0%, transparent 50%);
}

/* ── sidebar (collapsed state styling) ── */
[data-testid="stSidebar"] { display: none !important; }
[data-testid="collapsedControl"] { display: none !important; }

/* ── typography ── */
h1 { font-family: var(--ff) !important; font-size: 1.7rem !important;
     font-weight: 700 !important; color: var(--t0) !important;
     letter-spacing: -.03em !important; line-height: 1.1 !important; }
h2 { font-family: var(--ff) !important; font-size: 1.05rem !important;
     font-weight: 600 !important; color: var(--t0) !important; }
p  { font-family: var(--fi) !important; }

/* ── native Streamlit metric ── */
[data-testid="stMetric"] {
  background: var(--ink3) !important;
  border: 1px solid var(--line2) !important;
  border-radius: var(--r) !important;
  padding: .85rem 1rem .95rem !important;
  transition: border-color .2s, transform .18s !important;
}
[data-testid="stMetric"]:hover {
  border-color: var(--line3) !important;
  transform: translateY(-2px) !important;
}
[data-testid="stMetricValue"] {
  font-family: var(--ff) !important;
  font-size: 1.8rem !important;
  font-weight: 700 !important;
  color: var(--t0) !important;
  letter-spacing: -.04em !important;
}
[data-testid="stMetricLabel"] {
  color: var(--t3) !important;
  font-size: .58rem !important;
  text-transform: uppercase !important;
  letter-spacing: .16em !important;
}

/* ── dataframe ── */
[data-testid="stDataFrame"] {
  border: 1px solid var(--line2) !important;
  border-radius: var(--r) !important;
  overflow: hidden !important;
  box-shadow: 0 6px 32px rgba(0,0,0,.45) !important;
}
.dvn-scroller { background: var(--ink2) !important; }

/* ── file uploader ── */
[data-testid="stFileUploadDropzone"] {
  background: linear-gradient(155deg,
    rgba(184,135,12,.04) 0%, rgba(10,40,128,.03) 100%) !important;
  border: 1.5px dashed var(--gold2) !important;
  border-radius: 14px !important;
  padding: 2.5rem 2rem !important;
  transition: all .25s !important;
}
[data-testid="stFileUploadDropzone"]:hover {
  background: rgba(184,135,12,.07) !important;
  border-color: var(--gold3) !important;
  box-shadow: 0 0 40px rgba(184,135,12,.06) !important;
}
[data-testid="stFileUploadDropzone"] p,
[data-testid="stFileUploadDropzone"] span {
  color: var(--gold3) !important;
  font-family: var(--ff) !important;
  font-size: .9rem !important;
  font-weight: 500 !important;
}
[data-testid="stFileUploadDropzone"] small { color: var(--t3) !important; }

/* ── button ── */
.stButton > button {
  background: linear-gradient(135deg, var(--gold2) 0%, var(--gold) 100%) !important;
  color: #000 !important; border: none !important;
  border-radius: 8px !important; padding: .6rem 2rem !important;
  font-family: var(--ff) !important; font-weight: 700 !important;
  font-size: .8rem !important; letter-spacing: .07em !important;
  text-transform: uppercase !important;
  box-shadow: 0 2px 16px rgba(184,135,12,.25) !important;
  transition: all .17s !important;
}
.stButton > button:hover {
  background: linear-gradient(135deg, var(--gold3) 0%, var(--gold2) 100%) !important;
  box-shadow: 0 6px 24px rgba(184,135,12,.4) !important;
  transform: translateY(-2px) !important;
}

/* ── expander ── */
.streamlit-expanderHeader {
  background: var(--ink3) !important;
  border: 1px solid var(--line2) !important;
  border-radius: var(--r) !important;
  font-family: var(--ff) !important;
  font-size: .82rem !important; font-weight: 500 !important;
  color: var(--t1) !important;
  transition: all .18s !important;
}
.streamlit-expanderHeader:hover {
  background: var(--ink4) !important;
  border-color: var(--line3) !important;
}
.streamlit-expanderContent {
  background: var(--ink2) !important;
  border: 1px solid var(--line2) !important;
  border-top: none !important;
  border-radius: 0 0 var(--r) var(--r) !important;
}

/* ── select / radio ── */
.stSelectbox > div > div, .stMultiSelect > div > div {
  background: var(--ink3) !important;
  border: 1px solid var(--line2) !important;
  border-radius: 7px !important; color: var(--t1) !important;
}
.stSelectbox label, .stMultiSelect label {
  color: var(--t3) !important; font-size: .67rem !important;
  text-transform: uppercase !important; letter-spacing: .1em !important;
}
[data-testid="stRadio"] > label {
  color: var(--t3) !important; font-size: .67rem !important;
  text-transform: uppercase !important; letter-spacing: .1em !important;
}

/* ── misc ── */
.stAlert { border-radius: 8px !important; border-left-width: 3px !important; }
hr { border-color: var(--line2) !important; opacity: 1 !important; }
a  { color: var(--gold3) !important; text-decoration: none !important; }
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: var(--ink1); }
::-webkit-scrollbar-thumb { background: var(--line3); border-radius: 3px; }

/* ═══════════════════════════════════════════════════════════════
   KEYFRAMES
═══════════════════════════════════════════════════════════════ */
@keyframes fadeDown  { from{opacity:0;transform:translateY(-18px)} to{opacity:1;transform:translateY(0)} }
@keyframes fadeUp    { from{opacity:0;transform:translateY(16px)}  to{opacity:1;transform:translateY(0)} }
@keyframes fadeRight { from{opacity:0;transform:translateX(-14px)} to{opacity:1;transform:translateX(0)} }
@keyframes scaleIn   { from{opacity:0;transform:scale(.9)}         to{opacity:1;transform:scale(1)} }
@keyframes numUp     { from{opacity:0;transform:translateY(8px)}   to{opacity:1;transform:translateY(0)} }
@keyframes growBar   { from{width:0;opacity:0}                     to{width:100%;opacity:1} }
@keyframes pop       { 0%{transform:scale(.84);opacity:0}
                       55%{transform:scale(1.03)}
                       100%{transform:scale(1);opacity:1} }
@keyframes pulse     { 0%,100%{opacity:.6} 50%{opacity:1} }

/* ═══════════════════════════════════════════════════════════════
   CUSTOM COMPONENTS
═══════════════════════════════════════════════════════════════ */

/* — page header — */
.ph {
  padding: 2.25rem 0 0;
  animation: fadeDown .5s cubic-bezier(.22,1,.36,1) both;
}
.ph-eye {
  font-family: var(--fi);
  font-size: .55rem; font-weight: 500;
  letter-spacing: .28em; text-transform: uppercase;
  color: var(--gold2); margin-bottom: .3rem;
  animation: fadeDown .4s .05s ease both;
  animation-fill-mode: both;
}
.ph-rule {
  height: 1px; margin: .35rem 0 2rem;
  background: linear-gradient(90deg, var(--gold2) 0%, var(--line2) 30%, transparent 100%);
  animation: growBar .65s .1s ease both;
  animation-fill-mode: both;
}

/* — KPI card — */
.kc {
  background: var(--ink3);
  border-radius: var(--r);
  padding: .95rem 1.15rem 1.05rem;
  position: relative; overflow: hidden;
  animation: fadeUp .38s ease both;
  animation-fill-mode: both;
  transition: transform .18s, box-shadow .2s;
  cursor: default;
}
.kc:hover { transform: translateY(-4px); box-shadow: 0 14px 40px rgba(0,0,0,.55); }

/* top accent + inner glow via border + ::before */
.kc::before {
  content: ''; position: absolute;
  top: 0; left: 0; right: 0; height: 2px;
  border-radius: var(--r) var(--r) 0 0;
}
.kc::after {
  content: ''; position: absolute; inset: 0;
  border-radius: var(--r); pointer-events: none;
}

.kc.g  { border: 1px solid rgba(184,135,12,.30); }
.kc.g::before { background: linear-gradient(90deg, var(--gold2), transparent 70%); }
.kc.g::after  { background: linear-gradient(160deg, rgba(184,135,12,.06), transparent 55%); }
.kc.g:hover   { border-color: rgba(212,160,32,.55); }

.kc.r  { border: 1px solid rgba(176,24,24,.28); }
.kc.r::before { background: linear-gradient(90deg, var(--red2), transparent 70%); }
.kc.r::after  { background: linear-gradient(160deg, rgba(176,24,24,.05), transparent 55%); }
.kc.r:hover   { border-color: rgba(224,48,48,.45); }

.kc.o  { border: 1px solid rgba(154,62,8,.28); }
.kc.o::before { background: linear-gradient(90deg, var(--ora2), transparent 70%); }
.kc.o::after  { background: linear-gradient(160deg, rgba(154,62,8,.05), transparent 55%); }
.kc.o:hover   { border-color: rgba(200,92,24,.45); }

.kc.n  { border: 1px solid rgba(6,96,48,.28); }
.kc.n::before { background: linear-gradient(90deg, var(--grn2), transparent 70%); }
.kc.n::after  { background: linear-gradient(160deg, rgba(6,96,48,.05), transparent 55%); }
.kc.n:hover   { border-color: rgba(12,152,72,.45); }

.kc.u  { border: 1px solid rgba(10,40,128,.28); }
.kc.u::before { background: linear-gradient(90deg, var(--blu2), transparent 70%); }
.kc.u::after  { background: linear-gradient(160deg, rgba(10,40,128,.05), transparent 55%); }
.kc.u:hover   { border-color: rgba(20,72,192,.45); }

.kc-v {
  font-family: var(--ff);
  font-size: 2rem; font-weight: 700; line-height: 1.1;
  letter-spacing: -.045em;
  position: relative; z-index: 1;
  animation: numUp .38s .1s ease both;
  animation-fill-mode: both;
}
.kc-l {
  font-family: var(--fi);
  font-size: .57rem; font-weight: 500;
  text-transform: uppercase; letter-spacing: .18em;
  color: var(--t3); margin-top: 5px;
  position: relative; z-index: 1;
}
.kc.g .kc-v { color: var(--gold4); }
.kc.r .kc-v { color: var(--red3);  }
.kc.o .kc-v { color: var(--ora3);  }
.kc.n .kc-v { color: var(--grn3);  }
.kc.u .kc-v { color: var(--blu3);  }

/* — section header — */
.sh {
  display: flex; align-items: center; gap: .8rem;
  margin: 2.25rem 0 1.1rem;
}
.sh-icon {
  width: 32px; height: 32px; border-radius: 8px;
  display: flex; align-items: center; justify-content: center;
  font-size: .9rem; flex-shrink: 0;
}
.sh-icon.me  { background: rgba(184,135,12,.12); border: 1px solid rgba(184,135,12,.25); }
.sh-icon.aux { background: rgba(10,40,128,.12);  border: 1px solid rgba(20,72,192,.25); }
.sh-icon.oe  { background: rgba(6,96,48,.12);    border: 1px solid rgba(12,152,72,.25); }
.sh-body { flex: 1; }
.sh-title {
  font-family: var(--ff); font-size: 1rem; font-weight: 600;
  color: var(--t0); letter-spacing: -.01em; line-height: 1;
}
.sh-sub {
  font-family: var(--fm); font-size: .58rem; color: var(--t3);
  margin-top: 3px; letter-spacing: .04em;
}
.sh-rule {
  flex: 1; height: 1px;
  background: linear-gradient(90deg, var(--line2), transparent);
}
.sh-badge {
  font-family: var(--fm); font-size: .58rem; font-weight: 500;
  padding: 3px 9px; border-radius: 20px;
  background: var(--ink3); border: 1px solid var(--line2);
  color: var(--t2); flex-shrink: 0;
}

/* — filter bar — */
.fb {
  background: var(--ink3); border: 1px solid var(--line2);
  border-radius: var(--r); padding: .85rem 1.1rem;
  margin-bottom: .9rem;
}

/* — record count — */
.rc {
  font-family: var(--fm); font-size: .6rem; color: var(--t3);
  margin-bottom: .6rem; line-height: 1.7;
}
.rc b   { color: var(--t1); }
.rc .od { color: var(--red3);  font-weight: 700; }
.rc .hp { color: var(--ora3);  font-weight: 700; }
.rc .ok { color: var(--grn3);  font-weight: 700; }
.rc .nd { color: var(--blu3);  font-weight: 600; }

/* — status pill (inline HTML) — */
.spl {
  display: inline-flex; align-items: center; gap: 5px;
  padding: 2px 10px; border-radius: 20px;
  font-family: var(--fm); font-size: .6rem; font-weight: 600;
  letter-spacing: .05em;
}
.spl.od { background: rgba(176,24,24,.12);  border: 1px solid rgba(224,48,48,.22); color: var(--red3);  }
.spl.hp { background: rgba(154,62,8,.12);   border: 1px solid rgba(200,92,24,.22); color: var(--ora3);  }
.spl.ok { background: rgba(6,96,48,.12);    border: 1px solid rgba(12,152,72,.22); color: var(--grn3);  }
.spl.nd { background: rgba(10,40,128,.08);  border: 1px solid rgba(20,72,192,.18); color: var(--blu3);  }

/* — success banner — */
.success-banner {
  background: linear-gradient(135deg, rgba(6,96,48,.14), rgba(6,96,48,.04));
  border: 1px solid rgba(12,152,72,.28);
  border-radius: var(--r); padding: .9rem 1.4rem;
  color: var(--grn4); font-family: var(--ff);
  font-size: .88rem; font-weight: 500;
  animation: pop .5s cubic-bezier(.34,1.56,.64,1) both;
  display: flex; align-items: center; gap: .65rem;
}

/* — all-clear card — */
.allclear {
  background: rgba(6,96,48,.04); border: 1px solid rgba(6,96,48,.12);
  border-radius: var(--r); padding: 1.5rem; text-align: center;
  color: var(--grn3); font-family: var(--ff);
  font-size: .85rem; font-weight: 500;
}

/* — upload zone hint — */
.upload-hint {
  font-family: var(--fi); font-size: .75rem; color: var(--t3);
  line-height: 1.8; margin-top: .75rem;
}
.upload-hint b { color: var(--t2); font-weight: 500; }

/* — OE table ── */
.oe-section-label {
  font-family: var(--fm); font-size: .58rem; font-weight: 500;
  letter-spacing: .16em; text-transform: uppercase;
  color: var(--t3); padding: .5rem 0 .3rem;
  border-bottom: 1px solid var(--line2); margin-bottom: .5rem;
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
#  CONVERSION
# ══════════════════════════════════════════════════════════════════
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found — packages.txt must contain: libreoffice")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as t:
        t.write(raw); src = t.name
    outdir = tempfile.mkdtemp(prefix="lo_")
    out    = os.path.join(outdir, Path(src).stem + ".docx")
    pf     = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        r = subprocess.run(
            [soffice, "--headless", "--norestore", "--nofirststartwizard",
             f"-env:UserInstallation={pf}", "--convert-to", "docx",
             src, "--outdir", outdir],
            capture_output=True, timeout=120)
        if not os.path.exists(out):
            raise RuntimeError(r.stderr.decode("utf-8","ignore")[:400])
        with open(out, "rb") as f:
            return f.read()
    finally:
        for p in [src, out]:
            try:
                if os.path.exists(p): os.unlink(p)
            except Exception: pass
        shutil.rmtree(outdir, ignore_errors=True)


# ══════════════════════════════════════════════════════════════════
#  PARSER — validated VBA port, 13/13 ground truth
# ══════════════════════════════════════════════════════════════════

def _unique_grid(table) -> list:
    """
    Build 2-D list using only the FIRST occurrence of each merged cell
    (_tc identity dedup). Replicates Word COM behaviour that VBA relies on.
    Without this: ghost cylinders, wrong AUX actual_cyls, remarks bleed.
    """
    grid = []
    for row in table.rows:
        seen, cells = set(), []
        for cell in row.cells:
            cid = id(cell._tc)
            if cid not in seen:
                seen.add(cid)
                raw = re.sub(r'[\x0b\r]', '\n', cell.text).replace('\x07', '')
                lines = [ln.replace('\xa0', ' ').replace('\t', ' ').strip()
                         for ln in raw.split('\n') if ln.strip()]
                cells.append(lines[0] if lines else '')
        grid.append(cells)
    return grid


def _fl(txt: str) -> str:
    for part in re.split(r'[\r\n\x0b]+', str(txt or '')):
        s = re.sub(r'[\x07\xa0\t]+', ' ', part).strip()
        if s: return re.sub(r'  +', ' ', s)
    return ''


def _is_comp(name: str) -> bool:
    u = name.upper().strip()
    if not u or len(u) < 2: return False
    for bad in ('DESCRIPTION','REMARKS','COMPONENT','-','PERIODICITY',
                'PERIODICTLY','DATE OF LAST','RUNNING HOURS','TYPE:',
                'AUX. ENGINE','TURBOCHARGER','MAIN ENGINE'):
        if bad in u: return False
    if re.fullmatch(r'[\d./ ,:-]+', u): return False
    if len(u) > 55: return False
    return bool(re.search(r'[A-Za-z]', u))


def _clean_name(txt: str) -> str:
    t = _fl(txt)
    t = re.sub(r'(?i)ALEXIS\s*Date?', '', t)
    t = re.sub(r'(?i)Page\s*\d+\s*of\s*\d+', '', t)
    return re.sub(r'  +', ' ', t).strip()


def _date(txt: str) -> str:
    s = _fl(txt).strip()
    if not s or s in ('-', '1', '2', '1/', '/ 2', 'N/A', 'n/a'): return ''
    if len(s) > 20: return ''
    if re.match(r'^\d+$', s): return ''
    return s


def _num(txt: str) -> float:
    s = _fl(txt).strip().upper()
    if not s or s in ('', '-', 'N/A', 'CENTRAL'): return 0.0
    if any(w in s for w in ('MONTH','YEAR','WEEK','DAY','OBS')): return 0.0
    s = re.sub(r'([,\.])\s+', r'\1', s)
    m = re.search(r'\d[\d,\.]*', s)
    if not m: return 0.0
    block = m.group()
    sep = max(block.rfind('.'), block.rfind(','))
    if sep > 0 and len(block) - sep == 4:
        block = re.sub(r'[,\.]', '', block)
    elif sep > 0:
        block = re.sub(r'[,\.]', '', block[:sep])
    else:
        block = re.sub(r'[,\.]', '', block)
    try:    return float(block)
    except: return 0.0


def _status(hrs: float, period: float) -> str:
    if hrs <= 0 or period <= 0: return 'NO DATA'
    r = hrs / period
    if r >= 1.0: return 'OVERDUE'
    if r >= 0.8: return 'HIGH PRIORITY'
    return 'OK'


def _pct(hrs: float, period: float) -> float:
    return round(hrs / period, 4) if period and hrs else 0.0


def _make(cat, eng, unit, name, period, date, hrs) -> dict:
    return {
        'category': cat, 'engine_label': eng, 'unit': unit,
        'description': name, 'periodicity': period,
        'last_oh_date': date, 'hrs_since': hrs,
        'pct_used': _pct(hrs, period), 'status': _status(hrs, period),
    }


def _parse_me(grid: list) -> list:
    """
    Main Engine parser.
    Column layout (0-based unique indices — HARDCODED as in VBA):
      0 = component name  |  1 = periodicity  |  2 = row marker (1/2)
      3 = Cyl 1  |  4 = Cyl 2  |  …  |  rem_col = REMARKS (excluded)
    """
    if len(grid) < 3: return []
    if not any('MAIN ENGINE' in (grid[r][c] if c < len(grid[r]) else '').upper()
               for r in range(min(3, len(grid)))
               for c in range(min(12, len(grid[r])))): return []

    rem_col = len(grid[1]) if len(grid) > 1 else 99
    for ci, txt in enumerate(grid[1] if len(grid) > 1 else []):
        if 'REMARK' in txt.upper(): rem_col = ci; break

    FIRST_CYL  = 3
    PERIOD_COL = 1
    MARKER_COL = 2
    actual_cyls = max(1, min(7, rem_col - FIRST_CYL))

    end = len(grid)
    for r, row in enumerate(grid):
        if any(x in ' '.join(row).upper() for x in ('NOTE 1','TURBOCHARGER','AUX. ENGINE')):
            end = r; break

    result, r = [], 1
    while r < end - 1:
        name   = _clean_name(grid[r][0] if grid[r] else '')
        period = _num(grid[r][PERIOD_COL] if PERIOD_COL < len(grid[r]) else '')
        marker = (grid[r][MARKER_COL] if MARKER_COL < len(grid[r]) else '').strip()
        if _is_comp(name) and marker == '1':
            nxt = grid[r + 1] if r + 1 < len(grid) else []
            for cyl in range(1, actual_cyls + 1):
                ci   = FIRST_CYL + cyl - 1
                date = _date(grid[r][ci] if ci < len(grid[r]) else '')
                hrs  = _num(nxt[ci]      if ci < len(nxt)      else '')
                if not date and not hrs: continue
                result.append(_make('MAIN_ENGINE','ME',f'Cyl {cyl}',name,period,date,hrs))
            r += 2
        else:
            r += 1
    return result


def _parse_aux(grid: list) -> list:
    """
    Auxiliary Engine parser.
    Column layout (0-based unique indices):
      0 = name  |  1 = periodicity  |  2 = marker
      3 .. 3+N-1 = AUX-1 Cyls 1..N
      3+N .. 3+2N-1 = AUX-2
      3+2N .. 3+3N-1 = AUX-3
    N = actual_cyls, counted from descRow ucell 3 sequentially starting at 1.
    """
    desc_row = None
    for ri, row in enumerate(grid):
        if 'DESCRIPTION' in (row[0] if row else '').upper():
            desc_row = ri; break
    if desc_row is None: return []

    # Count cylinders: scan descRow from ucell 3 for "1,2,3,4,5,6"
    actual_cyls = 0
    for ci in range(3, len(grid[desc_row])):
        txt = grid[desc_row][ci].strip()
        if txt and re.match(r'^\d+$', txt):
            n = int(txt)
            if n == actual_cyls + 1: actual_cyls = n
            elif actual_cyls > 0:   break
        elif actual_cyls > 0:       break
    actual_cyls = max(1, min(6, actual_cyls)) if actual_cyls else 6

    AUX1 = 3
    AUX2 = AUX1 + actual_cyls
    AUX3 = AUX2 + actual_cyls

    result, r = [], desc_row + 1   # strict +1 lock (VBA THE FIX)
    while r < len(grid) - 1:
        name   = _clean_name(grid[r][0] if grid[r] else '')
        period = _num(grid[r][1] if len(grid[r]) > 1 else '')
        marker = (grid[r][2] if len(grid[r]) > 2 else '').strip()
        if _is_comp(name) and marker == '1':
            nxt = grid[r + 1] if r + 1 < len(grid) else []
            for cyl in range(1, actual_cyls + 1):
                for eng, start in (('AUX-1', AUX1), ('AUX-2', AUX2), ('AUX-3', AUX3)):
                    ci   = start + cyl - 1
                    date = _date(grid[r][ci] if ci < len(grid[r]) else '')
                    hrs  = _num(nxt[ci]      if ci < len(nxt)      else '')
                    if not date and not hrs: continue
                    result.append(_make('AUX_ENGINE',eng,f'Cyl {cyl}',name,period,date,hrs))
            r += 2
        else:
            r += 1
    return result


def _parse_oe(tables) -> list:
    oe = []
    # Table 1: Turbocharger / Coolers / A/C
    if len(tables) > 1:
        grid = _unique_grid(tables[1])
        SKIP = {'TURBOCHARGER','AUXILIARY BOILER','COOLERS','EXH GAS  BOILER',
                'A/C & REFR. COMPRESSORS','MAIN AIR COMPRESSORS',
                'PERIODICTLY','DATE OF LAST INSPECTION','RUN HRS',
                'DATE OF LAST CLEANING','DATE','PERIODICITY',''}
        for row in grid:
            def gc(ci): return row[ci] if ci < len(row) else ''
            for sec, dc, dtc, hrc in [
                ('Turbocharger / Aux Boiler', 0, 2, 3),
                ('Coolers / Exh Gas Boiler',  5, 6, 7),
                ('A/C & Compressors',         10,11,12),
            ]:
                desc = gc(dc)
                if not desc or desc.upper() in SKIP: continue
                dv = gc(dtc); hv = gc(hrc)
                if dv or hv:
                    oe.append({'section':sec,'description':desc,
                               'last_date':dv,'run_hrs':hv})
    # Table 3: D/G Equipment
    if len(tables) > 3:
        grid = _unique_grid(tables[3])
        DG = ['D/G 1','D/G 2','D/G 3']
        for ri, row in enumerate(grid):
            if ri == 0: continue
            def gc2(ci): return row[ci] if ci < len(row) else ''
            for dc, tc, ds in [(0,2,3),(9,11,12)]:
                desc = gc2(dc); rt = gc2(tc)
                if not desc or rt not in ('1','2'): continue
                for gi, gl in enumerate(DG):
                    val = gc2(ds+gi)
                    if not val: continue
                    key = f"{desc} — {gl}"
                    if rt == '1':
                        oe.append({'section':'D/G Equipment','description':key,
                                   'last_date':val,'run_hrs':''})
                    else:
                        for e in reversed(oe):
                            if e['description'] == key and e['run_hrs'] == '':
                                e['run_hrs'] = val; break
                        else:
                            oe.append({'section':'D/G Equipment','description':key,
                                       'last_date':'','run_hrs':val})
    return oe


def parse_doc_bytes(docx_bytes: bytes) -> dict:
    from docx import Document
    warns = []
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as t:
        t.write(docx_bytes); tp = t.name
    try:    doc = Document(tp)
    except Exception as e: raise ValueError(f"Cannot open document: {e}")
    finally:
        try: os.unlink(tp)
        except Exception: pass

    if not doc.tables:
        raise ValueError("No tables found — is this a TEC-004 report?")

    # Vessel + date
    vn = 'UNKNOWN'; rd = None
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt: continue
        if m := re.search(r"Vessel[\u2019\u2018']?s?\s+Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+?)(?:\s{2,}|\t)", txt, re.I):
            vn = _clean_name(m.group(1))
        if m := re.search(r"Date\s*:\s*(.+)", txt, re.I):
            rd = _date(m.group(1).strip())
        if vn != 'UNKNOWN' and rd: break
    if vn == 'UNKNOWN': warns.append("Could not extract vessel name.")

    # ME totals
    me_tot = me_mo = None
    g0 = _unique_grid(doc.tables[0])
    for cell_txt in (g0[0] if g0 else []):
        if m := re.search(r'Total Running Hours[\s:ǀ|]+([\d,]+)', cell_txt, re.I):
            me_tot = int(_num(m.group(1)))
        if m := re.search(r'This Month[\s:]+([\d,]+)', cell_txt, re.I):
            me_mo  = int(_num(m.group(1)))

    me_comps  = _parse_me(g0)
    aux_comps = _parse_aux(_unique_grid(doc.tables[2])) if len(doc.tables) > 2 else []
    oe        = _parse_oe(doc.tables)
    comps     = me_comps + aux_comps

    if not comps: warns.append("No components extracted.")

    return {
        'vessel_name':     vn,
        'report_date':     rd,
        'me_total_hrs':    me_tot,
        'me_this_month':   me_mo,
        'components':      comps,
        'me_comps':        me_comps,
        'aux_comps':       aux_comps,
        'other_equipment': oe,
        'warnings':        warns,
        'uploaded_at':     datetime.utcnow().isoformat(),
    }


from datetime import datetime


# ══════════════════════════════════════════════════════════════════
#  SESSION STATE — single source of truth, no DB timing race
# ══════════════════════════════════════════════════════════════════
if 'parsed' not in st.session_state:
    st.session_state.parsed = None   # holds the latest parsed result


# ══════════════════════════════════════════════════════════════════
#  TABLE ENGINE
# ══════════════════════════════════════════════════════════════════
_STATUS_ORD = {'OVERDUE': 0, 'HIGH PRIORITY': 1, 'OK': 2, 'NO DATA': 3}

# Status colour themes — high contrast, visible on dark backgrounds
_TH = {
    'OVERDUE': {
        'bg':  'background-color:#1e0404',
        'bgs': 'background-color:#2c0505',
        'ts':  'color:#ff5252;font-weight:700',
        'tm':  'color:#ff7878',
        'tn':  'color:#ff1a1a;font-weight:700',
        'td':  'color:#822222',
    },
    'HIGH PRIORITY': {
        'bg':  'background-color:#1e0e02',
        'bgs': 'background-color:#2a1403',
        'ts':  'color:#ffaa33;font-weight:700',
        'tm':  'color:#ffcc77',
        'tn':  'color:#ffcc00;font-weight:700',
        'td':  'color:#7a5010',
    },
    'OK': {
        'bg':  'background-color:#011008',
        'bgs': 'background-color:#01180a',
        'ts':  'color:#22ee66;font-weight:700',
        'tm':  'color:#44dd88',
        'tn':  'color:#22ee66;font-weight:700',
        'td':  'color:#0a4820',
    },
    'NO DATA': {
        'bg':  'background-color:#060c18',
        'bgs': 'background-color:#0a1122',
        'ts':  'color:#4488bb;font-weight:600',
        'tm':  'color:#336688',
        'tn':  'color:#4488bb',
        'td':  'color:#1a2e44',
    },
}

# Per-column style — dynamic dict so any column subset works
_CS = {
    'Status':      lambda t: f"{t['bgs']};{t['ts']}",
    'Component':   lambda t: f"{t['bg']};{t['tm']};font-weight:600",
    'Engine':      lambda t: f"{t['bg']};{t['td']}",
    'Unit':        lambda t: f"{t['bg']};{t['td']}",
    'Periodicity': lambda t: f"{t['bg']};{t['td']}",
    'Last O/H':    lambda t: f"{t['bg']};{t['td']}",
    'Hrs Since':   lambda t: f"{t['bg']};{t['tm']};font-weight:600",
    'Used %':      lambda t: f"{t['bg']};{t['tn']}",
}
_CS_DEF = lambda t: f"{t['bg']};{t['td']}"


def _sf(x):
    try:
        v = float(x)
        return None if pd.isna(v) else v
    except Exception: return None


def _cyl_n(u: str) -> int:
    m = re.search(r'\d+', str(u)); return int(m.group()) if m else 999


def _build(records: list, mode: str = 'matrix') -> pd.DataFrame:
    """
    Build display DataFrame from list of component dicts.
    mode: 'matrix'   → component A–Z → cylinder 1–N
          'priority'  → OVERDUE → HIGH PRIORITY → OK, then % used desc
    """
    if not records:
        return pd.DataFrame(columns=['Status','Component','Engine','Unit',
                                      'Periodicity','Last O/H','Hrs Since','Used %'])
    df = pd.DataFrame(records)
    df['_s'] = df['status'].map(lambda x: _STATUS_ORD.get(str(x), 4))
    df['_p'] = df['pct_used'].apply(lambda x: _sf(x) or 0.0)

    if mode == 'matrix':
        df['_k1'] = df['description'].str.upper()
        df['_k2'] = df['unit'].apply(_cyl_n)
        df = df.sort_values(['_k1','_k2']).drop(columns=['_k1','_k2','_s','_p'])
    else:  # priority
        df = df.sort_values(['_s','_p'], ascending=[True,False]).drop(columns=['_s','_p'])

    out = pd.DataFrame(index=range(len(df)))
    out['Status']      = df['status'].values
    out['Component']   = df['description'].values
    out['Engine']      = df['engine_label'].values
    out['Unit']        = df['unit'].values
    out['Periodicity'] = [int(float(x)) if _sf(x) else None for x in df['periodicity'].values]
    out['Last O/H']    = [str(x) if x and str(x) not in ('nan','None','') else '—'
                          for x in df['last_oh_date'].values]
    out['Hrs Since']   = [int(float(x)) if _sf(x) else None for x in df['hrs_since'].values]
    out['Used %']      = [round(float(x)*100,1) if _sf(x) else 0.0 for x in df['pct_used'].values]
    return out


def _style(df: pd.DataFrame):
    """Dynamic styling — list length always matches column count."""
    cols = list(df.columns)
    def _r(row):
        t = _TH.get(str(row.get('Status','')), _TH['NO DATA'])
        return [_CS.get(col, _CS_DEF)(t) for col in cols]
    return df.style.apply(_r, axis=1)


_CC = {
    'Status':      st.column_config.TextColumn('Status',       width=130),
    'Component':   st.column_config.TextColumn('Component',    width=230),
    'Engine':      st.column_config.TextColumn('Engine',       width=82),
    'Unit':        st.column_config.TextColumn('Unit',         width=70),
    'Periodicity': st.column_config.NumberColumn('Periodicity', format='%d hrs', width=112),
    'Last O/H':    st.column_config.TextColumn('Last O/H',     width=108),
    'Hrs Since':   st.column_config.NumberColumn('Hrs Since',   format='%d hrs', width=108),
    'Used %':      st.column_config.ProgressColumn(
                       'Used %', min_value=0, max_value=160,
                       format='%.1f%%', width=135),
}


def matrix(records: list, mode: str = 'matrix', height: int = None):
    """Render a component matrix. mode: 'matrix' | 'priority'"""
    if not records: st.info('No data.'); return
    tbl = _build(records, mode=mode)
    if tbl.empty: st.info('No data.'); return
    cfg = {k: v for k, v in _CC.items() if k in tbl.columns}
    h   = height or min(880, 38*len(tbl) + 44)
    st.dataframe(_style(tbl), use_container_width=True,
                 hide_index=True, height=h, column_config=cfg)


# ══════════════════════════════════════════════════════════════════
#  HTML HELPERS
# ══════════════════════════════════════════════════════════════════

def kpi(val, lbl: str, c: str = 'g', delay: float = 0.0) -> str:
    return (f'<div class="kc {c}" style="animation-delay:{delay}s">'
            f'<div class="kc-v">{val}</div>'
            f'<div class="kc-l">{lbl}</div></div>')


def section_header(icon: str, title: str, sub: str, badge: str,
                   icon_cls: str = 'me') -> str:
    return (f'<div class="sh">'
            f'<div class="sh-icon {icon_cls}">{icon}</div>'
            f'<div class="sh-body">'
            f'<div class="sh-title">{title}</div>'
            f'<div class="sh-sub">{sub}</div>'
            f'</div>'
            f'<div class="sh-rule"></div>'
            f'<div class="sh-badge">{badge}</div>'
            f'</div>')


def record_count(records: list, label: str = '') -> str:
    n   = len(records)
    od  = sum(1 for c in records if c['status']=='OVERDUE')
    hp  = sum(1 for c in records if c['status']=='HIGH PRIORITY')
    ok  = sum(1 for c in records if c['status']=='OK')
    nd  = sum(1 for c in records if c['status']=='NO DATA')
    parts = [f'<b>{n}</b> records']
    if od: parts.append(f'<span class="od">{od} overdue</span>')
    if hp: parts.append(f'<span class="hp">{hp} high priority</span>')
    if ok: parts.append(f'<span class="ok">{ok} OK</span>')
    if nd: parts.append(f'<span class="nd">{nd} no data</span>')
    return f'<div class="rc">{" · ".join(parts)}</div>'


# ══════════════════════════════════════════════════════════════════
#  PAGE
# ══════════════════════════════════════════════════════════════════

# ── header ───────────────────────────────────────────────────────
st.markdown(
    '<div class="ph">'
    '<div class="ph-eye">Running Hours Management System</div>'
    '<h1>Fleet Overview</h1>'
    '</div>'
    '<div class="ph-rule"></div>',
    unsafe_allow_html=True)

# ── upload ───────────────────────────────────────────────────────
with st.expander("Upload TEC-004 Report", expanded=(st.session_state.parsed is None)):
    uc, ic = st.columns([3, 2], gap='large')
    with uc:
        uploaded = st.file_uploader(
            'Drop .doc file here', type=['doc'], label_visibility='collapsed')
        if uploaded:
            st.markdown(
                '<div class="upload-hint">'
                f'<b>{uploaded.name}</b> &nbsp;·&nbsp; '
                f'{uploaded.size/1024:.1f} kB &nbsp;·&nbsp; ready to parse'
                '</div>',
                unsafe_allow_html=True)
    with ic:
        st.markdown(
            '<div class="upload-hint" style="padding:.6rem 0">'
            '<b>Accepted:</b> TEC-004 Running Hours Monthly Report (.doc)<br>'
            '<b>Extracted:</b> Vessel name · Report date · M/E totals<br>'
            '<b>Matrices:</b> Main Engine · AUX-1/2/3 · Other Equipment<br>'
            '<b>Status:</b> Computed per periodicity (OK / HP / Overdue)'
            '</div>',
            unsafe_allow_html=True)

    if uploaded:
        raw = uploaded.read()
        fh  = hashlib.md5(raw).hexdigest()

        # Don't re-parse the same file
        if (st.session_state.parsed is None or
                st.session_state.parsed.get('file_hash') != fh):

            with st.spinner('Converting .doc → .docx via LibreOffice…'):
                try:
                    docx = convert_doc_to_docx(raw)
                except Exception as e:
                    st.error(f'Conversion failed: {e}'); st.stop()

            with st.spinner('Parsing tables…'):
                try:
                    parsed = parse_doc_bytes(docx)
                    parsed['file_hash'] = fh
                    parsed['filename']  = uploaded.name
                except ValueError as e:
                    st.error(f'Parse failed: {e}'); st.stop()

            for w in parsed['warnings']:
                st.warning(f'⚠ {w}')

            if not parsed['components']:
                st.error('No components extracted. Verify this is a TEC-004 report.')
                st.stop()

            st.session_state.parsed = parsed
            st.markdown(
                f'<div class="success-banner">'
                f'<span style="font-size:1.3rem">✓</span>'
                f'<span><strong>{parsed["vessel_name"]}</strong> — '
                f'{len(parsed["components"])} components parsed · '
                f'{sum(1 for c in parsed["components"] if c["status"]=="OVERDUE")} overdue · '
                f'{sum(1 for c in parsed["components"] if c["status"]=="HIGH PRIORITY")} high priority'
                f'</span></div>',
                unsafe_allow_html=True)

# ── nothing loaded yet ───────────────────────────────────────────
if st.session_state.parsed is None:
    st.markdown(
        '<div style="height:35vh;display:flex;align-items:center;justify-content:center">'
        '<div style="text-align:center;color:var(--t3);font-family:var(--fi)">'
        '<div style="font-size:2.5rem;margin-bottom:.75rem;opacity:.3">⚓</div>'
        '<div style="font-size:.85rem;letter-spacing:.04em">'
        'Upload a TEC-004 report to view the matrices</div>'
        '</div></div>',
        unsafe_allow_html=True)
    st.stop()

# ── from here: data is guaranteed present ────────────────────────
p    = st.session_state.parsed
me   = p['me_comps']
aux  = p['aux_comps']
oe   = p['other_equipment']
all_ = p['components']

# ════════════════════════════════════════════════════════════════
#  KPI ROW
# ════════════════════════════════════════════════════════════════
n_od = sum(1 for c in all_ if c['status']=='OVERDUE')
n_hp = sum(1 for c in all_ if c['status']=='HIGH PRIORITY')
n_ok = sum(1 for c in all_ if c['status']=='OK')
n_me = len(me)
n_ax = len(aux)
me_t = p['me_total_hrs']
me_m = p['me_this_month']

c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
kpi_data = [
    (c1, p['vessel_name'],               'Vessel',         'u', 0.00),
    (c2, p['report_date'] or '—',        'Report Date',    'g', 0.05),
    (c3, f"{me_t:,}" if me_t else '—',  'M/E Total Hrs',  'g', 0.10),
    (c4, f"{me_m:,}" if me_m else '—',  'M/E This Month', 'g', 0.15),
    (c5, n_od,                            'Overdue',        'r', 0.20),
    (c6, n_hp,                            'High Priority',  'o', 0.25),
    (c7, n_ok,                            'OK',             'n', 0.30),
]
for col, val, lbl, clr, dly in kpi_data:
    with col: st.markdown(kpi(val, lbl, clr, dly), unsafe_allow_html=True)

st.markdown('<div style="height:.5rem"></div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════
#  MATRIX 1 — MAIN ENGINE
# ════════════════════════════════════════════════════════════════
me_od = sum(1 for c in me if c['status']=='OVERDUE')
me_hp = sum(1 for c in me if c['status']=='HIGH PRIORITY')

st.markdown(
    section_header(
        '⚙', 'Main Engine', f'{n_me} components',
        f'{me_od} OD · {me_hp} HP', 'me'),
    unsafe_allow_html=True)

# Filters
mf1, mf2, mf3 = st.columns([2,2,3])
with mf1:
    me_comp_opts = ['All'] + sorted({c['description'] for c in me})
    me_cf = st.selectbox('Component', me_comp_opts, key='me_cf')
with mf2:
    me_sf = st.selectbox('Status',
        ['All','Overdue only','High Priority +','OK only'], key='me_sf')
with mf3:
    me_sort = st.radio('Sort',
        ['Component → Cylinder', 'Priority → % Used'],
        horizontal=True, key='me_sort')

me_view = me[:]
if me_cf != 'All':                 me_view = [c for c in me_view if c['description']==me_cf]
if me_sf == 'Overdue only':        me_view = [c for c in me_view if c['status']=='OVERDUE']
elif me_sf == 'High Priority +':   me_view = [c for c in me_view if c['status'] in ('OVERDUE','HIGH PRIORITY')]
elif me_sf == 'OK only':           me_view = [c for c in me_view if c['status']=='OK']

st.markdown(record_count(me_view), unsafe_allow_html=True)

if me_view:
    matrix(me_view,
           mode='matrix' if 'Component' in me_sort else 'priority',
           height=min(820, 38*len(me_view)+44))
else:
    st.markdown('<div class="allclear">No records match the current filter</div>',
                unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════
#  MATRIX 2 — AUXILIARY ENGINES
# ════════════════════════════════════════════════════════════════
ax_od = sum(1 for c in aux if c['status']=='OVERDUE')
ax_hp = sum(1 for c in aux if c['status']=='HIGH PRIORITY')

st.markdown(
    section_header(
        '🔩', 'Auxiliary Engines', f'{n_ax} components',
        f'{ax_od} OD · {ax_hp} HP', 'aux'),
    unsafe_allow_html=True)

af1, af2, af3, af4 = st.columns([1.5,2,2,3])
with af1:
    eng_opts = ['All'] + sorted({c['engine_label'] for c in aux})
    ax_ef = st.selectbox('Engine', eng_opts, key='ax_ef')
with af2:
    ax_comp_opts = ['All'] + sorted({c['description'] for c in aux})
    ax_cf = st.selectbox('Component', ax_comp_opts, key='ax_cf')
with af3:
    ax_sf = st.selectbox('Status',
        ['All','Overdue only','High Priority +','OK only'], key='ax_sf')
with af4:
    ax_sort = st.radio('Sort',
        ['Component → Cylinder', 'Priority → % Used'],
        horizontal=True, key='ax_sort')

ax_view = aux[:]
if ax_ef != 'All':                 ax_view = [c for c in ax_view if c['engine_label']==ax_ef]
if ax_cf != 'All':                 ax_view = [c for c in ax_view if c['description']==ax_cf]
if ax_sf == 'Overdue only':        ax_view = [c for c in ax_view if c['status']=='OVERDUE']
elif ax_sf == 'High Priority +':   ax_view = [c for c in ax_view if c['status'] in ('OVERDUE','HIGH PRIORITY')]
elif ax_sf == 'OK only':           ax_view = [c for c in ax_view if c['status']=='OK']

st.markdown(record_count(ax_view), unsafe_allow_html=True)

if ax_view:
    matrix(ax_view,
           mode='matrix' if 'Component' in ax_sort else 'priority',
           height=min(820, 38*len(ax_view)+44))
else:
    st.markdown('<div class="allclear">No records match the current filter</div>',
                unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════
#  MATRIX 3 — OTHER EQUIPMENT
# ════════════════════════════════════════════════════════════════
st.markdown(
    section_header(
        '🛠', 'Other Equipment',
        'Turbocharger · Coolers · A/C · D/G',
        f'{len(oe)} records', 'oe'),
    unsafe_allow_html=True)

if not oe:
    st.markdown('<div class="allclear">No other equipment data found in this report</div>',
                unsafe_allow_html=True)
else:
    oe_df = pd.DataFrame(oe)
    for sec in sorted(oe_df['section'].unique()):
        st.markdown(
            f'<div class="oe-section-label">{sec}</div>',
            unsafe_allow_html=True)
        sd = oe_df[oe_df['section']==sec][['description','last_date','run_hrs']].copy()
        sd.columns = ['Description', 'Last Date / O/H', 'Run Hrs']
        st.dataframe(
            sd,
            use_container_width=True,
            hide_index=True,
            column_config={
                'Description':    st.column_config.TextColumn('Description',    width=300),
                'Last Date / O/H':st.column_config.TextColumn('Last Date / O/H',width=160),
                'Run Hrs':        st.column_config.TextColumn('Run Hrs',         width=120),
            })
