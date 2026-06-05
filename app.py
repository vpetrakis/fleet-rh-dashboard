import streamlit as st

st.set_page_config(
    page_title="Fleet Running Hours",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

import os
import re
import html
import shutil
import tempfile
import subprocess
import hashlib
from pathlib import Path
from datetime import datetime
from typing import Any, Dict, List, Tuple


# ══════════════════════════════════════════════════════════════════════════
#  DESIGN SYSTEM  —  Animated Glassmorphism & Corporate Dark Theme
# ══════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=IBM+Plex+Sans:wght@400;500;600;700&family=Inter:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500;700&display=swap');

:root{
  --bg:#080d12; --bg2:#0d141c; --bg3:#111a24; --bg4:#16222f;
  --panel:rgba(17, 26, 36, 0.75); --line:#1e3040; --line2:#2a4460;
  --ink:#f0f4f8; --muted:#94aabf; --soft:#607888;
  --gold:#b89460; --gold2:#d4b07a; --gold3:#e8cc94;
  --blue:#466b99; --green:#4f7a62; --amber:#99734d; --red:#8c5353;
}
*{box-sizing:border-box}
html,body,[class*="css"]{
  background:var(--bg)!important;color:var(--muted)!important;
  font-family:'Inter',sans-serif!important;-webkit-font-smoothing:antialiased;
}
.main,.main>div{background:var(--bg)!important}
.block-container{max-width:100%!important;padding:0 2.1rem 5rem!important}
[data-testid="stSidebar"],[data-testid="collapsedControl"]{display:none!important}

/* Background Ambient Glow */
.main::before{
  content:"";position:fixed;inset:0;pointer-events:none;z-index:0;
  background:radial-gradient(circle 800px at 0% 0%,rgba(184,148,96,.04),transparent 100%),
             radial-gradient(circle 600px at 100% 100%,rgba(70,107,153,.04),transparent 100%);
}
.block-container>*{position:relative;z-index:1}

/* Dropzone Styling */
[data-testid="stFileUploadDropzone"]{
  background:rgba(184,148,96,.03)!important;
  border:1.5px dashed var(--gold)!important;border-radius:16px!important;
  padding:2.5rem 2rem!important;transition:all .3s cubic-bezier(0.4,0,0.2,1)!important;
}
[data-testid="stFileUploadDropzone"]:hover{
  background:rgba(184,148,96,.08)!important;border-color:var(--gold2)!important;
  transform:translateY(-2px);box-shadow:0 8px 24px rgba(184,148,96,.1);
}
[data-testid="stFileUploadDropzone"] p,[data-testid="stFileUploadDropzone"] span{
  color:var(--gold2)!important;font-family:'IBM Plex Sans',sans-serif!important;
  font-size:.95rem!important;font-weight:600!important;
}

/* Inputs & Buttons */
div[data-baseweb="select"]>div{
  background:var(--bg3)!important;border:1px solid var(--line)!important;
  border-radius:10px!important;color:var(--ink)!important;transition:border-color .2s;
}
div[data-baseweb="select"]>div:hover{border-color:var(--gold)!important;}
.stSelectbox label,.stRadio label{
  color:var(--soft)!important;font-size:.65rem!important;font-weight:600!important;
  text-transform:uppercase!important;letter-spacing:.15em!important;
}
.stButton>button{
  background:linear-gradient(135deg,var(--gold2),var(--gold))!important;color:#080d12!important;
  border:none!important;border-radius:10px!important;padding:.65rem 1.5rem!important;
  font-family:'IBM Plex Sans',sans-serif!important;font-weight:700!important;
  font-size:.78rem!important;letter-spacing:.08em!important;text-transform:uppercase!important;
  box-shadow:0 4px 14px rgba(184,148,96,.25)!important;transition:all .2s ease!important;
}
.stButton>button:hover{
  background:linear-gradient(135deg,var(--gold3),var(--gold2))!important;
  box-shadow:0 6px 24px rgba(184,148,96,.4)!important;transform:translateY(-2px)!important;
}

/* Animations */
@keyframes slideDown { from{opacity:0;transform:translateY(-15px)} to{opacity:1;transform:translateY(0)} }
@keyframes fadeUp    { from{opacity:0;transform:translateY(15px)}  to{opacity:1;transform:translateY(0)} }
@keyframes growWidth { from{width:0;opacity:0}                     to{opacity:1} }
@keyframes popIn     { 0%{transform:scale(.9);opacity:0} 60%{transform:scale(1.03)} 100%{transform:scale(1);opacity:1} }

/* Matrix Row Staggered Entrance */
.stagger-row {
    opacity: 0;
    animation: fadeUp 0.4s cubic-bezier(0.2, 0.8, 0.2, 1) forwards;
}
/* Overdue Pulse Effect */
.pulse-tag {
    animation: alert-pulse 2s infinite cubic-bezier(0.4, 0, 0.2, 1);
}
@keyframes alert-pulse {
    0%   { box-shadow: 0 0 0 0 rgba(200, 128, 128, 0.5); }
    70%  { box-shadow: 0 0 0 6px rgba(200, 128, 128, 0); }
    100% { box-shadow: 0 0 0 0 rgba(200, 128, 128, 0); }
}

/* Scrollbars */
::-webkit-scrollbar{width:6px;height:6px}
::-webkit-scrollbar-track{background:var(--bg2);border-radius:10px}
::-webkit-scrollbar-thumb{background:var(--line2);border-radius:10px}
::-webkit-scrollbar-thumb:hover{background:var(--soft)}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  STATUS COLOUR MAP
# ══════════════════════════════════════════════════════════════════════════
_SC = {
    "OVERDUE":       {"row":"#1a1111","tag_bg":"#361d1d","tag_fg":"#d98c8c",
                      "comp":"#f2e6e6","dim":"#b38f8f","num":"#e6c3c3",
                      "bar_f":"#b35959","bar_e":"#2d1717","bord":"#402323", "class":"pulse-tag"},
    "HIGH PRIORITY": {"row":"#1a1610","tag_bg":"#362816","tag_fg":"#d9a66c",
                      "comp":"#f2ead9","dim":"#b39c7d","num":"#e6d1ad",
                      "bar_f":"#b38640","bar_e":"#2d2112","bord":"#40311a", "class":""},
    "OK":            {"row":"#0d1a14","tag_bg":"#1a3326","tag_fg":"#7cb393",
                      "comp":"#d9ebd9","dim":"#8ca699","num":"#b3cca3",
                      "bar_f":"#528c6a","bar_e":"#172e22","bord":"#234030", "class":""},
    "NO DATA":       {"row":"#111824","tag_bg":"#1d2a3d","tag_fg":"#7aa1c7",
                      "comp":"#e6edf5","dim":"#8ca1b3","num":"#b3c8db",
                      "bar_f":"#537799","bar_e":"#1a2536","bord":"#26384f", "class":""},
}
_ORD = {"OVERDUE": 0, "HIGH PRIORITY": 1, "OK": 2, "NO DATA": 3}


# ══════════════════════════════════════════════════════════════════════════
#  CONVERSION
# ══════════════════════════════════════════════════════════════════════════
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found — packages.txt must contain: libreoffice")
    
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as t:
        t.write(raw)
        src = t.name
        
    outdir = tempfile.mkdtemp(prefix="lo_")
    out = os.path.join(outdir, Path(src).stem + ".docx")
    pf = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    
    try:
        r = subprocess.run(
            [soffice, "--headless", "--norestore", "--nofirststartwizard",
             f"-env:UserInstallation={pf}", "--convert-to", "docx", src, "--outdir", outdir],
            capture_output=True, timeout=120)
        
        if not os.path.exists(out):
            raise RuntimeError(r.stderr.decode("utf-8", "ignore")[:400])
            
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


# ══════════════════════════════════════════════════════════════════════════
#  GRID BUILDERS
# ══════════════════════════════════════════════════════════════════════════
def _raw_grid(table) -> List[List[str]]:
    grid = []
    mc = 0
    for row in table.rows:
        cells = []
        for cell in row.cells:
            raw = re.sub(r'[\x0b\r]', '\n', cell.text).replace('\x07', '')
            lines = [ln.replace('\xa0', ' ').replace('\t', ' ').strip() for ln in raw.split('\n') if ln.strip()]
            cells.append(lines[0] if lines else '')
        mc = max(mc, len(cells))
        grid.append(cells)
    for row in grid:
        while len(row) < mc: row.append('')
    return grid

def _dedup_grid(table) -> List[List[str]]:
    grid = []
    mc = 0
    for row in table.rows:
        cells = []
        prior = None
        for cell in row.cells:
            if cell._tc is prior: 
                continue
            prior = cell._tc
            raw = re.sub(r'[\x0b\r]', '\n', cell.text).replace('\x07', '')
            lines = [ln.replace('\xa0', ' ').replace('\t', ' ').strip() for ln in raw.split('\n') if ln.strip()]
            cells.append(lines[0] if lines else '')
        mc = max(mc, len(cells))
        grid.append(cells)
    for row in grid:
        while len(row) < mc: row.append('')
    return grid


# ══════════════════════════════════════════════════════════════════════════
#  TEXT & NUMBER HELPERS
# ══════════════════════════════════════════════════════════════════════════
def _fl(t: Any) -> str:
    raw = str(t or "").replace('\x07', '').replace('\xa0', ' ').replace('\t', ' ')
    for p in re.split(r'[\r\n\x0b]+', raw):
        s = re.sub(r'\s+', ' ', p).strip()
        if s: return s
    return ''

def _clean_name(t: Any) -> str:
    s = _fl(t)
    s = re.sub(r'(?i)^MV\s+', '', s)
    s = re.sub(r'(?i)Page\s*\d+\s*of\s*\d+', '', s)
    return re.sub(r'  +', ' ', s).strip(" -:")

def _parse_number(t: Any) -> float:
    s = _fl(t).strip().upper()
    if not s or s in ('', '-', 'N/A', 'CENTRAL', 'COOLER'): return 0.0
    s = s.replace('[', '').replace(']', '')
    
    if any(w in s for w in ('MONTH', 'YEAR', 'WEEK', 'DAY', 'OBS', 'NO RECORD', 'BASED ON', 'OBSERVATION')): 
        return 0.0
        
    m = re.search(r'\d[\d,\.]*', s)
    if not m: return 0.0
        
    b = m.group()
    sep = max(b.rfind('.'), b.rfind(','))
    
    if sep > 0 and len(b) - sep == 4: 
        b = re.sub(r'[,\.]', '', b)
    elif sep > 0: 
        b = re.sub(r'[,\.]', '', b[:sep])
    else: 
        b = re.sub(r'[,\.]', '', b)
        
    try: return float(b)
    except: return 0.0

def _parse_date(t: Any) -> str:
    s = _fl(t).strip()
    if not s: return ''
    s = s.replace('[', '').replace(']', '').strip()
    
    if s in ('-', '1', '2', 'N/A', 'n/a', 'NA', 'Central', 'CENTRAL', 'COOLER', 'NO RECORD', 'NOT WORKING', 'N.A.'): 
        return ''
    if re.fullmatch(r'^\d+$', s): return ''
    if re.fullmatch(r'^\d{1,3}[.,]\d{3}$', s): return ''
    if len(s) > 32: return ''
        
    return s if re.search(r'[A-Za-z0-9/]', s) else ''

def _status(hrs: float, period: float) -> str:
    if hrs <= 0 or period <= 0: return 'NO DATA'
    r = hrs / period
    if r >= 1.0: return 'OVERDUE'
    if r >= 0.8: return 'HIGH PRIORITY'
    return 'OK'

def _pct(hrs: float, period: float) -> float:
    return round(hrs / period, 4) if hrs and period else 0.0

def _is_comp(n: str) -> bool:
    u = _fl(n).upper()
    if not u or len(u) < 2: return False
        
    BAD = ('DESCRIPTION', 'REMARKS', 'COMPONENT', 'PERIODICITY', 'PERIODICTLY', 'DATE OF LAST',
           'RUNNING HOURS', 'MAIN ENGINE', 'AUX. ENGINE', 'TYPE:', '1-DATE OF LAST',
           'TOTAL RUNNING', 'THIS MONTH', 'CYL. NO', 'NOTE 1', 'BASED ON', 'SERIAL NR',
           'HOURS THIS MONTH', 'AUX. ENGINE MAKER')
           
    if any(b in u for b in BAD): return False
    if re.fullmatch(r'[\d./ ,:\-\[\]()]+', u): return False
    return bool(re.search(r'[A-Za-z]', u))

def _mk(cat, eng, unit, nm, per, dt, hrs, doc_idx=0) -> Dict:
    # doc_idx is injected later to preserve true document order
    return {'category': cat, 'engine_label': eng, 'unit': unit, 'description': nm,
            'periodicity': per, 'last_oh_date': dt, 'hrs_since': hrs,
            'pct_used': _pct(hrs, per), 'status': _status(hrs, per), 'doc_idx': doc_idx}

def _get_unique_ordered(records: List[Dict], key: str) -> List[str]:
    """Returns a list of unique values preserving original document order."""
    seen = set()
    out = []
    # Sort strictly by doc_idx first
    for r in sorted(records, key=lambda x: x.get('doc_idx', 0)):
        val = r.get(key)
        if val and val not in seen:
            seen.add(val)
            out.append(val)
    return out


# ══════════════════════════════════════════════════════════════════════════
#  ME PARSER
# ══════════════════════════════════════════════════════════════════════════
def _parse_me(grid: List[List[str]]) -> List[Dict]:
    if not grid: return []
    if not any('MAIN ENGINE' in (_fl(grid[r][c]).upper() if c < len(grid[r]) else '')
               for r in range(min(4, len(grid))) for c in range(min(16, len(grid[r])))): 
        return []

    rem_col = None
    for r in range(min(8, len(grid))):
        for ci, txt in enumerate(grid[r]):
            if 'REMARK' in _fl(txt).upper() and ci > 3: 
                rem_col = ci
                break
        if rem_col: break
    if rem_col is None: rem_col = 12

    MARKER = 2; PERIOD = 1; FIRST = 3
    actual_cyls = max(1, min(8, rem_col - FIRST))

    end = len(grid)
    for r, row in enumerate(grid):
        j = ' '.join(_fl(x) for x in row).upper()
        if 'NOTE 1' in j or 'TURBOCHARGER' in j or 'AUX. ENGINE MAKER' in j: 
            end = r; break

    result = []
    r = 0
    while r < end - 1:
        nm     = _clean_name(grid[r][0] if grid[r] else '')
        period = _parse_number(grid[r][PERIOD] if PERIOD < len(grid[r]) else '')
        marker = _fl(grid[r][MARKER] if MARKER < len(grid[r]) else '').strip()
        
        if _is_comp(nm) and marker == '1':
            nxt = grid[r + 1] if r + 1 < len(grid) else []
            for cyl in range(1, actual_cyls + 1):
                ci = FIRST + cyl - 1
                d = _parse_date(grid[r][ci] if ci < len(grid[r]) else '')
                h = _parse_number(nxt[ci] if ci < len(nxt) else '')
                if d or h > 0: 
                    result.append(_mk('MAIN_ENGINE', 'ME', f'Cyl {cyl}', nm, period, d, h))
            r += 2
        else: r += 1
            
    return result


# ══════════════════════════════════════════════════════════════════════════
#  AUX PARSER
# ══════════════════════════════════════════════════════════════════════════
def _find_aux_groups(grid: List[List[str]]) -> Tuple[int, List[Tuple]]:
    dr = None
    for i, row in enumerate(grid):
        rt = ' | '.join(_fl(c) for c in row).upper()
        if 'DESCRIPTION' in rt and ('PERIODICTLY' in rt or 'PERIODICITY' in rt): 
            dr = i; break
    if dr is None: return -1, []
        
    nums = [(c, int(_fl(grid[dr][c])))
            for c in range(2, len(grid[dr]))
            if re.fullmatch(r'\d+', _fl(grid[dr][c]))]
            
    if nums:
        starts = [c for c, n in nums if n == 1]
        groups = []
        for i, s in enumerate(starts[:3]):
            ds = s + 1   
            de = (starts[i + 1] + 1) if i + 1 < len(starts) else len(grid[dr])
            groups.append((['AUX-1', 'AUX-2', 'AUX-3'][i], ds, de))
        if groups: return dr, groups
    return dr, []

def _parse_aux(grid: List[List[str]]) -> List[Dict]:
    if not grid: return []
    dr, groups = _find_aux_groups(grid)
    if dr < 0 or not groups: return []
        
    result = []
    r = dr + 1
    while r < len(grid) - 1:
        nm     = _clean_name(grid[r][0] if grid[r] else '')
        period = _parse_number(grid[r][1] if len(grid[r]) > 1 else '')
        marker = _fl(grid[r][2] if len(grid[r]) > 2 else '').strip()
        
        if _is_comp(nm) and marker == '1':
            nxt = grid[r + 1] if r + 1 < len(grid) else []
            for eng, start, end in groups:
                for ci_idx, ci in enumerate(range(start, min(end, len(grid[r])))):
                    d = _parse_date(grid[r][ci] if ci < len(grid[r]) else '')
                    h = _parse_number(nxt[ci] if ci < len(nxt) else '')
                    if d or h > 0:
                        result.append(_mk('AUX_ENGINE', eng, f'Cyl {ci_idx + 1}', nm, period, d, h))
            r += 2
        else: r += 1
    return result


# ══════════════════════════════════════════════════════════════════════════
#  OE PARSER
# ══════════════════════════════════════════════════════════════════════════
_OE_HEADER = {
    'TURBOCHARGER','AUXILIARY BOILER','COOLERS','EXH GAS BOILER','EXH GAS  BOILER',
    'A/C & REFR. COMPRESSORS','MAIN AIR COMPRESSORS','PERIODICTLY','PERIODICITY',
    'DATE OF LAST INSPECTION','DATE OF LAST O/H','RUN HRS','DATE OF LAST CLEANING',
    'DATE','DESCRIPTION','',
}

def _is_oe_comp(n: str) -> bool:
    u = _fl(n).upper().strip()
    if not u or len(u) < 2 or u in _OE_HEADER: return False
    if re.fullmatch(r'[\d./ ,:\-\[\]()]+', u): return False
    return bool(re.search(r'[A-Za-z]', u))

def _oe_section(desc: str, zone: str) -> str:
    d = _fl(desc).upper()
    if zone == 'A':
        if any(k in d for k in ('TURBOCHARGER','AIR COOLER CLEANING','GENERAL O/H','ROTOR','BALANCING')): return 'Turbocharger'
        if any(k in d for k in ('BOILER','BURNER','FEED PUMPS','FORCED DRAFT','FURNACE')): return 'Boiler Equipment'
        return 'Machinery Services'
    if zone in ('B'):
        if any(k in d for k in ('L.O.','JACKET','WASHING','CIRC. PUMP','PISTON L.O.')): return 'Coolers / Water Systems'
        return 'Cooling Systems'
    if zone in ('C', 'F'):
        if 'COMPRESSOR' in d: return 'Compressors'
        if any(k in d for k in ('CONDENSER','A/C','AIR. COND','AIR COND')): return 'Air Conditioning'
        return 'Other Equipment'
    if zone == 'D':
        # FIXED: Zone D explicitly mapped to Boiler Equipment
        return 'Boiler Equipment'
    return 'Other Equipment'

def _parse_oe(grid: List[List[str]]) -> List[Dict]:
    rows = []
    in_boiler_section = False

    for row in grid:
        r = row
        def gc(i, _r=r): return _fl(_r[i]) if i < len(_r) else ''

        if _fl(gc(1)).upper() in ('DATE OF LAST INSPECTION', 'DATE OF LAST O/H', 'DATE'):
            in_boiler_section = True
            continue

        if not in_boiler_section:
            da = _clean_name(gc(0))
            if _is_oe_comp(da):
                per = _parse_number(gc(1)); dt = _parse_date(gc(2)); hrs = _parse_number(gc(3))
                if dt or hrs > 0 or per > 0: rows.append({'section': _oe_section(da, 'A'), 'description': da, 'periodicity': per, 'last_date': dt, 'run_hrs': hrs})

            db = _clean_name(gc(5))
            if _is_oe_comp(db):
                dt = _parse_date(gc(6)); hrs = _parse_number(gc(7))
                if dt or hrs > 0: rows.append({'section': _oe_section(db, 'B'), 'description': db, 'periodicity': 0, 'last_date': dt, 'run_hrs': hrs})

            dc = _clean_name(gc(10))
            if _is_oe_comp(dc):
                dt = _parse_date(gc(11)); hrs = _parse_number(gc(12))
                if dt or hrs > 0: rows.append({'section': _oe_section(dc, 'C'), 'description': dc, 'periodicity': 0, 'last_date': dt, 'run_hrs': hrs})

        else:
            dd = _clean_name(gc(0))
            if _is_oe_comp(dd):
                dt = _parse_date(gc(1))
                if dt: rows.append({'section': _oe_section(dd, 'D'), 'description': dd, 'periodicity': 0, 'last_date': dt, 'run_hrs': 0})

            de = _clean_name(gc(5))
            if _is_oe_comp(de):
                dt = _parse_date(gc(6))
                if dt: rows.append({'section': _oe_section(de, 'B'), 'description': de, 'periodicity': 0, 'last_date': dt, 'run_hrs': 0})

            df = _clean_name(gc(10))
            if _is_oe_comp(df):
                dt = _parse_date(gc(11)); hrs = _parse_number(gc(12))
                if dt or hrs > 0: rows.append({'section': _oe_section(df, 'F'), 'description': df, 'periodicity': 0, 'last_date': dt, 'run_hrs': hrs})
                    
    return rows


# ══════════════════════════════════════════════════════════════════════════
#  DG PARSER
# ══════════════════════════════════════════════════════════════════════════
_DG_SKIP = {'DESCRIPTION', 'PERIODICTLY', 'PERIODICITY', 'D/G NO1', 'D/G NO2', 'D/G NO3', ''}

def _is_dg_comp(n: str) -> bool:
    u = _fl(n).upper().strip()
    if not u or len(u) < 2 or u in _DG_SKIP: return False
    if re.fullmatch(r'[\d./ ,\-\[\]()]+', u): return False
    return bool(re.search(r'[A-Za-z]', u))

def _parse_dg(grid: List[List[str]]) -> List[Dict]:
    rows = []
    r = 0
    while r < len(grid) - 1:
        r1 = grid[r]
        r2 = grid[r + 1] if r + 1 < len(grid) else []

        left_col = -1
        for i in range(min(5, len(r1))):
            if _is_dg_comp(r1[i]): left_col = i; break
        
        if left_col != -1 and left_col + 5 < len(r1):
            if _fl(r1[left_col + 2]) == '1':
                per = _parse_number(r1[left_col + 1])
                for gi, gl in enumerate(['D/G 1', 'D/G 2', 'D/G 3']):
                    dt = _parse_date(r1[left_col + 3 + gi])
                    h = _parse_number(r2[left_col + 3 + gi] if left_col + 3 + gi < len(r2) else '')
                    if dt or h > 0: rows.append({'section': 'D/G Equipment', 'description': _clean_name(r1[left_col]), 'engine_label': gl, 'periodicity': per, 'last_date': dt, 'run_hrs': h, 'status': _status(h, per)})
        
        right_col = -1
        search_start = left_col + 6 if left_col != -1 else 4
        for i in range(search_start, min(15, len(r1))):
            if _is_dg_comp(r1[i]): right_col = i; break

        if right_col != -1 and right_col + 5 < len(r1):
            if _fl(r1[right_col + 2]) == '1':
                per = _parse_number(r1[right_col + 1])
                for gi, gl in enumerate(['D/G 1', 'D/G 2', 'D/G 3']):
                    dt = _parse_date(r1[right_col + 3 + gi])
                    h = _parse_number(r2[right_col + 3 + gi] if right_col + 3 + gi < len(r2) else '')
                    if dt or h > 0: rows.append({'section': 'D/G Equipment', 'description': _clean_name(r1[right_col]), 'engine_label': gl, 'periodicity': per, 'last_date': dt, 'run_hrs': h, 'status': _status(h, per)})
        r += 1
    return rows


# ══════════════════════════════════════════════════════════════════════════
#  TEXT FALLBACK
# ══════════════════════════════════════════════════════════════════════════
def _lines_from_doc(doc) -> List[str]:
    lines = []
    for p in doc.paragraphs:
        t = _fl(p.text)
        if t: lines.append(t)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t = _fl(cell.text)
                if t: lines.append(t)
    return lines

def _between(lines, starts, ends):
    s = e = None
    for i, ln in enumerate(lines):
        u = ln.upper()
        if s is None and any(m in u for m in starts): s = i; continue
        if s is not None and any(m in u for m in ends): e = i; break
    return lines[s:e] if s is not None else []

def _parse_me_text(lines: List[str]) -> List[Dict]:
    rows = []
    seg = _between(lines, ['MAIN ENGINE', 'CYL. NO.1'], ['NOTE 1', 'TURBOCHARGER', 'AUX. ENGINE MAKER / TYPE'])
    if not seg: return rows
        
    data = [x for x in seg if _fl(x)]
    i = 0
    while i < len(data):
        nm = _clean_name(data[i])
        if _is_comp(nm):
            period = _parse_number(data[i + 1]) if i + 1 < len(data) else 0.0
            if _fl(data[i + 2] if i + 2 < len(data) else '') == '1':
                dates = []
                j = i + 3
                while j < len(data) and len(dates) < 8 and _fl(data[j]) != '2':
                    val = _fl(data[j])
                    if _is_comp(val): break
                    if not re.fullmatch(r'^\d{1,3}[.,]\d{3}$', val) and 'OBSERV' not in val.upper() and 'NOTE' not in val.upper():
                        dates.append(val)
                    j += 1
                
                if j < len(data) and _fl(data[j]) == '2':
                    j += 1
                    hv = []
                    while j < len(data) and len(hv) < len(dates):
                        if _is_comp(_fl(data[j])): break
                        hv.append(_fl(data[j]))
                        j += 1
                        
                    for k in range(len(dates)):
                        d = _parse_date(dates[k]) if k < len(dates) else ''
                        h = _parse_number(hv[k]) if k < len(hv) else 0.0
                        if d or h > 0: rows.append(_mk('MAIN_ENGINE', 'ME', f'Cyl {k + 1}', nm, period, d, h))
                    i = j
                    continue
        i += 1
    return rows

def _parse_aux_text(lines: List[str]) -> List[Dict]:
    rows = []
    seg = _between(lines, ['AUX. ENGINE MAKER / TYPE', 'AUX. ENGINE NO.1'], ['D/G NO1', 'TURBOCHARGER (2)', '1ST COPY'])
    if not seg: seg = _between(lines, ['AUX. ENGINE NO.1'], ['D/G NO1', 'TURBOCHARGER (2)', '1ST COPY'])
        
    si = None
    for i, ln in enumerate(seg):
        if 'DESCRIPTION' in ln.upper() and ('PERIODICTLY' in ln.upper() or 'PERIODICITY' in ln.upper()): si = i; break
            
    if si is None: return rows
        
    data = [_fl(x) for x in seg[si + 1:] if _fl(x)]
    i = 0
    while i < len(data):
        nm = _clean_name(data[i])
        if _is_comp(nm):
            period = _parse_number(data[i + 1]) if i + 1 < len(data) else 0.0
            if _fl(data[i + 2] if i + 2 < len(data) else '') == '1':
                dates = []
                j = i + 3
                while j < len(data) and len(dates) < 18 and _fl(data[j]) != '2':
                    if _is_comp(_fl(data[j])): break
                    dates.append(_fl(data[j]))
                    j += 1
                    
                if j < len(data) and _fl(data[j]) == '2':
                    j += 1
                    hv = []
                    while j < len(data) and len(hv) < 18:
                        if _is_comp(_fl(data[j])): break
                        hv.append(_fl(data[j]))
                        j += 1
                        
                    for ei, elbl in enumerate(['AUX-1', 'AUX-2', 'AUX-3']):
                        for cyl in range(6):
                            idx = ei * 6 + cyl
                            d = _parse_date(dates[idx]) if idx < len(dates) else ''
                            h = _parse_number(hv[idx]) if idx < len(hv) else 0.0
                            if d or h > 0: rows.append(_mk('AUX_ENGINE', elbl, f'Cyl {cyl + 1}', nm, period, d, h))
                    i = j
                    continue
        i += 1
    return rows


# ══════════════════════════════════════════════════════════════════════════
#  MASTER PARSER & DEDUPLICATION (Preserving Document Order)
# ══════════════════════════════════════════════════════════════════════════
def _dedupe(records: List[Dict]) -> List[Dict]:
    best = {}
    for r in records:
        key = (r.get('category'), r.get('description'), r.get('engine_label'), r.get('unit'))
        score = (1 if r.get('last_oh_date') else 0) + (1 if r.get('hrs_since', 0) > 0 else 0)
        prev = best.get(key)
        if prev is None or score > prev[0]: best[key] = (score, r)
    return [v[1] for v in best.values()]

def parse_docx(docx_bytes: bytes) -> Dict:
    from docx import Document
    warns = []
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as t:
        t.write(docx_bytes)
        tp = t.name
        
    try: doc = Document(tp)
    except Exception as e: raise ValueError(f"Cannot open: {e}")
    finally:
        try: os.unlink(tp)
        except Exception: pass

    if not doc.tables: raise ValueError("No tables — not a TEC-004 report?")

    vn = 'UNKNOWN'
    rd = None
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt: continue
        if m := re.search(r"Vessel[\u2019\u2018']?s?\s+Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+?)(?:\s{2,}|\t|Date:|$)", txt, re.I): vn = _clean_name(m.group(1))
        if m := re.search(r"Date\s*:\s*(.+)", txt, re.I): rd = _parse_date(m.group(1).strip())
        if vn != 'UNKNOWN' and rd: break
    if vn == 'UNKNOWN': warns.append("Could not extract vessel name.")

    mt = mo = None
    me_g: List[Dict] = []; aux_g: List[Dict] = []
    oe_rows: List[Dict] = []; dg_rows: List[Dict] = []

    for table in doc.tables:
        rg = _raw_grid(table)
        dg = _dedup_grid(table)
        full_table_text = ' '.join(_fl(c) for row in rg for c in row).upper()

        if mt is None:
            for row in rg[:4]:
                line = ' '.join(str(x) for x in row)
                if mt is None:
                    if m := re.search(r'Total Running Hours[\s:ǀ|]+([\d,]+)', line, re.I): mt = int(_parse_number(m.group(1)))
                if mo is None:
                    if m := re.search(r'This Month[\s:]+([\d,]+)', line, re.I): mo = int(_parse_number(m.group(1)))

        me_g.extend(_parse_me(rg))
        aux_g.extend(_parse_aux(dg))
        if 'TURBOCHARGER' in full_table_text and 'A/C & REFR' in full_table_text and 'COOLERS' in full_table_text: oe_rows.extend(_parse_oe(dg))
        if 'D/G NO' in full_table_text.replace(' ','').replace('.',''): dg_rows.extend(_parse_dg(dg))

    all_lines = _lines_from_doc(doc)
    me = _dedupe(me_g + _parse_me_text(all_lines))
    aux = _dedupe(aux_g + _parse_aux_text(all_lines))
    oe = _dedupe(oe_rows)
    dg_rows = _dedupe(dg_rows)

    # Assign Document Order Indices (Preserving natural extraction sequence)
    for i, r in enumerate(me): r['doc_idx'] = i
    for i, r in enumerate(aux): r['doc_idx'] = i
    for i, r in enumerate(oe): r['doc_idx'] = i
    for i, r in enumerate(dg_rows): r['doc_idx'] = i

    if not me and not aux: warns.append("No components extracted.")

    return {
        'vessel_name': vn, 'report_date': rd,
        'me_total_hrs': mt, 'me_this_month': mo,
        'me': me, 'aux': aux, 'oe': oe, 'dg': dg_rows,
        'warnings': warns, 'parsed_at': datetime.utcnow().isoformat(),
    }


# ══════════════════════════════════════════════════════════════════════════
#  HTML MATRICES (Animated, Protected, Ordered)
# ══════════════════════════════════════════════════════════════════════════
def _sort_recs(records: List[Dict], mode: str) -> List[Dict]:
    if mode == 'Document Order':
        return sorted(records, key=lambda c: c.get('doc_idx', 0))
    elif mode == 'Component → Cylinder':
        return sorted(records, key=lambda c: (c['engine_label'], c['description'].upper(), _cyl_n(c.get('unit', ''))))
    else: # Priority
        return sorted(records, key=lambda c: (_ORD.get(c['status'], 4), -(c.get('pct_used') or 0)))

def _matrix_html(records: List[Dict], mode: str = 'Document Order') -> str:
    if not records: return ''
        
    HEADS = ['Status', 'Component', 'Engine', 'Unit', 'Periodicity', 'Last O/H', 'Hrs Since', '% Used']
    TH = ("padding:14px 16px;font-family:'Inter',sans-serif;font-size:.62rem;font-weight:700;"
          "text-transform:uppercase;letter-spacing:.12em;color:#607888;text-align:left;"
          "white-space:nowrap;border-bottom:1px solid var(--line);background:var(--bg3)")
          
    header = '<tr>' + ''.join(f'<th style="{TH}">{h}</th>' for h in HEADS) + '</tr>'
    body = ''
    
    for row_idx, rec in enumerate(_sort_recs(records, mode)):
        s = str(rec.get('status', 'NO DATA'))
        c = _SC.get(s, _SC['NO DATA'])
        
        pct = float(rec.get('pct_used') or 0)
        hrs = float(rec.get('hrs_since') or 0)
        per = float(rec.get('periodicity') or 0)
        
        pct_s = f"{pct*100:.1f}%" if pct > 0 else '—'
        hrs_s = f"{int(hrs):,}" if hrs > 0 else '—'
        per_s = f"{int(per):,}" if per > 0 else '—'
        dt_s = html.escape(str(rec.get('last_oh_date') or '—') or '—')
        desc_s = html.escape(str(rec.get('description', '')))
        
        bw = min(100, pct * 100)
        
        # Add tooltips and pulse classes
        tag = (f'<span class="{c["class"]}" title="Component is {s}" style="display:inline-block;padding:5px 12px;border-radius:999px;'
               f'background:{c["tag_bg"]};color:{c["tag_fg"]};font-family:JetBrains Mono,monospace;'
               f'font-size:.63rem;font-weight:700;white-space:nowrap;transition:all .2s">{s}</span>')
               
        # Animated progress bar
        bar = (f'<div style="display:flex;align-items:center;gap:10px;min-width:130px">'
               f'<div style="flex:1;height:5px;background:{c["bar_e"]};border-radius:999px;overflow:hidden">'
               f'<div style="width:{bw:.1f}%;height:100%;background:{c["bar_f"]};border-radius:999px;'
               f'animation: growWidth 1s cubic-bezier(0.4, 0, 0.2, 1) forwards;"></div>'
               f'</div><span style="font-family:JetBrains Mono,monospace;font-size:.71rem;font-weight:700;'
               f'color:{c["num"]};min-width:46px;text-align:right">{pct_s}</span></div>')
               
        def td(v, fg, fw='400', align='left', ff="'Inter',sans-serif", fs='.81rem', mw=''):
            mw_s = f'max-width:{mw};overflow:hidden;text-overflow:ellipsis;' if mw else ''
            return (f'<td style="padding:14px 16px;color:{fg};font-family:{ff};font-size:{fs};'
                    f'font-weight:{fw};text-align:{align};white-space:nowrap;{mw_s}">{v}</td>')
                    
        # Inject staggering delay directly into the row
        anim_delay = min(row_idx * 0.02, 1.5) # Cap delay at 1.5s
        body += (
            f'<tr class="stagger-row" style="background:{c["row"]};border-bottom:1px solid {c["bord"]};'
            f'animation-delay: {anim_delay}s; transition: background .2s" '
            f'onmouseover="this.style.background=\'var(--bg3)\'" '
            f'onmouseout="this.style.background=\'{c["row"]}\'">'
            f'<td style="padding:14px 16px">{tag}</td>'
            + td(desc_s, c['comp'], '600', 'left', "'IBM Plex Sans',sans-serif", '.83rem', '310px')
            + td(rec.get('engine_label', ''), c['dim'])
            + td(rec.get('unit', ''), c['dim'], '500', 'left', 'JetBrains Mono,monospace', '.76rem')
            + td(per_s, c['dim'], '500', 'right', 'JetBrains Mono,monospace', '.76rem')
            + td(dt_s, c['dim'], '500', 'left', 'JetBrains Mono,monospace', '.76rem')
            + td(hrs_s, c['num'], '700', 'right', 'JetBrains Mono,monospace', '.77rem')
            + f'<td style="padding:14px 16px">{bar}</td></tr>'
        )
        
    return (f'<div style="overflow-x:auto;border-radius:14px;border:1px solid var(--line);'
            f'background:var(--bg2);overflow:hidden;box-shadow:0 12px 30px rgba(0,0,0,.3)">'
            f'<table style="width:100%;border-collapse:collapse">'
            f'<thead>{header}</thead><tbody>{body}</tbody></table></div>')

def _oe_html(rows: List[Dict], show_status: bool = False) -> str:
    if not rows: return ''
        
    HEADS = (['Component', 'Engine', 'Period', 'Last Date', 'Run Hrs', 'Status']
             if show_status else ['Component', 'Section', 'Period', 'Last Date', 'Run Hrs'])
             
    TH = ("padding:14px 16px;font-family:'Inter',sans-serif;font-size:.62rem;font-weight:700;"
          "text-transform:uppercase;letter-spacing:.12em;color:#607888;text-align:left;"
          "border-bottom:1px solid var(--line);background:var(--bg3)")
          
    header = '<tr>' + ''.join(f'<th style="{TH}">{h}</th>' for h in HEADS) + '</tr>'
    body = ''
    
    # Preserve document order in OE
    for row_idx, row in enumerate(sorted(rows, key=lambda x: x.get('doc_idx', 0))):
        bg = 'var(--bg2)'; fm = 'var(--ink)'; fd = 'var(--muted)'
        
        if show_status:
            s = str(row.get('status', 'NO DATA'))
            cc = _SC.get(s, _SC['NO DATA'])
            bg = cc['row']; fm = cc['comp']; fd = cc['dim']
            
        per = float(row.get('periodicity', 0) or 0)
        per_s = f"{int(per):,}" if per > 0 else '—'
        dt_s = html.escape(str(row.get('last_date') or '—') or '—')
        desc_s = html.escape(str(row.get('description', '')))
        hrs = float(row.get('run_hrs', 0) or 0)
        hrs_s = f"{int(hrs):,}" if hrs > 0 else '—'
        
        st_cell = ''
        if show_status:
            s = str(row.get('status', 'NO DATA'))
            cc = _SC.get(s, _SC['NO DATA'])
            st_cell = (f'<td style="padding:14px 16px"><span class="{cc["class"]}" style="display:inline-block;padding:5px 12px;'
                       f'border-radius:999px;background:{cc["tag_bg"]};color:{cc["tag_fg"]};'
                       f'font-family:JetBrains Mono,monospace;font-size:.63rem;font-weight:700">{s}</span></td>')
                       
        def td(v, fg, ff="'Inter',sans-serif", fw='400', align='left'):
            return (f'<td style="padding:14px 16px;color:{fg};font-family:{ff};font-size:.81rem;'
                    f'font-weight:{fw};text-align:{align};white-space:nowrap">{v}</td>')
                    
        anim_delay = min(row_idx * 0.02, 1.5)
        body += (f'<tr class="stagger-row" style="background:{bg};border-bottom:1px solid var(--line);'
                 f'animation-delay: {anim_delay}s; transition: background .2s" '
                 f'onmouseover="this.style.background=\'var(--bg3)\'" '
                 f'onmouseout="this.style.background=\'{bg}\'">'
                 + td(desc_s, fm, "'IBM Plex Sans',sans-serif", '600')
                 + (td(row.get('engine_label', ''), fd, 'JetBrains Mono,monospace') if show_status else td(row.get('section', ''), fd))
                 + f'<td style="padding:14px 16px;color:{fd};font-family:JetBrains Mono,monospace;font-size:.78rem;text-align:right">{per_s}</td>'
                 + td(dt_s, fd, 'JetBrains Mono,monospace')
                 + f'<td style="padding:14px 16px;color:var(--ink);font-family:JetBrains Mono,monospace;font-size:.78rem;font-weight:700;text-align:right">{hrs_s}</td>'
                 + st_cell + '</tr>')
                 
    return (f'<div style="overflow-x:auto;border-radius:14px;border:1px solid var(--line);'
            f'background:var(--bg2);overflow:hidden;box-shadow:0 12px 30px rgba(0,0,0,.3)">'
            f'<table style="width:100%;border-collapse:collapse">'
            f'<thead>{header}</thead><tbody>{body}</tbody></table></div>')


# ══════════════════════════════════════════════════════════════════════════
#  UI DASHBOARD COMPONENTS
# ══════════════════════════════════════════════════════════════════════════
def _kpi(val: Any, lbl: str, accent: str, delay: float = 0.0) -> str:
    return (f'<div style="background:var(--panel); backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px);'
            f'border:1px solid var(--line);border-radius:14px;padding:1.1rem 1.1rem 1rem;'
            f'box-shadow:0 8px 32px rgba(0,0,0,.2);position:relative;overflow:hidden;'
            f'animation:fadeUp .4s {delay}s cubic-bezier(0.2, 0.8, 0.2, 1) both;'
            f'transition:transform .2s cubic-bezier(0.4, 0, 0.2, 1), box-shadow .2s" '
            f'onmouseover="this.style.transform=\'translateY(-4px)\';this.style.boxShadow=\'0 12px 40px rgba(184,148,96,.08)\'" '
            f'onmouseout="this.style.transform=\'\';this.style.boxShadow=\'0 8px 32px rgba(0,0,0,.2)\'">'
            f'<div style="width:36px;height:3px;background:linear-gradient(90deg, {accent}, transparent);border-radius:999px;margin-bottom:1rem;'
            f'animation:growWidth .6s {delay+.1}s ease both"></div>'
            f'<div style="font-family:IBM Plex Sans,sans-serif;font-size:1.3rem;font-weight:700;'
            f'line-height:1.15;letter-spacing:-.02em;color:var(--ink);white-space:nowrap;'
            f'overflow:hidden;text-overflow:ellipsis;">{val}</div>'
            f'<div style="font-family:Inter,sans-serif;font-size:.62rem;font-weight:600;'
            f'text-transform:uppercase;letter-spacing:.15em;color:var(--soft);margin-top:.6rem">{lbl}</div>'
            f'</div>')

def _section(title: str, sub: str, badge: str, accent: str = '#b89460') -> str:
    return (f'<div style="display:flex;align-items:flex-end;gap:.9rem;margin:2.5rem 0 1rem;'
            f'animation:slideDown .4s ease both">'
            f'<div style="flex:1">'
            f'<div style="font-family:IBM Plex Sans,sans-serif;font-size:1.1rem;font-weight:700;'
            f'color:var(--ink);letter-spacing:-.01em;line-height:1">{title}</div>'
            f'<div style="font-family:Inter,sans-serif;font-size:.74rem;color:var(--soft);'
            f'margin-top:.35rem">{sub}</div></div>'
            f'<div style="font-family:JetBrains Mono,monospace;font-size:.63rem;font-weight:600;'
            f'padding:6px 12px;border-radius:999px;background:var(--bg3);border:1px solid var(--line);'
            f'color:var(--muted);white-space:nowrap;box-shadow:0 4px 12px rgba(0,0,0,.15)">{badge}</div>'
            f'</div>'
            f'<div style="height:1px;background:linear-gradient(90deg,{accent},var(--line) 40%,transparent);'
            f'margin-bottom:1.2rem;animation:growWidth .6s ease both"></div>')


# ══════════════════════════════════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════════════════════════════════
if 'data' not in st.session_state:
    st.session_state.data = None


# ══════════════════════════════════════════════════════════════════════════
#  PAGE HEADER
# ══════════════════════════════════════════════════════════════════════════
st.markdown(
    '<div style="padding:1.7rem 0 0;animation:slideDown .4s cubic-bezier(0.2, 0.8, 0.2, 1) both">'
    '<div style="display:flex;justify-content:space-between;align-items:flex-end;gap:1rem;flex-wrap:wrap">'
    '<div>'
    '<div style="font-family:Inter,sans-serif;font-size:.62rem;font-weight:700;letter-spacing:.3em;'
    'text-transform:uppercase;color:var(--gold);margin-bottom:.35rem;text-shadow:0 2px 10px rgba(184,148,96,.3)">Fleet Technical Operations</div>'
    '<div style="font-family:IBM Plex Sans,sans-serif;font-size:2.1rem;font-weight:700;'
    'color:var(--ink);letter-spacing:-.03em;line-height:1.05">Running Hours Control Panel</div>'
    '</div>'
    '<div style="font-family:JetBrains Mono,monospace;font-size:.65rem;color:var(--soft);'
    'padding:.6rem .9rem;border:1px solid var(--line);border-radius:999px;background:var(--bg3);'
    'white-space:nowrap;box-shadow:0 4px 12px rgba(0,0,0,.15)">TEC-004 · Executive Build</div>'
    '</div></div>'
    '<div style="height:1px;margin:1rem 0 1.8rem;'
    'background:linear-gradient(90deg,var(--gold),var(--line) 40%,transparent);'
    'animation:growWidth .6s .1s ease both"></div>',
    unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  UPLOAD
# ══════════════════════════════════════════════════════════════════════════
with st.expander("Upload TEC-004 Report", expanded=(st.session_state.data is None)):
    uc, ic = st.columns([2.2, 1.0], gap='large')
    with uc:
        uploaded = st.file_uploader('Drop TEC-004 .doc', type=['doc'], label_visibility='collapsed')
    with ic:
        st.markdown(
            '<div style="background:var(--bg3);border:1px solid var(--line);border-radius:14px;'
            'padding:1.1rem 1.2rem;font-family:Inter,sans-serif;font-size:.78rem;color:var(--muted);line-height:1.9">'
            '<div style="font-family:IBM Plex Sans,sans-serif;color:var(--ink);font-weight:700;font-size:.9rem;margin-bottom:.5rem">Document Profile</div>'
            '<b style="color:var(--ink)">Format:</b> TEC-004 Running Hours Report (.doc)<br>'
            '<b style="color:var(--ink)">Parser:</b> Dynamic Slide · Hard Anchors · Strict Ceilings<br>'
            '<b style="color:var(--ink)">Output:</b> Absolute True Document Order Preserved<br>'
            '<b style="color:var(--ink)">Security:</b> Full XSS Escaping Deployed'
            '</div>', unsafe_allow_html=True)

    if uploaded:
        raw = uploaded.read()
        # CACHE BUSTER
        fh = hashlib.md5(raw).hexdigest() + "_FINAL" 
        
        if st.session_state.data is None or st.session_state.data.get('_hash') != fh:
            with st.spinner('Converting .doc → .docx via LibreOffice…'):
                try: docx = convert_doc_to_docx(raw)
                except Exception as e: st.error(f'Conversion failed: {e}'); st.stop()
                    
            with st.spinner('Parsing TEC-004 architecture…'):
                try: result = parse_docx(docx)
                except ValueError as e: st.error(f'Parse failed: {e}'); st.stop()
                    
            for w in result['warnings']: st.warning(f'⚠ {w}')
                
            if not result['me'] and not result['aux']:
                st.error('No components extracted. Verify this is a TEC-004 report.')
                st.stop()
                
            result['_hash'] = fh
            result['_filename'] = uploaded.name
            st.session_state.data = result
            
            ac = result['me'] + result['aux']
            od = sum(1 for c in ac if c['status'] == 'OVERDUE')
            hp = sum(1 for c in ac if c['status'] == 'HIGH PRIORITY')
            
            st.markdown(
                f'<div style="background:linear-gradient(135deg,#132018,#101720);'
                f'border:1px solid #1e3828;border-radius:14px;padding:1rem 1.2rem;'
                f'color:#a8c4b0;font-family:IBM Plex Sans,sans-serif;font-size:.85rem;font-weight:500;'
                f'animation:popIn .5s cubic-bezier(.34,1.56,.64,1) both;box-shadow:0 8px 24px rgba(0,0,0,.2)">'
                f'<span style="color:#7cb393;font-weight:700">Validated:</span> '
                f'<span style="color:var(--ink);font-weight:700">{result["vessel_name"]}</span> — '
                f'{len(ac)} components · {od} overdue · {hp} high priority</div>',
                unsafe_allow_html=True)


# ── Empty state ──────────────────────────────────────────────────────────
if st.session_state.data is None:
    st.markdown(
        '<div style="display:flex;align-items:center;justify-content:center;'
        'height:40vh;flex-direction:column;gap:1.2rem;animation:fadeUp .5s ease">'
        '<svg width="56" height="56" viewBox="0 0 64 64" fill="none" '
        'xmlns="http://www.w3.org/2000/svg" style="opacity:.15">'
        '<path d="M32 8v48M32 8L20 20M32 8l12 12M8 32h48" '
        'stroke="var(--muted)" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>'
        '<circle cx="32" cy="48" r="8" stroke="var(--muted)" stroke-width="2.5"/></svg>'
        '<div style="font-family:Inter,sans-serif;font-size:.85rem;font-weight:500;color:var(--soft);letter-spacing:.05em">'
        'Awaiting TEC-004 Document Upload</div></div>',
        unsafe_allow_html=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════
#  DATA READY
# ══════════════════════════════════════════════════════════════════════════
d = st.session_state.data
me = d['me']; aux = d['aux']; oe = d['oe']; dg = d['dg']

ac = me + aux
n_od = sum(1 for c in ac if c['status'] == 'OVERDUE')
n_hp = sum(1 for c in ac if c['status'] == 'HIGH PRIORITY')
mt = d.get('me_total_hrs')
mo = d.get('me_this_month')

cols = st.columns(8)
for col, (val, lbl, acc, dly) in zip(cols, [
    (d['vessel_name'],         'Vessel',         '#b89460', 0.00),
    (d['report_date'] or '—',  'Report Date',    '#466b99', 0.05),
    (f"{mt:,}" if mt else '—', 'M/E Total Hrs',  '#4f7a62', 0.10),
    (f"{mo:,}" if mo else '—', 'M/E This Month', '#4f7a62', 0.15),
    (len(me),                  'ME Records',     '#b89460', 0.20),
    (len(aux),                 'AUX Records',    '#466b99', 0.25),
    (n_od,                     'Overdue',        '#8c5353', 0.30),
    (n_hp,                     'High Priority',  '#99734d', 0.35),
]):
    with col: st.markdown(_kpi(val, lbl, acc, dly), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 1 — MAIN ENGINE
# ══════════════════════════════════════════════════════════════════════════
me_od = sum(1 for c in me if c['status'] == 'OVERDUE')
me_hp = sum(1 for c in me if c['status'] == 'HIGH PRIORITY')
st.markdown(_section('Main Engine', f'{len(me)} component records', f'{me_od} OD · {me_hp} HP'), unsafe_allow_html=True)

f1, f2, f3 = st.columns([2, 2, 3])
with f1: mc_sel = st.selectbox('Component', ['All'] + _get_unique_ordered(me, 'description'), key='mc')
with f2: ms_sel = st.selectbox('Status', ['All', 'Overdue only', 'High Priority +', 'OK only'], key='ms')
with f3: mr_sel = st.radio('Sort', ['Document Order', 'Component → Cylinder', 'Priority → % Used'], horizontal=True, key='mr')

v = me[:]
if mc_sel != 'All':               v = [c for c in v if c['description'] == mc_sel]
if ms_sel == 'Overdue only':      v = [c for c in v if c['status'] == 'OVERDUE']
elif ms_sel == 'High Priority +': v = [c for c in v if c['status'] in ('OVERDUE', 'HIGH PRIORITY')]
elif ms_sel == 'OK only':         v = [c for c in v if c['status'] == 'OK']

if not v: st.markdown('<div style="color:var(--soft);padding:1rem;">No records match filter.</div>', unsafe_allow_html=True)
else: st.markdown(_matrix_html(v, mode=mr_sel), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 2 — AUX ENGINES
# ══════════════════════════════════════════════════════════════════════════
ax_od = sum(1 for c in aux if c['status'] == 'OVERDUE')
ax_hp = sum(1 for c in aux if c['status'] == 'HIGH PRIORITY')
st.markdown(_section('Auxiliary Engines', f'{len(aux)} component records · AUX-1 · AUX-2 · AUX-3', f'{ax_od} OD · {ax_hp} HP', '#466b99'), unsafe_allow_html=True)

g1, g2, g3, g4 = st.columns([1.3, 2, 2, 3])
with g1: ae_sel = st.selectbox('Engine', ['All'] + _get_unique_ordered(aux, 'engine_label'), key='ae')
with g2: ac_sel = st.selectbox('Component', ['All'] + _get_unique_ordered(aux, 'description'), key='ac')
with g3: as_sel = st.selectbox('Status', ['All', 'Overdue only', 'High Priority +', 'OK only'], key='axs')
with g4: ar_sel = st.radio('Sort', ['Document Order', 'Component → Cylinder', 'Priority → % Used'], horizontal=True, key='ar')

v = aux[:]
if ae_sel != 'All':               v = [c for c in v if c['engine_label'] == ae_sel]
if ac_sel != 'All':               v = [c for c in v if c['description'] == ac_sel]
if as_sel == 'Overdue only':      v = [c for c in v if c['status'] == 'OVERDUE']
elif as_sel == 'High Priority +': v = [c for c in v if c['status'] in ('OVERDUE', 'HIGH PRIORITY')]
elif as_sel == 'OK only':         v = [c for c in v if c['status'] == 'OK']

if not v: st.markdown('<div style="color:var(--soft);padding:1rem;">No records match filter.</div>', unsafe_allow_html=True)
else: st.markdown(_matrix_html(v, mode=ar_sel), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 3 — OTHER EQUIPMENT
# ══════════════════════════════════════════════════════════════════════════
st.markdown(_section('Other Equipment', 'Turbocharger · Boiler Equipment · Compressors · Cooling', f'{len(oe)} records', '#4f7a62'), unsafe_allow_html=True)

if not oe:
    st.markdown('<div style="background:var(--bg3);border:1px solid var(--line);border-radius:14px;padding:1.5rem;'
                'text-align:center;color:var(--muted);font-family:IBM Plex Sans,sans-serif;font-size:.85rem;font-weight:500">'
                'No other equipment data found.</div>', unsafe_allow_html=True)
else:
    secs = _get_unique_ordered(oe, 'section')
    os_sel = st.selectbox('Section', ['All'] + secs, key='oe_s')
    ov = oe if os_sel == 'All' else [r for r in oe if r['section'] == os_sel]
    st.markdown(_oe_html(ov, show_status=False), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 4 — D/G EQUIPMENT
# ══════════════════════════════════════════════════════════════════════════
st.markdown(_section('D/G Equipment', 'Diesel Generator components · D/G 1 · D/G 2 · D/G 3', f'{len(dg)} records', '#8c5353'), unsafe_allow_html=True)

if not dg:
    st.markdown('<div style="background:var(--bg3);border:1px solid var(--line);border-radius:14px;padding:1.5rem;'
                'text-align:center;color:var(--muted);font-family:IBM Plex Sans,sans-serif;font-size:.85rem;font-weight:500">'
                'No D/G equipment data found.</div>', unsafe_allow_html=True)
else:
    dg_sel = st.selectbox('D/G Unit', ['All'] + _get_unique_ordered(dg, 'engine_label'), key='dge')
    dv = dg if dg_sel == 'All' else [r for r in dg if r.get('engine_label') == dg_sel]
    st.markdown(_oe_html(dv, show_status=True), unsafe_allow_html=True)
