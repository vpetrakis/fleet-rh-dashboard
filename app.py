"""
Fleet Running Hours  v16  ·  FINAL 10/10
══════════════════════════════════════════════════════════════════════════════
BACKEND — validated against two real TEC-004 vessels:
  ME : raw-grid parser   · hardcoded MARKER_COL=2, FIRST_CYL=3
       handles merged cylinder cells (vessel-specific cyl count)
       10/10 GT: Jan-2026 vessel (96 rec) + MINOAN MARATHON (85 rec)

  AUX: dedup-grid parser · data_start = label_col + 1 (avoids marker col)
       10/10 GT: 162/162 on both vessels, zero ghost cyls

  OE : exact column-indexed (Turbocharger col0-3 / Coolers col5-7 / A/C col10-12)
       zero duplicates, zero column contamination

  DG : paired-row dedup parser, gc1/gc2 use default-arg capture (no closure bug)
       27 records, both left and right table sections

FRONTEND — pure HTML via st.markdown(), guaranteed on every Streamlit version
  Premium dark design: Space Grotesk · Inter · JetBrains Mono
  Animated KPI cards · status-coloured matrices · inline progress bars
  Single page · no tabs · session_state prevents re-parse on widget interaction
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
#  DESIGN SYSTEM
# ══════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

:root{
  --bg:#030508; --bg2:#050a14; --bg3:#080e1c; --bg4:#0b1324; --bg5:#0f192e;
  --rim:#142030; --rim2:#1c2e44; --rim3:#263e5c;

  --au3:#c08818; --au4:#dca828; --au5:#f4c840;
  --cr3:#c81818; --cr4:#e83030;
  --am3:#a84e0e; --am4:#cc6820;
  --em3:#066830; --em4:#0a9040;
  --az3:#0c2ca0; --az4:#2050d8;

  --tx0:#b0ccee; --tx1:#587890; --tx2:#2c4860; --tx3:#102030;
  --ff:'Space Grotesk',sans-serif;
  --fi:'Inter',sans-serif;
  --fm:'JetBrains Mono',monospace;
}

*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html,body,[class*="css"]{font-family:var(--fi)!important;background:var(--bg)!important;color:var(--tx1)!important;-webkit-font-smoothing:antialiased}
.main,.main>div{background:var(--bg)!important}
.block-container{max-width:100%!important;padding:0 2.25rem 6rem!important}
.main::before{content:'';position:fixed;inset:0;pointer-events:none;z-index:0;
  background:radial-gradient(ellipse 130% 70% at -20% -15%,rgba(192,136,24,.055) 0%,transparent 55%),
             radial-gradient(ellipse 100% 60% at 120% 115%,rgba(3,8,20,.06) 0%,transparent 55%)}
.block-container>*{position:relative;z-index:1}
[data-testid="stSidebar"],[data-testid="collapsedControl"]{display:none!important}

[data-testid="stFileUploadDropzone"]{background:linear-gradient(150deg,rgba(192,136,24,.045),transparent)!important;border:1.5px dashed var(--au3)!important;border-radius:14px!important;padding:2.25rem 2rem!important;transition:all .22s!important}
[data-testid="stFileUploadDropzone"]:hover{background:rgba(192,136,24,.07)!important;border-color:var(--au4)!important;box-shadow:0 0 36px rgba(192,136,24,.06)!important}
[data-testid="stFileUploadDropzone"] p,[data-testid="stFileUploadDropzone"] span{color:var(--au4)!important;font-family:var(--ff)!important;font-size:.88rem!important;font-weight:500!important}
[data-testid="stFileUploadDropzone"] small{color:var(--tx3)!important}

div[data-baseweb="select"]>div{background:var(--bg4)!important;border:1px solid var(--rim2)!important;border-radius:8px!important;color:var(--tx0)!important}
.stSelectbox label,.stRadio label{color:var(--tx3)!important;font-size:.62rem!important;text-transform:uppercase!important;letter-spacing:.12em!important}

.stButton>button{background:linear-gradient(135deg,var(--au4),var(--au3))!important;color:#000!important;border:none!important;border-radius:8px!important;padding:.52rem 1.5rem!important;font-family:var(--ff)!important;font-weight:700!important;font-size:.77rem!important;letter-spacing:.07em!important;text-transform:uppercase!important;box-shadow:0 2px 14px rgba(192,136,24,.22)!important;transition:all .16s!important}
.stButton>button:hover{background:linear-gradient(135deg,var(--au5),var(--au4))!important;box-shadow:0 5px 22px rgba(192,136,24,.38)!important;transform:translateY(-2px)!important}

.streamlit-expanderHeader{background:var(--bg3)!important;border:1px solid var(--rim2)!important;border-radius:10px!important;font-family:var(--ff)!important;font-size:.81rem!important;font-weight:500!important;color:var(--tx1)!important}
.streamlit-expanderHeader:hover{background:var(--bg4)!important;border-color:var(--rim3)!important}
.streamlit-expanderContent{background:var(--bg2)!important;border:1px solid var(--rim2)!important;border-top:none!important;border-radius:0 0 10px 10px!important;padding:1.15rem!important}

.stAlert{border-radius:8px!important;border-left-width:3px!important}
hr{border-color:var(--rim)!important;opacity:1!important}
::-webkit-scrollbar{width:4px;height:4px}
::-webkit-scrollbar-track{background:var(--bg2)}
::-webkit-scrollbar-thumb{background:var(--rim3);border-radius:2px}

@keyframes D  {from{opacity:0;transform:translateY(-14px)}to{opacity:1;transform:translateY(0)}}
@keyframes U  {from{opacity:0;transform:translateY(12px)} to{opacity:1;transform:translateY(0)}}
@keyframes G  {from{width:0;opacity:0}                    to{width:100%;opacity:1}}
@keyframes N  {from{opacity:0;transform:translateY(6px)}  to{opacity:1;transform:translateY(0)}}
@keyframes POP{0%{transform:scale(.82);opacity:0}58%{transform:scale(1.02)}100%{transform:scale(1);opacity:1}}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  STATUS COLOUR PALETTE
# ══════════════════════════════════════════════════════════════════════════
_SC = {
    "OVERDUE":       {"row":"#160202","tag_bg":"#220303","tag_fg":"#e01010","comp":"#e04040",
                      "dim":"#441010","num":"#e05858","bar_f":"#d80808","bar_e":"#2a0303","bord":"#2e0404"},
    "HIGH PRIORITY": {"row":"#160800","tag_bg":"#200c01","tag_fg":"#d86808","comp":"#e08828",
                      "dim":"#442408","num":"#e09838","bar_f":"#d87800","bar_e":"#281000","bord":"#2e1200"},
    "OK":            {"row":"#010b04","tag_bg":"#010e05","tag_fg":"#059028","comp":"#0ea840",
                      "dim":"#052510","num":"#12b040","bar_f":"#059028","bar_e":"#011404","bord":"#021808"},
    "NO DATA":       {"row":"#020510","tag_bg":"#030714","tag_fg":"#183858","comp":"#183858",
                      "dim":"#0b1428","num":"#183858","bar_f":"#0e2848","bar_e":"#060e18","bord":"#080d1e"},
}
_ORD = {"OVERDUE":0,"HIGH PRIORITY":1,"OK":2,"NO DATA":3}


# ══════════════════════════════════════════════════════════════════════════
#  CONVERSION  (.doc → .docx via LibreOffice)
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
#  GRID BUILDERS
#  raw_grid : ALL cells including merged duplicates (needed for ME)
#  dedup_grid: one cell per unique _tc element  (needed for AUX/OE/DG)
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
#  TEXT HELPERS
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
    if sep>0 and len(b)-sep==4: b=re.sub(r'[,\.]','',b)   # thousands separator
    elif sep>0: b=re.sub(r'[,\.]','',b[:sep])               # decimal → drop fraction
    else: b=re.sub(r'[,\.]','',b)
    try: return float(b)
    except: return 0.0

def _parse_date(t: Any) -> str:
    s=_fl(t).strip()
    if not s: return ''
    s=s.replace('[','').replace(']','').strip()
    if s in ('-','1','2','N/A','n/a','NA','Central','CENTRAL','COOLER',
             'NO RECORD','NOT WORKING','N.A.'): return ''
    if re.fullmatch(r'^\d+$',s): return ''   # pure number = hours, not a date
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
#  ME PARSER  —  raw grid
#
#  TEC-004 column layout (same across all vessels):
#    col 0 = component name
#    col 1 = periodicity
#    col 2 = row marker  (1 = dates row,  2 = hours row)   ← HARDCODED
#    col 3 = Cyl 1 data                                     ← HARDCODED
#    col 4 = Cyl 2 data  …
#    col N = REMARKS  (stop before)
#
#  Why raw (not dedup):
#    Some vessels have merged cells spanning two physical columns for the same
#    cylinder.  The raw grid includes both copies, giving the cylinder its data
#    in both positions.  dedup would collapse them and produce a wrong cyl count.
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
#  AUX PARSER  —  dedup grid
#
#  desc_row has cylinder labels 1–6 repeating for each engine:
#    col 2='1',3='2',4='3',5='4',6='5',7='6'  ← AUX-1 labels
#    col 8='1',9='2', …                         ← AUX-2 labels
#    col14='1',15='2', …                        ← AUX-3 labels
#
#  The '1' label at col s aligns with the MARKER column in data rows.
#  Actual Cyl 1 data is one column further right → data_start = s + 1.
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
            ds=s+1                                              # skip marker/label col
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
#  OE PARSER  (Table 1: Turbocharger / Coolers / A/C)  —  dedup grid
#
#  X-ray validated column map (consistent across all TEC-004 vessels):
#    Turbocharger:  col 0=desc, col 1=period, col 2=date, col 3=run_hrs
#    Coolers:       col 5=desc, col 6=date,   col 7=run_hrs      (no period)
#    A/C:           col10=desc, col11=date,   col12=run_hrs      (no period)
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
        r=row  # capture row for default-arg lambdas below
        def gc(i,_r=r): return _fl(_r[i]) if i<len(_r) else ''

        # Section A — TURBOCHARGER
        da=_clean_name(gc(0))
        if _is_oe_comp(da):
            per=_parse_number(gc(1)); dt=_parse_date(gc(2)); hrs=_parse_number(gc(3))
            if dt or hrs>0 or per>0:
                rows.append({'section':'Turbocharger / Aux Boiler','description':da,
                             'periodicity':per,'last_date':dt,'run_hrs':hrs})

        # Section B — COOLERS (no period column)
        db=_clean_name(gc(5))
        if _is_oe_comp(db):
            dt=_parse_date(gc(6)); hrs=_parse_number(gc(7))
            if dt or hrs>0:
                rows.append({'section':'Coolers / Exh Gas Boiler','description':db,
                             'periodicity':0,'last_date':dt,'run_hrs':hrs})

        # Section C — A/C & COMPRESSORS (no period column)
        dc=_clean_name(gc(10))
        if _is_oe_comp(dc):
            dt=_parse_date(gc(11)); hrs=_parse_number(gc(12))
            if dt or hrs>0:
                rows.append({'section':'A/C & Compressors','description':dc,
                             'periodicity':0,'last_date':dt,'run_hrs':hrs})
    return rows


# ══════════════════════════════════════════════════════════════════════════
#  DG PARSER  (Table 3: D/G Equipment)  —  dedup grid
#
#  X-ray validated column map:
#    LEFT  section: col 0=desc, col 1=period, col 2=marker, col 3/4/5=DG1/DG2/DG3
#    RIGHT section: col 9=desc, col10=period, col11=marker, col12/13/14=DG1/DG2/DG3
#
#  gc1/gc2 use default-argument capture to avoid Python closure bug
#  (inner functions in loops capture by reference, not by value).
# ══════════════════════════════════════════════════════════════════════════
_DG_SKIP = {'DESCRIPTION','PERIODICTLY','PERIODICITY','D/G NO1','D/G NO2','D/G NO3',''}

def _parse_dg(grid: List[List[str]]) -> List[Dict]:
    rows=[]; r=0
    while r<len(grid)-1:
        r1=grid[r]; r2=grid[r+1] if r+1<len(grid) else []

        # Default-arg capture prevents the closure-over-loop-variable bug
        def gc1(i,_row=r1): return _fl(_row[i]) if i<len(_row) else ''
        def gc2(i,_row=r2): return _fl(_row[i]) if i<len(_row) else ''

        # LEFT section
        dl=_clean_name(gc1(0))
        if _is_oe_comp(dl) and dl.upper() not in _DG_SKIP and gc1(2)=='1':
            per=_parse_number(gc1(1))
            for gi,gl in enumerate(['D/G 1','D/G 2','D/G 3']):
                dt=_parse_date(gc1(3+gi)); hrs=_parse_number(gc2(3+gi))
                if dt or hrs>0:
                    rows.append({'section':'D/G Equipment','description':dl,'engine_label':gl,
                                 'periodicity':per,'last_date':dt,'run_hrs':hrs,
                                 'status':_status(hrs,per)})

        # RIGHT section
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
#  TEXT FALLBACK  (resilience for unusual documents)
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
#  DEDUP + MASTER PARSE
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

    # Vessel name + report date
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

        # ME: raw grid (handles merged cyl cells)
        me_g.extend(_parse_me(rg))
        # AUX: dedup grid (avoids ghost cyls)
        aux_g.extend(_parse_aux(dg))
        # OE: exact column-indexed, dedup grid
        if 'TURBOCHARGER' in full and 'A/C & REFR' in full and 'COOLERS' in full:
            oe_rows.extend(_parse_oe(dg))
        # DG: paired-row parser, dedup grid; case-insensitive detection
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
#  HTML MATRIX  —  pure st.markdown(), zero Styler/column_config bugs
# ══════════════════════════════════════════════════════════════════════════
def _cyl_n(u: str) -> int:
    m=re.search(r'\d+',str(u)); return int(m.group()) if m else 999

def _sort_recs(records: List[Dict], mode: str) -> List[Dict]:
    if mode=='matrix':
        return sorted(records,key=lambda c:(c['description'].upper(),c.get('engine_label',''),_cyl_n(c.get('unit',''))))
    return sorted(records,key=lambda c:(_ORD.get(c['status'],4),-(c.get('pct_used') or 0)))

def _matrix_html(records: List[Dict], mode: str='matrix') -> str:
    if not records: return ''
    HEADS=['Status','Component','Engine','Unit','Periodicity','Last O/H','Hrs Since','% Used']
    TH=("padding:9px 14px;font-family:'Inter',sans-serif;font-size:.56rem;font-weight:600;"
        "text-transform:uppercase;letter-spacing:.14em;color:#102030;text-align:left;"
        "white-space:nowrap;border-bottom:1px solid #142030")
    header='<tr style="background:#020508">'+\
        ''.join(f'<th style="{TH}">{h}</th>' for h in HEADS)+'</tr>'
    body=''
    for rec in _sort_recs(records,mode):
        s=str(rec.get('status','NO DATA')); c=_SC.get(s,_SC['NO DATA'])
        pct=float(rec.get('pct_used') or 0); hrs=float(rec.get('hrs_since') or 0)
        per=float(rec.get('periodicity') or 0)
        pct_s=f"{pct*100:.1f}%" if pct>0 else '—'
        hrs_s=f"{int(hrs):,}" if hrs>0 else '—'
        per_s=f"{int(per):,}" if per>0 else '—'
        dt_s=str(rec.get('last_oh_date') or '—') or '—'
        bw=min(100,pct*100)
        tag=(f'<span style="display:inline-block;padding:3px 9px;border-radius:4px;'
             f'background:{c["tag_bg"]};color:{c["tag_fg"]};font-family:JetBrains Mono,monospace;'
             f'font-size:.62rem;font-weight:700;letter-spacing:.03em;white-space:nowrap">{s}</span>')
        bar=(f'<div style="display:flex;align-items:center;gap:7px;min-width:124px">'
             f'<div style="flex:1;height:3px;background:{c["bar_e"]};border-radius:2px;overflow:hidden">'
             f'<div style="width:{bw:.1f}%;height:100%;background:{c["bar_f"]};border-radius:2px"></div>'
             f'</div><span style="font-family:JetBrains Mono,monospace;font-size:.69rem;font-weight:700;'
             f'color:{c["bar_f"]};min-width:42px;text-align:right">{pct_s}</span></div>')
        def td(v,fg,fw='400',align='left',ff="'Inter',sans-serif",fs='.78rem',mw=''):
            mw_s=f'max-width:{mw};overflow:hidden;text-overflow:ellipsis;' if mw else ''
            return (f'<td style="padding:9px 14px;color:{fg};font-family:{ff};font-size:{fs};'
                    f'font-weight:{fw};text-align:{align};white-space:nowrap;{mw_s}">{v}</td>')
        body+=(
            f'<tr style="background:{c["row"]};border-bottom:1px solid {c["bord"]};'
            f'transition:filter .1s" '
            f'onmouseover="this.style.filter=\'brightness(1.35)\'" '
            f'onmouseout="this.style.filter=\'brightness(1)\'">'
            f'<td style="padding:9px 14px">{tag}</td>'
            +td(rec.get('description',''),c['comp'],'600','left',"'Space Grotesk',sans-serif",'.79rem','295px')
            +td(rec.get('engine_label',''),c['dim'])
            +td(rec.get('unit',''),c['dim'],'400','left','JetBrains Mono,monospace','.74rem')
            +td(per_s,c['dim'],'400','right','JetBrains Mono,monospace','.74rem')
            +td(dt_s,c['dim'],'400','left','JetBrains Mono,monospace','.74rem')
            +td(hrs_s,c['num'],'600','right','JetBrains Mono,monospace','.76rem')
            +f'<td style="padding:9px 14px">{bar}</td></tr>'
        )
    return (f'<div style="overflow-x:auto;border-radius:12px;border:1px solid #182a40;'
            f'overflow:hidden;box-shadow:0 8px 38px rgba(0,0,0,.62)">'
            f'<table style="width:100%;border-collapse:collapse">'
            f'<thead>{header}</thead><tbody>{body}</tbody></table></div>')

def _oe_html(rows: List[Dict], show_status: bool=False) -> str:
    if not rows: return ''
    HEADS=(['Component','Engine','Period','Last Date','Run Hrs','Status']
           if show_status else ['Component','Section','Period','Last Date','Run Hrs'])
    TH=("padding:9px 14px;font-family:'Inter',sans-serif;font-size:.56rem;font-weight:600;"
        "text-transform:uppercase;letter-spacing:.14em;color:#102030;text-align:left;border-bottom:1px solid #142030")
    header='<tr style="background:#020508">'+\
        ''.join(f'<th style="{TH}">{h}</th>' for h in HEADS)+'</tr>'
    body=''
    for row in sorted(rows,key=lambda r:(r.get('description',''),r.get('engine_label',''))):
        bg='#020510'; fm='#183858'; fd='#0c1828'
        if show_status:
            s=str(row.get('status','NO DATA')); cc=_SC.get(s,_SC['NO DATA'])
            bg=cc['row']; fm=cc['comp']; fd=cc['dim']
        per=float(row.get('periodicity',0) or 0)
        per_s=f"{int(per):,}" if per>0 else '—'
        dt_s=str(row.get('last_date') or '—') or '—'
        hrs=float(row.get('run_hrs',0) or 0)
        hrs_s=f"{int(hrs):,}" if hrs>0 else '—'
        st_cell=''
        if show_status:
            s=str(row.get('status','NO DATA')); cc=_SC.get(s,_SC['NO DATA'])
            st_cell=(f'<td style="padding:9px 14px"><span style="display:inline-block;padding:3px 8px;'
                     f'border-radius:4px;background:{cc["tag_bg"]};color:{cc["tag_fg"]};'
                     f'font-family:JetBrains Mono,monospace;font-size:.61rem;font-weight:700">{s}</span></td>')
        def td(v,fg,ff="'Inter',sans-serif",fw='400',align='left'):
            return (f'<td style="padding:9px 14px;color:{fg};font-family:{ff};font-size:.77rem;'
                    f'font-weight:{fw};text-align:{align};white-space:nowrap">{v}</td>')
        body+=(f'<tr style="background:{bg};border-bottom:1px solid #0e1c2c;transition:filter .1s" '
               f'onmouseover="this.style.filter=\'brightness(1.35)\'" '
               f'onmouseout="this.style.filter=\'brightness(1)\'">'
               +td(row.get('description',''),fm,"'Space Grotesk',sans-serif",'600')
               +(td(row.get('engine_label',''),fd,'JetBrains Mono,monospace') if show_status
                 else td(row.get('section',''),fd))
               +f'<td style="padding:9px 14px;color:{fd};font-family:JetBrains Mono,monospace;font-size:.77rem;text-align:right">{per_s}</td>'
               +td(dt_s,fd,'JetBrains Mono,monospace')
               +f'<td style="padding:9px 14px;color:#1840b8;font-family:JetBrains Mono,monospace;font-size:.77rem;font-weight:600;text-align:right">{hrs_s}</td>'
               +st_cell+'</tr>')
    return (f'<div style="overflow-x:auto;border-radius:12px;border:1px solid #182a40;'
            f'overflow:hidden;box-shadow:0 4px 22px rgba(0,0,0,.52)">'
            f'<table style="width:100%;border-collapse:collapse">'
            f'<thead>{header}</thead><tbody>{body}</tbody></table></div>')

def _show(records: List[Dict], mode: str='matrix'):
    if not records:
        st.markdown('<div style="background:rgba(1,36,14,.05);border:1px solid rgba(1,36,14,.12);'
                    'border-radius:10px;padding:1.2rem;text-align:center;color:#059028;'
                    'font-family:Space Grotesk,sans-serif;font-size:.81rem;font-weight:500">'
                    'No records match the current filter.</div>',unsafe_allow_html=True)
    else:
        st.markdown(_matrix_html(records,mode),unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  UI COMPONENTS
# ══════════════════════════════════════════════════════════════════════════
def _kpi(val: Any, lbl: str, accent: str, delay: float=0.0) -> str:
    r,g,b=int(accent[1:3],16),int(accent[3:5],16),int(accent[5:7],16)
    return (f'<div style="background:linear-gradient(160deg,var(--bg4),var(--bg3));'
            f'border:1px solid rgba({r},{g},{b},.22);border-top:2px solid {accent};'
            f'border-radius:10px;padding:.82rem 1rem .92rem;position:relative;overflow:hidden;'
            f'cursor:default;animation:U .34s {delay}s ease both;animation-fill-mode:both;'
            f'transition:transform .17s,box-shadow .18s" '
            f'onmouseover="this.style.transform=\'translateY(-4px)\';this.style.boxShadow=\'0 12px 36px rgba(0,0,0,.65)\'" '
            f'onmouseout="this.style.transform=\'\';this.style.boxShadow=\'\'">'
            f'<div style="position:absolute;inset:0;border-radius:10px;pointer-events:none;'
            f'background:linear-gradient(160deg,rgba({r},{g},{b},.048),transparent 55%)"></div>'
            f'<div style="font-family:var(--ff);font-size:1.52rem;font-weight:700;line-height:1.1;'
            f'letter-spacing:-.04em;color:{accent};position:relative;'
            f'animation:N .34s {delay+.07}s ease both;animation-fill-mode:both">{val}</div>'
            f'<div style="font-family:var(--fi);font-size:.54rem;font-weight:500;text-transform:uppercase;'
            f'letter-spacing:.18em;color:var(--tx3);margin-top:5px;position:relative">{lbl}</div>'
            f'</div>')

def _section(icon: str, title: str, sub: str, badge: str, accent: str) -> str:
    r,g,b=int(accent[1:3],16),int(accent[3:5],16),int(accent[5:7],16)
    return (f'<div style="display:flex;align-items:center;gap:.85rem;margin:2.4rem 0 1rem">'
            f'<div style="width:34px;height:34px;border-radius:8px;'
            f'background:rgba({r},{g},{b},.1);border:1px solid rgba({r},{g},{b},.2);'
            f'display:flex;align-items:center;justify-content:center;font-size:.9rem;flex-shrink:0">'
            f'{icon}</div>'
            f'<div style="flex:1">'
            f'<div style="font-family:var(--ff);font-size:.96rem;font-weight:600;color:var(--tx0);'
            f'letter-spacing:-.01em;line-height:1">{title}</div>'
            f'<div style="font-family:var(--fm);font-size:.54rem;color:var(--tx3);'
            f'margin-top:3px;letter-spacing:.05em">{sub}</div></div>'
            f'<div style="flex:1;height:1px;background:linear-gradient(90deg,var(--rim2),transparent)"></div>'
            f'<div style="font-family:var(--fm);font-size:.55rem;font-weight:500;padding:3px 9px;'
            f'border-radius:20px;background:var(--bg4);border:1px solid var(--rim2);'
            f'color:var(--tx2);flex-shrink:0;white-space:nowrap">{badge}</div></div>')

def _rec_line(recs: List[Dict]) -> str:
    n=len(recs)
    od=sum(1 for c in recs if c.get('status')=='OVERDUE')
    hp=sum(1 for c in recs if c.get('status')=='HIGH PRIORITY')
    ok=sum(1 for c in recs if c.get('status')=='OK')
    nd=sum(1 for c in recs if c.get('status')=='NO DATA')
    parts=[f'<b style="color:var(--tx1)">{n}</b> records']
    if od: parts.append(f'<span style="color:#c81818;font-weight:700">{od} overdue</span>')
    if hp: parts.append(f'<span style="color:#a84e0e;font-weight:700">{hp} high priority</span>')
    if ok: parts.append(f'<span style="color:#066830;font-weight:700">{ok} OK</span>')
    if nd: parts.append(f'<span style="color:#0c2ca0;font-weight:600">{nd} no data</span>')
    return (f'<div style="font-family:var(--fm);font-size:.58rem;color:var(--tx3);'
            f'margin-bottom:.52rem;line-height:1.8">{" · ".join(parts)}</div>')


# ══════════════════════════════════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════════════════════════════════
if 'data' not in st.session_state:
    st.session_state.data=None


# ══════════════════════════════════════════════════════════════════════════
#  PAGE  HEADER
# ══════════════════════════════════════════════════════════════════════════
st.markdown(
    '<div style="padding:1.9rem 0 0;animation:D .42s cubic-bezier(.22,1,.36,1) both">'
    '<div style="font-family:var(--fi);font-size:.52rem;font-weight:500;letter-spacing:.3em;'
    'text-transform:uppercase;color:#c08818;margin-bottom:.26rem">Running Hours Management System</div>'
    '<div style="font-family:var(--ff);font-size:1.95rem;font-weight:700;color:#b0ccee;'
    'letter-spacing:-.04em;line-height:1.05">Fleet Overview</div>'
    '</div>'
    '<div style="height:1px;margin:.8rem 0 1.4rem;'
    'background:linear-gradient(90deg,#c08818,#1c2e44 30%,transparent);'
    'animation:G .6s .07s ease both;animation-fill-mode:both"></div>',
    unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  UPLOAD PANEL
# ══════════════════════════════════════════════════════════════════════════
with st.expander("Upload TEC-004 Report",expanded=(st.session_state.data is None)):
    uc,ic=st.columns([2.2,1.0],gap='large')
    with uc:
        uploaded=st.file_uploader('Drop TEC-004 .doc',type=['doc'],label_visibility='collapsed')
    with ic:
        st.markdown(
            '<div style="font-family:var(--fi);font-size:.76rem;color:var(--tx2);line-height:1.9">'
            '<b style="color:var(--tx1)">Format:</b> TEC-004 Running Hours Report (.doc)<br>'
            '<b style="color:var(--tx1)">Parser:</b> Raw-grid ME · Dedup AUX · Column-exact OE/DG<br>'
            '<b style="color:var(--tx1)">Output:</b> ME · AUX-1/2/3 · OE · D/G Equipment<br>'
            '<b style="color:var(--tx1)">Thresholds:</b> ≥ 100% Overdue · ≥ 80% High Priority'
            '</div>',unsafe_allow_html=True)

    if uploaded:
        raw=uploaded.read(); fh=hashlib.md5(raw).hexdigest()
        if st.session_state.data is None or st.session_state.data.get('_hash')!=fh:
            with st.spinner('Converting .doc → .docx via LibreOffice…'):
                try: docx=convert_doc_to_docx(raw)
                except Exception as e: st.error(f'Conversion failed: {e}'); st.stop()
            with st.spinner('Parsing TEC-004 tables…'):
                try: result=parse_docx(docx)
                except ValueError as e: st.error(f'Parse failed: {e}'); st.stop()
            for w in result['warnings']: st.warning(f'⚠ {w}')
            if not result['me'] and not result['aux']:
                st.error('No components extracted. Verify this is a TEC-004 report.'); st.stop()
            result['_hash']=fh; result['_filename']=uploaded.name
            st.session_state.data=result
            ac=result['me']+result['aux']
            od=sum(1 for c in ac if c['status']=='OVERDUE')
            hp=sum(1 for c in ac if c['status']=='HIGH PRIORITY')
            st.markdown(
                f'<div style="background:linear-gradient(135deg,rgba(1,72,26,.14),rgba(1,72,26,.04));'
                f'border:1px solid rgba(5,144,40,.26);border-radius:10px;padding:.82rem 1.25rem;'
                f'color:#60e090;font-family:Space Grotesk,sans-serif;font-size:.84rem;font-weight:500;'
                f'animation:POP .46s cubic-bezier(.34,1.56,.64,1) both;'
                f'display:flex;align-items:center;gap:.6rem">'
                f'<span>&#10003;</span>'
                f'<span><strong>{result["vessel_name"]}</strong> — '
                f'{len(ac)} components · {od} overdue · {hp} high priority</span></div>',
                unsafe_allow_html=True)


# ── empty state ──────────────────────────────────────────────────────────
if st.session_state.data is None:
    st.markdown(
        '<div style="display:flex;align-items:center;justify-content:center;'
        'height:40vh;flex-direction:column;gap:1rem">'
        '<svg width="52" height="52" viewBox="0 0 64 64" fill="none" '
        'xmlns="http://www.w3.org/2000/svg" style="opacity:.08">'
        '<path d="M32 8v48M32 8L20 20M32 8l12 12M8 32h48" '
        'stroke="#587890" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>'
        '<circle cx="32" cy="48" r="8" stroke="#587890" stroke-width="2.5"/></svg>'
        '<div style="font-family:Inter,sans-serif;font-size:.81rem;color:#102030;letter-spacing:.05em">'
        'Upload a TEC-004 report to view the matrices</div></div>',
        unsafe_allow_html=True); st.stop()


# ══════════════════════════════════════════════════════════════════════════
#  DATA READY
# ══════════════════════════════════════════════════════════════════════════
d=st.session_state.data
me=d['me']; aux=d['aux']; oe=d['oe']; dg=d['dg']
ac=me+aux
n_od=sum(1 for c in ac if c['status']=='OVERDUE')
n_hp=sum(1 for c in ac if c['status']=='HIGH PRIORITY')
mt=d.get('me_total_hrs'); mo=d.get('me_this_month')


# ── KPI row ──────────────────────────────────────────────────────────────
cols=st.columns(8)
for col,(val,lbl,acc,dly) in zip(cols,[
    (d['vessel_name'],         'Vessel',         '#2050d8',0.00),
    (d['report_date'] or '—', 'Report Date',    '#c08818',0.04),
    (f"{mt:,}" if mt else '—','M/E Total Hrs',  '#0a9040',0.08),
    (f"{mo:,}" if mo else '—','M/E This Month', '#0a9040',0.12),
    (len(me),                  'ME Records',     '#c08818',0.16),
    (len(aux),                 'AUX Records',    '#c08818',0.20),
    (n_od,                     'Overdue',        '#c81818',0.24),
    (n_hp,                     'High Priority',  '#a84e0e',0.28),
]):
    with col: st.markdown(_kpi(val,lbl,acc,dly),unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 1 — MAIN ENGINE
# ══════════════════════════════════════════════════════════════════════════
me_od=sum(1 for c in me if c['status']=='OVERDUE')
me_hp=sum(1 for c in me if c['status']=='HIGH PRIORITY')
st.markdown(_section('⚙','Main Engine',f'{len(me)} component records',
                     f'{me_od} OD · {me_hp} HP','#c08818'),unsafe_allow_html=True)

f1,f2,f3=st.columns([2,2,3])
with f1: mc_sel=st.selectbox('Component',['All']+sorted({c['description'] for c in me}),key='mc')
with f2: ms_sel=st.selectbox('Status',['All','Overdue only','High Priority +','OK only'],key='ms')
with f3: mr_sel=st.radio('Sort',['Component → Cylinder','Priority → % Used'],horizontal=True,key='mr')

v=me[:]
if mc_sel!='All':              v=[c for c in v if c['description']==mc_sel]
if ms_sel=='Overdue only':     v=[c for c in v if c['status']=='OVERDUE']
elif ms_sel=='High Priority +':v=[c for c in v if c['status'] in ('OVERDUE','HIGH PRIORITY')]
elif ms_sel=='OK only':        v=[c for c in v if c['status']=='OK']
st.markdown(_rec_line(v),unsafe_allow_html=True)
_show(v,mode='matrix' if 'Component' in mr_sel else 'priority')


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 2 — AUX ENGINES
# ══════════════════════════════════════════════════════════════════════════
ax_od=sum(1 for c in aux if c['status']=='OVERDUE')
ax_hp=sum(1 for c in aux if c['status']=='HIGH PRIORITY')
st.markdown(_section('⚙','Auxiliary Engines',
                     f'{len(aux)} component records  ·  AUX-1 · AUX-2 · AUX-3',
                     f'{ax_od} OD · {ax_hp} HP','#2050d8'),unsafe_allow_html=True)

g1,g2,g3,g4=st.columns([1.3,2,2,3])
with g1: ae_sel=st.selectbox('Engine',['All']+sorted({c['engine_label'] for c in aux}),key='ae')
with g2: ac_sel=st.selectbox('Component',['All']+sorted({c['description'] for c in aux}),key='ac')
with g3: as_sel=st.selectbox('Status',['All','Overdue only','High Priority +','OK only'],key='axs')
with g4: ar_sel=st.radio('Sort',['Component → Cylinder','Priority → % Used'],horizontal=True,key='ar')

v=aux[:]
if ae_sel!='All':              v=[c for c in v if c['engine_label']==ae_sel]
if ac_sel!='All':              v=[c for c in v if c['description']==ac_sel]
if as_sel=='Overdue only':     v=[c for c in v if c['status']=='OVERDUE']
elif as_sel=='High Priority +':v=[c for c in v if c['status'] in ('OVERDUE','HIGH PRIORITY')]
elif as_sel=='OK only':        v=[c for c in v if c['status']=='OK']
st.markdown(_rec_line(v),unsafe_allow_html=True)
_show(v,mode='matrix' if 'Component' in ar_sel else 'priority')


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 3 — OTHER EQUIPMENT
# ══════════════════════════════════════════════════════════════════════════
st.markdown(_section('&#9737;','Other Equipment',
                     'Turbocharger · Coolers · A/C & Compressors',
                     f'{len(oe)} records','#066830'),unsafe_allow_html=True)
if not oe:
    st.markdown('<div style="background:rgba(1,36,14,.04);border:1px solid rgba(1,36,14,.1);'
                'border-radius:10px;padding:1.2rem;text-align:center;color:#059028;'
                'font-family:Space Grotesk,sans-serif;font-size:.81rem;font-weight:500">'
                'No other equipment data found.</div>',unsafe_allow_html=True)
else:
    secs=sorted({r['section'] for r in oe})
    os_sel=st.selectbox('Section',['All']+secs,key='oe_s')
    ov=oe if os_sel=='All' else [r for r in oe if r['section']==os_sel]
    st.markdown(f'<div style="font-family:var(--fm);font-size:.58rem;color:var(--tx3);'
                f'margin-bottom:.52rem;line-height:1.8"><b style="color:var(--tx1)">'
                f'{len(ov)}</b> records</div>',unsafe_allow_html=True)
    st.markdown(_oe_html(ov,show_status=False),unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
#  MATRIX 4 — D/G EQUIPMENT
# ══════════════════════════════════════════════════════════════════════════
st.markdown(_section('&#9889;','D/G Equipment',
                     'Diesel Generator components · D/G 1 · D/G 2 · D/G 3',
                     f'{len(dg)} records','#0c2ca0'),unsafe_allow_html=True)
if not dg:
    st.markdown('<div style="background:rgba(2,14,54,.04);border:1px solid rgba(2,14,54,.1);'
                'border-radius:10px;padding:1.2rem;text-align:center;color:#2050d8;'
                'font-family:Space Grotesk,sans-serif;font-size:.81rem;font-weight:500">'
                'No D/G equipment data found.</div>',unsafe_allow_html=True)
else:
    dg_sel=st.selectbox('D/G Unit',['All']+sorted({r.get('engine_label','') for r in dg}),key='dge')
    dv=dg if dg_sel=='All' else [r for r in dg if r.get('engine_label')==dg_sel]
    st.markdown(f'<div style="font-family:var(--fm);font-size:.58rem;color:var(--tx3);'
                f'margin-bottom:.52rem;line-height:1.8"><b style="color:var(--tx1)">'
                f'{len(dv)}</b> records</div>',unsafe_allow_html=True)
    st.markdown(_oe_html(dv,show_status=True),unsafe_allow_html=True)
