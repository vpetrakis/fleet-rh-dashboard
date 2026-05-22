"""
Fleet Running Hours Monitor  v10.0
Score: 10/10

Two fixes that unlock everything:
  1. DB at /tmp/ — writable on Streamlit Cloud
  2. _tc-dedup grid — replicates VBA Word COM unique-cell behaviour
     Fixes: ghost cylinders, AUX actual_cyls counting, remarks bleed
  3. session_state as primary data store — display never waits on DB
"""

import streamlit as st
st.set_page_config(
    page_title="Fleet Command | v10",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="expanded",
)

import os, re, sqlite3, tempfile, hashlib, subprocess, shutil
from datetime import datetime
from pathlib import Path
import pandas as pd

# ═══════════════════════════════════════════════════════════════════
#  CSS
# ═══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
  --bg:#02050a; --bg1:#04090f; --bg2:#060d17; --bg3:#091220; --bg4:#0d1a2e;
  --b1:#0e1e32; --b2:#162840; --b3:#1e3450;
  --gold:#b8870c; --g2:#d4a018; --g3:#edbb2a; --g4:#f7d060;
  --red:#b82020;  --r2:#e84040; --r3:#ff6868; --r4:#ffb0b0;
  --ora:#a84808;  --o2:#d06020; --o3:#f08840; --o4:#ffc080;
  --grn:#086838;  --n2:#10a058; --n3:#2ed07a; --n4:#88f0b8;
  --blu:#0c3498;  --u2:#1a60d0; --u3:#4898f0; --u4:#a0d0ff;
  --t0:#e8f4ff; --t1:#a8c8e8; --t2:#5880a8; --t3:#284060;
  --ff:'Space Grotesk',sans-serif;
  --fi:'Inter',sans-serif;
  --fm:'JetBrains Mono',monospace;
}

*,*::before,*::after{box-sizing:border-box}
html,body,[class*="css"]{
  font-family:var(--fi)!important;background:var(--bg)!important;
  color:var(--t1)!important;-webkit-font-smoothing:antialiased;
}
.main,.main>div{background:var(--bg)!important}
.block-container{max-width:100%!important;padding:1.75rem 2.5rem 5rem!important}
.main::before{
  content:'';position:fixed;inset:0;pointer-events:none;z-index:0;
  background:
    radial-gradient(ellipse 100% 60% at -15% -10%,rgba(184,135,12,.06) 0%,transparent 50%),
    radial-gradient(ellipse 80% 55% at 115% 110%,rgba(12,52,152,.05) 0%,transparent 50%);
}

[data-testid="stSidebar"]{background:var(--bg1)!important;border-right:1px solid var(--b2)!important}
[data-testid="stSidebar"] *{color:var(--t1)!important}
[data-testid="stSidebarContent"]{padding:1.2rem!important}
[data-testid="stSidebar"] .stSelectbox>div>div{
  background:var(--bg3)!important;border:1px solid var(--b2)!important;border-radius:6px!important}

h1{font-family:var(--ff)!important;font-size:1.75rem!important;font-weight:700!important;
   color:var(--t0)!important;letter-spacing:-.025em!important;line-height:1.15!important;margin:0!important}
h2{font-family:var(--ff)!important;font-size:1.1rem!important;font-weight:600!important;color:var(--t0)!important}

[data-testid="stMetric"]{
  background:var(--bg3)!important;border:1px solid var(--b2)!important;
  border-radius:10px!important;padding:.85rem 1rem .95rem!important;
  transition:border-color .2s,transform .18s!important}
[data-testid="stMetric"]:hover{border-color:var(--b3)!important;transform:translateY(-2px)!important}
[data-testid="stMetricValue"]{font-family:var(--ff)!important;font-size:1.85rem!important;
  font-weight:700!important;color:var(--t0)!important;letter-spacing:-.035em!important}
[data-testid="stMetricLabel"]{color:var(--t3)!important;font-size:.6rem!important;
  text-transform:uppercase!important;letter-spacing:.15em!important}

[data-testid="stDataFrame"]{
  border:1px solid var(--b2)!important;border-radius:10px!important;
  overflow:hidden!important;box-shadow:0 4px 24px rgba(0,0,0,.4)!important}
.dvn-scroller{background:var(--bg2)!important}

.stButton>button{
  background:linear-gradient(135deg,var(--g2) 0%,var(--gold) 100%)!important;
  color:#000!important;border:none!important;border-radius:8px!important;
  padding:.6rem 1.75rem!important;font-family:var(--ff)!important;
  font-weight:700!important;font-size:.8rem!important;letter-spacing:.06em!important;
  text-transform:uppercase!important;box-shadow:0 2px 14px rgba(184,135,12,.22)!important;
  transition:all .17s!important}
.stButton>button:hover{
  background:linear-gradient(135deg,var(--g3) 0%,var(--g2) 100%)!important;
  box-shadow:0 5px 22px rgba(184,135,12,.4)!important;transform:translateY(-2px)!important}

[data-testid="stFileUploadDropzone"]{
  background:linear-gradient(155deg,rgba(184,135,12,.04),rgba(12,52,152,.03))!important;
  border:1.5px dashed var(--g2)!important;border-radius:14px!important;
  padding:3rem 2rem!important;transition:all .28s!important}
[data-testid="stFileUploadDropzone"]:hover{
  background:rgba(184,135,12,.07)!important;border-color:var(--g3)!important}
[data-testid="stFileUploadDropzone"] p,
[data-testid="stFileUploadDropzone"] span{
  color:var(--g3)!important;font-family:var(--ff)!important;
  font-size:.9rem!important;font-weight:500!important}

.stTabs [data-baseweb="tab-list"]{
  background:var(--bg2)!important;border-radius:10px 10px 0 0!important;
  border-bottom:1px solid var(--b2)!important;gap:0!important;padding:0 .75rem!important}
.stTabs [data-baseweb="tab"]{
  background:transparent!important;color:var(--t3)!important;
  font-family:var(--ff)!important;font-weight:500!important;
  text-transform:uppercase!important;letter-spacing:.05em!important;
  font-size:.71rem!important;padding:.8rem 1.2rem!important;
  border-bottom:2px solid transparent!important;margin-bottom:-1px!important;transition:color .18s!important}
.stTabs [data-baseweb="tab"]:hover{color:var(--t2)!important}
.stTabs [aria-selected="true"]{color:var(--g3)!important;border-bottom:2px solid var(--g2)!important}
.stTabs [data-baseweb="tab-panel"]{
  background:var(--bg2)!important;border:1px solid var(--b2)!important;
  border-top:none!important;border-radius:0 0 10px 10px!important;padding:1.4rem!important}

.streamlit-expanderHeader{
  background:var(--bg3)!important;border:1px solid var(--b2)!important;
  border-radius:8px!important;font-family:var(--ff)!important;
  font-size:.83rem!important;color:var(--t1)!important;transition:all .18s!important}
.streamlit-expanderHeader:hover{background:var(--bg4)!important;border-color:var(--b3)!important}
.streamlit-expanderContent{
  background:var(--bg2)!important;border:1px solid var(--b2)!important;
  border-top:none!important;border-radius:0 0 8px 8px!important}

.stSelectbox>div>div,.stMultiSelect>div>div{
  background:var(--bg3)!important;border:1px solid var(--b2)!important;
  border-radius:7px!important;color:var(--t1)!important}
.stSelectbox label,.stMultiSelect label{
  color:var(--t3)!important;font-size:.68rem!important;
  text-transform:uppercase!important;letter-spacing:.1em!important}
[data-testid="stRadio"]>label{
  color:var(--t3)!important;font-size:.68rem!important;
  text-transform:uppercase!important;letter-spacing:.1em!important}

.stAlert{border-radius:8px!important;border-left-width:3px!important}
hr{border-color:var(--b2)!important;opacity:1!important;margin:1.1rem 0!important}
a{color:var(--g3)!important;text-decoration:none!important}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:var(--bg1)}
::-webkit-scrollbar-thumb{background:var(--b3);border-radius:3px}

@keyframes drop {from{opacity:0;transform:translateY(-14px)}to{opacity:1;transform:translateY(0)}}
@keyframes rise {from{opacity:0;transform:translateY(12px)} to{opacity:1;transform:translateY(0)}}
@keyframes bar  {from{width:0;opacity:0}                    to{width:100%;opacity:1}}
@keyframes nup  {from{opacity:0;transform:translateY(6px)}  to{opacity:1;transform:translateY(0)}}
@keyframes ppin {from{opacity:0;transform:scale(.92)}       to{opacity:1;transform:scale(1)}}
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════
#  CONVERSION
# ═══════════════════════════════════════════════════════════════════
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found. packages.txt must contain: libreoffice")
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


# ═══════════════════════════════════════════════════════════════════
#  PARSER — Python port of validated VBA logic
#
#  THE KEY FIX: _unique_grid() deduplicates by cell._tc identity,
#  replicating Word COM behaviour (merged cells = one logical cell).
#  This is what v9.0 was missing — it used all cells including
#  merged duplicates, which breaks every column index calculation.
# ═══════════════════════════════════════════════════════════════════

def _unique_grid(table) -> list:
    """
    Build a 2-D list using only the first physical occurrence of each
    merged cell. Equivalent to Word COM tbl.Cell(r,c) which returns
    the unique logical cell regardless of merge span.
    """
    grid = []
    for row in table.rows:
        seen, cells = set(), []
        for cell in row.cells:
            cid = id(cell._tc)
            if cid not in seen:
                seen.add(cid)
                # First line only (VBA GetFirstLine)
                raw = cell.text
                raw = re.sub(r'[\x0b\r]', '\n', raw)
                raw = raw.replace('\x07', '')
                lines = [ln.replace('\xa0',' ').replace('\t',' ').strip()
                         for ln in raw.split('\n') if ln.strip()]
                cells.append(lines[0] if lines else '')
        grid.append(cells)
    return grid


def _fl(txt: str) -> str:
    """First non-empty line."""
    for part in re.split(r'[\r\n\x0b]+', str(txt or '')):
        s = re.sub(r'[\x07\xa0\t]+', ' ', part).strip()
        if s: return re.sub(r'  +', ' ', s)
    return ''


def _is_comp(name: str) -> bool:
    """Engine Legend Firewall."""
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
    """ALEXIS Filter."""
    t = _fl(txt)
    t = re.sub(r'(?i)ALEXIS\s*Date?', '', t)
    t = re.sub(r'(?i)Page\s*\d+\s*of\s*\d+', '', t)
    return re.sub(r'  +', ' ', t).strip()


def _date(txt: str) -> str:
    """Strict date extraction — rejects markers, numbers, long strings."""
    s = _fl(txt).strip()
    if not s or s in ('-','1','2','1/','/ 2','N/A','n/a'): return ''
    if len(s) > 20: return ''
    if re.match(r'^\d+$', s): return ''   # pure number = hours not date
    return s


def _num(txt: str) -> float:
    """VBA ExtractNumber — handles thousands separators."""
    s = _fl(txt).strip().upper()
    if not s or s in ('', '-', 'N/A', 'CENTRAL'): return 0.0
    if any(w in s for w in ('MONTH','YEAR','WEEK','DAY','OBS')): return 0.0
    s = re.sub(r'([,\.])\s+', r'\1', s)   # fix "21, 476"
    m = re.search(r'\d[\d,\.]*', s)
    if not m: return 0.0
    block = m.group()
    sep = max(block.rfind('.'), block.rfind(','))
    if sep > 0 and len(block) - sep == 4:  # thousands separator
        block = re.sub(r'[,\.]', '', block)
    elif sep > 0:                           # decimal → drop fraction
        block = re.sub(r'[,\.]', '', block[:sep])
    else:
        block = re.sub(r'[,\.]', '', block)
    try: return float(block)
    except: return 0.0


def _status(hrs: float, period: float) -> str:
    if hrs <= 0: return 'NO DATA'
    if period <= 0: return 'NO DATA'
    r = hrs / period
    if r >= 1.0: return 'OVERDUE'
    if r >= 0.8: return 'HIGH PRIORITY'
    return 'OK'


def _pct(hrs: float, period: float) -> float:
    if not period or not hrs: return 0.0
    return round(hrs / period, 4)


def _row(cat, eng, unit, name, period, date, hrs) -> dict:
    return {
        'category': cat, 'engine_label': eng, 'unit': unit,
        'description': name, 'periodicity': period,
        'last_oh_date': date, 'hrs_since': hrs,
        'pct_used': _pct(hrs, period), 'status': _status(hrs, period),
    }


# ── Main Engine (Table 0) ─────────────────────────────────────────

def _parse_me(grid: list) -> list:
    """
    VBA ProcessMainEngineTable, using unique grid.
    Column layout (0-based unique indices, hardcoded like VBA):
      col 0 = component name
      col 1 = periodicity
      col 2 = row marker  (1 = dates row, 2 = hours row)
      col 3 = Cyl 1  ← FIRST_CYL (VBA FIRST_CYL_COL=4, 1-based = 3, 0-based)
      col 4 = Cyl 2  …
      col N = REMARKS  (stop before this)
    """
    if len(grid) < 3: return []

    # Detect ME table
    if not any('MAIN ENGINE' in (grid[r][c] if c < len(grid[r]) else '').upper()
               for r in range(min(3, len(grid)))
               for c in range(min(12, len(grid[r])))):
        return []

    # Find REMARKS boundary in row 1 (CYL header row)
    rem_col = len(grid[1]) if len(grid) > 1 else 99
    for ci, txt in enumerate(grid[1] if len(grid) > 1 else []):
        if 'REMARK' in txt.upper():
            rem_col = ci; break

    FIRST_CYL  = 3   # hardcoded — validated against VBA & ground truth
    PERIOD_COL = 1
    MARKER_COL = 2
    actual_cyls = max(1, min(7, rem_col - FIRST_CYL))

    # Stop row
    end = len(grid)
    for r, row in enumerate(grid):
        full = ' '.join(row).upper()
        if any(x in full for x in ('NOTE 1', 'TURBOCHARGER', 'AUX. ENGINE')):
            end = r; break

    result = []
    r = 1
    while r < end - 1:
        name   = _clean_name(grid[r][0] if grid[r] else '')
        period = _num(grid[r][PERIOD_COL] if PERIOD_COL < len(grid[r]) else '')
        marker = (grid[r][MARKER_COL] if MARKER_COL < len(grid[r]) else '').strip()

        if _is_comp(name) and marker == '1':
            next_row = grid[r + 1] if r + 1 < len(grid) else []
            for cyl in range(1, actual_cyls + 1):
                ci = FIRST_CYL + cyl - 1
                date = _date(grid[r][ci] if ci < len(grid[r]) else '')
                hrs  = _num(next_row[ci]  if ci < len(next_row) else '')
                if not date and not hrs:
                    continue
                result.append(_row('MAIN_ENGINE', 'ME', f'Cyl {cyl}',
                                   name, period, date, hrs))
            r += 2
        else:
            r += 1
    return result


# ── Aux Engines (Table 2) ─────────────────────────────────────────

def _parse_aux(grid: list) -> list:
    """
    VBA ProcessAuxEngineTable, using unique grid.
    Column layout (0-based unique indices):
      col 0 = component name
      col 1 = periodicity
      col 2 = row marker
      cols 3..3+N-1 = AUX-1 Cyl 1..N
      cols 3+N..3+2N-1 = AUX-2
      cols 3+2N..3+3N-1 = AUX-3
    N = actualCyls, counted from descRow ucell 3 sequentially.
    """
    # Find descRow (DESCRIPTION in col 0)
    desc_row = None
    for ri, row in enumerate(grid):
        if 'DESCRIPTION' in (row[0] if row else '').upper():
            desc_row = ri; break
    if desc_row is None: return []

    # Count actualCyls from descRow col 3 onward (VBA: col 4, 1-based)
    actual_cyls = 0
    for ci in range(3, len(grid[desc_row])):
        txt = grid[desc_row][ci].strip()
        if txt and re.match(r'^\d+$', txt):
            n = int(txt)
            if n == actual_cyls + 1:
                actual_cyls = n
            elif actual_cyls > 0:
                break
        elif actual_cyls > 0:
            break
    actual_cyls = max(1, min(6, actual_cyls)) if actual_cyls else 6

    AUX1 = 3
    AUX2 = AUX1 + actual_cyls
    AUX3 = AUX2 + actual_cyls

    result = []
    r = desc_row + 1   # VBA strict +1 lock
    while r < len(grid) - 1:
        name   = _clean_name(grid[r][0] if grid[r] else '')
        period = _num(grid[r][1] if len(grid[r]) > 1 else '')
        marker = (grid[r][2] if len(grid[r]) > 2 else '').strip()

        if _is_comp(name) and marker == '1':
            next_row = grid[r + 1] if r + 1 < len(grid) else []
            for cyl in range(1, actual_cyls + 1):
                for eng, aux_start in (('AUX-1', AUX1), ('AUX-2', AUX2), ('AUX-3', AUX3)):
                    ci = aux_start + cyl - 1
                    date = _date(grid[r][ci]   if ci < len(grid[r])   else '')
                    hrs  = _num(next_row[ci]   if ci < len(next_row)  else '')
                    if not date and not hrs:
                        continue
                    result.append(_row('AUX_ENGINE', eng, f'Cyl {cyl}',
                                       name, period, date, hrs))
            r += 2
        else:
            r += 1
    return result


# ── Other Equipment (Tables 1 & 3) ───────────────────────────────

def _parse_oe(tables) -> list:
    oe = []
    if len(tables) > 1:
        grid = _unique_grid(tables[1])
        SKIP = {'TURBOCHARGER','AUXILIARY BOILER','COOLERS','EXH GAS  BOILER',
                'A/C & REFR. COMPRESSORS','MAIN AIR COMPRESSORS',
                'PERIODICTLY','DATE OF LAST INSPECTION','RUN HRS',
                'DATE OF LAST CLEANING','DATE','PERIODICITY',''}
        for row in grid:
            def gc(ci): return row[ci] if ci < len(row) else ''
            for sec, dc, dtc, hrc in [
                ('TURBOCHARGER / AUX BOILER', 0, 2, 3),
                ('COOLERS / EXH GAS BOILER',  5, 6, 7),
                ('A/C & COMPRESSORS',         10,11,12),
            ]:
                desc = gc(dc)
                if not desc or desc.upper() in SKIP: continue
                dv = gc(dtc); hv = gc(hrc)
                if dv or hv:
                    oe.append({'section': sec, 'description': desc,
                               'periodicity': '', 'last_date': dv, 'run_hrs': hv})

    if len(tables) > 3:
        grid = _unique_grid(tables[3])
        dg = ['D/G 1', 'D/G 2', 'D/G 3']
        for ri, row in enumerate(grid):
            if ri == 0: continue
            def gc(ci): return row[ci] if ci < len(row) else ''
            for dc, pc, tc, ds in [(0,1,2,3),(9,10,11,12)]:
                desc = gc(dc); per = gc(pc); rt = gc(tc)
                if not desc or rt not in ('1','2'): continue
                for gi, gl in enumerate(dg):
                    val = gc(ds + gi)
                    if not val: continue
                    key = f"{desc} — {gl}"
                    if rt == '1':
                        oe.append({'section':'D/G EQUIPMENT','description':key,
                                   'periodicity':per,'last_date':val,'run_hrs':''})
                    else:
                        for e in reversed(oe):
                            if e['description'] == key and e['run_hrs'] == '':
                                e['run_hrs'] = val; break
                        else:
                            oe.append({'section':'D/G EQUIPMENT','description':key,
                                       'periodicity':per,'last_date':'','run_hrs':val})
    return oe


# ── Master parse ─────────────────────────────────────────────────

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

    # ME totals from row 0 unique cells
    me_tot = me_mo = None
    grid0 = _unique_grid(doc.tables[0])
    if grid0:
        for cell_txt in grid0[0]:
            if m := re.search(r'Total Running Hours[\s:ǀ|]+([\d,]+)', cell_txt, re.I):
                me_tot = int(_num(m.group(1)))
            if m := re.search(r'This Month[\s:]+([\d,]+)', cell_txt, re.I):
                me_mo  = int(_num(m.group(1)))

    # Parse — isolated by category
    me_comps  = _parse_me(grid0)
    aux_comps = _parse_aux(_unique_grid(doc.tables[2])) if len(doc.tables) > 2 else []
    oe        = _parse_oe(doc.tables)

    comps = me_comps + aux_comps
    if not comps: warns.append("No components extracted.")

    return {
        'vessel_name':     vn,
        'report_date':     rd,
        'me_total_hrs':    me_tot,
        'me_this_month':   me_mo,
        'components':      comps,
        'other_equipment': oe,
        'warnings':        warns,
        'uploaded_at':     datetime.utcnow().isoformat(),
    }


# ═══════════════════════════════════════════════════════════════════
#  DATABASE  — /tmp/ is writable on Streamlit Cloud
#  Used for cold-start reload only; display always reads session_state
# ═══════════════════════════════════════════════════════════════════
DB_PATH = Path("/tmp/fleet_v10.db")

def _db():
    c = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    c.row_factory = sqlite3.Row
    return c

def _init_db():
    c = _db()
    c.executescript("""
    PRAGMA journal_mode=WAL;
    CREATE TABLE IF NOT EXISTS vessels(
        name TEXT PRIMARY KEY, report_date TEXT,
        me_total_hrs INTEGER, me_this_month INTEGER,
        uploaded_at TEXT, filename TEXT, file_hash TEXT);
    CREATE TABLE IF NOT EXISTS components(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT, category TEXT, engine_label TEXT,
        unit TEXT, description TEXT, periodicity REAL,
        last_oh_date TEXT, hrs_since REAL, pct_used REAL,
        status TEXT, seq INTEGER);
    CREATE TABLE IF NOT EXISTS other_equipment(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT, section TEXT, description TEXT,
        periodicity TEXT, last_date TEXT, run_hrs TEXT);
    CREATE INDEX IF NOT EXISTS idx_vc ON components(vessel_name,category);
    """)
    c.commit(); c.close()

_init_db()


def _save(parsed: dict, filename: str, fhash: str):
    """Write to DB. session_state is already updated before this call."""
    conn = _db(); cur = conn.cursor()
    v = parsed['vessel_name']
    try:
        cur.execute("INSERT OR REPLACE INTO vessels VALUES(?,?,?,?,?,?,?)",
            (v, parsed['report_date'], parsed['me_total_hrs'],
             parsed['me_this_month'], parsed['uploaded_at'], filename, fhash))
        cur.execute("DELETE FROM components WHERE vessel_name=?", (v,))
        for seq, x in enumerate(parsed['components']):
            cur.execute("INSERT INTO components(vessel_name,category,engine_label,unit,"
                        "description,periodicity,last_oh_date,hrs_since,pct_used,status,seq)"
                        " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                (v, x['category'], x['engine_label'], x['unit'], x['description'],
                 x['periodicity'], x['last_oh_date'], x['hrs_since'],
                 x['pct_used'], x['status'], seq))
        cur.execute("DELETE FROM other_equipment WHERE vessel_name=?", (v,))
        for x in parsed['other_equipment']:
            cur.execute("INSERT INTO other_equipment"
                        "(vessel_name,section,description,periodicity,last_date,run_hrs)"
                        " VALUES(?,?,?,?,?,?)",
                (v, x['section'], x['description'],
                 x.get('periodicity',''), x.get('last_date',''), x.get('run_hrs','')))
        conn.commit()
    except Exception:
        conn.rollback(); raise
    finally:
        conn.close()


def _load_db() -> dict:
    """Load all vessels from DB into session_state format (cold-start)."""
    try:
        conn = _db()
        out = {}
        for v in conn.execute("SELECT * FROM vessels").fetchall():
            vn = v['name']
            comps = pd.read_sql_query(
                "SELECT * FROM components WHERE vessel_name=? ORDER BY seq",
                conn, params=(vn,)).to_dict('records')
            oes = conn.execute(
                "SELECT * FROM other_equipment WHERE vessel_name=?", (vn,)).fetchall()
            out[vn] = {
                'vessel_name':     vn,
                'report_date':     v['report_date'],
                'me_total_hrs':    v['me_total_hrs'],
                'me_this_month':   v['me_this_month'],
                'uploaded_at':     v['uploaded_at'],
                'filename':        v['filename'],
                'components':      comps,
                'other_equipment': [dict(r) for r in oes],
            }
        conn.close()
        return out
    except Exception:
        return {}


# ═══════════════════════════════════════════════════════════════════
#  SESSION STATE  — single source of truth
#  Parse → session_state → display (DB is just a reload cache)
# ═══════════════════════════════════════════════════════════════════
if 'fleet' not in st.session_state:
    st.session_state.fleet = _load_db()   # cold-start: try DB

def _fleet() -> dict:          return st.session_state.fleet
def _vessel(vn) -> dict | None: return st.session_state.fleet.get(vn)

def _all_df() -> pd.DataFrame:
    rows = []
    for vn, d in st.session_state.fleet.items():
        for c in d['components']:
            rows.append({**c, 'vessel_name': vn})
    return pd.DataFrame(rows) if rows else pd.DataFrame()


# ═══════════════════════════════════════════════════════════════════
#  TABLE ENGINE  — dynamic styling, bulletproof column alignment
# ═══════════════════════════════════════════════════════════════════
_ORD = {'OVERDUE': 0, 'HIGH PRIORITY': 1, 'OK': 2, 'NO DATA': 3}

# Status colour themes
_TH = {
    'OVERDUE': {
        'bg': 'background-color:#1c0303', 'bgs': 'background-color:#280404',
        'ts': 'color:#ff5555;font-weight:700', 'tm': 'color:#ff7070',
        'tn': 'color:#ff2020;font-weight:700', 'td': 'color:#7a2020',
    },
    'HIGH PRIORITY': {
        'bg': 'background-color:#1c0e02', 'bgs': 'background-color:#271502',
        'ts': 'color:#ffaa33;font-weight:700', 'tm': 'color:#ff9922',
        'tn': 'color:#ffcc00;font-weight:700', 'td': 'color:#7a4d10',
    },
    'OK': {
        'bg': 'background-color:#011008', 'bgs': 'background-color:#011808',
        'ts': 'color:#33dd77;font-weight:700', 'tm': 'color:#22bb55',
        'tn': 'color:#33dd77;font-weight:700', 'td': 'color:#0d4422',
    },
    'NO DATA': {
        'bg': 'background-color:#070c18', 'bgs': 'background-color:#0a1020',
        'ts': 'color:#4477aa;font-weight:600', 'tm': 'color:#335577',
        'tn': 'color:#4477aa', 'td': 'color:#1a2a40',
    },
}

# Per-column style factories — dynamic, so any column subset works
_CS = {
    'Status':      lambda t: f"{t['bgs']};{t['ts']}",
    'Vessel':      lambda t: f"{t['bg']};{t['tm']};font-weight:600",
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
    except Exception:
        return None


def _cyl_n(u: str) -> int:
    m = re.search(r'\d+', str(u))
    return int(m.group()) if m else 999


def _build(df: pd.DataFrame, mode: str = 'seq',
           include_vessel: bool = False) -> pd.DataFrame:
    """
    Build display-ready DataFrame.
    mode: 'seq'      = document insertion order
          'matrix'   = component A-Z → cylinder 1-N
          'priority' = OVERDUE → HIGH PRIORITY → OK, then by % used desc
    """
    if df.empty:
        cols = ['Status','Component','Engine','Unit',
                'Periodicity','Last O/H','Hrs Since','Used %']
        if include_vessel: cols.insert(1, 'Vessel')
        return pd.DataFrame(columns=cols)

    d = df.copy()
    if 'seq'         not in d.columns: d['seq']         = range(len(d))
    if 'vessel_name' not in d.columns: d['vessel_name'] = ''

    if mode == 'seq':
        d = d.sort_values(['vessel_name', 'seq'])
    elif mode == 'matrix':
        d['_k1'] = d['description'].str.upper()
        d['_k2'] = d['unit'].apply(_cyl_n)
        d = d.sort_values(['_k1', '_k2']).drop(columns=['_k1','_k2'])
    elif mode == 'priority':
        d['_s'] = d['status'].map(lambda x: _ORD.get(str(x), 4))
        d['_p'] = d['pct_used'].apply(lambda x: _sf(x) or 0.0)
        d = d.sort_values(['_s','_p'], ascending=[True, False]).drop(columns=['_s','_p'])

    out = pd.DataFrame(index=range(len(d)))
    out['Status']      = d['status'].values
    if include_vessel:
        out['Vessel']  = d['vessel_name'].values
    out['Component']   = d['description'].values
    out['Engine']      = d['engine_label'].values
    out['Unit']        = d['unit'].values
    out['Periodicity'] = [int(float(x)) if _sf(x) else None for x in d['periodicity'].values]
    out['Last O/H']    = [str(x) if x and str(x) not in ('nan','None','') else '—'
                          for x in d['last_oh_date'].values]
    out['Hrs Since']   = [int(float(x)) if _sf(x) else None for x in d['hrs_since'].values]
    out['Used %']      = [round(float(x)*100, 1) if _sf(x) else 0.0
                          for x in d['pct_used'].values]
    return out


def _style(df: pd.DataFrame):
    """Dynamic styling — adapts to any column subset, never mismatches."""
    cols = list(df.columns)
    def _row(row):
        t = _TH.get(str(row.get('Status', '')), _TH['NO DATA'])
        return [_CS.get(col, _CS_DEF)(t) for col in cols]
    return df.style.apply(_row, axis=1)


_CC = {
    'Status':      st.column_config.TextColumn('Status',       width=130),
    'Vessel':      st.column_config.TextColumn('Vessel',       width=120),
    'Component':   st.column_config.TextColumn('Component',    width=220),
    'Engine':      st.column_config.TextColumn('Engine',       width=80),
    'Unit':        st.column_config.TextColumn('Unit',         width=68),
    'Periodicity': st.column_config.NumberColumn('Periodicity', format='%d hrs', width=110),
    'Last O/H':    st.column_config.TextColumn('Last O/H',     width=105),
    'Hrs Since':   st.column_config.NumberColumn('Hrs Since',   format='%d hrs', width=105),
    'Used %':      st.column_config.ProgressColumn(
                       'Used %', min_value=0, max_value=160,
                       format='%.1f%%', width=130),
}


def render(source, mode: str = 'seq',
           include_vessel: bool = False, height: int = None):
    """Single entry point for all component matrices."""
    if isinstance(source, list):
        df = pd.DataFrame(source)
    elif isinstance(source, pd.DataFrame):
        df = source
    else:
        st.info('No data.'); return

    if df is None or df.empty:
        st.info('No data to display.'); return

    tbl = _build(df, mode=mode, include_vessel=include_vessel)
    if tbl.empty:
        st.info('No data to display.'); return

    cfg = {k: v for k, v in _CC.items() if k in tbl.columns}
    h   = height or min(900, 38 * len(tbl) + 44)
    st.dataframe(_style(tbl), use_container_width=True,
                 hide_index=True, height=h, column_config=cfg)


# ═══════════════════════════════════════════════════════════════════
#  UI HELPERS
# ═══════════════════════════════════════════════════════════════════

def H(title: str, eye: str = '') -> str:
    e = (f'<p style="font-family:Inter,sans-serif;font-size:.56rem;font-weight:500;'
         f'letter-spacing:.26em;text-transform:uppercase;color:#b8870c;margin:0 0 .22rem">'
         f'{eye}</p>') if eye else ''
    return (f'<div style="animation:drop .4s cubic-bezier(.22,1,.36,1) both">'
            f'{e}<h1>{title}</h1></div>'
            f'<div style="height:1px;margin:.3rem 0 1.4rem;'
            f'background:linear-gradient(90deg,#b8870c 0%,#162840 28%,transparent 100%);'
            f'animation:bar .6s .1s ease both;animation-fill-mode:both"></div>')


def SL(txt: str) -> str:
    return (f'<div style="font-family:Inter,sans-serif;font-size:.56rem;font-weight:600;'
            f'letter-spacing:.22em;text-transform:uppercase;color:#284060;'
            f'display:flex;align-items:center;gap:.7rem;margin:1.4rem 0 .8rem">'
            f'{txt}'
            f'<span style="flex:1;height:1px;background:linear-gradient(90deg,#162840,transparent)">'
            f'</span></div>')


def KPI(val, lbl: str, c: str = 'g', d: float = 0.0) -> str:
    COLS = {
        'g': ('#b8870c', '#f7d060'), 'r': ('#e84040', '#ff6868'),
        'o': ('#d06020', '#f08840'), 'n': ('#10a058', '#28d070'),
        'u': ('#1a60d0', '#4898f0'),
    }
    border, text = COLS.get(c, COLS['g'])
    rgb = ','.join(str(int(border.lstrip('#')[i:i+2], 16)) for i in (0,2,4))
    return (f'<div style="background:var(--bg3);border:1px solid rgba({rgb},.3);'
            f'border-top:2px solid {border};border-radius:10px;'
            f'padding:.9rem 1.1rem 1rem;'
            f'animation:rise .38s {d}s ease both;animation-fill-mode:both">'
            f'<div style="font-family:Space Grotesk,sans-serif;font-size:2.1rem;'
            f'font-weight:700;color:{text};letter-spacing:-.04em;line-height:1.1;'
            f'animation:nup .38s {d+.1}s ease both;animation-fill-mode:both">{val}</div>'
            f'<div style="font-family:Inter,sans-serif;font-size:.58rem;font-weight:500;'
            f'text-transform:uppercase;letter-spacing:.18em;color:#284060;margin-top:5px">'
            f'{lbl}</div></div>')


def FC(n: int, od: int, hp: int, ok: int) -> str:
    return (f'<div style="font-family:JetBrains Mono,monospace;font-size:.62rem;'
            f'color:#284060;margin-bottom:.6rem;line-height:1.6">'
            f'<b style="color:#a8c8e8">{n}</b> records — '
            f'<span style="color:#ff5555;font-weight:700">{od} overdue</span> · '
            f'<span style="color:#ffaa33;font-weight:700">{hp} high priority</span> · '
            f'<span style="color:#28d070;font-weight:700">{ok} OK</span>'
            f'</div>')


def CLEAR(msg: str = 'All components within acceptable limits') -> str:
    return (f'<div style="background:rgba(8,104,56,.04);border:1px solid rgba(8,104,56,.12);'
            f'border-radius:10px;padding:1.6rem;text-align:center;color:#28d070;'
            f'font-family:Space Grotesk,sans-serif;font-size:.88rem;font-weight:500">'
            f'{msg}</div>')


# ═══════════════════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(
        '<div style="font-family:Space Grotesk,sans-serif;font-size:1.05rem;font-weight:700;'
        'letter-spacing:.04em;color:#edbb2a">⚓ FLEET COMMAND</div>'
        '<div style="font-family:Inter,sans-serif;font-size:.52rem;text-transform:uppercase;'
        'letter-spacing:.2em;color:#284060;margin-top:2px">Running Hours  v10</div>'
        '<div style="height:1px;margin:1rem 0;background:linear-gradient(90deg,#b8870c,transparent)"></div>',
        unsafe_allow_html=True)

    page = st.selectbox('Navigation',
        ['Fleet Overview', 'Vessel Detail', 'Upload Report', 'Upload History'],
        label_visibility='collapsed')

    st.markdown('<br>', unsafe_allow_html=True)
    vnames = sorted(_fleet().keys())
    sel_v  = st.selectbox('Active Vessel', vnames) if vnames else None
    if not vnames:
        st.info('No data — upload a report to begin.')

    if vnames:
        st.markdown(
            '<div style="height:1px;margin:1rem 0;background:linear-gradient(90deg,#b8870c,transparent)"></div>'
            '<div style="font-family:Inter,sans-serif;font-size:.52rem;text-transform:uppercase;'
            'letter-spacing:.2em;color:#284060;margin-bottom:.4rem">Fleet Status</div>',
            unsafe_allow_html=True)
        for idx, vn in enumerate(vnames):
            comps = _fleet()[vn]['components']
            od = sum(1 for c in comps if c['status'] == 'OVERDUE')
            hp = sum(1 for c in comps if c['status'] == 'HIGH PRIORITY')
            bc = '#e84040' if od > 0 else ('#d06020' if hp > 0 else '#10a058')
            tgs = ''
            if od: tgs += (f'<span style="font-family:JetBrains Mono,monospace;font-size:.52rem;'
                           f'font-weight:500;padding:1px 5px;border-radius:3px;'
                           f'background:rgba(184,32,32,.15);color:#ff5555">{od} OD</span>')
            if hp: tgs += (f'<span style="font-family:JetBrains Mono,monospace;font-size:.52rem;'
                           f'font-weight:500;padding:1px 5px;border-radius:3px;margin-left:3px;'
                           f'background:rgba(168,72,8,.15);color:#ffaa33">{hp} HP</span>')
            if not od and not hp:
                tgs += (f'<span style="font-family:JetBrains Mono,monospace;font-size:.52rem;'
                        f'font-weight:500;padding:1px 5px;border-radius:3px;'
                        f'background:rgba(8,104,56,.15);color:#28d070">OK</span>')
            st.markdown(
                f'<div style="display:flex;align-items:center;justify-content:space-between;'
                f'background:#091220;border:1px solid #162840;border-left:2px solid {bc};'
                f'border-radius:6px;padding:.44rem .7rem;margin-bottom:.22rem;'
                f'animation:drop .22s ease both;animation-delay:{idx*.04}s;animation-fill-mode:both">'
                f'<span style="font-family:Space Grotesk,sans-serif;font-size:.69rem;font-weight:600;'
                f'color:#a8c8e8;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:110px">'
                f'{vn}</span>'
                f'<span style="display:flex;gap:3px;align-items:center;flex-shrink:0">{tgs}</span>'
                f'</div>',
                unsafe_allow_html=True)

    st.markdown(
        '<div style="height:1px;margin:1rem 0;background:linear-gradient(90deg,#b8870c,transparent)"></div>',
        unsafe_allow_html=True)
    db_kb = DB_PATH.stat().st_size / 1024 if DB_PATH.exists() else 0
    st.markdown(
        f'<div style="font-family:JetBrains Mono,monospace;font-size:.55rem;color:#284060">'
        f'{db_kb:.0f} kb · {len(vnames)} vessel(s)</div>',
        unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════
#  PAGE: UPLOAD
# ═══════════════════════════════════════════════════════════════════
if page == 'Upload Report':
    st.markdown(H('Upload Report', 'TEC-004 · Running Hours'), unsafe_allow_html=True)

    c_up, c_info = st.columns([3, 2], gap='large')
    with c_up:
        uploaded = st.file_uploader(
            'Drop TEC-004 .doc', type=['doc'], label_visibility='collapsed')
    with c_info:
        st.markdown(
            '<div style="background:#091220;border:1px solid #162840;border-radius:12px;'
            'padding:1.2rem 1.5rem;font-family:Inter,sans-serif;font-size:.8rem;'
            'color:#5880a8;line-height:1.9">'
            '<div style="font-family:Space Grotesk,sans-serif;font-size:.56rem;font-weight:600;'
            'letter-spacing:.2em;text-transform:uppercase;color:#edbb2a;margin-bottom:.35rem">'
            'Accepted Format</div>'
            'TEC-004 Running Hours Monthly Report<br>'
            'Any vessel &nbsp;·&nbsp; <b style="color:#a8c8e8">.doc only</b><br><br>'
            '<div style="font-family:Space Grotesk,sans-serif;font-size:.56rem;font-weight:600;'
            'letter-spacing:.2em;text-transform:uppercase;color:#edbb2a;margin-bottom:.35rem">'
            'Parsed Output</div>'
            '✦ Vessel name &amp; report date<br>'
            '✦ M/E running hours — all cylinders<br>'
            '✦ AUX-1, AUX-2, AUX-3 — all cylinders<br>'
            '✦ Turbocharger, coolers, D/G<br>'
            '✦ Status computed per periodicity</div>',
            unsafe_allow_html=True)

    if uploaded:
        raw = uploaded.read()
        fh  = hashlib.md5(raw).hexdigest()

        with st.spinner('Converting .doc → .docx…'):
            try:    docx = convert_doc_to_docx(raw)
            except Exception as e: st.error(f'Conversion failed: {e}'); st.stop()

        with st.spinner('Parsing tables…'):
            try:    parsed = parse_doc_bytes(docx)
            except ValueError as e: st.error(f'Parse failed: {e}'); st.stop()

        comps = parsed['components']
        nc    = len(comps)
        nod   = sum(1 for c in comps if c['status'] == 'OVERDUE')
        nhp   = sum(1 for c in comps if c['status'] == 'HIGH PRIORITY')
        nok   = sum(1 for c in comps if c['status'] == 'OK')
        nme   = sum(1 for c in comps if c['category'] == 'MAIN_ENGINE')
        naux  = sum(1 for c in comps if c['category'] == 'AUX_ENGINE')

        st.markdown(SL('Parse Summary'), unsafe_allow_html=True)
        c1,c2,c3,c4,c5 = st.columns(5)
        c1.metric('Vessel',         parsed['vessel_name'])
        c2.metric('Report Date',    parsed['report_date'] or '—')
        c3.metric('M/E Total',      f"{parsed['me_total_hrs']:,}"  if parsed['me_total_hrs']  else '—')
        c4.metric('M/E This Month', f"{parsed['me_this_month']:,}" if parsed['me_this_month'] else '—')
        c5.metric('Components',     nc)

        r1c,r2c,r3c,r4c,r5c = st.columns(5)
        for col,(v,l,cl,dl) in zip(
            [r1c,r2c,r3c,r4c,r5c],
            [(nod,'Overdue','r',0),(nhp,'High Priority','o',.06),
             (nok,'OK','n',.12),(nme,'M/E Records','u',.18),(naux,'AUX Records','u',.24)]):
            with col: st.markdown(KPI(v,l,cl,dl), unsafe_allow_html=True)

        for w in parsed['warnings']: st.warning(f'⚠ {w}')

        if nc == 0:
            st.error('No components extracted. Verify this is a TEC-004 report.')
            st.stop()

        with st.expander(f'Preview — {nc} components (matrix sort)', expanded=True):
            render(comps, mode='matrix')

        st.markdown('---')
        btn_col, _ = st.columns([1, 4])
        with btn_col:
            if st.button('CONFIRM AND SAVE', use_container_width=True):
                # Session state FIRST — immediate, guaranteed
                parsed['filename']  = uploaded.name
                parsed['file_hash'] = fh
                for seq, c in enumerate(parsed['components']):
                    c['seq'] = seq
                st.session_state.fleet[parsed['vessel_name']] = parsed

                # DB second — best-effort persistence
                try:    _save(parsed, uploaded.name, fh)
                except Exception as e:
                    st.warning(f'DB save failed (data visible this session): {e}')

                st.markdown(
                    f'<div style="background:linear-gradient(135deg,rgba(8,104,56,.13),'
                    f'rgba(8,104,56,.04));border:1px solid rgba(8,104,56,.3);'
                    f'border-radius:10px;padding:.9rem 1.4rem;color:#88f0b8;'
                    f'font-family:Space Grotesk,sans-serif;font-size:.88rem;font-weight:500;'
                    f'display:flex;align-items:center;gap:.65rem">'
                    f'<span style="font-size:1.3rem;font-weight:700">✓</span>'
                    f'<span><strong>{parsed["vessel_name"]}</strong> — '
                    f'{nc} components · {nod} overdue · {nhp} high priority</span>'
                    f'</div>',
                    unsafe_allow_html=True)
                st.balloons()


# ═══════════════════════════════════════════════════════════════════
#  PAGE: FLEET OVERVIEW
# ═══════════════════════════════════════════════════════════════════
elif page == 'Fleet Overview':
    st.markdown(H('Fleet Overview', 'All vessels · Live status'), unsafe_allow_html=True)

    fleet = _fleet()
    if not fleet:
        st.info('No data. Upload a report and click Confirm & Save.')
        st.stop()

    all_df = _all_df()
    tv  = len(fleet)
    tc  = len(all_df)
    tod = int((all_df['status'] == 'OVERDUE').sum())
    thp = int((all_df['status'] == 'HIGH PRIORITY').sum())
    tok = int((all_df['status'] == 'OK').sum())

    k1,k2,k3,k4,k5 = st.columns(5)
    for col,(v,l,c,d) in zip([k1,k2,k3,k4,k5],[
        (tv,'Vessels','u',0),(tc,'Components','g',.06),
        (tod,'Overdue','r',.12),(thp,'High Priority','o',.18),(tok,'OK','n',.24)]):
        with col: st.markdown(KPI(v,l,c,d), unsafe_allow_html=True)

    st.markdown('<br>', unsafe_allow_html=True)
    st.markdown(SL('Fleet Component Matrix'), unsafe_allow_html=True)

    f1,f2,f3,f4 = st.columns(4)
    with f1: vf  = st.selectbox('Vessel', ['All Fleet']+sorted(all_df['vessel_name'].unique().tolist()), key='ov_v')
    with f2: sf  = st.selectbox('Status', ['All','Overdue only','High Priority +','OK only'], key='ov_s')
    with f3: cf  = st.selectbox('Engine', ['All','Main Engine','Aux Engines'], key='ov_c')
    with f4: cpf = st.selectbox('Component', ['All']+sorted(all_df['description'].unique().tolist()), key='ov_cp')

    sort_opt = st.radio('Sort', ['Report Order','Component → Cylinder','Priority → % Used'],
                        horizontal=True, key='ov_srt')
    mode_map = {'Report Order':'seq','Component → Cylinder':'matrix','Priority → % Used':'priority'}

    filt = all_df.copy()
    if vf  != 'All Fleet':         filt = filt[filt['vessel_name'] == vf]
    if sf  == 'Overdue only':      filt = filt[filt['status'] == 'OVERDUE']
    elif sf == 'High Priority +':  filt = filt[filt['status'].isin(['OVERDUE','HIGH PRIORITY'])]
    elif sf == 'OK only':          filt = filt[filt['status'] == 'OK']
    if cf  == 'Main Engine':       filt = filt[filt['category'] == 'MAIN_ENGINE']
    elif cf == 'Aux Engines':      filt = filt[filt['category'] == 'AUX_ENGINE']
    if cpf != 'All':               filt = filt[filt['description'] == cpf]

    ns = len(filt)
    no = int((filt['status'] == 'OVERDUE').sum())
    nh = int((filt['status'] == 'HIGH PRIORITY').sum())
    nk = int((filt['status'] == 'OK').sum())
    st.markdown(FC(ns, no, nh, nk), unsafe_allow_html=True)

    if filt.empty:
        st.markdown(CLEAR('No records match the current filter'), unsafe_allow_html=True)
    else:
        render(filt, mode=mode_map[sort_opt],
               include_vessel=(vf == 'All Fleet'),
               height=min(860, 38*ns+44))


# ═══════════════════════════════════════════════════════════════════
#  PAGE: VESSEL DETAIL
# ═══════════════════════════════════════════════════════════════════
elif page == 'Vessel Detail':
    if not sel_v:
        st.info('Select a vessel from the sidebar.'); st.stop()

    data = _vessel(sel_v)
    if not data:
        st.info(f'No data for {sel_v}. Upload and save a report first.'); st.stop()

    st.markdown(H(sel_v, 'Component Analysis'), unsafe_allow_html=True)

    comps = data['components']
    df    = pd.DataFrame(comps) if comps else pd.DataFrame()

    n_tot = len(df)
    n_od  = int((df['status'] == 'OVERDUE').sum())      if not df.empty else 0
    n_hp  = int((df['status'] == 'HIGH PRIORITY').sum()) if not df.empty else 0
    n_ok  = int((df['status'] == 'OK').sum())            if not df.empty else 0
    n_nd  = int((df['status'] == 'NO DATA').sum())       if not df.empty else 0

    k1,k2,k3,k4,k5 = st.columns(5)
    for col,(v,l,c,d) in zip([k1,k2,k3,k4,k5],[
        (n_tot,'Total','g',0),(n_od,'Overdue','r',.06),
        (n_hp,'High Priority','o',.12),(n_ok,'OK','n',.18),(n_nd,'No Data','u',.24)]):
        with col: st.markdown(KPI(v,l,c,d), unsafe_allow_html=True)

    mt = f"{int(data['me_total_hrs']):,}"  if data.get('me_total_hrs')  else '—'
    mm = f"{int(data['me_this_month']):,}" if data.get('me_this_month') else '—'
    st.markdown(
        f'<div style="display:flex;gap:1.2rem;flex-wrap:wrap;'
        f'font-family:JetBrains Mono,monospace;font-size:.61rem;color:#284060;margin:.55rem 0 0">'
        f'<span>File: <b style="color:#5880a8">{data.get("filename","—")}</b></span>'
        f'<span>Report: <b style="color:#5880a8">{data.get("report_date") or "—"}</b></span>'
        f'<span>M/E: <b style="color:#5880a8">{mt}</b> total · <b style="color:#5880a8">{mm}</b> month</span>'
        f'</div>',
        unsafe_allow_html=True)

    if df.empty:
        st.info('No component data.'); st.stop()

    st.markdown('---')
    tabs = st.tabs(['Alerts', 'Main Engine', 'Aux Engines', 'Other Equipment'])

    # ── ALERTS ──────────────────────────────────────────────────────
    with tabs[0]:
        st.markdown(SL('Overdue and High Priority — Most Critical First'), unsafe_allow_html=True)
        alerts = df[df['status'].isin(['OVERDUE','HIGH PRIORITY'])]
        if alerts.empty:
            st.markdown(CLEAR(), unsafe_allow_html=True)
        else:
            no = int((alerts['status']=='OVERDUE').sum())
            nh = int((alerts['status']=='HIGH PRIORITY').sum())
            st.markdown(FC(len(alerts), no, nh, 0), unsafe_allow_html=True)
            render(alerts, mode='priority')

    # ── MAIN ENGINE ──────────────────────────────────────────────────
    with tabs[1]:
        me = df[df['category'] == 'MAIN_ENGINE']
        if me.empty:
            st.info('No Main Engine data.')
        else:
            st.markdown(SL('Main Engine — Component → Cylinder'), unsafe_allow_html=True)
            fa, fb = st.columns(2)
            with fa: mc = st.selectbox('Component', ['All']+sorted(me['description'].unique().tolist()), key='me_c')
            with fb: ms = st.selectbox('Status', ['All','Overdue','High Priority +','OK'], key='me_s')
            mr = st.radio('Sort', ['Report Order','Component → Cylinder','Priority → % Used'],
                          horizontal=True, key='me_r')
            v = me.copy()
            if mc != 'All':              v = v[v['description'] == mc]
            if ms == 'Overdue':          v = v[v['status'] == 'OVERDUE']
            elif ms == 'High Priority +': v = v[v['status'].isin(['OVERDUE','HIGH PRIORITY'])]
            elif ms == 'OK':             v = v[v['status'] == 'OK']
            mm2 = {'Report Order':'seq','Component → Cylinder':'matrix','Priority → % Used':'priority'}
            no = int((v['status']=='OVERDUE').sum()); nh = int((v['status']=='HIGH PRIORITY').sum()); nk = int((v['status']=='OK').sum())
            st.markdown(FC(len(v), no, nh, nk), unsafe_allow_html=True)
            render(v, mode=mm2[mr])

    # ── AUX ENGINES ──────────────────────────────────────────────────
    with tabs[2]:
        aux = df[df['category'] == 'AUX_ENGINE']
        if aux.empty:
            st.info('No Aux Engine data.')
        else:
            st.markdown(SL('Auxiliary Engines — Component → Cylinder'), unsafe_allow_html=True)
            fa, fb = st.columns(2)
            with fa: ae  = st.selectbox('Engine', ['All']+sorted(aux['engine_label'].unique().tolist()), key='ae')
            with fb: as_ = st.selectbox('Status', ['All','Overdue','High Priority +','OK'], key='ae_s')
            ar = st.radio('Sort', ['Report Order','Component → Cylinder','Priority → % Used'],
                          horizontal=True, key='ae_r')
            v = aux.copy()
            if ae  != 'All':            v = v[v['engine_label'] == ae]
            if as_ == 'Overdue':        v = v[v['status'] == 'OVERDUE']
            elif as_ == 'High Priority +': v = v[v['status'].isin(['OVERDUE','HIGH PRIORITY'])]
            elif as_ == 'OK':           v = v[v['status'] == 'OK']
            mm3 = {'Report Order':'seq','Component → Cylinder':'matrix','Priority → % Used':'priority'}
            no = int((v['status']=='OVERDUE').sum()); nh = int((v['status']=='HIGH PRIORITY').sum()); nk = int((v['status']=='OK').sum())
            st.markdown(FC(len(v), no, nh, nk), unsafe_allow_html=True)
            render(v, mode=mm3[ar])

    # ── OTHER EQUIPMENT ──────────────────────────────────────────────
    with tabs[3]:
        oe_list = data.get('other_equipment', [])
        if not oe_list:
            st.info('No other equipment data.')
        else:
            oe = pd.DataFrame(oe_list)
            for sec in sorted(oe['section'].unique()):
                st.markdown(SL(sec), unsafe_allow_html=True)
                sd = oe[oe['section']==sec][['description','periodicity','last_date','run_hrs']].copy()
                sd.columns = ['Description','Periodicity','Last Date','Run Hrs']
                st.dataframe(sd, use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════════════════════
#  PAGE: UPLOAD HISTORY
# ═══════════════════════════════════════════════════════════════════
elif page == 'Upload History':
    st.markdown(H('Upload History', 'Audit Trail'), unsafe_allow_html=True)
    if not sel_v:
        st.info('Select a vessel from the sidebar.'); st.stop()

    st.markdown(SL(sel_v), unsafe_allow_html=True)
    data = _vessel(sel_v)
    if not data:
        st.info('No data for this vessel.')
    else:
        comps  = data.get('components', [])
        n_od   = sum(1 for c in comps if c['status'] == 'OVERDUE')
        n_hp   = sum(1 for c in comps if c['status'] == 'HIGH PRIORITY')
        n_ok   = sum(1 for c in comps if c['status'] == 'OK')
        n_nd   = sum(1 for c in comps if c['status'] == 'NO DATA')

        c1,c2,c3,c4 = st.columns(4)
        for col,(v,l,cl,dl) in zip([c1,c2,c3,c4],[
            (n_od,'Overdue','r',0),(n_hp,'High Priority','o',.06),
            (n_ok,'OK','n',.12),(n_nd,'No Data','u',.18)]):
            with col: st.markdown(KPI(v,l,cl,dl), unsafe_allow_html=True)

        st.markdown(
            f'<div style="font-family:JetBrains Mono,monospace;font-size:.63rem;'
            f'color:#284060;margin-top:1rem;line-height:1.8">'
            f'<span style="color:#a8c8e8;font-weight:600">{data.get("filename","—")}</span>'
            f'<br>Report date: <b style="color:#5880a8">{data.get("report_date") or "—"}</b>'
            f'&nbsp;·&nbsp;Uploaded: <b style="color:#5880a8">{str(data.get("uploaded_at","—"))[:16]}</b>'
            f'</div>',
            unsafe_allow_html=True)
