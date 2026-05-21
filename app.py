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
from typing import Optional
import pandas as pd


# ═══════════════════════════════════════════════════════════════════
#  GLOBAL CSS  — injected once at startup
# ═══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

/* ── RESET ── */
*,*::before,*::after{box-sizing:border-box}
html,body,[class*="css"]{
  font-family:'Inter',sans-serif!important;
  background:#020509!important;
  color:#b0c8e8!important;
  -webkit-font-smoothing:antialiased;
}
.main,.main>div{background:#020509!important}
.block-container{max-width:100%!important;padding:1.75rem 2.5rem 5rem!important}

/* ── AMBIENT GLOW ── */
.main::before{
  content:'';position:fixed;inset:0;pointer-events:none;z-index:0;
  background:
    radial-gradient(ellipse 100% 60% at -15% -10%,rgba(184,135,12,.07) 0%,transparent 50%),
    radial-gradient(ellipse 80%  55% at 115% 110%,rgba(12,52,152,.06) 0%,transparent 50%);
}

/* ── SIDEBAR ── */
[data-testid="stSidebar"]{
  background:#050a14!important;
  border-right:1px solid #162640!important;
}
[data-testid="stSidebar"] *{color:#b0c8e8!important}
[data-testid="stSidebarContent"]{padding:1.25rem!important}
[data-testid="stSidebar"] .stSelectbox>div>div{
  background:#0a1224!important;border:1px solid #162640!important;border-radius:6px!important
}

/* ── HEADINGS ── */
h1{font-family:'Space Grotesk',sans-serif!important;font-size:1.75rem!important;
   font-weight:700!important;color:#edf4ff!important;letter-spacing:-.025em!important;line-height:1.15!important}
h2{font-family:'Space Grotesk',sans-serif!important;font-size:1.1rem!important;
   font-weight:600!important;color:#edf4ff!important}

/* ── METRICS ── */
[data-testid="stMetric"]{
  background:#0a1224!important;border:1px solid #162640!important;
  border-radius:10px!important;padding:.9rem 1.1rem 1rem!important;
  position:relative!important;overflow:hidden!important;
  transition:border-color .2s,transform .18s!important
}
[data-testid="stMetric"]:hover{border-color:#1e3354!important;transform:translateY(-2px)!important}
[data-testid="stMetricValue"]{font-family:'Space Grotesk',sans-serif!important;font-size:1.9rem!important;
  font-weight:700!important;color:#edf4ff!important;letter-spacing:-.035em!important}
[data-testid="stMetricLabel"]{color:#304060!important;font-size:.6rem!important;
  text-transform:uppercase!important;letter-spacing:.15em!important}

/* ── DATAFRAME ── */
[data-testid="stDataFrame"]{
  border:1px solid #162640!important;
  border-radius:10px!important;
  overflow:hidden!important;
  box-shadow:0 4px 24px rgba(0,0,0,.4)!important;
}
.dvn-scroller{background:#070d1c!important}

/* ── BUTTON ── */
.stButton>button{
  background:linear-gradient(135deg,#d4a018 0%,#b8870c 100%)!important;
  color:#000!important;border:none!important;border-radius:8px!important;
  padding:.6rem 1.75rem!important;font-family:'Space Grotesk',sans-serif!important;
  font-weight:700!important;font-size:.8rem!important;letter-spacing:.06em!important;
  text-transform:uppercase!important;
  box-shadow:0 2px 14px rgba(184,135,12,.22)!important;transition:all .17s!important
}
.stButton>button:hover{
  background:linear-gradient(135deg,#edbb2a 0%,#d4a018 100%)!important;
  box-shadow:0 5px 22px rgba(184,135,12,.4)!important;transform:translateY(-2px)!important
}

/* ── FILE UPLOADER ── */
[data-testid="stFileUploadDropzone"]{
  background:linear-gradient(155deg,rgba(184,135,12,.04) 0%,rgba(12,52,152,.03) 100%)!important;
  border:1.5px dashed #d4a018!important;border-radius:14px!important;
  padding:3rem 2rem!important;transition:all .28s!important
}
[data-testid="stFileUploadDropzone"]:hover{
  background:rgba(184,135,12,.07)!important;border-color:#edbb2a!important
}
[data-testid="stFileUploadDropzone"] p,
[data-testid="stFileUploadDropzone"] span{
  color:#edbb2a!important;font-family:'Space Grotesk',sans-serif!important;
  font-size:.92rem!important;font-weight:500!important
}
[data-testid="stFileUploadDropzone"] small{color:#304060!important}

/* ── TABS ── */
.stTabs [data-baseweb="tab-list"]{
  background:#070d1c!important;border-radius:10px 10px 0 0!important;
  border-bottom:1px solid #162640!important;gap:0!important;padding:0 .75rem!important
}
.stTabs [data-baseweb="tab"]{
  background:transparent!important;color:#304060!important;
  font-family:'Space Grotesk',sans-serif!important;font-weight:500!important;
  text-transform:uppercase!important;letter-spacing:.05em!important;font-size:.72rem!important;
  padding:.8rem 1.2rem!important;border-bottom:2px solid transparent!important;
  margin-bottom:-1px!important;transition:color .18s!important
}
.stTabs [data-baseweb="tab"]:hover{color:#6080a8!important}
.stTabs [aria-selected="true"]{color:#edbb2a!important;border-bottom:2px solid #b8870c!important}
.stTabs [data-baseweb="tab-panel"]{
  background:#070d1c!important;border:1px solid #162640!important;border-top:none!important;
  border-radius:0 0 10px 10px!important;padding:1.4rem!important
}

/* ── EXPANDER ── */
.streamlit-expanderHeader{
  background:#0a1224!important;border:1px solid #162640!important;
  border-radius:8px!important;font-family:'Space Grotesk',sans-serif!important;
  font-size:.83rem!important;color:#b0c8e8!important;transition:all .18s!important
}
.streamlit-expanderHeader:hover{background:#0e182e!important;border-color:#1e3354!important}
.streamlit-expanderContent{
  background:#070d1c!important;border:1px solid #162640!important;
  border-top:none!important;border-radius:0 0 8px 8px!important
}

/* ── INPUTS ── */
.stSelectbox>div>div,.stMultiSelect>div>div{
  background:#0a1224!important;border:1px solid #162640!important;
  border-radius:7px!important;color:#b0c8e8!important
}
.stSelectbox label,.stMultiSelect label{
  color:#304060!important;font-size:.68rem!important;
  text-transform:uppercase!important;letter-spacing:.1em!important
}
[data-testid="stRadio"]>label{
  color:#304060!important;font-size:.68rem!important;
  text-transform:uppercase!important;letter-spacing:.1em!important
}
[data-testid="stRadio"]>div{gap:.5rem!important}
[data-testid="stRadio"] label span{
  font-family:'Space Grotesk',sans-serif!important;font-size:.72rem!important;
  font-weight:500!important;color:#6080a8!important
}

/* ── MISC ── */
.stAlert{border-radius:8px!important;border-left-width:3px!important}
hr{border-color:#162640!important;opacity:1!important;margin:1.2rem 0!important}
a{color:#edbb2a!important;text-decoration:none!important}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:#050a14}
::-webkit-scrollbar-thumb{background:#1e3354;border-radius:3px}
::-webkit-scrollbar-thumb:hover{color:#304060}

/* ── ANIMATIONS ── */
@keyframes drop  {from{opacity:0;transform:translateY(-14px)}to{opacity:1;transform:translateY(0)}}
@keyframes rise  {from{opacity:0;transform:translateY(12px)} to{opacity:1;transform:translateY(0)}}
@keyframes push  {from{opacity:0;transform:translateX(-12px)}to{opacity:1;transform:translateX(0)}}
@keyframes scl   {from{opacity:0;transform:scale(.9)}        to{opacity:1;transform:scale(1)}}
@keyframes numup {from{opacity:0;transform:translateY(7px)}  to{opacity:1;transform:translateY(0)}}
@keyframes bar   {from{width:0;opacity:0}                    to{width:100%;opacity:1}}
</style>
""", unsafe_allow_html=True)



# ═══════════════════════════════════════════════════════════════════
#  DATABASE  — /tmp/ for Streamlit Cloud writability
# ═══════════════════════════════════════════════════════════════════
DB_PATH = Path("/tmp/running_hours.db")

def get_db():
    c = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    c.row_factory = sqlite3.Row; return c

def init_db():
    c = get_db()
    c.executescript("""
    PRAGMA journal_mode=WAL;
    CREATE TABLE IF NOT EXISTS vessels(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        created_at TEXT NOT NULL DEFAULT(datetime('now')));
    CREATE TABLE IF NOT EXISTS upload_log(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL, filename TEXT NOT NULL,
        file_hash TEXT NOT NULL, report_date TEXT,
        me_total_hrs INTEGER, me_this_month INTEGER,
        component_count INTEGER DEFAULT 0,
        uploaded_at TEXT NOT NULL DEFAULT(datetime('now')));
    CREATE TABLE IF NOT EXISTS components(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL, category TEXT NOT NULL,
        engine_label TEXT NOT NULL, unit TEXT NOT NULL,
        description TEXT NOT NULL, periodicity REAL,
        last_oh_date TEXT, last_oh_hrs REAL,
        hrs_since REAL, pct_used REAL, status TEXT NOT NULL,
        seq INTEGER DEFAULT 0,
        updated_at TEXT NOT NULL DEFAULT(datetime('now')));
    CREATE TABLE IF NOT EXISTS other_equipment(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL, section TEXT NOT NULL,
        description TEXT NOT NULL, periodicity TEXT,
        last_date TEXT, run_hrs TEXT,
        updated_at TEXT NOT NULL DEFAULT(datetime('now')));
    CREATE INDEX IF NOT EXISTS idx_cv   ON components(vessel_name);
    CREATE INDEX IF NOT EXISTS idx_cs   ON components(status);
    CREATE INDEX IF NOT EXISTS idx_cseq ON components(vessel_name, seq);
    """)
    c.commit(); c.close()

init_db()


# ═══════════════════════════════════════════════════════════════════
#  CONVERSION
# ═══════════════════════════════════════════════════════════════════
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
            [soffice,"--headless","--norestore","--nofirststartwizard",
             f"-env:UserInstallation={pf}","--convert-to","docx",src,"--outdir",outdir],
            capture_output=True, timeout=120)
        if not os.path.exists(out):
            raise RuntimeError(r.stderr.decode("utf-8","ignore")[:400])
        with open(out,"rb") as f: return f.read()
    finally:
        for p in [src, out]:
            try:
                if os.path.exists(p): os.unlink(p)
            except Exception: pass
        shutil.rmtree(outdir, ignore_errors=True)


# ═══════════════════════════════════════════════════════════════════
#  PARSER  — Architecturally Correct
#
#  Phase 1: Isolated ME/AUX parsers with hardcoded category tags
#  Phase 2: Strict right-side wall on REMARKS column
#  Phase 3: Car-wash sanitizer strips vessel metadata artifacts
#  Phase 4: Exact-match rejection of row-type markers
#  Phase 5: Dynamic anchor mapping from X-Ray validated column positions
# ═══════════════════════════════════════════════════════════════════

# ── Cell utilities ────────────────────────────────────────────────

def _ucells(row):
    """First occurrence of each physical cell (_tc identity deduplication)."""
    seen, out = set(), []
    for ci, cell in enumerate(row.cells):
        cid = id(cell._tc)
        if cid not in seen:
            seen.add(cid); out.append((ci, cell))
    return out

def _all_cells(row):
    """All cells including merged duplicates, as {col_index: text}."""
    return {ci: cell.text.strip() for ci, cell in enumerate(row.cells)}

# ── Phase 3: Car-wash sanitizer ───────────────────────────────────

# Artifacts that Word injects into cells from page headers / footers
_ARTIFACTS = re.compile(
    r'\b(ALEXIS|Date|Page|Signature|Vessel|MV|M/V)\b'
    r'|[\u200b\u200c\u200d\ufeff\xa0]',
    re.IGNORECASE
)

def _wash(txt: str) -> str:
    """Strip known Word export artifacts and normalise whitespace."""
    s = _ARTIFACTS.sub(' ', str(txt or ''))
    return re.sub(r'\s+', ' ', s).strip()

# ── Phase 4: Strict type converters ──────────────────────────────

# Strings that are row-type markers, not data
_REJECT_EXACT = {'1','2','1/','/ 2','-','--','n/a','na',''}

def _date(raw: str) -> Optional[str]:
    """Parse date. Returns None for markers, remarks text, and unparseable strings."""
    s = _wash(raw).strip()
    if not s or s.lower() in _REJECT_EXACT: return None
    # Phase 4: reject if more than 20 chars (free text / remarks)
    if len(s) > 20: return None
    # Pure number = hours value, not date
    if re.match(r'^\d+$', s.replace(',','').replace('.','')): return None
    rn = re.sub(r'\bSEPT\b','SEP', s, flags=re.I)
    rn = re.sub(r'\bJUNE\b','JUN', rn, flags=re.I)
    rn = re.sub(r'\bJULY\b','JUL', rn, flags=re.I)
    for fmt in ('%d %b %y','%d %B %y','%d %b %Y','%d %B %Y',
                '%d/%m/%y','%d/%m/%Y','%d-%m-%y','%d-%m-%Y',
                '%b %Y','%B %Y','%Y-%m-%d'):
        for v in (rn, rn.upper(), rn.title(), s, s.upper()):
            try: return datetime.strptime(v, fmt).strftime('%Y-%m-%d')
            except ValueError: pass
    return None  # reject unparseable (remarks text)

def _hrs(raw: str) -> Optional[float]:
    """
    Extract running hours.
    Phase 4: reject standalone row-type markers, brackets, free text.
    """
    s = _wash(raw).strip().replace(',','')
    # Strip brackets e.g. '[608]' -> '608'
    s = s.strip('[]')
    if not s or s.lower() in _REJECT_EXACT: return None
    # Must start with digit
    if not re.match(r'^\d', s): return None
    # Phase 4: reject if unreasonably long (free text contamination)
    if len(s) > 8: return None
    m = re.search(r'\d+', s)
    if not m: return None
    v = float(m.group())
    return v if v > 0 else None

def _per(raw: str) -> Optional[float]:
    s = _wash(raw).strip()
    s = re.sub(r'\.(?=\d{3}(\.|$))', '', s)
    s = re.sub(r'[^0-9.]', '', s)
    try: return float(s) if s else None
    except: return None

def _stat(h, p) -> str:
    if h is None or p is None or p == 0: return 'NO DATA'
    r = h / p
    if r > 1.0: return 'OVERDUE'
    if r >= 0.80: return 'HIGH PRIORITY'
    return 'OK'

def _pct(h, p) -> float:
    if h is None or p is None or p == 0: return 0.0
    return round(h / p, 4)

def _is_comp(txt: str) -> bool:
    """True if this looks like a component name, not a header or artifact."""
    t = _wash(txt).strip()
    if not t or len(t) < 2: return False
    if re.match(r'^1[\-\s]?DATE OF', t, re.I): return False
    if re.match(r'^(PERIODICITY|CYL|MAIN ENGINE|TYPE:|TOTAL RUNNING|THIS MONTH|REMARKS|DATE|DESCRIPTION|AUX\.)$', t, re.I): return False
    if not re.search(r'[A-Za-z]', t): return False
    return True


# ── Phase 1 + 2 + 5: Main Engine parser ───────────────────────────

def _parse_me(table) -> list:
    """
    Parse Main Engine table (Table 0).
    Phase 2: Hard stop at REMARKS column boundary.
    Phase 5: Column positions anchored from X-Ray validated layout.
    """
    comps = []; seq = 0
    rows = table.rows
    if len(rows) < 2: return comps

    # Detect cylinder columns using _tc identity (Phase 2: note REMARKS col)
    cyl_cols = []    # [(data_col_index, "Cyl N")]
    remarks_col = None
    for ci, cell in _ucells(rows[1]):
        txt = cell.text.strip()
        if re.match(r'REMARKS', txt, re.I):
            remarks_col = ci  # Phase 2: hard wall
            break
        if m := re.search(r'CYL\s*\.?\s*No\s*\.?\s*(\d+)', txt, re.I):
            cyl_cols.append((ci, f"Cyl {int(m.group(1))}"))

    if not cyl_cols: return comps

    # Phase 2: filter out any cyl col that reaches or exceeds REMARKS
    if remarks_col is not None:
        cyl_cols = [(ci, lbl) for ci, lbl in cyl_cols if ci < remarks_col]

    i = 2
    while i < len(rows) - 1:
        # Phase 3: read unique cells as dict
        r1 = {ci: _wash(cell.text) for ci, cell in _ucells(rows[i])}
        r2 = {ci: _wash(cell.text) for ci, cell in _ucells(rows[i+1])} if i+1 < len(rows) else {}

        name = r1.get(0, '')
        if not _is_comp(name): i += 1; continue

        type1 = r1.get(2, '').strip()
        type2 = r2.get(2, '').strip()
        name2 = r2.get(0, '')

        if type1 == '1' and type2 == '2' and name == name2:
            p = _per(r1.get(1, ''))
            for ci, lbl in cyl_cols:
                # Phase 2: hard stop enforced by cyl_cols already filtered
                d = _date(r1.get(ci, ''))
                h = _hrs(r2.get(ci, ''))
                if d is None and h is None: continue
                comps.append({
                    'seq': seq, 'category': 'MAIN_ENGINE', 'engine_label': 'ME',
                    'unit': lbl, 'description': name, 'periodicity': p,
                    'last_oh_date': d, 'last_oh_hrs': h, 'hrs_since': h,
                    'pct_used': _pct(h, p), 'status': _stat(h, p),
                })
                seq += 1
            i += 2
        else:
            i += 1
    return comps


# ── Phase 1 + 5: Aux Engine parser ────────────────────────────────

def _build_aux_cyl_map(table) -> dict:
    """
    Build data_col -> (engine_label, cyl_label) mapping for AUX engine table.

    Phase 5 algorithm (X-Ray validated):
    1. From row 0 UNIQUE: find engine block start columns
    2. From row 4 ALL cells with _tc: group merged cols by cyl number per engine block
    3. For each cyl group: data_col = cols[-2] if len >= 2 else cols[-1]
       (avoids the boundary-bleed col at the end of each group)
    4. Deduplicate by (engine_label, cyl_num) — keep first occurrence only
    """
    rows = table.rows
    if len(rows) < 5: return {}

    # Engine block start columns from row 0 unique
    r0u = {ci: cell.text.strip() for ci, cell in _ucells(rows[0])}
    eng_starts = sorted([
        (int(re.search(r'No\.?\s*(\d+)', txt, re.I).group(1)), ci)
        for ci, txt in r0u.items()
        if re.search(r'Aux\.\s*Engine\s*No', txt, re.I)
    ])
    if not eng_starts: return {}

    # Cyl groups from row 4 ALL cells
    r4_all = [(ci, cell.text.strip(), id(cell._tc)) for ci, cell in enumerate(rows[4].cells)]

    cyl_map = {}        # data_col -> (eng_lbl, cyl_lbl)
    seen_pairs = set()  # (eng_lbl, cyl_num) — prevent duplicate from boundary bleed

    for idx, (eng_num, eng_start) in enumerate(eng_starts):
        eng_end = eng_starts[idx+1][1] if idx+1 < len(eng_starts) else 9999
        lbl = f"AUX-{eng_num}"

        # Group cols by _tc within this engine block, keep only cyl-number cells
        tc_groups = {}
        for ci, txt, tc_id in r4_all:
            if eng_start <= ci < eng_end and re.match(r'^\d+$', txt.strip()):
                if tc_id not in tc_groups:
                    tc_groups[tc_id] = {'cyl_num': int(txt), 'cols': []}
                tc_groups[tc_id]['cols'].append(ci)

        for tc_id, info in sorted(tc_groups.items(), key=lambda x: x[1]['cols'][0]):
            cyl_num = info['cyl_num']
            pair = (lbl, cyl_num)
            if pair in seen_pairs: continue  # skip boundary-bleed duplicate
            seen_pairs.add(pair)
            cols = sorted(info['cols'])
            # Use second-to-last col to avoid the boundary transition column
            data_col = cols[-2] if len(cols) >= 2 else cols[-1]
            cyl_map[data_col] = (lbl, f"Cyl {cyl_num}")

    return cyl_map


def _parse_aux(table, seq_start: int = 1000) -> list:
    """
    Parse Auxiliary Engine table (Table 2).
    Phase 1: ALL results tagged 'AUX_ENGINE' — no cross-contamination with ME.
    """
    comps = []; seq = seq_start
    rows = table.rows
    if len(rows) < 6: return comps

    cyl_map = _build_aux_cyl_map(table)
    if not cyl_map: return comps

    i = 5
    while i < len(rows) - 1:
        r1 = {ci: _wash(cell.text) for ci, cell in _ucells(rows[i])}
        r2 = {ci: _wash(cell.text) for ci, cell in _ucells(rows[i+1])} if i+1 < len(rows) else {}

        name = r1.get(0, '')
        if not _is_comp(name): i += 1; continue

        type1 = r1.get(2, '').strip()
        type2 = r2.get(2, '').strip()
        name2 = r2.get(0, '')

        if type1 in ('1', '2') and name == name2:
            p = _per(r1.get(1, ''))
            # Use ALL cells for data (not unique) to access exact column positions
            r1_all = _all_cells(rows[i])
            r2_all = _all_cells(rows[i+1]) if i+1 < len(rows) else {}
            for dc, (elbl, ulbl) in sorted(cyl_map.items()):
                d = _date(r1_all.get(dc, ''))
                h = _hrs(r2_all.get(dc, ''))
                if d is None and h is None: continue
                comps.append({
                    'seq': seq, 'category': 'AUX_ENGINE', 'engine_label': elbl,
                    'unit': ulbl, 'description': name, 'periodicity': p,
                    'last_oh_date': d, 'last_oh_hrs': h, 'hrs_since': h,
                    'pct_used': _pct(h, p), 'status': _stat(h, p),
                })
                seq += 1
            i += 2
        else:
            i += 1
    return comps


# ── Table 1: Turbocharger / Coolers / A/C ─────────────────────────

def _parse_table1(table) -> list:
    """Parse the Turbo / Coolers / A/C table (Table 1)."""
    oe = []
    SKIP = {'TURBOCHARGER','AUXILIARY BOILER','COOLERS','EXH GAS  BOILER',
            'A/C & REFR. COMPRESSORS','MAIN AIR COMPRESSORS',
            'PERIODICTLY','DATE OF LAST INSPECTION','RUN HRS',
            'DATE OF LAST CLEANING','DATE','PERIODICITY',''}
    for row in table.rows:
        cells = {ci: cell.text.strip() for ci, cell in _ucells(row)}
        for sec, dc, datec, hrsc in [
            ('TURBOCHARGER / AUX BOILER', 0, 2, 3),
            ('COOLERS / EXH GAS BOILER',  5, 6, 7),
            ('A/C & COMPRESSORS',         10,11,12),
        ]:
            desc = cells.get(dc, '')
            if not desc or desc.upper() in SKIP: continue
            dv = cells.get(datec, ''); hv = cells.get(hrsc, '')
            if dv or hv:
                oe.append({'section': sec, 'description': desc,
                           'periodicity': '', 'last_date': dv, 'run_hrs': hv})
    return oe


# ── Table 3: D/G Equipment ────────────────────────────────────────

def _parse_table3(table) -> list:
    oe = []; dg = ['D/G 1','D/G 2','D/G 3']
    for ri, row in enumerate(table.rows):
        cells = {ci: cell.text.strip() for ci, cell in _ucells(row)}
        if ri == 0: continue
        for dc, pc, tc, ds in [(0,1,2,3),(9,10,11,12)]:
            desc = cells.get(dc,''); per = cells.get(pc,''); rt = cells.get(tc,'')
            if not desc or rt not in ('1','2'): continue
            for gi, gl in enumerate(dg):
                val = cells.get(ds+gi,'')
                if not val: continue
                key = f"{desc} — {gl}"
                if rt == '1':
                    oe.append({'section':'D/G EQUIPMENT','description':key,
                               'periodicity':per,'last_date':_date(val) or val,'run_hrs':''})
                else:
                    for e in reversed(oe):
                        if e['description']==key and e['run_hrs']=='':
                            e['run_hrs']=val; break
                    else:
                        oe.append({'section':'D/G EQUIPMENT','description':key,
                                   'periodicity':per,'last_date':'','run_hrs':val})
    return oe


# ── Master parse function ─────────────────────────────────────────

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

    # Vessel name + report date
    vn = 'UNKNOWN'; rd = None
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt: continue
        vm = re.search(r"Vessel[\u2019\u2018']?s?\s+Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+?)(?:\s{2,}|\t)", txt, re.I)
        dm = re.search(r"Date\s*:\s*(.+)", txt, re.I)
        if vm: vn = vm.group(1).strip()
        if dm: rd = _date(dm.group(1).strip())
        if vm or dm: break
    if vn == 'UNKNOWN': warns.append("Could not extract vessel name from header.")

    # M/E totals
    me_tot = me_mo = None
    for _, cell in _ucells(doc.tables[0].rows[0]):
        x = cell.text
        if m := re.search(r'Total Running Hours[\s:ǀ|]+([\d,]+)', x, re.I):
            try: me_tot = int(m.group(1).replace(',',''))
            except: pass
        if m := re.search(r'This Month[\s:]+([\d,]+)', x, re.I):
            try: me_mo = int(m.group(1).replace(',',''))
            except: pass

    # Phase 1: completely isolated parsers
    comps = _parse_me(doc.tables[0])
    if len(doc.tables) > 2:
        comps += _parse_aux(doc.tables[2], seq_start=len(comps))

    # Other equipment
    oe = []
    if len(doc.tables) > 1: oe += _parse_table1(doc.tables[1])
    if len(doc.tables) > 3: oe += _parse_table3(doc.tables[3])

    if not comps:
        warns.append("No components extracted. Check document structure.")

    return {
        'vessel_name': vn, 'report_date': rd,
        'me_total_hrs': me_tot, 'me_this_month': me_mo,
        'components': comps, 'other_equipment': oe, 'warnings': warns,
    }


# ═══════════════════════════════════════════════════════════════════
#  DB HELPERS
# ═══════════════════════════════════════════════════════════════════
def save_parsed(parsed, filename, fhash):
    conn = get_db(); c = conn.cursor()
    now = datetime.utcnow().isoformat()+'Z'; v = parsed['vessel_name']
    try:
        c.execute("INSERT OR IGNORE INTO vessels(name,created_at) VALUES(?,?)", (v,now))
        c.execute("INSERT INTO upload_log(vessel_name,filename,file_hash,report_date,me_total_hrs,me_this_month,component_count,uploaded_at) VALUES(?,?,?,?,?,?,?,?)",
            (v,filename,fhash,parsed['report_date'],parsed['me_total_hrs'],parsed['me_this_month'],len(parsed['components']),now))
        c.execute("DELETE FROM components WHERE vessel_name=?", (v,))
        for x in parsed['components']:
            c.execute("INSERT INTO components(vessel_name,category,engine_label,unit,description,periodicity,last_oh_date,last_oh_hrs,hrs_since,pct_used,status,seq,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (v,x['category'],x['engine_label'],x['unit'],x['description'],
                 x['periodicity'],x['last_oh_date'],x['last_oh_hrs'],
                 x['hrs_since'],x['pct_used'],x['status'],x['seq'],now))
        c.execute("DELETE FROM other_equipment WHERE vessel_name=?", (v,))
        for x in parsed['other_equipment']:
            c.execute("INSERT INTO other_equipment(vessel_name,section,description,periodicity,last_date,run_hrs,updated_at) VALUES(?,?,?,?,?,?,?)",
                (v,x['section'],x['description'],x.get('periodicity',''),x.get('last_date',''),x.get('run_hrs',''),now))
        conn.commit()
    except Exception: conn.rollback(); raise
    finally: conn.close()

@st.cache_data(ttl=10)
def get_vessels():
    c=get_db(); r=c.execute("SELECT name FROM vessels ORDER BY name").fetchall(); c.close()
    return [x['name'] for x in r]

@st.cache_data(ttl=10)
def get_comps(vessel):
    c=get_db()
    df=pd.read_sql_query("SELECT * FROM components WHERE vessel_name=? ORDER BY seq",c,params=(vessel,))
    c.close(); return df

@st.cache_data(ttl=10)
def get_oe(vessel):
    c=get_db()
    df=pd.read_sql_query("SELECT * FROM other_equipment WHERE vessel_name=? ORDER BY section,description",c,params=(vessel,))
    c.close(); return df

@st.cache_data(ttl=10)
def get_history(vessel):
    c=get_db()
    df=pd.read_sql_query("SELECT filename,report_date,me_total_hrs,me_this_month,component_count,uploaded_at FROM upload_log WHERE vessel_name=? ORDER BY uploaded_at DESC LIMIT 20",c,params=(vessel,))
    c.close(); return df

@st.cache_data(ttl=10)
def get_summary():
    c=get_db()
    df=pd.read_sql_query("""
        SELECT c.vessel_name,
            SUM(CASE WHEN c.status='OVERDUE'       THEN 1 ELSE 0 END) AS overdue,
            SUM(CASE WHEN c.status='HIGH PRIORITY' THEN 1 ELSE 0 END) AS high_priority,
            SUM(CASE WHEN c.status='OK'            THEN 1 ELSE 0 END) AS ok,
            COUNT(*) AS total,
            MAX(u.uploaded_at) AS last_upload,
            MAX(u.me_total_hrs) AS me_total_hrs
        FROM components c
        LEFT JOIN upload_log u ON u.vessel_name=c.vessel_name
        GROUP BY c.vessel_name ORDER BY overdue DESC,high_priority DESC
    """,c); c.close(); return df

@st.cache_data(ttl=10)
def get_all_comps():
    c=get_db()
    df=pd.read_sql_query("SELECT * FROM components ORDER BY vessel_name,seq",c)
    c.close(); return df


# ═══════════════════════════════════════════════════════════════════
#  TABLE ENGINE  — dynamic styling, any column count
# ═══════════════════════════════════════════════════════════════════

_TH = {
    'OVERDUE':       {'bg':'#1f0505','bgs':'#2d0707','ts':'#ff6b6b','tm':'#ff8888','tn':'#ff2222','td':'#663030'},
    'HIGH PRIORITY': {'bg':'#1e0d02','bgs':'#2c1403','ts':'#ffaa44','tm':'#ffa030','tn':'#ffcc00','td':'#664422'},
    'OK':            {'bg':'#021208','bgs':'#031c0c','ts':'#40d880','tm':'#20b860','tn':'#40d880','td':'#0e4020'},
    'NO DATA':       {'bg':'#080e18','bgs':'#0b1420','ts':'#4878aa','tm':'#305878','tn':'#4878aa','td':'#182838'},
}

_CS = {
    'Status':      lambda c: f"background-color:{c['bgs']};color:{c['ts']};font-weight:700",
    'Vessel':      lambda c: f"background-color:{c['bg']};color:{c['tm']};font-weight:600",
    'Component':   lambda c: f"background-color:{c['bg']};color:{c['tm']};font-weight:600",
    'Engine':      lambda c: f"background-color:{c['bg']};color:{c['td']}",
    'Unit':        lambda c: f"background-color:{c['bg']};color:{c['td']}",
    'Periodicity': lambda c: f"background-color:{c['bg']};color:{c['td']}",
    'Last O/H':    lambda c: f"background-color:{c['bg']};color:{c['td']}",
    'Hrs Since':   lambda c: f"background-color:{c['bg']};color:{c['tm']};font-weight:600",
    'Used':        lambda c: f"background-color:{c['bg']};color:{c['tn']};font-weight:700",
}
_CS_DEFAULT = lambda c: f"background-color:{c['bg']};color:{c['td']}"




# ═══════════════════════════════════════════════════════════════════
#  TABLE ENGINE  — dynamic styling, bulletproof for any column count
# ═══════════════════════════════════════════════════════════════════

_ORD = {'OVERDUE': 0, 'HIGH PRIORITY': 1, 'OK': 2, 'NO DATA': 3}

_TH = {
    'OVERDUE': {
        'bg':  'background-color:#1c0303',
        'bgs': 'background-color:#280404',
        'ts':  'color:#ff5555;font-weight:700',
        'tm':  'color:#ff7070',
        'tn':  'color:#ff2020;font-weight:700',
        'td':  'color:#7a2020',
    },
    'HIGH PRIORITY': {
        'bg':  'background-color:#1c0e02',
        'bgs': 'background-color:#271502',
        'ts':  'color:#ffaa33;font-weight:700',
        'tm':  'color:#ff9922',
        'tn':  'color:#ffcc00;font-weight:700',
        'td':  'color:#7a4d10',
    },
    'OK': {
        'bg':  'background-color:#011008',
        'bgs': 'background-color:#01180a',
        'ts':  'color:#33dd77;font-weight:700',
        'tm':  'color:#22bb55',
        'tn':  'color:#33dd77;font-weight:700',
        'td':  'color:#0d4422',
    },
    'NO DATA': {
        'bg':  'background-color:#070c18',
        'bgs': 'background-color:#0a1020',
        'ts':  'color:#4477aa;font-weight:600',
        'tm':  'color:#335577',
        'tn':  'color:#4477aa',
        'td':  'color:#1a2a40',
    },
}

# Per-column style rule (callable with theme dict)
_CS = {
    'Status':      lambda t: f"{t['bgs']};{t['ts']}",
    'Vessel':      lambda t: f"{t['bg']};{t['tm']};font-weight:600",
    'Component':   lambda t: f"{t['bg']};{t['tm']};font-weight:600",
    'Engine':      lambda t: f"{t['bg']};{t['td']}",
    'Unit':        lambda t: f"{t['bg']};{t['td']}",
    'Periodicity': lambda t: f"{t['bg']};{t['td']}",
    'Last O/H':    lambda t: f"{t['bg']};{t['td']}",
    'Hrs Since':   lambda t: f"{t['bg']};{t['tm']};font-weight:600",
    'Used':        lambda t: f"{t['bg']};{t['tn']}",
}
_CS_DEF = lambda t: f"{t['bg']};{t['td']}"


def _sf(x):
    try:
        v = float(x)
        return None if pd.isna(v) else v
    except Exception:
        return None


def _cyl_n(u) -> int:
    m = re.search(r'\d+', str(u))
    return int(m.group()) if m else 999


def build_table(df: pd.DataFrame, mode: str = 'seq',
                include_vessel: bool = False) -> pd.DataFrame:
    """
    Build display-ready DataFrame.
    mode: 'seq' = document order | 'matrix' = component A-Z → cyl 1-N
          'priority' = OVERDUE first, then by % used desc
    """
    if df.empty:
        cols = ['Status', 'Component', 'Engine', 'Unit',
                'Periodicity', 'Last O/H', 'Hrs Since', 'Used']
        if include_vessel:
            cols.insert(1, 'Vessel')
        return pd.DataFrame(columns=cols)

    d = df.copy()
    if 'seq' not in d.columns:
        d['seq'] = range(len(d))
    if 'vessel_name' not in d.columns:
        d['vessel_name'] = ''

    if mode == 'seq':
        d = d.sort_values(['vessel_name', 'seq'], ascending=True)
    elif mode == 'matrix':
        d['_k1'] = d['description'].str.upper()
        d['_k2'] = d['unit'].apply(_cyl_n)
        d = d.sort_values(['_k1', '_k2']).drop(columns=['_k1', '_k2'])
    elif mode == 'priority':
        d['_s'] = d['status'].map(lambda x: _ORD.get(str(x), 4))
        d['_p'] = d['pct_used'].apply(lambda x: _sf(x) or 0.0)
        d = d.sort_values(['_s', '_p'], ascending=[True, False])
        d = d.drop(columns=['_s', '_p'])

    n = len(d)
    out = pd.DataFrame(index=range(n))
    out['Status']      = d['status'].values
    if include_vessel:
        out['Vessel']  = d['vessel_name'].values
    out['Component']   = d['description'].values
    out['Engine']      = d['engine_label'].values
    out['Unit']        = d['unit'].values
    out['Periodicity'] = [
        int(float(x)) if _sf(x) else None
        for x in d['periodicity'].values
    ]
    out['Last O/H']    = [
        str(x) if x and str(x) not in ('nan', 'None', '') else '—'
        for x in d['last_oh_date'].values
    ]
    out['Hrs Since']   = [
        int(float(x)) if _sf(x) else None
        for x in d['hrs_since'].values
    ]
    out['Used']        = [
        round(float(x) * 100, 1) if _sf(x) else 0.0
        for x in d['pct_used'].values
    ]
    return out


def style_table(df: pd.DataFrame):
    """
    Apply per-row background + text colours.
    Builds style list dynamically from df.columns — works for any column count.
    """
    cols = list(df.columns)

    def _row(row):
        t = _TH.get(str(row.get('Status', '')), _TH['NO DATA'])
        return [_CS.get(col, _CS_DEF)(t) for col in cols]

    return df.style.apply(_row, axis=1)


_CC = {
    'Status':      st.column_config.TextColumn('Status',      width=130),
    'Vessel':      st.column_config.TextColumn('Vessel',      width=120),
    'Component':   st.column_config.TextColumn('Component',   width=220),
    'Engine':      st.column_config.TextColumn('Engine',      width=80),
    'Unit':        st.column_config.TextColumn('Unit',        width=68),
    'Periodicity': st.column_config.NumberColumn('Periodicity', format='%d hrs',  width=110),
    'Last O/H':    st.column_config.TextColumn('Last O/H',    width=105),
    'Hrs Since':   st.column_config.NumberColumn('Hrs Since',  format='%d hrs',   width=105),
    'Used':        st.column_config.ProgressColumn(
                       'Used %', min_value=0, max_value=160,
                       format='%.1f%%', width=130),
}


def render_table(df, mode: str = 'seq',
                 include_vessel: bool = False, height: int = None):
    """Single entry point — always renders correctly."""
    if isinstance(df, list):
        df = pd.DataFrame(df)
    if df is None or df.empty:
        st.info('No data to display.')
        return
    tbl = build_table(df, mode=mode, include_vessel=include_vessel)
    if tbl.empty:
        st.info('No data to display.')
        return
    cfg = {k: v for k, v in _CC.items() if k in tbl.columns}
    h = height or min(900, 38 * len(tbl) + 44)
    st.dataframe(
        style_table(tbl),
        use_container_width=True,
        hide_index=True,
        height=h,
        column_config=cfg,
    )


# ═══════════════════════════════════════════════════════════════════
#  UI HELPERS
# ═══════════════════════════════════════════════════════════════════

def _kpi(val, lbl, color='gold', delay=0.0):
    c = {
        'gold': ('rgba(184,135,12,.28)', '#b8870c', '#f7d060'),
        'red':  ('rgba(184,32,32,.22)',  '#e84040', '#ff7070'),
        'ora':  ('rgba(168,72,8,.22)',   '#d06020', '#f08840'),
        'grn':  ('rgba(8,104,56,.22)',   '#10a058', '#2ed07a'),
        'blu':  ('rgba(12,52,152,.22)',  '#1a60d0', '#5090f0'),
    }.get(color, ('rgba(184,135,12,.28)', '#b8870c', '#f7d060'))
    border_col, accent, text_col = c
    return f'''<div style="
        background:linear-gradient(160deg,rgba(255,255,255,.02) 0%,transparent 100%);
        border:1px solid {border_col};
        border-top:2px solid {accent};
        border-radius:10px;
        padding:1rem 1.2rem 1.1rem;
        animation:rise .4s ease both;
        animation-delay:{delay}s;
        animation-fill-mode:both;
        transition:transform .18s,box-shadow .2s;
    " onmouseover="this.style.transform='translateY(-3px)';this.style.boxShadow='0 12px 32px rgba(0,0,0,.5)'"
       onmouseout="this.style.transform='';this.style.boxShadow=''">
        <div style="font-family:'Space Grotesk',sans-serif;font-size:2.1rem;
                    font-weight:700;color:{text_col};letter-spacing:-.04em;
                    line-height:1.1;animation:numup .4s .1s ease both;
                    animation-delay:{delay+.1}s;animation-fill-mode:both">
            {val}
        </div>
        <div style="font-family:'Inter',sans-serif;font-size:.58rem;font-weight:500;
                    text-transform:uppercase;letter-spacing:.18em;
                    color:#304060;margin-top:5px">
            {lbl}
        </div>
    </div>'''


def _ph(title: str, eyebrow: str = '') -> str:
    eye = (f'<div style="font-family:Inter,sans-serif;font-size:.56rem;font-weight:500;'
           f'letter-spacing:.26em;text-transform:uppercase;color:#b8870c;'
           f'margin-bottom:.25rem">{eyebrow}</div>') if eyebrow else ''
    return f'''<div style="animation:drop .4s cubic-bezier(.22,1,.36,1) both">
        {eye}
        <h1 style="margin:0;font-family:Space Grotesk,sans-serif;font-size:1.75rem;
                   font-weight:700;color:#edf4ff;letter-spacing:-.025em;line-height:1.15">
            {title}
        </h1>
    </div>
    <div style="height:1px;margin:.3rem 0 1.5rem;
                background:linear-gradient(90deg,#b8870c 0%,#162640 30%,transparent 100%);
                animation:bar .6s .1s ease both;animation-fill-mode:both"></div>'''


def _sl(txt: str) -> str:
    return f'''<div style="
        font-family:Inter,sans-serif;font-size:.56rem;font-weight:600;
        letter-spacing:.22em;text-transform:uppercase;color:#304060;
        display:flex;align-items:center;gap:.7rem;margin:1.5rem 0 .85rem">
        {txt}
        <span style="flex:1;height:1px;
                     background:linear-gradient(90deg,#162640,transparent)"></span>
    </div>'''


def _fc(n: int, od: int, hp: int, ok: int) -> str:
    return f'''<div style="font-family:'JetBrains Mono',monospace;font-size:.62rem;
                           color:#304060;margin-bottom:.65rem;line-height:1.6">
        <b style="color:#b0c8e8">{n}</b> records —
        <span style="color:#ff5555;font-weight:700">{od} overdue</span> ·
        <span style="color:#ffaa33;font-weight:700">{hp} high priority</span> ·
        <span style="color:#33dd77;font-weight:700">{ok} OK</span>
    </div>'''


def _ps_row(*items) -> str:
    """items: list of (value, label, color_class)"""
    cols = {'red':'#ff5555','ora':'#ffaa33','grn':'#33dd77','blu':'#5090f0'}
    accent = {'red':'#e84040','ora':'#d06020','grn':'#10a058','blu':'#1a60d0'}
    cards = ''
    for i, (val, lbl, clr) in enumerate(items):
        c = cols.get(clr, '#5090f0')
        a = accent.get(clr, '#1a60d0')
        cards += f'''<div style="background:#0a1224;border:1px solid #162640;
            border-top:2px solid {a};border-radius:8px;
            padding:.55rem .95rem .6rem;min-width:80px;
            animation:scl .3s ease both;animation-delay:{i*.06}s;animation-fill-mode:both">
            <div style="font-family:Space Grotesk,sans-serif;font-size:1.6rem;
                        font-weight:700;line-height:1;letter-spacing:-.03em;color:{c}">{val}</div>
            <div style="font-family:Inter,sans-serif;font-size:.55rem;text-transform:uppercase;
                        letter-spacing:.14em;color:#304060;margin-top:3px">{lbl}</div>
        </div>'''
    return f'<div style="display:flex;gap:.55rem;flex-wrap:wrap;margin:.85rem 0 1.3rem">{cards}</div>'


# ═══════════════════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown('''
    <div style="font-family:Space Grotesk,sans-serif;font-size:1.05rem;font-weight:700;
                letter-spacing:.04em;color:#edbb2a;display:flex;align-items:center;gap:.4rem">
        ⚓ FLEET MONITOR
    </div>
    <div style="font-family:Inter,sans-serif;font-size:.52rem;text-transform:uppercase;
                letter-spacing:.2em;color:#304060;margin-top:2px">
        Running Hours System
    </div>
    <div style="height:1px;margin:1rem 0;
                background:linear-gradient(90deg,#b8870c,transparent)"></div>
    ''', unsafe_allow_html=True)

    page = st.selectbox('Navigation',
        ['Fleet Overview', 'Vessel Detail', 'Upload Report', 'Upload History'],
        label_visibility='collapsed')

    st.markdown('<br>', unsafe_allow_html=True)
    vessels = get_vessels()
    sel_v = st.selectbox('Active Vessel', vessels) if vessels else None
    if not vessels:
        st.info('No data — upload a report to begin.')

    if vessels:
        smry = get_summary()
        if not smry.empty:
            st.markdown(
                '<div style="height:1px;margin:1rem 0;background:linear-gradient(90deg,#b8870c,transparent)"></div>'
                '<div style="font-family:Inter,sans-serif;font-size:.52rem;text-transform:uppercase;'
                'letter-spacing:.2em;color:#304060;margin-bottom:.45rem">Vessel Status</div>',
                unsafe_allow_html=True)
            for idx, (_, row) in enumerate(smry.iterrows()):
                od = int(row['overdue']); hp = int(row['high_priority'])
                bc = '#e84040' if od > 0 else ('#d06020' if hp > 0 else '#10a058')
                tags = ''
                if od > 0:
                    tags += f'<span style="font-family:JetBrains Mono,monospace;font-size:.52rem;font-weight:500;padding:1px 5px;border-radius:3px;background:rgba(184,32,32,.15);color:#ff5555">{od} OD</span>'
                if hp > 0:
                    tags += f'<span style="font-family:JetBrains Mono,monospace;font-size:.52rem;font-weight:500;padding:1px 5px;border-radius:3px;background:rgba(168,72,8,.15);color:#ffaa33;margin-left:3px">{hp} HP</span>'
                if od == 0 and hp == 0:
                    tags += '<span style="font-family:JetBrains Mono,monospace;font-size:.52rem;font-weight:500;padding:1px 5px;border-radius:3px;background:rgba(8,104,56,.15);color:#33dd77">OK</span>'
                st.markdown(
                    f'<div style="display:flex;align-items:center;justify-content:space-between;'
                    f'background:#0a1224;border:1px solid #162640;border-left:2px solid {bc};'
                    f'border-radius:6px;padding:.44rem .7rem;margin-bottom:.22rem;'
                    f'animation:push .22s ease both;animation-delay:{idx*.04}s;animation-fill-mode:both">'
                    f'<span style="font-family:Space Grotesk,sans-serif;font-size:.69rem;font-weight:600;'
                    f'color:#b0c8e8;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:110px">'
                    f'{row["vessel_name"]}</span>'
                    f'<span style="display:flex;gap:3px;align-items:center;flex-shrink:0">{tags}</span>'
                    f'</div>',
                    unsafe_allow_html=True)

    st.markdown(
        '<div style="height:1px;margin:1rem 0;background:linear-gradient(90deg,#b8870c,transparent)"></div>',
        unsafe_allow_html=True)
    db_kb = DB_PATH.stat().st_size / 1024 if DB_PATH.exists() else 0
    st.markdown(
        f'<div style="font-family:JetBrains Mono,monospace;font-size:.55rem;color:#304060">'
        f'{db_kb:.0f} kb · {len(vessels)} vessel(s) · v6.1</div>',
        unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════
#  PAGE: UPLOAD
# ═══════════════════════════════════════════════════════════════════
if page == 'Upload Report':
    st.markdown(_ph('Upload Report', 'TEC-004 · Running Hours'), unsafe_allow_html=True)

    c_up, c_info = st.columns([3, 2], gap='large')
    with c_up:
        uploaded = st.file_uploader('Drop your .doc file here',
                                    type=['doc'], label_visibility='collapsed')
    with c_info:
        st.markdown('''
        <div style="background:#0a1224;border:1px solid #162640;border-radius:12px;
                    padding:1.25rem 1.5rem;font-family:Inter,sans-serif;font-size:.8rem;
                    color:#6080a8;line-height:1.9">
            <div style="font-family:Space Grotesk,sans-serif;font-size:.56rem;font-weight:600;
                        letter-spacing:.2em;text-transform:uppercase;color:#edbb2a;margin-bottom:.35rem">
                Accepted Format</div>
            TEC-004 Running Hours Monthly Report<br>
            Any vessel &nbsp;·&nbsp; <b style="color:#b0c8e8">.doc format only</b><br><br>
            <div style="font-family:Space Grotesk,sans-serif;font-size:.56rem;font-weight:600;
                        letter-spacing:.2em;text-transform:uppercase;color:#edbb2a;margin-bottom:.35rem">
                What Gets Parsed</div>
            ✦ Vessel name &amp; report date<br>
            ✦ M/E total &amp; monthly running hours<br>
            ✦ All M/E components per cylinder (Cyl 1–7)<br>
            ✦ Aux engines (AUX-1, AUX-2, AUX-3 × 6 cyls)<br>
            ✦ Turbocharger, coolers, D/G equipment<br>
            ✦ Status auto-computed per periodicity
        </div>''', unsafe_allow_html=True)

    if uploaded:
        raw = uploaded.read()
        fh  = hashlib.md5(raw).hexdigest()

        with st.spinner('Converting .doc → .docx via LibreOffice…'):
            try:
                docx = convert_doc_to_docx(raw)
            except Exception as e:
                st.error(f'**Conversion failed:** {e}')
                st.stop()

        with st.spinner('Parsing tables…'):
            try:
                parsed = parse_doc_bytes(docx)
            except ValueError as e:
                st.error(f'**Parse failed:** {e}')
                st.stop()

        comps = parsed['components']
        nc    = len(comps)
        nod   = sum(1 for c in comps if c['status'] == 'OVERDUE')
        nhp   = sum(1 for c in comps if c['status'] == 'HIGH PRIORITY')
        nok   = sum(1 for c in comps if c['status'] == 'OK')
        nme   = sum(1 for c in comps if c['category'] == 'MAIN_ENGINE')
        naux  = sum(1 for c in comps if c['category'] == 'AUX_ENGINE')
        noe   = len(parsed['other_equipment'])

        st.markdown(_sl('Parse Summary'), unsafe_allow_html=True)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric('Vessel',         parsed['vessel_name'])
        c2.metric('Report Date',    parsed['report_date'] or '—')
        c3.metric('M/E Total Hrs',  f"{parsed['me_total_hrs']:,}"  if parsed['me_total_hrs']  else '—')
        c4.metric('M/E This Month', f"{parsed['me_this_month']:,}" if parsed['me_this_month'] else '—')
        c5.metric('Components',     nc)

        st.markdown(_ps_row(
            (nod,  'Overdue',      'red'),
            (nhp,  'High Priority','ora'),
            (nok,  'OK',           'grn'),
            (nme,  'M/E Records',  'blu'),
            (naux, 'AUX Records',  'blu'),
        ), unsafe_allow_html=True)

        for w in parsed['warnings']:
            st.warning(f'⚠ {w}')

        if nc == 0:
            st.error('No components extracted. Ensure this is a valid TEC-004 report.')
            st.stop()

        st.markdown('---')
        cb, _ = st.columns([1, 4])
        with cb:
            if st.button('CONFIRM AND SAVE', use_container_width=True):
                save_parsed(parsed, uploaded.name, fh)
                for fn in [get_vessels, get_comps, get_oe,
                           get_history, get_summary, get_all_comps]:
                    fn.clear()
                st.markdown(f'''
                <div style="background:linear-gradient(135deg,rgba(8,104,56,.13),rgba(8,104,56,.04));
                            border:1px solid rgba(8,104,56,.3);border-radius:10px;
                            padding:.9rem 1.4rem;color:#90efc0;font-family:Space Grotesk,sans-serif;
                            font-size:.88rem;font-weight:500;
                            display:flex;align-items:center;gap:.65rem">
                    <span style="font-size:1.3rem;font-weight:700">✓</span>
                    <span><strong>{parsed['vessel_name']}</strong> saved —
                    {nc} components · {nod} overdue · {nhp} high priority</span>
                </div>''', unsafe_allow_html=True)
                st.balloons()


# ═══════════════════════════════════════════════════════════════════
#  PAGE: FLEET OVERVIEW
# ═══════════════════════════════════════════════════════════════════
elif page == 'Fleet Overview':
    st.markdown(_ph('Fleet Overview', 'All vessels · Live status'), unsafe_allow_html=True)

    smry     = get_summary()
    all_comp = get_all_comps()

    if smry.empty or all_comp.empty:
        st.info('No data loaded. Upload a report and click Confirm & Save.')
        st.stop()

    tv  = len(smry)
    tc  = len(all_comp)
    tod = int((all_comp['status'] == 'OVERDUE').sum())
    thp = int((all_comp['status'] == 'HIGH PRIORITY').sum())
    tok = int((all_comp['status'] == 'OK').sum())

    k1, k2, k3, k4, k5 = st.columns(5)
    for col, (val, lbl, clr, dly) in zip(
        [k1, k2, k3, k4, k5],
        [(tv,  'Vessels',       'blu',  0.00),
         (tc,  'Components',    'gold', 0.06),
         (tod, 'Overdue',       'red',  0.12),
         (thp, 'High Priority', 'ora',  0.18),
         (tok, 'OK',            'grn',  0.24)]):
        with col:
            st.markdown(_kpi(val, lbl, clr, dly), unsafe_allow_html=True)

    st.markdown('<br>', unsafe_allow_html=True)
    st.markdown(_sl('Fleet Component Matrix'), unsafe_allow_html=True)

    # Filters
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        vf = st.selectbox('Vessel',
            ['All Fleet'] + sorted(all_comp['vessel_name'].unique().tolist()),
            key='ov_v')
    with f2:
        sf = st.selectbox('Status',
            ['All', 'Overdue only', 'High Priority +', 'OK only'],
            key='ov_s')
    with f3:
        cf = st.selectbox('Engine Type',
            ['All', 'Main Engine', 'Aux Engines'],
            key='ov_c')
    with f4:
        cpf = st.selectbox('Component',
            ['All'] + sorted(all_comp['description'].unique().tolist()),
            key='ov_cp')

    sort_opt = st.radio(
        'Sort by',
        ['Report Order', 'Component → Cylinder', 'Priority → % Used'],
        horizontal=True, key='ov_srt')
    mode_map = {
        'Report Order':          'seq',
        'Component → Cylinder':  'matrix',
        'Priority → % Used':     'priority',
    }

    # Apply filters
    filt = all_comp.copy()
    if vf != 'All Fleet':
        filt = filt[filt['vessel_name'] == vf]
    if sf == 'Overdue only':
        filt = filt[filt['status'] == 'OVERDUE']
    elif sf == 'High Priority +':
        filt = filt[filt['status'].isin(['OVERDUE', 'HIGH PRIORITY'])]
    elif sf == 'OK only':
        filt = filt[filt['status'] == 'OK']
    if cf == 'Main Engine':
        filt = filt[filt['category'] == 'MAIN_ENGINE']
    elif cf == 'Aux Engines':
        filt = filt[filt['category'] == 'AUX_ENGINE']
    if cpf != 'All':
        filt = filt[filt['description'] == cpf]

    ns = len(filt)
    no = int((filt['status'] == 'OVERDUE').sum())
    nh = int((filt['status'] == 'HIGH PRIORITY').sum())
    nk = int((filt['status'] == 'OK').sum())

    st.markdown(_fc(ns, no, nh, nk), unsafe_allow_html=True)

    if filt.empty:
        st.markdown(
            '<div style="background:rgba(8,104,56,.04);border:1px solid rgba(8,104,56,.12);'
            'border-radius:10px;padding:1.6rem;text-align:center;color:#33dd77;'
            'font-family:Space Grotesk,sans-serif;font-size:.88rem;font-weight:500">'
            'No records match the current filter</div>',
            unsafe_allow_html=True)
    else:
        render_table(
            filt,
            mode=mode_map[sort_opt],
            include_vessel=(vf == 'All Fleet'),
            height=min(860, 38 * ns + 44))


# ═══════════════════════════════════════════════════════════════════
#  PAGE: VESSEL DETAIL
# ═══════════════════════════════════════════════════════════════════
elif page == 'Vessel Detail':
    if not sel_v:
        st.info('Select a vessel from the sidebar.')
        st.stop()

    st.markdown(_ph(sel_v, 'Component Analysis'), unsafe_allow_html=True)

    df = get_comps(sel_v)
    oe = get_oe(sel_v)

    if df.empty:
        st.info('No data for this vessel. Upload and save a report first.')
        st.stop()

    n_tot = len(df)
    n_od  = int((df['status'] == 'OVERDUE').sum())
    n_hp  = int((df['status'] == 'HIGH PRIORITY').sum())
    n_ok  = int((df['status'] == 'OK').sum())
    n_nd  = int((df['status'] == 'NO DATA').sum())

    k1, k2, k3, k4, k5 = st.columns(5)
    for col, (val, lbl, clr, dly) in zip(
        [k1, k2, k3, k4, k5],
        [(n_tot, 'Total',         'gold', 0.00),
         (n_od,  'Overdue',       'red',  0.06),
         (n_hp,  'High Priority', 'ora',  0.12),
         (n_ok,  'OK',            'grn',  0.18),
         (n_nd,  'No Data',       'blu',  0.24)]):
        with col:
            st.markdown(_kpi(val, lbl, clr, dly), unsafe_allow_html=True)

    hist = get_history(sel_v)
    if not hist.empty:
        last = hist.iloc[0]
        mt = f"{int(last['me_total_hrs']):,}" if pd.notna(last['me_total_hrs']) else '—'
        mm = f"{int(last['me_this_month']):,}" if pd.notna(last['me_this_month']) else '—'
        st.markdown(
            f'<div style="display:flex;gap:1.2rem;flex-wrap:wrap;font-family:JetBrains Mono,monospace;'
            f'font-size:.61rem;color:#304060;margin:.55rem 0 0">'
            f'<span>File: <b style="color:#6080a8">{last["filename"]}</b></span>'
            f'<span>Report: <b style="color:#6080a8">{last["report_date"] or "—"}</b></span>'
            f'<span>M/E: <b style="color:#6080a8">{mt}</b> total · <b style="color:#6080a8">{mm}</b> this month</span>'
            f'<span>Saved: <b style="color:#6080a8">{str(last["uploaded_at"])[:16]}</b></span>'
            f'</div>',
            unsafe_allow_html=True)

    st.markdown('---')

    tabs = st.tabs(['Alerts', 'Main Engine', 'Aux Engines', 'Other Equipment'])

    # ── TAB 0: ALERTS ────────────────────────────────────────────────
    with tabs[0]:
        st.markdown(_sl('Overdue and High Priority — Most Critical First'),
                    unsafe_allow_html=True)
        alerts = df[df['status'].isin(['OVERDUE', 'HIGH PRIORITY'])]
        if alerts.empty:
            st.markdown(
                '<div style="background:rgba(8,104,56,.04);border:1px solid rgba(8,104,56,.12);'
                'border-radius:10px;padding:1.6rem;text-align:center;color:#33dd77;'
                'font-family:Space Grotesk,sans-serif;font-size:.88rem;font-weight:500">'
                'All components within acceptable limits</div>',
                unsafe_allow_html=True)
        else:
            no = int((alerts['status'] == 'OVERDUE').sum())
            nh = int((alerts['status'] == 'HIGH PRIORITY').sum())
            st.markdown(_fc(len(alerts), no, nh, 0), unsafe_allow_html=True)
            render_table(alerts, mode='priority')

    # ── TAB 1: MAIN ENGINE ───────────────────────────────────────────
    with tabs[1]:
        me = df[df['category'] == 'MAIN_ENGINE']
        if me.empty:
            st.info('No Main Engine data.')
        else:
            st.markdown(_sl('Main Engine — sorted by component then cylinder'),
                        unsafe_allow_html=True)
            fa, fb = st.columns(2)
            with fa:
                mc = st.selectbox(
                    'Filter component',
                    ['All'] + sorted(me['description'].unique().tolist()),
                    key='me_c')
            with fb:
                ms = st.selectbox(
                    'Filter status',
                    ['All', 'Overdue', 'High Priority +', 'OK'],
                    key='me_s')
            mr = st.radio(
                'Sort',
                ['Report Order', 'Component → Cylinder', 'Priority → % Used'],
                horizontal=True, key='me_r')

            v = me.copy()
            if mc != 'All':
                v = v[v['description'] == mc]
            if ms == 'Overdue':
                v = v[v['status'] == 'OVERDUE']
            elif ms == 'High Priority +':
                v = v[v['status'].isin(['OVERDUE', 'HIGH PRIORITY'])]
            elif ms == 'OK':
                v = v[v['status'] == 'OK']

            mmap = {'Report Order': 'seq',
                    'Component → Cylinder': 'matrix',
                    'Priority → % Used': 'priority'}
            no = int((v['status'] == 'OVERDUE').sum())
            nh = int((v['status'] == 'HIGH PRIORITY').sum())
            nk = int((v['status'] == 'OK').sum())
            st.markdown(_fc(len(v), no, nh, nk), unsafe_allow_html=True)
            render_table(v, mode=mmap[mr])

    # ── TAB 2: AUX ENGINES ───────────────────────────────────────────
    with tabs[2]:
        aux = df[df['category'] == 'AUX_ENGINE']
        if aux.empty:
            st.info('No Aux Engine data.')
        else:
            st.markdown(_sl('Auxiliary Engines — sorted by component then cylinder'),
                        unsafe_allow_html=True)
            fa, fb = st.columns(2)
            with fa:
                ae = st.selectbox(
                    'Filter engine',
                    ['All'] + sorted(aux['engine_label'].unique().tolist()),
                    key='ae')
            with fb:
                as_ = st.selectbox(
                    'Filter status',
                    ['All', 'Overdue', 'High Priority +', 'OK'],
                    key='ae_s')
            ar = st.radio(
                'Sort',
                ['Report Order', 'Component → Cylinder', 'Priority → % Used'],
                horizontal=True, key='ae_r')

            v = aux.copy()
            if ae != 'All':
                v = v[v['engine_label'] == ae]
            if as_ == 'Overdue':
                v = v[v['status'] == 'OVERDUE']
            elif as_ == 'High Priority +':
                v = v[v['status'].isin(['OVERDUE', 'HIGH PRIORITY'])]
            elif as_ == 'OK':
                v = v[v['status'] == 'OK']

            mmap = {'Report Order': 'seq',
                    'Component → Cylinder': 'matrix',
                    'Priority → % Used': 'priority'}
            no = int((v['status'] == 'OVERDUE').sum())
            nh = int((v['status'] == 'HIGH PRIORITY').sum())
            nk = int((v['status'] == 'OK').sum())
            st.markdown(_fc(len(v), no, nh, nk), unsafe_allow_html=True)
            render_table(v, mode=mmap[ar])

    # ── TAB 3: OTHER EQUIPMENT ───────────────────────────────────────
    with tabs[3]:
        if oe.empty:
            st.info('No other equipment data.')
        else:
            for sec in sorted(oe['section'].unique()):
                st.markdown(_sl(sec), unsafe_allow_html=True)
                sd = oe[oe['section'] == sec][
                    ['description', 'periodicity', 'last_date', 'run_hrs']
                ].copy()
                sd.columns = ['Description', 'Periodicity', 'Last Date', 'Run Hrs']
                st.dataframe(sd, use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════════════════════
#  PAGE: UPLOAD HISTORY
# ═══════════════════════════════════════════════════════════════════
elif page == 'Upload History':
    st.markdown(_ph('Upload History', 'Audit Trail'), unsafe_allow_html=True)
    if not sel_v:
        st.info('Select a vessel from the sidebar.')
        st.stop()
    st.markdown(_sl(sel_v), unsafe_allow_html=True)
    hist = get_history(sel_v)
    if hist.empty:
        st.info('No upload history for this vessel.')
    else:
        d = hist.copy()
        d.columns = ['Filename', 'Report Date', 'M/E Total', 'M/E Month',
                     'Components', 'Uploaded']
        d['M/E Total'] = d['M/E Total'].apply(
            lambda x: f"{int(x):,}" if pd.notna(x) else '—')
        d['M/E Month'] = d['M/E Month'].apply(
            lambda x: f"{int(x):,}" if pd.notna(x) else '—')
        st.dataframe(d, use_container_width=True, hide_index=True)
