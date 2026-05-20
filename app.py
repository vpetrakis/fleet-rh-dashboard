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
# CSS
# ═══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

:root{
  --bg:#03060d; --bg1:#06091a; --bg2:#080e20; --bg3:#0b1228; --bg4:#0f1830;
  --b1:#0f1c35; --b2:#182840; --b3:#223350;
  --gold:#c89a14; --gold2:#e0b422; --gold3:#f5cc44;
  --red:#cc2828;  --red2:#ff5c5c;  --red3:#ff9090;
  --ora:#b85518;  --ora2:#ff8833;  --ora3:#ffcc00;
  --grn:#0d8a4a;  --grn2:#22c55e; --grn3:#6ee7b7;
  --blu:#1444a8;  --blu2:#3b82f6; --blu3:#93c5fd;
  --t0:#f2f7ff; --t1:#c0d0e8; --t2:#7e95b6; --t3:#3d5472;
  --ff:'Space Grotesk',sans-serif;
  --fi:'Inter',sans-serif;
  --fm:'JetBrains Mono',monospace;
}

*,*::before,*::after{box-sizing:border-box}
html,body,[class*="css"]{
  font-family:var(--fi)!important;
  background:var(--bg)!important;
  color:var(--t1)!important;
  -webkit-font-smoothing:antialiased;
}
.main,.main>div{background:var(--bg)!important}
.block-container{max-width:100%!important;padding:1.75rem 2.25rem 5rem!important}

.main::before{
  content:'';position:fixed;inset:0;pointer-events:none;z-index:0;
  background:
    radial-gradient(ellipse 90% 50% at -10% -5%,rgba(200,154,20,.06) 0%,transparent 55%),
    radial-gradient(ellipse 70% 45% at 110% 105%,rgba(20,68,168,.05) 0%,transparent 55%);
}

/* sidebar */
[data-testid="stSidebar"]{background:var(--bg1)!important;border-right:1px solid var(--b2)!important}
[data-testid="stSidebar"] *{color:var(--t1)!important}
[data-testid="stSidebarContent"]{padding:1.25rem!important}
[data-testid="stSidebar"] .stSelectbox>div>div{
  background:var(--bg3)!important;border:1px solid var(--b2)!important;border-radius:6px!important}

/* headings */
h1{font-family:var(--ff)!important;font-size:1.8rem!important;font-weight:700!important;
   color:var(--t0)!important;letter-spacing:-.02em!important;line-height:1.2!important}
h2{font-family:var(--ff)!important;font-size:1.2rem!important;font-weight:600!important;color:var(--t0)!important}

/* metrics */
[data-testid="stMetric"]{
  background:var(--bg3)!important;border:1px solid var(--b2)!important;
  border-radius:10px!important;padding:.9rem 1.1rem 1rem!important;
  position:relative!important;overflow:hidden!important;
  transition:border-color .25s,transform .2s!important}
[data-testid="stMetric"]:hover{border-color:var(--b3)!important;transform:translateY(-2px)!important}
[data-testid="stMetricValue"]{font-family:var(--ff)!important;font-size:2rem!important;
  font-weight:700!important;color:var(--t0)!important;letter-spacing:-.03em!important}
[data-testid="stMetricLabel"]{color:var(--t3)!important;font-size:.62rem!important;
  text-transform:uppercase!important;letter-spacing:.15em!important}

/* dataframe */
[data-testid="stDataFrame"]{
  border:1px solid var(--b2)!important;border-radius:10px!important;
  overflow:hidden!important;box-shadow:0 4px 24px rgba(0,0,0,.3)!important}
.dvn-scroller{background:var(--bg2)!important}

/* button */
.stButton>button{
  background:linear-gradient(135deg,var(--gold) 0%,#8a6a08 100%)!important;
  color:#000!important;border:none!important;border-radius:8px!important;
  padding:.65rem 1.75rem!important;font-family:var(--ff)!important;
  font-weight:700!important;font-size:.82rem!important;
  letter-spacing:.06em!important;text-transform:uppercase!important;
  box-shadow:0 2px 14px rgba(200,154,20,.2)!important;transition:all .18s!important}
.stButton>button:hover{
  background:linear-gradient(135deg,var(--gold2) 0%,var(--gold) 100%)!important;
  box-shadow:0 5px 22px rgba(200,154,20,.38)!important;transform:translateY(-2px)!important}

/* uploader */
[data-testid="stFileUploadDropzone"]{
  background:linear-gradient(160deg,rgba(200,154,20,.04) 0%,rgba(20,68,168,.03) 100%)!important;
  border:1.5px dashed var(--gold)!important;border-radius:14px!important;
  padding:3rem 2rem!important;transition:all .3s!important}
[data-testid="stFileUploadDropzone"]:hover{
  background:rgba(200,154,20,.07)!important;border-color:var(--gold2)!important}
[data-testid="stFileUploadDropzone"] p,
[data-testid="stFileUploadDropzone"] span{
  color:var(--gold2)!important;font-family:var(--ff)!important;
  font-size:.95rem!important;font-weight:500!important}

/* tabs */
.stTabs [data-baseweb="tab-list"]{
  background:var(--bg2)!important;border-radius:10px 10px 0 0!important;
  border-bottom:1px solid var(--b2)!important;gap:0!important;padding:0 .75rem!important}
.stTabs [data-baseweb="tab"]{
  background:transparent!important;color:var(--t3)!important;
  font-family:var(--ff)!important;font-weight:500!important;
  text-transform:uppercase!important;letter-spacing:.04em!important;
  font-size:.74rem!important;padding:.85rem 1.2rem!important;
  border-bottom:2px solid transparent!important;margin-bottom:-1px!important}
.stTabs [data-baseweb="tab"]:hover{color:var(--t2)!important}
.stTabs [aria-selected="true"]{color:var(--gold2)!important;border-bottom:2px solid var(--gold)!important}
.stTabs [data-baseweb="tab-panel"]{
  background:var(--bg2)!important;border:1px solid var(--b2)!important;
  border-top:none!important;border-radius:0 0 10px 10px!important;padding:1.5rem!important}

/* expander */
.streamlit-expanderHeader{
  background:var(--bg3)!important;border:1px solid var(--b2)!important;
  border-radius:8px!important;font-family:var(--ff)!important;
  font-size:.85rem!important;color:var(--t1)!important;transition:all .2s!important}
.streamlit-expanderHeader:hover{background:var(--bg4)!important;border-color:var(--b3)!important}
.streamlit-expanderContent{
  background:var(--bg2)!important;border:1px solid var(--b2)!important;
  border-top:none!important;border-radius:0 0 8px 8px!important}

/* inputs */
.stSelectbox>div>div,.stMultiSelect>div>div{
  background:var(--bg3)!important;border:1px solid var(--b2)!important;
  border-radius:7px!important;color:var(--t1)!important}
.stSelectbox label,.stMultiSelect label{
  color:var(--t3)!important;font-size:.7rem!important;
  text-transform:uppercase!important;letter-spacing:.1em!important}

/* misc */
.stAlert{border-radius:8px!important;border-left-width:3px!important}
hr{border-color:var(--b2)!important;opacity:1!important;margin:1.25rem 0!important}
a{color:var(--gold2)!important;text-decoration:none!important}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:var(--bg1)}
::-webkit-scrollbar-thumb{background:var(--b3);border-radius:3px}

/* ── animations ── */
@keyframes slideD{from{opacity:0;transform:translateY(-16px)}to{opacity:1;transform:translateY(0)}}
@keyframes slideU{from{opacity:0;transform:translateY(14px)} to{opacity:1;transform:translateY(0)}}
@keyframes slideR{from{opacity:0;transform:translateX(-14px)}to{opacity:1;transform:translateX(0)}}
@keyframes numIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
@keyframes popIn{from{opacity:0;transform:scale(.9)}to{opacity:1;transform:scale(1)}}
@keyframes succ{0%{transform:scale(.85);opacity:0}55%{transform:scale(1.02)}100%{transform:scale(1);opacity:1}}
@keyframes gline{from{width:0;opacity:0}to{width:100%;opacity:1}}

/* ── page header ── */
.ph{animation:slideD .4s cubic-bezier(.22,1,.36,1) both}
.ph h1{margin:0}
.ph-eye{font-family:var(--fi);font-size:.58rem;font-weight:500;letter-spacing:.22em;
  text-transform:uppercase;color:var(--gold);margin-bottom:.25rem;
  animation:slideD .35s .05s ease both;animation-fill-mode:both}
.ph-line{height:1px;margin:.3rem 0 1.5rem;
  background:linear-gradient(90deg,var(--gold) 0%,var(--b2) 30%,transparent 100%);
  animation:gline .65s .1s ease both;animation-fill-mode:both}

/* ── KPI card ── */
.kc{background:var(--bg3);border:1px solid var(--b2);border-radius:10px;
  padding:.95rem 1.15rem 1.05rem;position:relative;overflow:hidden;
  animation:slideU .4s ease both;animation-fill-mode:both;
  transition:border-color .25s,transform .2s,box-shadow .25s;cursor:default}
.kc:hover{transform:translateY(-4px);box-shadow:0 14px 40px rgba(0,0,0,.5)}
.kc::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:10px 10px 0 0}
.kc::after{content:'';position:absolute;inset:0;border-radius:10px;pointer-events:none}
.kc.gold{border-color:rgba(200,154,20,.3)}
.kc.gold::before{background:linear-gradient(90deg,var(--gold),transparent 70%)}
.kc.gold::after{background:linear-gradient(160deg,rgba(200,154,20,.05) 0%,transparent 60%)}
.kc.gold:hover{border-color:rgba(224,180,34,.5)}
.kc.red{border-color:rgba(204,40,40,.25)}
.kc.red::before{background:linear-gradient(90deg,var(--red),transparent 70%)}
.kc.red::after{background:linear-gradient(160deg,rgba(204,40,40,.05) 0%,transparent 60%)}
.kc.red:hover{border-color:rgba(255,92,92,.4)}
.kc.orange{border-color:rgba(184,85,24,.25)}
.kc.orange::before{background:linear-gradient(90deg,var(--ora),transparent 70%)}
.kc.orange::after{background:linear-gradient(160deg,rgba(184,85,24,.05) 0%,transparent 60%)}
.kc.orange:hover{border-color:rgba(255,136,51,.4)}
.kc.green{border-color:rgba(13,138,74,.25)}
.kc.green::before{background:linear-gradient(90deg,var(--grn),transparent 70%)}
.kc.green::after{background:linear-gradient(160deg,rgba(13,138,74,.05) 0%,transparent 60%)}
.kc.green:hover{border-color:rgba(34,197,94,.4)}
.kc.blue{border-color:rgba(20,68,168,.25)}
.kc.blue::before{background:linear-gradient(90deg,var(--blu),transparent 70%)}
.kc.blue::after{background:linear-gradient(160deg,rgba(20,68,168,.05) 0%,transparent 60%)}
.kc.blue:hover{border-color:rgba(59,130,246,.4)}
.kc-val{font-family:var(--ff);font-size:2.1rem;font-weight:700;line-height:1.1;
  letter-spacing:-.04em;position:relative;z-index:1;
  animation:numIn .4s .1s ease both;animation-fill-mode:both}
.kc-lbl{font-family:var(--fi);font-size:.58rem;font-weight:500;
  text-transform:uppercase;letter-spacing:.16em;color:var(--t3);
  margin-top:4px;position:relative;z-index:1}
.kc.gold  .kc-val{color:var(--gold3)}
.kc.red   .kc-val{color:var(--red2)}
.kc.orange.kc-val{color:var(--ora2)}
.kc.orange .kc-val{color:var(--ora2)}
.kc.green .kc-val{color:var(--grn2)}
.kc.blue  .kc-val{color:var(--blu2)}

/* ── section label ── */
.sl{font-family:var(--fi);font-size:.57rem;font-weight:600;
  letter-spacing:.22em;text-transform:uppercase;color:var(--t3);
  display:flex;align-items:center;gap:.7rem;margin:1.5rem 0 .9rem}
.sl::after{content:'';flex:1;height:1px;background:linear-gradient(90deg,var(--b2),transparent)}

/* ── parse stats ── */
.ps-row{display:flex;gap:.6rem;flex-wrap:wrap;margin:.9rem 0 1.4rem}
.ps{background:var(--bg3);border:1px solid var(--b2);border-radius:9px;
  padding:.6rem 1rem .65rem;min-width:82px;
  animation:popIn .3s ease both;animation-fill-mode:both;
  transition:border-color .2s,transform .15s;position:relative;overflow:hidden}
.ps:hover{transform:translateY(-2px);border-color:var(--b3)}
.ps::before{content:'';position:absolute;top:0;left:0;right:0;height:2px}
.ps.red::before{background:var(--red)} .ps.orange::before{background:var(--ora)}
.ps.green::before{background:var(--grn)} .ps.blue::before{background:var(--blu)}
.ps-val{font-family:var(--ff);font-size:1.65rem;font-weight:700;line-height:1;letter-spacing:-.03em}
.ps-lbl{font-family:var(--fi);font-size:.57rem;text-transform:uppercase;letter-spacing:.14em;color:var(--t3);margin-top:3px}
.ps.red    .ps-val{color:var(--red2)}
.ps.orange .ps-val{color:var(--ora2)}
.ps.green  .ps-val{color:var(--grn2)}
.ps.blue   .ps-val{color:var(--blu2)}

/* ── info card ── */
.ic{background:var(--bg3);border:1px solid var(--b2);border-radius:12px;
  padding:1.3rem 1.5rem;font-family:var(--fi);font-size:.82rem;
  color:var(--t2);line-height:1.9;animation:slideR .4s ease both;animation-fill-mode:both}
.ic-title{font-family:var(--ff);font-size:.57rem;font-weight:600;
  letter-spacing:.2em;text-transform:uppercase;color:var(--gold2);margin-bottom:.35rem}

/* ── success / all-clear ── */
.sb{background:linear-gradient(135deg,rgba(13,138,74,.12),rgba(13,138,74,.04));
  border:1px solid rgba(13,138,74,.3);border-radius:10px;
  padding:1rem 1.5rem;color:var(--grn3);font-family:var(--ff);
  font-size:.9rem;font-weight:500;animation:succ .5s cubic-bezier(.34,1.56,.64,1) both;
  display:flex;align-items:center;gap:.7rem}
.ac{background:rgba(13,138,74,.04);border:1px solid rgba(13,138,74,.12);
  border-radius:10px;padding:1.75rem;text-align:center;
  color:var(--grn3);font-family:var(--ff);font-size:.9rem;font-weight:500}

/* ── sidebar widgets ── */
.logo{font-family:var(--ff);font-size:1.1rem;font-weight:700;
  letter-spacing:.04em;color:var(--gold2);display:flex;align-items:center;gap:.45rem}
.logo-tag{font-family:var(--fi);font-size:.55rem;text-transform:uppercase;
  letter-spacing:.2em;color:var(--t3);margin-top:3px}
.logo-rule{height:1px;margin:1.1rem 0;background:linear-gradient(90deg,var(--gold),transparent)}
.vc{display:flex;align-items:center;justify-content:space-between;
  background:var(--bg3);border:1px solid var(--b2);border-radius:7px;
  padding:.48rem .75rem;margin-bottom:.25rem;
  animation:slideR .25s ease both;animation-fill-mode:both;transition:all .2s}
.vc:hover{border-color:var(--b3);background:var(--bg4)}
.vc.crit{border-left:2px solid var(--red)}
.vc.warn{border-left:2px solid var(--ora)}
.vc.safe{border-left:2px solid var(--grn)}
.vc-name{font-family:var(--ff);font-size:.71rem;font-weight:600;color:var(--t1);
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:112px}
.vc-tags{display:flex;gap:4px;align-items:center;flex-shrink:0}
.vt{font-family:var(--fm);font-size:.54rem;font-weight:500;padding:1px 5px;border-radius:3px}
.vt.od{background:rgba(204,40,40,.15);color:var(--red2)}
.vt.hp{background:rgba(184,85,24,.15);color:var(--ora2)}
.vt.ok{background:rgba(13,138,74,.15);color:var(--grn2)}

/* ── vessel hero ── */
.vh{background:var(--bg3);border:1px solid var(--b2);border-radius:12px;
  padding:1.1rem 1.5rem;display:flex;align-items:center;
  justify-content:space-between;flex-wrap:wrap;gap:.9rem;
  margin:.4rem 0 1.4rem;animation:slideD .38s ease both}
.vh.crit{border-left:4px solid var(--red)}
.vh.warn{border-left:4px solid var(--ora)}
.vh.safe{border-left:4px solid var(--grn)}
.vh-name{font-family:var(--ff);font-size:1.2rem;font-weight:700;color:var(--t0)}
.vh-meta{font-family:var(--fm);font-size:.63rem;color:var(--t3);margin-top:3px}
.vh-stats{display:flex;gap:.8rem;align-items:center;flex-wrap:wrap}
.vh-s{text-align:center}
.vh-sv{font-family:var(--ff);font-size:1.45rem;font-weight:800;line-height:1}
.vh-sl{font-family:var(--fi);font-size:.56rem;text-transform:uppercase;
  letter-spacing:.14em;color:var(--t3);margin-top:2px}
.sev{border-radius:5px;padding:4px 12px;font-family:var(--fm);
  font-size:.65rem;font-weight:700;letter-spacing:.07em}

/* ── filter count ── */
.fc{font-family:var(--fm);font-size:.64rem;color:var(--t3);margin-bottom:.65rem}
.fc b{color:var(--t1)} .fc .od{color:var(--red2);font-weight:700}
.fc .hp{color:var(--ora2);font-weight:700} .fc .ok{color:var(--grn2);font-weight:700}

/* ── meta line ── */
.ml{display:flex;gap:1.2rem;flex-wrap:wrap;
  font-family:var(--fm);font-size:.63rem;color:var(--t3);margin:.6rem 0 0}
.ml b{color:var(--t2);font-weight:500}
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════
# DATABASE
# ═══════════════════════════════════════════════════════════════════
DB_PATH = Path("running_hours.db")

def get_db():
    c = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    c.row_factory = sqlite3.Row
    return c

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
    CREATE INDEX IF NOT EXISTS idx_cv ON components(vessel_name);
    CREATE INDEX IF NOT EXISTS idx_cs ON components(status);
    CREATE INDEX IF NOT EXISTS idx_cseq ON components(vessel_name,seq);
    """)
    c.commit(); c.close()

init_db()


# ═══════════════════════════════════════════════════════════════════
# CONVERSION
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
# PARSER — uses _tc identity to handle merged cells correctly
# ═══════════════════════════════════════════════════════════════════

def _unique_cells(row):
    """Return (col_index, cell) for the FIRST occurrence of each physical cell.
    Merged cells share the same _tc XML element — deduplicating by id(_tc)
    prevents double-counting CYL columns and REMARKS columns."""
    seen, out = set(), []
    for ci, cell in enumerate(row.cells):
        cid = id(cell._tc)
        if cid not in seen:
            seen.add(cid); out.append((ci, cell))
    return out

def _clean(x) -> str:
    return re.sub(r"\s+", " ", str(x or "").replace("\xa0", " ")).strip()

def _parse_date(raw: str) -> Optional[str]:
    raw = _clean(raw)
    if not raw or raw.upper() in ("N/A","NA","-","--"): return None
    if re.match(r"^\d+$", raw): return None  # pure number = hours, not date
    rn = re.sub(r"\bSEPT\b","SEP", raw, flags=re.I)
    rn = re.sub(r"\bJUNE\b","JUN", rn,  flags=re.I)
    rn = re.sub(r"\bJULY\b","JUL", rn,  flags=re.I)
    fmts = ["%d %b %y","%d %B %y","%d %b %Y","%d %B %Y",
            "%d/%m/%y","%d/%m/%Y","%d-%m-%y","%d-%m-%Y",
            "%b %Y","%B %Y","%Y-%m-%d"]
    for fmt in fmts:
        for v in (rn, rn.upper(), rn.title(), raw, raw.upper()):
            try: return datetime.strptime(v, fmt).strftime("%Y-%m-%d")
            except ValueError: pass
    return None  # reject unparseable strings (like remarks text)

def _parse_hrs(raw: str) -> Optional[float]:
    """Extract hours — only from strings that start with a digit."""
    s = _clean(raw)
    if not s or not re.match(r"^\d", s): return None
    # Take first number (handles "17560\n95" → 17560)
    m = re.search(r"\d+", s.replace(",", ""))
    return float(m.group()) if m else None

def _clean_period(raw: str) -> Optional[float]:
    s = _clean(raw)
    # Remove thousand-separator dots: "16.000" → "16000"
    s = re.sub(r"\.(?=\d{3}(\.|$))", "", s)
    s = re.sub(r"[^0-9.]", "", s)
    try: return float(s) if s else None
    except ValueError: return None

def _status(h, p) -> str:
    if h is None or p is None or p == 0: return "NO DATA"
    r = h / p
    if r > 1.0: return "OVERDUE"
    if r >= 0.80: return "HIGH PRIORITY"
    return "OK"

def _pct(h, p) -> float:
    if h is None or p is None or p == 0: return 0.0
    return round(h / p, 4)

def _is_component_name(txt: str) -> bool:
    """True if the cell looks like a component description."""
    t = txt.strip()
    if not t or len(t) < 2: return False
    # Reject header patterns
    if re.match(r"^1[\-\s]?DATE OF", t, re.I): return False
    if re.match(r"^(PERIODICITY|CYL|MAIN ENGINE|TYPE:|TOTAL RUNNING|THIS MONTH|REMARKS|DATE)$", t, re.I): return False
    # Must contain at least one letter
    if not re.search(r"[A-Za-z]", t): return False
    return True


def parse_doc_bytes(docx_bytes: bytes) -> dict:
    from docx import Document

    warns = []
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as t:
        t.write(docx_bytes); tp = t.name
    try:    doc = Document(tp)
    except Exception as e: raise ValueError(f"Cannot open document: {e}")
    finally:
        try: os.unlink(tp)
        except Exception: pass

    if not doc.tables:
        raise ValueError("No tables found — is this a TEC-004 report?")

    # ── Vessel name + report date ─────────────────────────────────
    vessel_name = "UNKNOWN"; report_date = None
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt: continue
        vm = re.search(r"Vessel[\u2019\u2018']?s?\s+Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+?)(?:\s{2,}|\t)", txt, re.IGNORECASE)
        dm = re.search(r"Date\s*:\s*(.+)", txt, re.IGNORECASE)
        if vm: vessel_name = vm.group(1).strip()
        if dm: report_date = _parse_date(dm.group(1).strip())
        if vm or dm: break
    if vessel_name == "UNKNOWN":
        warns.append("Could not extract vessel name from header.")

    # ── Table 0: Main Engine ──────────────────────────────────────
    me_total = me_month = None
    components = []; seq = 0
    t0 = doc.tables[0]

    # Row 0: ME totals (unique cells only)
    for _, cell in _unique_cells(t0.rows[0]):
        txt = cell.text
        if m := re.search(r"Total Running Hours[\s:ǀ|]+([\d,]+)", txt, re.IGNORECASE):
            try: me_total = int(m.group(1).replace(",",""))
            except ValueError: pass
        if m := re.search(r"This Month[\s:]+([\d,]+)", txt, re.IGNORECASE):
            try: me_month = int(m.group(1).replace(",",""))
            except ValueError: pass

    # Row 1: cylinder column detection (unique cells → no double-counting)
    cyl_cols = []
    if len(t0.rows) > 1:
        for ci, cell in _unique_cells(t0.rows[1]):
            if m := re.search(r"CYL\s*\.?\s*No\s*\.?\s*(\d+)", cell.text.strip(), re.IGNORECASE):
                cyl_cols.append((ci, f"Cyl {int(m.group(1))}"))

    # Rows 2+: component pairs
    # Row type 1 = dates, Row type 2 = hours; col[2] holds the type indicator
    rows = t0.rows; i = 2
    while i < len(rows) - 1:
        # Build col→text maps using UNIQUE cells only
        r1 = {ci: cell.text.strip() for ci, cell in _unique_cells(rows[i])}
        r2 = {ci: cell.text.strip() for ci, cell in _unique_cells(rows[i+1])} if i+1 < len(rows) else {}

        name = r1.get(0, "")
        if not _is_component_name(name): i += 1; continue

        type1 = r1.get(2, ""); type2 = r2.get(2, ""); name2 = r2.get(0, "")

        if type1 == "1" and type2 == "2" and name == name2:
            period = _clean_period(r1.get(1, ""))
            for ci, lbl in cyl_cols:
                d = _parse_date(r1.get(ci, ""))
                h = _parse_hrs(r2.get(ci, ""))
                if d is None and h is None: continue
                components.append({
                    "seq": seq, "category": "MAIN_ENGINE", "engine_label": "ME",
                    "unit": lbl, "description": name, "periodicity": period,
                    "last_oh_date": d, "last_oh_hrs": h, "hrs_since": h,
                    "pct_used": _pct(h, period), "status": _status(h, period),
                })
                seq += 1
            i += 2
        else:
            i += 1

    # ── Table 1: Turbocharger / Coolers / A/C ─────────────────────
    other_equip = []
    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        SKIP = {"TURBOCHARGER","AUXILIARY BOILER","COOLERS","EXH GAS  BOILER",
                "A/C & REFR. COMPRESSORS","MAIN AIR COMPRESSORS",
                "PERIODICTLY","DATE OF LAST INSPECTION","RUN HRS",
                "DATE OF LAST CLEANING","DATE","PERIODICITY",""}
        for row in t1.rows:
            cells = {ci: cell.text.strip() for ci, cell in _unique_cells(row)}
            for sec, dc, datec, hrsc in [
                ("TURBOCHARGER / AUX BOILER", 0, 1, 3),
                ("COOLERS / EXH GAS BOILER",  5, 6, 8),
                ("A/C & COMPRESSORS",         10,11,12),
            ]:
                desc = cells.get(dc, "")
                if not desc or desc.upper() in SKIP: continue
                dv = cells.get(datec, ""); hv = cells.get(hrsc, "")
                if dv or hv:
                    other_equip.append({"section": sec, "description": desc,
                                        "periodicity": "", "last_date": dv, "run_hrs": hv})

    # ── Table 2: Auxiliary Engines ────────────────────────────────
    if len(doc.tables) > 2:
        t2 = doc.tables[2]; rows2 = t2.rows; eblocks = []
        if rows2:
            hdr = {ci: cell.text.strip() for ci, cell in _unique_cells(rows2[0])}
            tot = {ci: cell.text.strip() for ci, cell in _unique_cells(rows2[2])} if len(rows2) > 2 else {}
            seen_e = set()
            for ci, txt in hdr.items():
                if m := re.search(r"Aux\.\s*Engine\s*No\.?\s*(\d+)", txt, re.IGNORECASE):
                    lbl = f"AUX-{int(m.group(1))}"
                    if lbl not in seen_e:
                        seen_e.add(lbl); eblocks.append((lbl, ci))

        # Cylinder column map from row 4
        cyl_map = {}
        if len(rows2) > 4:
            r4 = {ci: cell.text.strip() for ci, cell in _unique_cells(rows2[4])}
            for ei, (elbl, es) in enumerate(eblocks):
                ee = eblocks[ei+1][1] if ei+1 < len(eblocks) else max(r4.keys(), default=0)+1
                seen_c: list = []
                for ci, txt in r4.items():
                    if es <= ci < ee:
                        if m := re.search(r"(\d+)", txt):
                            cn = int(m.group(1))
                            if cn not in seen_c:
                                seen_c.append(cn); cyl_map[ci] = (elbl, f"Cyl {cn}")

        i2 = 5
        while i2 < len(rows2) - 1:
            r1 = {ci: cell.text.strip() for ci, cell in _unique_cells(rows2[i2])}
            r2 = {ci: cell.text.strip() for ci, cell in _unique_cells(rows2[i2+1])} if i2+1 < len(rows2) else {}
            name = r1.get(0, "")
            if not _is_component_name(name): i2 += 1; continue
            type1 = r1.get(2, ""); type2 = r2.get(2, ""); name2 = r2.get(0, "")
            if type1 in ("1","2") and name == name2:
                period = _clean_period(r1.get(1, ""))
                for ci, (elbl, ulbl) in cyl_map.items():
                    d = _parse_date(r1.get(ci, ""))
                    h = _parse_hrs(r2.get(ci, ""))
                    if d is None and h is None: continue
                    components.append({
                        "seq": seq, "category": "AUX_ENGINE", "engine_label": elbl,
                        "unit": ulbl, "description": name, "periodicity": period,
                        "last_oh_date": d, "last_oh_hrs": h, "hrs_since": h,
                        "pct_used": _pct(h, period), "status": _status(h, period),
                    })
                    seq += 1
                i2 += 2
            else: i2 += 1

    # ── Table 3: D/G Equipment ────────────────────────────────────
    if len(doc.tables) > 3:
        t3 = doc.tables[3]; dglbls = ["D/G 1","D/G 2","D/G 3"]
        for ri, row in enumerate(t3.rows):
            cells = {ci: cell.text.strip() for ci, cell in _unique_cells(row)}
            if ri == 0: continue
            for dc, pc, tc, ds in [(0,1,2,3),(9,10,11,12)]:
                desc = cells.get(dc,""); per = cells.get(pc,""); rt = cells.get(tc,"")
                if not desc or rt not in ("1","2"): continue
                for gi, gl in enumerate(dglbls):
                    val = cells.get(ds+gi, "")
                    if not val: continue
                    key = f"{desc} — {gl}"
                    if rt == "1":
                        other_equip.append({"section":"D/G EQUIPMENT","description":key,
                                            "periodicity":per,"last_date":_parse_date(val) or val,"run_hrs":""})
                    else:
                        for e in reversed(other_equip):
                            if e["description"]==key and e["run_hrs"]=="":
                                e["run_hrs"]=val; break
                        else:
                            other_equip.append({"section":"D/G EQUIPMENT","description":key,
                                                "periodicity":per,"last_date":"","run_hrs":val})

    return {
        "vessel_name": vessel_name, "report_date": report_date,
        "me_total_hrs": me_total, "me_this_month": me_month,
        "components": components, "other_equipment": other_equip, "warnings": warns,
    }


# ═══════════════════════════════════════════════════════════════════
# DB HELPERS
# ═══════════════════════════════════════════════════════════════════
def save_parsed(parsed, filename, fhash):
    conn = get_db(); c = conn.cursor()
    now = datetime.utcnow().isoformat()+"Z"; v = parsed["vessel_name"]
    try:
        c.execute("INSERT OR IGNORE INTO vessels(name,created_at) VALUES(?,?)", (v,now))
        c.execute("INSERT INTO upload_log(vessel_name,filename,file_hash,report_date,me_total_hrs,me_this_month,component_count,uploaded_at) VALUES(?,?,?,?,?,?,?,?)",
            (v,filename,fhash,parsed["report_date"],parsed["me_total_hrs"],parsed["me_this_month"],len(parsed["components"]),now))
        c.execute("DELETE FROM components WHERE vessel_name=?", (v,))
        for x in parsed["components"]:
            c.execute("INSERT INTO components(vessel_name,category,engine_label,unit,description,periodicity,last_oh_date,last_oh_hrs,hrs_since,pct_used,status,seq,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (v,x["category"],x["engine_label"],x["unit"],x["description"],
                 x["periodicity"],x["last_oh_date"],x["last_oh_hrs"],
                 x["hrs_since"],x["pct_used"],x["status"],x["seq"],now))
        c.execute("DELETE FROM other_equipment WHERE vessel_name=?", (v,))
        for x in parsed["other_equipment"]:
            c.execute("INSERT INTO other_equipment(vessel_name,section,description,periodicity,last_date,run_hrs,updated_at) VALUES(?,?,?,?,?,?,?)",
                (v,x["section"],x["description"],x.get("periodicity",""),x.get("last_date",""),x.get("run_hrs",""),now))
        conn.commit()
    except Exception:
        conn.rollback(); raise
    finally:
        conn.close()

@st.cache_data(ttl=10)
def get_vessels():
    c=get_db(); r=c.execute("SELECT name FROM vessels ORDER BY name").fetchall(); c.close()
    return [x["name"] for x in r]

@st.cache_data(ttl=10)
def get_comps(vessel):
    c=get_db()
    df=pd.read_sql_query("SELECT * FROM components WHERE vessel_name=? ORDER BY seq", c, params=(vessel,))
    c.close(); return df

@st.cache_data(ttl=10)
def get_oe(vessel):
    c=get_db()
    df=pd.read_sql_query("SELECT * FROM other_equipment WHERE vessel_name=? ORDER BY section,description", c, params=(vessel,))
    c.close(); return df

@st.cache_data(ttl=10)
def get_history(vessel):
    c=get_db()
    df=pd.read_sql_query("SELECT filename,report_date,me_total_hrs,me_this_month,component_count,uploaded_at FROM upload_log WHERE vessel_name=? ORDER BY uploaded_at DESC LIMIT 20", c, params=(vessel,))
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
            MAX(u.me_total_hrs) AS me_total_hrs,
            MAX(u.report_date) AS report_date
        FROM components c
        LEFT JOIN upload_log u ON u.vessel_name=c.vessel_name
        GROUP BY c.vessel_name ORDER BY overdue DESC,high_priority DESC
    """, c); c.close(); return df

@st.cache_data(ttl=10)
def get_all_comps():
    c=get_db(); df=pd.read_sql_query("SELECT * FROM components ORDER BY vessel_name,seq", c); c.close(); return df


# ═══════════════════════════════════════════════════════════════════
# TABLE ENGINE
# ═══════════════════════════════════════════════════════════════════

# Status theme — single source of truth
_TH = {
    "OVERDUE":       {"bg":"#1f0505","bgs":"#2d0707","ts":"#ff6b6b","tm":"#ff8080","tn":"#ff3333","td":"#6b3333"},
    "HIGH PRIORITY": {"bg":"#1e0d02","bgs":"#2d1503","ts":"#ffaa44","tm":"#ff9933","tn":"#ffcc00","td":"#6b4422"},
    "OK":            {"bg":"#021208","bgs":"#042010","ts":"#4ade80","tm":"#22c55e","tn":"#4ade80","td":"#0f4023"},
    "NO DATA":       {"bg":"#090e18","bgs":"#0c1422","ts":"#5580aa","tm":"#3d5878","tn":"#5580aa","td":"#1a2a38"},
}

_STATUS_ORD = {"OVERDUE":0,"HIGH PRIORITY":1,"OK":2,"NO DATA":3}

def _sf(x) -> Optional[float]:
    try:
        v = float(x); return None if pd.isna(v) else v
    except: return None

def _cyl(u) -> int:
    m = re.search(r"\d+", str(u)); return int(m.group()) if m else 999


def build_table(df: pd.DataFrame, mode: str = "seq") -> pd.DataFrame:
    """
    Build the display DataFrame.

    mode='seq'      → original document order (seq column)
    mode='matrix'   → sorted component A-Z then cylinder 1-N
    mode='priority' → OVERDUE first, then HIGH PRIORITY, within each by % used desc
    """
    if df.empty:
        return pd.DataFrame(columns=["Status","Vessel","Component","Engine",
                                      "Unit","Periodicity","Last O/H","Hrs Since","Used"])
    d = df.copy()

    # Ensure required columns exist
    for col in ["vessel_name","category","seq"]:
        if col not in d.columns:
            d[col] = "" if col != "seq" else range(len(d))

    if mode == "seq":
        d = d.sort_values(["vessel_name","seq"], ascending=[True,True])
    elif mode == "matrix":
        d["_k1"] = d["description"].str.upper()
        d["_k2"] = d["unit"].apply(_cyl)
        d = d.sort_values(["_k1","_k2"]).drop(columns=["_k1","_k2"])
    elif mode == "priority":
        d["_s"] = d["status"].map(lambda x: _STATUS_ORD.get(str(x),4))
        d["_p"] = d["pct_used"].apply(lambda x: _sf(x) or 0.0)
        d = d.sort_values(["_s","_p"], ascending=[True,False]).drop(columns=["_s","_p"])

    out = pd.DataFrame(index=range(len(d)))
    out["Status"]      = d["status"].values
    out["Vessel"]      = d["vessel_name"].values if "vessel_name" in d.columns else ""
    out["Component"]   = d["description"].values
    out["Engine"]      = d["engine_label"].values
    out["Unit"]        = d["unit"].values
    out["Periodicity"] = [int(float(x)) if _sf(x) else None for x in d["periodicity"].values]
    out["Last O/H"]    = [str(x) if x and str(x) not in ("nan","None","") else "—"
                          for x in d["last_oh_date"].values]
    out["Hrs Since"]   = [int(float(x)) if _sf(x) else None for x in d["hrs_since"].values]
    out["Used"]        = [round(float(x)*100,1) if _sf(x) else 0.0 for x in d["pct_used"].values]
    return out


def style_table(df: pd.DataFrame):
    def rs(row):
        c = _TH.get(str(row.get("Status","")), _TH["NO DATA"])
        return [
            f"background-color:{c['bgs']};color:{c['ts']};font-weight:700",   # Status
            f"background-color:{c['bg']};color:{c['tm']};font-weight:600",    # Vessel
            f"background-color:{c['bg']};color:{c['tm']};font-weight:600",    # Component
            f"background-color:{c['bg']};color:{c['td']}",                    # Engine
            f"background-color:{c['bg']};color:{c['td']}",                    # Unit
            f"background-color:{c['bg']};color:{c['td']}",                    # Periodicity
            f"background-color:{c['bg']};color:{c['td']}",                    # Last O/H
            f"background-color:{c['bg']};color:{c['tm']};font-weight:600",    # Hrs Since
            f"background-color:{c['bg']};color:{c['tn']};font-weight:700",    # Used
        ]
    return df.style.apply(rs, axis=1)


_COLCFG = {
    "Status":      st.column_config.TextColumn("Status",      width=130),
    "Vessel":      st.column_config.TextColumn("Vessel",      width=120),
    "Component":   st.column_config.TextColumn("Component",   width=210),
    "Engine":      st.column_config.TextColumn("Engine",      width=80),
    "Unit":        st.column_config.TextColumn("Unit",        width=68),
    "Periodicity": st.column_config.NumberColumn("Periodicity",format="%d",        width=100),
    "Last O/H":    st.column_config.TextColumn("Last O/H",    width=100),
    "Hrs Since":   st.column_config.NumberColumn("Hrs Since",  format="%d hrs",    width=100),
    "Used":        st.column_config.ProgressColumn("Used", min_value=0, max_value=160,
                                                    format="%.1f%%", width=120),
}

# Column config without the Vessel column (for single-vessel views)
_COLCFG_NV = {k:v for k,v in _COLCFG.items() if k != "Vessel"}


def render_table(df, height=None, mode="seq", show_vessel=False):
    if isinstance(df, list): df = pd.DataFrame(df)
    if df.empty: st.info("No data to display."); return
    tbl = build_table(df, mode=mode)
    if not show_vessel and "Vessel" in tbl.columns:
        tbl = tbl.drop(columns=["Vessel"])
        cfg = _COLCFG_NV
    else:
        cfg = _COLCFG
    h = height or min(900, 38*len(tbl)+44)
    st.dataframe(style_table(tbl), use_container_width=True, hide_index=True, height=h, column_config=cfg)


# ═══════════════════════════════════════════════════════════════════
# UI HELPERS
# ═══════════════════════════════════════════════════════════════════
def kpi(val, lbl, color="gold", delay=0):
    return (f'<div class="kc {color}" style="animation-delay:{delay}s">'
            f'<div class="kc-val">{val}</div><div class="kc-lbl">{lbl}</div></div>')

def ph(icon, title, eye=""):
    e = f'<div class="ph-eye">{eye}</div>' if eye else ""
    return f'<div class="ph">{e}<h1>{icon}&nbsp;&nbsp;{title}</h1></div><div class="ph-line"></div>'

def sl(txt):
    return f'<div class="sl">{txt}</div>'

def fc(total, od, hp, ok):
    return (f'<div class="fc">Showing <b>{total}</b> records — '
            f'<span class="od">{od} overdue</span> · '
            f'<span class="hp">{hp} high priority</span> · '
            f'<span class="ok">{ok} OK</span></div>')


# ═══════════════════════════════════════════════════════════════════
# SIDEBAR
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
            st.markdown('<div style="font-family:var(--fi);font-size:.54rem;text-transform:uppercase;'
                        'letter-spacing:.2em;color:var(--t3);margin-bottom:.45rem">Vessel Status</div>',
                        unsafe_allow_html=True)
            for idx, (_, row) in enumerate(smry.iterrows()):
                od=int(row["overdue"]); hp=int(row["high_priority"])
                cls = "crit" if od>0 else ("warn" if hp>0 else "safe")
                tags = ""
                if od>0: tags += f'<span class="vt od">{od} OD</span>'
                if hp>0: tags += f'<span class="vt hp">{hp} HP</span>'
                if od==0 and hp==0: tags += '<span class="vt ok">✓ OK</span>'
                st.markdown(
                    f'<div class="vc {cls}" style="animation-delay:{idx*.04}s">'
                    f'<div class="vc-name">{row["vessel_name"]}</div>'
                    f'<div class="vc-tags">{tags}</div></div>',
                    unsafe_allow_html=True)

    st.markdown('<div class="logo-rule"></div>', unsafe_allow_html=True)
    db_kb = DB_PATH.stat().st_size/1024 if DB_PATH.exists() else 0
    st.markdown(f'<div style="font-family:var(--fm);font-size:.56rem;color:var(--t3)">'
                f'db {db_kb:.0f} kb · {len(vessels)} vessels · v5.1</div>',
                unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════
# PAGE: UPLOAD
# ═══════════════════════════════════════════════════════════════════
if page == "📤  Upload Report":
    st.markdown(ph("📤","Upload Report","TEC-004 · Running Hours"), unsafe_allow_html=True)

    col_up, col_info = st.columns([3,2], gap="large")
    with col_up:
        uploaded = st.file_uploader("file", type=["doc"], label_visibility="collapsed")
    with col_info:
        st.markdown("""
        <div class="ic">
          <div class="ic-title">Accepted Format</div>
          TEC-004 Running Hours Monthly Report<br>
          Any vessel &nbsp;·&nbsp; <b>.doc format only</b><br><br>
          <div class="ic-title">What Gets Parsed</div>
          ✦ Vessel name &amp; report date<br>
          ✦ M/E total &amp; monthly running hours<br>
          ✦ All M/E components per cylinder<br>
          ✦ Aux engines (up to 3 × 7 cylinders)<br>
          ✦ Turbocharger, coolers, D/G equipment<br>
          ✦ Status auto-computed per periodicity
        </div>""", unsafe_allow_html=True)

    if uploaded:
        raw = uploaded.read()
        fh  = hashlib.md5(raw).hexdigest()

        with st.spinner("Converting .doc → .docx and parsing…"):
            try:   docx = convert_doc_to_docx(raw)
            except Exception as e: st.error(f"Conversion failed: `{e}`"); st.stop()
            try:   parsed = parse_doc_bytes(docx)
            except ValueError as e: st.error(f"Parse failed: `{e}`"); st.stop()

        comps = parsed["components"]
        nc    = len(comps)
        nod   = sum(1 for c in comps if c["status"]=="OVERDUE")
        nhp   = sum(1 for c in comps if c["status"]=="HIGH PRIORITY")
        nok   = sum(1 for c in comps if c["status"]=="OK")
        noe   = len(parsed["other_equipment"])

        st.markdown(sl("Parsed Data Summary"), unsafe_allow_html=True)
        c1,c2,c3,c4,c5 = st.columns(5)
        c1.metric("Vessel",          parsed["vessel_name"])
        c2.metric("Report Date",     parsed["report_date"] or "—")
        c3.metric("M/E Total Hrs",   f"{parsed['me_total_hrs']:,}"  if parsed["me_total_hrs"]  else "—")
        c4.metric("M/E This Month",  f"{parsed['me_this_month']:,}" if parsed["me_this_month"] else "—")
        c5.metric("Components",      nc)

        st.markdown(f"""
        <div class="ps-row">
          <div class="ps red"    style="animation-delay:0s">
            <div class="ps-val">{nod}</div><div class="ps-lbl">Overdue</div></div>
          <div class="ps orange" style="animation-delay:.07s">
            <div class="ps-val">{nhp}</div><div class="ps-lbl">High Priority</div></div>
          <div class="ps green"  style="animation-delay:.14s">
            <div class="ps-val">{nok}</div><div class="ps-lbl">OK</div></div>
          <div class="ps blue"   style="animation-delay:.21s">
            <div class="ps-val">{noe}</div><div class="ps-lbl">Other Equip</div></div>
        </div>""", unsafe_allow_html=True)

        for w in parsed["warnings"]: st.warning(f"⚠ {w}")

        if nc == 0:
            st.error("No components were extracted. The document may not be a valid TEC-004 report.")
            st.stop()

        st.markdown("---")
        cb, _ = st.columns([1,4])
        with cb:
            if st.button("✅  CONFIRM & SAVE", use_container_width=True):
                save_parsed(parsed, uploaded.name, fh)
                for fn in [get_vessels,get_comps,get_oe,get_history,get_summary,get_all_comps]:
                    fn.clear()
                st.markdown(f"""
                <div class="sb"><span style="font-size:1.35rem">✓</span>
                  <span><b>{parsed['vessel_name']}</b> saved —
                  {nc} components · {nod} overdue · {nhp} high priority</span>
                </div>""", unsafe_allow_html=True)
                st.balloons()


# ═══════════════════════════════════════════════════════════════════
# PAGE: FLEET OVERVIEW — holistic matrix
# ═══════════════════════════════════════════════════════════════════
elif page == "🗺️  Fleet Overview":
    st.markdown(ph("🗺️","Fleet Overview","All vessels · Live status"), unsafe_allow_html=True)

    smry     = get_summary()
    all_comp = get_all_comps()

    if smry.empty or all_comp.empty:
        st.info("No data loaded. Upload a report to begin."); st.stop()

    # ── Fleet KPIs ──────────────────────────────────────────────────
    tv=len(smry); tc=int(all_comp.shape[0])
    tod=int((all_comp["status"]=="OVERDUE").sum())
    thp=int((all_comp["status"]=="HIGH PRIORITY").sum())
    tok=int((all_comp["status"]=="OK").sum())

    k1,k2,k3,k4,k5 = st.columns(5)
    for col,(val,lbl,clr,dly) in zip([k1,k2,k3,k4,k5],[
        (tv,"Vessels","blue",0),(tc,"Components","gold",.07),
        (tod,"Overdue","red",.14),(thp,"High Priority","orange",.21),(tok,"OK","green",.28)]):
        with col: st.markdown(kpi(val,lbl,clr,dly), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(sl("Fleet Component Matrix"), unsafe_allow_html=True)

    # ── Filters ─────────────────────────────────────────────────────
    f1,f2,f3,f4 = st.columns(4)
    with f1:
        vnames = ["All Fleet"] + sorted(all_comp["vessel_name"].unique().tolist())
        vf = st.selectbox("Vessel", vnames, key="ov_v")
    with f2:
        sf = st.selectbox("Status",
            ["All","🔴 Overdue only","🟡 High Priority +","🟢 OK only"], key="ov_s")
    with f3:
        cats = ["All","Main Engine","Aux Engines"]
        cf = st.selectbox("Engine Type", cats, key="ov_c")
    with f4:
        comp_opts = ["All"] + sorted(all_comp["description"].unique().tolist())
        compf = st.selectbox("Component", comp_opts, key="ov_comp")

    # ── Sort ────────────────────────────────────────────────────────
    sort_opt = st.radio("Sort by",
        ["📋 Report order","🔤 Component → Cylinder","⚠️ Priority → % Used"],
        horizontal=True, key="ov_sort")
    sort_mode = {"📋 Report order":"seq","🔤 Component → Cylinder":"matrix","⚠️ Priority → % Used":"priority"}[sort_opt]

    # Apply filters
    filt = all_comp.copy()
    if vf != "All Fleet":    filt = filt[filt["vessel_name"]==vf]
    if sf == "🔴 Overdue only":  filt = filt[filt["status"]=="OVERDUE"]
    elif sf == "🟡 High Priority +": filt = filt[filt["status"].isin(["OVERDUE","HIGH PRIORITY"])]
    elif sf == "🟢 OK only":        filt = filt[filt["status"]=="OK"]
    if cf == "Main Engine":  filt = filt[filt["category"]=="MAIN_ENGINE"]
    elif cf == "Aux Engines":filt = filt[filt["category"]=="AUX_ENGINE"]
    if compf != "All":       filt = filt[filt["description"]==compf]

    ns=len(filt)
    no=int((filt["status"]=="OVERDUE").sum())
    nh=int((filt["status"]=="HIGH PRIORITY").sum())
    nk=int((filt["status"]=="OK").sum())
    st.markdown(fc(ns,no,nh,nk), unsafe_allow_html=True)

    if filt.empty:
        st.markdown('<div class="ac">✓ No records match the current filter</div>', unsafe_allow_html=True)
    else:
        render_table(filt, mode=sort_mode, show_vessel=(vf=="All Fleet"),
                     height=min(860, 38*ns+44))


# ═══════════════════════════════════════════════════════════════════
# PAGE: VESSEL DETAIL
# ═══════════════════════════════════════════════════════════════════
elif page == "🚢  Vessel Detail":
    if not sel_v: st.info("Select a vessel from the sidebar."); st.stop()
    st.markdown(ph("🚢", sel_v, "Component Analysis"), unsafe_allow_html=True)

    df = get_comps(sel_v); oe = get_oe(sel_v)
    if df.empty: st.info("No data for this vessel."); st.stop()

    n_tot=len(df); n_od=int((df["status"]=="OVERDUE").sum())
    n_hp=int((df["status"]=="HIGH PRIORITY").sum())
    n_ok=int((df["status"]=="OK").sum())
    n_nd=int((df["status"]=="NO DATA").sum())

    k1,k2,k3,k4,k5=st.columns(5)
    for col,(val,lbl,clr,dly) in zip([k1,k2,k3,k4,k5],[
        (n_tot,"Total","gold",0),(n_od,"Overdue","red",.07),
        (n_hp,"High Priority","orange",.14),(n_ok,"OK","green",.21),(n_nd,"No Data","blue",.28)]):
        with col: st.markdown(kpi(val,lbl,clr,dly), unsafe_allow_html=True)

    hist = get_history(sel_v)
    if not hist.empty:
        last = hist.iloc[0]
        mt = f"{int(last['me_total_hrs']):,}" if pd.notna(last['me_total_hrs']) else "—"
        mm = f"{int(last['me_this_month']):,}" if pd.notna(last['me_this_month']) else "—"
        st.markdown(f"""
        <div class="ml">
          <span>📄 <b>{last['filename']}</b></span>
          <span>Report: <b>{last['report_date'] or '—'}</b></span>
          <span>M/E: <b>{mt}</b> total · <b>{mm}</b> this month</span>
          <span>Uploaded: <b>{str(last['uploaded_at'])[:16]}</b></span>
        </div>""", unsafe_allow_html=True)

    st.markdown("---")
    tabs = st.tabs(["⚠️  Alerts","⚙️  Main Engine","🔩  Aux Engines","🛠️  Other Equipment"])

    # ── ALERTS (priority sort) ────────────────────────────────────
    with tabs[0]:
        st.markdown(sl("Overdue & High Priority — Most Urgent First"), unsafe_allow_html=True)
        alerts = df[df["status"].isin(["OVERDUE","HIGH PRIORITY"])]
        if alerts.empty:
            st.markdown('<div class="ac">✓ All components within acceptable limits</div>',
                        unsafe_allow_html=True)
        else:
            no=int((alerts["status"]=="OVERDUE").sum())
            nh=int((alerts["status"]=="HIGH PRIORITY").sum())
            st.markdown(fc(len(alerts),no,nh,0), unsafe_allow_html=True)
            render_table(alerts, mode="priority")

    # ── MAIN ENGINE ───────────────────────────────────────────────
    with tabs[1]:
        me = df[df["category"]=="MAIN_ENGINE"]
        if me.empty: st.info("No Main Engine data.")
        else:
            fa,fb,fc_ = st.columns(3)
            with fa: mc=st.selectbox("Component",["All"]+sorted(me["description"].unique().tolist()),key="me_c")
            with fb: ms=st.selectbox("Status",["All","🔴 Overdue","🟡 High Priority +","🟢 OK"],key="me_s")
            with fc_: ms_rt=st.radio("Sort",["📋 Report","🔤 Matrix","⚠️ Priority"],horizontal=True,key="me_rt")
            v=me.copy()
            if mc!="All": v=v[v["description"]==mc]
            if ms=="🔴 Overdue": v=v[v["status"]=="OVERDUE"]
            elif ms=="🟡 High Priority +": v=v[v["status"].isin(["OVERDUE","HIGH PRIORITY"])]
            elif ms=="🟢 OK": v=v[v["status"]=="OK"]
            sm={"📋 Report":"seq","🔤 Matrix":"matrix","⚠️ Priority":"priority"}[ms_rt]
            no=int((v["status"]=="OVERDUE").sum()); nh=int((v["status"]=="HIGH PRIORITY").sum()); nk=int((v["status"]=="OK").sum())
            st.markdown(sl("Main Engine Components"), unsafe_allow_html=True)
            st.markdown(fc(len(v),no,nh,nk), unsafe_allow_html=True)
            render_table(v, mode=sm)

    # ── AUX ENGINES ───────────────────────────────────────────────
    with tabs[2]:
        aux = df[df["category"]=="AUX_ENGINE"]
        if aux.empty: st.info("No Aux Engine data.")
        else:
            fa,fb,fc_ = st.columns(3)
            with fa: ae=st.selectbox("Engine",["All"]+sorted(aux["engine_label"].unique().tolist()),key="aux_e")
            with fb: as_=st.selectbox("Status",["All","🔴 Overdue","🟡 High Priority +","🟢 OK"],key="aux_s")
            with fc_: as_rt=st.radio("Sort",["📋 Report","🔤 Matrix","⚠️ Priority"],horizontal=True,key="aux_rt")
            v=aux.copy()
            if ae!="All": v=v[v["engine_label"]==ae]
            if as_=="🔴 Overdue": v=v[v["status"]=="OVERDUE"]
            elif as_=="🟡 High Priority +": v=v[v["status"].isin(["OVERDUE","HIGH PRIORITY"])]
            elif as_=="🟢 OK": v=v[v["status"]=="OK"]
            sm={"📋 Report":"seq","🔤 Matrix":"matrix","⚠️ Priority":"priority"}[as_rt]
            no=int((v["status"]=="OVERDUE").sum()); nh=int((v["status"]=="HIGH PRIORITY").sum()); nk=int((v["status"]=="OK").sum())
            st.markdown(sl("Auxiliary Engine Components"), unsafe_allow_html=True)
            st.markdown(fc(len(v),no,nh,nk), unsafe_allow_html=True)
            render_table(v, mode=sm)

    # ── OTHER EQUIPMENT ───────────────────────────────────────────
    with tabs[3]:
        if oe.empty: st.info("No other equipment data.")
        else:
            for sec in sorted(oe["section"].unique()):
                st.markdown(sl(sec), unsafe_allow_html=True)
                sd=oe[oe["section"]==sec][["description","periodicity","last_date","run_hrs"]].copy()
                sd.columns=["Description","Periodicity","Last Date","Run Hrs"]
                st.dataframe(sd, use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════════════════════
# PAGE: UPLOAD HISTORY
# ═══════════════════════════════════════════════════════════════════
elif page == "📋  Upload History":
    st.markdown(ph("📋","Upload History","Audit Trail"), unsafe_allow_html=True)
    if not sel_v: st.info("Select a vessel from the sidebar."); st.stop()
    st.markdown(sl(sel_v), unsafe_allow_html=True)
    hist = get_history(sel_v)
    if hist.empty:
        st.info("No upload history for this vessel.")
    else:
        d=hist.copy()
        d.columns=["Filename","Report Date","M/E Total Hrs","M/E This Month","Components","Uploaded At"]
        d["M/E Total Hrs"] =d["M/E Total Hrs"].apply(lambda x:f"{int(x):,}" if pd.notna(x) else "—")
        d["M/E This Month"]=d["M/E This Month"].apply(lambda x:f"{int(x):,}" if pd.notna(x) else "—")
        st.dataframe(d, use_container_width=True, hide_index=True)
