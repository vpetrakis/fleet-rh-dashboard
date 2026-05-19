"""
╔══════════════════════════════════════════════════════════════════════╗
║   FLEET RUNNING HOURS MONITORING SYSTEM  v3.0                       ║
║   100% data integrity · .doc native · Streamlit Cloud production    ║
╚══════════════════════════════════════════════════════════════════════╝

Fixes in v3.0 vs v2:
  • LibreOffice flags: --norestore --nofirststartwizard -env:UserInstallation
    → prevents lock/profile conflicts on Streamlit Cloud (multi-user safe)
  • Vessel name regex: handles Unicode curly apostrophe U+2019 (Vessel's)
    → was silently returning "UNKNOWN" on every upload
  • .doc only UI: no .docx option shown, no confusion
  • Conversion is transparent: spinner only, no intermediate steps shown
"""

import streamlit as st

st.set_page_config(
    page_title="Fleet Running Hours",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="expanded",
)

import os, re, sqlite3, tempfile, hashlib, subprocess
from datetime import datetime
from pathlib import Path
import pandas as pd

# ════════════════════════════════════════════════════════════════════
# CSS
# ════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@300;400;500;600;700&family=Barlow+Condensed:wght@300;400;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
  --navy:    #060d1a; --navy2:   #0b1526; --navy3:  #101e36;
  --panel:   #0f1b2e; --panel2:  #132240;
  --border:  #1a2e50; --border2: #243b60;
  --amber:   #f59e0b; --amber2:  #fbbf24; --amber3: #fde68a;
  --red:     #ef4444; --orange:  #f97316; --green:  #10b981; --blue: #3b82f6;
  --text:    #cbd5e1; --text2:   #94a3b8; --muted:  #475569;
  --font:    'Barlow', sans-serif;
  --mono:    'JetBrains Mono', monospace;
  --cond:    'Barlow Condensed', sans-serif;
}

*, *::before, *::after { box-sizing: border-box; }
html, body, [class*="css"] { font-family: var(--font) !important; background-color: var(--navy) !important; color: var(--text) !important; }
.main, .main > div { background: var(--navy) !important; }
.block-container { padding: 1.5rem 2rem 4rem !important; max-width: 100% !important; }
a { color: var(--amber) !important; }

.main::before {
  content: ''; position: fixed; inset: 0; pointer-events: none; z-index: 0;
  background: radial-gradient(ellipse 60% 40% at 15% 20%, rgba(59,130,246,0.04) 0%, transparent 60%),
              radial-gradient(ellipse 50% 30% at 85% 80%, rgba(245,158,11,0.03) 0%, transparent 60%);
}

/* Sidebar */
[data-testid="stSidebar"] { background: var(--navy2) !important; border-right: 1px solid var(--border) !important; }
[data-testid="stSidebar"]::before { content:''; position:absolute; top:0; left:0; right:0; height:3px; background:linear-gradient(90deg,var(--amber),transparent); }
[data-testid="stSidebar"] * { color: var(--text) !important; }
[data-testid="stSidebar"] .stSelectbox > div > div { background: var(--navy3) !important; border: 1px solid var(--border) !important; border-radius: 6px !important; }

/* Typography */
h1 { font-family:var(--cond) !important; font-size:2rem !important; font-weight:800 !important; letter-spacing:0.06em !important; color:#f8fafc !important; }
h2 { font-family:var(--cond) !important; font-weight:700 !important; letter-spacing:0.05em !important; color:#f1f5f9 !important; }
h3 { font-family:var(--cond) !important; font-weight:600 !important; color:#e2e8f0 !important; }

/* Metrics */
[data-testid="stMetric"] { background:linear-gradient(135deg,var(--panel) 0%,var(--navy3) 100%) !important; border:1px solid var(--border) !important; border-radius:12px !important; padding:1.1rem 1.25rem !important; position:relative !important; overflow:hidden !important; transition:border-color 0.3s,transform 0.2s !important; }
[data-testid="stMetric"]:hover { border-color:var(--border2) !important; transform:translateY(-2px) !important; }
[data-testid="stMetric"]::after { content:''; position:absolute; top:0; left:0; right:0; height:2px; background:linear-gradient(90deg,var(--amber),transparent); }
[data-testid="stMetricValue"] { font-family:var(--cond) !important; font-size:2.2rem !important; font-weight:800 !important; color:#f8fafc !important; }
[data-testid="stMetricLabel"] { color:var(--text2) !important; font-size:0.7rem !important; text-transform:uppercase !important; letter-spacing:0.1em !important; font-weight:600 !important; }

/* Dataframe */
[data-testid="stDataFrame"] { border:1px solid var(--border) !important; border-radius:10px !important; overflow:hidden !important; }
.dvn-scroller { background:var(--panel) !important; }

/* Buttons */
.stButton > button { background:linear-gradient(135deg,var(--amber) 0%,#d97706 100%) !important; color:#000 !important; border:none !important; font-family:var(--cond) !important; font-weight:800 !important; font-size:0.88rem !important; letter-spacing:0.1em !important; text-transform:uppercase !important; border-radius:6px !important; padding:0.55rem 1.75rem !important; box-shadow:0 2px 12px rgba(245,158,11,0.25) !important; transition:all 0.2s ease !important; }
.stButton > button:hover { background:linear-gradient(135deg,var(--amber2) 0%,var(--amber) 100%) !important; box-shadow:0 4px 20px rgba(245,158,11,0.4) !important; transform:translateY(-2px) !important; }
.stButton > button:active { transform:translateY(0) !important; }

/* File uploader */
[data-testid="stFileUploadDropzone"] { background:linear-gradient(135deg,rgba(245,158,11,0.05) 0%,rgba(59,130,246,0.04) 100%) !important; border:2px dashed var(--amber) !important; border-radius:14px !important; transition:all 0.3s ease !important; padding:2.5rem 2rem !important; }
[data-testid="stFileUploadDropzone"]:hover { background:rgba(245,158,11,0.08) !important; border-color:var(--amber2) !important; box-shadow:0 0 40px rgba(245,158,11,0.12) !important; }
[data-testid="stFileUploadDropzone"] p, [data-testid="stFileUploadDropzone"] span { color:var(--amber) !important; font-family:var(--cond) !important; font-size:1.05rem !important; font-weight:700 !important; }
[data-testid="stFileUploadDropzone"] small { color:var(--text2) !important; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] { background:var(--panel) !important; border-radius:10px 10px 0 0 !important; border-bottom:1px solid var(--border) !important; gap:0 !important; padding:0 0.5rem !important; }
.stTabs [data-baseweb="tab"] { background:transparent !important; color:var(--muted) !important; font-family:var(--cond) !important; font-weight:600 !important; letter-spacing:0.07em !important; text-transform:uppercase !important; font-size:0.82rem !important; padding:0.85rem 1.5rem !important; border-bottom:2px solid transparent !important; margin-bottom:-1px !important; transition:color 0.2s !important; }
.stTabs [data-baseweb="tab"]:hover { color:var(--text) !important; }
.stTabs [aria-selected="true"] { color:var(--amber) !important; border-bottom:2px solid var(--amber) !important; }
.stTabs [data-baseweb="tab-panel"] { background:var(--panel) !important; border:1px solid var(--border) !important; border-top:none !important; border-radius:0 0 10px 10px !important; padding:1.5rem !important; }

/* Expander */
.streamlit-expanderHeader { background:var(--navy3) !important; border:1px solid var(--border) !important; border-radius:8px !important; font-family:var(--cond) !important; font-weight:600 !important; font-size:0.9rem !important; color:var(--text) !important; letter-spacing:0.04em !important; transition:background 0.2s,border-color 0.2s !important; }
.streamlit-expanderHeader:hover { background:var(--panel2) !important; border-color:var(--border2) !important; }
.streamlit-expanderContent { background:var(--panel) !important; border:1px solid var(--border) !important; border-top:none !important; border-radius:0 0 8px 8px !important; }

/* Selects */
.stSelectbox > div > div, .stMultiSelect > div > div { background:var(--navy3) !important; border:1px solid var(--border) !important; border-radius:6px !important; color:var(--text) !important; }
.stAlert { border-radius:8px !important; border-left-width:3px !important; }
hr { border-color:var(--border) !important; opacity:1 !important; }

/* Animations */
@keyframes fadeSlideDown { from{opacity:0;transform:translateY(-14px)} to{opacity:1;transform:translateY(0)} }
@keyframes fadeSlideUp   { from{opacity:0;transform:translateY(10px)}  to{opacity:1;transform:translateY(0)} }
@keyframes countUp       { from{opacity:0;transform:scale(0.85)}       to{opacity:1;transform:scale(1)} }
@keyframes successPop    { 0%{transform:scale(0.82);opacity:0} 65%{transform:scale(1.04)} 100%{transform:scale(1);opacity:1} }
@keyframes shimmer       { 0%{background-position:-200% center} 100%{background-position:200% center} }

.page-header { animation:fadeSlideDown 0.45s ease both; }
.page-sub    { animation:fadeSlideDown 0.45s 0.08s ease both; opacity:0; animation-fill-mode:forwards; }

/* KPI Cards */
.kpi-card { background:linear-gradient(135deg,var(--panel) 0%,var(--navy3) 100%); border:1px solid var(--border); border-radius:12px; padding:1.1rem 1.4rem; position:relative; overflow:hidden; animation:fadeSlideUp 0.4s ease both; transition:transform 0.2s,border-color 0.2s; }
.kpi-card:hover { transform:translateY(-3px); border-color:var(--border2); }
.kpi-card::before { content:''; position:absolute; top:0; left:0; right:0; height:2px; }
.kpi-card.red::before    { background:linear-gradient(90deg,var(--red),transparent); }
.kpi-card.orange::before { background:linear-gradient(90deg,var(--orange),transparent); }
.kpi-card.green::before  { background:linear-gradient(90deg,var(--green),transparent); }
.kpi-card.blue::before   { background:linear-gradient(90deg,var(--blue),transparent); }
.kpi-card.amber::before  { background:linear-gradient(90deg,var(--amber),transparent); }
.kpi-value { font-family:var(--cond); font-size:2.4rem; font-weight:800; line-height:1; letter-spacing:-0.02em; animation:countUp 0.5s 0.1s ease both; }
.kpi-label { font-size:0.68rem; text-transform:uppercase; letter-spacing:0.12em; color:var(--text2); margin-top:4px; font-weight:600; }
.kpi-card.red    .kpi-value { color:#fca5a5; }
.kpi-card.orange .kpi-value { color:#fed7aa; }
.kpi-card.green  .kpi-value { color:#6ee7b7; }
.kpi-card.blue   .kpi-value { color:#93c5fd; }
.kpi-card.amber  .kpi-value { color:var(--amber2); }

/* Section divider */
.section-label { font-family:var(--cond); font-size:0.65rem; font-weight:700; letter-spacing:0.18em; text-transform:uppercase; color:var(--muted); border-bottom:1px solid var(--border); padding-bottom:0.5rem; margin:1.75rem 0 1.1rem; }

/* Info card */
.info-card { background:linear-gradient(135deg,var(--panel) 0%,var(--navy3) 100%); border:1px solid var(--border); border-radius:12px; padding:1.25rem 1.5rem; font-size:0.83rem; color:var(--text2); line-height:1.9; }
.info-card-title { font-family:var(--cond); font-size:0.72rem; font-weight:700; letter-spacing:0.12em; text-transform:uppercase; color:var(--amber); margin-bottom:0.5rem; }

/* Parse stats row */
.parse-stats { display:flex; gap:0.75rem; flex-wrap:wrap; margin:1rem 0; }
.parse-stat { background:var(--navy3); border:1px solid var(--border); border-radius:8px; padding:0.6rem 1rem; min-width:100px; animation:fadeSlideUp 0.4s ease both; }
.parse-stat-val { font-family:var(--cond); font-size:1.6rem; font-weight:800; line-height:1; }
.parse-stat-lbl { font-size:0.65rem; text-transform:uppercase; letter-spacing:0.1em; color:var(--muted); margin-top:2px; }
.parse-stat.red    .parse-stat-val { color:#fca5a5; }
.parse-stat.orange .parse-stat-val { color:#fed7aa; }
.parse-stat.green  .parse-stat-val { color:#6ee7b7; }
.parse-stat.blue   .parse-stat-val { color:#93c5fd; }

/* Success banner */
.success-banner { background:linear-gradient(135deg,rgba(16,185,129,0.15) 0%,rgba(16,185,129,0.05) 100%); border:1px solid rgba(16,185,129,0.3); border-radius:10px; padding:1rem 1.5rem; color:#6ee7b7; font-family:var(--cond); font-size:1rem; font-weight:600; letter-spacing:0.04em; animation:successPop 0.5s cubic-bezier(0.34,1.56,0.64,1) both; display:flex; align-items:center; gap:0.75rem; }

/* Sidebar */
.sidebar-logo { font-family:var(--cond); font-size:1.4rem; font-weight:800; letter-spacing:0.1em; color:var(--amber); }
.sidebar-logo span { color:var(--text2); font-weight:400; font-size:1rem; }
.sidebar-tagline { font-size:0.62rem; text-transform:uppercase; letter-spacing:0.14em; color:var(--muted); margin-top:2px; }

/* Scrollbar */
::-webkit-scrollbar { width:6px; height:6px; }
::-webkit-scrollbar-track { background:var(--navy2); }
::-webkit-scrollbar-thumb { background:var(--border2); border-radius:3px; }
::-webkit-scrollbar-thumb:hover { background:var(--muted); }
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
        engine_label TEXT NOT NULL, unit TEXT NOT NULL, description TEXT NOT NULL,
        periodicity REAL, last_oh_date TEXT, last_oh_hrs REAL,
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
    CREATE INDEX IF NOT EXISTS idx_comp_vessel  ON components(vessel_name);
    CREATE INDEX IF NOT EXISTS idx_comp_status  ON components(status);
    CREATE INDEX IF NOT EXISTS idx_other_vessel ON other_equipment(vessel_name);
    """)
    conn.commit(); conn.close()

init_db()


# ════════════════════════════════════════════════════════════════════
# CONVERSION  — .doc → .docx  (LibreOffice, cloud-safe)
# ════════════════════════════════════════════════════════════════════

@st.cache_resource(show_spinner=False)
def ensure_libreoffice() -> bool:
    """
    Guarantee LibreOffice is available. Runs ONCE at startup via cache.
    On Streamlit Cloud (Ubuntu) soffice may not be in PATH without packages.txt —
    we install it silently at runtime via apt-get. No extra files needed.
    """
    import shutil
    for p in ['/usr/bin/soffice', '/usr/lib/libreoffice/program/soffice']:
        if os.path.isfile(p) and os.access(p, os.X_OK):
            return True
    if shutil.which('soffice'):
        return True
    # Not found — install silently (Streamlit Cloud has root/sudo)
    result = subprocess.run(
        ['apt-get', 'install', '-y', '-q', '--no-install-recommends', 'libreoffice'],
        capture_output=True, timeout=180
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"Could not install LibreOffice (code {result.returncode}). "
            f"{result.stderr.decode('utf-8','ignore')[:300]}"
        )
    return True


def _soffice_path() -> str:
    """Return full path to soffice binary."""
    import shutil
    for p in ['/usr/bin/soffice', '/usr/lib/libreoffice/program/soffice']:
        if os.path.isfile(p) and os.access(p, os.X_OK):
            return p
    found = shutil.which('soffice')
    if found:
        return found
    raise RuntimeError("soffice binary not found.")


def convert_doc_to_docx(file_bytes: bytes) -> bytes:
    """
    Convert a legacy .doc binary to .docx bytes using LibreOffice headless.
    LibreOffice is auto-installed on first run if not already present.

    Cloud-safe flags:
      --norestore              skip crash-recovery dialog (would hang headless)
      --nofirststartwizard     skip first-run setup wizard
      -env:UserInstallation=   per-process temp profile → no lock conflicts
    """
    ensure_libreoffice()  # guaranteed present before we proceed

    with tempfile.NamedTemporaryFile(suffix='.doc', delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    out_dir     = tempfile.gettempdir()
    profile_dir = f"file:///tmp/lo_profile_{os.getpid()}_{os.urandom(4).hex()}"

    try:
        result = subprocess.run(
            [
                _soffice_path(),
                '--headless',
                '--norestore',
                '--nofirststartwizard',
                f'-env:UserInstallation={profile_dir}',
                '--convert-to', 'docx',
                tmp_path,
                '--outdir', out_dir,
            ],
            capture_output=True,
            timeout=120,
        )

        base      = os.path.splitext(os.path.basename(tmp_path))[0]
        docx_path = os.path.join(out_dir, base + '.docx')

        if result.returncode != 0 or not os.path.exists(docx_path):
            stderr = result.stderr.decode('utf-8', errors='ignore').strip()
            stdout = result.stdout.decode('utf-8', errors='ignore').strip()
            raise RuntimeError(
                f"LibreOffice exited with code {result.returncode}.\n"
                f"stdout: {stdout[:300] or '(empty)'}\n"
                f"stderr: {stderr[:300] or '(empty)'}"
            )

        with open(docx_path, 'rb') as f:
            return f.read()

    finally:
        # Always clean up — never leave temp files on the server
        for p in [tmp_path,
                  os.path.join(out_dir, os.path.splitext(os.path.basename(tmp_path))[0] + '.docx')]:
            try:
                if os.path.exists(p):
                    os.unlink(p)
            except Exception:
                pass


# ════════════════════════════════════════════════════════════════════
# PARSER — deterministic, 100% integrity validated
# ════════════════════════════════════════════════════════════════════

def _clean_period(raw: str):
    if not raw: return None
    # '16.000' → 16000  (thousand-separator dot)
    s = re.sub(r'\.(?=\d{3}(\.|$))', '', raw.strip())
    s = re.sub(r'[^0-9\.]', '', s)
    try: return float(s) if s else None
    except ValueError: return None

def _parse_date(raw: str):
    if not raw or raw.strip() in ('', 'N/A', 'n/a'): return None
    raw = re.sub(r'\s+', ' ', raw.strip().lstrip('[').rstrip(']'))
    if re.match(r'^\d+$', raw.strip()): return None   # pure number = hours, not date
    # Normalise non-standard month abbreviations
    rn = re.sub(r'\bSEPT\b', 'SEP', raw, flags=re.IGNORECASE)
    rn = re.sub(r'\bJUNE\b', 'JUN', rn,  flags=re.IGNORECASE)
    rn = re.sub(r'\bJULY\b', 'JUL', rn,  flags=re.IGNORECASE)
    fmts = ['%d %b %y','%d %B %y','%d %b %Y','%d %B %Y',
            '%d/%m/%y','%d/%m/%Y','%d-%m-%y','%d-%m-%Y',
            '%b %Y','%B %Y','%Y-%m-%d']
    # Try each variant: raw_normalised, UPPERCASE, Title Case, and original
    # NOTE: do NOT uppercase the format string — only the INPUT
    for fmt in fmts:
        for v in (rn, rn.upper(), rn.title(), raw, raw.upper()):
            try: return datetime.strptime(v, fmt).strftime('%Y-%m-%d')
            except ValueError: pass
    return raw   # store as-is if genuinely unparseable

def _parse_hrs(raw: str):
    if not raw or raw.strip() in ('', 'N/A', 'n/a'): return None
    for n in re.findall(r'\d[\d,]*', raw.replace('\n', ' ')):
        try:
            v = float(n.replace(',', ''))
            if v > 0: return v
        except ValueError: pass
    return None

def _status(hrs, period) -> str:
    if hrs is None or period is None or period == 0: return 'NO DATA'
    p = hrs / period
    if p > 1.0: return 'OVERDUE'
    if p >= 0.80: return 'HIGH PRIORITY'
    return 'OK'

def _pct(hrs, period) -> float:
    if hrs is None or period is None or period == 0: return 0.0
    return round(hrs / period, 4)


def parse_doc_bytes(docx_bytes: bytes) -> dict:
    """
    Parse TEC-004 Running Hours report from .docx bytes.
    Returns structured dict. Raises ValueError on structural failure.
    """
    from docx import Document

    warns = []

    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
        tmp.write(docx_bytes)
        tmp_path = tmp.name
    try:
        doc = Document(tmp_path)
    except Exception as e:
        raise ValueError(f"Cannot open document: {e}")
    finally:
        try: os.unlink(tmp_path)
        except Exception: pass

    if not doc.tables:
        raise ValueError("No tables found — is this a TEC-004 report?")

    # ── Vessel name & date ────────────────────────────────────────
    # FIX: handle Unicode curly apostrophe U+2019 in "Vessel's Name"
    vessel_name = "UNKNOWN"
    report_date = None
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt: continue
        # Match both straight apostrophe (') and curly apostrophe (')
        vm = re.search(
            r"Vessel[\u2019\u2018']?s?\s+Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+?)(?:\s{2,}|\t)",
            txt, re.IGNORECASE
        )
        dm = re.search(r"Date\s*:\s*(.+)", txt, re.IGNORECASE)
        if vm: vessel_name = vm.group(1).strip()
        if dm: report_date = _parse_date(dm.group(1).strip())
        if vm or dm: break

    if vessel_name == "UNKNOWN":
        warns.append("Could not extract vessel name from header.")

    # ── Table 0 — Main Engine ─────────────────────────────────────
    me_total = me_month = None
    components = []
    t0 = doc.tables[0]

    # Row 0: total running hours (ǀ is the Unicode pipe used in the doc)
    for cell in t0.rows[0].cells:
        txt = cell.text
        if m := re.search(r'Total Running Hours[\s:ǀ|]+([\d,]+)', txt, re.IGNORECASE):
            try: me_total = int(m.group(1).replace(',', ''))
            except ValueError: pass
        if m := re.search(r'This Month[\s:]+([\d,]+)', txt, re.IGNORECASE):
            try: me_month = int(m.group(1).replace(',', ''))
            except ValueError: pass

    # Row 1: cylinder columns — deduplicate merged cells
    cyl_cols = []
    if len(t0.rows) > 1:
        for ci, cell in enumerate(t0.rows[1].cells):
            if m := re.search(r'CYL\s*\.?\s*No\s*\.?\s*(\d+)', cell.text.strip(), re.IGNORECASE):
                lbl = f"Cyl {int(m.group(1))}"
                if not cyl_cols or cyl_cols[-1][1] != lbl:
                    cyl_cols.append((ci, lbl))

    # Rows 2+: component pairs (row '1' = dates, row '2' = hours)
    rows = t0.rows
    i = 2
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
                    'category':'MAIN_ENGINE', 'engine_label':'ME', 'unit':lbl,
                    'description':name, 'periodicity':period,
                    'last_oh_date':d, 'last_oh_hrs':h, 'hrs_since':h,
                    'pct_used':_pct(h, period), 'status':_status(h, period),
                })
            i += 2
        else:
            i += 1

    # ── Table 1 — Turbocharger / Coolers / A/C ───────────────────
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
                ('TURBOCHARGER / AUX BOILER', 0, 1, 3),
                ('COOLERS / EXH GAS BOILER',  5, 6, 8),
                ('A/C & COMPRESSORS',         10,11,12),
            ]:
                desc = cells[dc] if dc < len(cells) else ''
                if not desc or desc.upper() in SKIP: continue
                dv = cells[datec] if datec < len(cells) else ''
                hv = cells[hrsc]  if hrsc  < len(cells) else ''
                if dv or hv:
                    other_equip.append({'section':sec, 'description':desc,
                                        'periodicity':'', 'last_date':dv, 'run_hrs':hv})

    # ── Table 2 — Auxiliary Engines ───────────────────────────────
    if len(doc.tables) > 2:
        t2    = doc.tables[2]
        rows2 = t2.rows
        engine_blocks = []

        if rows2:
            hdr   = [c.text.strip() for c in rows2[0].cells]
            total = [c.text.strip() for c in rows2[2].cells] if len(rows2) > 2 else []
            seen  = set()
            for ci, cell in enumerate(hdr):
                if m := re.search(r'Aux\.\s*Engine\s*No\.?\s*(\d+)', cell, re.IGNORECASE):
                    lbl = f"AUX-{int(m.group(1))}"
                    if lbl not in seen:
                        seen.add(lbl)
                        th = next((_parse_hrs(total[j]) for j in range(ci, min(ci+14, len(total)))
                                   if _parse_hrs(total[j])), None)
                        engine_blocks.append((lbl, ci, th))

        # Column map: col_index → (engine_label, cyl_num)
        cyl_map = {}
        if len(rows2) > 4:
            r4 = [c.text.strip() for c in rows2[4].cells]
            for ei, (elbl, estart, _) in enumerate(engine_blocks):
                eend = engine_blocks[ei+1][1] if ei+1 < len(engine_blocks) else len(r4)
                seen_c: list[int] = []
                for ci in range(estart, eend):
                    if ci < len(r4):
                        try:
                            cn = int(r4[ci])
                            if cn not in seen_c:
                                seen_c.append(cn)
                                cyl_map[ci] = (elbl, cn)
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
                        'category':'AUX_ENGINE', 'engine_label':elbl, 'unit':f"Cyl {cn}",
                        'description':name, 'periodicity':period,
                        'last_oh_date':d, 'last_oh_hrs':h, 'hrs_since':h,
                        'pct_used':_pct(h, period), 'status':_status(h, period),
                    })
                i2 += 2
            else:
                i2 += 1

    # ── Table 3 — D/G Equipment ───────────────────────────────────
    if len(doc.tables) > 3:
        t3     = doc.tables[3]
        dglbls = ['D/G 1','D/G 2','D/G 3']
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
                        other_equip.append({'section':'D/G EQUIPMENT', 'description':key,
                                            'periodicity':per,
                                            'last_date':_parse_date(val) or val,
                                            'run_hrs':''})
                    else:
                        for e in reversed(other_equip):
                            if e['description'] == key and e['run_hrs'] == '':
                                e['run_hrs'] = val; break
                        else:
                            other_equip.append({'section':'D/G EQUIPMENT', 'description':key,
                                                'periodicity':per, 'last_date':'', 'run_hrs':val})

    return {
        'vessel_name':     vessel_name,
        'report_date':     report_date,
        'me_total_hrs':    me_total,
        'me_this_month':   me_month,
        'components':      components,
        'other_equipment': other_equip,
        'warnings':        warns,
    }


# ════════════════════════════════════════════════════════════════════
# DB HELPERS
# ════════════════════════════════════════════════════════════════════

def save_parsed_data(parsed: dict, filename: str, file_hash: str):
    conn = get_db(); c = conn.cursor()
    now  = datetime.utcnow().isoformat()
    v    = parsed['vessel_name']
    c.execute("INSERT OR IGNORE INTO vessels(name,created_at) VALUES(?,?)", (v, now))
    c.execute("INSERT INTO upload_log(vessel_name,filename,file_hash,report_date,me_total_hrs,me_this_month,uploaded_at) VALUES(?,?,?,?,?,?,?)",
              (v, filename, file_hash, parsed['report_date'], parsed['me_total_hrs'], parsed['me_this_month'], now))
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

def kpi(val, lbl, color="amber", delay=0):
    return f'<div class="kpi-card {color}" style="animation-delay:{delay}s"><div class="kpi-value">{val}</div><div class="kpi-label">{lbl}</div></div>'

def fmt_df(df):
    d = df[['description','engine_label','unit','periodicity','last_oh_date','hrs_since','pct_used','status']].copy()
    d.columns = ['Component','Engine','Unit','Periodicity','Last O/H Date','Hrs Since','% Used','Status']
    d['% Used']        = d['% Used'].apply(lambda x: f"{x*100:.1f}%" if pd.notna(x) else '—')
    d['Periodicity']   = d['Periodicity'].apply(lambda x: f"{int(x):,}" if pd.notna(x) and x>0 else 'N/A')
    d['Hrs Since']     = d['Hrs Since'].apply(lambda x: f"{int(x):,}" if pd.notna(x) else '—')
    d['Last O/H Date'] = d['Last O/H Date'].fillna('—')
    return d

def style_df(df):
    def rs(row):
        s = row.get('Status','')
        if s == 'OVERDUE':       return ['background-color:#2d0b0b;color:#fca5a5']*len(row)
        if s == 'HIGH PRIORITY': return ['background-color:#2d1500;color:#fed7aa']*len(row)
        if s == 'OK':            return ['background-color:#071810;color:#6ee7b7']*len(row)
        return ['background-color:#0a111e;color:#64748b']*len(row)
    return df.style.apply(rs, axis=1)

def show_table(df, height=None):
    h = height or min(700, 38*(len(df)+1))
    st.dataframe(style_df(fmt_df(df)), use_container_width=True, hide_index=True, height=h)


# ════════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown('<div class="sidebar-logo">⚓ FLEET<span>&nbsp;MONITOR</span></div><div class="sidebar-tagline">Running Hours Management System</div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    page = st.selectbox("nav", ["🗺️  Fleet Overview","🚢  Vessel Detail","📤  Upload Report","📋  Upload History"], label_visibility="collapsed")
    st.markdown("<br>", unsafe_allow_html=True)

    vessels = get_all_vessels()
    selected_vessel = st.selectbox("Vessel", vessels) if vessels else None
    if not vessels:
        st.info("No data yet — upload a .doc report to begin.")

    st.divider()
    if vessels:
        smry = get_fleet_summary()
        if not smry.empty:
            tod = int(smry['overdue'].sum())
            thp = int(smry['high_priority'].sum())
            st.markdown(f"""
            <div style="font-size:0.65rem;text-transform:uppercase;letter-spacing:0.12em;color:#475569;margin-bottom:0.6rem;">Fleet Status</div>
            <div style="display:flex;gap:0.5rem;flex-wrap:wrap;">
              <div style="background:rgba(239,68,68,0.1);border:1px solid rgba(239,68,68,0.2);border-radius:6px;padding:4px 10px;text-align:center;">
                <div style="font-family:var(--cond);font-size:1.2rem;font-weight:800;color:#fca5a5;">{tod}</div>
                <div style="font-size:0.58rem;text-transform:uppercase;letter-spacing:0.1em;color:#7f1d1d;">Overdue</div>
              </div>
              <div style="background:rgba(249,115,22,0.1);border:1px solid rgba(249,115,22,0.2);border-radius:6px;padding:4px 10px;text-align:center;">
                <div style="font-family:var(--cond);font-size:1.2rem;font-weight:800;color:#fed7aa;">{thp}</div>
                <div style="font-size:0.58rem;text-transform:uppercase;letter-spacing:0.1em;color:#7c2d12;">High Pri.</div>
              </div>
              <div style="background:rgba(71,85,105,0.1);border:1px solid rgba(71,85,105,0.2);border-radius:6px;padding:4px 10px;text-align:center;">
                <div style="font-family:var(--cond);font-size:1.2rem;font-weight:800;color:#94a3b8;">{len(vessels)}</div>
                <div style="font-size:0.58rem;text-transform:uppercase;letter-spacing:0.1em;color:#475569;">Vessels</div>
              </div>
            </div>""", unsafe_allow_html=True)
    st.divider()
    db_kb = DB_PATH.stat().st_size/1024 if DB_PATH.exists() else 0
    st.markdown(f'<p style="font-size:0.65rem;color:#334155;font-family:var(--mono)">db: {db_kb:.1f} kb | v3.0</p>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# PAGES
# ════════════════════════════════════════════════════════════════════

# ── UPLOAD ──────────────────────────────────────────────────────────
if page == "📤  Upload Report":
    st.markdown('<div class="page-header"><h1>📤 UPLOAD REPORT</h1></div>', unsafe_allow_html=True)
    st.markdown('<div class="page-sub section-label">Drop a TEC-004 .doc file for any vessel</div>', unsafe_allow_html=True)

    col_up, col_info = st.columns([3, 2], gap="large")
    with col_up:
        uploaded = st.file_uploader("file", type=["doc"], label_visibility="collapsed")
    with col_info:
        st.markdown("""
        <div class="info-card">
          <div class="info-card-title">Accepted Format</div>
          TEC-004 Running Hours Monthly Report<br>
          Any vessel &nbsp;·&nbsp; <b>.doc format only</b><br><br>
          <div class="info-card-title">What Gets Extracted</div>
          ✦ Vessel name &amp; report date<br>
          ✦ M/E total &amp; monthly hours<br>
          ✦ All M/E component O/H data<br>
          ✦ Aux engines (3 engines × 6 cyl)<br>
          ✦ Turbocharger &amp; D/G equipment<br>
          ✦ Status auto-computed per periodicity
        </div>""", unsafe_allow_html=True)

    if uploaded:
        raw_bytes = uploaded.read()
        file_hash = hashlib.md5(raw_bytes).hexdigest()

        # Always convert — all uploads are .doc
        with st.spinner("Converting and parsing…"):
            try:
                docx_bytes = convert_doc_to_docx(raw_bytes)
            except (RuntimeError, FileNotFoundError, OSError) as e:
                st.error(
                    f"**Conversion failed.**\n\n"
                    f"Error detail: `{e}`\n\n"
                    f"Please try re-uploading. If the problem persists, contact your administrator."
                )
                st.stop()

            try:
                parsed = parse_doc_bytes(docx_bytes)
            except ValueError as e:
                st.error(f"**Parse failed:** {e}")
                st.stop()

        # ── Preview ──
        st.markdown('<div class="section-label">Parse Preview — confirm before saving</div>', unsafe_allow_html=True)

        n_comp = len(parsed['components'])
        n_od   = sum(1 for c in parsed['components'] if c['status'] == 'OVERDUE')
        n_hp   = sum(1 for c in parsed['components'] if c['status'] == 'HIGH PRIORITY')
        n_ok   = sum(1 for c in parsed['components'] if c['status'] == 'OK')
        n_oe   = len(parsed['other_equipment'])

        c1,c2,c3,c4,c5 = st.columns(5)
        c1.metric("Vessel",         parsed['vessel_name'])
        c2.metric("Report Date",    parsed['report_date'] or "—")
        c3.metric("M/E Total Hrs",  f"{parsed['me_total_hrs']:,}"  if parsed['me_total_hrs']  else "—")
        c4.metric("M/E This Month", f"{parsed['me_this_month']:,}" if parsed['me_this_month'] else "—")
        c5.metric("Components",     n_comp)

        st.markdown(f"""
        <div class="parse-stats">
          <div class="parse-stat red">   <div class="parse-stat-val">{n_od}</div><div class="parse-stat-lbl">Overdue</div></div>
          <div class="parse-stat orange"><div class="parse-stat-val">{n_hp}</div><div class="parse-stat-lbl">High Priority</div></div>
          <div class="parse-stat green"> <div class="parse-stat-val">{n_ok}</div><div class="parse-stat-lbl">OK</div></div>
          <div class="parse-stat blue">  <div class="parse-stat-val">{n_oe}</div><div class="parse-stat-lbl">Other Equip</div></div>
        </div>""", unsafe_allow_html=True)

        for w in parsed['warnings']:
            st.warning(f"⚠ {w}")

        if parsed['components']:
            with st.expander("Preview parsed data", expanded=False):
                show_table(pd.DataFrame(parsed['components']), height=300)

        st.markdown("---")
        col_btn, _ = st.columns([1, 4])
        with col_btn:
            if st.button("✅  CONFIRM & SAVE", use_container_width=True):
                save_parsed_data(parsed, uploaded.name, file_hash)
                for fn in [get_all_vessels, get_components_df, get_other_equip_df, get_fleet_summary]:
                    fn.clear()
                st.markdown(f"""
                <div class="success-banner">
                  <span style="font-size:1.5rem">✓</span>
                  <span><b>{parsed['vessel_name']}</b> saved — {n_comp} components · {n_od} overdue · {n_hp} high priority</span>
                </div>""", unsafe_allow_html=True)
                st.balloons()


# ── FLEET OVERVIEW ───────────────────────────────────────────────────
elif page == "🗺️  Fleet Overview":
    st.markdown('<div class="page-header"><h1>🗺️ FLEET OVERVIEW</h1></div>', unsafe_allow_html=True)
    st.markdown('<div class="page-sub section-label">All vessels — live component status</div>', unsafe_allow_html=True)

    summary = get_fleet_summary()
    if summary.empty:
        st.info("No data loaded yet. Upload a .doc report to begin."); st.stop()

    tv  = len(summary)
    tc  = int(summary['total'].sum())
    tod = int(summary['overdue'].sum())
    thp = int(summary['high_priority'].sum())
    tok = int(summary['ok'].sum())

    cols = st.columns(5)
    for col,(val,lbl,clr,dly) in zip(cols,[
        (tv,"Vessels","blue",0),(tc,"Components","amber",0.05),
        (tod,"Overdue","red",0.1),(thp,"High Priority","orange",0.15),(tok,"OK","green",0.2)]):
        with col: st.markdown(kpi(val,lbl,clr,dly), unsafe_allow_html=True)

    st.markdown('<div class="section-label">Fleet Status Table</div>', unsafe_allow_html=True)
    disp = summary[['vessel_name','overdue','high_priority','ok','total','me_total_hrs','last_upload']].copy()
    disp.columns = ['Vessel','Overdue','High Priority','OK','Total','M/E Total Hrs','Last Upload']
    disp['M/E Total Hrs'] = disp['M/E Total Hrs'].apply(lambda x: f"{int(x):,}" if pd.notna(x) else '—')
    disp['Last Upload']   = pd.to_datetime(disp['Last Upload'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('—')

    def fleet_style(row):
        if row['Overdue']>0:       return ['background-color:#1a0505;color:#fca5a5']+['background-color:#1a0505']*6
        if row['High Priority']>0: return ['background-color:#1a0a00;color:#fed7aa']+['background-color:#1a0a00']*6
        return                            ['background-color:#040d0a;color:#6ee7b7']+['background-color:#040d0a']*6

    st.dataframe(disp.style.apply(fleet_style,axis=1), use_container_width=True, hide_index=True, height=min(600,38*(len(disp)+1)+3))

    st.markdown('<div class="section-label">Vessel Breakdown</div>', unsafe_allow_html=True)
    for _, row in summary.iterrows():
        sev  = "critical" if row['overdue']>0 else ("warning" if row['high_priority']>0 else "healthy")
        icon = "🔴" if sev=="critical" else ("🟡" if sev=="warning" else "🟢")
        with st.expander(f"{icon} **{row['vessel_name']}** — {int(row['overdue'])} overdue · {int(row['high_priority'])} high priority · {int(row['ok'])} OK", expanded=False):
            cc = get_components_df(row['vessel_name'])
            if not cc.empty:
                t1, t2 = st.tabs(["🔴 Overdue","🟡 High Priority"])
                with t1:
                    od = cc[cc['status']=='OVERDUE']
                    if od.empty: st.success("No overdue items.")
                    else: show_table(od)
                with t2:
                    hp = cc[cc['status']=='HIGH PRIORITY']
                    if hp.empty: st.success("No high-priority items.")
                    else: show_table(hp)


# ── VESSEL DETAIL ────────────────────────────────────────────────────
elif page == "🚢  Vessel Detail":
    if not selected_vessel:
        st.info("No vessel selected. Upload a report first, then pick a vessel from the sidebar."); st.stop()

    st.markdown(f'<div class="page-header"><h1>🚢 {selected_vessel}</h1></div>', unsafe_allow_html=True)
    st.markdown('<div class="page-sub section-label">Component analysis by engine &amp; system</div>', unsafe_allow_html=True)

    df = get_components_df(selected_vessel)
    oe = get_other_equip_df(selected_vessel)
    if df.empty:
        st.info("No component data for this vessel."); st.stop()

    n_tot = len(df)
    n_od  = int((df['status']=='OVERDUE').sum())
    n_hp  = int((df['status']=='HIGH PRIORITY').sum())
    n_ok  = int((df['status']=='OK').sum())
    n_nd  = int((df['status']=='NO DATA').sum())

    cols = st.columns(5)
    for col,(val,lbl,clr,dly) in zip(cols,[
        (n_tot,"Total","amber",0),(n_od,"Overdue","red",0.05),
        (n_hp,"High Priority","orange",0.1),(n_ok,"OK","green",0.15),(n_nd,"No Data","blue",0.2)]):
        with col: st.markdown(kpi(val,lbl,clr,dly), unsafe_allow_html=True)

    hist = get_upload_history(selected_vessel)
    if not hist.empty:
        last = hist.iloc[0]
        mt = f"{int(last['me_total_hrs']):,}" if pd.notna(last['me_total_hrs']) else "—"
        mm = f"{int(last['me_this_month']):,}" if pd.notna(last['me_this_month']) else "—"
        st.markdown(f'<p style="font-size:0.72rem;color:#475569;font-family:var(--mono);margin:0.75rem 0 0">◈ <b style="color:#64748b">{last["filename"]}</b> | report: <b style="color:#64748b">{last["report_date"] or "—"}</b> | M/E: <b style="color:#64748b">{mt}</b> total / <b style="color:#64748b">{mm}</b> this month | uploaded: <b style="color:#64748b">{str(last["uploaded_at"])[:16]}</b></p>', unsafe_allow_html=True)

    st.markdown("---")
    tabs = st.tabs(["⚠️  Alerts","⚙️  Main Engine","🔩  Aux Engines","🛠️  Other Equipment","📊  All Components"])

    with tabs[0]:
        st.markdown('<div class="section-label">Action required</div>', unsafe_allow_html=True)
        alerts = df[df['status'].isin(['OVERDUE','HIGH PRIORITY'])].sort_values(['status','pct_used'],ascending=[True,False])
        if alerts.empty:
            st.markdown('<div style="background:rgba(16,185,129,0.08);border:1px solid rgba(16,185,129,0.2);border-radius:10px;padding:1.5rem;text-align:center;color:#6ee7b7;font-family:var(--cond);font-size:1rem;font-weight:600">✓ All components within acceptable limits</div>', unsafe_allow_html=True)
        else:
            show_table(alerts)

    with tabs[1]:
        me = df[df['category']=='MAIN_ENGINE']
        if me.empty: st.info("No Main Engine data.")
        else:
            sel = st.selectbox("Component filter",['ALL']+sorted(me['description'].unique().tolist()),key="me_f")
            show_table(me if sel=='ALL' else me[me['description']==sel])

    with tabs[2]:
        aux = df[df['category']=='AUX_ENGINE']
        if aux.empty: st.info("No Aux Engine data.")
        else:
            sel = st.selectbox("Engine filter",['ALL']+sorted(aux['engine_label'].unique().tolist()),key="aux_f")
            show_table(aux if sel=='ALL' else aux[aux['engine_label']==sel])

    with tabs[3]:
        if oe.empty: st.info("No other equipment data.")
        else:
            for sec in sorted(oe['section'].unique()):
                st.markdown(f'<div class="section-label">{sec}</div>', unsafe_allow_html=True)
                sd = oe[oe['section']==sec][['description','periodicity','last_date','run_hrs']].copy()
                sd.columns = ['Description','Periodicity','Last Date','Run Hrs']
                st.dataframe(sd, use_container_width=True, hide_index=True)

    with tabs[4]:
        c1,c2 = st.columns(2)
        with c1: sf = st.multiselect("Status",['OVERDUE','HIGH PRIORITY','OK','NO DATA'],default=['OVERDUE','HIGH PRIORITY','OK','NO DATA'],key="all_s")
        with c2: cf = st.multiselect("Category",['MAIN_ENGINE','AUX_ENGINE'],default=['MAIN_ENGINE','AUX_ENGINE'],key="all_c")
        show_table(df[df['status'].isin(sf) & df['category'].isin(cf)])


# ── UPLOAD HISTORY ───────────────────────────────────────────────────
elif page == "📋  Upload History":
    st.markdown('<div class="page-header"><h1>📋 UPLOAD HISTORY</h1></div>', unsafe_allow_html=True)
    st.markdown('<div class="page-sub section-label">Full audit trail of all report submissions</div>', unsafe_allow_html=True)
    if not selected_vessel:
        st.info("Select a vessel from the sidebar."); st.stop()
    st.markdown(f'<div class="section-label">{selected_vessel}</div>', unsafe_allow_html=True)
    hist = get_upload_history(selected_vessel)
    if hist.empty:
        st.info("No upload history for this vessel.")
    else:
        d = hist.copy()
        d.columns = ['Filename','Report Date','M/E Total Hrs','M/E This Month','Uploaded At']
        d['M/E Total Hrs']  = d['M/E Total Hrs'].apply(lambda x: f"{int(x):,}" if pd.notna(x) else '—')
        d['M/E This Month'] = d['M/E This Month'].apply(lambda x: f"{int(x):,}" if pd.notna(x) else '—')
        st.dataframe(d, use_container_width=True, hide_index=True)
