import streamlit as st
import os
import re
import sqlite3
import tempfile
import hashlib
import subprocess
import shutil
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional
import pandas as pd
from docx import Document

st.set_page_config(
    page_title="Fleet Running Hours",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# PREMIUM UI LAYER
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
  --bg:#03060d; --bg1:#06091a; --bg2:#080e20; --bg3:#0b1228; --bg4:#0f1830;
  --b1:#0f1c35; --b2:#182840; --b3:#223350;
  --gold:#c89a14; --gold2:#e0b422; --gold3:#f5cc44;
  --red:#cc2828; --red2:#ff5c5c;
  --orange:#b85518; --ora2:#ff8833;
  --green:#0d8a4a; --grn2:#22c55e; --grn3:#6ee7b7;
  --blue:#1444a8; --blu2:#3b82f6;
  --t0:#f2f7ff; --t1:#c0d0e8; --t2:#7e95b6; --t3:#425372;
  --ff:'Space Grotesk',sans-serif;
  --fi:'Inter',sans-serif;
  --fm:'JetBrains Mono',monospace;
}

html, body, [class*="css"] {
  background: var(--bg) !important;
  color: var(--t1) !important;
  font-family: var(--fi) !important;
  -webkit-font-smoothing: antialiased;
}
.main, .main > div { background: var(--bg) !important; }
.block-container { max-width: 100% !important; padding: 1.8rem 2rem 4rem !important; }

.main::before{
  content:"";
  position:fixed; inset:0; pointer-events:none; z-index:0;
  background:
    radial-gradient(ellipse 90% 50% at -10% -5%, rgba(200,154,20,.06) 0%, transparent 55%),
    radial-gradient(ellipse 70% 45% at 110% 105%, rgba(20,68,168,.05) 0%, transparent 55%);
}

[data-testid="stSidebar"]{
  background:var(--bg1)!important;
  border-right:1px solid var(--b2)!important;
}
[data-testid="stSidebar"] * { color: var(--t1)!important; }
[data-testid="stSidebarContent"] { padding:1.25rem!important; }

h1,h2,h3 {
  font-family:var(--ff)!important;
  color:var(--t0)!important;
  letter-spacing:-.02em!important;
}
h1 { font-size:1.8rem!important; font-weight:700!important; }
h2 { font-size:1.25rem!important; font-weight:600!important; }
h3 { font-size:1rem!important; font-weight:600!important; }

.stButton > button {
  background: linear-gradient(135deg, var(--gold) 0%, #8a6a08 100%) !important;
  color:#000!important; border:none!important;
  border-radius:8px!important; padding:.7rem 1.25rem!important;
  font-family:var(--ff)!important; font-weight:700!important;
  letter-spacing:.05em!important; text-transform:uppercase!important;
}
.stButton > button:hover { transform:translateY(-1px); }

[data-testid="stMetric"]{
  background: var(--bg3)!important;
  border:1px solid var(--b2)!important;
  border-radius:12px!important;
  padding:.9rem 1rem 1rem!important;
}
[data-testid="stMetricValue"]{
  font-family:var(--ff)!important;
  color:var(--t0)!important;
  font-weight:700!important;
}
[data-testid="stMetricLabel"]{
  color:var(--t3)!important;
  text-transform:uppercase!important;
  letter-spacing:.12em!important;
  font-size:.66rem!important;
}

[data-testid="stDataFrame"]{
  border:1px solid var(--b2)!important;
  border-radius:12px!important;
  overflow:hidden!important;
}

.stTabs [data-baseweb="tab-list"]{
  background:var(--bg2)!important;
  border:1px solid var(--b2)!important;
  border-radius:12px 12px 0 0!important;
  gap:0!important;
}
.stTabs [data-baseweb="tab"]{
  color:var(--t3)!important;
  font-family:var(--ff)!important;
  text-transform:uppercase!important;
  letter-spacing:.05em!important;
  font-size:.74rem!important;
}
.stTabs [aria-selected="true"]{
  color:var(--gold2)!important;
  border-bottom:2px solid var(--gold)!important;
}
.stTabs [data-baseweb="tab-panel"]{
  background:var(--bg2)!important;
  border:1px solid var(--b2)!important;
  border-top:none!important;
  border-radius:0 0 12px 12px!important;
  padding:1.25rem!important;
}
</style>
""", unsafe_allow_html=True)

def kpi(val, lbl, color="gold"):
    return f'''
    <div style="background:var(--bg3);border:1px solid var(--b2);border-radius:12px;padding:1rem 1.1rem">
      <div style="font-family:var(--ff);font-size:2rem;font-weight:700;color:var(--t0);line-height:1">{val}</div>
      <div style="color:var(--t3);text-transform:uppercase;letter-spacing:.15em;font-size:.62rem;margin-top:.4rem">{lbl}</div>
    </div>
    '''

# ============================================================
# DB LAYER
# ============================================================
DB_PATH = Path("running_hours.db")

def get_db():
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.execute("PRAGMA journal_mode=WAL")
    c.executescript("""
    CREATE TABLE IF NOT EXISTS vessels(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        created_at TEXT NOT NULL DEFAULT(datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS upload_log(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL,
        filename TEXT NOT NULL,
        file_hash TEXT NOT NULL,
        report_date TEXT,
        me_total_hrs INTEGER,
        me_this_month INTEGER,
        parse_confidence REAL DEFAULT 0,
        component_count INTEGER DEFAULT 0,
        warning_count INTEGER DEFAULT 0,
        uploaded_at TEXT NOT NULL DEFAULT(datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS components(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL,
        category TEXT NOT NULL,
        engine_label TEXT NOT NULL,
        unit TEXT NOT NULL,
        description TEXT NOT NULL,
        periodicity REAL,
        last_oh_date TEXT,
        last_oh_hrs REAL,
        hrs_since REAL,
        pct_used REAL,
        status TEXT NOT NULL,
        updated_at TEXT NOT NULL DEFAULT(datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS other_equipment(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL,
        section TEXT NOT NULL,
        description TEXT NOT NULL,
        periodicity TEXT,
        last_date TEXT,
        run_hrs TEXT,
        updated_at TEXT NOT NULL DEFAULT(datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_cv ON components(vessel_name);
    CREATE INDEX IF NOT EXISTS idx_cs ON components(status);
    """)
    conn.commit()
    conn.close()

init_db()

# ============================================================
# HELPERS
# ============================================================
HEADER_PATTERNS = [
    "1-DATE OF LAST O/H 2- RUNNING HOURS SINCE LAST O/H",
    "1-DATE OF LAST O/H 2- RUNNING HOURS SINCE LAST O/ H",
    "DATE OF LAST O/H",
    "RUNNING HOURS SINCE LAST O/H",
    "TOTAL RUNNING HOURS",
    "THIS MONTH",
    "MAIN ENGINE",
    "TYPE: MAN B&W",
    "MAN B&W –5S60MC",
]

def clean_text(x: Any) -> str:
    if x is None:
        return ""
    return re.sub(r"\s+", " ", str(x).replace("\xa0", " ")).strip()

def is_header_like(text: str) -> bool:
    t = clean_text(text).upper()
    if not t:
        return False
    for pat in HEADER_PATTERNS:
        if pat.upper() in t:
            return True
    # Long all‑caps with digits & punctuation starting with "1-"
    if re.match(r"^1[- ]?DATE OF LAST", t):
        return True
    return False

def looks_like_name(x: str) -> bool:
    t = clean_text(x)
    if not t or len(t) < 3:
        return False
    u = t.upper()
    if is_header_like(u):
        return False
    # Pure numbers / dates / hours only → not a component name
    if re.fullmatch(r"[\d./ ,:-]+", u):
        return False
    # Reject generic labels
    if u in {"DATE", "RUN HRS", "RUN HOURS", "PERIODICITY"}:
        return False
    return True

def safe_float(x: Any) -> Optional[float]:
    s = clean_text(x).replace(",", "")
    if not s:
        return None
    m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group())
    except:
        return None

def parse_date(raw: Any) -> Optional[str]:
    raw = clean_text(raw)
    if not raw or raw.upper() in {"NA", "N/A", "-", "--"}:
        return None
    raw = re.sub(r"SEPT", "SEP", raw, flags=re.I)
    for fmt in [
        "%d %b %y", "%d %B %y", "%d %b %Y", "%d %B %Y",
        "%d/%m/%y", "%d/%m/%Y", "%d-%m-%y", "%d-%m-%Y",
        "%b %Y", "%B %Y", "%Y-%m-%d"
    ]:
        for v in (raw, raw.upper(), raw.title()):
            try:
                return datetime.strptime(v, fmt).strftime("%Y-%m-%d")
            except:
                pass
    return raw

def parse_periodicity(raw: Any) -> Optional[float]:
    s = clean_text(raw).upper()
    s = s.replace("O/H", "").replace("HRS", "").replace("HRS.", "").replace("HR", "")
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    return float(m.group(1)) if m else None

def extract_number(raw: Any) -> Optional[float]:
    s = clean_text(raw).replace(",", "")
    m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None

def status_from(hrs_since: Optional[float], periodicity: Optional[float]) -> str:
    if hrs_since is None or periodicity is None or periodicity <= 0:
        return "NO DATA"
    r = hrs_since / periodicity
    if r >= 1.0:
        return "OVERDUE"
    if r >= 0.80:
        return "HIGH PRIORITY"
    return "OK"

def pct_used(hrs_since: Optional[float], periodicity: Optional[float]) -> float:
    if hrs_since is None or periodicity is None or periodicity <= 0:
        return 0.0
    return round(hrs_since / periodicity, 4)

def cyl_sort(unit: Any) -> int:
    m = re.search(r"(\d+)", str(unit))
    return int(m.group(1)) if m else 999

def engine_sort_key(engine_label: str) -> tuple:
    # MAIN_ENGINE / ME first, then AUX‑n
    lab = (engine_label or "").upper()
    if lab in {"ME", "MAIN ENGINE", "MAIN_ENGINE"}:
        return (0, 0)
    m = re.search(r"AUX[-\s]*(\d+)", lab)
    if m:
        return (1, int(m.group(1)))
    return (2, 999)

# ============================================================
# DOC CONVERSION
# ============================================================
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found. Add 'libreoffice' in packages.")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as t:
        t.write(raw)
        src = t.name
    outdir = tempfile.mkdtemp(prefix="docconv_")
    outdocx = os.path.join(outdir, Path(src).stem + ".docx")
    profile = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        res = subprocess.run(
            [soffice, "--headless", "--norestore", "--nofirststartwizard",
             f"-env:UserInstallation={profile}", "--convert-to", "docx", src, "--outdir", outdir],
            capture_output=True, timeout=120
        )
        if not os.path.exists(outdocx):
            raise RuntimeError(res.stderr.decode("utf-8", errors="ignore")[:400])
        with open(outdocx, "rb") as f:
            return f.read()
    finally:
        for p in [src, outdocx]:
            try:
                if os.path.exists(p):
                    os.unlink(p)
            except:
                pass
        shutil.rmtree(outdir, ignore_errors=True)

def open_docx_bytes(docx_bytes: bytes) -> Document:
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as t:
        t.write(docx_bytes)
        p = t.name
    try:
        return Document(p)
    finally:
        try:
            os.unlink(p)
        except:
            pass

def table_grid(table) -> List[List[str]]:
    return [[clean_text(c.text) for c in row.cells] for row in table.rows]

def row_text(row: List[str]) -> str:
    return " | ".join([clean_text(x) for x in row if clean_text(x)])

def doc_all_text(doc: Document) -> str:
    parts = [clean_text(p.text) for p in doc.paragraphs if clean_text(p.text)]
    for table in doc.tables:
        for row in table_grid(table):
            txt = row_text(row)
            if txt:
                parts.append(txt)
    return "\n".join(parts)

def detect_vessel_name(doc: Document) -> str:
    pats = [
        r"VESSEL\s*NAME[:\s]+(?:M/V\s*)?([A-Z0-9 \-_/]{2,})",
        r"\bM/?V\s+([A-Z0-9 \-_/]{2,})",
        r"NAME OF VESSEL[:\s]+([A-Z0-9 \-_/]{2,})",
    ]
    for p in doc.paragraphs:
        txt = clean_text(p.text)
        for pat in pats:
            m = re.search(pat, txt, flags=re.I)
            if m:
                return clean_text(m.group(1))
    alltxt = doc_all_text(doc).upper()
    for pat in pats:
        m = re.search(pat, alltxt, flags=re.I)
        if m:
            return clean_text(m.group(1))
    return "UNKNOWN"

def detect_report_date(doc: Document) -> Optional[str]:
    for p in doc.paragraphs:
        txt = clean_text(p.text)
        m = re.search(r"(?:REPORT\s*)?DATE[:\s]+(.+)", txt, flags=re.I)
        if m:
            return parse_date(m.group(1))
    return None

def detect_me_totals(doc: Document):
    txt = doc_all_text(doc)
    total = None
    month = None
    m1 = re.search(r"TOTAL\s+RUNNING\s+HOURS[^0-9]{0,20}([\d,]+)", txt, flags=re.I)
    m2 = re.search(r"THIS\s+MONTH[^0-9]{0,20}([\d,]+)", txt, flags=re.I)
    if m1:
        total = int(m1.group(1).replace(",", ""))
    if m2:
        month = int(m2.group(1).replace(",", ""))
    return total, month

# ============================================================
# PARSING CORE
# ============================================================
def parse_main_engine(table) -> List[Dict[str, Any]]:
    grid = table_grid(table)
    comps: List[Dict[str, Any]] = []
    if not grid:
        return comps
    max_cols = max(len(r) for r in grid)

    # Discover cylinder columns by header text
    cyl_cols = []
    for ci in range(max_cols):
        txt = " ".join(r[ci] if ci < len(r) else "" for r in grid[:4])
        m = re.search(r"CYL(?:INDER)?\s*NO\.?\s*(\d+)", txt, flags=re.I)
        if m:
            cyl_cols.append((ci, f"Cyl {m.group(1)}"))
    if not cyl_cols:
        cyl_cols = [(i, f"Cyl {i-1}") for i in range(2, min(max_cols, 14))]

    i = 0
    while i < len(grid) - 1:
        r1 = grid[i]
        r2 = grid[i + 1]
        name = r1[0] if r1 else ""
        if not looks_like_name(name):
            i += 1
            continue
        periodicity = parse_periodicity(r1[1] if len(r1) > 1 else None)
        for ci, unit in cyl_cols:
            d = parse_date(r1[ci] if ci < len(r1) else None)
            h = extract_number(r2[ci] if ci < len(r2) else None)
            # Accept rows where either date OR hours exists
            if d is None and h is None:
                continue
            comps.append({
                "category": "MAIN_ENGINE",
                "engine_label": "ME",
                "unit": unit,
                "description": clean_text(name),
                "periodicity": periodicity,
                "last_oh_date": d,
                "last_oh_hrs": h,
                "hrs_since": h,
                "pct_used": pct_used(h, periodicity),
                "status": status_from(h, periodicity),
            })
        i += 2
    return comps

def parse_aux_engine(table) -> List[Dict[str, Any]]:
    grid = table_grid(table)
    comps: List[Dict[str, Any]] = []
    if not grid:
        return comps

    header = grid[:5]
    blocks = []
    seen = set()
    for row in header:
        for ci, cell in enumerate(row):
            m = re.search(r"AUX(?:ILIARY)?\s*ENGINE\s*NO\.?\s*(\d+)", cell, flags=re.I)
            if m:
                label = f"AUX-{m.group(1)}"
                if label not in seen:
                    seen.add(label)
                    blocks.append((label, ci))
    if not blocks:
        blocks = [("AUX-1", 2), ("AUX-2", 6), ("AUX-3", 10)]

    cyl_map = {}
    if len(grid) > 4:
        hdr_row = grid[4]
        for idx, (label, start) in enumerate(blocks):
            end = blocks[idx + 1][1] if idx + 1 < len(blocks) else len(hdr_row)
            for ci in range(start, end):
                if ci < len(hdr_row):
                    m = re.search(r"\b(\d{1,2})\b", hdr_row[ci])
                    if m:
                        cyl_map[ci] = (label, f"Cyl {m.group(1)}")

    i = 5
    while i < len(grid) - 1:
        r1 = grid[i]
        r2 = grid[i + 1]
        name = r1[0] if r1 else ""
        if not looks_like_name(name):
            i += 1
            continue
        periodicity = parse_periodicity(r1[1] if len(r1) > 1 else None)
        for ci, (eng, unit) in cyl_map.items():
            d = parse_date(r1[ci] if ci < len(r1) else None)
            h = extract_number(r2[ci] if ci < len(r2) else None)
            if d is None and h is None:
                continue
            comps.append({
                "category": "AUX_ENGINE",
                "engine_label": eng,
                "unit": unit,
                "description": clean_text(name),
                "periodicity": periodicity,
                "last_oh_date": d,
                "last_oh_hrs": h,
                "hrs_since": h,
                "pct_used": pct_used(h, periodicity),
                "status": status_from(h, periodicity),
            })
        i += 2
    return comps

def parse_other_equipment(doc: Document) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for table in doc.tables:
        for row in table_grid(table):
            joined = clean_text(" ".join(row)).upper()
            if any(x in joined for x in ["TURBOCHARGER", "COOLERS", "COMPRESSORS", "BOILER", "D/G"]):
                desc = next((c for c in row if looks_like_name(c)), "")
                if desc:
                    rows.append({
                        "section": "OTHER EQUIPMENT",
                        "description": desc,
                        "periodicity": "",
                        "last_date": "",
                        "run_hrs": "",
                    })
    # Deduplicate
    out = []
    seen = set()
    for r in rows:
        key = (r["section"], r["description"])
        if key not in seen:
            seen.add(key)
            out.append(r)
    return out

def dedupe_components(comps: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    out = []
    seen = set()
    for x in comps:
        key = (
            x["category"],
            x["engine_label"],
            x["unit"],
            x["description"],
            x.get("last_oh_date"),
            x.get("hrs_since"),
        )
        if key not in seen:
            seen.add(key)
            out.append(x)
    return out

def filter_header_noise(comps: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    out = []
    for c in comps:
        desc = clean_text(c.get("description", ""))
        if is_header_like(desc):
            continue
        if c["status"] == "NO DATA":
            if not c.get("periodicity") and not c.get("last_oh_date") and c.get("hrs_since") is None:
                # Pure NO DATA with no numbers and header‑like description → drop
                continue
        out.append(c)
    return out

def parse_doc_bytes(docx_bytes: bytes) -> Dict[str, Any]:
    doc = open_docx_bytes(docx_bytes)
    warnings: List[str] = []
    vessel_name = detect_vessel_name(doc)
    report_date = detect_report_date(doc)
    me_total_hrs, me_this_month = detect_me_totals(doc)

    if vessel_name == "UNKNOWN":
        warnings.append("Could not confidently extract vessel name.")

    main_table = doc.tables[0] if len(doc.tables) > 0 else None
    aux_table = doc.tables[2] if len(doc.tables) > 2 else None

    components: List[Dict[str, Any]] = []
    if main_table:
        components.extend(parse_main_engine(main_table))
    else:
        warnings.append("No main engine table found.")
    if aux_table:
        components.extend(parse_aux_engine(aux_table))
    else:
        warnings.append("No auxiliary engine table found.")

    components = dedupe_components(components)
    components = filter_header_noise(components)

    other_equipment = parse_other_equipment(doc)

    if len(components) < 5:
        warnings.append(f"Low component count extracted ({len(components)}).")

    confidence = 0.0
    if vessel_name != "UNKNOWN":
        confidence += 0.2
    if report_date:
        confidence += 0.1
    if len(components) >= 8:
        confidence += 0.3
    if any(c["category"] == "MAIN_ENGINE" for c in components):
        confidence += 0.2
    if any(c["category"] == "AUX_ENGINE" for c in components):
        confidence += 0.15
    if not warnings:
        confidence += 0.05
    confidence = round(min(confidence, 1.0), 2)

    return {
        "vessel_name": vessel_name,
        "report_date": report_date,
        "me_total_hrs": me_total_hrs,
        "me_this_month": me_this_month,
        "components": components,
        "other_equipment": other_equipment,
        "warnings": warnings,
        "parse_confidence": confidence,
    }

# ============================================================
# SAVE
# ============================================================
def save_parsed(parsed: Dict[str, Any], filename: str, file_hash: str):
    if parsed["parse_confidence"] < 0.35 or len(parsed["components"]) < 5:
        raise ValueError("Parse confidence too low to commit safely.")
    conn = get_db()
    c = conn.cursor()
    now = datetime.utcnow().isoformat() + "Z"
    vessel = parsed["vessel_name"]
    try:
        c.execute("INSERT OR IGNORE INTO vessels(name, created_at) VALUES (?, ?)", (vessel, now))
        c.execute("""
            INSERT INTO upload_log(
                vessel_name, filename, file_hash, report_date, me_total_hrs,
                me_this_month, parse_confidence, component_count, warning_count, uploaded_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            vessel, filename, file_hash, parsed["report_date"], parsed["me_total_hrs"],
            parsed["me_this_month"], parsed["parse_confidence"],
            len(parsed["components"]), len(parsed["warnings"]), now
        ))
        c.execute("DELETE FROM components WHERE vessel_name = ?", (vessel,))
        for x in parsed["components"]:
            c.execute("""
                INSERT INTO components(
                    vessel_name, category, engine_label, unit, description,
                    periodicity, last_oh_date, last_oh_hrs, hrs_since, pct_used, status, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                vessel, x["category"], x["engine_label"], x["unit"], x["description"],
                x["periodicity"], x["last_oh_date"], x["last_oh_hrs"], x["hrs_since"],
                x["pct_used"], x["status"], now
            ))
        c.execute("DELETE FROM other_equipment WHERE vessel_name = ?", (vessel,))
        for x in parsed["other_equipment"]:
            c.execute("""
                INSERT INTO other_equipment(
                    vessel_name, section, description, periodicity, last_date, run_hrs, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                vessel, x["section"], x["description"], x["periodicity"],
                x["last_date"], x["run_hrs"], now
            ))
        conn.commit()
    except:
        conn.rollback()
        raise
    finally:
        conn.close()

# ============================================================
# QUERIES
# ============================================================
@st.cache_data(ttl=10)
def get_vessels() -> List[str]:
    conn = get_db()
    rows = conn.execute("SELECT name FROM vessels ORDER BY name").fetchall()
    conn.close()
    return [r["name"] for r in rows]

@st.cache_data(ttl=10)
def get_components(vessel: str) -> pd.DataFrame:
    conn = get_db()
    df = pd.read_sql_query("SELECT * FROM components WHERE vessel_name = ?", conn, params=(vessel,))
    conn.close()
    return df

@st.cache_data(ttl=10)
def get_other_equipment(vessel: str) -> pd.DataFrame:
    conn = get_db()
    df = pd.read_sql_query(
        "SELECT * FROM other_equipment WHERE vessel_name = ? ORDER BY section, description",
        conn, params=(vessel,)
    )
    conn.close()
    return df

@st.cache_data(ttl=10)
def get_history(vessel: str) -> pd.DataFrame:
    conn = get_db()
    df = pd.read_sql_query("""
        SELECT filename, report_date, me_total_hrs, me_this_month,
               parse_confidence, component_count, warning_count, uploaded_at
        FROM upload_log
        WHERE vessel_name = ?
        ORDER BY uploaded_at DESC
        LIMIT 20
    """, conn, params=(vessel,))
    conn.close()
    return df

@st.cache_data(ttl=10)
def get_summary() -> pd.DataFrame:
    conn = get_db()
    df = pd.read_sql_query("""
        SELECT vessel_name,
               SUM(CASE WHEN status='OVERDUE' THEN 1 ELSE 0 END) AS overdue,
               SUM(CASE WHEN status='HIGH PRIORITY' THEN 1 ELSE 0 END) AS high_priority,
               SUM(CASE WHEN status='OK' THEN 1 ELSE 0 END) AS ok,
               SUM(CASE WHEN status='NO DATA' THEN 1 ELSE 0 END) AS no_data,
               COUNT(*) AS total
        FROM components
        GROUP BY vessel_name
        ORDER BY overdue DESC, high_priority DESC, vessel_name ASC
    """, conn)
    conn.close()
    return df

@st.cache_data(ttl=10)
def get_all_fleet_components() -> pd.DataFrame:
    conn = get_db()
    df = pd.read_sql_query("SELECT * FROM components", conn)
    conn.close()
    return df

# ============================================================
# DISPLAY TABLES (WITH ORDERING ME→AUX, CYL 1..N)
# ============================================================
STATUS_THEME = {
    "OVERDUE": {"bg":"#2d0707", "status":"#ff6b6b", "main":"#ff8080", "accent":"#ff3333", "dim":"#773333"},
    "HIGH PRIORITY": {"bg":"#2d1503", "status":"#ffaa44", "main":"#ff9933", "accent":"#ffcc00", "dim":"#774422"},
    "OK": {"bg":"#042010", "status":"#4ade80", "main":"#22c55e", "accent":"#4ade80", "dim":"#0f4023"},
    "NO DATA": {"bg":"#0c1422", "status":"#7da3d8", "main":"#5f7fa6", "accent":"#7da3d8", "dim":"#2a3950"},
}

def build_display_df(df: pd.DataFrame, priority: bool=False) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=[
            "Status","Vessel","Component","Engine","Unit","Periodicity","Last O/H","Hrs Since","Used"
        ])
    d = df.copy()

    # Custom ordering: category/engine/cylinder/description
    if priority:
        order_map = {"OVERDUE":0, "HIGH PRIORITY":1, "OK":2, "NO DATA":3}
        d["_ord"] = d["status"].map(lambda x: order_map.get(str(x), 9))
        d["_pct"] = d["pct_used"].fillna(0)
    d["_cat_rank"] = d["category"].map(
        lambda c: 0 if str(c).upper().startswith("MAIN") else (1 if str(c).upper().startswith("AUX") else 2)
    )
    d["_eng_rank"] = d["engine_label"].map(lambda e: engine_sort_key(str(e))[1])
    d["_eng_group"] = d["engine_label"].map(lambda e: engine_sort_key(str(e))[0])
    d["_cyl"] = d["unit"].map(cyl_sort)
    d["_desc"] = d["description"].astype(str).str.upper()

    sort_cols = ["_cat_rank", "_eng_group", "_eng_rank", "_cyl", "_desc"]
    asc = [True, True, True, True, True]
    if priority:
        sort_cols = ["_ord", "_pct"] + sort_cols
        asc = [True, False] + asc
    if "vessel_name" in d.columns:
        sort_cols = ["vessel_name"] + sort_cols
        asc = [True] + asc

    d = d.sort_values(sort_cols, ascending=asc)

    keep = ["status","vessel_name","description","engine_label","unit",
            "periodicity","last_oh_date","hrs_since","pct_used"]
    d = d[keep]

    out = pd.DataFrame()
    out["Status"] = d["status"]
    out["Vessel"] = d["vessel_name"] if "vessel_name" in d.columns else ""
    out["Component"] = d["description"]
    out["Engine"] = d["engine_label"]
    out["Unit"] = d["unit"]
    out["Periodicity"] = d["periodicity"].apply(lambda x: int(x) if pd.notna(x) else None)
    out["Last O/H"] = d["last_oh_date"].fillna("")
    out["Hrs Since"] = d["hrs_since"].apply(lambda x: int(x) if pd.notna(x) else None)
    out["Used"] = d["pct_used"].apply(lambda x: round(float(x) * 100, 1) if pd.notna(x) else 0.0)
    return out

def apply_style(df: pd.DataFrame):
    def style_row(row):
        s = STATUS_THEME.get(str(row["Status"]), STATUS_THEME["NO DATA"])
        return [
            f"background-color:{s['bg']}; color:{s['status']}; font-weight:700",
            f"background-color:{s['bg']}; color:{s['main']}; font-weight:600",
            f"background-color:{s['bg']}; color:{s['main']}",
            f"background-color:{s['bg']}; color:{s['dim']}",
            f"background-color:{s['bg']}; color:{s['dim']}",
            f"background-color:{s['bg']}; color:{s['dim']}",
            f"background-color:{s['bg']}; color:{s['main']}; font-weight:600",
            f"background-color:{s['bg']}; color:{s['accent']}; font-weight:700",
            f"background-color:{s['bg']}; color:{s['accent']}; font-weight:700",
        ]
    return df.style.apply(style_row, axis=1)

COLCFG = {
    "Status": st.column_config.TextColumn("Status", width=130),
    "Vessel": st.column_config.TextColumn("Vessel", width=130),
    "Component": st.column_config.TextColumn("Component", width=230),
    "Engine": st.column_config.TextColumn("Engine", width=90),
    "Unit": st.column_config.TextColumn("Unit", width=80),
    "Periodicity": st.column_config.NumberColumn("Periodicity", format="%d", width=105),
    "Last O/H": st.column_config.TextColumn("Last O/H", width=110),
    "Hrs Since": st.column_config.NumberColumn("Hrs Since", format="%d hrs", width=110),
    "Used": st.column_config.ProgressColumn("Used", min_value=0, max_value=160, format="%.1f%%", width=120),
}

def render_table(df: pd.DataFrame, height: Optional[int]=None, priority: bool=False):
    if isinstance(df, list):
        df = pd.DataFrame(df)
    if df.empty:
        st.info("No data to display.")
        return
    tbl = build_display_df(df, priority=priority)
    h = height or min(900, 38 * len(tbl) + 44)
    st.dataframe(
        apply_style(tbl),
        use_container_width=True,
        hide_index=True,
        height=h,
        column_config=COLCFG
    )

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("<h2>FLEET MONITOR</h2>", unsafe_allow_html=True)
    page = st.selectbox(
        "Navigation",
        ["Fleet Overview", "Vessel Detail", "Upload Report", "Upload History"],
        label_visibility="collapsed",
    )
    vessels = get_vessels()
    selected_vessel = st.selectbox(
        "Active Vessel", vessels if vessels else ["—"], disabled=not bool(vessels)
    )
    st.markdown("<hr>", unsafe_allow_html=True)
    db_kb = DB_PATH.stat().st_size / 1024 if DB_PATH.exists() else 0
    st.write(f"DB: {DB_PATH.name} · {db_kb:.0f} kB · {len(vessels)} vessels")

# ============================================================
# PAGES
# ============================================================
if page == "Fleet Overview":
    st.markdown("<h1>Fleet Master Matrix</h1>", unsafe_allow_html=True)
    smry = get_summary()
    all_comps = get_all_fleet_components()
    if smry.empty or all_comps.empty:
        st.info("No data loaded. Upload a report to begin.")
        st.stop()

    cols = st.columns(5)
    metrics = [
        (len(smry), "Vessels", "blue"),
        (len(all_comps), "Components", "gold"),
        (int((all_comps["status"] == "OVERDUE").sum()), "Overdue", "red"),
        (int((all_comps["status"] == "HIGH PRIORITY").sum()), "High Priority", "orange"),
        (int((all_comps["status"] == "OK").sum()), "OK", "green"),
    ]
    for c, (v, l, cl) in zip(cols, metrics):
        with c:
            st.markdown(kpi(v, l, cl), unsafe_allow_html=True)

    vessel_f = st.selectbox("Filter Vessel", ["All Fleet"] + sorted(all_comps["vessel_name"].unique().tolist()))
    cat_f = st.selectbox("Filter Machinery Type", ["All", "Main Engine", "Aux Engines"])
    status_f = st.selectbox("Filter Status", ["All", "Critical Focus", "Overdue Only", "High Priority Only", "OK Only", "No Data Only"])
    comp_f = st.selectbox("Filter Component", ["All"] + sorted(all_comps["description"].unique().tolist()))

    filt = all_comps.copy()
    if vessel_f != "All Fleet":
        filt = filt[filt["vessel_name"] == vessel_f]
    if cat_f == "Main Engine":
        filt = filt[filt["category"] == "MAIN_ENGINE"]
    elif cat_f == "Aux Engines":
        filt = filt[filt["category"] == "AUX_ENGINE"]

    if status_f == "Critical Focus":
        filt = filt[filt["status"].isin(["OVERDUE", "HIGH PRIORITY"])]
    elif status_f == "Overdue Only":
        filt = filt[filt["status"] == "OVERDUE"]
    elif status_f == "High Priority Only":
        filt = filt[filt["status"] == "HIGH PRIORITY"]
    elif status_f == "OK Only":
        filt = filt[filt["status"] == "OK"]
    elif status_f == "No Data Only":
        filt = filt[filt["status"] == "NO DATA"]

    if comp_f != "All":
        filt = filt[filt["description"] == comp_f]

    render_table(filt, priority=True)

elif page == "Vessel Detail":
    if not vessels or selected_vessel == "—":
        st.info("Select a vessel from the sidebar.")
        st.stop()
    st.markdown(f"<h1>{selected_vessel}</h1>", unsafe_allow_html=True)
    df = get_components(selected_vessel)
    oe = get_other_equipment(selected_vessel)
    if df.empty:
        st.info("No data for this vessel.")
        st.stop()

    cols = st.columns(5)
    metrics = [
        (len(df), "Total", "gold"),
        (int((df["status"] == "OVERDUE").sum()), "Overdue", "red"),
        (int((df["status"] == "HIGH PRIORITY").sum()), "High Priority", "orange"),
        (int((df["status"] == "OK").sum()), "OK", "green"),
        (int((df["status"] == "NO DATA").sum()), "No Data", "blue"),
    ]
    for c, (v, l, cl) in zip(cols, metrics):
        with c:
            st.markdown(kpi(v, l, cl), unsafe_allow_html=True)

    hist = get_history(selected_vessel)
    if not hist.empty:
        last = hist.iloc[0]
        st.write(f"Latest file: {last['filename']} · Confidence {float(last['parse_confidence']):.2f}")

    tabs = st.tabs(["Alerts", "Main Engine", "Aux Engines", "Other Equipment"])
    with tabs[0]:
        alerts = df[df["status"].isin(["OVERDUE", "HIGH PRIORITY"])]
        render_table(alerts, priority=True)
    with tabs[1]:
        me = df[df["category"] == "MAIN_ENGINE"]
        render_table(me)
    with tabs[2]:
        aux = df[df["category"] == "AUX_ENGINE"]
        render_table(aux)
    with tabs[3]:
        if oe.empty:
            st.info("No other equipment data.")
        else:
            st.dataframe(oe, use_container_width=True, hide_index=True)

elif page == "Upload Report":
    st.markdown("<h1>Upload Report</h1>", unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload TEC-004 .doc", type=["doc"], label_visibility="collapsed")
    if uploaded:
        raw = uploaded.read()
        file_hash = hashlib.md5(raw).hexdigest()
        with st.spinner("Converting and parsing..."):
            docx_bytes = convert_doc_to_docx(raw)
            parsed = parse_doc_bytes(docx_bytes)

        st.write(f"Vessel: {parsed['vessel_name']}")
        st.write(f"Report date: {parsed['report_date'] or '—'}")
        st.write(f"Confidence: {parsed['parse_confidence']:.2f}")
        st.write(f"Components: {len(parsed['components'])}")
        if parsed["warnings"]:
            for w in parsed["warnings"]:
                st.warning(w)

        pre = pd.DataFrame(parsed["components"])
        if not pre.empty:
            st.subheader("Preview")
            render_table(pre, priority=True)

        if st.button("Commit to Database"):
            try:
                save_parsed(parsed, uploaded.name, file_hash)
                for fn in [get_vessels, get_components, get_other_equipment, get_history, get_summary, get_all_fleet_components]:
                    fn.clear()
                st.success("Saved successfully.")
            except Exception as e:
                st.error(str(e))

elif page == "Upload History":
    st.markdown("<h1>Upload History</h1>", unsafe_allow_html=True)
    if not vessels or selected_vessel == "—":
        st.info("Select a vessel from the sidebar.")
        st.stop()
    hist = get_history(selected_vessel)
    if hist.empty:
        st.info("No history available.")
    else:
        st.dataframe(hist, use_container_width=True, hide_index=True)
