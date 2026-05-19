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
from typing import Optional, List, Dict, Tuple

import pandas as pd
from docx import Document

st.set_page_config(
    page_title="Fleet Running Hours",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# PREMIUM UI
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

h1,h2,h3 { font-family:var(--ff)!important; color:var(--t0)!important; letter-spacing:-.02em!important; }
h1 { font-size:1.8rem!important; font-weight:700!important; }
h2 { font-size:1.25rem!important; font-weight:600!important; }
h3 { font-size:1rem!important; font-weight:600!important; }

.stSelectbox > div > div,
.stMultiSelect > div > div,
.stTextInput > div > div > input {
  background: var(--bg3)!important;
  border: 1px solid var(--b2)!important;
  color: var(--t1)!important;
  border-radius: 8px!important;
}

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
  padding: .9rem 1rem 1rem!important;
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
  background:var(--bg2)!important;
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

hr { border-color: var(--b2)!important; }

.kpi {
  background: var(--bg3);
  border:1px solid var(--b2);
  border-radius:12px;
  padding:1rem 1.1rem;
}
.kpi .v {
  font-family:var(--ff);
  font-size:2rem;
  font-weight:700;
  line-height:1;
}
.kpi .l {
  color:var(--t3);
  text-transform:uppercase;
  letter-spacing:.15em;
  font-size:.62rem;
  margin-top:.4rem;
}
.kpi.gold .v{color:var(--gold3);}
.kpi.red .v{color:var(--red2);}
.kpi.orange .v{color:var(--ora2);}
.kpi.green .v{color:var(--grn2);}
.kpi.blue .v{color:var(--blu2);}

.panel {
  background:var(--bg3);
  border:1px solid var(--b2);
  border-radius:14px;
  padding:1rem 1.1rem;
}
.muted { color:var(--t2)!important; }
.smallcap {
  color:var(--t3);
  text-transform:uppercase;
  letter-spacing:.16em;
  font-size:.66rem;
  font-family:var(--fi);
}
.hero {
  margin-bottom:1rem;
}
.hero-line {
  height:1px;
  background:linear-gradient(90deg,var(--gold),var(--b2),transparent);
  margin:.5rem 0 1.2rem;
}
.code {
  font-family:var(--fm);
  color:var(--t2);
  font-size:.74rem;
}
.ok   { color:var(--grn2)!important; }
.warn { color:var(--ora2)!important; }
.bad  { color:var(--red2)!important; }

[data-testid="stFileUploader"]{
  background:linear-gradient(160deg,rgba(200,154,20,.04),rgba(20,68,168,.03))!important;
  border-radius:14px!important;
}
[data-testid="stFileUploadDropzone"]{
  border:1.5px dashed var(--gold)!important;
  border-radius:14px!important;
  padding:2.5rem 1.5rem!important;
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# DB
# ============================================================
DB_PATH = Path("runninghours.db")

def get_db():
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.execute("PRAGMA journal_mode=WAL")
    c.executescript("""
    CREATE TABLE IF NOT EXISTS vessels (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS upload_log (
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
        uploaded_at TEXT NOT NULL DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS components (
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
        updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS other_equipment (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL,
        section TEXT NOT NULL,
        description TEXT NOT NULL,
        periodicity TEXT,
        last_date TEXT,
        run_hrs TEXT,
        updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );

    CREATE INDEX IF NOT EXISTS idx_components_vessel ON components(vessel_name);
    CREATE INDEX IF NOT EXISTS idx_components_status ON components(status);
    CREATE INDEX IF NOT EXISTS idx_log_vessel ON upload_log(vessel_name);
    """)
    conn.commit()
    conn.close()

init_db()

# ============================================================
# HELPERS
# ============================================================
def clean_text(x: str) -> str:
    if x is None:
        return ""
    x = str(x).replace("\xa0", " ")
    x = re.sub(r"\s+", " ", x).strip()
    return x

def normalize_for_match(x: str) -> str:
    x = clean_text(x).upper()
    x = x.replace("O/H", "OH")
    x = x.replace("A/E", "AUX ENGINE")
    x = x.replace("M/E", "MAIN ENGINE")
    x = re.sub(r"[^A-Z0-9 /.-]", "", x)
    return x

def safe_float(x) -> Optional[float]:
    if x is None:
        return None
    s = clean_text(str(x))
    if not s or s.upper() in {"NA", "N/A", "-", "--"}:
        return None
    s = s.replace(",", "")
    m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group())
    except:
        return None

def safe_int(x) -> Optional[int]:
    v = safe_float(x)
    return int(v) if v is not None else None

def parse_date(raw: str) -> Optional[str]:
    raw = clean_text(raw)
    if not raw or raw.upper() in {"NA", "N/A", "-", "--"}:
        return None

    raw = re.sub(r"SEPT", "SEP", raw, flags=re.I)
    raw = re.sub(r"\s+", " ", raw)
    candidates = [raw, raw.upper(), raw.title()]

    fmts = [
        "%d %b %y", "%d %B %y", "%d %b %Y", "%d %B %Y",
        "%d/%m/%y", "%d/%m/%Y", "%d-%m-%y", "%d-%m-%Y",
        "%b %Y", "%B %Y", "%Y-%m-%d"
    ]
    for c in candidates:
        for fmt in fmts:
            try:
                return datetime.strptime(c, fmt).strftime("%Y-%m-%d")
            except:
                pass
    return raw

def extract_positive_number(raw: str) -> Optional[float]:
    if raw is None:
        return None
    txt = clean_text(str(raw)).replace(",", "")
    nums = re.findall(r"[-+]?\d+(?:\.\d+)?", txt)
    for n in nums:
        try:
            v = float(n)
            if v >= 0:
                return v
        except:
            pass
    return None

def parse_periodicity(raw: str) -> Optional[float]:
    if raw is None:
        return None
    s = clean_text(str(raw)).upper()
    s = s.replace("O/H", "")
    s = s.replace("HRS", "")
    s = s.replace("HRS.", "")
    s = s.replace("HR", "")
    s = s.replace(",", "")
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    if not m:
        return None
    try:
        return float(m.group(1))
    except:
        return None

def calc_pct_used(hrs_since: Optional[float], periodicity: Optional[float]) -> float:
    if hrs_since is None or periodicity is None or periodicity <= 0:
        return 0.0
    return round(hrs_since / periodicity, 4)

def derive_status(hrs_since: Optional[float], periodicity: Optional[float]) -> str:
    if hrs_since is None or periodicity is None or periodicity <= 0:
        return "NO DATA"
    ratio = hrs_since / periodicity
    if ratio >= 1.0:
        return "OVERDUE"
    if ratio >= 0.80:
        return "HIGH PRIORITY"
    return "OK"

def kpi(val, lbl, color="gold"):
    return f"""
    <div class="kpi {color}">
        <div class="v">{val}</div>
        <div class="l">{lbl}</div>
    </div>
    """

def hero(title: str, eyebrow: str = ""):
    top = f'<div class="smallcap">{eyebrow}</div>' if eyebrow else ""
    return f'<div class="hero">{top}<h1>{title}</h1><div class="hero-line"></div></div>'

# ============================================================
# DOC CONVERSION
# ============================================================
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found. Add libreoffice in packages.txt.")

    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as tmp:
        tmp.write(raw)
        doc_path = tmp.name

    out_dir = tempfile.mkdtemp(prefix="docconv_")
    out_docx = os.path.join(out_dir, Path(doc_path).stem + ".docx")
    profile = f"file:///tmp/lo_profile_{os.getpid()}_{os.urandom(4).hex()}"

    try:
        res = subprocess.run(
            [
                soffice,
                "--headless",
                "--norestore",
                "--nofirststartwizard",
                f"-env:UserInstallation={profile}",
                "--convert-to", "docx",
                doc_path,
                "--outdir", out_dir,
            ],
            capture_output=True,
            timeout=120,
        )
        if not os.path.exists(out_docx):
            stderr = res.stderr.decode("utf-8", errors="ignore")[:500]
            raise RuntimeError(f"LibreOffice conversion failed. {stderr}")
        with open(out_docx, "rb") as f:
            return f.read()
    finally:
        try:
            os.unlink(doc_path)
        except:
            pass
        try:
            if os.path.exists(out_docx):
                os.unlink(out_docx)
        except:
            pass
        try:
            shutil.rmtree(out_dir, ignore_errors=True)
        except:
            pass

# ============================================================
# DOC INSPECTION
# ============================================================
def open_docx_bytes(docx_bytes: bytes) -> Document:
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp.write(docx_bytes)
        temp_path = tmp.name
    try:
        doc = Document(temp_path)
    finally:
        try:
            os.unlink(temp_path)
        except:
            pass
    return doc

def table_to_grid(table) -> List[List[str]]:
    return [[clean_text(cell.text) for cell in row.cells] for row in table.rows]

def row_text(row: List[str]) -> str:
    return " | ".join([clean_text(x) for x in row if clean_text(x)])

def all_doc_text(doc: Document) -> str:
    parts = []
    for p in doc.paragraphs:
        t = clean_text(p.text)
        if t:
            parts.append(t)
    for table in doc.tables:
        grid = table_to_grid(table)
        for row in grid:
            txt = row_text(row)
            if txt:
                parts.append(txt)
    return "\n".join(parts)

def detect_vessel_name(doc: Document) -> str:
    patterns = [
        r"VESSEL\s*NAME[:\s]+(?:M/V\s*)?([A-Z0-9 \-_/]{2,})",
        r"\bM/?V\s+([A-Z0-9 \-_/]{2,})",
        r"NAME OF VESSEL[:\s]+([A-Z0-9 \-_/]{2,})",
    ]
    for p in doc.paragraphs:
        txt = clean_text(p.text)
        if not txt:
            continue
        u = txt.upper()
        for pat in patterns:
            m = re.search(pat, u, flags=re.I)
            if m:
                name = clean_text(m.group(1))
                name = re.sub(r"\s{2,}", " ", name)
                return name
    joined = all_doc_text(doc).upper()
    for pat in patterns:
        m = re.search(pat, joined, flags=re.I)
        if m:
            return clean_text(m.group(1))
    return "UNKNOWN"

def detect_report_date(doc: Document) -> Optional[str]:
    patterns = [
        r"DATE[:\s]+([A-Z0-9 /-]{4,20})",
        r"REPORT\s*DATE[:\s]+([A-Z0-9 /-]{4,20})",
        r"MONTH[:\s]+([A-Z0-9 /-]{4,20})",
    ]
    for p in doc.paragraphs:
        txt = clean_text(p.text)
        if not txt:
            continue
        for pat in patterns:
            m = re.search(pat, txt, flags=re.I)
            if m:
                return parse_date(m.group(1))
    return None

def detect_me_totals(doc: Document) -> Tuple[Optional[int], Optional[int]]:
    total = None
    month = None
    text = all_doc_text(doc)

    m1 = re.search(r"TOTAL\s+RUNNING\s+HOURS[^0-9]{0,20}([\d,]+)", text, flags=re.I)
    m2 = re.search(r"THIS\s+MONTH[^0-9]{0,20}([\d,]+)", text, flags=re.I)
    if m1:
        total = safe_int(m1.group(1))
    if m2:
        month = safe_int(m2.group(1))
    return total, month

# ============================================================
# ROBUST PARSER
# ============================================================
MAIN_ENGINE_KEYWORDS = [
    "MAIN ENGINE", "ME", "CYL", "CYLINDER", "RUNNING HOURS", "TOTAL RUNNING HOURS"
]
AUX_ENGINE_KEYWORDS = [
    "AUX. ENGINE", "AUX ENGINE", "A/E", "DG 1", "DG 2", "DG 3", "AUX ENGINE NO"
]
OTHER_SECTION_HINTS = [
    "TURBOCHARGER", "AUXILIARY BOILER", "COOLERS", "EXH GAS BOILER",
    "AC REFR. COMPRESSORS", "MAIN AIR COMPRESSORS", "DG EQUIPMENT"
]

def table_score(grid: List[List[str]], keywords: List[str]) -> int:
    score = 0
    for row in grid[:8]:
        txt = normalize_for_match(row_text(row))
        for kw in keywords:
            if kw in txt:
                score += 1
    return score

def find_best_tables(doc: Document):
    candidates = []
    for idx, table in enumerate(doc.tables):
        grid = table_to_grid(table)
        candidates.append({
            "idx": idx,
            "grid": grid,
            "me_score": table_score(grid, MAIN_ENGINE_KEYWORDS),
            "aux_score": table_score(grid, AUX_ENGINE_KEYWORDS),
            "other_score": table_score(grid, OTHER_SECTION_HINTS),
        })
    me_tables = sorted([x for x in candidates if x["me_score"] > 0], key=lambda x: x["me_score"], reverse=True)
    aux_tables = sorted([x for x in candidates if x["aux_score"] > 0], key=lambda x: x["aux_score"], reverse=True)
    other_tables = sorted([x for x in candidates if x["other_score"] > 0], key=lambda x: x["other_score"], reverse=True)
    return me_tables, aux_tables, other_tables, candidates

def find_cylinder_columns(header_rows: List[List[str]]) -> List[Tuple[int, str]]:
    out = []
    seen = set()
    for row in header_rows:
        for ci, cell in enumerate(row):
            txt = normalize_for_match(cell)
            m = re.search(r"CYL(?:INDER)?\s*NO\.?\s*(\d+)", txt)
            if m:
                label = f"Cyl {m.group(1)}"
                if label not in seen:
                    seen.add(label)
                    out.append((ci, label))
            elif txt in {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"}:
                label = f"Cyl {txt}"
                if label not in seen:
                    seen.add(label)
                    out.append((ci, label))
    return out

def looks_like_component_name(text: str) -> bool:
    t = normalize_for_match(text)
    if not t:
        return False
    rejects = {
        "DATE", "RUN HRS", "RUN HOURS", "PERIODICITY", "THIS MONTH",
        "TOTAL RUNNING HOURS", "LAST OH DATE", "LAST OH HOURS"
    }
    if t in rejects:
        return False
    if len(t) < 3:
        return False
    if re.fullmatch(r"[\d./ -]+", t):
        return False
    return True

def parse_main_engine_table(grid: List[List[str]]) -> List[Dict]:
    comps = []
    if not grid:
        return comps

    header_rows = grid[:4]
    cyl_cols = find_cylinder_columns(header_rows)

    if not cyl_cols:
        max_cols = max(len(r) for r in grid) if grid else 0
        for c in range(2, min(max_cols, 14)):
            cyl_cols.append((c, f"Cyl {c-1}"))

    i = 0
    while i < len(grid) - 1:
        r1 = grid[i]
        r2 = grid[i + 1] if i + 1 < len(grid) else []

        name = clean_text(r1[0]) if r1 else ""
        if not looks_like_component_name(name):
            i += 1
            continue

        periodicity = parse_periodicity(r1[1] if len(r1) > 1 else None)

        for ci, label in cyl_cols:
            d = parse_date(r1[ci] if ci < len(r1) else None)
            h = extract_positive_number(r2[ci] if ci < len(r2) else None)

            if d is None and h is None:
                continue

            comps.append({
                "category": "MAINENGINE",
                "engine_label": "ME",
                "unit": label,
                "description": name,
                "periodicity": periodicity,
                "last_oh_date": d,
                "last_oh_hrs": h,
                "hrs_since": h,
                "pct_used": calc_pct_used(h, periodicity),
                "status": derive_status(h, periodicity),
            })
        i += 2
    return comps

def detect_aux_blocks(grid: List[List[str]]) -> List[Tuple[str, int]]:
    blocks = []
    seen = set()
    for row in grid[:5]:
        for ci, cell in enumerate(row):
            txt = normalize_for_match(cell)
            m = re.search(r"AUX(?:ILIARY)?\s*ENGINE\s*NO\.?\s*(\d+)", txt)
            if m:
                label = f"AUX-{m.group(1)}"
                if label not in seen:
                    seen.add(label)
                    blocks.append((label, ci))
            m2 = re.search(r"DG\s*(\d+)", txt)
            if m2:
                label = f"AUX-{m2.group(1)}"
                if label not in seen:
                    seen.add(label)
                    blocks.append((label, ci))
    return blocks

def parse_aux_engine_table(grid: List[List[str]]) -> List[Dict]:
    comps = []
    if not grid:
        return comps

    blocks = detect_aux_blocks(grid)
    if not blocks:
        max_cols = max(len(r) for r in grid) if grid else 0
        approx = [2, 6, 10]
        blocks = [(f"AUX-{n+1}", col) for n, col in enumerate(approx) if col < max_cols]

    cyl_map = {}
    if len(grid) >= 5:
        row4 = grid[4]
        for label, start_col in blocks:
            end_col = next((c for l, c in blocks if c > start_col), len(row4))
            for ci in range(start_col, end_col):
                if ci < len(row4):
                    m = re.search(r"\b(\d{1,2})\b", clean_text(row4[ci]))
                    if m:
                        cyl_map[ci] = (label, f"Cyl {m.group(1)}")

    i = 5
    while i < len(grid) - 1:
        r1 = grid[i]
        r2 = grid[i + 1] if i + 1 < len(grid) else []

        name = clean_text(r1[0]) if r1 else ""
        if not looks_like_component_name(name):
            i += 1
            continue

        periodicity = parse_periodicity(r1[1] if len(r1) > 1 else None)

        for ci, (eng_label, cyl_label) in cyl_map.items():
            d = parse_date(r1[ci] if ci < len(r1) else None)
            h = extract_positive_number(r2[ci] if ci < len(r2) else None)

            if d is None and h is None:
                continue

            comps.append({
                "category": "AUXENGINE",
                "engine_label": eng_label,
                "unit": cyl_label,
                "description": name,
                "periodicity": periodicity,
                "last_oh_date": d,
                "last_oh_hrs": h,
                "hrs_since": h,
                "pct_used": calc_pct_used(h, periodicity),
                "status": derive_status(h, periodicity),
            })
        i += 2

    return comps

def parse_other_equipment_tables(doc: Document) -> List[Dict]:
    rows = []
    for table in doc.tables:
        grid = table_to_grid(table)
        for row in grid:
            joined = normalize_for_match(row_text(row))
            for hint in OTHER_SECTION_HINTS:
                if hint in joined:
                    pass

        for row in grid:
            cells = [clean_text(c) for c in row]
            joined = normalize_for_match(" ".join(cells))
            if any(h in joined for h in OTHER_SECTION_HINTS):
                continue

            if len(cells) < 2:
                continue

            desc = None
            for c in cells:
                if looks_like_component_name(c) and len(c) > 3:
                    desc = c
                    break
            if not desc:
                continue

            date_val = None
            hrs_val = None
            per_val = None
            for c in cells:
                if date_val is None:
                    d = parse_date(c)
                    if d and d != c:
                        date_val = d
                if hrs_val is None:
                    h = extract_positive_number(c)
                    if h is not None:
                        hrs_val = str(int(h)) if float(h).is_integer() else str(h)
                if per_val is None and re.search(r"HR|DAY|MONTH|YEAR|\d", c, flags=re.I):
                    per_val = clean_text(c)

            if date_val or hrs_val:
                rows.append({
                    "section": "OTHER EQUIPMENT",
                    "description": desc,
                    "periodicity": per_val or "",
                    "last_date": date_val or "",
                    "run_hrs": hrs_val or "",
                })

    dedup = []
    seen = set()
    for r in rows:
        key = (r["section"], r["description"], r["last_date"], r["run_hrs"])
        if key not in seen:
            seen.add(key)
            dedup.append(r)
    return dedup

def deduplicate_components(comps: List[Dict]) -> List[Dict]:
    out = []
    seen = set()
    for x in comps:
        key = (
            x["category"], x["engine_label"], x["unit"],
            x["description"], x.get("last_oh_date"), x.get("hrs_since")
        )
        if key not in seen:
            seen.add(key)
            out.append(x)
    return out

def assess_parse_quality(vessel_name: str, comps: List[Dict], warnings: List[str]) -> float:
    score = 0.0
    if vessel_name and vessel_name != "UNKNOWN":
        score += 0.20
    if len(comps) >= 10:
        score += 0.30
    if any(c["category"] == "MAINENGINE" for c in comps):
        score += 0.20
    if any(c["category"] == "AUXENGINE" for c in comps):
        score += 0.15
    usable = [c for c in comps if c.get("periodicity") and c.get("hrs_since") is not None]
    if len(usable) >= 8:
        score += 0.10
    if len(warnings) == 0:
        score += 0.05
    return round(min(score, 1.0), 2)

def parse_doc_bytes(docx_bytes: bytes) -> Dict:
    doc = open_docx_bytes(docx_bytes)
    warnings = []

    vessel_name = detect_vessel_name(doc)
    if vessel_name == "UNKNOWN":
        warnings.append("Could not confidently extract vessel name.")

    report_date = detect_report_date(doc)
    me_total_hrs, me_this_month = detect_me_totals(doc)

    me_tables, aux_tables, other_tables, candidates = find_best_tables(doc)

    components = []

    if me_tables:
        components.extend(parse_main_engine_table(me_tables[0]["grid"]))
    else:
        warnings.append("No main engine candidate table detected.")

    if aux_tables:
        components.extend(parse_aux_engine_table(aux_tables[0]["grid"]))
    else:
        warnings.append("No auxiliary engine candidate table detected.")

    other_equipment = parse_other_equipment_tables(doc)

    components = deduplicate_components(components)

    if not components:
        warnings.append("No components extracted from document.")
    elif len(components) < 8:
        warnings.append(f"Low component count extracted ({len(components)}).")

    confidence = assess_parse_quality(vessel_name, components, warnings)

    return {
        "vessel_name": vessel_name,
        "report_date": report_date,
        "me_total_hrs": me_total_hrs,
        "me_this_month": me_this_month,
        "components": components,
        "other_equipment": other_equipment,
        "warnings": warnings,
        "parse_confidence": confidence,
        "table_diagnostics": [
            {
                "table_index": t["idx"],
                "me_score": t["me_score"],
                "aux_score": t["aux_score"],
                "other_score": t["other_score"],
                "rows": len(t["grid"]),
                "cols": max((len(r) for r in t["grid"]), default=0),
            }
            for t in candidates
        ],
    }

# ============================================================
# DB SAVE / FETCH
# ============================================================
def save_parsed(parsed: Dict, filename: str, file_hash: str):
    vessel_name = parsed["vessel_name"]
    now = datetime.utcnow().isoformat() + "Z"

    confidence = float(parsed.get("parse_confidence", 0))
    component_count = len(parsed.get("components", []))
    warning_count = len(parsed.get("warnings", []))

    if confidence < 0.35 or component_count < 5:
        raise ValueError(
            f"Parse confidence too low to commit safely "
            f"(confidence={confidence}, components={component_count})."
        )

    conn = get_db()
    c = conn.cursor()
    try:
        c.execute(
            "INSERT OR IGNORE INTO vessels(name, created_at) VALUES (?, ?)",
            (vessel_name, now)
        )

        c.execute("""
            INSERT INTO upload_log(
                vessel_name, filename, file_hash, report_date,
                me_total_hrs, me_this_month, parse_confidence,
                component_count, warning_count, uploaded_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            vessel_name,
            filename,
            file_hash,
            parsed.get("report_date"),
            parsed.get("me_total_hrs"),
            parsed.get("me_this_month"),
            confidence,
            component_count,
            warning_count,
            now
        ))

        c.execute("DELETE FROM components WHERE vessel_name = ?", (vessel_name,))
        for x in parsed["components"]:
            c.execute("""
                INSERT INTO components(
                    vessel_name, category, engine_label, unit, description,
                    periodicity, last_oh_date, last_oh_hrs, hrs_since,
                    pct_used, status, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                vessel_name,
                x["category"],
                x["engine_label"],
                x["unit"],
                x["description"],
                x.get("periodicity"),
                x.get("last_oh_date"),
                x.get("last_oh_hrs"),
                x.get("hrs_since"),
                x.get("pct_used"),
                x.get("status"),
                now
            ))

        c.execute("DELETE FROM other_equipment WHERE vessel_name = ?", (vessel_name,))
        for x in parsed["other_equipment"]:
            c.execute("""
                INSERT INTO other_equipment(
                    vessel_name, section, description, periodicity,
                    last_date, run_hrs, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                vessel_name,
                x["section"],
                x["description"],
                x.get("periodicity", ""),
                x.get("last_date", ""),
                x.get("run_hrs", ""),
                now
            ))

        conn.commit()
    except:
        conn.rollback()
        raise
    finally:
        conn.close()

@st.cache_data(ttl=10)
def get_vessels():
    conn = get_db()
    rows = conn.execute("SELECT name FROM vessels ORDER BY name").fetchall()
    conn.close()
    return [r["name"] for r in rows]

@st.cache_data(ttl=10)
def get_components(vessel_name: str):
    conn = get_db()
    df = pd.read_sql_query(
        "SELECT * FROM components WHERE vessel_name = ?",
        conn,
        params=(vessel_name,)
    )
    conn.close()
    return df

@st.cache_data(ttl=10)
def get_other_equipment(vessel_name: str):
    conn = get_db()
    df = pd.read_sql_query(
        "SELECT * FROM other_equipment WHERE vessel_name = ? ORDER BY section, description",
        conn,
        params=(vessel_name,)
    )
    conn.close()
    return df

@st.cache_data(ttl=10)
def get_history(vessel_name: str):
    conn = get_db()
    df = pd.read_sql_query("""
        SELECT filename, report_date, me_total_hrs, me_this_month,
               parse_confidence, component_count, warning_count, uploaded_at
        FROM upload_log
        WHERE vessel_name = ?
        ORDER BY uploaded_at DESC
        LIMIT 20
    """, conn, params=(vessel_name,))
    conn.close()
    return df

@st.cache_data(ttl=10)
def get_summary():
    conn = get_db()
    df = pd.read_sql_query("""
        SELECT
            vessel_name,
            SUM(CASE WHEN status = 'OVERDUE' THEN 1 ELSE 0 END) AS overdue,
            SUM(CASE WHEN status = 'HIGH PRIORITY' THEN 1 ELSE 0 END) AS high_priority,
            SUM(CASE WHEN status = 'OK' THEN 1 ELSE 0 END) AS ok,
            SUM(CASE WHEN status = 'NO DATA' THEN 1 ELSE 0 END) AS no_data,
            COUNT(*) AS total
        FROM components
        GROUP BY vessel_name
        ORDER BY overdue DESC, high_priority DESC, vessel_name ASC
    """, conn)
    conn.close()
    return df

@st.cache_data(ttl=10)
def get_all_fleet_components():
    conn = get_db()
    df = pd.read_sql_query("SELECT * FROM components", conn)
    conn.close()
    return df

# ============================================================
# DISPLAY TABLES
# ============================================================
STATUS_THEME = {
    "OVERDUE": {"bg":"#2d0707", "status":"#ff6b6b", "main":"#ff8080", "accent":"#ff3333", "dim":"#773333"},
    "HIGH PRIORITY": {"bg":"#2d1503", "status":"#ffaa44", "main":"#ff9933", "accent":"#ffcc00", "dim":"#774422"},
    "OK": {"bg":"#042010", "status":"#4ade80", "main":"#22c55e", "accent":"#4ade80", "dim":"#0f4023"},
    "NO DATA": {"bg":"#0c1422", "status":"#7da3d8", "main":"#5f7fa6", "accent":"#7da3d8", "dim":"#2a3950"},
}

def cyl_sort(unit: str) -> int:
    m = re.search(r"(\d+)", str(unit))
    return int(m.group(1)) if m else 999

def build_display_df(df: pd.DataFrame, priority: bool = False) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=[
            "Status", "Vessel", "Component", "Engine", "Unit",
            "Periodicity", "Last OH", "Hrs Since", "Used"
        ])

    d = df.copy()

    if priority:
        order_map = {"OVERDUE": 0, "HIGH PRIORITY": 1, "OK": 2, "NO DATA": 3}
        d["_ord"] = d["status"].map(lambda x: order_map.get(str(x), 9))
        d["_pct"] = d["pct_used"].fillna(0)
        sort_cols = ["_ord", "_pct"]
        asc = [True, False]
        if "vessel_name" in d.columns:
            sort_cols.append("vessel_name")
            asc.append(True)
        d = d.sort_values(sort_cols, ascending=asc).drop(columns=["_ord", "_pct"])
    else:
        d["_k1"] = d["description"].astype(str).str.upper()
        d["_k2"] = d["unit"].apply(cyl_sort)
        sort_cols = ["_k1", "_k2"]
        asc = [True, True]
        if "vessel_name" in d.columns:
            sort_cols = ["vessel_name"] + sort_cols
            asc = [True] + asc
        d = d.sort_values(sort_cols, ascending=asc).drop(columns=["_k1", "_k2"])

    out = pd.DataFrame(index=range(len(d)))
    out["Status"] = d["status"].fillna("")
    out["Vessel"] = d["vessel_name"] if "vessel_name" in d.columns else ""
    out["Component"] = d["description"].fillna("")
    out["Engine"] = d["engine_label"].fillna("")
    out["Unit"] = d["unit"].fillna("")
    out["Periodicity"] = d["periodicity"].apply(lambda x: int(x) if pd.notna(x) else None)
    out["Last OH"] = d["last_oh_date"].fillna("")
    out["Hrs Since"] = d["hrs_since"].apply(lambda x: int(x) if pd.notna(x) else None)
    out["Used"] = d["pct_used"].apply(lambda x: round(float(x) * 100, 1) if pd.notna(x) else 0.0)
    return out

def apply_style(df: pd.DataFrame):
    def style_row(row):
        s = STATUS_THEME.get(str(row.get("Status")), STATUS_THEME["NO DATA"])
        return [
            f"background-color:{s['bg']}; color:{s['status']}; font-weight:700",
            f"background-color:{s['bg']}; color:{s['main']}; font-weight:600",
            f"background-color:{s['bg']}; color:{s['dim']}",
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
    "Last OH": st.column_config.TextColumn("Last OH", width=110),
    "Hrs Since": st.column_config.NumberColumn("Hrs Since", format="%d hrs", width=110),
    "Used": st.column_config.ProgressColumn("Used", min_value=0, max_value=160, format="%.1f%%", width=120),
}

def render_table(df: pd.DataFrame, height: Optional[int] = None, priority: bool = False):
    if isinstance(df, list):
        df = pd.DataFrame(df)
    if df.empty:
        st.info("No data to display.")
        return
    tbl = build_display_df(df, priority=priority)
    h = height or min(900, 38 * len(tbl) + 42)
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
    st.markdown("<h2 style='margin-bottom:.15rem'>FLEET MONITOR</h2>", unsafe_allow_html=True)
    st.markdown("<div class='smallcap' style='margin-bottom:1rem'>Running Hours Intelligence System</div>", unsafe_allow_html=True)

    page = st.selectbox(
        "Navigation",
        ["Fleet Overview", "Vessel Detail", "Upload Report", "Upload History"],
        label_visibility="collapsed",
    )

    vessels = get_vessels()
    selected_vessel = st.selectbox("Active Vessel", vessels if vessels else ["—"], disabled=not bool(vessels))

    st.markdown("<hr>", unsafe_allow_html=True)
    db_kb = DB_PATH.stat().st_size / 1024 if DB_PATH.exists() else 0
    st.markdown(
        f"<div class='code'>db {db_kb:.0f} kb · {len(vessels)} vessels · hardened parser build</div>",
        unsafe_allow_html=True
    )

# ============================================================
# FLEET OVERVIEW
# ============================================================
if page == "Fleet Overview":
    st.markdown(hero("Fleet Master Matrix", "Universal Fleet Telemetry"), unsafe_allow_html=True)

    summary = get_summary()
    all_comps = get_all_fleet_components()

    if summary.empty or all_comps.empty:
        st.info("No data loaded. Upload a report to begin.")
        st.stop()

    total_vessels = len(summary)
    total_components = len(all_comps)
    overdue = int((all_comps["status"] == "OVERDUE").sum())
    highp = int((all_comps["status"] == "HIGH PRIORITY").sum())
    ok = int((all_comps["status"] == "OK").sum())

    cols = st.columns(5)
    with cols[0]: st.markdown(kpi(total_vessels, "Vessels", "blue"), unsafe_allow_html=True)
    with cols[1]: st.markdown(kpi(total_components, "Components", "gold"), unsafe_allow_html=True)
    with cols[2]: st.markdown(kpi(overdue, "Overdue", "red"), unsafe_allow_html=True)
    with cols[3]: st.markdown(kpi(highp, "High Priority", "orange"), unsafe_allow_html=True)
    with cols[4]: st.markdown(kpi(ok, "OK", "green"), unsafe_allow_html=True)

    st.markdown("<div class='smallcap' style='margin:1rem 0 .5rem'>Universal Component Control Grid</div>", unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns([1.4, 1.4, 1.6, 2.2])

    with c1:
        vessel_filter = st.selectbox("Filter Vessel Context", ["All Fleet"] + sorted(all_comps["vessel_name"].dropna().unique().tolist()))
    with c2:
        category_filter = st.selectbox("Filter Machinery Type", ["All", "Main Engine", "Aux Engines"])
    with c3:
        status_filter = st.selectbox("Filter Component Urgency", ["All Statuses", "Critical Focus", "Overdue Only", "High Priority Only", "OK Only", "No Data Only"])
    with c4:
        component_filter = st.selectbox("Search Component Definition", ["All"] + sorted(all_comps["description"].dropna().unique().tolist()))

    filt = all_comps.copy()

    if vessel_filter != "All Fleet":
        filt = filt[filt["vessel_name"] == vessel_filter]
    if category_filter == "Main Engine":
        filt = filt[filt["category"] == "MAINENGINE"]
    elif category_filter == "Aux Engines":
        filt = filt[filt["category"] == "AUXENGINE"]

    if status_filter == "Critical Focus":
        filt = filt[filt["status"].isin(["OVERDUE", "HIGH PRIORITY"])]
    elif status_filter == "Overdue Only":
        filt = filt[filt["status"] == "OVERDUE"]
    elif status_filter == "High Priority Only":
        filt = filt[filt["status"] == "HIGH PRIORITY"]
    elif status_filter == "OK Only":
        filt = filt[filt["status"] == "OK"]
    elif status_filter == "No Data Only":
        filt = filt[filt["status"] == "NO DATA"]

    if component_filter != "All":
        filt = filt[filt["description"] == component_filter]

    st.markdown(
        f"<div class='code' style='margin:.4rem 0 1rem'>showing {len(filt)} records · "
        f"{int((filt['status']=='OVERDUE').sum())} overdue · "
        f"{int((filt['status']=='HIGH PRIORITY').sum())} high priority · "
        f"{int((filt['status']=='OK').sum())} ok</div>",
        unsafe_allow_html=True
    )

    if filt.empty:
        st.warning("No records match the current filter matrix.")
    else:
        render_table(filt, height=min(950, 38 * len(filt) + 44), priority=True)

# ============================================================
# VESSEL DETAIL
# ============================================================
elif page == "Vessel Detail":
    if not vessels:
        st.info("No vessel data available yet.")
        st.stop()

    if selected_vessel == "—":
        st.info("Select a vessel from the sidebar.")
        st.stop()

    st.markdown(hero(selected_vessel, "Component Analysis"), unsafe_allow_html=True)

    df = get_components(selected_vessel)
    oe = get_other_equipment(selected_vessel)

    if df.empty:
        st.info("No data for this vessel.")
        st.stop()

    total = len(df)
    overdue = int((df["status"] == "OVERDUE").sum())
    highp = int((df["status"] == "HIGH PRIORITY").sum())
    ok = int((df["status"] == "OK").sum())
    nodata = int((df["status"] == "NO DATA").sum())

    cols = st.columns(5)
    with cols[0]: st.markdown(kpi(total, "Total", "gold"), unsafe_allow_html=True)
    with cols[1]: st.markdown(kpi(overdue, "Overdue", "red"), unsafe_allow_html=True)
    with cols[2]: st.markdown(kpi(highp, "High Priority", "orange"), unsafe_allow_html=True)
    with cols[3]: st.markdown(kpi(ok, "OK", "green"), unsafe_allow_html=True)
    with cols[4]: st.markdown(kpi(nodata, "No Data", "blue"), unsafe_allow_html=True)

    hist = get_history(selected_vessel)
    if not hist.empty:
        last = hist.iloc[0]
        st.markdown(
            f"""
            <div class="panel" style="margin:.7rem 0 1rem">
                <div class="smallcap">Latest Accepted Upload</div>
                <div class="code" style="margin-top:.55rem">
                    file: <b>{last['filename']}</b><br>
                    report: <b>{last['report_date'] or '-'}</b><br>
                    me total: <b>{int(last['me_total_hrs']) if pd.notna(last['me_total_hrs']) else '-'}</b> ·
                    this month: <b>{int(last['me_this_month']) if pd.notna(last['me_this_month']) else '-'}</b><br>
                    confidence: <b>{float(last['parse_confidence']):.2f}</b> ·
                    components: <b>{int(last['component_count'])}</b> ·
                    warnings: <b>{int(last['warning_count'])}</b><br>
                    uploaded: <b>{str(last['uploaded_at'])[:16]}</b>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    tabs = st.tabs(["Alerts", "Main Engine", "Aux Engines", "Other Equipment"])

    with tabs[0]:
        alerts = df[df["status"].isin(["OVERDUE", "HIGH PRIORITY"])]
        if alerts.empty:
            st.success("All machinery components are within acceptable operational bounds.")
        else:
            render_table(alerts, priority=True)

    with tabs[1]:
        me = df[df["category"] == "MAINENGINE"]
        if me.empty:
            st.info("No Main Engine data available.")
        else:
            a, b = st.columns(2)
            with a:
                comp = st.selectbox("Machinery Element", ["All"] + sorted(me["description"].dropna().unique().tolist()), key="me_comp")
            with b:
                stat = st.selectbox("Machinery Status", ["All", "Overdue only", "High Priority only", "OK only", "No Data only"], key="me_stat")

            v = me.copy()
            if comp != "All":
                v = v[v["description"] == comp]
            if stat == "Overdue only":
                v = v[v["status"] == "OVERDUE"]
            elif stat == "High Priority only":
                v = v[v["status"] == "HIGH PRIORITY"]
            elif stat == "OK only":
                v = v[v["status"] == "OK"]
            elif stat == "No Data only":
                v = v[v["status"] == "NO DATA"]

            render_table(v)

    with tabs[2]:
        aux = df[df["category"] == "AUXENGINE"]
        if aux.empty:
            st.info("No Auxiliary Engine data available.")
        else:
            a, b = st.columns(2)
            with a:
                eng = st.selectbox("Aux Generator Node", ["All"] + sorted(aux["engine_label"].dropna().unique().tolist()), key="aux_eng")
            with b:
                stat = st.selectbox("Node Condition", ["All", "Overdue only", "High Priority only", "OK only", "No Data only"], key="aux_stat")

            v = aux.copy()
            if eng != "All":
                v = v[v["engine_label"] == eng]
            if stat == "Overdue only":
                v = v[v["status"] == "OVERDUE"]
            elif stat == "High Priority only":
                v = v[v["status"] == "HIGH PRIORITY"]
            elif stat == "OK only":
                v = v[v["status"] == "OK"]
            elif stat == "No Data only":
                v = v[v["status"] == "NO DATA"]

            render_table(v)

    with tabs[3]:
        if oe.empty:
            st.info("No auxiliary plant or extension machinery data located.")
        else:
            st.dataframe(
                oe.rename(columns={
                    "section": "Section",
                    "description": "Machinery Description",
                    "periodicity": "Maintenance Periodicity",
                    "last_date": "Inspection Date",
                    "run_hrs": "Logged Hours"
                }),
                use_container_width=True,
                hide_index=True,
                height=500
            )

# ============================================================
# UPLOAD REPORT
# ============================================================
elif page == "Upload Report":
    st.markdown(hero("Upload Report", "TEC-004 Log Processing"), unsafe_allow_html=True)

    left, right = st.columns([1.7, 1.1], gap="large")

    with left:
        uploaded = st.file_uploader("Upload file", type=["doc"], label_visibility="collapsed")

    with right:
        st.markdown("""
        <div class="panel">
            <div class="smallcap">Accepted Specification</div>
            <div style="margin-top:.6rem; line-height:1.8">
                TEC-004 Running Hours Monthly Log Report<br>
                Native <b>.doc</b> binary streams only
            </div>
            <div class="smallcap" style="margin-top:1rem">Parser Guarantees</div>
            <div style="margin-top:.5rem; line-height:1.8" class="muted">
                Vessel and report metadata detection<br>
                Main engine + auxiliary engine table discovery<br>
                Robust row-pair extraction with fallbacks<br>
                Confidence-scored pre-commit validation<br>
                Safe commit blocking on low-confidence parses
            </div>
        </div>
        """, unsafe_allow_html=True)

    if uploaded:
        raw = uploaded.read()
        file_hash = hashlib.md5(raw).hexdigest()

        with st.spinner("Converting .doc and executing hardened extraction pipeline..."):
            try:
                docx_bytes = convert_doc_to_docx(raw)
            except Exception as e:
                st.error(f"Document conversion failed: {e}")
                st.stop()

            try:
                parsed = parse_doc_bytes(docx_bytes)
            except Exception as e:
                st.error(f"Telemetry extraction failed: {e}")
                st.stop()

        comps = parsed["components"]
        nc = len(comps)
        overdue = sum(1 for c in comps if c["status"] == "OVERDUE")
        highp = sum(1 for c in comps if c["status"] == "HIGH PRIORITY")
        ok = sum(1 for c in comps if c["status"] == "OK")
        nodata = sum(1 for c in comps if c["status"] == "NO DATA")
        other_n = len(parsed["other_equipment"])
        conf = parsed["parse_confidence"]

        cols = st.columns(6)
        cols[0].metric("Asset", parsed["vessel_name"])
        cols[1].metric("Report Window", parsed["report_date"] or "-")
        cols[2].metric("ME Accumulated", f"{parsed['me_total_hrs']:,}" if parsed["me_total_hrs"] else "-")
        cols[3].metric("Monthly Increment", f"{parsed['me_this_month']:,}" if parsed["me_this_month"] else "-")
        cols[4].metric("Data Channels", nc)
        cols[5].metric("Confidence", f"{conf:.2f}")

        st.markdown(
            f"""
            <div class="panel" style="margin:1rem 0">
                <div class="smallcap">Parse Health</div>
                <div class="code" style="margin-top:.55rem">
                    overdue={overdue} · high_priority={highp} · ok={ok} · no_data={nodata} · other_equipment={other_n}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        if conf >= 0.75:
            st.success("High-confidence parse. Safe to commit.")
        elif conf >= 0.50:
            st.warning("Moderate-confidence parse. Review preview carefully before commit.")
        else:
            st.error("Low-confidence parse. Commit is blocked until extraction is reliable enough.")

        if parsed["warnings"]:
            for w in parsed["warnings"]:
                st.warning(w)

        diag = pd.DataFrame(parsed["table_diagnostics"])
        if not diag.empty:
            with st.expander("Table diagnostics"):
                st.dataframe(diag, use_container_width=True, hide_index=True)

        st.markdown("<div class='smallcap' style='margin:1rem 0 .5rem'>Extracted Telemetry Preview</div>", unsafe_allow_html=True)

        df_preview = pd.DataFrame(comps) if comps else pd.DataFrame()
        tabs = st.tabs(["Main Engine Matrix", "Aux Engines Matrix", "Other Equipment"])

        with tabs[0]:
            if not df_preview.empty:
                meprev = df_preview[df_preview["category"] == "MAINENGINE"]
                if not meprev.empty:
                    render_table(meprev, height=420, priority=True)
                else:
                    st.info("No Main Engine telemetry extracted.")
            else:
                st.info("No component data available.")

        with tabs[1]:
            if not df_preview.empty:
                auxprev = df_preview[df_preview["category"] == "AUXENGINE"]
                if not auxprev.empty:
                    render_table(auxprev, height=420, priority=True)
                else:
                    st.info("No Auxiliary Engine telemetry extracted.")
            else:
                st.info("No component data available.")

        with tabs[2]:
            if parsed["other_equipment"]:
                oep = pd.DataFrame(parsed["other_equipment"])
                st.dataframe(oep, use_container_width=True, hide_index=True, height=420)
            else:
                st.info("No Other Equipment data extracted.")

        cbtn, _ = st.columns([1.2, 4])
        with cbtn:
            commit_disabled = conf < 0.35 or nc < 5 or parsed["vessel_name"] == "UNKNOWN"
            if st.button("COMMIT STREAM TO DATABASE", use_container_width=True, disabled=commit_disabled):
                try:
                    save_parsed(parsed, uploaded.name, file_hash)
                    for fn in [get_vessels, get_components, get_other_equipment, get_summary, get_all_fleet_components, get_history]:
                        fn.clear()
                    st.success(
                        f"System telemetry confirmed. {parsed['vessel_name']} committed with {nc} component rows."
                    )
                    st.balloons()
                except Exception as e:
                    st.error(str(e))

# ============================================================
# UPLOAD HISTORY
# ============================================================
elif page == "Upload History":
    st.markdown(hero("Upload History", "System Audit Trails"), unsafe_allow_html=True)

    if not vessels or selected_vessel == "—":
        st.info("Select a vessel from the sidebar.")
        st.stop()

    hist = get_history(selected_vessel)
    if hist.empty:
        st.info("No upload trail entries recorded for this vessel.")
    else:
        d = hist.copy()
        if "me_total_hrs" in d.columns:
            d["me_total_hrs"] = d["me_total_hrs"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
        if "me_this_month" in d.columns:
            d["me_this_month"] = d["me_this_month"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
        d = d.rename(columns={
            "filename": "Logged Filename",
            "report_date": "Extracted Target Date",
            "me_total_hrs": "ME Combined Total",
            "me_this_month": "ME Monthly Increment",
            "parse_confidence": "Parse Confidence",
            "component_count": "Component Count",
            "warning_count": "Warning Count",
            "uploaded_at": "Transaction Timestamp",
        })
        st.dataframe(d, use_container_width=True, hide_index=True, height=520)
