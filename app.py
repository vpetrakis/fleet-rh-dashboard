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

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap');

:root {
  --bg:#060b16; --bg1:#0b1324; --bg2:#0f1a31; --bg3:#13203b; --bg4:#182746;
  --b1:#1b2a46; --b2:#24365a; --b3:#304772;
  --gold:#d1a12b; --gold2:#e4ba4a; --gold3:#f3d06b;
  --red:#d84a4a; --red2:#ff7a7a;
  --orange:#df7f2f; --ora2:#ffad5a;
  --green:#1fa96b; --grn2:#39d98a; --grn3:#97f0c3;
  --blue:#2d6cdf; --blu2:#6ba1ff;
  --t0:#f4f8ff; --t1:#d1ddf2; --t2:#94a9c8; --t3:#5d7396;
  --ff:'Inter',sans-serif;
  --fm:'JetBrains Mono',monospace;
}

html, body, [class*="css"] {
  background: var(--bg) !important;
  color: var(--t1) !important;
  font-family: var(--ff) !important;
}
.main, .main > div { background: var(--bg) !important; }
.block-container { max-width: 100% !important; padding: 1.5rem 1.75rem 3rem !important; }

.main::before{
  content:"";
  position:fixed; inset:0; pointer-events:none; z-index:0;
  background:
    radial-gradient(ellipse 80% 45% at -10% -10%, rgba(209,161,43,.09) 0%, transparent 55%),
    radial-gradient(ellipse 70% 45% at 110% 110%, rgba(45,108,223,.08) 0%, transparent 55%);
}

[data-testid="stSidebar"]{
  background:var(--bg1)!important;
  border-right:1px solid var(--b2)!important;
}
[data-testid="stSidebar"] * { color: var(--t1)!important; }
[data-testid="stSidebarContent"] { padding:1.1rem!important; }

h1,h2,h3 { color:var(--t0)!important; letter-spacing:-.02em!important; }
h1 { font-size:1.7rem!important; font-weight:700!important; }
h2 { font-size:1.12rem!important; font-weight:700!important; }
h3 { font-size:.96rem!important; font-weight:600!important; }

.stSelectbox > div > div,
.stMultiSelect > div > div,
.stTextInput > div > div > input {
  background: var(--bg3)!important;
  border: 1px solid var(--b2)!important;
  color: var(--t1)!important;
  border-radius: 10px!important;
}

.stButton > button {
  background: linear-gradient(135deg, var(--gold) 0%, #8f6913 100%) !important;
  color:#07111f!important; border:none!important;
  border-radius:10px!important; padding:.72rem 1.1rem!important;
  font-weight:800!important; letter-spacing:.05em!important; text-transform:uppercase!important;
}
.stButton > button:hover { transform:translateY(-1px); filter:brightness(1.03); }

[data-testid="stMetric"]{
  background: linear-gradient(180deg, rgba(255,255,255,.02), rgba(255,255,255,.00)), var(--bg3)!important;
  border:1px solid var(--b2)!important;
  border-radius:14px!important;
  padding: .9rem 1rem 1rem!important;
}
[data-testid="stMetricValue"]{ color:var(--t0)!important; font-weight:800!important; }
[data-testid="stMetricLabel"]{
  color:var(--t3)!important; text-transform:uppercase!important;
  letter-spacing:.13em!important; font-size:.66rem!important;
}

[data-testid="stDataFrame"]{
  border:1px solid var(--b2)!important;
  border-radius:14px!important;
  overflow:hidden!important;
  background:var(--bg2)!important;
}

.stTabs [data-baseweb="tab-list"]{
  background:var(--bg2)!important;
  border:1px solid var(--b2)!important;
  border-radius:14px 14px 0 0!important;
  gap:0!important;
}
.stTabs [data-baseweb="tab"]{
  color:var(--t3)!important;
  font-weight:700!important; text-transform:uppercase!important;
  letter-spacing:.05em!important; font-size:.73rem!important;
}
.stTabs [aria-selected="true"]{
  color:var(--gold2)!important;
  border-bottom:2px solid var(--gold)!important;
}
.stTabs [data-baseweb="tab-panel"]{
  background:var(--bg2)!important;
  border:1px solid var(--b2)!important;
  border-top:none!important;
  border-radius:0 0 14px 14px!important;
  padding:1.15rem!important;
}

hr { border-color: var(--b2)!important; }

.hero { margin-bottom: 1rem; }
.eyebrow {
  color: var(--gold2); font-size: .66rem; text-transform: uppercase; letter-spacing: .18em;
  margin-bottom: .25rem; font-weight: 700;
}
.hero-line {
  height: 1px; background: linear-gradient(90deg, var(--gold), var(--b2), transparent); margin-top:.55rem;
}
.panel {
  background: linear-gradient(180deg, rgba(255,255,255,.02), rgba(255,255,255,.00)), var(--bg3);
  border: 1px solid var(--b2); border-radius: 14px; padding: 1rem 1.05rem;
}
.code { font-family: var(--fm); color: var(--t2); font-size: .74rem; }
.smallcap { color:var(--t3); text-transform:uppercase; letter-spacing:.16em; font-size:.66rem; }

.kpi {
  background: var(--bg3); border:1px solid var(--b2); border-radius:14px; padding:1rem 1.05rem;
}
.kpi .v { font-size:1.95rem; font-weight:800; line-height:1; }
.kpi .l { color:var(--t3); text-transform:uppercase; letter-spacing:.16em; font-size:.62rem; margin-top:.45rem; }
.kpi.gold .v{color:var(--gold3);} .kpi.red .v{color:var(--red2);} .kpi.orange .v{color:var(--ora2);} .kpi.green .v{color:var(--grn2);} .kpi.blue .v{color:var(--blu2);} 

.badge { display:inline-block; padding:.28rem .58rem; border-radius:999px; font-size:.68rem; font-weight:700; letter-spacing:.04em; }
.badge.ok{background:rgba(31,169,107,.12); color:var(--grn2); border:1px solid rgba(31,169,107,.28);} 
.badge.warn{background:rgba(223,127,47,.12); color:var(--ora2); border:1px solid rgba(223,127,47,.28);} 
.badge.bad{background:rgba(216,74,74,.12); color:var(--red2); border:1px solid rgba(216,74,74,.28);} 
.badge.info{background:rgba(45,108,223,.12); color:var(--blu2); border:1px solid rgba(45,108,223,.28);} 

[data-testid="stFileUploadDropzone"]{
  border:1.5px dashed var(--gold)!important;
  border-radius:14px!important;
  padding:2.4rem 1.3rem!important;
  background: linear-gradient(160deg, rgba(209,161,43,.05), rgba(45,108,223,.04))!important;
}
</style>
""", unsafe_allow_html=True)

DB_PATH = Path("runninghours.db")


def hero(title: str, eyebrow: str = ""):
    return f"""
    <div class='hero'>
      {f"<div class='eyebrow'>{eyebrow}</div>" if eyebrow else ""}
      <h1>{title}</h1>
      <div class='hero-line'></div>
    </div>
    """


def kpi(value, label, color="gold"):
    return f"""
    <div class='kpi {color}'>
      <div class='v'>{value}</div>
      <div class='l'>{label}</div>
    </div>
    """


def get_db():
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    c = conn.cursor()
    c.execute("PRAGMA journal_mode=WAL")
    c.executescript(
        """
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
        CREATE INDEX IF NOT EXISTS idx_upload_vessel ON upload_log(vessel_name);
        """
    )
    conn.commit()
    conn.close()


init_db()


def clean_text(x) -> str:
    if x is None:
        return ""
    x = str(x).replace("\xa0", " ")
    x = re.sub(r"\s+", " ", x).strip()
    return x


def normalized(x) -> str:
    x = clean_text(x).upper()
    x = x.replace("O/H", "OH")
    x = x.replace("PERIODICTLY", "PERIODICITY")
    x = re.sub(r"[^A-Z0-9 /().:-]", "", x)
    return x


def safe_number(x) -> Optional[float]:
    s = clean_text(x)
    if not s or s.upper() in {"N/A", "NA", "-", "--"}:
        return None
    s = s.strip("[]")
    s = s.replace(",", "")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group())
    except Exception:
        return None


def parse_periodicity(x) -> Optional[float]:
    s = clean_text(x)
    if not s:
        return None
    u = s.upper()
    if "OBSERVATION" in u or u in {"N/A", "NA", "-", "--"}:
        return None
    m = re.search(r"\d[\d.,]*", s)
    if not m:
        return None
    raw = m.group().replace(",", "")
    if "." in raw and raw.count(".") == 1:
        left, right = raw.split(".")
        if right in {"000", "00", "0"}:
            raw = left + right
    raw = raw.replace(".", "") if raw.count(".") > 1 else raw
    try:
        return float(raw)
    except Exception:
        return None


def parse_date(x) -> Optional[str]:
    s = clean_text(x)
    if not s or s.upper() in {"N/A", "NA", "-", "--"}:
        return None
    s = s.strip("[]")
    s = re.sub(r"\.", " ", s)
    s = re.sub(r"SEPT", "SEP", s, flags=re.I)
    s = re.sub(r"JUNE", "JUN", s, flags=re.I)
    s = re.sub(r"JULY", "JUL", s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip()
    candidates = [s, s.upper(), s.title()]
    fmts = [
        "%d %b %y", "%d %b %Y", "%d %B %y", "%d %B %Y",
        "%d/%m/%y", "%d/%m/%Y", "%d-%m-%y", "%d-%m-%Y",
        "%Y-%m-%d"
    ]
    for c in candidates:
        for f in fmts:
            try:
                return datetime.strptime(c, f).strftime("%Y-%m-%d")
            except Exception:
                pass
    return None


def derive_status(hrs_since: Optional[float], periodicity: Optional[float]) -> str:
    if hrs_since is None or periodicity is None or periodicity <= 0:
        return "NO DATA"
    ratio = hrs_since / periodicity
    if ratio >= 1.0:
        return "OVERDUE"
    if ratio >= 0.80:
        return "HIGH PRIORITY"
    return "OK"


def pct_used(hrs_since: Optional[float], periodicity: Optional[float]) -> float:
    if hrs_since is None or periodicity is None or periodicity <= 0:
        return 0.0
    return round(hrs_since / periodicity, 4)


def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found. Add libreoffice to packages.txt.")

    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as t:
        t.write(raw)
        in_path = t.name

    out_dir = tempfile.mkdtemp(prefix="tec004_")
    out_docx = os.path.join(out_dir, Path(in_path).stem + ".docx")
    profile = f"file:///tmp/lo_profile_{os.getpid()}_{os.urandom(4).hex()}"

    try:
        r = subprocess.run(
            [
                soffice,
                "--headless",
                "--norestore",
                "--nofirststartwizard",
                f"-env:UserInstallation={profile}",
                "--convert-to", "docx",
                in_path,
                "--outdir", out_dir,
            ],
            capture_output=True,
            timeout=120,
        )
        if not os.path.exists(out_docx):
            stderr = r.stderr.decode("utf-8", errors="ignore")[:500]
            raise RuntimeError(f"Conversion failed. {stderr}")
        with open(out_docx, "rb") as f:
            return f.read()
    finally:
        try:
            os.unlink(in_path)
        except Exception:
            pass
        try:
            if os.path.exists(out_docx):
                os.unlink(out_docx)
        except Exception:
            pass
        shutil.rmtree(out_dir, ignore_errors=True)


def open_docx_bytes(docx_bytes: bytes) -> Document:
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as t:
        t.write(docx_bytes)
        p = t.name
    try:
        doc = Document(p)
    finally:
        try:
            os.unlink(p)
        except Exception:
            pass
    return doc


def table_grid(table):
    return [[clean_text(c.text) for c in row.cells] for row in table.rows]


def detect_vessel(doc: Document) -> str:
    patterns = [
        r"VESSEL.?S NAME[:\s]+(?:MV\s+)?([A-Z0-9\- ]{2,})",
        r"NAME OF VESSEL[:\s]+([A-Z0-9\- ]{2,})",
        r"\bMV\s+([A-Z0-9\- ]{2,})",
    ]
    texts = [clean_text(p.text).upper() for p in doc.paragraphs if clean_text(p.text)]
    for table in doc.tables:
        for row in table_grid(table):
            texts.append(" ".join(row).upper())
    for txt in texts:
        for pat in patterns:
            m = re.search(pat, txt, flags=re.I)
            if m:
                return clean_text(m.group(1))
    return "UNKNOWN"


def detect_report_date(doc: Document) -> Optional[str]:
    texts = [clean_text(p.text) for p in doc.paragraphs if clean_text(p.text)]
    for table in doc.tables:
        for row in table_grid(table):
            texts.append(" ".join(row))
    for txt in texts:
        m = re.search(r"DATE[:\s]+([A-Z0-9 /.-]{5,25})", txt, flags=re.I)
        if m:
            d = parse_date(m.group(1))
            if d:
                return d
    return None


def detect_me_totals(doc: Document) -> Tuple[Optional[int], Optional[int]]:
    text = []
    for p in doc.paragraphs:
        t = clean_text(p.text)
        if t:
            text.append(t)
    for table in doc.tables:
        for row in table_grid(table):
            text.append(" ".join(row))
    joined = "\n".join(text)
    total = None
    month = None
    m1 = re.search(r"TOTAL RUNNING HOURS[:\s]*([\d,]+)", joined, flags=re.I)
    m2 = re.search(r"THIS MONTH[:\s]*([\d,]+)", joined, flags=re.I)
    if m1:
        total = int(m1.group(1).replace(",", ""))
    if m2:
        month = int(m2.group(1).replace(",", ""))
    return total, month


def is_me_table(grid) -> bool:
    text = " ".join(" ".join(r) for r in grid[:6]).upper()
    return "MAIN ENGINE" in text and "PERIODICITY" in text and "CYL" in text


def is_aux_table(grid) -> bool:
    text = " ".join(" ".join(r) for r in grid[:8]).upper()
    return "AUX. ENGINE" in text or ("DESCRIPTION" in text and "D/G NO1" in text) or ("DESCRIPTION" in text and "AUX. ENGINE NO.1" in text)


def find_me_table(doc: Document):
    for table in doc.tables:
        grid = table_grid(table)
        if is_me_table(grid):
            return grid
    return None


def find_aux_text_rows(doc: Document):
    rows = []
    for table in doc.tables:
        grid = table_grid(table)
        if is_aux_table(grid):
            rows.extend(grid)
    return rows


def parse_me_table(grid) -> Tuple[List[Dict], List[str]]:
    warnings = []
    if not grid:
        return [], ["Main engine table not found."]

    header_idx = None
    cyl_cols = []

    for i, row in enumerate(grid[:5]):
        for j, cell in enumerate(row):
            m = re.search(r"CYL\.?\s*NO\.?\s*(\d+)", normalized(cell))
            if m:
                header_idx = i
                cyl_cols.append((j, f"Cyl {m.group(1)}"))

    if not cyl_cols:
        warnings.append("Could not detect cylinder headers in main engine table.")
        return [], warnings

    comps = []
    i = header_idx + 1
    while i < len(grid) - 1:
        row1 = grid[i]
        row2 = grid[i + 1]

        desc = clean_text(row1[0]) if len(row1) > 0 else ""
        if not desc:
            i += 1
            continue

        udesc = normalized(desc)
        if any(x in udesc for x in ["NOTE 1", "NOTE 2", "DATE OF LAST", "RUNNING HOURS SINCE LAST", "TOTAL RUNNING HOURS"]):
            i += 1
            continue

        periodicity = parse_periodicity(row1[1] if len(row1) > 1 else None)

        for col, label in cyl_cols:
            date_val = None
            hrs_val = None

            # Expected structure in source: row1[col] = '1', row1[col+1] = date; row2[col] = '2', row2[col+1] = hours
            c11 = clean_text(row1[col]) if col < len(row1) else ""
            c12 = clean_text(row1[col + 1]) if col + 1 < len(row1) else ""
            c21 = clean_text(row2[col]) if col < len(row2) else ""
            c22 = clean_text(row2[col + 1]) if col + 1 < len(row2) else ""

            if c11 == "1":
                date_val = parse_date(c12)
            else:
                date_val = parse_date(c11) or parse_date(c12)

            if c21 == "2":
                hrs_val = safe_number(c22)
            else:
                hrs_val = safe_number(c21)
                if hrs_val is None:
                    hrs_val = safe_number(c22)

            if date_val is None and hrs_val is None:
                continue

            comps.append({
                "category": "MAINENGINE",
                "engine_label": "ME",
                "unit": label,
                "description": desc,
                "periodicity": periodicity,
                "last_oh_date": date_val,
                "last_oh_hrs": hrs_val,
                "hrs_since": hrs_val,
                "pct_used": pct_used(hrs_val, periodicity),
                "status": derive_status(hrs_val, periodicity),
            })

        i += 2

    if not comps:
        warnings.append("No main engine components extracted.")
    return comps, warnings


def parse_other_equipment(doc: Document) -> List[Dict]:
    out = []
    section = None
    for table in doc.tables:
        grid = table_grid(table)
        for row in grid:
            row_join = normalized(" ".join(row))
            if "TURBOCHARGER" in row_join and "COOLERS" in row_join:
                section = "OTHER EQUIPMENT"
                continue
            if "AUXILIARY BOILER" in row_join and "MAIN AIR COMPRESSORS" in row_join:
                section = "OTHER EQUIPMENT"
                continue
            if not section:
                continue
            cells = [clean_text(c) for c in row if clean_text(c)]
            if len(cells) < 2:
                continue
            desc = cells[0]
            if normalized(desc) in {"GENERAL O/H", "BALANCING OF ROTOR SHAFT", "AIR COOLER CLEANING", "FURNACE INSPECTION", "BURNER ATOMIZER", "FORCED DRAFT FAN", "FEED PUMPS NO.1", "FEED PUMPS NO.2"} or len(cells) >= 3:
                out.append({
                    "section": section,
                    "description": desc,
                    "periodicity": cells[1] if len(cells) > 1 else "",
                    "last_date": parse_date(cells[2]) if len(cells) > 2 else "",
                    "run_hrs": str(int(safe_number(cells[3]))) if len(cells) > 3 and safe_number(cells[3]) is not None else "",
                })
    dedup = []
    seen = set()
    for r in out:
        key = (r["section"], r["description"], r["periodicity"], r["last_date"], r["run_hrs"])
        if key not in seen:
            seen.add(key)
            dedup.append(r)
    return dedup


def parse_doc_bytes(docx_bytes: bytes) -> Dict:
    doc = open_docx_bytes(docx_bytes)
    warnings = []
    vessel_name = detect_vessel(doc)
    report_date = detect_report_date(doc)
    me_total_hrs, me_this_month = detect_me_totals(doc)

    me_grid = find_me_table(doc)
    components, me_warn = parse_me_table(me_grid)
    warnings.extend(me_warn)

    other_equipment = parse_other_equipment(doc)

    confidence = 0.0
    if vessel_name != "UNKNOWN":
        confidence += 0.2
    if report_date:
        confidence += 0.1
    if len(components) >= 8:
        confidence += 0.35
    if any(c["category"] == "MAINENGINE" for c in components):
        confidence += 0.2
    if any(c.get("periodicity") for c in components):
        confidence += 0.1
    if len(warnings) == 0:
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


def save_parsed(parsed: Dict, filename: str, file_hash: str):
    confidence = float(parsed.get("parse_confidence", 0))
    component_count = len(parsed.get("components", []))
    warning_count = len(parsed.get("warnings", []))
    vessel_name = parsed["vessel_name"]
    now = datetime.utcnow().isoformat() + "Z"

    if confidence < 0.45 or component_count < 6:
        raise ValueError(f"Commit blocked: parse confidence too low ({confidence}) or too few rows ({component_count}).")

    conn = get_db()
    c = conn.cursor()
    try:
        c.execute("INSERT OR IGNORE INTO vessels(name, created_at) VALUES (?, ?)", (vessel_name, now))
        c.execute(
            """
            INSERT INTO upload_log(
                vessel_name, filename, file_hash, report_date,
                me_total_hrs, me_this_month, parse_confidence,
                component_count, warning_count, uploaded_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                vessel_name, filename, file_hash, parsed.get("report_date"),
                parsed.get("me_total_hrs"), parsed.get("me_this_month"),
                confidence, component_count, warning_count, now
            ),
        )

        c.execute("DELETE FROM components WHERE vessel_name = ?", (vessel_name,))
        for x in parsed["components"]:
            c.execute(
                """
                INSERT INTO components(
                    vessel_name, category, engine_label, unit, description,
                    periodicity, last_oh_date, last_oh_hrs, hrs_since,
                    pct_used, status, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    vessel_name, x["category"], x["engine_label"], x["unit"], x["description"],
                    x.get("periodicity"), x.get("last_oh_date"), x.get("last_oh_hrs"), x.get("hrs_since"),
                    x.get("pct_used"), x.get("status"), now
                ),
            )

        c.execute("DELETE FROM other_equipment WHERE vessel_name = ?", (vessel_name,))
        for x in parsed["other_equipment"]:
            c.execute(
                """
                INSERT INTO other_equipment(vessel_name, section, description, periodicity, last_date, run_hrs, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    vessel_name, x["section"], x["description"], x.get("periodicity", ""),
                    x.get("last_date", ""), x.get("run_hrs", ""), now,
                ),
            )

        conn.commit()
    except Exception:
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
    df = pd.read_sql_query("SELECT * FROM components WHERE vessel_name = ?", conn, params=(vessel_name,))
    conn.close()
    return df


@st.cache_data(ttl=10)
def get_other_equipment(vessel_name: str):
    conn = get_db()
    df = pd.read_sql_query("SELECT * FROM other_equipment WHERE vessel_name = ? ORDER BY section, description", conn, params=(vessel_name,))
    conn.close()
    return df


@st.cache_data(ttl=10)
def get_history(vessel_name: str):
    conn = get_db()
    df = pd.read_sql_query(
        """
        SELECT filename, report_date, me_total_hrs, me_this_month, parse_confidence, component_count, warning_count, uploaded_at
        FROM upload_log
        WHERE vessel_name = ?
        ORDER BY uploaded_at DESC
        LIMIT 20
        """,
        conn,
        params=(vessel_name,),
    )
    conn.close()
    return df


@st.cache_data(ttl=10)
def get_summary():
    conn = get_db()
    df = pd.read_sql_query(
        """
        SELECT vessel_name,
               SUM(CASE WHEN status='OVERDUE' THEN 1 ELSE 0 END) AS overdue,
               SUM(CASE WHEN status='HIGH PRIORITY' THEN 1 ELSE 0 END) AS high_priority,
               SUM(CASE WHEN status='OK' THEN 1 ELSE 0 END) AS ok,
               SUM(CASE WHEN status='NO DATA' THEN 1 ELSE 0 END) AS no_data,
               COUNT(*) AS total
        FROM components
        GROUP BY vessel_name
        ORDER BY overdue DESC, high_priority DESC, vessel_name ASC
        """,
        conn,
    )
    conn.close()
    return df


@st.cache_data(ttl=10)
def get_all_components():
    conn = get_db()
    df = pd.read_sql_query("SELECT * FROM components", conn)
    conn.close()
    return df


STATUS_THEME = {
    "OVERDUE": {"bg": "#2d1014", "fg": "#ff7a7a", "mid": "#ff9c9c", "dim": "#6f3940"},
    "HIGH PRIORITY": {"bg": "#2f1d0d", "fg": "#ffad5a", "mid": "#ffc27d", "dim": "#73533a"},
    "OK": {"bg": "#10261c", "fg": "#39d98a", "mid": "#97f0c3", "dim": "#395a49"},
    "NO DATA": {"bg": "#142033", "fg": "#6ba1ff", "mid": "#9dbfff", "dim": "#41536f"},
}


def cyl_sort(x: str) -> int:
    m = re.search(r"(\d+)", str(x))
    return int(m.group(1)) if m else 999


def build_display_df(df: pd.DataFrame, priority: bool = False) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Status", "Vessel", "Component", "Engine", "Unit", "Periodicity", "Last OH", "Hrs Since", "Used"])

    d = df.copy()
    if priority:
        order_map = {"OVERDUE": 0, "HIGH PRIORITY": 1, "OK": 2, "NO DATA": 3}
        d["_ord"] = d["status"].map(lambda s: order_map.get(str(s), 9))
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
        theme = STATUS_THEME.get(str(row.get("Status")), STATUS_THEME["NO DATA"])
        return [
            f"background-color:{theme['bg']}; color:{theme['fg']}; font-weight:700",
            f"background-color:{theme['bg']}; color:{theme['mid']}; font-weight:600",
            f"background-color:{theme['bg']}; color:{theme['dim']}",
            f"background-color:{theme['bg']}; color:{theme['dim']}",
            f"background-color:{theme['bg']}; color:{theme['dim']}",
            f"background-color:{theme['bg']}; color:{theme['dim']}",
            f"background-color:{theme['bg']}; color:{theme['mid']}; font-weight:600",
            f"background-color:{theme['bg']}; color:{theme['fg']}; font-weight:700",
            f"background-color:{theme['bg']}; color:{theme['fg']}; font-weight:700",
        ]
    return df.style.apply(style_row, axis=1)


COLCFG = {
    "Status": st.column_config.TextColumn("Status", width=130),
    "Vessel": st.column_config.TextColumn("Vessel", width=130),
    "Component": st.column_config.TextColumn("Component", width=230),
    "Engine": st.column_config.TextColumn("Engine", width=90),
    "Unit": st.column_config.TextColumn("Unit", width=85),
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
    h = height or min(920, 38 * len(tbl) + 44)
    st.dataframe(apply_style(tbl), use_container_width=True, hide_index=True, height=h, column_config=COLCFG)


with st.sidebar:
    st.markdown("<h2 style='margin-bottom:.15rem'>FLEET MONITOR</h2>", unsafe_allow_html=True)
    st.markdown("<div class='smallcap' style='margin-bottom:1rem'>Running Hours Intelligence</div>", unsafe_allow_html=True)

    page = st.selectbox("Navigation", ["Fleet Overview", "Vessel Detail", "Upload Report", "Upload History"], label_visibility="collapsed")
    vessels = get_vessels()
    selected_vessel = st.selectbox("Active Vessel", vessels if vessels else ["—"], disabled=not bool(vessels))

    st.markdown("<hr>", unsafe_allow_html=True)
    db_kb = DB_PATH.stat().st_size / 1024 if DB_PATH.exists() else 0
    st.markdown(f"<div class='code'>db {db_kb:.0f} kb · {len(vessels)} vessels · hardened tec004 parser</div>", unsafe_allow_html=True)


if page == "Fleet Overview":
    st.markdown(hero("Fleet Master Matrix", "Universal Fleet Telemetry"), unsafe_allow_html=True)
    summary = get_summary()
    all_df = get_all_components()

    if summary.empty or all_df.empty:
        st.info("No data loaded. Upload a report to begin.")
        st.stop()

    cols = st.columns(5)
    with cols[0]: st.markdown(kpi(len(summary), "Vessels", "blue"), unsafe_allow_html=True)
    with cols[1]: st.markdown(kpi(len(all_df), "Components", "gold"), unsafe_allow_html=True)
    with cols[2]: st.markdown(kpi(int((all_df['status'] == 'OVERDUE').sum()), "Overdue", "red"), unsafe_allow_html=True)
    with cols[3]: st.markdown(kpi(int((all_df['status'] == 'HIGH PRIORITY').sum()), "High Priority", "orange"), unsafe_allow_html=True)
    with cols[4]: st.markdown(kpi(int((all_df['status'] == 'OK').sum()), "OK", "green"), unsafe_allow_html=True)

    st.markdown("<div class='smallcap' style='margin:1rem 0 .5rem'>Universal Component Control Grid</div>", unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns([1.4, 1.4, 1.6, 2.2])
    with c1:
        vessel_filter = st.selectbox("Filter Vessel", ["All Fleet"] + sorted(all_df["vessel_name"].dropna().unique().tolist()))
    with c2:
        category_filter = st.selectbox("Filter Category", ["All", "Main Engine", "Aux Engines"])
    with c3:
        status_filter = st.selectbox("Filter Status", ["All Statuses", "Critical Focus", "Overdue Only", "High Priority Only", "OK Only", "No Data Only"])
    with c4:
        component_filter = st.selectbox("Search Component", ["All"] + sorted(all_df["description"].dropna().unique().tolist()))

    filt = all_df.copy()
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
        f"<div class='code' style='margin:.4rem 0 1rem'>showing {len(filt)} records · {int((filt['status']=='OVERDUE').sum())} overdue · {int((filt['status']=='HIGH PRIORITY').sum())} high priority · {int((filt['status']=='OK').sum())} ok</div>",
        unsafe_allow_html=True,
    )

    if filt.empty:
        st.warning("No records match current filters.")
    else:
        render_table(filt, priority=True)


elif page == "Vessel Detail":
    if not vessels or selected_vessel == "—":
        st.info("Select a vessel from the sidebar.")
        st.stop()

    st.markdown(hero(selected_vessel, "Component Analysis"), unsafe_allow_html=True)
    df = get_components(selected_vessel)
    oe = get_other_equipment(selected_vessel)

    if df.empty:
        st.info("No data for this vessel.")
        st.stop()

    cols = st.columns(5)
    with cols[0]: st.markdown(kpi(len(df), "Total", "gold"), unsafe_allow_html=True)
    with cols[1]: st.markdown(kpi(int((df['status'] == 'OVERDUE').sum()), "Overdue", "red"), unsafe_allow_html=True)
    with cols[2]: st.markdown(kpi(int((df['status'] == 'HIGH PRIORITY').sum()), "High Priority", "orange"), unsafe_allow_html=True)
    with cols[3]: st.markdown(kpi(int((df['status'] == 'OK').sum()), "OK", "green"), unsafe_allow_html=True)
    with cols[4]: st.markdown(kpi(int((df['status'] == 'NO DATA').sum()), "No Data", "blue"), unsafe_allow_html=True)

    hist = get_history(selected_vessel)
    if not hist.empty:
        last = hist.iloc[0]
        st.markdown(
            f"""
            <div class='panel' style='margin:.8rem 0 1rem'>
              <div class='smallcap'>Latest Accepted Upload</div>
              <div class='code' style='margin-top:.55rem'>
                file: <b>{last['filename']}</b><br>
                report: <b>{last['report_date'] or '-'}</b><br>
                me total: <b>{int(last['me_total_hrs']) if pd.notna(last['me_total_hrs']) else '-'}</b> · this month: <b>{int(last['me_this_month']) if pd.notna(last['me_this_month']) else '-'}</b><br>
                confidence: <b>{float(last['parse_confidence']):.2f}</b> · components: <b>{int(last['component_count'])}</b> · warnings: <b>{int(last['warning_count'])}</b><br>
                uploaded: <b>{str(last['uploaded_at'])[:16]}</b>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    tabs = st.tabs(["Alerts", "Main Engine", "Other Equipment"])
    with tabs[0]:
        alerts = df[df["status"].isin(["OVERDUE", "HIGH PRIORITY"])]
        if alerts.empty:
            st.success("All machinery components are within acceptable bounds.")
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
        if oe.empty:
            st.info("No Other Equipment data available.")
        else:
            st.dataframe(
                oe.rename(columns={
                    "section": "Section",
                    "description": "Machinery Description",
                    "periodicity": "Maintenance Periodicity",
                    "last_date": "Inspection Date",
                    "run_hrs": "Logged Hours",
                }),
                use_container_width=True,
                hide_index=True,
                height=500,
            )


elif page == "Upload Report":
    st.markdown(hero("Upload Report", "TEC-004 Log Processing"), unsafe_allow_html=True)
    left, right = st.columns([1.7, 1.1], gap="large")

    with left:
        uploaded = st.file_uploader("Upload file", type=["doc"], label_visibility="collapsed")

    with right:
        st.markdown(
            """
            <div class='panel'>
              <div class='smallcap'>Accepted Specification</div>
              <div style='margin-top:.55rem; line-height:1.8'>
                TEC-004 Running Hours Monthly Log Report<br>
                Native <b>.doc</b> binary streams only
              </div>
              <div class='smallcap' style='margin-top:1rem'>Hardened Extraction Rules</div>
              <div style='margin-top:.5rem; line-height:1.8'>
                Main-engine paired-row parsing<br>
                Marker-cell protection for 1 / 2 rows<br>
                Thousands-aware periodicity parsing<br>
                Safe commit blocking on low confidence<br>
                Audit trail for every upload
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    if uploaded:
        raw = uploaded.read()
        file_hash = hashlib.md5(raw).hexdigest()

        with st.spinner("Converting report and executing hardened extraction pipeline..."):
            try:
                docx_bytes = convert_doc_to_docx(raw)
                parsed = parse_doc_bytes(docx_bytes)
            except Exception as e:
                st.error(f"Extraction failed: {e}")
                st.stop()

        comps = parsed["components"]
        nc = len(comps)
        overdue = sum(1 for c in comps if c["status"] == "OVERDUE")
        highp = sum(1 for c in comps if c["status"] == "HIGH PRIORITY")
        ok = sum(1 for c in comps if c["status"] == "OK")
        nod = sum(1 for c in comps if c["status"] == "NO DATA")
        conf = parsed["parse_confidence"]

        cols = st.columns(6)
        cols[0].metric("Asset", parsed["vessel_name"])
        cols[1].metric("Report Date", parsed["report_date"] or "-")
        cols[2].metric("ME Total", f"{parsed['me_total_hrs']:,}" if parsed["me_total_hrs"] else "-")
        cols[3].metric("This Month", f"{parsed['me_this_month']:,}" if parsed["me_this_month"] else "-")
        cols[4].metric("Rows", nc)
        cols[5].metric("Confidence", f"{conf:.2f}")

        st.markdown(
            f"<div class='panel' style='margin:1rem 0'><div class='smallcap'>Parse Health</div><div class='code' style='margin-top:.55rem'>overdue={overdue} · high_priority={highp} · ok={ok} · no_data={nod} · other_equipment={len(parsed['other_equipment'])}</div></div>",
            unsafe_allow_html=True,
        )

        if conf >= 0.80:
            st.success("High-confidence parse. Safe to review and commit.")
        elif conf >= 0.55:
            st.warning("Moderate-confidence parse. Review before commit.")
        else:
            st.error("Low-confidence parse. Commit blocked.")

        if parsed["warnings"]:
            for w in parsed["warnings"]:
                st.warning(w)

        st.markdown("<div class='smallcap' style='margin:1rem 0 .5rem'>Extracted Telemetry Preview</div>", unsafe_allow_html=True)
        preview_df = pd.DataFrame(comps) if comps else pd.DataFrame()
        tabs = st.tabs(["Main Engine Matrix", "Other Equipment"])
        with tabs[0]:
            if not preview_df.empty:
                me = preview_df[preview_df["category"] == "MAINENGINE"]
                if not me.empty:
                    render_table(me, height=440, priority=True)
                else:
                    st.info("No Main Engine telemetry extracted.")
            else:
                st.info("No component data available.")
        with tabs[1]:
            if parsed["other_equipment"]:
                st.dataframe(pd.DataFrame(parsed["other_equipment"]), use_container_width=True, hide_index=True, height=420)
            else:
                st.info("No Other Equipment data extracted.")

        cbtn, _ = st.columns([1.2, 4])
        with cbtn:
            disabled = conf < 0.45 or nc < 6 or parsed["vessel_name"] == "UNKNOWN"
            if st.button("COMMIT STREAM TO DATABASE", use_container_width=True, disabled=disabled):
                try:
                    save_parsed(parsed, uploaded.name, file_hash)
                    for fn in [get_vessels, get_components, get_other_equipment, get_history, get_summary, get_all_components]:
                        fn.clear()
                    st.success(f"Telemetry committed. {parsed['vessel_name']} stored with {nc} component rows.")
                    st.balloons()
                except Exception as e:
                    st.error(str(e))


elif page == "Upload History":
    st.markdown(hero("Upload History", "System Audit Trails"), unsafe_allow_html=True)
    if not vessels or selected_vessel == "—":
        st.info("Select a vessel from the sidebar.")
        st.stop()
    hist = get_history(selected_vessel)
    if hist.empty:
        st.info("No upload history recorded for this vessel.")
    else:
        d = hist.copy()
        d["me_total_hrs"] = d["me_total_hrs"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
        d["me_this_month"] = d["me_this_month"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
        d = d.rename(columns={
            "filename": "Logged Filename",
            "report_date": "Extracted Date",
            "me_total_hrs": "ME Combined Total",
            "me_this_month": "ME Monthly Increment",
            "parse_confidence": "Parse Confidence",
            "component_count": "Component Count",
            "warning_count": "Warning Count",
            "uploaded_at": "Transaction Timestamp",
        })
        st.dataframe(d, use_container_width=True, hide_index=True, height=520)
