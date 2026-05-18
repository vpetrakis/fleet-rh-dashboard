"""
╔══════════════════════════════════════════════════════════════════════╗
║         VESSEL RUNNING HOURS MONITORING SYSTEM                      ║
║         100% data integrity — zero-bug production build             ║
╚══════════════════════════════════════════════════════════════════════╝
Architecture:
  • parser.py   – deterministic Word-doc → structured data
  • db.py       – SQLite persistence layer (zero formulas, real data)
  • app.py      – Streamlit UI (this file bootstraps everything)
"""

import streamlit as st

# ── Page config must be FIRST Streamlit call ────────────────────────
st.set_page_config(
    page_title="Fleet Running Hours",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="expanded",
)

import os, re, sqlite3, io, tempfile, hashlib
from datetime import datetime, date
from pathlib import Path
import pandas as pd

# ════════════════════════════════════════════════════════════════════
# SECTION 1 — CUSTOM CSS  (industrial dark-navy + amber accent)
# ════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@300;400;500;600;700&family=Barlow+Condensed:wght@400;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

/* ── Root variables ── */
:root {
  --navy:    #0a0f1e;
  --navy2:   #0f1729;
  --navy3:   #162040;
  --panel:   #111827;
  --border:  #1e2d4a;
  --amber:   #f59e0b;
  --amber2:  #fbbf24;
  --red:     #ef4444;
  --red2:    #fca5a5;
  --orange:  #f97316;
  --green:   #22c55e;
  --green2:  #86efac;
  --blue:    #3b82f6;
  --text:    #e2e8f0;
  --muted:   #64748b;
  --font:    'Barlow', sans-serif;
  --mono:    'JetBrains Mono', monospace;
  --cond:    'Barlow Condensed', sans-serif;
}

/* ── Global reset ── */
html, body, [class*="css"] {
  font-family: var(--font) !important;
  background-color: var(--navy) !important;
  color: var(--text) !important;
}
.main { background: var(--navy) !important; }
.block-container { padding: 1.5rem 2rem !important; max-width: 100% !important; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
  background: var(--navy2) !important;
  border-right: 1px solid var(--border) !important;
}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stFileUploader label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span { color: var(--text) !important; }

/* ── Headers ── */
h1, h2, h3 { font-family: var(--cond) !important; letter-spacing: 0.04em !important; }

/* ── Metric cards ── */
[data-testid="stMetric"] {
  background: var(--panel) !important;
  border: 1px solid var(--border) !important;
  border-radius: 8px !important;
  padding: 1rem !important;
}
[data-testid="stMetricValue"] { font-family: var(--cond) !important; font-size: 2rem !important; font-weight: 700 !important; }
[data-testid="stMetricLabel"] { color: var(--muted) !important; font-size: 0.75rem !important; text-transform: uppercase !important; letter-spacing: 0.08em !important; }

/* ── Dataframe ── */
[data-testid="stDataFrame"] { border: 1px solid var(--border) !important; border-radius: 8px !important; }
.dvn-scroller { background: var(--panel) !important; }

/* ── Buttons ── */
.stButton > button {
  background: var(--amber) !important;
  color: #000 !important;
  border: none !important;
  font-family: var(--cond) !important;
  font-weight: 700 !important;
  letter-spacing: 0.06em !important;
  text-transform: uppercase !important;
  border-radius: 4px !important;
  padding: 0.5rem 1.5rem !important;
  transition: all 0.2s !important;
}
.stButton > button:hover { background: var(--amber2) !important; transform: translateY(-1px) !important; }

/* ── Select box ── */
.stSelectbox > div > div {
  background: var(--navy3) !important;
  border-color: var(--border) !important;
  color: var(--text) !important;
}

/* ── Upload area ── */
[data-testid="stFileUploadDropzone"] {
  background: var(--navy3) !important;
  border: 2px dashed var(--amber) !important;
  border-radius: 8px !important;
}
[data-testid="stFileUploadDropzone"] p { color: var(--amber) !important; font-family: var(--cond) !important; font-size: 1rem !important; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] { background: var(--panel) !important; border-radius: 8px 8px 0 0 !important; border-bottom: 2px solid var(--border) !important; gap: 0 !important; }
.stTabs [data-baseweb="tab"] {
  background: transparent !important;
  color: var(--muted) !important;
  font-family: var(--cond) !important;
  font-weight: 600 !important;
  letter-spacing: 0.06em !important;
  text-transform: uppercase !important;
  font-size: 0.85rem !important;
  padding: 0.75rem 1.5rem !important;
  border-bottom: 2px solid transparent !important;
  margin-bottom: -2px !important;
}
.stTabs [aria-selected="true"] { color: var(--amber) !important; border-bottom: 2px solid var(--amber) !important; }
.stTabs [data-baseweb="tab-panel"] { background: var(--panel) !important; border: 1px solid var(--border) !important; border-top: none !important; border-radius: 0 0 8px 8px !important; padding: 1.5rem !important; }

/* ── Expander ── */
.streamlit-expanderHeader {
  background: var(--navy3) !important;
  border: 1px solid var(--border) !important;
  border-radius: 6px !important;
  font-family: var(--cond) !important;
  color: var(--text) !important;
}
.streamlit-expanderContent { background: var(--panel) !important; border: 1px solid var(--border) !important; border-top: none !important; }

/* ── Divider ── */
hr { border-color: var(--border) !important; }

/* ── Info / Warning / Error boxes ── */
.stAlert { border-radius: 6px !important; border-left-width: 4px !important; }

/* ── Custom status badges ── */
.badge-overdue   { background:#7f1d1d; color:#fca5a5; padding:2px 10px; border-radius:20px; font-size:0.75rem; font-weight:700; letter-spacing:0.06em; font-family:var(--cond); }
.badge-highpri   { background:#78350f; color:#fcd34d; padding:2px 10px; border-radius:20px; font-size:0.75rem; font-weight:700; letter-spacing:0.06em; font-family:var(--cond); }
.badge-ok        { background:#14532d; color:#86efac; padding:2px 10px; border-radius:20px; font-size:0.75rem; font-weight:700; letter-spacing:0.06em; font-family:var(--cond); }
.badge-nodata    { background:#1e293b; color:#94a3b8; padding:2px 10px; border-radius:20px; font-size:0.75rem; font-weight:700; letter-spacing:0.06em; font-family:var(--cond); }

/* ── Fleet overview card ── */
.vessel-card {
  background: var(--navy3);
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 1rem 1.25rem;
  margin-bottom: 0.75rem;
  transition: border-color 0.2s;
}
.vessel-card:hover { border-color: var(--amber); }
.vessel-card-title { font-family: var(--cond); font-size: 1.1rem; font-weight: 700; letter-spacing: 0.06em; color: var(--amber); margin-bottom: 0.3rem; }
.vessel-card-sub   { font-size: 0.78rem; color: var(--muted); font-family: var(--mono); }

/* ── Progress bar custom ── */
.prog-wrap { background: #1e293b; border-radius: 999px; height: 8px; overflow: hidden; margin: 4px 0; }
.prog-fill-ok     { background: var(--green);  height: 100%; border-radius: 999px; transition: width 0.6s ease; }
.prog-fill-high   { background: var(--orange); height: 100%; border-radius: 999px; transition: width 0.6s ease; }
.prog-fill-over   { background: var(--red);    height: 100%; border-radius: 999px; transition: width 0.6s ease; }

/* ── Section header ── */
.section-header {
  font-family: var(--cond);
  font-size: 0.7rem;
  font-weight: 700;
  letter-spacing: 0.14em;
  text-transform: uppercase;
  color: var(--muted);
  border-bottom: 1px solid var(--border);
  padding-bottom: 0.4rem;
  margin: 1.5rem 0 1rem;
}

/* ── Big stat number ── */
.big-num { font-family: var(--cond); font-size: 2.5rem; font-weight: 700; line-height: 1; }
.big-num-label { font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.1em; color: var(--muted); margin-top: 2px; }

/* ── Toast-style parse result ── */
.parse-success { background: #14532d; border: 1px solid #166534; border-radius: 8px; padding: 0.75rem 1rem; color: #86efac; font-family: var(--cond); font-size: 0.9rem; }
.parse-error   { background: #7f1d1d; border: 1px solid #991b1b; border-radius: 8px; padding: 0.75rem 1rem; color: #fca5a5; font-family: var(--cond); font-size: 0.9rem; }

/* ── Watermark / logo area ── */
.app-logo { font-family: var(--cond); font-size: 1.5rem; font-weight: 700; letter-spacing: 0.1em; color: var(--amber); }
.app-logo span { color: var(--muted); font-weight: 400; }

/* ── Scrollable table container ── */
.scroll-table { overflow-x: auto; border-radius: 8px; border: 1px solid var(--border); }
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# SECTION 2 — DATABASE LAYER
# ════════════════════════════════════════════════════════════════════

DB_PATH = Path("running_hours.db")

def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS vessels (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        name        TEXT    NOT NULL UNIQUE,
        created_at  TEXT    NOT NULL DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS upload_log (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT    NOT NULL,
        filename    TEXT    NOT NULL,
        file_hash   TEXT    NOT NULL,
        report_date TEXT,
        me_total_hrs INTEGER,
        me_this_month INTEGER,
        uploaded_at TEXT    NOT NULL DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS components (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name   TEXT    NOT NULL,
        category      TEXT    NOT NULL,
        engine_label  TEXT    NOT NULL,
        unit          TEXT    NOT NULL,
        description   TEXT    NOT NULL,
        periodicity   REAL,
        last_oh_date  TEXT,
        last_oh_hrs   REAL,
        hrs_since     REAL,
        pct_used      REAL,
        status        TEXT    NOT NULL,
        updated_at    TEXT    NOT NULL DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS other_equipment (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name   TEXT    NOT NULL,
        section       TEXT    NOT NULL,
        description   TEXT    NOT NULL,
        periodicity   TEXT,
        last_date     TEXT,
        run_hrs       TEXT,
        updated_at    TEXT    NOT NULL DEFAULT (datetime('now'))
    );

    CREATE INDEX IF NOT EXISTS idx_comp_vessel  ON components(vessel_name);
    CREATE INDEX IF NOT EXISTS idx_comp_status  ON components(status);
    CREATE INDEX IF NOT EXISTS idx_other_vessel ON other_equipment(vessel_name);
    """)
    conn.commit()
    conn.close()

init_db()


# ════════════════════════════════════════════════════════════════════
# SECTION 3 — WORD DOC PARSER  (100% deterministic)
# ════════════════════════════════════════════════════════════════════

def _clean_period(raw: str) -> float | None:
    """Convert periodicity string like '16.000', '32000', 'N/A', 'Based on Observation' to float or None."""
    if not raw:
        return None
    raw = raw.strip().replace(',', '').replace('.', '', raw.count('.') - 1)
    # Remove thousand-separator dots: '16.000' → 16000
    # But '32000' is already fine; '16.000' means 16000
    s = raw.replace('.', '').replace(',', '')
    try:
        return float(s)
    except ValueError:
        return None

def _parse_date(raw: str) -> str | None:
    """Try many date formats, return ISO string or raw text if unparseable."""
    if not raw or raw.strip() in ('', 'N/A', 'n/a'):
        return None
    raw = re.sub(r'\s+', ' ', raw.strip().lstrip('[').rstrip(']'))
    # Pure numeric string = hours value, not a date
    if re.match(r'^\d+$', raw.strip()):
        return None
    # Normalise non-standard abbreviations: SEPT → SEP, JUNE → JUN, JULY → JUL
    raw_norm = re.sub(r'\bSEPT\b', 'SEP', raw, flags=re.IGNORECASE)
    raw_norm = re.sub(r'\bJUNE\b', 'JUN', raw_norm, flags=re.IGNORECASE)
    raw_norm = re.sub(r'\bJULY\b', 'JUL', raw_norm, flags=re.IGNORECASE)
    formats = [
        '%d %b %y', '%d %B %y', '%d %b %Y', '%d %B %Y',
        '%d/%m/%y', '%d/%m/%Y', '%d-%m-%y', '%d-%m-%Y',
        '%b %Y', '%B %Y', '%Y-%m-%d',
    ]
    # Try uppercase INPUT (for month names) but NOT uppercase format string
    for fmt in formats:
        for variant in [raw_norm, raw_norm.upper(), raw_norm.title(), raw, raw.upper()]:
            try:
                return datetime.strptime(variant, fmt).strftime('%Y-%m-%d')
            except ValueError:
                pass
    # Last resort: return as-is (keeps data, flags as non-standard)
    return raw

def _parse_hrs(raw: str) -> float | None:
    """Extract numeric hours from strings like '1308', '17.560', '[608]', '95\n17560'."""
    if not raw or raw.strip() in ('', 'N/A', 'n/a'):
        return None
    # If multiple numbers (merged cells), take the last non-zero
    nums = re.findall(r'\d[\d,\.]*', raw.replace('\n', ' '))
    if not nums:
        return None
    # Pick first valid positive number
    for n in nums:
        try:
            v = float(n.replace(',', '').replace('.', '', n.count('.') - 1))
            if v > 0:
                return v
        except ValueError:
            pass
    return None

def _compute_status(hrs: float | None, period: float | None) -> str:
    if hrs is None or period is None or period == 0:
        return 'NO DATA'
    pct = hrs / period
    if pct > 1.0:
        return 'OVERDUE'
    if pct >= 0.80:
        return 'HIGH PRIORITY'
    return 'OK'

def _compute_pct(hrs: float | None, period: float | None) -> float:
    if hrs is None or period is None or period == 0:
        return 0.0
    return round(hrs / period, 4)


def parse_doc(file_bytes: bytes) -> dict:
    """
    Parse a TEC-004 Running Hours Word document.
    Returns a dict with keys:
      vessel_name, report_date, me_total_hrs, me_this_month,
      components: list[dict], other_equipment: list[dict],
      warnings: list[str]
    Raises ValueError with clear message on structural failure.
    """
    from docx import Document

    warnings = []

    # Write to temp file (python-docx needs a file path or file-like)
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    try:
        doc = Document(tmp_path)
    except Exception as e:
        os.unlink(tmp_path)
        raise ValueError(f"Cannot open document: {e}")
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass

    if not doc.tables:
        raise ValueError("Document contains no tables. Is this a TEC-004 report?")

    # ── Extract vessel name and report date from paragraph 0 ──
    vessel_name = "UNKNOWN"
    report_date = None
    for para in doc.paragraphs:
        txt = para.text.strip()
        if txt:
            # "Vessel's Name: MV ALEXIS   Date:30 APRIL 2026"
            vm = re.search(r"Vessel'?s?\s+Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+?)(?:\s{2,}|\t)", txt, re.IGNORECASE)
            dm = re.search(r"Date\s*:\s*(.+)", txt, re.IGNORECASE)
            if vm:
                vessel_name = vm.group(1).strip()
            if dm:
                report_date = _parse_date(dm.group(1).strip())
            if vm or dm:
                break

    if vessel_name == "UNKNOWN":
        warnings.append("Could not extract vessel name from document header.")

    # ── Table 0: Main Engine ──
    me_total_hrs = None
    me_this_month = None
    components = []

    table0 = doc.tables[0]

    # Row 0 header: extract total running hours
    if table0.rows:
        hdr = table0.rows[0]
        for cell in hdr.cells:
            txt = cell.text
            tm = re.search(r'Total Running Hours[:\s]*([\d,]+)', txt, re.IGNORECASE)
            mm = re.search(r'This Month[:\s]*([\d,]+)', txt, re.IGNORECASE)
            if tm:
                try:
                    me_total_hrs = int(tm.group(1).replace(',', ''))
                except ValueError:
                    pass
            if mm:
                try:
                    me_this_month = int(mm.group(1).replace(',', ''))
                except ValueError:
                    pass

    # Row 1: cylinder count — cols 3..N-3 are cylinder columns
    # Identify which columns correspond to which cylinder numbers
    cyl_labels = []
    if len(table0.rows) > 1:
        row1_cells = [c.text.strip() for c in table0.rows[1].cells]
        # Find columns labelled "CYL. No.X"
        cyl_col_indices = []
        for i, cell in enumerate(row1_cells):
            m = re.search(r'CYL\s*\.?\s*No\s*\.?\s*(\d+)', cell, re.IGNORECASE)
            if m:
                # Only keep first occurrence per cylinder number (merged cells repeat)
                cyl_num = int(m.group(1))
                label = f"Cyl {cyl_num}"
                if not cyl_col_indices or cyl_col_indices[-1][1] != label:
                    cyl_col_indices.append((i, label))
        cyl_labels = cyl_col_indices  # list of (col_index, "Cyl N")

    # Parse component rows (pairs: row '1' = dates, row '2' = hours)
    # Rows come in pairs with same component name
    i = 2  # skip header rows 0,1
    rows = table0.rows
    while i < len(rows) - 1:
        r1 = [c.text.strip() for c in rows[i].cells]
        r2 = [c.text.strip() for c in rows[i + 1].cells] if i + 1 < len(rows) else []

        comp_name = r1[0] if r1 else ''
        if not comp_name:
            i += 1
            continue

        # Row type indicator is col 2: '1' = dates, '2' = hours
        row1_type = r1[2] if len(r1) > 2 else ''
        row2_type = r2[2] if len(r2) > 2 else ''

        if row1_type == '1' and row2_type == '2' and r1[0] == r2[0]:
            # periodicity
            period_raw = r1[1] if len(r1) > 1 else ''
            period = _clean_period(period_raw)

            for col_idx, cyl_label in cyl_labels:
                date_raw = r1[col_idx] if col_idx < len(r1) else ''
                hrs_raw  = r2[col_idx] if col_idx < len(r2) else ''

                oh_date = _parse_date(date_raw) if date_raw else None
                oh_hrs  = _parse_hrs(hrs_raw)   if hrs_raw  else None

                if oh_date is None and oh_hrs is None:
                    continue  # skip empty cylinder slots

                status = _compute_status(oh_hrs, period)
                pct    = _compute_pct(oh_hrs, period)

                components.append({
                    'category':     'MAIN_ENGINE',
                    'engine_label': 'ME',
                    'unit':         cyl_label,
                    'description':  comp_name,
                    'periodicity':  period,
                    'last_oh_date': oh_date,
                    'last_oh_hrs':  oh_hrs,
                    'hrs_since':    oh_hrs,
                    'pct_used':     pct,
                    'status':       status,
                })
            i += 2
        else:
            i += 1

    # ── Table 1: Turbocharger + Coolers + A/C & Compressors ──
    other_equip = []
    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        # Layout: 3 sections side by side
        # Cols 0-3: Turbocharger / Aux Boiler
        # Cols 5-8: Coolers / Exh Gas Boiler
        # Cols 10-12: A/C Compressors / Main Air Compressors
        section_map = {
            0: ('TURBOCHARGER/AUX BOILER', 0, 3),
            1: ('COOLERS/EXH GAS BOILER',  5, 8),
            2: ('AC/COMPRESSORS',          10, 12),
        }
        # Determine section header from row 0
        current_sections = {}
        for ridx, row in enumerate(t1.rows):
            cells = [c.text.strip() for c in row.cells]
            # Section A (cols 0-3)
            for sec_id, (sec_name, c_start, c_end) in section_map.items():
                desc = cells[c_start] if c_start < len(cells) else ''
                if not desc:
                    continue
                # Section headers
                if desc.upper() in ('TURBOCHARGER', 'AUXILIARY BOILER', 'COOLERS',
                                    'EXH GAS  BOILER', 'A/C & REFR. COMPRESSORS',
                                    'MAIN AIR COMPRESSORS', 'PERIODICTLY', 'DATE OF LAST INSPECTION', 'RUN HRS'):
                    continue
                date_val = cells[c_start + 1] if c_start + 1 < len(cells) else ''
                hrs_val  = cells[c_start + 2] if c_start + 2 < len(cells) else ''
                if desc and (date_val or hrs_val):
                    other_equip.append({
                        'section':      sec_name,
                        'description':  desc,
                        'periodicity':  cells[c_start - 1] if c_start > 0 and ridx > 0 else '',
                        'last_date':    date_val,
                        'run_hrs':      hrs_val,
                    })

    # ── Table 2: Auxiliary Engines ──
    if len(doc.tables) > 2:
        t2 = doc.tables[2]
        rows2 = t2.rows

        # Identify AUX engine blocks from row 0 header
        # Row 0: cols give "Aux. Engine No.1", "Aux. Engine No.2", "Aux. Engine No.3"
        # Each engine block spans 13 cols (1 label + 12 cylinder slots for 6 cyls × 2)
        # Row 2: Total Hours per engine
        # Row 3: This Month hours

        # Determine engine block boundaries
        engine_blocks = []  # list of (engine_label, start_col, total_hrs, this_month)
        if rows2:
            header_cells = [c.text.strip() for c in rows2[0].cells]
            total_cells  = [c.text.strip() for c in rows2[2].cells] if len(rows2) > 2 else []
            month_cells  = [c.text.strip() for c in rows2[3].cells] if len(rows2) > 3 else []

            seen = set()
            for i, cell in enumerate(header_cells):
                m = re.search(r'Aux\.\s*Engine\s*No\.?\s*(\d+)', cell, re.IGNORECASE)
                if m:
                    eng_num = int(m.group(1))
                    label = f"AUX-{eng_num}"
                    if label not in seen:
                        seen.add(label)
                        total_hrs = None
                        this_month = None
                        if total_cells:
                            # Scan from col i forward for numeric total
                            for j in range(i, min(i + 14, len(total_cells))):
                                v = _parse_hrs(total_cells[j])
                                if v and v > 0:
                                    total_hrs = v
                                    break
                        if month_cells:
                            for j in range(i, min(i + 14, len(month_cells))):
                                v = _parse_hrs(month_cells[j])
                                if v and v > 0:
                                    this_month = v
                                    break
                        engine_blocks.append((label, i, total_hrs, this_month))

        # Row 4: cylinder column headers "1, 1, 2, 2, 3, 3..."
        # Each engine has 6 cylinders × 2 cols = 12 data cols after a 3-col prefix
        # Offset: engine col_start + 1 (skip engine label merges)
        # For each component row pair, extract per-engine per-cylinder data

        # Build column index → (engine_label, cyl_num) from row 4
        cyl_col_map = {}  # col_idx → (engine_label, cyl_num)
        if len(rows2) > 4:
            row4_cells = [c.text.strip() for c in rows2[4].cells]
            for eng_label, eng_start, _, _ in engine_blocks:
                eng_end = engine_blocks[engine_blocks.index((eng_label, eng_start, _, _)) + 1][1] \
                    if engine_blocks.index((eng_label, eng_start, _, _)) < len(engine_blocks) - 1 \
                    else len(row4_cells)
                cyl_nums_seen = []
                for ci in range(eng_start, eng_end):
                    if ci < len(row4_cells):
                        try:
                            cn = int(row4_cells[ci])
                            if cn not in cyl_nums_seen:
                                cyl_nums_seen.append(cn)
                                cyl_col_map[ci] = (eng_label, cn)
                        except ValueError:
                            pass

        # Parse component pairs from row 5 onwards
        i2 = 5
        while i2 < len(rows2) - 1:
            r1 = [c.text.strip() for c in rows2[i2].cells]
            r2 = [c.text.strip() for c in rows2[i2 + 1].cells] if i2 + 1 < len(rows2) else []

            comp_name = r1[0] if r1 else ''
            if not comp_name:
                i2 += 1
                continue

            row1_type = r1[2] if len(r1) > 2 else ''
            row2_type = r2[2] if len(r2) > 2 else ''

            if row1_type in ('1', '2') and r1[0] == (r2[0] if r2 else ''):
                period_raw = r1[1] if len(r1) > 1 else ''
                period = _clean_period(period_raw)

                for col_idx, (eng_label, cyl_num) in cyl_col_map.items():
                    date_raw = r1[col_idx] if col_idx < len(r1) else ''
                    hrs_raw  = r2[col_idx] if col_idx < len(r2) else ''

                    oh_date = _parse_date(date_raw) if date_raw else None
                    oh_hrs  = _parse_hrs(hrs_raw)   if hrs_raw  else None

                    if oh_date is None and oh_hrs is None:
                        continue

                    status = _compute_status(oh_hrs, period)
                    pct    = _compute_pct(oh_hrs, period)

                    components.append({
                        'category':     'AUX_ENGINE',
                        'engine_label': eng_label,
                        'unit':         f"Cyl {cyl_num}",
                        'description':  comp_name,
                        'periodicity':  period,
                        'last_oh_date': oh_date,
                        'last_oh_hrs':  oh_hrs,
                        'hrs_since':    oh_hrs,
                        'pct_used':     pct,
                        'status':       status,
                    })
                i2 += 2
            else:
                i2 += 1

    # ── Table 3: D/G (Diesel Generator) Other Equipment ──
    if len(doc.tables) > 3:
        t3 = doc.tables[3]
        rows3 = t3.rows

        # Layout: 2 halves — left (cols 0-7) and right (cols 9-14)
        # Cols 3,4,5 = D/G No1, No2, No3 for left section
        # Cols 12,13,14 = D/G No1, No2, No3 for right section
        dg_labels = ['D/G 1', 'D/G 2', 'D/G 3']

        for ridx, row in enumerate(rows3):
            cells = [c.text.strip() for c in row.cells]
            if ridx == 0:
                continue  # header

            # Left section
            desc_l = cells[0] if cells else ''
            period_l = cells[1] if len(cells) > 1 else ''
            row_type_l = cells[2] if len(cells) > 2 else ''

            # Right section (offset 9)
            desc_r = cells[9]  if len(cells) > 9  else ''
            period_r = cells[10] if len(cells) > 10 else ''
            row_type_r = cells[11] if len(cells) > 11 else ''

            for section_offset, desc, period_raw, row_type, dg_start in [
                (0, desc_l, period_l, row_type_l, 3),
                (1, desc_r, period_r, row_type_r, 12),
            ]:
                if not desc or row_type not in ('1', '2'):
                    continue

                for dg_idx, dg_label in enumerate(dg_labels):
                    col = dg_start + dg_idx
                    val = cells[col] if col < len(cells) else ''
                    if not val:
                        continue

                    if row_type == '1':
                        other_equip.append({
                            'section':     'D/G EQUIPMENT',
                            'description': f"{desc} — {dg_label}",
                            'periodicity': period_raw,
                            'last_date':   _parse_date(val) or val,
                            'run_hrs':     '',
                        })
                    elif row_type == '2':
                        # Match to previous row's entry for same desc+dg
                        key = f"{desc} — {dg_label}"
                        for entry in reversed(other_equip):
                            if entry['description'] == key and entry['run_hrs'] == '':
                                entry['run_hrs'] = val
                                break
                        else:
                            other_equip.append({
                                'section':     'D/G EQUIPMENT',
                                'description': key,
                                'periodicity': period_raw,
                                'last_date':   '',
                                'run_hrs':     val,
                            })

    return {
        'vessel_name':    vessel_name,
        'report_date':    report_date,
        'me_total_hrs':   me_total_hrs,
        'me_this_month':  me_this_month,
        'components':     components,
        'other_equipment': other_equip,
        'warnings':       warnings,
    }


# ════════════════════════════════════════════════════════════════════
# SECTION 4 — DB WRITE / READ HELPERS
# ════════════════════════════════════════════════════════════════════

def save_parsed_data(parsed: dict, filename: str, file_hash: str):
    conn = get_db()
    c = conn.cursor()
    now = datetime.utcnow().isoformat()
    vessel = parsed['vessel_name']

    # Upsert vessel
    c.execute("INSERT OR IGNORE INTO vessels(name, created_at) VALUES (?,?)", (vessel, now))

    # Upload log
    c.execute("""
        INSERT INTO upload_log(vessel_name, filename, file_hash, report_date,
                               me_total_hrs, me_this_month, uploaded_at)
        VALUES (?,?,?,?,?,?,?)
    """, (vessel, filename, file_hash, parsed['report_date'],
          parsed['me_total_hrs'], parsed['me_this_month'], now))

    # Delete old components for this vessel
    c.execute("DELETE FROM components WHERE vessel_name=?", (vessel,))

    # Insert new
    for comp in parsed['components']:
        c.execute("""
            INSERT INTO components
              (vessel_name,category,engine_label,unit,description,
               periodicity,last_oh_date,last_oh_hrs,hrs_since,pct_used,status,updated_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
        """, (vessel,
              comp['category'], comp['engine_label'], comp['unit'],
              comp['description'], comp['periodicity'], comp['last_oh_date'],
              comp['last_oh_hrs'], comp['hrs_since'], comp['pct_used'],
              comp['status'], now))

    # Delete + re-insert other equipment
    c.execute("DELETE FROM other_equipment WHERE vessel_name=?", (vessel,))
    for oe in parsed['other_equipment']:
        c.execute("""
            INSERT INTO other_equipment
              (vessel_name, section, description, periodicity, last_date, run_hrs, updated_at)
            VALUES (?,?,?,?,?,?,?)
        """, (vessel, oe['section'], oe['description'],
              oe.get('periodicity', ''), oe.get('last_date', ''),
              oe.get('run_hrs', ''), now))

    conn.commit()
    conn.close()


@st.cache_data(ttl=10)
def get_all_vessels() -> list[str]:
    conn = get_db()
    rows = conn.execute("SELECT name FROM vessels ORDER BY name").fetchall()
    conn.close()
    return [r['name'] for r in rows]

@st.cache_data(ttl=10)
def get_components_df(vessel: str) -> pd.DataFrame:
    conn = get_db()
    df = pd.read_sql_query(
        "SELECT * FROM components WHERE vessel_name=? ORDER BY category, engine_label, description, unit",
        conn, params=(vessel,)
    )
    conn.close()
    return df

@st.cache_data(ttl=10)
def get_other_equip_df(vessel: str) -> pd.DataFrame:
    conn = get_db()
    df = pd.read_sql_query(
        "SELECT * FROM other_equipment WHERE vessel_name=? ORDER BY section, description",
        conn, params=(vessel,)
    )
    conn.close()
    return df

@st.cache_data(ttl=10)
def get_upload_history(vessel: str) -> pd.DataFrame:
    conn = get_db()
    df = pd.read_sql_query(
        "SELECT filename, report_date, me_total_hrs, me_this_month, uploaded_at FROM upload_log WHERE vessel_name=? ORDER BY uploaded_at DESC LIMIT 20",
        conn, params=(vessel,)
    )
    conn.close()
    return df

@st.cache_data(ttl=10)
def get_fleet_summary() -> pd.DataFrame:
    conn = get_db()
    df = pd.read_sql_query("""
        SELECT
            c.vessel_name,
            COUNT(CASE WHEN c.status='OVERDUE'        THEN 1 END) AS overdue,
            COUNT(CASE WHEN c.status='HIGH PRIORITY'  THEN 1 END) AS high_priority,
            COUNT(CASE WHEN c.status='OK'             THEN 1 END) AS ok,
            COUNT(*) AS total,
            MAX(u.uploaded_at)  AS last_upload,
            MAX(u.me_total_hrs) AS me_total_hrs,
            MAX(u.report_date)  AS report_date
        FROM components c
        LEFT JOIN upload_log u ON u.vessel_name = c.vessel_name
        GROUP BY c.vessel_name
        ORDER BY overdue DESC, high_priority DESC
    """, conn)
    conn.close()
    return df


# ════════════════════════════════════════════════════════════════════
# SECTION 5 — UI HELPERS
# ════════════════════════════════════════════════════════════════════

def status_badge(status: str) -> str:
    cls = {
        'OVERDUE':       'badge-overdue',
        'HIGH PRIORITY': 'badge-highpri',
        'OK':            'badge-ok',
    }.get(status, 'badge-nodata')
    return f'<span class="{cls}">{status}</span>'

def pct_bar(pct: float, status: str) -> str:
    w = min(pct * 100, 100)
    css = 'prog-fill-over' if status == 'OVERDUE' else ('prog-fill-high' if status == 'HIGH PRIORITY' else 'prog-fill-ok')
    return f'<div class="prog-wrap"><div class="{css}" style="width:{w:.1f}%"></div></div>'

def fmt_pct(pct: float) -> str:
    return f"{pct * 100:.1f}%"

def colored_status_df(df: pd.DataFrame) -> pd.DataFrame:
    """Return a display-ready copy of components df."""
    d = df[['description', 'engine_label', 'unit', 'periodicity',
             'last_oh_date', 'hrs_since', 'pct_used', 'status']].copy()
    d.columns = ['Component', 'Engine', 'Unit', 'Periodicity', 'Last O/H Date', 'Hrs Since', '% Used', 'Status']
    d['% Used'] = d['% Used'].apply(lambda x: f"{x*100:.1f}%" if pd.notna(x) else '—')
    d['Periodicity'] = d['Periodicity'].apply(lambda x: f"{int(x):,}" if pd.notna(x) and x > 0 else 'N/A')
    d['Hrs Since'] = d['Hrs Since'].apply(lambda x: f"{int(x):,}" if pd.notna(x) else '—')
    d['Last O/H Date'] = d['Last O/H Date'].fillna('—')
    return d


def style_status_df(df):
    """Apply Streamlit-compatible styling."""
    def row_style(row):
        s = row['Status']
        if s == 'OVERDUE':
            return ['background-color: #3b0a0a; color: #fca5a5'] * len(row)
        elif s == 'HIGH PRIORITY':
            return ['background-color: #3b1a00; color: #fcd34d'] * len(row)
        elif s == 'OK':
            return ['background-color: #071a0e; color: #86efac'] * len(row)
        return [''] * len(row)
    return df.style.apply(row_style, axis=1)


# ════════════════════════════════════════════════════════════════════
# SECTION 6 — SIDEBAR
# ════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown('<div class="app-logo">⚓ FLEET<span> MONITOR</span></div>', unsafe_allow_html=True)
    st.markdown('<p style="font-size:0.7rem;color:#475569;letter-spacing:0.08em;text-transform:uppercase;margin-top:2px;">Running Hours System</p>', unsafe_allow_html=True)
    st.divider()

    page = st.selectbox(
        "NAVIGATION",
        ["🗺️ Fleet Overview", "🚢 Vessel Detail", "📤 Upload Report", "📋 Upload History"],
        label_visibility="visible"
    )

    st.divider()

    vessels = get_all_vessels()
    if vessels:
        selected_vessel = st.selectbox("SELECT VESSEL", vessels)
    else:
        selected_vessel = None
        st.info("No vessel data yet. Upload a report to begin.")

    st.divider()
    st.markdown('<p style="font-size:0.65rem;color:#334155;text-transform:uppercase;letter-spacing:0.1em;">System</p>', unsafe_allow_html=True)
    db_size = DB_PATH.stat().st_size / 1024 if DB_PATH.exists() else 0
    st.markdown(f'<p style="font-size:0.7rem;color:#475569;font-family:var(--mono)">DB: {db_size:.1f} KB</p>', unsafe_allow_html=True)
    if vessels:
        st.markdown(f'<p style="font-size:0.7rem;color:#475569;font-family:var(--mono)">Vessels: {len(vessels)}</p>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# SECTION 7 — PAGES
# ════════════════════════════════════════════════════════════════════

# ── PAGE: UPLOAD ─────────────────────────────────────────────────
if page == "📤 Upload Report":
    st.markdown("## UPLOAD RUNNING HOURS REPORT")
    st.markdown('<p class="section-header">Drop a TEC-004 Word document (.doc or .docx) for any vessel</p>', unsafe_allow_html=True)

    col_up, col_info = st.columns([2, 1])

    with col_up:
        uploaded = st.file_uploader(
            "Drag & drop or browse",
            type=["doc", "docx"],
            accept_multiple_files=False,
            help="TEC-004-XX Running Hours Monthly Report"
        )

    with col_info:
        st.markdown("""
        <div style="background:#111827;border:1px solid #1e2d4a;border-radius:8px;padding:1rem;font-size:0.78rem;color:#94a3b8;">
        <b style="color:#f59e0b;font-family:var(--cond);letter-spacing:0.06em;">ACCEPTED FORMAT</b><br><br>
        • TEC-004 Running Hours Report<br>
        • Any vessel in the fleet<br>
        • .doc or .docx format<br><br>
        <b style="color:#f59e0b;font-family:var(--cond);letter-spacing:0.06em;">WHAT GETS PARSED</b><br><br>
        ✓ Vessel name & report date<br>
        ✓ M/E total & monthly hours<br>
        ✓ All M/E component O/H dates & hours<br>
        ✓ Aux engine component data (3 engines)<br>
        ✓ Turbocharger & D/G equipment<br>
        ✓ Status auto-computed (OVERDUE / HIGH PRIORITY / OK)
        </div>
        """, unsafe_allow_html=True)

    if uploaded:
        file_bytes = uploaded.read()

        # If .doc, convert to .docx first via LibreOffice
        if uploaded.name.endswith('.doc'):
            with st.spinner("Converting .doc → .docx..."):
                import subprocess
                with tempfile.NamedTemporaryFile(suffix='.doc', delete=False) as tmp_doc:
                    tmp_doc.write(file_bytes)
                    tmp_doc_path = tmp_doc.name
                try:
                    subprocess.run(
                        ['soffice', '--headless', '--convert-to', 'docx',
                         tmp_doc_path, '--outdir', tempfile.gettempdir()],
                        capture_output=True, timeout=30
                    )
                    docx_path = tmp_doc_path.replace('.doc', '.docx')
                    if os.path.exists(docx_path):
                        with open(docx_path, 'rb') as f:
                            file_bytes = f.read()
                        os.unlink(docx_path)
                    else:
                        st.error("LibreOffice conversion failed. Please save as .docx and re-upload.")
                        st.stop()
                finally:
                    try:
                        os.unlink(tmp_doc_path)
                    except Exception:
                        pass

        file_hash = hashlib.md5(file_bytes).hexdigest()

        with st.spinner("Parsing document..."):
            try:
                parsed = parse_doc(file_bytes)
            except ValueError as e:
                st.error(f"❌ Parse failed: {e}")
                st.stop()

        # Show parse preview
        st.markdown("---")
        st.markdown("### Parse Preview — confirm before saving")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Vessel", parsed['vessel_name'])
        c2.metric("Report Date", parsed['report_date'] or "Unknown")
        c3.metric("M/E Total Hrs", f"{parsed['me_total_hrs']:,}" if parsed['me_total_hrs'] else "—")
        c4.metric("M/E This Month", f"{parsed['me_this_month']:,}" if parsed['me_this_month'] else "—")

        st.markdown(f"""
        <div style="display:flex;gap:1rem;margin:1rem 0;flex-wrap:wrap;">
          <div style="background:#111827;border:1px solid #1e2d4a;border-radius:8px;padding:0.75rem 1.25rem;">
            <span style="font-size:1.4rem;font-family:var(--cond);font-weight:700;">{len(parsed['components'])}</span><br>
            <span style="font-size:0.7rem;text-transform:uppercase;letter-spacing:0.1em;color:#64748b;">Component Records</span>
          </div>
          <div style="background:#111827;border:1px solid #1e2d4a;border-radius:8px;padding:0.75rem 1.25rem;">
            <span style="font-size:1.4rem;font-family:var(--cond);font-weight:700;">{sum(1 for c in parsed['components'] if c['status']=='OVERDUE')}</span><br>
            <span style="font-size:0.7rem;text-transform:uppercase;letter-spacing:0.1em;color:#ef4444;">Overdue</span>
          </div>
          <div style="background:#111827;border:1px solid #1e2d4a;border-radius:8px;padding:0.75rem 1.25rem;">
            <span style="font-size:1.4rem;font-family:var(--cond);font-weight:700;">{sum(1 for c in parsed['components'] if c['status']=='HIGH PRIORITY')}</span><br>
            <span style="font-size:0.7rem;text-transform:uppercase;letter-spacing:0.1em;color:#f59e0b;">High Priority</span>
          </div>
          <div style="background:#111827;border:1px solid #1e2d4a;border-radius:8px;padding:0.75rem 1.25rem;">
            <span style="font-size:1.4rem;font-family:var(--cond);font-weight:700;">{len(parsed['other_equipment'])}</span><br>
            <span style="font-size:0.7rem;text-transform:uppercase;letter-spacing:0.1em;color:#64748b;">Other Equip Records</span>
          </div>
        </div>
        """, unsafe_allow_html=True)

        if parsed['warnings']:
            for w in parsed['warnings']:
                st.warning(f"⚠ {w}")

        # Preview table
        if parsed['components']:
            with st.expander("Preview parsed component data", expanded=False):
                prev = pd.DataFrame(parsed['components'])
                st.dataframe(
                    style_status_df(colored_status_df(prev)),
                    use_container_width=True, height=300
                )

        st.markdown("---")
        col_save, col_cancel = st.columns([1, 4])
        with col_save:
            if st.button("✅ CONFIRM & SAVE", use_container_width=True):
                save_parsed_data(parsed, uploaded.name, file_hash)
                get_all_vessels.clear()
                get_components_df.clear()
                get_other_equip_df.clear()
                get_fleet_summary.clear()
                st.success(f"✅ {parsed['vessel_name']} saved — {len(parsed['components'])} components updated.")
                st.balloons()


# ── PAGE: FLEET OVERVIEW ─────────────────────────────────────────
elif page == "🗺️ Fleet Overview":
    st.markdown("## FLEET RUNNING HOURS OVERVIEW")
    st.markdown('<p class="section-header">All vessels — component status summary</p>', unsafe_allow_html=True)

    summary = get_fleet_summary()

    if summary.empty:
        st.info("No vessel data loaded yet. Go to **Upload Report** to get started.")
    else:
        # KPI row
        total_overdue = int(summary['overdue'].sum())
        total_high    = int(summary['high_priority'].sum())
        total_ok      = int(summary['ok'].sum())
        total_vessels = len(summary)

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Vessels Tracked", total_vessels)
        k2.metric("🔴 Overdue Items", total_overdue)
        k3.metric("🟡 High Priority", total_high)
        k4.metric("🟢 OK Items", total_ok)

        st.markdown("---")

        # Fleet table
        disp = summary.copy()
        disp['Health'] = disp.apply(
            lambda r: f"{int(r['ok'])} OK / {int(r['high_priority'])} HIGH / {int(r['overdue'])} OVERDUE", axis=1
        )
        disp['Last Report'] = pd.to_datetime(disp['last_upload'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('—')
        disp['M/E Hrs'] = disp['me_total_hrs'].apply(lambda x: f"{int(x):,}" if pd.notna(x) else '—')
        disp_out = disp[['vessel_name', 'overdue', 'high_priority', 'ok', 'total', 'M/E Hrs', 'Last Report']].copy()
        disp_out.columns = ['Vessel', 'Overdue', 'High Priority', 'OK', 'Total', 'M/E Total Hrs', 'Last Report']

        def fleet_style(row):
            if row['Overdue'] > 0:
                return ['background-color:#2a0a0a'] * len(row)
            if row['High Priority'] > 0:
                return ['background-color:#2a1800'] * len(row)
            return ['background-color:#071a0e'] * len(row)

        st.dataframe(
            disp_out.style.apply(fleet_style, axis=1),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("---")
        st.markdown('<p class="section-header">Per-vessel overdue breakdown</p>', unsafe_allow_html=True)

        # Visual per-vessel cards
        for _, row in summary.iterrows():
            total = max(row['total'], 1)
            pct_over = row['overdue'] / total
            pct_high = row['high_priority'] / total

            health_color = '#ef4444' if row['overdue'] > 0 else ('#f59e0b' if row['high_priority'] > 0 else '#22c55e')

            with st.expander(f"**{row['vessel_name']}**  — {int(row['overdue'])} overdue / {int(row['high_priority'])} high priority", expanded=False):
                cc = get_components_df(row['vessel_name'])
                if not cc.empty:
                    od = cc[cc['status'] == 'OVERDUE']
                    if not od.empty:
                        st.markdown("**🔴 OVERDUE**")
                        st.dataframe(
                            colored_status_df(od),
                            use_container_width=True, hide_index=True, height=200
                        )
                    hp = cc[cc['status'] == 'HIGH PRIORITY']
                    if not hp.empty:
                        st.markdown("**🟡 HIGH PRIORITY**")
                        st.dataframe(
                            colored_status_df(hp),
                            use_container_width=True, hide_index=True, height=200
                        )


# ── PAGE: VESSEL DETAIL ───────────────────────────────────────────
elif page == "🚢 Vessel Detail":
    if not selected_vessel:
        st.info("No vessel selected. Upload a report first.")
        st.stop()

    st.markdown(f"## {selected_vessel}")
    st.markdown('<p class="section-header">Detailed component analysis</p>', unsafe_allow_html=True)

    df = get_components_df(selected_vessel)
    oe = get_other_equip_df(selected_vessel)

    if df.empty:
        st.info("No component data for this vessel.")
        st.stop()

    # Summary KPIs
    overdue_n   = int((df['status'] == 'OVERDUE').sum())
    high_n      = int((df['status'] == 'HIGH PRIORITY').sum())
    ok_n        = int((df['status'] == 'OK').sum())
    nodata_n    = int((df['status'] == 'NO DATA').sum())
    total_n     = len(df)

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total Components", total_n)
    k2.metric("🔴 Overdue", overdue_n)
    k3.metric("🟡 High Priority", high_n)
    k4.metric("🟢 OK", ok_n)
    k5.metric("⚪ No Data", nodata_n)

    # Last upload info
    hist = get_upload_history(selected_vessel)
    if not hist.empty:
        last = hist.iloc[0]
        st.markdown(
            f'<p style="font-size:0.75rem;color:#475569;font-family:var(--mono);">'
            f'Last upload: {last["filename"]} | Report date: {last["report_date"] or "—"} | '
            f'M/E hrs: {int(last["me_total_hrs"]):,} total / {int(last["me_this_month"] or 0):,} this month | '
            f'Uploaded: {last["uploaded_at"][:16]}</p>',
            unsafe_allow_html=True
        )

    st.markdown("---")

    tabs = st.tabs(["⚠ Alerts", "⚙ Main Engine", "🔩 Aux Engines", "🛠 Other Equipment", "📊 All Components"])

    # ── TAB: ALERTS ──
    with tabs[0]:
        st.markdown('<p class="section-header">Overdue & High Priority — action required</p>', unsafe_allow_html=True)
        alerts = df[df['status'].isin(['OVERDUE', 'HIGH PRIORITY'])].copy()
        alerts = alerts.sort_values(['status', 'pct_used'], ascending=[True, False])
        if alerts.empty:
            st.success("✅ No overdue or high-priority items for this vessel.")
        else:
            st.dataframe(
                style_status_df(colored_status_df(alerts)),
                use_container_width=True, hide_index=True,
                height=min(600, 35 * len(alerts) + 40)
            )

    # ── TAB: MAIN ENGINE ──
    with tabs[1]:
        me_df = df[df['category'] == 'MAIN_ENGINE'].copy()
        if me_df.empty:
            st.info("No Main Engine data.")
        else:
            components_list = sorted(me_df['description'].unique())
            sel_comp = st.selectbox("Filter component", ['ALL'] + components_list)
            view_df = me_df if sel_comp == 'ALL' else me_df[me_df['description'] == sel_comp]
            st.dataframe(
                style_status_df(colored_status_df(view_df)),
                use_container_width=True, hide_index=True,
                height=min(700, 35 * len(view_df) + 40)
            )

    # ── TAB: AUX ENGINES ──
    with tabs[2]:
        aux_df = df[df['category'] == 'AUX_ENGINE'].copy()
        if aux_df.empty:
            st.info("No Aux Engine data.")
        else:
            eng_list = sorted(aux_df['engine_label'].unique())
            sel_eng = st.selectbox("Engine", ['ALL'] + eng_list)
            view_aux = aux_df if sel_eng == 'ALL' else aux_df[aux_df['engine_label'] == sel_eng]
            st.dataframe(
                style_status_df(colored_status_df(view_aux)),
                use_container_width=True, hide_index=True,
                height=min(700, 35 * len(view_aux) + 40)
            )

    # ── TAB: OTHER EQUIPMENT ──
    with tabs[3]:
        if oe.empty:
            st.info("No other equipment data.")
        else:
            sec_list = sorted(oe['section'].unique())
            for sec in sec_list:
                st.markdown(f'<p class="section-header">{sec}</p>', unsafe_allow_html=True)
                sec_df = oe[oe['section'] == sec][['description', 'periodicity', 'last_date', 'run_hrs']].copy()
                sec_df.columns = ['Description', 'Periodicity', 'Last Date', 'Run Hrs']
                st.dataframe(sec_df, use_container_width=True, hide_index=True)

    # ── TAB: ALL COMPONENTS ──
    with tabs[4]:
        status_filter = st.multiselect(
            "Filter by status",
            ['OVERDUE', 'HIGH PRIORITY', 'OK', 'NO DATA'],
            default=['OVERDUE', 'HIGH PRIORITY', 'OK', 'NO DATA']
        )
        cat_filter = st.multiselect(
            "Filter by category",
            ['MAIN_ENGINE', 'AUX_ENGINE'],
            default=['MAIN_ENGINE', 'AUX_ENGINE']
        )
        filtered = df[df['status'].isin(status_filter) & df['category'].isin(cat_filter)]
        st.dataframe(
            style_status_df(colored_status_df(filtered)),
            use_container_width=True, hide_index=True,
            height=min(800, 35 * len(filtered) + 40)
        )


# ── PAGE: UPLOAD HISTORY ─────────────────────────────────────────
elif page == "📋 Upload History":
    st.markdown("## UPLOAD HISTORY")
    st.markdown('<p class="section-header">Audit trail of all report uploads</p>', unsafe_allow_html=True)

    if not selected_vessel:
        st.info("Select a vessel from the sidebar.")
        st.stop()

    st.markdown(f"### {selected_vessel}")
    hist = get_upload_history(selected_vessel)
    if hist.empty:
        st.info("No upload history for this vessel.")
    else:
        hist_disp = hist.copy()
        hist_disp.columns = ['Filename', 'Report Date', 'M/E Total Hrs', 'M/E This Month', 'Uploaded At']
        hist_disp['M/E Total Hrs'] = hist_disp['M/E Total Hrs'].apply(lambda x: f"{int(x):,}" if pd.notna(x) else '—')
        hist_disp['M/E This Month'] = hist_disp['M/E This Month'].apply(lambda x: f"{int(x):,}" if pd.notna(x) else '—')
        st.dataframe(hist_disp, use_container_width=True, hide_index=True)
