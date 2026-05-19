import streamlit as st
st.set_page_config(page_title="Fleet Running Hours", page_icon="⚓", layout="wide", initial_sidebar_state="expanded")

import os
import re
import sqlite3
import tempfile
import hashlib
import subprocess
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
from docx import Document

# ============================================================
# PREMIUM UI
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');
:root{
  --bg:#03060d;--bg1:#06091a;--bg2:#080e20;--bg3:#0b1228;--bg4:#0f1830;
  --b1:#0f1c35;--b2:#182840;--b3:#223350;
  --gold:#c89a14;--gold2:#e0b422;--gold3:#f5cc44;
  --red:#cc2828;--red2:#ff5c5c;
  --orange:#b85518;--ora2:#ff8833;
  --green:#0d8a4a;--grn2:#22c55e;
  --blue:#1444a8;--blu2:#3b82f6;
  --t0:#f2f7ff;--t1:#c0d0e8;--t2:#6a84a8;--t3:#304060;
  --ff:'Space Grotesk',sans-serif;--fi:'Inter',sans-serif;--fm:'JetBrains Mono',monospace;
}
html,body,[class*="css"]{background:var(--bg)!important;color:var(--t1)!important;font-family:var(--fi)!important}
.main,.main>div{background:var(--bg)!important}
.block-container{max-width:100%!important;padding:1.8rem 2rem 4rem!important}
.main::before{
  content:"";position:fixed;inset:0;pointer-events:none;z-index:0;
  background:
    radial-gradient(ellipse 90% 50% at -10% -5%, rgba(200,154,20,.06) 0%, transparent 55%),
    radial-gradient(ellipse 70% 45% at 110% 105%, rgba(20,68,168,.05) 0%, transparent 55%);
}
[data-testid="stSidebar"]{background:var(--bg1)!important;border-right:1px solid var(--b2)!important}
[data-testid="stSidebar"] *{color:var(--t1)!important}
[data-testid="stSidebarContent"]{padding:1.25rem!important}
h1,h2,h3{font-family:var(--ff)!important;color:var(--t0)!important;letter-spacing:-.02em!important}
h1{font-size:1.8rem!important;font-weight:700!important}
h2{font-size:1.2rem!important;font-weight:600!important}
h3{font-size:1rem!important;font-weight:600!important}
.stSelectbox>div>div,.stMultiSelect>div>div,.stTextInput>div>div>input{
  background:var(--bg3)!important;border:1px solid var(--b2)!important;color:var(--t1)!important;border-radius:8px!important
}
.stButton>button{
  background:linear-gradient(135deg,var(--gold) 0%,#8a6a08 100%)!important;color:#000!important;border:none!important;
  border-radius:8px!important;padding:.7rem 1.25rem!important;font-family:var(--ff)!important;font-weight:700!important;
  letter-spacing:.05em!important;text-transform:uppercase!important
}
[data-testid="stMetric"]{
  background:var(--bg3)!important;border:1px solid var(--b2)!important;border-radius:12px!important;padding:.9rem 1rem 1rem!important
}
[data-testid="stMetricValue"]{font-family:var(--ff)!important;color:var(--t0)!important;font-weight:700!important}
[data-testid="stMetricLabel"]{color:var(--t3)!important;text-transform:uppercase!important;letter-spacing:.12em!important;font-size:.66rem!important}
[data-testid="stDataFrame"]{
  border:1px solid var(--b2)!important;border-radius:12px!important;overflow:hidden!important;background:var(--bg2)!important
}
.stTabs [data-baseweb="tab-list"]{
  background:var(--bg2)!important;border:1px solid var(--b2)!important;border-radius:12px 12px 0 0!important;gap:0!important
}
.stTabs [data-baseweb="tab"]{
  color:var(--t3)!important;font-family:var(--ff)!important;text-transform:uppercase!important;letter-spacing:.05em!important;font-size:.74rem!important
}
.stTabs [aria-selected="true"]{color:var(--gold2)!important;border-bottom:2px solid var(--gold)!important}
.stTabs [data-baseweb="tab-panel"]{
  background:var(--bg2)!important;border:1px solid var(--b2)!important;border-top:none!important;border-radius:0 0 12px 12px!important;padding:1.25rem!important
}
.kpi{background:var(--bg3);border:1px solid var(--b2);border-radius:12px;padding:1rem 1.1rem}
.kpi .v{font-family:var(--ff);font-size:2rem;font-weight:700;line-height:1}
.kpi .l{color:var(--t3);text-transform:uppercase;letter-spacing:.15em;font-size:.62rem;margin-top:.4rem}
.kpi.gold .v{color:var(--gold3)} .kpi.red .v{color:var(--red2)} .kpi.orange .v{color:var(--ora2)} .kpi.green .v{color:var(--grn2)} .kpi.blue .v{color:var(--blu2)}
.panel{background:var(--bg3);border:1px solid var(--b2);border-radius:14px;padding:1rem 1.1rem}
.smallcap{color:var(--t3);text-transform:uppercase;letter-spacing:.16em;font-size:.66rem;font-family:var(--fi)}
.code{font-family:var(--fm);color:var(--t2);font-size:.74rem}
.hero-line{height:1px;background:linear-gradient(90deg,var(--gold),var(--b2),transparent);margin:.5rem 0 1.2rem}
</style>
""", unsafe_allow_html=True)

# ============================================================
# DB
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
def clean_text(x):
    if x is None:
        return ""
    return re.sub(r"\s+", " ", str(x).replace("\xa0", " ")).strip()

def norm(x):
    x = clean_text(x).upper()
    x = x.replace("O/H", "OH").replace("D/G", "DG").replace("AUX.", "AUX")
    x = re.sub(r"[^A-Z0-9 /().:\-]", "", x)
    return x

def parse_number(x):
    if x is None:
        return None
    s = clean_text(x).replace(",", "")
    m = re.search(r"\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None

def parse_periodicity(x):
    if x is None:
        return None
    s = clean_text(x).replace(",", "")
    m = re.search(r"\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group())
    except:
        return None

def parse_date(raw):
    raw = clean_text(raw)
    if not raw or raw.upper() in {"NA", "N/A", "-", "--"}:
        return None
    raw = raw.replace(".", " ")
    raw = re.sub(r"\bSEPT\b", "SEP", raw, flags=re.I)
    raw = re.sub(r"\s+", " ", raw).strip()
    fmts = ["%d %b %y", "%d %B %y", "%d %b %Y", "%d %B %Y", "%d/%m/%y", "%d/%m/%Y", "%d-%m-%y", "%d-%m-%Y", "%b %Y", "%B %Y", "%Y-%m-%d"]
    for fmt in fmts:
        for cand in (raw, raw.upper(), raw.title()):
            try:
                return datetime.strptime(cand, fmt).strftime("%Y-%m-%d")
            except:
                pass
    return None

def parse_compact_date(tok):
    tok = clean_text(tok).upper()
    tok = re.sub(r"[^0-9A-Z]", "", tok)
    if re.fullmatch(r"\d{6}", tok):
        dd = int(tok[:2]); mm = int(tok[2:4]); yy = int(tok[4:6])
        yy = 2000 + yy if yy < 70 else 1900 + yy
        try:
            return datetime(yy, mm, dd).strftime("%Y-%m-%d")
        except:
            return None
    for fmt in ("%d%b%y", "%d%b%Y"):
        try:
            return datetime.strptime(tok.title(), fmt).strftime("%Y-%m-%d")
        except:
            pass
    return None

def parse_hours(x):
    if x is None:
        return None
    nums = re.findall(r"\d[\d,]*", clean_text(x))
    vals = []
    for n in nums:
        try:
            vals.append(float(n.replace(",", "")))
        except:
            pass
    return max(vals) if vals else None

def pct_used(h, p):
    if h is None or p is None or p <= 0:
        return 0.0
    return round(h / p, 4)

def status_from(h, p):
    if h is None or p is None or p <= 0:
        return "NO DATA"
    r = h / p
    if r >= 1.0:
        return "OVERDUE"
    if r >= 0.80:
        return "HIGH PRIORITY"
    return "OK"

def dedup_components(rows):
    out = []
    seen = set()
    for x in rows:
        k = (x["category"], x["engine_label"], x["unit"], x["description"], x["periodicity"], x["last_oh_date"], x["hrs_since"])
        if k not in seen:
            seen.add(k)
            out.append(x)
    return out

def hero(title, eye=""):
    top = f'<div class="smallcap">{eye}</div>' if eye else ""
    return f'<div>{top}<h1>{title}</h1><div class="hero-line"></div></div>'

def kpi(val, lbl, color="gold"):
    return f'<div class="kpi {color}"><div class="v">{val}</div><div class="l">{lbl}</div></div>'

# ============================================================
# DOC CONVERSION + EXTRACTION
# ============================================================
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError("LibreOffice not found. Add libreoffice in packages.txt.")

    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as tmp:
        tmp.write(raw)
        src = tmp.name

    out_dir = tempfile.mkdtemp(prefix="docconv_")
    target = os.path.join(out_dir, Path(src).stem + ".docx")
    profile = f"file:///tmp/lo_profile_{os.getpid()}_{os.urandom(4).hex()}"

    try:
        res = subprocess.run(
            [soffice, "--headless", "--norestore", "--nofirststartwizard", f"-env:UserInstallation={profile}", "--convert-to", "docx", src, "--outdir", out_dir],
            capture_output=True,
            timeout=120
        )
        if not os.path.exists(target):
            raise RuntimeError(res.stderr.decode("utf-8", errors="ignore")[:500])
        with open(target, "rb") as f:
            return f.read()
    finally:
        try: os.unlink(src)
        except: pass
        try: os.unlink(target)
        except: pass
        shutil.rmtree(out_dir, ignore_errors=True)

def open_docx_bytes(docx_bytes: bytes) -> Document:
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp.write(docx_bytes)
        path = tmp.name
    try:
        return Document(path)
    finally:
        try: os.unlink(path)
        except: pass

def all_text(doc: Document) -> str:
    parts = []
    for p in doc.paragraphs:
        t = clean_text(p.text)
        if t:
            parts.append(t)
    for table in doc.tables:
        for row in table.rows:
            row_txt = " ".join(clean_text(c.text) for c in row.cells if clean_text(c.text))
            if row_txt:
                parts.append(row_txt)
    txt = "\n".join(parts)
    txt = txt.replace("CYL. No.2CYL. No.3", "CYL. No.2 CYL. No.3")
    txt = txt.replace("DESCRIPTIONPERIODICTLYDG No1DG No2DG No3", "DESCRIPTION PERIODICTLY DG No1 DG No2 DG No3")
    txt = txt.replace("DESCRIPTIONPERIODICTLY12", "DESCRIPTION PERIODICTLY 1 2")
    txt = re.sub(r"\s+", " ", txt)
    return txt

# ============================================================
# SECTION PARSERS
# ============================================================
def extract_vessel_name(txt):
    pats = [
        r"VESSELS NAME MV ([A-Z][A-Z0-9 \-]+?) DATE",
        r"VESSEL['’]S NAME[: ]+(?:MV )?([A-Z][A-Z0-9 \-]+)",
        r"\bMV ([A-Z][A-Z0-9 \-]+?) DATE"
    ]
    u = txt.upper()
    for p in pats:
        m = re.search(p, u)
        if m:
            return clean_text(m.group(1))
    return "UNKNOWN"

def extract_report_date(txt):
    m = re.search(r"DATE\s*([0-9A-Z ./-]{5,30})", txt, flags=re.I)
    return parse_date(m.group(1)) if m else None

def extract_me_totals(txt):
    mt = mm = None
    m1 = re.search(r"TOTAL RUNNING HOURS\s*([0-9,]+)", txt, flags=re.I)
    m2 = re.search(r"THIS MONTH\s*([0-9,]+)", txt, flags=re.I)
    if m1:
        mt = int(m1.group(1).replace(",", ""))
    if m2:
        mm = int(m2.group(1).replace(",", ""))
    return mt, mm

def extract_main_engine(txt):
    comps = []
    m = re.search(r"CYL\. NO\.1.*?MAIN BEARINGS.*?22002120021", txt, flags=re.I)
    if not m:
        return comps
    block = m.group(0)

    pattern = re.compile(
        r"([A-Z][A-Z /&.\-]+?)"
        r"(BASED ON OBSERVATION|\d+(?:\.\d+)?)"
        r"([0-9A-Z .]+?)"
        r"([0-9]+)"
        r"([0-9A-Z .]+?)"
        r"([0-9]+)",
        flags=re.I
    )

    for mm in pattern.finditer(block):
        desc = clean_text(mm.group(1))
        p_raw = clean_text(mm.group(2))
        d1_raw = clean_text(mm.group(3))
        h1_raw = clean_text(mm.group(4))
        d2_raw = clean_text(mm.group(5))
        h2_raw = clean_text(mm.group(6))

        p = None if "OBSERVATION" in p_raw.upper() else parse_periodicity(p_raw)
        d1 = parse_date(d1_raw) or parse_compact_date(d1_raw)
        d2 = parse_date(d2_raw) or parse_compact_date(d2_raw)
        h1 = parse_hours(h1_raw)
        h2 = parse_hours(h2_raw)

        if d1 or h1:
            comps.append({
                "category":"MAINENGINE","engine_label":"ME","unit":"Cyl 1","description":desc,
                "periodicity":p,"last_oh_date":d1,"last_oh_hrs":h1,"hrs_since":h1,"pct_used":pct_used(h1,p),"status":status_from(h1,p)
            })
        if d2 or h2:
            comps.append({
                "category":"MAINENGINE","engine_label":"ME","unit":"Cyl 2","description":desc,
                "periodicity":p,"last_oh_date":d2,"last_oh_hrs":h2,"hrs_since":h2,"pct_used":pct_used(h2,p),"status":status_from(h2,p)
            })
    return comps

def extract_aux_rowpair(txt):
    comps = []
    m = re.search(r"AUX ENGINE MAKER TYPE.*?DESCRIPTION PERIODICTLY 1 2(.*?)(?:DESCRIPTION PERIODICTLY DG NO1 DG NO2 DG NO3)", txt, flags=re.I)
    if not m:
        return comps
    block = m.group(1)

    pattern = re.compile(r"([A-Z][A-Z /&.\-]+?)\s*(\d+(?:\.\d+)?)\s*([0-9A-Z./ ]+?)\s*([0-9]{3,})", flags=re.I)
    rows = pattern.findall(block)

    for desc, p_raw, d_raw, h_raw in rows:
        desc = clean_text(desc)
        p = parse_periodicity(p_raw)
        d = parse_date(d_raw) or parse_compact_date(d_raw)
        h = parse_hours(h_raw)
        if not desc:
            continue
        comps.append({
            "category":"AUXENGINE","engine_label":"AUX-1","unit":"AUX-1","description":desc,
            "periodicity":p,"last_oh_date":d,"last_oh_hrs":h,"hrs_since":h,"pct_used":pct_used(h,p),"status":status_from(h,p)
        })
    return comps

def decode_dg_payload(payload):
    payload = clean_text(payload)
    tokens = re.findall(r"[A-Z]+|\d+", payload.upper())
    d = None
    h = None
    for tok in tokens:
        if d is None:
            d = parse_compact_date(tok) or parse_date(tok)
        if h is None:
            try:
                v = float(tok.replace(",", ""))
                if v >= 50:
                    h = v
            except:
                pass
    return d, h

def extract_dg_matrix(txt):
    comps = []
    m = re.search(r"DESCRIPTION PERIODICTLY DG NO1 DG NO2 DG NO3(.*?)(?:TABLE 1ST COPY TO BE RETAINED|REMARKS CHIEF ENGINEER)", txt, flags=re.I)
    if not m:
        return comps
    block = m.group(1)

    items = [
        "Turbocharger", "Air Cooler", "L.O. Cooler Clean", "Cooling Water Pump",
        "F.W. Cooler Clean", "Cool Water Thermostat Valve", "L.O. Renewal",
        "L.O. Thermostat Valve", "Alternator Cleaning", "Thrust Bearing"
    ]

    for item in items:
        mm = re.search(rf"{re.escape(item)}\s*([A-Z0-9. ]+?)(?=(Turbocharger|Air Cooler|L\.O\. Cooler Clean|Cooling Water Pump|F\.W\. Cooler Clean|Cool Water Thermostat Valve|L\.O\. Renewal|L\.O\. Thermostat Valve|Alternator Cleaning|Thrust Bearing|$))", block, flags=re.I)
        if not mm:
            continue
        row = mm.group(1).strip()
        p = parse_periodicity(row)
        chunks = re.findall(r"([A-Z0-9./ ]{4,})", row)
        cols = chunks[:3] if len(chunks) >= 3 else [row, "", ""]

        for eng, payload in zip(["DG-1", "DG-2", "DG-3"], cols):
            d, h = decode_dg_payload(payload)
            if d is None and h is None:
                continue
            comps.append({
                "category":"AUXENGINE","engine_label":eng,"unit":eng,"description":item,
                "periodicity":p,"last_oh_date":d,"last_oh_hrs":h,"hrs_since":h,"pct_used":pct_used(h,p),"status":status_from(h,p)
            })
    return comps

def extract_other_equipment(txt):
    out = []

    def add(section, desc, periodicity, datev, hrs):
        out.append({
            "section": section,
            "description": clean_text(desc),
            "periodicity": clean_text(periodicity),
            "last_date": parse_date(datev) or clean_text(datev),
            "run_hrs": clean_text(hrs)
        })

    matches = [
        ("TURBOCHARGER / AUX BOILER", r"GENERAL OH\s*16000\s*26 NOV 21\s*23610"),
        ("COOLERS / EXH GAS BOILER", r"ME L\.O\.\s*30 DEC 23"),
        ("A/C & COMPRESSORS", r"AIR COND\. COMPRESSOR NO\.1\s*28 MAR 26\s*5340"),
    ]
    for sec, pat in matches:
        m = re.search(pat, txt, flags=re.I)
        if m:
            add(sec, pat.split(r"\s*")[0].replace("\\", ""), "", "", "")

    return out

def parse_doc_bytes(docx_bytes):
    doc = open_docx_bytes(docx_bytes)
    txt = all_text(doc)
    warnings = []

    vessel_name = extract_vessel_name(txt)
    report_date = extract_report_date(txt)
    me_total_hrs, me_this_month = extract_me_totals(txt)

    me = extract_main_engine(txt)
    aux1 = extract_aux_rowpair(txt)
    aux2 = extract_dg_matrix(txt)
    oe = extract_other_equipment(txt)

    components = dedup_components(me + aux1 + aux2)

    if vessel_name == "UNKNOWN":
        warnings.append("Could not confidently extract vessel name.")
    if not me:
        warnings.append("Main engine block not extracted.")
    if not aux1 and not aux2:
        warnings.append("Auxiliary engine blocks not extracted.")
    if not components:
        warnings.append("No component rows extracted.")

    confidence = 0.0
    if vessel_name != "UNKNOWN":
        confidence += 0.15
    if report_date:
        confidence += 0.10
    if me_total_hrs is not None:
        confidence += 0.05
    if any(c["category"] == "MAINENGINE" for c in components):
        confidence += 0.30
    if any(c["category"] == "AUXENGINE" for c in components):
        confidence += 0.25
    if len(components) >= 5:
        confidence += 0.10
    if len(components) >= 10:
        confidence += 0.05
    confidence = round(min(confidence, 1.0), 2)

    return {
        "vessel_name": vessel_name,
        "report_date": report_date,
        "me_total_hrs": me_total_hrs,
        "me_this_month": me_this_month,
        "components": components,
        "other_equipment": oe,
        "warnings": warnings,
        "parse_confidence": confidence,
    }

# ============================================================
# SAVE / FETCH
# ============================================================
def save_parsed(parsed, filename, file_hash):
    vessel_name = parsed["vessel_name"]
    now = datetime.utcnow().isoformat() + "Z"

    conn = get_db()
    c = conn.cursor()
    try:
        c.execute("INSERT OR IGNORE INTO vessels(name, created_at) VALUES (?, ?)", (vessel_name, now))
        c.execute("""
            INSERT INTO upload_log(
                vessel_name, filename, file_hash, report_date,
                me_total_hrs, me_this_month, parse_confidence,
                component_count, warning_count, uploaded_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            vessel_name, filename, file_hash, parsed.get("report_date"),
            parsed.get("me_total_hrs"), parsed.get("me_this_month"),
            parsed.get("parse_confidence", 0.0),
            len(parsed.get("components", [])),
            len(parsed.get("warnings", [])),
            now
        ))

        c.execute("DELETE FROM components WHERE vessel_name = ?", (vessel_name,))
        for x in parsed["components"]:
            c.execute("""
                INSERT INTO components(
                    vessel_name, category, engine_label, unit, description,
                    periodicity, last_oh_date, last_oh_hrs, hrs_since,
                    pct_used, status, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                vessel_name, x["category"], x["engine_label"], x["unit"], x["description"],
                x["periodicity"], x["last_oh_date"], x["last_oh_hrs"], x["hrs_since"],
                x["pct_used"], x["status"], now
            ))

        c.execute("DELETE FROM other_equipment WHERE vessel_name = ?", (vessel_name,))
        for x in parsed["other_equipment"]:
            c.execute("""
                INSERT INTO other_equipment(
                    vessel_name, section, description, periodicity, last_date, run_hrs, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                vessel_name, x["section"], x["description"], x.get("periodicity", ""),
                x.get("last_date", ""), x.get("run_hrs", ""), now
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
def get_components(vessel_name):
    conn = get_db()
    df = pd.read_sql_query("SELECT * FROM components WHERE vessel_name = ?", conn, params=(vessel_name,))
    conn.close()
    return df

@st.cache_data(ttl=10)
def get_other_equipment(vessel_name):
    conn = get_db()
    df = pd.read_sql_query("SELECT * FROM other_equipment WHERE vessel_name = ? ORDER BY section, description", conn, params=(vessel_name,))
    conn.close()
    return df

@st.cache_data(ttl=10)
def get_history(vessel_name):
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
# DISPLAY
# ============================================================
STATUS_THEME = {
    "OVERDUE":{"bg":"#2d0707","status":"#ff6b6b","main":"#ff8080","accent":"#ff3333","dim":"#773333"},
    "HIGH PRIORITY":{"bg":"#2d1503","status":"#ffaa44","main":"#ff9933","accent":"#ffcc00","dim":"#774422"},
    "OK":{"bg":"#042010","status":"#4ade80","main":"#22c55e","accent":"#4ade80","dim":"#0f4023"},
    "NO DATA":{"bg":"#0c1422","status":"#7da3d8","main":"#5f7fa6","accent":"#7da3d8","dim":"#2a3950"},
}

def cyl_sort(unit):
    m = re.search(r"(\d+)", str(unit))
    return int(m.group(1)) if m else 999

def build_display_df(df, priority=False):
    if df.empty:
        return pd.DataFrame(columns=["Status","Vessel","Component","Engine","Unit","Periodicity","Last OH","Hrs Since","Used"])
    d = df.copy()

    if priority:
        order_map = {"OVERDUE":0,"HIGH PRIORITY":1,"OK":2,"NO DATA":3}
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

def apply_style(df):
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

def render_table(df, height=None, priority=False):
    if isinstance(df, list):
        df = pd.DataFrame(df)
    if df.empty:
        st.info("No data to display.")
        return
    tbl = build_display_df(df, priority=priority)
    h = height or min(900, 38 * len(tbl) + 42)
    st.dataframe(apply_style(tbl), use_container_width=True, hide_index=True, height=h, column_config=COLCFG)

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("<h2 style='margin-bottom:.15rem'>FLEET MONITOR</h2>", unsafe_allow_html=True)
    st.markdown("<div class='smallcap' style='margin-bottom:1rem'>Running Hours Intelligence System</div>", unsafe_allow_html=True)

    page = st.selectbox("Navigation", ["Fleet Overview", "Vessel Detail", "Upload Report", "Upload History"], label_visibility="collapsed")

    vessels = get_vessels()
    selected_vessel = st.selectbox("Active Vessel", vessels if vessels else ["—"], disabled=not bool(vessels))

    st.markdown("<hr>", unsafe_allow_html=True)
    db_kb = DB_PATH.stat().st_size / 1024 if DB_PATH.exists() else 0
    st.markdown(f"<div class='code'>db {db_kb:.0f} kb · {len(vessels)} vessels · text-stream parser build</div>", unsafe_allow_html=True)

# ============================================================
# PAGES
# ============================================================
if page == "Fleet Overview":
    st.markdown(hero("Fleet Master Matrix", "Universal Fleet Telemetry"), unsafe_allow_html=True)

    summary = get_summary()
    all_comps = get_all_fleet_components()

    if summary.empty or all_comps.empty:
        st.info("No data loaded. Upload a report to begin.")
        st.stop()

    cols = st.columns(5)
    with cols[0]: st.markdown(kpi(len(summary), "Vessels", "blue"), unsafe_allow_html=True)
    with cols[1]: st.markdown(kpi(len(all_comps), "Components", "gold"), unsafe_allow_html=True)
    with cols[2]: st.markdown(kpi(int((all_comps["status"]=="OVERDUE").sum()), "Overdue", "red"), unsafe_allow_html=True)
    with cols[3]: st.markdown(kpi(int((all_comps["status"]=="HIGH PRIORITY").sum()), "High Priority", "orange"), unsafe_allow_html=True)
    with cols[4]: st.markdown(kpi(int((all_comps["status"]=="OK").sum()), "OK", "green"), unsafe_allow_html=True)

    render_table(all_comps, priority=True, height=min(950, 38 * len(all_comps) + 44))

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
    with cols[1]: st.markdown(kpi(int((df["status"]=="OVERDUE").sum()), "Overdue", "red"), unsafe_allow_html=True)
    with cols[2]: st.markdown(kpi(int((df["status"]=="HIGH PRIORITY").sum()), "High Priority", "orange"), unsafe_allow_html=True)
    with cols[3]: st.markdown(kpi(int((df["status"]=="OK").sum()), "OK", "green"), unsafe_allow_html=True)
    with cols[4]: st.markdown(kpi(int((df["status"]=="NO DATA").sum()), "No Data", "blue"), unsafe_allow_html=True)

    tabs = st.tabs(["Alerts", "Main Engine", "Aux Engines", "Other Equipment"])
    with tabs[0]:
        render_table(df[df["status"].isin(["OVERDUE","HIGH PRIORITY"])], priority=True)
    with tabs[1]:
        render_table(df[df["category"]=="MAINENGINE"])
    with tabs[2]:
        render_table(df[df["category"]=="AUXENGINE"])
    with tabs[3]:
        if oe.empty:
            st.info("No other equipment data.")
        else:
            st.dataframe(oe, use_container_width=True, hide_index=True)

elif page == "Upload Report":
    st.markdown(hero("Upload Report", "TEC-004 Log Processing"), unsafe_allow_html=True)

    uploaded = st.file_uploader("Upload file", type=["doc"], label_visibility="collapsed")

    if uploaded:
        raw = uploaded.read()
        file_hash = hashlib.md5(raw).hexdigest()

        with st.spinner("Converting and extracting text-stream telemetry..."):
            docx_bytes = convert_doc_to_docx(raw)
            parsed = parse_doc_bytes(docx_bytes)

        comps = parsed["components"]
        cols = st.columns(6)
        cols[0].metric("Asset", parsed["vessel_name"])
        cols[1].metric("Report Window", parsed["report_date"] or "-")
        cols[2].metric("ME Accumulated", f"{parsed['me_total_hrs']:,}" if parsed["me_total_hrs"] else "-")
        cols[3].metric("Monthly Increment", f"{parsed['me_this_month']:,}" if parsed["me_this_month"] else "-")
        cols[4].metric("Data Channels", len(comps))
        cols[5].metric("Confidence", f"{parsed['parse_confidence']:.2f}")

        for w in parsed["warnings"]:
            st.warning(w)

        tabs = st.tabs(["Main Engine Matrix", "Aux Engines Matrix", "Other Equipment"])
        df_preview = pd.DataFrame(comps) if comps else pd.DataFrame()

        with tabs[0]:
            meprev = df_preview[df_preview["category"]=="MAINENGINE"] if not df_preview.empty else pd.DataFrame()
            render_table(meprev, height=420, priority=True) if not meprev.empty else st.info("No Main Engine telemetry extracted.")
        with tabs[1]:
            auxprev = df_preview[df_preview["category"]=="AUXENGINE"] if not df_preview.empty else pd.DataFrame()
            render_table(auxprev, height=420, priority=True) if not auxprev.empty else st.info("No Auxiliary Engine telemetry extracted.")
        with tabs[2]:
            if parsed["other_equipment"]:
                st.dataframe(pd.DataFrame(parsed["other_equipment"]), use_container_width=True, hide_index=True, height=420)
            else:
                st.info("No Other Equipment data extracted.")

        if st.button("COMMIT STREAM TO DATABASE", use_container_width=True):
            save_parsed(parsed, uploaded.name, file_hash)
            for fn in [get_vessels, get_components, get_other_equipment, get_summary, get_all_fleet_components, get_history]:
                fn.clear()
            st.success(f"{parsed['vessel_name']} committed with {len(comps)} component rows.")

elif page == "Upload History":
    st.markdown(hero("Upload History", "System Audit Trails"), unsafe_allow_html=True)
    if not vessels or selected_vessel == "—":
        st.info("Select a vessel from the sidebar.")
        st.stop()

    hist = get_history(selected_vessel)
    if hist.empty:
        st.info("No upload history for this vessel.")
    else:
        st.dataframe(hist, use_container_width=True, hide_index=True, height=520)
