"""
╔══════════════════════════════════════════════════════════════════════╗
║   FLEET COMMAND & TELEMETRY SYSTEM  v9.0                             ║
║   Geometric ETL Parser · Idempotent DB · Enterprise Edition          ║
╚══════════════════════════════════════════════════════════════════════╝
"""

import streamlit as st

st.set_page_config(
    page_title="Fleet Command | v9.0",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="expanded",
)

import os, re, sqlite3, tempfile, hashlib, subprocess, shutil
from datetime import datetime
from pathlib import Path
import pandas as pd
from docx import Document

# ═══════════════════════════════════════════════════════════════════
#  GLOBAL UI STEALTH DESIGN SYSTEM
# ═══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');
:root {
  --bg: #03060d; --bg1: #06091a; --bg2: #080e20; --bg3: #0b1228; --bg4: #0f1830;
  --b1: #0f1c35; --b2: #182840; --b3: #223350;
  --gold: #c89a14; --gold2:#e0b422; --gold3:#f5cc44;
  --red: #cc2828; --red2: #ff5c5c; --orange: #b85518; --ora2: #ff8833;
  --green: #0d8a4a; --grn2: #22c55e; --blue: #1444a8; --blu2: #3b82f6;
  --t0: #f2f7ff; --t1: #c0d0e8; --t2: #6a84a8; --t3: #304060;
  --ff: 'Space Grotesk', sans-serif; --fi: 'Inter', sans-serif; --fm: 'JetBrains Mono', monospace;
}
*, *::before, *::after { box-sizing: border-box; }
html, body, [class*="css"] { font-family: var(--fi)!important; background: var(--bg)!important; color: var(--t1)!important; -webkit-font-smoothing: antialiased; }
.main, .main > div { background: var(--bg)!important; }
.block-container { padding: 2rem 2.5rem 5rem!important; max-width: 100%!important; }
.main::before { content: ''; position: fixed; inset: 0; pointer-events: none; z-index: 0; background: radial-gradient(ellipse 90% 50% at -10% -5%, rgba(200,154,20,0.06) 0%, transparent 55%), radial-gradient(ellipse 70% 45% at 110% 105%, rgba(20,68,168,0.05) 0%, transparent 55%); }
[data-testid="stSidebar"] { background: var(--bg1)!important; border-right: 1px solid var(--b2)!important; }
[data-testid="stSidebar"] * { color: var(--t1)!important; }
[data-testid="stSidebarContent"] { padding: 1.5rem 1.25rem!important; }
h1 { font-family: var(--ff)!important; font-size: 1.8rem!important; font-weight: 700!important; color: var(--t0)!important; letter-spacing: -0.02em!important; line-height: 1.2!important; }
h2 { font-family: var(--ff)!important; font-size: 1.2rem!important; font-weight: 600!important; color: var(--t0)!important; }
h3 { font-family: var(--ff)!important; font-size: 1rem!important; font-weight: 500!important; color: var(--t1)!important; }
[data-testid="stMetric"] { background: var(--bg3)!important; border: 1px solid var(--b2)!important; border-radius: 10px!important; padding: 1rem 1.2rem 1.1rem!important; }
[data-testid="stMetricValue"] { font-family: var(--ff)!important; font-size: 2rem!important; font-weight: 700!important; color: var(--t0)!important; }
[data-testid="stMetricLabel"] { font-family: var(--fi)!important; color: var(--t3)!important; font-size: 0.62rem!important; text-transform: uppercase!important; letter-spacing: 0.15em!important; }
[data-testid="stDataFrame"] { border: 1px solid var(--b2)!important; border-radius: 10px!important; overflow: hidden!important; box-shadow: 0 4px 24px rgba(0,0,0,0.35)!important; }
.stButton > button { background: linear-gradient(135deg, var(--gold) 0%, #8a6a08 100%)!important; color: #000!important; border: none!important; font-family: var(--ff)!important; font-weight: 600!important; font-size: 0.82rem!important; letter-spacing: 0.06em!important; text-transform: uppercase!important; border-radius: 7px!important; padding: .6rem 1.8rem!important; box-shadow: 0 2px 14px rgba(200,154,20,.2),inset 0 1px 0 rgba(255,255,255,.1)!important; transition: all .18s!important; }
.stButton > button:hover { transform: translateY(-2px)!important; }
.stTabs [data-baseweb="tab-list"] { background: var(--bg2)!important; border-radius: 10px 10px 0 0!important; border-bottom: 1px solid var(--b2)!important; padding: 0 1rem!important; }
.stTabs [data-baseweb="tab"] { background: transparent!important; color: var(--t3)!important; font-family: var(--ff)!important; font-weight: 500!important; font-size: .75rem!important; text-transform: uppercase!important; padding: .85rem 1.3rem!important; border-bottom: 2px solid transparent!important; }
.stTabs [aria-selected="true"] { color: var(--gold2)!important; border-bottom: 2px solid var(--gold)!important; }
.stTabs [data-baseweb="tab-panel"] { background: var(--bg2)!important; border: 1px solid var(--b2)!important; border-top: none!important; border-radius: 0 0 10px 10px!important; padding: 1.5rem!important; }
.kc { background: var(--bg3); border: 1px solid var(--b2); border-radius: 10px; padding: 1rem 1.2rem 1.1rem; position: relative; overflow: hidden; }
.kc::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; border-radius: 10px 10px 0 0; }
.kc.gold::before { background: var(--gold); } .kc.red::before { background: var(--red); } .kc.orange::before { background: var(--orange); } .kc.green::before { background: var(--green); } .kc.blue::before { background: var(--blue); }
.kc-val { font-family: var(--ff); font-size: 2.2rem; font-weight: 700; line-height: 1.1; letter-spacing: -.04em; }
.kc.gold .kc-val { color: var(--gold3); } .kc.red .kc-val { color: var(--red2); } .kc.orange .kc-val { color: var(--ora2); } .kc.green .kc-val { color: var(--grn2); } .kc.blue .kc-val { color: var(--blu2); }
.kc-lbl { font-family: var(--fi); font-size: .6rem; font-weight: 500; text-transform: uppercase; letter-spacing: .16em; color: var(--t3); margin-top: 5px; }
.sl { font-family: var(--fi); font-size: .58rem; font-weight: 600; letter-spacing: .22em; text-transform: uppercase; color: var(--t3); display: flex; align-items: center; gap: .75rem; margin: 1.75rem 0 1rem; }
.sl::after { content: ''; flex: 1; height: 1px; background: linear-gradient(90deg, var(--b2), transparent); }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════
#  DATABASE LAYER (IDEMPOTENT UPSERT ARCHITECTURE)
# ═══════════════════════════════════════════════════════════════════
DB_PATH = Path("running_hours_v9.db")

def get_db():
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    c = get_db()
    c.execute("PRAGMA journal_mode=WAL;")
    c.executescript("""
    CREATE TABLE IF NOT EXISTS vessels(
        name TEXT PRIMARY KEY,
        created_at TEXT NOT NULL DEFAULT(datetime('now')));
        
    CREATE TABLE IF NOT EXISTS upload_log(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL, filename TEXT NOT NULL,
        file_hash TEXT NOT NULL, report_date TEXT,
        me_total_hrs INTEGER, me_this_month INTEGER,
        uploaded_at TEXT NOT NULL DEFAULT(datetime('now')));
        
    CREATE TABLE IF NOT EXISTS components(
        vessel_name TEXT NOT NULL, category TEXT NOT NULL,
        engine_label TEXT NOT NULL, unit TEXT NOT NULL,
        description TEXT NOT NULL, periodicity REAL,
        last_oh_date TEXT, last_oh_hrs REAL,
        hrs_since REAL, pct_used REAL, status TEXT NOT NULL,
        updated_at TEXT NOT NULL DEFAULT(datetime('now')),
        PRIMARY KEY (vessel_name, category, engine_label, unit, description)
    );
    
    CREATE TABLE IF NOT EXISTS quarantine(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL, filename TEXT NOT NULL,
        raw_text TEXT, category TEXT, presumed_unit TEXT,
        extracted_hrs REAL, extracted_date TEXT,
        uploaded_at TEXT NOT NULL DEFAULT(datetime('now'))
    );
    """)
    c.commit(); c.close()

init_db()


# ═══════════════════════════════════════════════════════════════════
#  CONVERSION  — Headless File Decryption Engine
# ═══════════════════════════════════════════════════════════════════
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice): raise RuntimeError("LibreOffice not found in environment.")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as t:
        t.write(raw); tp = t.name
    od = tempfile.gettempdir()
    base = os.path.splitext(os.path.basename(tp))[0]
    dp = os.path.join(od, base + ".docx")
    pf = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        r = subprocess.run([soffice,"--headless","--norestore","--nofirststartwizard", f"-env:UserInstallation={pf}","--convert-to","docx",tp,"--outdir",od], capture_output=True, timeout=120)
        if not os.path.exists(dp): raise RuntimeError(f"Conversion failed: {r.stderr.decode('utf-8','ignore')[:300]}")
        with open(dp,"rb") as f: return f.read()
    finally:
        for p in [tp,dp]:
            try:
                if os.path.exists(p): os.unlink(p)
            except Exception: pass


# ═══════════════════════════════════════════════════════════════════
#  THE PYTHON "CAR WASH" (STRICT TEXT SANITIZATION)
# ═══════════════════════════════════════════════════════════════════
def get_first_line(raw: str) -> str:
    """Mimics VBA GetFirstLine: Splits on carriage returns, isolates primary text."""
    if not raw: return ""
    c = str(raw).replace('\x0b', '\n').replace('\r', '\n').replace('\x07', '')
    lines = [line.strip() for line in c.split('\n') if line.strip()]
    return lines[0] if lines else ""

def clean_cell_text(raw: str) -> str:
    """Purges hex chars and normalizes spacing."""
    if not raw: return ""
    txt = get_first_line(raw)
    txt = txt.replace('\xa0', ' ').replace('\t', ' ')
    while "  " in txt: txt = txt.replace("  ", " ")
    return txt.strip()

def is_valid_component(name: str) -> bool:
    """The Engine Legend Firewall."""
    u = name.upper()
    if u in ["", "DESCRIPTION", "REMARKS", "COMPONENT", "-"]: return False
    if "DATE OF LAST" in u or "RUNNING HOURS" in u or "PERIODICITY" in u or "TYPE:" in u: return False
    if re.fullmatch(r"[\d./ ,:-]+", u): return False # Kills pure numbers
    if len(u) > 55: return False # Kills remarks paragraphs
    return True

def clean_component_name(raw: str) -> str:
    """The ALEXIS Filter."""
    t = clean_cell_text(raw)
    t = re.sub(r'(?i)ALEXIS\s*Date', '', t)
    t = re.sub(r'(?i)Page\s*\d+\s*of\s*\d+', '', t)
    return t.strip()

def extract_number(raw: str) -> float:
    """Bulletproof geometric number extraction."""
    c = clean_cell_text(raw).upper()
    if c in ["N/A", "CENTRAL", "-", ""]: return 0.0
    if any(x in c for x in ["MONTH", "YEAR", "WEEK", "DAY", "OBSERVATION"]): return 0.0
    
    # Gap closer (fixes typos like "21, 476")
    c = c.replace(", ", ",").replace(". ", ".")
    
    num_str = ""
    locked = False
    for char in c:
        if char.isdigit():
            locked = True
            num_str += char
        elif char in [".", ","] and locked:
            num_str += char
        elif locked:
            break
            
    if not num_str: return 0.0
    num_str = num_str.replace(",", "")
    try: return float(num_str)
    except: return 0.0

def clean_date(raw: str) -> str:
    """Strict Date boundaries."""
    c = clean_cell_text(raw)
    if c in ["", "-", "1", "2", "1/", "/2", "N/A", "n/a"]: return ""
    if len(c) > 20: return ""
    return c

def get_status(hrs: float, period: float) -> str:
    if hrs <= 0: return "NO DATA"
    if period <= 0: return "NO DATA"
    ratio = hrs / period
    if ratio >= 1.0: return "OVERDUE"
    if ratio >= 0.8: return "HIGH PRIORITY"
    return "OK"


# ═══════════════════════════════════════════════════════════════════
#  GEOMETRIC PARSER PIPELINES (VBA TRANSLATION)
# ═══════════════════════════════════════════════════════════════════
def process_me_table(grid: list) -> tuple[list, list]:
    comps, quarantine = [], []
    if len(grid) < 5: return comps, quarantine
    
    is_me = False
    for r in range(min(4, len(grid))):
        for c in range(len(grid[r])):
            if "MAIN ENGINE" in clean_cell_text(grid[r][c]).upper(): is_me = True
    if not is_me: return comps, quarantine

    # 1. Map Coordinates
    per_col, rem_col = 1, len(grid[0])-1
    for c in range(len(grid[0])):
        h_txt = "".join([clean_cell_text(grid[r][c]).upper() for r in range(min(3, len(grid)))])
        if "PERIOD" in h_txt: per_col = c
        if "REMARK" in h_txt: rem_col = c

    actual_cyls = rem_col - per_col - 1
    if actual_cyls < 5: actual_cyls = 6
    if actual_cyls > 7: actual_cyls = 7 # Strict physical bounds
    
    first_cyl_col = per_col + 2
    marker_col = per_col + 1

    me_end = len(grid)
    for r in range(len(grid)):
        f_row = "".join([clean_cell_text(cell).upper() for cell in grid[r]])
        if "NOTE 1" in f_row or "TURBOCHARGER" in f_row or "AUX. ENGINE" in f_row:
            me_end = r
            break

    # 2. Extract Data using Anchor Lock
    for r in range(1, me_end - 1): # -1 because we read r+1
        raw_name = grid[r][0]
        c_name = clean_component_name(raw_name)
        period = extract_number(grid[r][per_col]) if per_col < len(grid[r]) else 0.0
        marker = clean_cell_text(grid[r][marker_col]) if marker_col < len(grid[r]) else ""

        if "1" in marker.split(): # Only trigger on explicit Marker row
            for cyl in range(1, actual_cyls + 1):
                c = first_cyl_col + cyl - 1
                if c < len(grid[r]) and c < len(grid[r+1]):
                    date_val = clean_date(grid[r][c])
                    hrs_val = extract_number(grid[r+1][c])
                    
                    if date_val or hrs_val > 0:
                        data = {
                            "category": "MAIN_ENGINE", "engine_label": "ME", "unit": f"Cyl {cyl}",
                            "description": c_name, "periodicity": period, "last_oh_date": date_val,
                            "last_oh_hrs": hrs_val, "hrs_since": hrs_val,
                            "pct_used": (hrs_val/period if period>0 else 0), "status": get_status(hrs_val, period)
                        }
                        if is_valid_component(c_name): comps.append(data)
                        else: quarantine.append({"raw": raw_name, "hrs": hrs_val, "unit": f"ME Cyl {cyl}"})
    return comps, quarantine

def process_aux_table(grid: list) -> tuple[list, list]:
    comps, quarantine = [], []
    aux_start, aux_end = 0, len(grid)
    
    for r in range(len(grid)):
        f_row = "".join([clean_cell_text(cell).upper() for cell in grid[r]])
        if "AUX. ENGINE MAKER" in f_row or "AUX. ENGINE NO" in f_row:
            aux_start = r; break
            
    if aux_start == 0: return comps, quarantine
    
    desc_row = 0
    for r in range(aux_start, len(grid)):
        check = clean_cell_text(grid[r][0]).upper()
        if "DESCRIPTION" in check and desc_row == 0: desc_row = r
        if "TURBOCHARGER" in check or "D/G" in check:
            aux_end = r; break

    actual_cyls = 0
    if desc_row > 0:
        for c in range(3, len(grid[desc_row])):
            val = clean_cell_text(grid[desc_row][c])
            if val.isdigit() and int(val) == actual_cyls + 1: actual_cyls += 1
            elif actual_cyls > 0: break
            
    if actual_cyls == 0 or actual_cyls > 6: actual_cyls = 6
    
    marker_col, aux1_start = 2, 3
    aux2_start = aux1_start + actual_cyls
    aux3_start = aux2_start + actual_cyls
    
    start_ex = desc_row + 1 if desc_row > 0 else aux_start
    
    for r in range(start_ex, aux_end - 1):
        raw_name = grid[r][0]
        c_name = clean_component_name(raw_name)
        period = extract_number(grid[r][1]) if 1 < len(grid[r]) else 0.0
        marker = clean_cell_text(grid[r][marker_col]) if marker_col < len(grid[r]) else ""
        
        if "1" in marker.split():
            for engine, offset in [("AUX-1", aux1_start), ("AUX-2", aux2_start), ("AUX-3", aux3_start)]:
                for cyl in range(1, actual_cyls + 1):
                    c = offset + cyl - 1
                    if c < len(grid[r]) and c < len(grid[r+1]):
                        date_val = clean_date(grid[r][c])
                        hrs_val = extract_number(grid[r+1][c])
                        if date_val or hrs_val > 0:
                            data = {
                                "category": "AUX_ENGINE", "engine_label": engine, "unit": f"Cyl {cyl}",
                                "description": c_name, "periodicity": period, "last_oh_date": date_val,
                                "last_oh_hrs": hrs_val, "hrs_since": hrs_val,
                                "pct_used": (hrs_val/period if period>0 else 0), "status": get_status(hrs_val, period)
                            }
                            if is_valid_component(c_name): comps.append(data)
                            else: quarantine.append({"raw": raw_name, "hrs": hrs_val, "unit": f"{engine} Cyl {cyl}"})
    return comps, quarantine

def parse_doc_bytes(docx_bytes: bytes) -> dict:
    with tempfile.NamedTemporaryFile(suffix='.docx',delete=False) as t:
        t.write(docx_bytes); tp=t.name
    try: doc=Document(tp)
    except Exception as e: raise ValueError(f"Cannot open document: {e}")
    finally:
        try: os.unlink(tp)
        except Exception: pass

    vn = 'UNKNOWN'; rd = None; mt = None; mm = None
    
    # Metadata Sweep
    for p in doc.paragraphs:
        txt = clean_cell_text(p.text)
        if m := re.search(r"Vessel[\u2019\u2018']?s?\s+Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+?)(?:\s{2,}|\t)", txt, re.I): vn = clean_cell_text(m.group(1))
        if m := re.search(r"Date\s*:\s*(.+)", txt, re.I): rd = clean_date(m.group(1))
    
    all_comps, all_quar = [], []
    for table in doc.tables:
        grid = [[cell.text for cell in row.cells] for row in table.rows]
        
        # M/E Totals
        if len(grid) > 0 and len(grid[0]) > 0:
            c0 = clean_cell_text(grid[0][0])
            if m := re.search(r'Total Running Hours[\s:ǀ|]+([\d,]+)', c0, re.I): mt = extract_number(m.group(1))
            if m := re.search(r'This Month[\s:]+([\d,]+)', c0, re.I): mm = extract_number(m.group(1))
                
        mc, mq = process_me_table(grid)
        ac, aq = process_aux_table(grid)
        all_comps.extend(mc); all_comps.extend(ac)
        all_quar.extend(mq); all_quar.extend(aq)

    return {
        "vessel_name": vn, "report_date": rd, 
        "me_total_hrs": mt, "me_this_month": mm,
        "components": all_comps, "quarantine": all_quar
    }


# ═══════════════════════════════════════════════════════════════════
#  ETL LOAD LAYER (UPSERT)
# ═══════════════════════════════════════════════════════════════════
def save_parsed(parsed: dict, filename: str, fhash: str):
    conn = get_db(); c = conn.cursor()
    now = datetime.utcnow().isoformat() + "Z"
    v = parsed['vessel_name']
    
    try:
        c.execute("INSERT OR IGNORE INTO vessels(name,created_at) VALUES(?,?)",(v,now))
        c.execute("INSERT INTO upload_log(vessel_name,filename,file_hash,report_date,me_total_hrs,me_this_month,uploaded_at) VALUES(?,?,?,?,?,?,?)",
            (v,filename,fhash,parsed['report_date'],parsed['me_total_hrs'],parsed['me_this_month'],now))
        
        # IDEMPOTENT UPSERT: Replaces data only if the exact part is found. Never deletes tables.
        for x in parsed['components']:
            c.execute("""
                INSERT OR REPLACE INTO components
                (vessel_name, category, engine_label, unit, description, periodicity, last_oh_date, last_oh_hrs, hrs_since, pct_used, status, updated_at) 
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
            """, (v, x['category'], x['engine_label'], x['unit'], x['description'], x['periodicity'], x['last_oh_date'], x['last_oh_hrs'], x['hrs_since'], x['pct_used'], x['status'], now))
            
        for q in parsed['quarantine']:
            c.execute("INSERT INTO quarantine(vessel_name, filename, raw_text, presumed_unit, extracted_hrs, uploaded_at) VALUES(?,?,?,?,?,?)",
                      (v, filename, q['raw'], q['unit'], q['hrs'], now))
        conn.commit()
    except Exception as e:
        conn.rollback() 
        raise e
    finally:
        conn.close()


# ═══════════════════════════════════════════════════════════════════
#  UI DASHBOARD QUERIES
# ═══════════════════════════════════════════════════════════════════
def get_vessels():
    c=get_db(); r=c.execute("SELECT name FROM vessels ORDER BY name").fetchall(); c.close()
    return [x['name'] for x in r]

def get_comps(vessel):
    c=get_db(); df=pd.read_sql_query("SELECT * FROM components WHERE vessel_name=?",c,params=(vessel,)); c.close(); return df

def get_quarantine(vessel):
    c=get_db(); df=pd.read_sql_query("SELECT raw_text, presumed_unit, extracted_hrs, uploaded_at FROM quarantine WHERE vessel_name=? ORDER BY id DESC LIMIT 50",c,params=(vessel,)); c.close(); return df

def get_summary():
    c=get_db()
    df=pd.read_sql_query("""
        SELECT c.vessel_name,
            COUNT(CASE WHEN c.status='OVERDUE' THEN 1 END) AS overdue,
            COUNT(CASE WHEN c.status='HIGH PRIORITY' THEN 1 END) AS high_priority,
            COUNT(CASE WHEN c.status='OK' THEN 1 END) AS ok, COUNT(*) AS total
        FROM components c GROUP BY c.vessel_name ORDER BY overdue DESC
    """,c); c.close(); return df

def get_all_comps():
    c = get_db(); df = pd.read_sql_query("SELECT * FROM components", c); c.close(); return df

_S = {
    'OVERDUE':       {'bg':'#1f0505','bgs':'#2d0707','ts':'#ff6b6b','tm':'#ff8080','tn':'#ff3333','td':'#773333'},
    'HIGH PRIORITY': {'bg':'#1e0d02','bgs':'#2d1503','ts':'#ffaa44','tm':'#ff9933','tn':'#ffcc00','td':'#774422'},
    'OK':            {'bg':'#021208','bgs':'#042010','ts':'#4ade80','tm':'#22c55e','tn':'#4ade80','td':'#0f4023'},
    '_':             {'bg':'#090e18','bgs':'#0c1422','ts':'#4a6688','tm':'#334d66','tn':'#334d66','td':'#1a2a38'},
}

def render_table(df: pd.DataFrame, priority: bool = False):
    if df.empty: st.info("No data to display."); return
    d = df.copy()
    if priority:
        _ORD = {'OVERDUE':0,'HIGH PRIORITY':1,'OK':2,'NO DATA':3}
        d['_s'] = d['status'].map(lambda s: _ORD.get(str(s),4))
        d['_p'] = d['pct_used'].fillna(0)
        d = d.sort_values(['_s','_p'], ascending=[True,False]).drop(columns=['_s','_p'])
    
    out = pd.DataFrame()
    out['Status']      = d['status'].values
    out['Vessel']      = d.get('vessel_name', pd.Series(['—']*len(d))).values
    out['Component']   = d['description'].values
    out['Engine']      = d['engine_label'].values
    out['Unit']        = d['unit'].values
    out['Periodicity'] = [int(x) if pd.notna(x) and x>0 else None for x in d['periodicity']]
    out['Last O/H']    = [str(x) if x and str(x)!='nan' else '—' for x in d['last_oh_date']]
    out['Hrs Since']   = [int(x) if pd.notna(x) and x>0 else None for x in d['hrs_since']]
    out['% Used']      = [round(float(x)*100,1) if pd.notna(x) else 0.0 for x in d['pct_used']]

    def rs(row):
        c = _S.get(str(row.get('Status','')), _S['_'])
        return [f"background-color:{c['bg']};color:{c['ts']}"] + [f"background-color:{c['bg']};color:{c['td']}"]*6 + [f"background-color:{c['bg']};color:{c['tn']}"]*2

    st.dataframe(out.style.apply(rs, axis=1), use_container_width=True, hide_index=True, height=min(800, 38*(len(out)+1)+4),
        column_config={"% Used": st.column_config.ProgressColumn("% Used", min_value=0, max_value=160, format="%.1f%%")})

def kpi(val, lbl, color="gold"): return f'<div class="kc {color}"><div class="kc-val">{val}</div><div class="kc-lbl">{lbl}</div></div>'
def sl(txt): return f'<div class="sl">{txt}</div>'


# ═══════════════════════════════════════════════════════════════════
#  APP ROUTER
# ═══════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown('<div style="font-family:var(--ff);font-size:1.15rem;font-weight:700;color:var(--gold2)">⚓ FLEET COMMAND</div><hr>', unsafe_allow_html=True)
    page = st.selectbox("Navigation", ["🗺️ Fleet Overview", "🚢 Vessel Detail", "📤 Upload Report", "☣️ Quarantine Log"], label_visibility="collapsed")
    vessels = get_vessels()
    sel_v = st.selectbox("Active Context", vessels) if vessels else None

if page == "🗺️ Fleet Overview":
    st.markdown("<h1>🗺️ Fleet Master Matrix</h1>", unsafe_allow_html=True)
    smry = get_summary(); all_comps = get_all_comps()
    if all_comps.empty: st.info("Database empty. Upload a report to initialize ETl."); st.stop()
    
    k1,k2,k3,k4,k5 = st.columns(5)
    with k1: st.markdown(kpi(len(smry), "Vessels", "blue"), unsafe_allow_html=True)
    with k2: st.markdown(kpi(len(all_comps), "Components", "gold"), unsafe_allow_html=True)
    with k3: st.markdown(kpi((all_comps['status'] == 'OVERDUE').sum(), "Overdue", "red"), unsafe_allow_html=True)
    with k4: st.markdown(kpi((all_comps['status'] == 'HIGH PRIORITY').sum(), "High Priority", "orange"), unsafe_allow_html=True)
    with k5: st.markdown(kpi((all_comps['status'] == 'OK').sum(), "OK", "green"), unsafe_allow_html=True)
    
    st.markdown(sl("Universal Component Control Grid"), unsafe_allow_html=True)
    render_table(all_comps, priority=True)

elif page == "🚢 Vessel Detail":
    if not sel_v: st.info("Select an active vessel."); st.stop()
    st.markdown(f"<h1>🚢 {sel_v} Component Analysis</h1>", unsafe_allow_html=True)
    df = get_comps(sel_v)
    if df.empty: st.info("No parsed data."); st.stop()
    
    k1,k2,k3,k4 = st.columns(4)
    with k1: st.markdown(kpi(len(df), "Total Nodes", "gold"), unsafe_allow_html=True)
    with k2: st.markdown(kpi((df['status'] == 'OVERDUE').sum(), "Overdue", "red"), unsafe_allow_html=True)
    with k3: st.markdown(kpi((df['status'] == 'HIGH PRIORITY').sum(), "High Priority", "orange"), unsafe_allow_html=True)
    with k4: st.markdown(kpi((df['status'] == 'OK').sum(), "Healthy", "green"), unsafe_allow_html=True)
    
    st.markdown(sl("Component Roster"), unsafe_allow_html=True)
    render_table(df, priority=True)

elif page == "📤 Upload Report":
    st.markdown("<h1>📤 Pipeline Ingestion</h1>", unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload TEC-004 .doc", type=["doc"])
    
    if uploaded:
        raw = uploaded.read(); fh = hashlib.md5(raw).hexdigest()
        with st.spinner("Executing Geometric Extraction..."):
            docx = convert_doc_to_docx(raw)
            parsed = parse_doc_bytes(docx)
            
        c1, c2, c3 = st.columns(3)
        c1.metric("Vessel Detected", parsed['vessel_name'])
        c2.metric("Components Successfully Mapped", len(parsed['components']))
        c3.metric("Quarantined Artifacts", len(parsed['quarantine']))
        
        st.markdown(sl("Commit Authorization"), unsafe_allow_html=True)
        if st.button("✅ EXECUTE DATABASE UPSERT", use_container_width=True):
            save_parsed(parsed, uploaded.name, fh)
            st.success(f"Upsert Complete: {len(parsed['components'])} nodes committed to DB structure.")
            st.balloons()

elif page == "☣️ Quarantine Log":
    if not sel_v: st.info("Select an active vessel."); st.stop()
    st.markdown(f"<h1>☣️ {sel_v} Quarantine Deck</h1>", unsafe_allow_html=True)
    q = get_quarantine(sel_v)
    if q.empty: st.success("No unmapped data artifacts detected for this vessel.")
    else:
        st.warning(f"{len(q)} data rows contained valid running hours but the component name was rejected by the engine firewall.")
        st.dataframe(q, use_container_width=True, hide_index=True)
