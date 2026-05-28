"""
Fleet Command Telemetry - V20 (The Line-Break Resilient Build)
Architecture: Space-Merged Cell Sanitization + Extended Lookahead Boundaries
"""
import streamlit as st
st.set_page_config(page_title="Fleet Running Hours", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

import os
import re
import shutil
import tempfile
import subprocess
from pathlib import Path
import pandas as pd

# ══════════════════════════════════════════════════════════════════
#  NATIVE UI THEME
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;600;700&family=Inter:wght@400;500;600&display=swap');
:root { --bg: #071019; --bg2: #0c1623; --line: #1b2d44; --gold: #c99818; --t0: #ebf3ff; --t1: #a9bdd4; }
html, body, [class*="css"] { background: var(--bg)!important; color: var(--t1)!important; font-family: 'Inter', sans-serif!important; }
.main, .block-container { background: var(--bg)!important; }
[data-testid="stSidebar"], [data-testid="collapsedControl"] { display: none!important; }
.hero-k { font-size: .66rem; letter-spacing: .24em; text-transform: uppercase; color: var(--gold); font-weight: 700; }
.hero-h { font-family: 'Space Grotesk', sans-serif; font-size: 1.8rem; font-weight: 700; color: var(--t0); line-height: 1.1; margin-top: .2rem; }
.hero-rule { height: 1px; margin: 1rem 0; background: linear-gradient(90deg, var(--gold), var(--line), transparent); }
.metric-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; margin: 1.5rem 0; }
.metric { background: var(--bg2); border: 1px solid var(--line); border-radius: 10px; padding: 1rem; border-top: 2px solid var(--gold); }
.metric-v { font-family: 'Space Grotesk', sans-serif; font-size: 1.5rem; font-weight: 700; color: var(--t0); }
.metric-l { font-size: .6rem; text-transform: uppercase; letter-spacing: .15em; color: #71879f; margin-top: 5px; }
[data-testid="stFileUploadDropzone"] { background: rgba(201,152,24,.05)!important; border: 1.5px dashed rgba(201,152,24,.4)!important; border-radius: 12px!important; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
#  THE MARINE ALLOWLIST
# ══════════════════════════════════════════════════════════════════
ME_COMPONENTS = {
    "CYLINDER COVER", "PISTON ASSEMBLY", "STUFFING BOX", "PISTON CROWN", 
    "CYLINDER LINER", "EXHAUST VALVE", "STARTING VALVE", "SAFETY VALVE", 
    "FUEL VALVES", "FUEL PUMP", "PLUNGER AND BARREL(RENEWAL)", 
    "PLUNGER AND BARREL", "FUEL PUMP SUCTION VALVE", 
    "FUEL PUMP PUNCTURE VALVE", "CROSSHEAD BEARINGS", 
    "BOTTOM END BEARINGS", "MAIN BEARINGS"
}

AUX_COMPONENTS = {
    "CYLINDER HEAD", "PISTON", "CONNECTING ROD", "CYLINDER LINERS", 
    "FUEL VALVES (1)", "FUEL PUMPS", "CRANK PIN BEARING", "MAIN BEARING", 
    "ADJUST VALVE HEAD CLEARANCE"
}

OE_COMPONENTS = {
    "TURBOCHARGER (2)", "TURBOCHARGER (3)", "AIR COOLER", "L.O. COOLER CLEAN", 
    "F.W. COOLER CLEAN", "COOL WATER THERMOSTAT VALVE", "L.O. RENEWAL", 
    "L.O. THERMOSTAT VALVE", "ALTERNATOR CLEANING", "THRUST BEARING", 
    "TURBOCHARGER", "COOLERS", "A/C & REFR. COMPRESSORS", "GENERAL O/H", 
    "M/E L.O.", "AIR COND. COMPRESSOR NO.1", "BALANCING OF ROTOR SHAFT", 
    "JACKET FW NO.1", "AIR COND. COMPRESSOR NO.2", "AIR COOLER CLEANING", 
    "PISTON L.O.", "AIR. COND. COOLER CLEANING", "ATMOSPHERIC CONDENSER", 
    "REFRIGERATION COMPRESSOR NO.1", "REFRIGERATION COMPRESSOR NO.2", 
    "AUXILIARY BOILER", "EXH GAS BOILER", "MAIN AIR COMPRESSORS", 
    "FURNACE INSPECTION", "WASHING THE TUBES", "STARTING MAIN AIR COMPRESSOR NO.1", 
    "BURNER ATOMIZER", "O/H CIRC. PUMP NO.1", "STARTING MAIN AIR COMPRESSOR NO.2", 
    "FORCED DRAFT FAN", "O/H CIRC. PUMP NO.2", "SERVICE AIR COMPRESSOR", 
    "FEED PUMPS NO.1", "EMERGENCY AIR COMPRESSOR NO.", "FEED PUMPS NO.2",
    "COOLING WATER PUMP" # Added dynamically
}

# ══════════════════════════════════════════════════════════════════
#  TEXT SANITIZATION HELPERS
# ══════════════════════════════════════════════════════════════════
def fl(txt) -> str:
    """CRITICAL FIX: Joins multi-line cell text (like CYLINDER\nHEAD) with a space."""
    if txt is None: return ""
    raw = str(txt).replace('\x07', '').replace('\xa0', ' ').replace('\t', ' ')
    lines = [line.strip() for line in raw.split('\n') if line.strip()]
    return " ".join(lines) if lines else ""

def normalize_comp(txt) -> str:
    return re.sub(r'\s+', ' ', fl(txt).upper()).strip(" :-#")

def parse_num(txt) -> float:
    s = fl(txt).upper().replace('[', '').replace(']', '')
    if not s or any(w in s for w in ('N/A', 'MONTH', 'YEAR', 'DAY', 'OBS', 'CENTRAL', 'COOLER')): return 0.0
    m = re.search(r'\d[\d,\.]*', s)
    if not m: return 0.0
    val = re.sub(r'[,\.]', '', m.group())
    try: return float(val)
    except: return 0.0

def parse_date(txt) -> str:
    s = fl(txt).strip().replace('[', '').replace(']', '')
    if not s or s in ('-', '1', '2', 'N/A') or re.fullmatch(r'\d+', s): return ''
    return s

def get_status(hrs: float, period: float) -> str:
    if hrs <= 0 or period <= 0: return '🔵 NO DATA'
    if hrs / period >= 1.0: return '🔴 OVERDUE'
    if hrs / period >= 0.8: return '🟠 HIGH PRIORITY'
    return '🟢 OK'

# ══════════════════════════════════════════════════════════════════
#  CONVERTER
# ══════════════════════════════════════════════════════════════════
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice): raise RuntimeError("LibreOffice not found.")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as t:
        t.write(raw); src = t.name
    outdir = tempfile.mkdtemp(prefix="lo_")
    out = os.path.join(outdir, Path(src).stem + ".docx")
    profile = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        subprocess.run([soffice, "--headless", "--norestore", "--nofirststartwizard", f"-env:UserInstallation={profile}", "--convert-to", "docx", src, "--outdir", outdir], capture_output=True, timeout=120)
        with open(out, "rb") as f: return f.read()
    finally:
        for p in [src, out]:
            try: os.unlink(p)
            except: pass
        shutil.rmtree(outdir, ignore_errors=True)

# ══════════════════════════════════════════════════════════════════
#  THE "ANCHOR & BOUNDED-REACH" PARSER
# ══════════════════════════════════════════════════════════════════
def extract_telemetry(docx_bytes: bytes):
    from docx import Document
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as t:
        t.write(docx_bytes); tp = t.name
    doc = Document(tp)
    os.unlink(tp)

    cells = []
    for p in doc.paragraphs:
        if fl(p.text): cells.append(fl(p.text))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cells.append(fl(cell.text))

    me_rows, aux_rows, oe_rows = [], [], []
    vessel, report_date, me_total, me_month = "UNKNOWN", "—", 0.0, 0.0
    
    doc_text = " ".join([c for c in cells if c]).upper()
    if m := re.search(r'TOTAL RUNNING HOURS[\s:ǀ|]*([\d,]+)', doc_text): me_total = parse_num(m.group(1))
    if m := re.search(r'THIS MONTH[\s:]*([\d,]+)', doc_text): me_month = parse_num(m.group(1))
    for i, c in enumerate(cells):
        if 'VESSEL' in c.upper() and 'DATE' in c.upper():
            if m := re.search(r"Name\s*:\s*(?:MV\s+)?(.+?)\s+Date\s*:\s*(.+)$", c, re.I):
                vessel, report_date = re.sub(r'(?i)^MV\s+', '', fl(m.group(1))), parse_date(m.group(2))

    section = "GLOBAL"
    processed_me, processed_aux = set(), set()

    i = 0
    while i < len(cells):
        val = cells[i].upper()
        norm_val = normalize_comp(cells[i])
        
        # Determine Active Section
        if 'MAIN ENGINE' in val: section = "ME"
        elif 'AUX. ENGINE NO' in val or 'D/G NO' in val: section = "AUX"
        elif 'TURBOCHARGER' in val or 'A/C & REFR' in val: section = "OE"

        # ---- MAIN ENGINE ----
        if section == "ME" and norm_val in ME_COMPONENTS and norm_val not in processed_me:
            processed_me.add(norm_val)
            idx1, idx2, period = -1, -1, 0.0
            
            for offset in range(1, 20):
                if i + offset >= len(cells): break
                if cells[i + offset] == '1':
                    idx1 = i + offset
                    period = parse_num(cells[idx1 - 1])
                    break
                    
            if idx1 != -1:
                for offset in range(1, 30):
                    if idx1 + offset >= len(cells): break
                    if cells[idx1 + offset] == '2':
                        idx2 = idx1 + offset
                        break

            if idx1 != -1 and idx2 != -1:
                dates = [parse_date(cells[idx1 + k]) for k in range(1, 8) if idx1 + k < len(cells)]
                hours = [parse_num(cells[idx2 + k]) for k in range(1, 8) if idx2 + k < len(cells)]
                
                for c_idx in range(7):
                    d = dates[c_idx] if c_idx < len(dates) else ''
                    h = hours[c_idx] if c_idx < len(hours) else 0.0
                    if d or h > 0:
                        me_rows.append({
                            'Status': get_status(h, period), 'Component': norm_val, 'Engine': 'ME', 
                            'Unit': f'Cyl {c_idx+1}', 'Periodicity': int(period), 'Last O/H': d if d else '—', 
                            'Hrs Since': int(h), 'Used %': f"{round((h/period)*100, 1)}%" if period else "0.0%"
                        })
                i = idx2 + 7
                continue

        # ---- AUXILIARY ENGINE ----
        elif section == "AUX" and norm_val in AUX_COMPONENTS and norm_val not in processed_aux:
            processed_aux.add(norm_val)
            idx1, idx2, period = -1, -1, 0.0
            
            # Expanded Lookahead to account for 18 columns
            for offset in range(1, 25):
                if i + offset >= len(cells): break
                if cells[i + offset] == '1':
                    idx1 = i + offset
                    period = parse_num(cells[idx1 - 1])
                    break
                    
            if idx1 != -1:
                # Dramatically expanded lookahead to guarantee marker '2' is found past the 18 date columns
                for offset in range(1, 60):
                    if idx1 + offset >= len(cells): break
                    if cells[idx1 + offset] == '2':
                        idx2 = idx1 + offset
                        break

            if idx1 != -1 and idx2 != -1:
                dates = [parse_date(cells[idx1 + k]) for k in range(1, 19) if idx1 + k < len(cells)]
                hours = [parse_num(cells[idx2 + k]) for k in range(1, 19) if idx2 + k < len(cells)]
                
                for c_idx in range(18):
                    d = dates[c_idx] if c_idx < len(dates) else ''
                    h = hours[c_idx] if c_idx < len(hours) else 0.0
                    if d or h > 0:
                        eng_num = (c_idx // 6) + 1
                        cyl_num = (c_idx % 6) + 1
                        aux_rows.append({
                            'Status': get_status(h, period), 'Component': norm_val, 'Engine': f'AUX-{eng_num}', 
                            'Unit': f'Cyl {cyl_num}', 'Periodicity': int(period), 'Last O/H': d if d else '—', 
                            'Hrs Since': int(h), 'Used %': f"{round((h/period)*100, 1)}%" if period else "0.0%"
                        })
                i = idx2 + 18
                continue

        # ---- OTHER EQUIPMENT ----
        elif norm_val in OE_COMPONENTS:
            d, h = '', 0.0
            for offset in range(1, 10):
                if i + offset >= len(cells): break
                tok = cells[i + offset]
                if not d and ('/' in tok or '-' in tok or re.search(r'[A-Za-z]{3}\s+\d{2,4}', tok)): d = parse_date(tok)
                if h == 0.0 and re.match(r'^\d[\d,\.]*$', tok) and tok not in ('1', '2'): h = parse_num(tok)
            
            if d or h > 0:
                oe_rows.append({
                    'Section': 'Other Equipment', 'Description': norm_val, 
                    'Last Date': d if d else '—', 'Run Hrs': int(h) if h > 0 else 0
                })
                i += 1 
                continue
                
        i += 1

    oe_clean = [dict(t) for t in {tuple(d.items()) for d in oe_rows}]
    return vessel, report_date, me_total, me_month, me_rows, aux_rows, oe_clean

# ══════════════════════════════════════════════════════════════════
#  CRASH-PROOF UI BUILDER
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero-k">Running Hours Management System</div>
<div class="hero-h">TEC-004 Extraction Matrix</div>
<div class="hero-rule"></div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("Upload TEC-004 Report (.doc)", type=['doc'])

if uploaded:
    with st.spinner("Executing Zero-Trust Extraction..."):
        try:
            docx_data = convert_doc_to_docx(uploaded.read())
            vessel, report_date, me_total, me_month, me_data, aux_data, oe_data = extract_telemetry(docx_data)
            
            n_od = sum(1 for c in me_data + aux_data if 'OVERDUE' in c['Status'])
            n_hp = sum(1 for c in me_data + aux_data if 'HIGH PRIORITY' in c['Status'])

            st.markdown(f"""
            <div class="metric-grid">
              <div class="metric"><div class="metric-v">{vessel}</div><div class="metric-l">Vessel</div></div>
              <div class="metric"><div class="metric-v">{report_date}</div><div class="metric-l">Report Date</div></div>
              <div class="metric"><div class="metric-v">{me_total:,}</div><div class="metric-l">ME Total Hrs</div></div>
              <div class="metric"><div class="metric-v">{me_month:,}</div><div class="metric-l">ME This Month</div></div>
            </div>
            """, unsafe_allow_html=True)

            tab1, tab2, tab3 = st.tabs([f"⚙ Main Engine ({len(me_data)})", f"🔩 Aux Engines ({len(aux_data)})", f"🛠 Other Equipment ({len(oe_data)})"])

            with tab1:
                if not me_data: st.info("No Main Engine records found.")
                else:
                    df_me = pd.DataFrame(me_data)
                    df_me['_cyl'] = df_me['Unit'].str.extract(r'(\d+)').astype(float)
                    df_me = df_me.sort_values(by=['Component', '_cyl']).drop(columns=['_cyl']).astype(str)
                    st.dataframe(df_me, use_container_width=True, hide_index=True)

            with tab2:
                if not aux_data: st.info("No Auxiliary Engine records found.")
                else:
                    df_aux = pd.DataFrame(aux_data)
                    df_aux['_eng'] = df_aux['Engine'].str.extract(r'(\d+)').astype(float)
                    df_aux['_cyl'] = df_aux['Unit'].str.extract(r'(\d+)').astype(float)
                    df_aux = df_aux.sort_values(by=['Component', '_eng', '_cyl']).drop(columns=['_eng', '_cyl']).astype(str)
                    st.dataframe(df_aux, use_container_width=True, hide_index=True)

            with tab3:
                if not oe_data: st.info("No Other Equipment records found.")
                else:
                    df_oe = pd.DataFrame(oe_data).astype(str)
                    st.dataframe(df_oe, use_container_width=True, hide_index=True)

        except Exception as e:
            st.error(f"Execution Failed: {e}")
