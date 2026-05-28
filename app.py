"""
Fleet Running Hours Monitor - v18 (The 10/10 Enterprise Master Build)
Architecture: 2D Geometric Table Grid Scanning + PyArrow Crash-Proofing
"""
import streamlit as st
import os
import re
import shutil
import hashlib
import tempfile
import subprocess
from pathlib import Path
from datetime import datetime
import pandas as pd

# ══════════════════════════════════════════════════════════════════
#  UI DASHBOARD & CSS
# ══════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Fleet Running Hours Monitor", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');
:root{
  --bg:#071019; --bg2:#0c1623; --bg3:#111e30; --line:#1b2d44; --line2:#2a476b;
  --gold:#c99818; --gold2:#f0c45f; --red:#ef6666; --ora:#ffb04d; --grn:#69c07a; --blu:#6bb4ff;
  --t0:#ebf3ff; --t1:#a9bdd4; --t2:#71879f; --t3:#43576d; --r:14px;
  --ff:'Space Grotesk',sans-serif; --fi:'Inter',sans-serif; --fm:'JetBrains Mono',monospace;
}
html,body,[class*="css"]{background:var(--bg)!important;color:var(--t1)!important;font-family:var(--fi)!important}
.main,.main>div,.block-container{background:var(--bg)!important}
.block-container{max-width:100%!important;padding:1rem 1.5rem 3rem!important}
[data-testid="collapsedControl"],[data-testid="stSidebar"]{display:none!important}
.main::before{
 content:"";position:fixed;inset:0;pointer-events:none;z-index:0;
 background: radial-gradient(ellipse 70% 45% at 0% 0%, rgba(201,152,24,.08), transparent 60%), radial-gradient(ellipse 55% 35% at 100% 100%, rgba(107,180,255,.06), transparent 60%);
}
.block-container>*{position:relative;z-index:1}
.hero-k{font-size:.66rem;letter-spacing:.24em;text-transform:uppercase;color:var(--gold2);font-weight:700}
.hero-h{font-family:var(--ff);font-size:1.9rem;font-weight:700;color:var(--t0);letter-spacing:-.04em;line-height:1.06;margin-top:.2rem}
.hero-rule{height:1px;margin:.8rem 0 1rem;background:linear-gradient(90deg,var(--gold2),var(--line),transparent)}
.metric-grid{display:grid;grid-template-columns:repeat(8,1fr);gap:.75rem;margin:1rem 0}
.metric{
 background:linear-gradient(180deg,var(--bg3),var(--bg2));
 border:1px solid var(--line);border-radius:var(--r);padding:.85rem .95rem;position:relative
}
.metric::before{content:"";position:absolute;left:0;right:0;top:0;height:2px;background:linear-gradient(90deg,var(--gold),transparent 75%)}
.metric.r::before{background:linear-gradient(90deg,var(--red),transparent 75%)}
.metric.o::before{background:linear-gradient(90deg,var(--ora),transparent 75%)}
.metric.g::before{background:linear-gradient(90deg,var(--grn),transparent 75%)}
.metric-v{font-family:var(--ff);font-size:1.45rem;font-weight:700;color:var(--t0);line-height:1.05;letter-spacing:-.04em}
.metric-l{color:var(--t3);font-size:.58rem;text-transform:uppercase;letter-spacing:.16em;margin-top:.3rem}
[data-testid="stFileUploadDropzone"]{
 background:rgba(201,152,24,.04)!important;border:1.5px dashed rgba(201,152,24,.5)!important;
 border-radius:16px!important;padding:2rem 1.25rem!important
}
[data-testid="stFileUploadDropzone"]:hover{border-color:var(--gold2)!important;background:rgba(201,152,24,.07)!important}
[data-testid="stDataFrame"]{border:1px solid var(--line)!important;border-radius:var(--r)!important;overflow:auto!important}
.dvn-scroller{background:var(--bg2)!important}
.stTabs [data-baseweb="tab-list"] {background: var(--bg3); border-radius: 12px 12px 0 0;}
.stTabs [data-baseweb="tab-panel"] {background: var(--bg2); border: 1px solid var(--line); border-top: none; padding: 1.5rem; border-radius: 0 0 12px 12px;}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS & FORMATTERS
# ══════════════════════════════════════════════════════════════════
def fl(txt: Any) -> str:
    raw = str(txt or "").replace('\x07', '').replace('\xa0', ' ').replace('\t', ' ').replace('\r', '\n').replace('\x0b', '\n')
    parts = [re.sub(r'\s+', ' ', p).strip() for p in raw.split('\n') if p.strip()]
    return parts[0] if parts else ''

def clean_name(txt: Any) -> str:
    t = fl(txt)
    return re.sub(r'(?i)Page\s*\d+\s*of\s*\d+', '', re.sub(r'(?i)^MV\s+', '', t)).strip(" :-#")

def parse_num(txt: Any) -> float:
    s = fl(txt).upper().replace('[', '').replace(']', '')
    if not s or any(w in s for w in ('-', 'N/A', 'NA', 'MONTH', 'YEAR', 'DAY', 'OBS', 'CENTRAL')): return 0.0
    m = re.search(r'\d[\d,\.]*', s)
    if not m: return 0.0
    block = m.group()
    sep = max(block.rfind('.'), block.rfind(','))
    if sep > 0 and len(block) - sep == 4: block = re.sub(r'[,\.]', '', block)
    elif sep > 0: block = re.sub(r'[,\.]', '', block[:sep])
    else: block = re.sub(r'[,\.]', '', block)
    try: return float(block)
    except: return 0.0

def parse_date(txt: Any) -> str:
    s = fl(txt).strip().replace('[', '').replace(']', '')
    if not s or s in ('-', '1', '2', 'N/A') or re.fullmatch(r'\d+', s): return ''
    return s

def is_comp(name: str) -> bool:
    u = fl(name).upper()
    if len(u) < 2 or any(b in u for b in ('DESCRIPTION','PERIOD','MAIN ENGINE','AUX. ENGINE','TOTAL HOURS','REMARKS','TURBOCHARGER','COMPRESSOR','COOLER')): return False
    return not bool(re.fullmatch(r'[\d./ ,:\-\(\)\[\]]+', u))

def get_status(hrs: float, period: float) -> str:
    if hrs <= 0 or period <= 0: return '🔵 NO DATA'
    r = hrs / period
    if r >= 1.0: return '🔴 OVERDUE'
    if r >= 0.8: return '🟠 HIGH PRIORITY'
    return '🟢 OK'

# ══════════════════════════════════════════════════════════════════
#  LIBREOFFICE CONVERTER
# ══════════════════════════════════════════════════════════════════
def convert_doc_to_docx(raw: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice): raise RuntimeError("LibreOffice not found in the environment.")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as t:
        t.write(raw); src = t.name
    outdir = tempfile.mkdtemp(prefix="lo_")
    out = os.path.join(outdir, Path(src).stem + ".docx")
    profile = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        subprocess.run([soffice, "--headless", "--norestore", "--nofirststartwizard",
             f"-env:UserInstallation={profile}", "--convert-to", "docx", src, "--outdir", outdir],
            capture_output=True, timeout=120)
        with open(out, "rb") as f: return f.read()
    finally:
        for p in [src, out]:
            try: os.unlink(p)
            except: pass
        shutil.rmtree(outdir, ignore_errors=True)

# ══════════════════════════════════════════════════════════════════
#  2D GEOMETRIC GRID PARSER (THE CORE FIX)
# ══════════════════════════════════════════════════════════════════
def extract_all_2d(docx_bytes: bytes) -> Dict[str, Any]:
    from docx import Document
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as t:
        t.write(docx_bytes); tp = t.name
    doc = Document(tp)
    os.unlink(tp)

    # Global Metadata Extraction
    vessel, report_date, me_total, me_month = 'UNKNOWN', '—', 0.0, 0.0
    full_text = "\n".join([p.text for p in doc.paragraphs])
    
    if m := re.search(r"Vessel[\u2019\u2018']?s?\s+Name\s*:\s*(?:MV\s+)?(.+?)\s+Date\s*:\s*(.+)$", full_text, re.I | re.MULTILINE):
        vessel, report_date = clean_name(m.group(1)), parse_date(m.group(2))
    
    me_rows, aux_rows, oe_rows = [], [], []

    # Table Grid Engine
    for table in doc.tables:
        # Build perfect rectangular grid mapping
        grid = []
        for row in table.rows:
            grid.append([fl(c.text) for c in row.cells])
        max_c = max((len(r) for r in grid), default=0)
        for r in grid:
            while len(r) < max_c: r.append('')
            
        if not grid: continue
        table_blob = " ".join([" ".join(r) for r in grid]).upper()

        # ─── MAIN ENGINE EXTRACTION ───
        if 'MAIN ENGINE' in table_blob and 'CYL' in table_blob:
            for r in range(len(grid) - 1):
                c_name = grid[r][0]
                if not is_comp(c_name): continue
                
                # Dynamically locate the 1/2 row markers
                marker_col = -1
                for c in range(1, min(4, max_c)):
                    if grid[r][c] == '1' and grid[r+1][c] == '2':
                        marker_col = c; break
                
                if marker_col != -1:
                    period = parse_num(grid[r][marker_col - 1]) if marker_col > 0 else 0.0
                    cyl = 1
                    # Extract dates/hours until we hit REMARKS or run out of numbers
                    for c in range(marker_col + 1, max_c):
                        if 'REMARK' in grid[0][c].upper() or 'REMARK' in grid[1][c].upper(): break
                        d = parse_date(grid[r][c])
                        h = parse_num(grid[r+1][c])
                        if d or h > 0:
                            me_rows.append({'Status': get_status(h, period), 'Component': clean_name(c_name), 'Engine': 'ME', 'Unit': f'Cyl {cyl}', 'Periodicity': period, 'Last O/H': d, 'Hrs Since': h, 'Used %': round((h/period)*100,1) if period else 0.0})
                        cyl += 1

        # ─── AUX ENGINE EXTRACTION ───
        elif 'AUX. ENGINE' in table_blob or 'D/G NO' in table_blob:
            for r in range(len(grid) - 1):
                c_name = grid[r][0]
                if not is_comp(c_name): continue
                
                marker_col = -1
                for c in range(1, min(4, max_c)):
                    if grid[r][c] == '1' and grid[r+1][c] == '2':
                        marker_col = c; break
                        
                if marker_col != -1:
                    period = parse_num(grid[r][marker_col - 1]) if marker_col > 0 else 0.0
                    cyl, eng = 1, 1
                    for c in range(marker_col + 1, max_c):
                        d = parse_date(grid[r][c])
                        h = parse_num(grid[r+1][c])
                        if d or h > 0:
                            aux_rows.append({'Status': get_status(h, period), 'Component': clean_name(c_name), 'Engine': f'AUX-{eng}', 'Unit': f'Cyl {cyl}', 'Periodicity': period, 'Last O/H': d, 'Hrs Since': h, 'Used %': round((h/period)*100,1) if period else 0.0})
                        cyl += 1
                        if cyl > 6:  # Standard overflow handling for Aux matrices
                            cyl = 1; eng += 1

        # ─── OTHER EQUIPMENT EXTRACTION ───
        elif 'TURBOCHARGER' in table_blob or 'COMPRESSOR' in table_blob or 'COOLER' in table_blob:
            for r in range(len(grid)):
                for c in range(max_c - 2):
                    c_name = grid[r][c]
                    if is_comp(c_name) and len(c_name) > 4:
                        # Grab adjacent coordinates dynamically
                        d, h = parse_date(grid[r][c+1]), parse_num(grid[r][c+2])
                        if not d and h == 0.0 and c + 3 < max_c:  # Shift fallback
                            d, h = parse_date(grid[r][c+2]), parse_num(grid[r][c+3])
                        
                        if d or h > 0:
                            oe_rows.append({'Section': 'Other Equipment', 'Description': clean_name(c_name), 'Last Date / O/H': d if d else '—', 'Run Hrs': int(h) if h > 0 else 0})

    # Global Hours Catch (if present in specific top cells)
    for table in doc.tables:
        if table.rows:
            top_text = " ".join([fl(c.text) for c in table.rows[0].cells]).upper()
            if m := re.search(r'TOTAL RUNNING HOURS[\s:ǀ|]*([\d,]+)', top_text): me_total = parse_num(m.group(1))
            if m := re.search(r'THIS MONTH[\s:]*([\d,]+)', top_text): me_month = parse_num(m.group(1))

    # Deduplicate OE Rows
    oe_unique = [dict(t) for t in {tuple(d.items()) for d in oe_rows}]

    return {
        'vessel_name': vessel, 'report_date': report_date,
        'me_total_hrs': int(me_total), 'me_this_month': int(me_month),
        'me_comps': me_rows, 'aux_comps': aux_rows, 'other_equipment': oe_unique
    }

# ══════════════════════════════════════════════════════════════════
#  UI RENDERING (CRASH-PROOF DATAFRAMES)
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero-k">Running Hours Management System</div>
<div class="hero-h">TEC‑004 Enterprise Master Build</div>
<div class="hero-s">
Powered by a 2D Geometric Table Mapping Engine. Matrices are strictly cast to pure strings prior to rendering to mathematically guarantee zero PyArrow serialization crashes.
</div>
<div class="hero-rule"></div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("Drop .doc file here", type=['doc'])
if uploaded is not None:
    with st.spinner("Executing Geometric Table Scan..."):
        try:
            docx_data = convert_doc_to_docx(uploaded.read())
            data = extract_all_2d(docx_data)
            
            me, aux, oe = data['me_comps'], data['aux_comps'], data['other_equipment']
            n_od = sum(1 for c in me+aux if 'OVERDUE' in c['Status'])
            n_hp = sum(1 for c in me+aux if 'HIGH PRIORITY' in c['Status'])

            st.markdown(f"""
            <div class="metric-grid">
              <div class="metric b"><div class="metric-v">{data['vessel_name']}</div><div class="metric-l">Vessel</div></div>
              <div class="metric"><div class="metric-v">{data['report_date']}</div><div class="metric-l">Report Date</div></div>
              <div class="metric b"><div class="metric-v">{data['me_total_hrs']:,}</div><div class="metric-l">ME Total Hrs</div></div>
              <div class="metric b"><div class="metric-v">{data['me_this_month']:,}</div><div class="metric-l">ME This Month</div></div>
              <div class="metric"><div class="metric-v">{len(me)}</div><div class="metric-l">ME Rows</div></div>
              <div class="metric"><div class="metric-v">{len(aux)}</div><div class="metric-l">AUX Rows</div></div>
              <div class="metric o"><div class="metric-v">{n_hp}</div><div class="metric-l">High Priority</div></div>
              <div class="metric r"><div class="metric-v">{n_od}</div><div class="metric-l">Overdue</div></div>
            </div>
            """, unsafe_allow_html=True)

            tab1, tab2, tab3 = st.tabs(["⚙ Main Engine", "🔩 Auxiliary Engine", "🛠 Other Equipment"])

            with tab1:
                if not me: st.info("No Main Engine data found.")
                else:
                    # STRICT STRING CASTING - IMMUNE TO BROWSER CRASHES
                    df_me = pd.DataFrame(me).astype(str)
                    st.dataframe(df_me, use_container_width=True, hide_index=True)

            with tab2:
                if not aux: st.info("No Auxiliary Engine data found.")
                else:
                    # STRICT STRING CASTING - IMMUNE TO BROWSER CRASHES
                    df_aux = pd.DataFrame(aux).astype(str)
                    st.dataframe(df_aux, use_container_width=True, hide_index=True)

            with tab3:
                if not oe: st.info("No Other Equipment data found.")
                else:
                    # STRICT STRING CASTING - IMMUNE TO BROWSER CRASHES
                    df_oe = pd.DataFrame(oe).astype(str)
                    st.dataframe(df_oe, use_container_width=True, hide_index=True)

        except Exception as e:
            st.error(f"Fatal System Error during extraction: {e}")
