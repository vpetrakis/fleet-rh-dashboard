import streamlit as st
import pandas as pd
import numpy as np
import re

# --- 1. PREMIUM EXECUTIVE INTERFACE CANVAS (STEALTH LUXURY DESIGN) ---
st.set_page_config(page_title="Vessel Operational Command Tower", page_icon="⚓", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;700&family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap');
    
    /* Global Canvas Reset - Ultra-Premium Dark Minimalist */
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
        background: radial-gradient(circle at 50% 10%, #0F172A 0%, #020617 100%);
        color: #F1F5F9;
    }
    
    /* High-End Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #030712;
        border-right: 1px solid rgba(255, 255, 255, 0.05);
    }

    /* Keyframe Layer Transitions */
    @keyframes smoothEntrance {
        from { opacity: 0; transform: scale(0.99) translateY(15px); }
        to { opacity: 1; transform: scale(1) translateY(0); }
    }
    @keyframes pulseCritical {
        0% { border-color: rgba(239, 68, 68, 0.2); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.2); }
        50% { border-color: rgba(239, 68, 68, 0.8); box-shadow: 0 0 20px 0 rgba(239, 68, 68, 0.3); }
        100% { border-color: rgba(239, 68, 68, 0.2); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.2); }
    }

    /* Glassmorphism Metric Grid */
    .deck-container {
        display: flex;
        gap: 24px;
        margin-bottom: 35px;
    }
    .premium-card {
        flex: 1;
        background: rgba(15, 23, 42, 0.6);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.05);
        border-radius: 16px;
        padding: 28px;
        transition: all 0.4s cubic-bezier(0.16, 1, 0.3, 1);
        animation: smoothEntrance 0.6s cubic-bezier(0.16, 1, 0.3, 1) both;
    }
    .premium-card:hover {
        transform: translateY(-4px);
        background: rgba(30, 41, 59, 0.5);
        border-color: rgba(56, 189, 248, 0.4);
        box-shadow: 0 20px 30px rgba(0, 0, 0, 0.4);
    }
    .card-pulse-critical {
        animation: smoothEntrance 0.6s ease-out both, pulseCritical 3s infinite ease-in-out;
    }
    .meta-lbl {
        font-size: 10px;
        text-transform: uppercase;
        letter-spacing: 2.5px;
        color: #94A3B8;
        font-weight: 600;
    }
    .meta-val {
        font-size: 34px;
        font-weight: 700;
        margin-top: 14px;
        color: #FFFFFF;
        letter-spacing: -0.5px;
    }
    
    /* Code/Numeric Fine Controls */
    .mono-txt {
        font-family: 'JetBrains Mono', monospace;
        font-weight: 700;
    }
    
    /* Streamlit UI Component Customization Overrides */
    button[data-testid="stMarkdownContainer"] p {
        font-size: 14px !important;
        font-weight: 600 !important;
        letter-spacing: 0.5px;
    }
    div[data-testid="stFileUploadDropzone"] {
        background-color: rgba(15, 23, 42, 0.4) !important;
        border: 1px dashed rgba(255, 255, 255, 0.15) !important;
        border-radius: 16px !important;
        padding: 40px !important;
        transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
    }
    div[data-testid="stFileUploadDropzone"]:hover {
        border-color: #38BDF8 !important;
        background-color: rgba(30, 41, 59, 0.4) !important;
    }
    </style>
""", unsafe_allow_html=True)

THRESHOLD_RED = 1.0
THRESHOLD_YELLOW = 0.8

# --- 2. FORTRAN BACKEND DATA PIPELINE (100% INTEGRITY PARSER) ---
def clean_extracted_number(val) -> float:
    """Advanced string cleaner. Strips telemetry characters, fixes European decimals and typo bounds."""
    if pd.isna(val) or val == "" or val == "-":
        return 0.0
    s = str(val).upper().strip().replace('[', '').replace(']', '').replace(' ', '')
    if any(x in s for x in ["N/A", "CENTRAL", "OBSERVATION", "COOLER"]):
        return 0.0
    match = re.search(r'[\d\.\,]+', s)
    if not match:
        return 0.0
    num_str = match.group(0)
    if ',' in num_str and '.' in num_str:
        num_str = num_str.replace(',', '')
    elif num_str.count(',') == 1:
        if len(num_str.split(',')[1]) != 3:
            num_str = num_str.replace(',', '.')
        else:
            num_str = num_str.replace(',', '')
    try:
        res = float(num_str)
        return 0.0 if res > 250000 else res
    except ValueError:
        return 0.0

def execute_stream_ingestion(file_bytes) -> tuple:
    """Primary file reader looking for Word cell separation markers (\x07)."""
    text = file_bytes.decode('utf-8', errors='ignore')
    
    vessel, date_str = "UNKNOWN ASSET", "UNKNOWN DATE"
    v_m = re.search(r"Vessel’s\s*Name:\s*([^\t\x07\r\n]+)", text, re.IGNORECASE)
    if v_m: vessel = v_m.group(1).strip()
    d_m = re.search(r"Date:\s*([\d\s\w]+)", text, re.IGNORECASE)
    if d_m: date_str = d_m.group(1).strip()

    records = []
    
    # Machinery Categorization Mapping Tables
    me_definitions = [
        ("CYLINDER COVER", 16000), ("PISTON ASSEMBLY", 16000), ("STUFFING BOX", 16000), 
        ("PISTON CROWN", 32000), ("CYLINDER LINER", 16000), ("EXHAUST VALVE", 16000), 
        ("STARTING VALVE", 12000), ("SAFETY VALVE", 12000), ("FUEL VALVES", 8000), 
        ("FUEL PUMP", 16000), ("PLUNGER AND BARREL(RENEWAL)", 32000), ("FUEL PUMP SUCTION VALVE", 8000),
        ("FUEL PUMP PUNCTURE VALVE", 8000), ("CROSSHEAD BEARINGS", 32000), ("BOTTOM END BEARINGS", 32000), ("MAIN BEARINGS", 32000)
    ]
    
    for comp, periodicity in me_definitions:
        block = re.search(rf"{re.escape(comp)}\x07.*?\x072\x07([^\x0d]+)", text, re.DOTALL | re.IGNORECASE)
        if block:
            tokens = [t.strip() for t in block.group(1).split('\x07') if t.strip()]
            for idx, h_val in enumerate(tokens[:7]):
                records.append({
                    "Subsystem": "MAIN PROPULSION", "Component Group": comp, "Location Unit": f"Cyl No.{idx+1}",
                    "Baseline Interval (Hrs)": float(periodicity), "Current Running Hours": clean_extracted_number(h_val)
                })

    aux_definitions = [
        ("Cylinder Head", 12000), ("Piston", 10000), ("Connecting Rod", 10000), 
        ("Cylinder Liners", 10000), ("Fuel Valves (1)", 2000), ("Fuel Pumps", 5000),
        ("Crank Pin Bearing", 12000), ("Main Bearing", 12000), ("Adjust Valve Head Clearance", 1200)
    ]
    
    for comp, periodicity in aux_definitions:
        block = re.search(rf"{re.escape(comp)}\x07.*?\x072\x07([^\x0d]+)", text, re.DOTALL | re.IGNORECASE)
        if block:
            tokens = [t.strip() for t in block.group(1).split('\x07') if t.strip()]
            for i in range(1, 4):
                for cyl in range(6):
                    token_idx = ((i - 1) * 6) + cyl
                    if token_idx < len(tokens):
                        records.append({
                            "Subsystem": f"AUX GENERATOR No.{i}", "Component Group": comp.replace(" (1)", ""), "Location Unit": f"Cyl No.{cyl+1}",
                            "Baseline Interval (Hrs)": float(periodicity), "Current Running Hours": clean_extracted_number(tokens[token_idx])
                        })

    # Miscellaneous/Auxiliary Equipment Sub-Matrix
    misc_definitions = [
        ("GENERAL O/H", 16000, "TURBOCHARGER", "M/E T/C"), ("BALANCING OF ROTOR SHAFT", 32000, "TURBOCHARGER", "M/E T/C"),
        ("AIR COOLER CLEANING", 4000, "COOLERS", "M/E Air Cooler"), ("AIR COND. COMPRESSOR NO.1", 10000, "A/C SYSTEMS", "Compressor 1"),
        ("AIR COND. COMPRESSOR NO.2", 10000, "A/C SYSTEMS", "Compressor 2"), ("REFRIGERATION COMPRESSOR NO.1", 10000, "REFRIGERATION", "Compressor 1")
    ]
    for label, per, sub, unit in misc_definitions:
        m = re.search(rf"{re.escape(label)}\x07.*?\x07([\d\.\,\s\[\]]+)\x07", text, re.IGNORECASE)
        if m:
            records.append({
                "Subsystem": sub, "Component Group": label, "Location Unit": unit,
                "Baseline Interval (Hrs)": float(per), "Current Running Hours": clean_extracted_number(m.group(1))
            })

    return vessel, date_str, pd.DataFrame(records)

# --- 3. BRANDED SIDEBAR CONSOLE ---
with st.sidebar:
    st.markdown("<h2 style='color:#38BDF8; font-family:monospace; letter-spacing: 1px;'>⚓ PROPULSION OS</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color:#64748B; font-size:12px;'>Enterprise Marine Automation</p>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("### Active Module")
    st.markdown("<span style='color:#F1F5F9; font-weight:600;'>📊 Core Diagnostics</span>", unsafe_allow_html=True)
    st.markdown("---")
    st.caption("Fleet Analytics Platform Engine | v3.5.0")

# --- 4. MAIN PROPULSION CONTROL DASHBOARD ---
st.markdown("<h1 style='color:#FFFFFF; margin-bottom: 0px; font-weight:700; letter-spacing:-1px;'>Vessel Running Hours Intelligence</h1>", unsafe_allow_html=True)
st.markdown("<p style='color:#94A3B8; font-size:14px; margin-bottom: 30px;'>Ingest legacy telemetry blocks to isolate component degradation profiles instantly.</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("", type=["doc"])

if uploaded_file is not None:
    if 'raw_df' not in st.session_state or st.session_state.get('uploaded_file_name') != uploaded_file.name:
        # Initial parsing cycle and session memory caching
        vessel_name, report_date, raw_dataframe = execute_stream_ingestion(uploaded_file.read())
        st.session_state.raw_df = raw_dataframe
        st.session_state.vessel_name = vessel_name
        st.session_state.report_date = report_date
        st.session_state.uploaded_file_name = uploaded_file.name

    # --- HUMAN-IN-THE-LOOP STAGING GATE ---
    st.markdown("### 🛠️ Telemetry Ingestion Guard")
    st.caption("Verify or adjust parsed runtime logs before committing entries to the operations ledger.")
    
    verified_df = st.data_editor(
        st.session_state.raw_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Subsystem": st.column_config.TextColumn("Machinery Subsystem", disabled=True),
            "Component Group": st.column_config.TextColumn("Equipment Group", disabled=True),
            "Location Unit": st.column_config.TextColumn("Location Unit", disabled=True),
            "Baseline Interval (Hrs)": st.column_config.NumberColumn("Maintenance Limit", format="%d"),
            "Current Running Hours": st.column_config.NumberColumn("Extracted Value (Editable)", format="%.1f")
        }
    )
    
    # Process analytics metrics on the active verified data frame
    if not verified_df.empty:
        df = verified_df.copy()
        df['Lifecycle Consumed (%)'] = np.where(df['Baseline Interval (Hrs)'] > 0, df['Current Running Hours'] / df['Baseline Interval (Hrs)'], 0.0)
        
        # Mapping component status parameters
        conditions = [(df['Current Running Hours'] == 0), (df['Lifecycle Consumed (%)'] >= THRESHOLD_RED), (df['Lifecycle Consumed (%)'] >= THRESHOLD_YELLOW)]
        df['Status'] = np.select(conditions, ['NO DATA', 'OVERDUE', 'HIGH PRIORITY'], default='OK')

        # Split DataFrames for Dashboard Reporting
        crit_df = df[df['Status'] == 'OVERDUE']
        warn_df = df[df['Status'] == 'HIGH PRIORITY']
        health_factor = max(0.0, 100.0 - ((len(crit_df) * 3.0 + len(warn_df) * 1.0) / len(df) * 100))

        st.markdown("---")
        
        # High-End Luxury Metric Grid Injection
        st.markdown(f"""
            <div class="deck-container">
                <div class="premium-card">
                    <div class="meta-lbl">Operational Context</div>
                    <div class="meta-val" style="color:#38BDF8;">{st.session_state.vessel_name}</div>
                    <div style="color:#64748B; font-size:11px; margin-top:8px; font-weight:500;">Reference Date: {st.session_state.report_date}</div>
                </div>
                <div class="premium-card {'card-pulse-critical' if len(crit_df)>0 else ''}">
                    <div class="meta-lbl">Critical Interrupts</div>
                    <div class="meta-val" style="color:#F87171;">{len(crit_df)} Items</div>
                    <div style="color:#64748B; font-size:11px; margin-top:8px; font-weight:500;">Immediate Overhaul Obligatory</div>
                </div>
                <div class="premium-card">
                    <div class="meta-lbl">Pending Risk Vectors</div>
                    <div class="meta-val" style="color:#FB923C;">{len(warn_df)} Items</div>
                    <div style="color:#64748B; font-size:11px; margin-top:8px; font-weight:500;">Approaching Mechanical Threshold</div>
                </div>
                <div class="premium-card">
                    <div class="meta-lbl">Calculated Fleet Health</div>
                    <div class="meta-val" style="color:#34D399;">{health_factor:.1f}%</div>
                    <div style="color:#64748B; font-size:11px; margin-top:8px; font-weight:500;">Aggregated Fatigue Configuration</div>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # --- RECONFIGURED CLUSTER VIEWS ---
        tab1, tab2, tab3 = st.tabs(["🔥 Risk Exceptions Matrix", "🔩 Main Propulsion Hierarchy", "⚡ Auxiliary Generation Plant"])
        
        ui_config = {
            "Subsystem": st.column_config.TextColumn("Subsystem"),
            "Component Group": st.column_config.TextColumn("Component Classification"),
            "Location Unit": st.column_config.TextColumn("Location"),
            "Baseline Interval (Hrs)": st.column_config.NumberColumn("Interval Limit (Hrs)", format="%d"),
            "Current Running Hours": st.column_config.NumberColumn("Running Hours", format="%.1f"),
            "Lifecycle Consumed (%)": st.column_config.ProgressColumn("Fatigue Spectrum", format="%.1f%%", min_value=0.0, max_value=1.5),
            "Status": st.column_config.TextColumn("Diagnostic State")
        }

        def style_matrix_cells(val):
            if val == 'OVERDUE': return 'background-color: rgba(239, 68, 68, 0.15); color: #F87171; font-weight: bold;'
            elif val == 'HIGH PRIORITY': return 'background-color: rgba(251, 146, 60, 0.15); color: #FB923C; font-weight: bold;'
            return ''

        with tab1:
            st.markdown("<p style='color:#94A3B8; margin-top:10px;'>Isolating elements that have violated safe lifecycle parameters.</p>", unsafe_allow_html=True)
            risk_df = df[df['Status'].isin(['OVERDUE', 'HIGH PRIORITY'])].sort_values(by='Lifecycle Consumed (%)', ascending=False)
            if not risk_df.empty:
                st.dataframe(risk_df.style.map(style_matrix_cells, subset=['Status']), use_container_width=True, hide_index=True, column_config=ui_config)
            else:
                st.success("All mechanical elements operating inside safe parameters.")

        with tab2:
            st.markdown("<p style='color:#94A3B8; margin-top:10px;'>Sequential engineering view mapping components directly along the main crankshaft profile.</p>", unsafe_allow_html=True)
            me_display = df[df['Subsystem'] == 'MAIN PROPULSION'].sort_values(by=['Component Group', 'Location Unit'])
            st.dataframe(me_display.style.map(style_matrix_cells, subset=['Status']), use_container_width=True, hide_index=True, column_config=ui_config)

        with tab3:
            st.markdown("<p style='color:#94A3B8; margin-top:10px;'>Comprehensive distribution matrix for active prime movers and auxiliary engines.</p>", unsafe_allow_html=True)
            aux_display = df[df['Subsystem'].str.contains('AUX')].sort_values(by=['Subsystem', 'Component Group', 'Location Unit'])
            st.dataframe(aux_display.style.map(style_matrix_cells, subset=['Status']), use_container_width=True, hide_index=True, column_config=ui_config)
