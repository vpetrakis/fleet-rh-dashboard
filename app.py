import streamlit as st
import pandas as pd
import numpy as np
import re

# --- 1. THE PINNACLE OF DESIGN: STEALTH LUXURY SAAS INTERFACE ---
st.set_page_config(page_title="Propulsion Command Control", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap');
    
    /* Absolute System Canvas Reset */
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
        background: radial-gradient(circle at 50% -20%, #0F172A 0%, #020617 100%);
        color: #F8FAFC;
    }
    
    /* Hide Default Streamlit Elements for a Pure Native App Feel */
    [data-testid="stHeader"], footer {visibility: hidden;}
    
    /* Animation Framework */
    @keyframes subtleReveal {
        from { opacity: 0; transform: translateY(12px); }
        to { opacity: 1; transform: translateY(0); }
    }
    @keyframes radarPulse {
        0% { border-color: rgba(248, 113, 113, 0.2); box-shadow: 0 0 0 0 rgba(248, 113, 113, 0.2); }
        50% { border-color: rgba(248, 113, 113, 0.6); box-shadow: 0 0 25px 0 rgba(248, 113, 113, 0.25); }
        100% { border-color: rgba(248, 113, 113, 0.2); box-shadow: 0 0 0 0 rgba(248, 113, 113, 0.2); }
    }

    /* Premium Control Deck Grid Layout */
    .dashboard-deck {
        display: flex;
        gap: 24px;
        margin-bottom: 35px;
    }
    
    .dashboard-card {
        flex: 1;
        background: linear-gradient(180deg, rgba(30, 41, 59, 0.4) 0%, rgba(15, 23, 42, 0.6) 100%);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.04);
        border-radius: 16px;
        padding: 26px;
        transition: all 0.4s cubic-bezier(0.16, 1, 0.3, 1);
        animation: subtleReveal 0.6s cubic-bezier(0.16, 1, 0.3, 1) both;
    }
    
    .dashboard-card:hover {
        transform: translateY(-4px);
        background: rgba(30, 41, 59, 0.6);
        border-color: rgba(56, 189, 248, 0.3);
        box-shadow: 0 24px 48px rgba(0, 0, 0, 0.5);
    }
    
    .card-critical {
        animation: subtleReveal 0.6s ease-out both, radarPulse 3s infinite ease-in-out;
    }
    
    .metric-title {
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 2.5px;
        color: #64748B;
        font-weight: 600;
    }
    
    .metric-data {
        font-size: 36px;
        font-weight: 700;
        margin-top: 12px;
        color: #FFFFFF;
        letter-spacing: -0.5px;
    }
    
    /* Sleek Native Typography Elements */
    .mono-value {
        font-family: 'JetBrains Mono', monospace;
        font-weight: 600;
    }

    /* Streamlit Form Input & Component Modifications */
    div[data-testid="stFileUploadDropzone"] {
        background-color: rgba(15, 23, 42, 0.3) !important;
        border: 1px dashed rgba(255, 255, 255, 0.12) !important;
        border-radius: 16px !important;
        padding: 35px !important;
        transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
    }
    
    div[data-testid="stFileUploadDropzone"]:hover {
        border-color: #38BDF8 !important;
        background-color: rgba(30, 41, 59, 0.3) !important;
    }
    
    /* Tabs Interface Polish */
    button[data-testid="stMarkdownContainer"] p {
        font-size: 14px !important;
        font-weight: 600 !important;
        letter-spacing: 0.5px;
    }
    
    /* Clean Layout Containers */
    .section-wrapper {
        background: rgba(15, 23, 42, 0.2);
        border: 1px solid rgba(255, 255, 255, 0.03);
        border-radius: 16px;
        padding: 24px;
        margin-bottom: 25px;
    }
    </style>
""", unsafe_allow_html=True)

THRESHOLD_RED = 1.0
THRESHOLD_YELLOW = 0.8

# --- 2. FORENSIC BACKEND TELEMETRY ENGINE (REPLACES VBA PARSING RULES) ---
def clean_extracted_number(val) -> float:
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
    text = file_bytes.decode('utf-8', errors='ignore')
    
    vessel, date_str = "UNKNOWN ASSET", "UNKNOWN DATE"
    v_m = re.search(r"Vessel’s\s*Name:\s*([^\t\x07\r\n]+)", text, re.IGNORECASE)
    if v_m: vessel = v_m.group(1).strip()
    d_m = re.search(r"Date:\s*([\d\s\w]+)", text, re.IGNORECASE)
    if d_m: date_str = d_m.group(1).strip()

    records = []
    
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

# --- 3. CORE FRONTEND SYSTEM CONTROLS ---
st.markdown("<h1 style='color:#FFFFFF; margin-bottom: 0px; font-weight:700; letter-spacing:-1px;'>Vessel Running Hours Intelligence</h1>", unsafe_allow_html=True)
st.markdown("<p style='color:#64748B; font-size:14px; margin-bottom: 30px;'>Unified operating deck for log ingestion, data auditing, and telemetry validation.</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("", type=["doc"])

if uploaded_file is not None:
    if 'raw_df' not in st.session_state or st.session_state.get('uploaded_file_name') != uploaded_file.name:
        vessel_name, report_date, raw_dataframe = execute_stream_ingestion(uploaded_file.read())
        st.session_state.raw_df = raw_dataframe
        st.session_state.vessel_name = vessel_name
        st.session_state.report_date = report_date
        st.session_state.uploaded_file_name = uploaded_file.name

    # --- NATIVE INGESTION GUARD: IMMERSIVE MAIN LAYOUT FLOW ---
    st.markdown("<div class='section-wrapper'>", unsafe_allow_html=True)
    st.markdown("<h3 style='color:#FFFFFF; margin-top:0px; margin-bottom:4px; font-size:16px; font-weight:600;'>🛠️ Real-Time Telemetry Audit Ledger</h3>", unsafe_allow_html=True)
    st.markdown("<p style='color:#64748B; font-size:12px; margin-bottom:16px;'>Data firewall active. Review or manually overwrite parsed machinery metrics directly inside the inline grid grid below before downstream dashboard rendering.</p>", unsafe_allow_html=True)
    
    verified_df = st.data_editor(
        st.session_state.raw_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Subsystem": st.column_config.TextColumn("Machinery Subsystem", disabled=True),
            "Component Group": st.column_config.TextColumn("Equipment Group", disabled=True),
            "Location Unit": st.column_config.TextColumn("Location Location", disabled=True),
            "Baseline Interval (Hrs)": st.column_config.NumberColumn("Maintenance Limit", format="%d"),
            "Current Running Hours": st.column_config.NumberColumn("Running Hours (Editable Field)", format="%.1f")
        }
    )
    st.markdown("</div>", unsafe_allow_html=True)
    
    # Process core computations live on the adjusted inline frame
    if not verified_df.empty:
        df = verified_df.copy()
        df['Lifecycle Consumed (%)'] = np.where(df['Baseline Interval (Hrs)'] > 0, df['Current Running Hours'] / df['Baseline Interval (Hrs)'], 0.0)
        
        conditions = [(df['Current Running Hours'] == 0), (df['Lifecycle Consumed (%)'] >= THRESHOLD_RED), (df['Lifecycle Consumed (%)'] >= THRESHOLD_YELLOW)]
        df['Status'] = np.select(conditions, ['NO DATA', 'OVERDUE', 'HIGH PRIORITY'], default='OK')

        crit_df = df[df['Status'] == 'OVERDUE']
        warn_df = df[df['Status'] == 'HIGH PRIORITY']
        health_factor = max(0.0, 100.0 - ((len(crit_df) * 3.0 + len(warn_df) * 1.0) / len(df) * 100))

        # Executive Glassmorphic Metric Display Row
        st.markdown(f"""
            <div class="dashboard-deck">
                <div class="dashboard-card">
                    <div class="metric-title">Asset Target Profile</div>
                    <div class="metric-data" style="color:#38BDF8;">{st.session_state.vessel_name}</div>
                    <div style="color:#475569; font-size:11px; margin-top:8px; font-weight:600;">Log Reference: {st.session_state.report_date}</div>
                </div>
                <div class="dashboard-card {'dashboard-card card-critical' if len(crit_df)>0 else ''}">
                    <div class="metric-title">Critical Interrupt vectors</div>
                    <div class="metric-data" style="color:#F87171;">{len(crit_df)} Items</div>
                    <div style="color:#475569; font-size:11px; margin-top:8px; font-weight:600;">Immediate Overhaul Action Demanded</div>
                </div>
                <div class="dashboard-card">
                    <div class="metric-title">Impending Lifecycle Risks</div>
                    <div class="metric-data" style="color:#FB923C;">{len(warn_df)} Items</div>
                    <div style="color:#475569; font-size:11px; margin-top:8px; font-weight:600;">Approaching Mechanical Limitation</div>
                </div>
                <div class="dashboard-card">
                    <div class="metric-title">Aggregated Fleet Health Index</div>
                    <div class="metric-data" style="color:#34D399;">{health_factor:.1f}%</div>
                    <div style="color:#475569; font-size:11px; margin-top:8px; font-weight:600;">Total Mechanical Structural Score</div>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # --- CATEGORIZED SYSTEM TAB DECK ---
        tab1, tab2, tab3 = st.tabs(["🔥 Risk Exceptions Matrix", "🔩 Main Plant Hierarchy", "⚡ Auxiliary Generation Plant"])
        
        ui_table_config = {
            "Subsystem": st.column_config.TextColumn("Subsystem"),
            "Component Group": st.column_config.TextColumn("Component Classification"),
            "Location Unit": st.column_config.TextColumn("Location"),
            "Baseline Interval (Hrs)": st.column_config.NumberColumn("Interval Limit (Hrs)", format="%d"),
            "Current Running Hours": st.column_config.NumberColumn("Running Hours", format="%.1f"),
            "Lifecycle Consumed (%)": st.column_config.ProgressColumn("Fatigue Curve", format="%.1f%%", min_value=0.0, max_value=1.5),
            "Status": st.column_config.TextColumn("Diagnostic Condition State")
        }

        def color_row_states(val):
            if val == 'OVERDUE': return 'background-color: rgba(239, 68, 68, 0.12); color: #F87171; font-weight: 600;'
            elif val == 'HIGH PRIORITY': return 'background-color: rgba(251, 146, 60, 0.12); color: #FB923C; font-weight: 600;'
            return ''

        with tab1:
            risk_df = df[df['Status'].isin(['OVERDUE', 'HIGH PRIORITY'])].sort_values(by='Lifecycle Consumed (%)', ascending=False)
            if not risk_df.empty:
                st.dataframe(risk_df.style.map(color_row_states, subset=['Status']), use_container_width=True, hide_index=True, column_config=ui_table_config)
            else:
                st.success("Data validation complete: All mechanical components operating inside safe lifecycle boundaries.")

        with tab2:
            me_display = df[df['Subsystem'] == 'MAIN PROPULSION'].sort_values(by=['Component Group', 'Location Unit'])
            st.dataframe(me_display.style.map(color_row_states, subset=['Status']), use_container_width=True, hide_index=True, column_config=ui_table_config)

        with tab3:
            aux_display = df[df['Subsystem'].str.contains('AUX')].sort_values(by=['Subsystem', 'Component Group', 'Location Unit'])
            st.dataframe(aux_display.style.map(color_row_states, subset=['Status']), use_container_width=True, hide_index=True, column_config=ui_table_config)
