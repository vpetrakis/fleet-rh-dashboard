import streamlit as st
import pandas as pd
import numpy as np
import re

# --- 1. THE ARCHITECTURAL APEX: NATIVE ULTRA-PREMIUM COMMAND COCKPIT ---
st.set_page_config(page_title="Propulsion Command Control", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600;700&family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap');
    
    /* Absolute Canvas Reset - High-End Minimalist Stealth Theme */
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
        background: radial-gradient(circle at 50% -20%, #0F172A 0%, #020617 100%);
        color: #F8FAFC;
    }
    
    [data-testid="stHeader"], footer {visibility: hidden;}
    
    /* Advanced Interface Presentation Animations */
    @keyframes smoothReveal {
        from { opacity: 0; transform: translateY(20px) scale(0.98); }
        to { opacity: 1; transform: translateY(0) scale(1); }
    }
    @keyframes radarPulse {
        0% { border-color: rgba(248, 113, 113, 0.15); box-shadow: 0 0 0 0 rgba(248, 113, 113, 0.15); }
        50% { border-color: rgba(248, 113, 113, 0.5); box-shadow: 0 0 25px 0 rgba(248, 113, 113, 0.2); }
        100% { border-color: rgba(248, 113, 113, 0.15); box-shadow: 0 0 0 0 rgba(248, 113, 113, 0.15); }
    }

    /* Premium Glassmorphic KPI Row Layout */
    .dashboard-deck {
        display: flex;
        gap: 24px;
        margin-bottom: 30px;
    }
    
    .dashboard-card {
        flex: 1;
        background: linear-gradient(180deg, rgba(30, 41, 59, 0.3) 0%, rgba(15, 23, 42, 0.5) 100%);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.05);
        border-radius: 16px;
        padding: 26px;
        transition: all 0.4s cubic-bezier(0.16, 1, 0.3, 1);
        animation: smoothReveal 0.6s cubic-bezier(0.16, 1, 0.3, 1) both;
    }
    
    .dashboard-card:hover {
        transform: translateY(-4px);
        background: rgba(30, 41, 59, 0.5);
        border-color: rgba(56, 189, 248, 0.3);
        box-shadow: 0 24px 48px rgba(0, 0, 0, 0.45);
    }
    
    /* Cascading Animation Delays */
    .dashboard-deck > div:nth-child(1) { animation-delay: 0.1s; }
    .dashboard-deck > div:nth-child(2) { animation-delay: 0.2s; }
    .dashboard-deck > div:nth-child(3) { animation-delay: 0.3s; }
    .dashboard-deck > div:nth-child(4) { animation-delay: 0.4s; }

    .card-critical { animation: smoothReveal 0.6s ease-out 0.2s both, radarPulse 3s infinite ease-in-out; }
    
    .metric-title { font-size: 11px; text-transform: uppercase; letter-spacing: 2px; color: #64748B; font-weight: 600; }
    .metric-data { font-size: 38px; font-weight: 700; margin-top: 8px; color: #FFFFFF; font-family: 'JetBrains Mono', monospace; letter-spacing: -1px; }

    /* Custom Spreadsheet Wrapper Styles */
    div[data-testid="stDataFrame"] {
        border: 1px solid rgba(255, 255, 255, 0.05) !important;
        border-radius: 12px !important;
        overflow: hidden !important;
        background-color: #090D1A !important;
        animation: smoothReveal 0.8s ease-out 0.5s both;
    }
    
    /* File Upload Area Customizations */
    div[data-testid="stFileUploadDropzone"] {
        background-color: rgba(15, 23, 42, 0.3) !important;
        border: 1px dashed rgba(56, 189, 248, 0.2) !important;
        border-radius: 16px !important;
        padding: 40px !important;
        animation: smoothReveal 0.6s ease-out both;
    }
    div[data-testid="stFileUploadDropzone"]:hover {
        border-color: #38BDF8 !important;
        background-color: rgba(30, 41, 59, 0.4) !important;
        box-shadow: 0 0 30px rgba(56, 189, 248, 0.1);
    }
    
    /* Modernized Tab Layout Styling */
    button[data-testid="stMarkdownContainer"] p {
        font-size: 14px !important; font-weight: 600 !important; letter-spacing: 0.5px;
    }
    
    /* Custom Native Chart Container */
    .native-chart-box {
        background: rgba(15, 23, 42, 0.3);
        border: 1px solid rgba(255, 255, 255, 0.05);
        border-radius: 16px;
        padding: 24px;
        animation: smoothReveal 0.7s ease-out 0.4s both;
        height: 100%;
    }
    </style>
""", unsafe_allow_html=True)

THRESHOLD_RED = 1.0
THRESHOLD_YELLOW = 0.8

# --- 2. FORENSIC TELEMETRY STREAM ENGINE (100% LINEAR CELL MAPPING) ---
def clean_extracted_number(val) -> float:
    if pd.isna(val) or val == "" or val == "-": return 0.0
    s = str(val).upper().strip().replace('[', '').replace(']', '').replace(' ', '')
    if any(x in s for x in ["N/A", "CENTRAL", "OBSERVATION", "COOLER"]): return 0.0
    match = re.search(r'[\d\.\,]+', s)
    if not match: return 0.0
    num_str = match.group(0)
    if ',' in num_str and '.' in num_str: num_str = num_str.replace(',', '')
    elif num_str.count(',') == 1:
        if len(num_str.split(',')[1]) != 3: num_str = num_str.replace(',', '.')
        else: num_str = num_str.replace(',', '')
    try:
        res = float(num_str)
        return 0.0 if res > 250000 else res
    except ValueError: return 0.0

def execute_stream_ingestion(file_bytes) -> tuple:
    text = file_bytes.decode('utf-8', errors='ignore')
    vessel, date_str = "UNKNOWN ASSET", "UNKNOWN DATE"
    v_m = re.search(r"Vessel’s\s*Name:\s*([^\t\x07\r\n]+)", text, re.IGNORECASE)
    if v_m: vessel = v_m.group(1).strip()
    d_m = re.search(r"Date:\s*([\d\s\w]+)", text, re.IGNORECASE)
    if d_m: date_str = d_m.group(1).strip()

    raw_tokens = text.split('\x07')
    tokens = [t.replace('\r', '').replace('\n', '').strip() for t in raw_tokens]
    records = []
    
    me_definitions = [
        ("CYLINDER COVER", 16000), ("PISTON ASSEMBLY", 16000), ("STUFFING BOX", 16000), 
        ("PISTON CROWN", 32000), ("CYLINDER LINER", 16000), ("EXHAUST VALVE", 16000), 
        ("STARTING VALVE", 12000), ("SAFETY VALVE", 12000), ("FUEL VALVES", 8000), 
        ("FUEL PUMP", 16000), ("PLUNGER AND BARREL(RENEWAL)", 32000), ("FUEL PUMP SUCTION VALVE", 8000),
        ("FUEL PUMP PUNCTURE VALVE", 8000), ("CROSSHEAD BEARINGS", 32000), ("BOTTOM END BEARINGS", 32000), ("MAIN BEARINGS", 32000)
    ]
    for comp, periodicity in me_definitions:
        for idx, token in enumerate(tokens):
            if token.upper() == comp:
                scan_idx = idx + 1
                while scan_idx < len(tokens) and tokens[scan_idx] != '2': scan_idx += 1
                if scan_idx < len(tokens) and tokens[scan_idx] == '2':
                    for cyl in range(7):
                        if scan_idx + 1 + cyl < len(tokens):
                            records.append({
                                "Subsystem": "MAIN ENGINE", "Component Group": comp, "Location Unit": f"Cyl No.{cyl+1}",
                                "Baseline Interval (Hrs)": float(periodicity), "Current Running Hours": clean_extracted_number(tokens[scan_idx + 1 + cyl])
                            })
                break

    aux_definitions = [
        ("Cylinder Head", 12000), ("Piston", 10000), ("Connecting Rod", 10000), 
        ("Cylinder Liners", 10000), ("Fuel Valves", 2000), ("Fuel Pumps", 5000),
        ("Crank Pin Bearing", 12000), ("Main Bearing", 12000), ("Adjust Valve Head Clearance", 1200)
    ]
    for comp, periodicity in aux_definitions:
        for idx, token in enumerate(tokens):
            if comp.upper() in token.upper():
                scan_idx = idx + 1
                while scan_idx < len(tokens) and tokens[scan_idx] != '2': scan_idx += 1
                if scan_idx < len(tokens) and tokens[scan_idx] == '2':
                    for i in range(1, 4):
                        for cyl in range(6):
                            token_offset = ((i - 1) * 6) + cyl
                            if scan_idx + 1 + token_offset < len(tokens):
                                records.append({
                                    "Subsystem": "AUX ENGINE", "Component Group": comp.upper(), "Location Unit": f"DG No.{i} - Cyl No.{cyl+1}",
                                    "Baseline Interval (Hrs)": float(periodicity), "Current Running Hours": clean_extracted_number(tokens[scan_idx + 1 + token_offset])
                                })
                break

    misc_definitions = [
        ("GENERAL O/H", 16000, "OTHER EQUIPMENT", "M/E T/C"), ("BALANCING OF ROTOR SHAFT", 32000, "OTHER EQUIPMENT", "M/E T/C"),
        ("AIR COOLER CLEANING", 4000, "OTHER EQUIPMENT", "M/E Air Cooler"), ("AIR COND. COMPRESSOR NO.1", 10000, "OTHER EQUIPMENT", "Compressor 1"),
        ("AIR COND. COMPRESSOR NO.2", 10000, "OTHER EQUIPMENT", "Compressor 2"), ("REFRIGERATION COMPRESSOR NO.1", 10000, "OTHER EQUIPMENT", "Compressor 1")
    ]
    for label, per, sub, unit in misc_definitions:
        for idx, token in enumerate(tokens):
            if label.upper() in token.upper():
                scan_idx = idx + 1
                hrs_val = "0"
                if scan_idx < len(tokens):
                    t1 = tokens[scan_idx]
                    if any(m in t1.upper() for m in ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC", "/"]):
                        if scan_idx + 1 < len(tokens): hrs_val = tokens[scan_idx + 1]
                    else:
                        if scan_idx + 1 < len(tokens):
                            t2 = tokens[scan_idx + 1]
                            if any(m in t2.upper() for m in ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC", "/"]):
                                if scan_idx + 2 < len(tokens): hrs_val = tokens[scan_idx + 2]
                            else: hrs_val = t2
                records.append({
                    "Subsystem": sub, "Component Group": label, "Location Unit": unit,
                    "Baseline Interval (Hrs)": float(per), "Current Running Hours": clean_extracted_number(hrs_val)
                })
                break

    return vessel, date_str, pd.DataFrame(records)

# --- 3. FRONTEND NAVIGATION CONTROL DECK ---
st.markdown("<h1 style='color:#FFFFFF; margin-bottom: 0px; font-weight:700; letter-spacing:-1.5px;'>Vessel Telemetry Operations</h1>", unsafe_allow_html=True)
st.markdown("<p style='color:#64748B; font-size:15px; margin-bottom: 30px;'>Drag legacy telemetry blocks below for automated ingestion, live visual diagnostics, and component structural analysis.</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("", type=["doc"])

if uploaded_file is not None:
    with st.spinner("Decrypting binary streams and rendering telemetry matrices..."):
        vessel_name, report_date, df = execute_stream_ingestion(uploaded_file.read())
    
    if not df.empty:
        df['Lifecycle Consumed (%)'] = np.where(df['Baseline Interval (Hrs)'] > 0, df['Current Running Hours'] / df['Baseline Interval (Hrs)'], 0.0)
        conditions = [(df['Current Running Hours'] == 0), (df['Lifecycle Consumed (%)'] >= THRESHOLD_RED), (df['Lifecycle Consumed (%)'] >= THRESHOLD_YELLOW)]
        df['Status'] = np.select(conditions, ['NO DATA', 'OVERDUE', 'HIGH PRIORITY'], default='OK')

        crit_df = df[df['Status'] == 'OVERDUE']
        warn_df = df[df['Status'] == 'HIGH PRIORITY']
        health_factor = max(0.0, 100.0 - ((len(crit_df) * 3.0 + len(warn_df) * 1.0) / len(df) * 100))

        # --- ANIMATED KPI DECK ---
        st.markdown(f"""
            <div class="dashboard-deck">
                <div class="dashboard-card">
                    <div class="metric-title">Active Target Profile</div>
                    <div class="metric-data" style="color:#38BDF8;">{vessel_name}</div>
                    <div style="color:#475569; font-size:12px; margin-top:8px; font-weight:600;">Log Reference: {report_date}</div>
                </div>
                <div class="dashboard-card {'card-critical' if len(crit_df)>0 else ''}">
                    <div class="metric-title">Critical Overhauls</div>
                    <div class="metric-data" style="color:#F87171;">{len(crit_df)}<span style="font-size:16px; color:#64748B; font-weight:500;"> Items</span></div>
                    <div style="color:#475569; font-size:12px; margin-top:8px; font-weight:600;">Immediate Action Mandatory</div>
                </div>
                <div class="dashboard-card">
                    <div class="metric-title">High Priority Risks</div>
                    <div class="metric-data" style="color:#FB923C;">{len(warn_df)}<span style="font-size:16px; color:#64748B; font-weight:500;"> Items</span></div>
                    <div style="color:#475569; font-size:12px; margin-top:8px; font-weight:600;">Approaching Absolute Threshold</div>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # --- NATIVE CSS VISUALIZATIONS (ZERO DEPENDENCIES) ---
        col1, col2 = st.columns([1, 2])
        
        # Color Logic for Health Bar
        health_color = "#34D399" if health_factor > 80 else ("#FB923C" if health_factor > 50 else "#F87171")
        
        with col1:
            st.markdown(f"""
                <div class="native-chart-box" style="text-align: center; display: flex; flex-direction: column; justify-content: center;">
                    <div style="color: #94A3B8; font-size: 13px; text-transform: uppercase; letter-spacing: 2px; font-weight: 600; margin-bottom: 25px;">Aggregate Health Index</div>
                    <div style="font-size: 54px; font-weight: 700; color: {health_color}; font-family: 'JetBrains Mono', monospace; line-height: 1;">{health_factor:.1f}<span style="font-size: 24px;">%</span></div>
                    <div style="width: 100%; background: rgba(255,255,255,0.05); border-radius: 10px; height: 12px; margin-top: 30px; overflow: hidden; border: 1px solid rgba(255,255,255,0.02);">
                        <div style="width: {health_factor}%; background: {health_color}; height: 100%; border-radius: 10px; box-shadow: 0 0 15px {health_color}; transition: width 1.5s cubic-bezier(0.16, 1, 0.3, 1);"></div>
                    </div>
                </div>
            """, unsafe_allow_html=True)

        with col2:
            st.markdown("""
                <div class="native-chart-box">
                    <div style="color: #94A3B8; font-size: 13px; text-transform: uppercase; letter-spacing: 2px; font-weight: 600; margin-bottom: 15px;">Top Structural Fatigue Profiles</div>
            """, unsafe_allow_html=True)
            
            top_degraded = df[df['Lifecycle Consumed (%)'] > 0].sort_values(by='Lifecycle Consumed (%)', ascending=False).head(5)
            
            for _, row in top_degraded.iterrows():
                val = min(row['Lifecycle Consumed (%)'] * 100, 100) # Cap at 100% for bar
                bar_color = "#F87171" if row['Status'] == 'OVERDUE' else ("#FB923C" if row['Status'] == 'HIGH PRIORITY' else "#38BDF8")
                
                st.markdown(f"""
                    <div style="margin-bottom: 12px;">
                        <div style="display: flex; justify-content: space-between; font-size: 12px; margin-bottom: 4px;">
                            <span style="color: #E2E8F0; font-weight: 500;">{row['Component Group']} <span style="color:#64748B;">({row['Location Unit']})</span></span>
                            <span style="color: {bar_color}; font-family: 'JetBrains Mono', monospace; font-weight: 700;">{row['Lifecycle Consumed (%)']*100:.1f}%</span>
                        </div>
                        <div style="width: 100%; background: rgba(255,255,255,0.05); border-radius: 4px; height: 6px; overflow: hidden;">
                            <div style="width: {val}%; background: {bar_color}; height: 100%; border-radius: 4px;"></div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)

        # --- SECTOR DATA MATRICES ---
        tab1, tab2, tab3 = st.tabs(["⚙️ Main Engine Flat Matrix", "⚡ Aux Generator Plant Matrix", "🛠️ Ext. Equipment Matrix"])
        
        ui_table_config = {
            "Subsystem": st.column_config.TextColumn("Subsystem Component"),
            "Component Group": st.column_config.TextColumn("Component Classification"),
            "Location Unit": st.column_config.TextColumn("Location Profile"),
            "Baseline Interval (Hrs)": st.column_config.NumberColumn("Interval Limit (Hrs)", format="%d"),
            "Current Running Hours": st.column_config.NumberColumn("Running Hours", format="%.1f"),
            "Lifecycle Consumed (%)": st.column_config.ProgressColumn("Fatigue Spectrum", format="%.1f%%", min_value=0.0, max_value=1.5),
            "Status": st.column_config.TextColumn("Diagnostic Condition State")
        }

        def color_row_states(val):
            if val == 'OVERDUE': return 'background-color: rgba(239, 68, 68, 0.15); color: #F87171; font-weight: 600;'
            elif val == 'HIGH PRIORITY': return 'background-color: rgba(251, 146, 60, 0.15); color: #FB923C; font-weight: 600;'
            return ''

        with tab1:
            me_display = df[df['Subsystem'] == 'MAIN ENGINE'].sort_values(by=['Component Group', 'Location Unit'])
            st.dataframe(me_display.style.map(color_row_states, subset=['Status']), use_container_width=True, hide_index=True, column_config=ui_table_config)

        with tab2:
            aux_display = df[df['Subsystem'] == 'AUX ENGINE'].sort_values(by=['Component Group', 'Location Unit'])
            st.dataframe(aux_display.style.map(color_row_states, subset=['Status']), use_container_width=True, hide_index=True, column_config=ui_table_config)

        with tab3:
            misc_display = df[df['Subsystem'] == 'OTHER EQUIPMENT'].sort_values(by=['Component Group', 'Location Unit'])
            st.dataframe(misc_display.style.map(color_row_states, subset=['Status']), use_container_width=True, hide_index=True, column_config=ui_table_config)
