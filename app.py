import streamlit as st
import pandas as pd
import numpy as np
import re
import plotly.graph_objects as go
import plotly.express as px

# --- 1. THE PINNACLE OF UI ARCHITECTURE: STEALTH ANIMATIONS & GLASSMORPHISM ---
st.set_page_config(page_title="Propulsion Command OS", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;700&display=swap');
    
    /* Global Canvas Reset - Ultra-Dark Space Theme */
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Inter', sans-serif;
        background: radial-gradient(circle at 50% 0%, #0B1121 0%, #02040A 100%);
        color: #E2E8F0;
    }
    
    [data-testid="stHeader"], footer {visibility: hidden;}
    
    /* --- ADVANCED KEYFRAME ANIMATIONS --- */
    @keyframes staggeredFadeUp {
        0% { opacity: 0; transform: translateY(30px) scale(0.98); }
        100% { opacity: 1; transform: translateY(0) scale(1); }
    }
    
    @keyframes holographicPulse {
        0% { box-shadow: 0 0 0 0 rgba(56, 189, 248, 0.4); border-color: rgba(56, 189, 248, 0.4); }
        50% { box-shadow: 0 0 20px 0 rgba(56, 189, 248, 0.1); border-color: rgba(56, 189, 248, 0.8); }
        100% { box-shadow: 0 0 0 0 rgba(56, 189, 248, 0.4); border-color: rgba(56, 189, 248, 0.4); }
    }

    @keyframes criticalRadar {
        0% { box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.3); border-color: rgba(239, 68, 68, 0.5); }
        50% { box-shadow: 0 0 30px 5px rgba(239, 68, 68, 0.2); border-color: rgba(239, 68, 68, 1); }
        100% { box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.3); border-color: rgba(239, 68, 68, 0.5); }
    }

    @keyframes scanline {
        0% { transform: translateY(-100%); }
        100% { transform: translateY(100%); }
    }

    /* --- GLASSMORPHIC KPI CARDS WITH STAGGERED ENTRANCE --- */
    .dashboard-deck {
        display: flex;
        gap: 20px;
        margin-bottom: 30px;
    }
    
    .kpi-card {
        flex: 1;
        position: relative;
        overflow: hidden;
        background: linear-gradient(145deg, rgba(30, 41, 59, 0.4) 0%, rgba(15, 23, 42, 0.6) 100%);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.05);
        border-radius: 16px;
        padding: 24px;
        transition: all 0.4s cubic-bezier(0.16, 1, 0.3, 1);
        animation: staggeredFadeUp 0.7s cubic-bezier(0.16, 1, 0.3, 1) both;
    }
    
    .kpi-card:hover {
        transform: translateY(-8px);
        background: linear-gradient(145deg, rgba(30, 41, 59, 0.6) 0%, rgba(15, 23, 42, 0.8) 100%);
        border-color: rgba(56, 189, 248, 0.3);
        box-shadow: 0 20px 40px rgba(0, 0, 0, 0.6);
    }
    
    /* Staggering the animation delays for a cascading load effect */
    .dashboard-deck > div:nth-child(1) { animation-delay: 0.1s; }
    .dashboard-deck > div:nth-child(2) { animation-delay: 0.2s; }
    .dashboard-deck > div:nth-child(3) { animation-delay: 0.3s; }
    .dashboard-deck > div:nth-child(4) { animation-delay: 0.4s; }

    .card-critical { animation: staggeredFadeUp 0.7s ease-out 0.2s both, criticalRadar 2.5s infinite ease-in-out; }
    .card-active { animation: staggeredFadeUp 0.7s ease-out 0.1s both, holographicPulse 3s infinite ease-in-out; }

    .kpi-title { font-size: 11px; text-transform: uppercase; letter-spacing: 2px; color: #94A3B8; font-weight: 600; }
    .kpi-value { font-size: 38px; font-weight: 700; margin-top: 8px; color: #FFFFFF; font-family: 'JetBrains Mono', monospace; letter-spacing: -1px; }
    
    /* --- CUSTOM UPLOAD ZONE --- */
    div[data-testid="stFileUploadDropzone"] {
        background: rgba(15, 23, 42, 0.3) !important;
        border: 2px dashed rgba(56, 189, 248, 0.2) !important;
        border-radius: 16px !important;
        padding: 40px !important;
        transition: all 0.4s cubic-bezier(0.16, 1, 0.3, 1);
        animation: staggeredFadeUp 0.8s ease-out both;
    }
    div[data-testid="stFileUploadDropzone"]:hover {
        border-color: #38BDF8 !important;
        background: rgba(15, 23, 42, 0.6) !important;
        box-shadow: 0 0 30px rgba(56, 189, 248, 0.15) !important;
    }

    /* --- DATAFRAME & TABS STYLING --- */
    div[data-testid="stDataFrame"] {
        border: 1px solid rgba(255, 255, 255, 0.05) !important;
        border-radius: 12px !important;
        background: #090D1A !important;
        animation: staggeredFadeUp 0.9s ease-out 0.5s both;
    }
    button[data-testid="stMarkdownContainer"] p {
        font-size: 15px !important; font-weight: 600 !important; color: #E2E8F0;
    }
    
    .chart-container {
        animation: staggeredFadeUp 0.8s ease-out 0.4s both;
        background: rgba(15, 23, 42, 0.3);
        border: 1px solid rgba(255, 255, 255, 0.03);
        border-radius: 16px;
        padding: 15px;
        margin-bottom: 30px;
    }
    </style>
""", unsafe_allow_html=True)

THRESHOLD_RED = 1.0
THRESHOLD_YELLOW = 0.8

# --- 2. FORENSIC TELEMETRY ENGINE (100% INTEGRITY PRESERVED) ---
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

# --- 3. FRONTEND UI & ANIMATED VISUALIZATION MODULES ---
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
                <div class="kpi-card card-active">
                    <div class="kpi-title">Active Target Profile</div>
                    <div class="kpi-value" style="color:#38BDF8;">{vessel_name}</div>
                    <div style="color:#475569; font-size:12px; margin-top:8px; font-weight:600;">Log Reference: {report_date}</div>
                </div>
                <div class="kpi-card {'card-critical' if len(crit_df)>0 else ''}">
                    <div class="kpi-title">Critical Overhauls</div>
                    <div class="kpi-value" style="color:#F87171;">{len(crit_df)}<span style="font-size:16px; color:#64748B;"> Items</span></div>
                    <div style="color:#475569; font-size:12px; margin-top:8px; font-weight:600;">Immediate Action Mandatory</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-title">High Priority Risks</div>
                    <div class="kpi-value" style="color:#FB923C;">{len(warn_df)}<span style="font-size:16px; color:#64748B;"> Items</span></div>
                    <div style="color:#475569; font-size:12px; margin-top:8px; font-weight:600;">Approaching Absolute Threshold</div>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # --- INTERACTIVE 3D/CANVAS PLOTLY VISUALIZATIONS ---
        col_gauge, col_bar = st.columns([1, 2])
        
        with col_gauge:
            st.markdown("<div class='chart-container'>", unsafe_allow_html=True)
            # High-end Animated Gauge Chart
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number",
                value=health_factor,
                domain={'x': [0, 1], 'y': [0, 1]},
                title={'text': "Aggregate Health Index", 'font': {'size': 18, 'color': '#94A3B8'}},
                number={'suffix': "%", 'font': {'size': 40, 'color': '#FFFFFF'}},
                gauge={
                    'axis': {'range': [None, 100], 'tickwidth': 1, 'tickcolor': "darkblue"},
                    'bar': {'color': "#34D399" if health_factor > 80 else ("#FB923C" if health_factor > 50 else "#F87171")},
                    'bgcolor': "rgba(0,0,0,0)",
                    'borderwidth': 2,
                    'bordercolor': "rgba(255,255,255,0.1)",
                    'steps': [
                        {'range': [0, 50], 'color': 'rgba(239, 68, 68, 0.2)'},
                        {'range': [50, 80], 'color': 'rgba(251, 146, 60, 0.2)'},
                        {'range': [80, 100], 'color': 'rgba(52, 211, 153, 0.2)'}],
                }
            ))
            fig_gauge.update_layout(paper_bgcolor="rgba(0,0,0,0)", font={'color': "#F8FAFC", 'family': "Inter"}, height=280, margin=dict(l=20, r=20, t=50, b=20))
            st.plotly_chart(fig_gauge, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with col_bar:
            st.markdown("<div class='chart-container'>", unsafe_allow_html=True)
            # Advanced Degradation Bar Chart
            top_degraded = df[df['Lifecycle Consumed (%)'] > 0].sort_values(by='Lifecycle Consumed (%)', ascending=False).head(8)
            
            # Map colors dynamically based on status
            color_discrete_map = {'OVERDUE': '#F87171', 'HIGH PRIORITY': '#FB923C', 'OK': '#38BDF8'}
            
            fig_bar = px.bar(
                top_degraded, 
                x='Lifecycle Consumed (%)', 
                y='Component Group', 
                color='Status',
                orientation='h',
                color_discrete_map=color_discrete_map,
                title="Top Structural Fatigue Profiles",
                hover_data=['Location Unit', 'Current Running Hours']
            )
            fig_bar.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", 
                font={'color': "#94A3B8", 'family': "Inter"},
                height=280, margin=dict(l=10, r=20, t=40, b=20),
                xaxis=dict(tickformat=".0%", gridcolor="rgba(255,255,255,0.05)"),
                yaxis=dict(autorange="reversed")
            )
            st.plotly_chart(fig_bar, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

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
