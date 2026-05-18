import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta

# --- 1. CORPORATE LUXURY INTERFACE CANVAS (GLASSMORPHISM & ANIMATIONS) ---
st.set_page_config(page_title="Vessel Degradation Control Tower", page_icon="⚓", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;700&family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap');
    
    /* Core Canvas Typography Reset */
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
        background: radial-gradient(circle at 50% 50%, #0D1B2A 0%, #010811 100%);
        color: #E0E1DD;
    }
    
    /* Premium Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #0B132B;
        border-right: 1px solid #1B263B;
    }

    /* Keyframe Layer Movements */
    @keyframes cardEntrance {
        from { opacity: 0; transform: scale(0.97) translateY(20px); }
        to { opacity: 1; transform: scale(1) translateY(0); }
    }
    @keyframes criticalPulse {
        0% { box-shadow: 0 0 0 0 rgba(230, 57, 70, 0.4); border-color: #E63946; }
        70% { box-shadow: 0 0 0 15px rgba(230, 57, 70, 0); border-color: #E63946; }
        100% { box-shadow: 0 0 0 0 rgba(230, 57, 70, 0); border-color: #1B263B; }
    }

    /* Glassmorphism Metric Grid */
    .deck-container {
        display: flex;
        gap: 20px;
        margin-bottom: 30px;
    }
    .premium-card {
        flex: 1;
        background: rgba(27, 38, 59, 0.4);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px;
        padding: 24px;
        transition: all 0.4s cubic-bezier(0.16, 1, 0.3, 1);
        animation: cardEntrance 0.6s cubic-bezier(0.16, 1, 0.3, 1) both;
    }
    .premium-card:hover {
        transform: translateY(-5px);
        background: rgba(27, 38, 59, 0.6);
        border-color: #415A77;
        box-shadow: 0 20px 40px rgba(0, 0, 0, 0.5);
    }
    .card-pulse-critical {
        animation: cardEntrance 0.6s ease-out both, criticalPulse 2.5s infinite ease-in-out;
    }
    .meta-lbl {
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 2px;
        color: #8D99AE;
        font-weight: 600;
    }
    .meta-val {
        font-size: 38px;
        font-weight: 700;
        margin-top: 12px;
        color: #FFFFFF;
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    /* Code/Numeric Clean Renderings */
    .mono-txt {
        font-family: 'JetBrains Mono', monospace;
        font-weight: 700;
    }
    
    /* Customizing Streamlit Tab Navigation Bar */
    button[data-testid="stMarkdownContainer"] p {
        font-size: 15px !important;
        font-weight: 600 !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. BULLETPROOF PARSING ENGINE WITH DELTA SANITIZATION ---
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
        # Structural Delta Protection: Sanity-check single entries against maximum operational logical bounds
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

# --- 3. COMMAND CONSOLE MENU STRUCTURE ---
with st.sidebar:
    st.markdown("<h2 style='color:#5BC0BE; font-family:monospace;'>🎛️ PROPULSION OS</h2>", unsafe_allow_html=True)
    st.markdown("---")
    navigation = st.radio("Control View Menu", ["Vessel Overview", "Risk Control Hub", "System Calibration"])
    st.markdown("---")
    st.markdown("### Predictive Parameters")
    daily_vessel_runtime = st.slider("Assumed Daily Sea Runtime (Hrs)", 0, 24, 18)
    st.markdown("---")
    st.caption("Fleet Analytics Platform Engine | v3.4.1 Build 2026")

# --- 4. DATA PIPELINE INTERACTION AND RENDERING CONSOLE ---
if navigation == "Vessel Overview":
    st.markdown("<h1 style='color:#FFFFFF; margin-bottom: 0px; font-weight:700;'>🚢 Vessel Operational Command Tower</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#8D99AE;'>Ingest text telemetry blocks to isolate mechanical degradation vectors.</p>", unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("", type=["doc"])

    if uploaded_file is not None:
        if 'raw_df' not in st.session_state or st.session_state.get('uploaded_file_name') != uploaded_file.name:
            # First extraction and caching
            vessel_name, report_date, raw_dataframe = execute_stream_ingestion(uploaded_file.read())
            st.session_state.raw_df = raw_dataframe
            st.session_state.vessel_name = vessel_name
            st.session_state.report_date = report_date
            st.session_state.uploaded_file_name = uploaded_file.name

        # --- ADVANCED FEATURE: HUMAN-IN-THE-LOOP STAGING GATE ---
        st.markdown("### 🛠️ Interactive Data Verification Grid")
        st.caption("Verify and correct parsed values below before compiling the executive diagnostics dashboard.")
        
        # Deploy data editor for verification
        verified_df = st.data_editor(
            st.session_state.raw_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Subsystem": st.column_config.TextColumn("Subsystem Component", disabled=True),
                "Component Group": st.column_config.TextColumn("Equipment Group", disabled=True),
                "Location Unit": st.column_config.TextColumn("Location Unit", disabled=True),
                "Baseline Interval (Hrs)": st.column_config.NumberColumn("Baseline Limit", format="%d"),
                "Current Running Hours": st.column_config.NumberColumn("Extracted Hours (Editable)", format="%.1f")
            }
        )
        
        # Trigger Analytics Calculations on verified dataset
        if not verified_df.empty:
            df = verified_df.copy()
            df['Lifecycle Consumed (%)'] = np.where(df['Baseline Interval (Hrs)'] > 0, df['Current Running Hours'] / df['Baseline Interval (Hrs)'], 0.0)
            
            # Predictive failure logic mapping remaining service life to exact calendar deadlines
            df['Hours Remaining'] = np.maximum(0.0, df['Baseline Interval (Hrs)'] - df['Current Running Hours'])
            
            # Prevent dividing by 0 if runtime is sliding scale
            days_remaining = np.where(daily_vessel_runtime > 0, df['Hours Remaining'] / daily_vessel_runtime, 9999)
            
            # Vectorized timestamp scheduling mapping directly from system time
            current_date = datetime.now()
            df['Overhaul Deadline Date'] = [
                (current_date + timedelta(days=float(d))).strftime('%d %b %Y') if d < 5000 else "N/A"
                for d in days_remaining
            ]

            # Re-evaluating thresholds
            conditions = [(df['Current Running Hours'] == 0), (df['Lifecycle Consumed (%)'] >= 1.0), (df['Lifecycle Consumed (%)'] >= 0.8)]
            df['Status'] = np.select(conditions, ['NO DATA', 'OVERDUE', 'HIGH PRIORITY'], default='OK')

            # Aggregate Metric Calculations
            crit_df = df[df['Status'] == 'OVERDUE']
            warn_df = df[df['Status'] == 'HIGH PRIORITY']
            health_factor = max(0.0, 100.0 - ((len(crit_df) * 3.0 + len(warn_df) * 1.0) / len(df) * 100))

            st.markdown("---")
            st.markdown("### 📊 Fleet Diagnostics Summary Matrix")

            # HTML Neumorphic Metric Grid Layout Rendering
            st.markdown(f"""
                <div class="deck-container">
                    <div class="premium-card">
                        <div class="meta-lbl">Operational Context</div>
                        <div class="meta-val" style="color:#5BC0BE;">{st.session_state.vessel_name}</div>
                        <div style="color:#8D99AE; font-size:12px; margin-top:6px;">Report Reference: {st.session_state.report_date}</div>
                    </div>
                    <div class="premium-card {'card-pulse-critical' if len(crit_df)>0 else ''}">
                        <div class="meta-lbl">Critical Interrupts</div>
                        <div class="meta-val" style="color:#E63946;">{len(crit_df)} Items</div>
                        <div style="color:#8D99AE; font-size:12px; margin-top:6px;">Immediate Intervention Demanded</div>
                    </div>
                    <div class="premium-card">
                        <div class="meta-lbl">Pending Risks</div>
                        <div class="meta-val" style="color:#F4A261;">{len(warn_df)} Items</div>
                        <div style="color:#8D99AE; font-size:12px; margin-top:6px;">Approaching Life Cycle Threshold</div>
                    </div>
                    <div class="premium-card">
                        <div class="meta-lbl">Calculated Fleet Health</div>
                        <div class="meta-val" style="color:#2A9D8F;">{health_factor:.1f}%</div>
                        <div style="color:#8D99AE; font-size:12px; margin-top:6px;">Total Structural Integrity Score</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)

            # --- CUSTOM EQUIPMENT SORTING VIEWS ---
            tab1, tab2, tab3 = st.tabs(["🔥 Critical Interventions Log", "🔩 Main Engine Analysis", "⚡ Auxiliary Plant Log"])
            
            ui_config = {
                "Subsystem": st.column_config.TextColumn("Subsystem"),
                "Component Group": st.column_config.TextColumn("Equipment Component"),
                "Location Unit": st.column_config.TextColumn("Location"),
                "Baseline Interval (Hrs)": st.column_config.NumberColumn("Interval (Hrs)", format="%d"),
                "Current Running Hours": st.column_config.NumberColumn("Current Hours", format="%.1f"),
                "Lifecycle Consumed (%)": st.column_config.ProgressColumn("Fatigue Spectrum", format="%.1f%%", min_value=0.0, max_value=1.5),
                "Overhaul Deadline Date": st.column_config.TextColumn("Predicted Deadline Target"),
                "Status": st.column_config.TextColumn("Status")
            }

            def highlight_matrix_rows(val):
                if val == 'OVERDUE': return 'background-color: rgba(230, 57, 70, 0.25); color: #E63946; font-weight: bold;'
                elif val == 'HIGH PRIORITY': return 'background-color: rgba(244, 162, 97, 0.25); color: #F4A261; font-weight: bold;'
                return ''

            with tab1:
                st.subheader("Isolated Strategic Failure Elements")
                risk_df = df[df['Status'].isin(['OVERDUE', 'HIGH PRIORITY'])].sort_values(by='Lifecycle Consumed (%)', ascending=False)
                if not risk_df.empty:
                    st.dataframe(risk_df.style.map(highlight_matrix_rows, subset=['Status']), use_container_width=True, hide_index=True, column_config=ui_config)
                else:
                    st.success("No system exceptions or structural life runouts flagged.")

            with tab2:
                st.subheader("Main Propulsion System Hierarchy Log")
                # Grouping systematically matching structural layout sequence
                me_display = df[df['Subsystem'] == 'MAIN PROPULSION'].sort_values(by=['Component Group', 'Location Unit'])
                st.dataframe(me_display.style.map(highlight_matrix_rows, subset=['Status']), use_container_width=True, hide_index=True, column_config=ui_config)

            with tab3:
                st.subheader("Auxiliary Plant Allocation Matrix")
                aux_display = df[df['Subsystem'].str.contains('AUX')].sort_values(by=['Subsystem', 'Component Group', 'Location Unit'])
                st.dataframe(aux_display.style.map(highlight_matrix_rows, subset=['Status']), use_container_width=True, hide_index=True, column_config=ui_config)

elif navigation == "Risk Control Hub":
    st.markdown("<h1>🎯 Fleet Predictive Risk Matrix</h1>", unsafe_allow_html=True)
    st.info("The dynamic Risk Matrix module tracks lifecycle failure distributions across active operational sectors.")
else:
    st.markdown("<h1>⚙️ System Calibration Panel</h1>", unsafe_allow_html=True)
    st.info("Modify regex string identifiers and baseline calculation configurations.")
