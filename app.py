import streamlit as st
import pandas as pd
import numpy as np
import re

# --- 1. THE ARCHITECTURAL FRONTEND MASTERPIECE (CUSTOM CSS & ANIMATIONS) ---
st.set_page_config(page_title="Marine Operations Control", page_icon="⚓", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Global Canvas Reset */
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Inter', sans-serif;
        background-color: #0B132B; /* Deep Navy Maritime Background */
        color: #F4F6F9;
    }
    
    /* Sidebar Styling Override */
    [data-testid="stSidebar"] {
        background-color: #1C2541;
        border-right: 1px solid #3A506B;
    }

    /* CSS Keyframe Animations */
    @keyframes slideUpFade {
        from { opacity: 0; transform: translateY(30px); }
        to { opacity: 1; transform: translateY(0); }
    }
    @keyframes pulseCritical {
        0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(255, 75, 75, 0.7); }
        70% { transform: scale(1); box-shadow: 0 0 0 12px rgba(255, 75, 75, 0); }
        100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(255, 75, 75, 0); }
    }
    @keyframes glowPulse {
        0% { border-color: #3A506B; }
        50% { border-color: #5BC0BE; }
        100% { border-color: #3A506B; }
    }

    /* Executive Glassmorphic KPI Cards */
    .kpi-deck {
        display: flex;
        gap: 20px;
        margin-bottom: 25px;
        animation: slideUpFade 0.5s cubic-bezier(0.16, 1, 0.3, 1) both;
    }
    .kpi-card {
        flex: 1;
        background: linear-gradient(145deg, #1C2541, #131A33);
        border: 1px solid #3A506B;
        border-radius: 14px;
        padding: 24px;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
        transition: all 0.4s cubic-bezier(0.16, 1, 0.3, 1);
    }
    .kpi-card:hover {
        transform: translateY(-6px);
        border-color: #5BC0BE;
        box-shadow: 0 12px 40px 0 rgba(91, 192, 190, 0.2);
    }
    .kpi-label {
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 1.5px;
        color: #9A9EAB;
        font-weight: 600;
    }
    .kpi-value {
        font-size: 36px;
        font-weight: 700;
        margin-top: 10px;
        color: #FFFFFF;
    }
    
    /* Live Pulsing Indicator Dots */
    .indicator-dot {
        width: 12px;
        height: 12px;
        border-radius: 50%;
        display: inline-block;
        margin-right: 8px;
    }
    .dot-critical { background-color: #FF4B4B; animation: pulseCritical 2s infinite; }
    .dot-warning { background-color: #FFAA00; }
    .dot-nominal { background-color: #00E676; }

    /* Custom File Upload Drag & Drop Interface Wrapper */
    div[data-testid="stFileUploadDropzone"] {
        background-color: #1C2541 !important;
        border: 2px dashed #3A506B !important;
        border-radius: 12px !important;
        padding: 30px !important;
        animation: glowPulse 4s infinite ease-in-out;
        transition: all 0.3s ease;
    }
    div[data-testid="stFileUploadDropzone"]:hover {
        border-color: #5BC0BE !important;
        background-color: #222C4E !important;
    }

    /* Tabs Customization */
    button[data-testid="stMarkdownContainer"] {
        color: #FFFFFF !important;
    }
    </style>
""", unsafe_allow_html=True)

THRESHOLD_RED = 1.0
THRESHOLD_YELLOW = 0.8

# --- 2. FORTRAN BACKEND DATA PIPELINE ---
def extract_numeric_value(text: str) -> float:
    if not isinstance(text, str):
        return 0.0
    cleaned = text.strip().replace('[', '').replace(']', '').replace(' ', '')
    if any(ignore in cleaned.upper() for ignore in ["N/A", "CENTRAL", "OBSERVATION", "-", "COOLER"]):
        return 0.0
    match = re.search(r'[\d\.\,]+', cleaned)
    if not match:
        return 0.0
    num_str = match.group(0)
    if ',' in num_str and '.' in num_str:
        num_str = num_str.replace(',', '')
    elif num_str.count(',') == 1:
        parts = num_str.split(',')
        if len(parts[1]) != 3:
            num_str = num_str.replace(',', '.')
        else:
            num_str = num_str.replace(',', '')
    try:
        return float(num_str)
    except ValueError:
        return 0.0

def parse_legacy_doc_stream(raw_bytes) -> tuple:
    text = raw_bytes.decode('utf-8', errors='ignore')
    vessel = "UNKNOWN VESSEL"
    date_str = "UNKNOWN DATE"
    
    v_match = re.search(r"Vessel’s\s*Name:\s*([^\t\x07\r\n]+)", text, re.IGNORECASE)
    if v_match: vessel = v_match.group(1).strip()
    d_match = re.search(r"Date:\s*([\d\s\w]+)", text, re.IGNORECASE)
    if d_match: date_str = d_match.group(1).strip()

    parsed_records = []
    
    me_patterns = [
        ("CYLINDER COVER", 16000), ("PISTON ASSEMBLY", 16000), ("STUFFING BOX", 16000), 
        ("PISTON CROWN", 32000), ("CYLINDER LINER", 0), ("EXHAUST VALVE", 16000), 
        ("STARTING VALVE", 12000), ("SAFETY VALVE", 12000), ("FUEL VALVES", 8000), 
        ("FUEL PUMP", 16000), ("CROSSHEAD BEARINGS", 32000), ("BOTTOM END BEARINGS", 32000), ("MAIN BEARINGS", 32000)
    ]
    
    for comp, periodicity in me_patterns:
        comp_block = re.search(rf"{re.escape(comp)}\x07.*?\x072\x07([^\x0d]+)", text, re.DOTALL | re.IGNORECASE)
        if comp_block:
            hours_line = comp_block.group(1)
            hours_tokens = [t.strip() for t in hours_line.split('\x07') if t.strip()]
            for idx, h_val in enumerate(hours_tokens[:7]):
                hrs = extract_numeric_value(h_val)
                parsed_records.append({
                    "Component": comp, "System": "MAIN ENGINE", "Unit": f"Cyl No.{idx+1}",
                    "Periodicity": periodicity, "Running Hours": hrs
                })

    aux_patterns = [
        ("Cylinder Head", 12000), ("Piston", 10000), ("Connecting Rod", 10000), 
        ("Cylinder Liners", 10000), ("Fuel Valves (1)", 2000), ("Fuel Pumps", 5000),
        ("Crank Pin Bearing", 12000), ("Main Bearing", 12000)
    ]
    
    for comp, periodicity in aux_patterns:
        aux_block = re.search(rf"{re.escape(comp)}\x07.*?\x072\x07([^\x0d]+)", text, re.DOTALL | re.IGNORECASE)
        if aux_block:
            hours_line = aux_block.group(1)
            hours_tokens = [t.strip() for t in hours_line.split('\x07') if t.strip()]
            for i in range(1, 4):
                start_idx = (i - 1) * 6
                for cyl in range(6):
                    token_idx = start_idx + cyl
                    if token_idx < len(hours_tokens):
                        hrs = extract_numeric_value(hours_tokens[token_idx])
                        parsed_records.append({
                            "Component": comp.replace(" (1)", ""), "System": f"AUX ENGINE No.{i}", "Unit": f"Cyl No.{cyl+1}",
                            "Periodicity": periodicity, "Running Hours": hrs
                        })

    df = pd.DataFrame(parsed_records)
    if not df.empty:
        df['% Used'] = np.where(df['Periodicity'] > 0, df['Running Hours'] / df['Periodicity'], 0.0)
        conditions = [
            (df['Running Hours'] == 0) | (df['Periodicity'] == 0),
            (df['% Used'] >= THRESHOLD_RED),
            (df['% Used'] >= THRESHOLD_YELLOW)
        ]
        choices = ['NO DATA', 'OVERDUE', 'HIGH PRIORITY']
        df['Status'] = np.select(conditions, choices, default='OK')
        
    return vessel, date_str, df

# --- 3. SIDEBAR NAVIGATION CONSOLE ---
with st.sidebar:
    st.markdown("<h2 style='color:#5BC0BE;'>🛰️ Fleet Control</h2>", unsafe_allow_html=True)
    st.markdown("---")
    selected_view = st.radio("Navigation Window", ["Operations Dashboard", "System Live Logs", "Fleet Risk Map"])
    st.markdown("---")
    st.markdown("### Operational Settings")
    dynamic_red = st.slider("Overdue Limit (%)", 80, 120, 100) / 100.0

# --- 4. CONTROL TOWER INTERFACE ---
if selected_view == "Operations Dashboard":
    st.markdown("<h1 style='color:#FFFFFF; margin-bottom: 0px;'>⚓ Fleet Running Hours Intelligence</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#9A9EAB;'>Drag and drop raw text telemetry logs to evaluate degradation profiles.</p>", unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("", type=["doc"])

    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        vessel_name, report_date, data_df = parse_legacy_doc_stream(file_bytes)
        
        if not data_df.empty:
            # Re-evaluate limits dynamically based on sidebar slider inputs
            data_df['% Used'] = np.where(data_df['Periodicity'] > 0, data_df['Running Hours'] / data_df['Periodicity'], 0.0)
            conditions = [
                (data_df['Running Hours'] == 0),
                (data_df['% Used'] >= dynamic_red),
                (data_df['% Used'] >= THRESHOLD_YELLOW)
            ]
            data_df['Status'] = np.select(conditions, ['NO DATA', 'OVERDUE', 'HIGH PRIORITY'], default='OK')

            overdue_count = len(data_df[data_df['Status'] == 'OVERDUE'])
            warning_count = len(data_df[data_df['Status'] == 'HIGH PRIORITY'])
            health_score = max(0.0, 100.0 - ((overdue_count * 2.5 + warning_count * 1.0) / len(data_df) * 100))

            # HTML Injection for High-End Animated Metric Deck
            st.markdown(f"""
                <div class="kpi-deck">
                    <div class="kpi-card">
                        <div class="kpi-label">Target Asset Context</div>
                        <div class="kpi-value" style="color:#5BC0BE;">{vessel_name}</div>
                        <div style="color:#9A9EAB; font-size:12px; margin-top:5px;">Log Reference: {report_date}</div>
                    </div>
                    <div class="kpi-card">
                        <div class="kpi-label"><span class="indicator-dot dot-critical"></span>Critical Interrupts</div>
                        <div class="kpi-value" style="color:#FF4B4B;">{overdue_count} Items</div>
                        <div style="color:#9A9EAB; font-size:12px; margin-top:5px;">Immediate Overhaul Action Required</div>
                    </div>
                    <div class="kpi-card">
                        <div class="kpi-label"><span class="indicator-dot dot-warning"></span>Impending Interventions</div>
                        <div class="kpi-value" style="color:#FFAA00;">{warning_count} Items</div>
                        <div style="color:#9A9EAB; font-size:12px; margin-top:5px;">Approaching Operational Limit</div>
                    </div>
                    <div class="kpi-card">
                        <div class="kpi-label">Hull Efficiency Status</div>
                        <div class="kpi-value" style="color:#00E676;">{health_score:.1f}%</div>
                        <div style="color:#9A9EAB; font-size:12px; margin-top:5px;">Calculated Structural Integrity</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)

            # Executive Interactive Filters Matrix
            tab1, tab2, tab3 = st.tabs(["🔥 Risk Exceptions Matrix", "🔩 Main Propagation Matrix", "⚡ Auxiliary Propulsion Matrix"])
            
            # Interactive Progress Column Framework Layout Config
            ui_column_config = {
                "Component": st.column_config.TextColumn("Component Identifier", width="medium"),
                "System": st.column_config.TextColumn("Machinery Subsystem", width="medium"),
                "Unit": st.column_config.TextColumn("Location", width="small"),
                "Periodicity": st.column_config.NumberColumn("Limit Interval (Hrs)", format="%d"),
                "Running Hours": st.column_config.NumberColumn("Accumulated Running Hours", format="%.1f"),
                "% Used": st.column_config.ProgressColumn("Structural Fatigue Curve", format="%.1f%%", min_value=0, max_value=1.5),
                "Status": st.column_config.TextColumn("Diagnostic State")
            }

            with tab1:
                # Filter strictly for action elements sorted by highest degradation curves
                risk_df = data_df[data_df['Status'].isin(['OVERDUE', 'HIGH PRIORITY'])].sort_values(by='% Used', ascending=False)
                if not risk_df.empty:
                    st.dataframe(risk_df, use_container_width=True, hide_index=True, column_config=ui_column_config)
                else:
                    st.success("All systems operating within normal degradation profiles.")

            with tab2:
                # Hierarchical Filtering Strategy: System -> Component -> Unit
                me_df = data_df[data_df['System'] == 'MAIN ENGINE'].sort_values(by=['Component', 'Unit'])
                st.dataframe(me_df, use_container_width=True, hide_index=True, column_config=ui_column_config)

            with tab3:
                aux_df = data_df[data_df['System'].str.contains('AUX')].sort_values(by=['System', 'Component', 'Unit'])
                st.dataframe(aux_df, use_container_width=True, hide_index=True, column_config=ui_column_config)
else:
    st.info("Additional console module selected. Main parsing metrics are contained on the Operations Dashboard.")
