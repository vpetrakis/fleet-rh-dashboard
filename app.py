import streamlit as st
import pandas as pd
import numpy as np
import re
import io

# --- 1. CONFIGURATION & 10/10 AESTHETICS ---
st.set_page_config(page_title="Fleet RH Dashboard", page_icon="🚢", layout="wide", initial_sidebar_state="collapsed")

# Injecting Custom CSS for Premium Executive Look
st.markdown("""
    <style>
    /* Metric Cards */
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        padding: 5% 5% 5% 10%;
        border-radius: 8px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        border-left: 5px solid #1f77b4;
    }
    
    /* Status Badges */
    .badge-overdue { background-color: #ffcccc; color: #cc0000; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 14px; }
    .badge-warning { background-color: #fff2cc; color: #b38600; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 14px; }
    .badge-ok { background-color: #cce5ff; color: #004085; padding: 4px 8px; border-radius: 4px; font-weight: 600; font-size: 14px; }
    
    /* Headers */
    h1, h2, h3 { color: #2c3e50; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    </style>
""", unsafe_allow_html=True)

THRESHOLD_RED = 1.0
THRESHOLD_YELLOW = 0.8

# --- 2. THE BULLETPROOF BACKEND ---
def extract_numeric_value(text: str) -> float:
    """Ultra-resilient numeric extractor. Strips out all text, brackets, and fixes regional commas."""
    if not isinstance(text, str):
        return 0.0
    
    cleaned = text.strip().replace('[', '').replace(']', '').replace(' ', '')
    if any(ignore in cleaned.upper() for ignore in ["N/A", "CENTRAL", "OBSERVATION", "-", "COOLER"]):
        return 0.0
        
    # Regex to capture only numeric patterns and decimal/comma separators
    match = re.search(r'[\d\.\,]+', cleaned)
    if not match:
        return 0.0
        
    num_str = match.group(0)
    
    # Handle European vs US formats safely
    if ',' in num_str and '.' in num_str:
        num_str = num_str.replace(',', '')  # 1,234.56 -> 1234.56
    elif num_str.count(',') == 1:
        parts = num_str.split(',')
        if len(parts[1]) != 3:  # If not exactly 3 digits after comma, it's a decimal (e.g., 100,5)
            num_str = num_str.replace(',', '.')
        else:
            num_str = num_str.replace(',', '') # It's a thousands separator (e.g., 10,000)

    try:
        return float(num_str)
    except ValueError:
        return 0.0

def parse_legacy_doc_stream(raw_bytes) -> tuple:
    """Decodes binary .doc files and hunts for \x07 cell delimiters."""
    # Decode ignoring binary garbage
    text = raw_bytes.decode('utf-8', errors='ignore')
    
    # Structural Metadata Extraction
    vessel = "UNKNOWN VESSEL"
    date_str = "UNKNOWN DATE"
    
    v_match = re.search(r"Vessel’s\s*Name:\s*([^\t\x07\r\n]+)", text, re.IGNORECASE)
    if v_match: vessel = v_match.group(1).strip()
        
    d_match = re.search(r"Date:\s*([\d\s\w]+)", text, re.IGNORECASE)
    if d_match: date_str = d_match.group(1).strip()

    parsed_records = []
    
    # 1. Main Engine Logic (Vertical Pair Reading)
    me_patterns = [
        ("CYLINDER COVER", 16000), ("PISTON ASSEMBLY", 16000), ("STUFFING BOX", 16000), 
        ("PISTON CROWN", 32000), ("CYLINDER LINER", 0), ("EXHAUST VALVE", 16000), 
        ("STARTING VALVE", 12000), ("SAFETY VALVE", 12000), ("FUEL VALVES", 8000), 
        ("FUEL PUMP", 16000), ("CROSSHEAD BEARINGS", 32000), ("BOTTOM END BEARINGS", 32000), ("MAIN BEARINGS", 32000)
    ]
    
    for comp, periodicity in me_patterns:
        # Look for the component name, followed by any junk, then the '2' marker, then the hours row
        comp_block = re.search(rf"{re.escape(comp)}\x07.*?\x072\x07([^\x0d]+)", text, re.DOTALL | re.IGNORECASE)
        if comp_block:
            hours_line = comp_block.group(1)
            hours_tokens = [t.strip() for t in hours_line.split('\x07') if t.strip()]
            
            for idx, h_val in enumerate(hours_tokens[:7]):  # Max 7 cylinders for Main Engine
                hrs = extract_numeric_value(h_val)
                parsed_records.append({
                    "Component": comp, "System": "MAIN ENGINE", "Unit": f"Cyl No.{idx+1}",
                    "Periodicity": periodicity, "Running Hours": hrs
                })

    # 2. Auxiliary Engines Logic (Horizontal Reading across 3 Engines)
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
            
            for i in range(1, 4): # D/G 1, 2, 3
                start_idx = (i - 1) * 6
                for cyl in range(6): # 6 Cylinders per Aux Engine
                    token_idx = start_idx + cyl
                    if token_idx < len(hours_tokens):
                        hrs = extract_numeric_value(hours_tokens[token_idx])
                        parsed_records.append({
                            "Component": comp.replace(" (1)", ""), "System": f"AUX ENGINE No.{i}", "Unit": f"Cyl No.{cyl+1}",
                            "Periodicity": periodicity, "Running Hours": hrs
                        })

    # 3. Create DataFrame and Calculate Status
    df = pd.DataFrame(parsed_records)
    if not df.empty:
        # Avoid division by zero
        df['% Used'] = np.where(df['Periodicity'] > 0, df['Running Hours'] / df['Periodicity'], 0.0)
        
        conditions = [
            (df['Running Hours'] == 0) | (df['Periodicity'] == 0),
            (df['% Used'] >= THRESHOLD_RED),
            (df['% Used'] >= THRESHOLD_YELLOW)
        ]
        choices = ['NO DATA', 'OVERDUE', 'HIGH PRIORITY']
        df['Status'] = np.select(conditions, choices, default='OK')
        
    return vessel, date_str, df

def convert_df_to_csv(df):
    """Converts dataframe to a CSV for downloading."""
    return df.to_csv(index=False).encode('utf-8')

# --- 3. FRONTEND DASHBOARD ---
st.title("🚢 Fleet Running Hours Intelligence")
st.write("Upload a legacy `.doc` report to instantly extract, validate, and visualize fleet diagnostics.")

uploaded_file = st.file_uploader("Upload Vessel Report (.doc)", type=["doc"])

if uploaded_file is not None:
    with st.spinner('Parsing binary document structures...'):
        file_bytes = uploaded_file.read()
        vessel_name, report_date, data_df = parse_legacy_doc_stream(file_bytes)
        
    if data_df.empty:
        st.error("🚨 Extraction Failed: Could not locate table structures in this document. Ensure the file follows the standard format.")
    else:
        st.success(f"Successfully extracted {len(data_df)} records from {vessel_name}.")
        
        # --- TOP ROW: KPI METRICS ---
        st.markdown(f"### {vessel_name} | <span style='font-size:18px; color:gray;'>Report Date: {report_date}</span>", unsafe_allow_html=True)
        
        overdue_df = data_df[data_df['Status'] == 'OVERDUE']
        warning_df = data_df[data_df['Status'] == 'HIGH PRIORITY']
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Components Monitored", f"{len(data_df)}")
        c2.metric("Critical Overdue Items", f"{len(overdue_df)}", delta="Action Required", delta_color="inverse")
        c3.metric("High Priority Warnings", f"{len(warning_df)}", delta="Approaching Limit", delta_color="off")
        
        # Health Score Calculation
        health_score = 100 - ((len(overdue_df) * 2 + len(warning_df)) / len(data_df) * 100)
        c4.metric("Vessel Health Score", f"{health_score:.1f}%")

        st.divider()

        # --- TABBED NAVIGATION ---
        tab1, tab2 = st.tabs(["⚠️ Action Matrix", "📋 Full Diagnostic Log"])
        
        # UI Configuration for the Dataframe columns
        column_config = {
            "Component": st.column_config.TextColumn("Engine Component", width="medium"),
            "System": st.column_config.TextColumn("System", width="medium"),
            "Unit": st.column_config.TextColumn("Unit", width="small"),
            "Periodicity": st.column_config.NumberColumn("Periodicity (Hrs)", format="%d"),
            "Running Hours": st.column_config.NumberColumn("Running Hours", format="%.1f"),
            "% Used": st.column_config.ProgressColumn(
                "Lifecycle Used",
                help="Percentage of component lifecycle consumed",
                format="%.1f%%",
                min_value=0,
                max_value=1,
            ),
            "Status": st.column_config.TextColumn("Status")
        }
        
        def highlight_status(val):
            """Applies color to the dataframe based on status."""
            if val == 'OVERDUE': return 'background-color: #ffcccc; color: #cc0000; font-weight: bold'
            elif val == 'HIGH PRIORITY': return 'background-color: #fff2cc; color: #b38600; font-weight: bold'
            elif val == 'OK': return 'color: #004085;'
            return 'color: #6c757d;'

        with tab1:
            st.subheader("High Priority Interventions")
            action_items = pd.concat([overdue_df, warning_df]).sort_values(by='% Used', ascending=False)
            
            if not action_items.empty:
                styled_action = action_items.style.map(highlight_status, subset=['Status'])
                st.dataframe(styled_action, use_container_width=True, hide_index=True, column_config=column_config)
            else:
                st.info("✅ All systems are operating within safe lifecycle thresholds. No immediate action required.")
                
        with tab2:
            st.subheader("Complete Equipment Log")
            
            # Add Download Button for Local Backup
            csv_data = convert_df_to_csv(data_df)
            st.download_button(
                label="📥 Download Data as CSV",
                data=csv_data,
                file_name=f"{vessel_name.replace(' ', '_')}_RH_Report.csv",
                mime='text/csv',
            )
            
            styled_full = data_df.style.map(highlight_status, subset=['Status'])
            st.dataframe(styled_full, use_container_width=True, hide_index=True, column_config=column_config)
