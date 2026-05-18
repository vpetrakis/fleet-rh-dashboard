import streamlit as st
import pandas as pd
# from backend.parser import extract_vessel_data 

# --- CONFIGURATION ---
st.set_page_config(page_title="Fleet RH Monitor", layout="wide")

# Styling to mimic your "Premium Layout"
st.markdown("""
    <style>
    .metric-card { background-color: #F0F0F0; border: 1px solid #C8C8C8; padding: 15px; border-radius: 5px; text-align: center;}
    .overdue { color: #FF0000; font-weight: bold; font-size: 24px;}
    .warning { color: #FFC000; font-weight: bold; font-size: 24px;}
    </style>
""", unsafe_allow_html=True)

# --- UI LOGIC ---
st.title("🚢 Running Hours Monitoring System")

# 1. File Ingestion
uploaded_file = st.file_uploader("Upload Vessel Report (.docx)", type=["docx"])

if uploaded_file is not None:
    # 2. Process Data (Mocking the dataframe for demonstration)
    with st.spinner("Analyzing document metrics..."):
        # df = extract_vessel_data(uploaded_file)
        
        # Mock Data representing the output of the parser
        data = {
            "Component": ["Piston Ring", "Fuel Injector", "Turbo Bearings"],
            "Type": ["MAIN ENGINE", "AUX ENGINE", "OTHER EQUIP"],
            "Periodicity": [10000, 4000, 8000],
            "Hours": [10500, 3500, 2000],
            "Ratio": [1.05, 0.875, 0.25],
            "Status": ["OVERDUE", "HIGH PRIORITY", "OK"]
        }
        df = pd.DataFrame(data)

    # 3. Dashboard Metrics (Replaces your CreateDashboardCard VBA)
    st.markdown("### Fleet Health Overview")
    col1, col2, col3 = st.columns(3)
    
    overdue_count = len(df[df['Status'] == 'OVERDUE'])
    warning_count = len(df[df['Status'] == 'HIGH PRIORITY'])
    
    with col1:
        st.markdown(f"<div class='metric-card'><h4>MAIN ENGINE</h4><div class='overdue'>{overdue_count} Overdue</div><div class='warning'>{warning_count} Warning</div></div>", unsafe_allow_html=True)
    with col2:
        st.markdown(f"<div class='metric-card'><h4>AUX ENGINES</h4><div class='overdue'>0 Overdue</div><div class='warning'>1 Warning</div></div>", unsafe_allow_html=True)
    with col3:
        st.markdown(f"<div class='metric-card'><h4>OTHER EQUIP</h4><div class='overdue'>0 Overdue</div><div class='warning'>0 Warning</div></div>", unsafe_allow_html=True)

    st.divider()

    # 4. Premium Data Presentation
    st.subheader("⚠️ Alert Matrix")
    alerts_df = df[df['Status'].isin(['OVERDUE', 'HIGH PRIORITY'])]
    
    # Apply Pandas Styling to mimic your Excel Red/Yellow cell highlighting
    def highlight_status(val):
        if val == 'OVERDUE':
            return 'background-color: #FFC7CE; color: #9C0006; font-weight: bold'
        elif val == 'HIGH PRIORITY':
            return 'background-color: #FFF2CC; color: #9C6500; font-weight: bold'
        elif val == 'OK':
            return 'background-color: #C6EFCE; color: #006100;'
        return ''

    styled_alerts = alerts_df.style.map(highlight_status, subset=['Status']).format({'Ratio': "{:.1%}"})
    st.dataframe(styled_alerts, use_container_width=True, hide_index=True)
