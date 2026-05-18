import docx
import pandas as pd
from backend.firewalls import clean_number, clean_date
from backend.config import THRESHOLD_RED, THRESHOLD_YELLOW

def extract_vessel_data(file_path):
    """Reads the docx and extracts the running hours into a structured format."""
    doc = docx.Document(file_path)
    extracted_data = []

    # Iterate through all tables in the Word document
    for table in doc.tables:
        for r_idx, row in enumerate(table.rows):
            # This logic mimics your heuristic "PERIOD" and "REMARK" hunters
            row_data = [cell.text.strip().upper() for cell in row.cells]
            
            # Example heuristic extraction (simplified for the master architecture)
            if "MAIN ENGINE" in "".join(row_data):
                # Trigger Main Engine processing logic
                pass 
                
            # Assume we extracted a component, period, date, and hours:
            # This is where your specific table coordinate logic goes
            # ... [Extraction logic adapted from your VBA] ...
            
            # For demonstration of the pipeline, let's append structured data
            # extracted_data.append({
            #    "Component": comp, "Type": engine_type, "Periodicity": period, 
            #    "Date": date, "Hours": hrs
            # })

    df = pd.DataFrame(extracted_data)
    return calculate_status(df)

def calculate_status(df):
    """Calculates the ratios and assigns OK, HIGH PRIORITY, or OVERDUE."""
    if df.empty:
        return df
        
    df['Hours'] = df['Hours'].apply(clean_number)
    df['Periodicity'] = df['Periodicity'].apply(clean_number)
    
    # Calculate ratio safely avoiding division by zero
    df['Ratio'] = np.where(df['Periodicity'] > 0, df['Hours'] / df['Periodicity'], 0)
    
    # Vectorized condition mapping (Incredibly fast compared to VBA loops)
    conditions = [
        (df['Hours'] == 0),
        (df['Ratio'] >= THRESHOLD_RED),
        (df['Ratio'] >= THRESHOLD_YELLOW)
    ]
    choices = ['NO DATA', 'OVERDUE', 'HIGH PRIORITY']
    df['Status'] = np.select(conditions, choices, default='OK')
    
    return df
