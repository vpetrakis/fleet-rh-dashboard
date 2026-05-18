import re
import pandas as pd
import numpy as np

def clean_number(val) -> float:
    """Replaces VBA ExtractNumber. Fixes spaces, commas, and strips text."""
    if pd.isna(val) or val == "" or val == "-":
        return 0.0
    
    val_str = str(val).upper().strip()
    if any(word in val_str for word in ["N/A", "CENTRAL", "OBSERVATION", "MONTH", "YEAR"]):
        return 0.0

    # Extract only numbers, commas, and periods using Regex
    # Fixes typos like "21, 476" -> "21476"
    match = re.search(r'[\d\s\,\.]+', val_str)
    if not match:
        return 0.0
        
    num_str = match.group(0).replace(" ", "")
    
    # Handle European vs US decimal formats smartly
    if ',' in num_str and '.' in num_str:
        num_str = num_str.replace(',', '') # Assume comma is thousands
    elif num_str.count(',') == 1 and len(num_str.split(',')[1]) != 3:
        num_str = num_str.replace(',', '.') # Assume comma is decimal

    try:
        return float(num_str)
    except ValueError:
        return 0.0

def clean_date(val) -> str:
    """Validates date strings."""
    if pd.isna(val):
        return "-"
    val_str = str(val).strip()
    if val_str == "" or val_str == "-" or len(val_str) > 20:
        return "-"
    return val_str
