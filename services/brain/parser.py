import os
import re

# 🚢 THE FLEET REGISTRY: Define your vessels, their cylinder counts, and custom limits here.
FLEET_CONFIG = {
    "MV ALEXIS": {
        "me_cylinders": 5,      # 5-Cylinder Main Engine
        "dg_count": 3,          # 3 Diesel Generators
        "aux_cylinders": 6,     # 6-Cylinder Aux Engines
        "me_comps": {
            "CYLINDER COVER": 16000, "PISTON ASSEMBLY": 16000, "STUFFING BOX": 16000, 
            "PISTON CROWN": 32000, "CYLINDER LINER": 32000, "EXHAUST VALVE": 16000, 
            "STARTING VALVE": 12000, "SAFETY VALVE": 12000, "FUEL VALVES": 8000, 
            "FUEL PUMP": 16000, "CROSSHEAD BEARINGS": 32000, "BOTTOM END BEARINGS": 32000, 
            "MAIN BEARINGS": 32000
        },
        "dg_comps": {
            "Turbocharger (3)": 12000, "Air Cooler": 5000, "L.O. Cooler Clean": 8000,
            "Cooling Water Pump": 5000, "F.W. Cooler Clean": 4000,
            "Cool Water Thermostat Valve": 5000, "L.O. Renewal": 1500,
            "L.O. Thermostat Valve": 10000, "Alternator Cleaning": 5000,
            "Thrust Bearing": 12000
        },
        "aux_comps": {
            "Cylinder Head": 12000, "Piston": 10000, "Connecting Rod": 10000,
            "Cylinder Liners": 10000, "Fuel Valves (1)": 2000, "Fuel Pumps": 5000,
            "Crank Pin Bearing": 12000, "Main Bearing": 12000, "Adjust Valve Head Clearance": 1200
        }
    }
    # To add another ship, just copy the block above and name it "MV OLYMPIA" (for example), 
    # then change "me_cylinders" to 6 or 7.
}

def extract_vessel_data(file_path):
    if not os.path.exists(file_path): 
        return {"status": "error", "detail": "File not found"}

    try:
        with open(file_path, 'rb') as f:
            raw_data = f.read().decode('latin-1', errors='ignore')
        
        clean_text = raw_data.replace('\x07', ' ').replace('\r', ' ').replace('\n', ' ')
        clean_text = re.sub(r'\s+', ' ', clean_text)

        # 1. AUTO-DETECT THE VESSEL
        vessel_match = re.search(r"Vessel.*?Name:\s*([A-Za-z0-9\s]+?)\s*(?:Date|Month|1-DATE)", clean_text, re.IGNORECASE)
        detected_name = vessel_match.group(1).strip() if vessel_match else "MV ALEXIS"
        
        # Pull the specific vessel's configuration (Defaults to MV ALEXIS if unknown)
        config = FLEET_CONFIG.get(detected_name, FLEET_CONFIG["MV ALEXIS"])
        matrices = {"main_engine": [], "aux_engine": [], "dg_general": []}

        def extract_row(comp_name, count, search_name=None):
            search_term = search_name if search_name else comp_name
            idx = clean_text.upper().find(search_term.upper())
            if idx != -1:
                chunk = clean_text[idx:idx+800] 
                pattern = r"\b2\b\s*((?:\[?[\d,]+\]?\s*){" + str(count) + r",})"
                match = re.search(pattern, chunk)
                if match:
                    tokens = re.findall(r"\[?([\d,]+)\]?", match.group(1))
                    return [t.replace(',', '') for t in tokens][:count]
            return ["-"] * count

        # 2. DYNAMIC MAIN ENGINE EXTRACTION
        for comp, limit in config["me_comps"].items():
            vals = extract_row(comp, config["me_cylinders"])
            row_data = {"component": comp, "limit": limit}
            for i in range(config["me_cylinders"]):
                row_data[f"cyl{i+1}"] = vals[i] if i < len(vals) else "-"
            matrices["main_engine"].append(row_data)

        # 3. DYNAMIC D/G EXTRACTION
        for comp, limit in config["dg_comps"].items():
            vals = extract_row(comp, config["dg_count"])
            row_data = {"component": comp, "limit": limit}
            for i in range(config["dg_count"]):
                row_data[f"dg{i+1}"] = vals[i] if i < len(vals) else "-"
            matrices["dg_general"].append(row_data)

        # 4. DYNAMIC AUX ENGINE EXTRACTION
        total_aux_vals = config["aux_cylinders"] * config["dg_count"]
        for comp, limit in config["aux_comps"].items():
            search_term = comp
            if comp == "Piston": search_term = f"Piston {limit}" 
            if comp == "Main Bearing": search_term = f"Main Bearing {limit}"
            
            vals = extract_row(comp, total_aux_vals, search_name=search_term)
            row_data = {"component": comp, "limit": limit}
            
            for g in range(config["dg_count"]):
                start = g * config["aux_cylinders"]
                end = start + config["aux_cylinders"]
                row_data[f"dg{g+1}"] = vals[start:end] if start < len(vals) else []
            matrices["aux_engine"].append(row_data)

        # Send the blueprint config to the frontend so it knows how to draw the table
        return {
            "status": "success", 
            "vessel": detected_name,
            "config": {
                "me_cylinders": config["me_cylinders"],
                "dg_count": config["dg_count"],
                "aux_cylinders": config["aux_cylinders"]
            },
            "matrices": matrices
        }

    except Exception as e:
        print(f"Parser Fatal Error: {e}")
        return {"status": "error", "detail": str(e)}