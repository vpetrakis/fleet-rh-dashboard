import docx, re, subprocess, os

def convert_doc_to_docx(file_path):
    output_dir = os.path.dirname(file_path) or "."
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'docx', file_path, '--outdir', output_dir])
    return file_path + "x"

def parse_running_hours_report(file_path):
    if file_path.endswith('.doc'):
        file_path = convert_doc_to_docx(file_path)
    doc = docx.Document(file_path)
    data = {"vessel_name": "Unknown", "components": []}
    
    for para in doc.paragraphs:
        if "Vessel’s Name:" in para.text:
            data["vessel_name"] = para.text.split(":")[1].split("Date")[0].strip() # [cite: 72]

    for table in doc.tables:
        if "CYL. No.1" in table.rows[0].text: # [cite: 73]
            current_comp = ""
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if cells[0] and not cells[0].isdigit():
                    current_comp = cells[0]
                if cells[1] == '2': # Row '2' is Running Hours [cite: 78]
                    for i in range(2, 9): # Cylinders 1-7 [cite: 73]
                        if i < len(cells) and cells[i]:
                            data["components"].append({
                                "name": current_comp,
                                "cyl": i-1,
                                "hours": int(re.sub(r'[^0-9]', '', cells[i]))
                            })
    return data