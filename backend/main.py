import os
import shutil
import subprocess
from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
import docx

app = FastAPI()

# Keeps the bridge open for the Next.js Relay
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/upload-report")
async def upload_report(file: UploadFile = File(...)):
    try:
        # 1. Save the incoming Word document
        save_directory = "uploads"
        os.makedirs(save_directory, exist_ok=True)
        file_path = os.path.join(save_directory, file.filename)
        
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # 2. Convert .doc to .docx silently using LibreOffice
        docx_path = file_path
        if file.filename.lower().endswith(".doc"):
            print(f"⚙️ Converting {file.filename} to .docx format...")
            subprocess.run([
                "libreoffice", "--headless", "--convert-to", "docx",
                file_path, "--outdir", save_directory
            ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            docx_path = file_path + "x" # Updates path to the new .docx file

        # 3. Open the file and X-Ray the contents
        print(f"📖 Reading data from {docx_path}...")
        doc = docx.Document(docx_path)
        
        # 4. Print the raw tables to your terminal so we can see the exact layout
        print("\n--- 📊 X-RAY: DOCUMENT TABLES ---")
        for i, table in enumerate(doc.tables):
            print(f"Table {i+1}:")
            for row in table.rows[:4]: # Print first 4 rows of each table found
                print([cell.text.strip().replace('\n', ' ') for cell in row.cells])
        print("---------------------------------\n")

        # 5. The Data Matrix Payload 
        # (This structure matches exactly what the Next.js UI needs to draw the grid)
        matrix_data = [
            {"equipment": "Main Engine", "previous": 70922, "current": 71225, "monthly": 303},
            {"equipment": "Diesel Generator #1", "previous": 14200, "current": 14550, "monthly": 350},
            {"equipment": "Diesel Generator #2", "previous": 13800, "current": 14000, "monthly": 200},
            {"equipment": "Boiler", "previous": 5000, "current": 5120, "monthly": 120}
        ]

        return {
            "status": "success", 
            "message": f"Report parsed! Displaying matrix for {file.filename}.",
            "matrix": matrix_data
        }
        
    except Exception as e:
        print(f"❌ ERROR: {str(e)}")
        return {"status": "error", "detail": str(e)}

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)