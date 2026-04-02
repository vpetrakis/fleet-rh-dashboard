import os
import shutil
from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
from parser import extract_vessel_data

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

@app.post("/api/upload-report")
async def upload_report(file: UploadFile = File(...)):
    try:
        os.makedirs("uploads", exist_ok=True)
        save_path = os.path.join("uploads", file.filename)
        with open(save_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # The parser now handles the dictionary formatting
        result = extract_vessel_data(save_path)
        return result
        
    except Exception as e:
        return {"status": "error", "detail": str(e)}

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)