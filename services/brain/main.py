"""
Fleet Command Nexus — main.py
FastAPI backend.  Run from services/brain/:
    python3 main.py
"""

import asyncio
import hashlib
import json
import logging
import os
import uuid
from datetime import datetime, timezone
from pathlib import Path

import uvicorn
from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("fleet_nexus")

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR    = Path(__file__).parent
MEMORY_FILE = BASE_DIR / "vessel_memory.json"
TEMP_DIR    = BASE_DIR / "tmp"
TEMP_DIR.mkdir(exist_ok=True)

ALLOWED_EXT    = {".doc", ".docx"}
MAX_BYTES      = 15 * 1024 * 1024  # 15 MB

# ── App ───────────────────────────────────────────────────────────────────────
app = FastAPI(title="Fleet Command Nexus API", version="5.0.0")

app.add_middleware(
    CORSMiddleware,
    # Allow both localhost and GitHub Codespaces origins
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


# ── Persistence helpers ───────────────────────────────────────────────────────

def _load() -> dict | None:
    try:
        if MEMORY_FILE.exists() and MEMORY_FILE.stat().st_size > 2:
            data = json.loads(MEMORY_FILE.read_text("utf-8"))
            if data.get("status") == "success" and data.get("vessel"):
                return data
    except Exception as e:
        log.warning("Could not read vessel_memory.json: %s", e)
    return None


def _save(payload: dict) -> None:
    try:
        payload["_saved_at"] = datetime.now(timezone.utc).isoformat()
        MEMORY_FILE.write_text(json.dumps(payload, indent=2, ensure_ascii=False), "utf-8")
        log.info("Saved → vessel_memory.json  (vessel=%s)", payload.get("vessel"))
    except Exception as e:
        log.warning("Could not write vessel_memory.json: %s", e)


# ── Routes ────────────────────────────────────────────────────────────────────

@app.get("/api/init")
async def init():
    """Return last saved parse on browser refresh."""
    data = _load()
    if data:
        log.info("Boot state: %s  integrity=%s%%", data.get("vessel"), data.get("integrity"))
        return JSONResponse(content=data)
    return JSONResponse(content={"status": "idle"})


@app.post("/api/upload-report")
async def upload_report(file: UploadFile = File(...)):
    """Accept a .doc/.docx report, parse it, persist and return result."""
    fname = file.filename or "unknown"

    # Validate extension
    suffix = Path(fname).suffix.lower()
    if suffix not in ALLOWED_EXT:
        raise HTTPException(400, f"Unsupported file type '{suffix}'. Use .doc or .docx.")

    # Stream to temp file with size cap
    tmp_path = TEMP_DIR / f"{uuid.uuid4().hex}_{fname}"
    written = 0
    try:
        with tmp_path.open("wb") as fh:
            while chunk := await file.read(65_536):
                written += len(chunk)
                if written > MAX_BYTES:
                    raise HTTPException(413, f"File exceeds {MAX_BYTES//1024//1024} MB limit.")
                fh.write(chunk)

        if written == 0:
            raise HTTPException(400, "Uploaded file is empty.")

        log.info("Received '%s'  %d KB", fname, written // 1024)

        # Parse in thread pool (CPU-bound work must not block the event loop)
        from parser import extract_vessel_data          # noqa: PLC0415
        loop   = asyncio.get_running_loop()
        result = await loop.run_in_executor(None, extract_vessel_data, str(tmp_path))

        if not isinstance(result, dict):
            raise HTTPException(500, "Parser returned unexpected data type.")

        if result.get("status") == "error":
            raise HTTPException(422, result.get("detail", "Unknown parse error."))

        # Enrich with upload metadata
        result["_upload"] = {
            "original_filename": fname,
            "bytes":             written,
            "uploaded_at":       datetime.now(timezone.utc).isoformat(),
        }
        _save(result)
        return JSONResponse(content=result)

    except HTTPException:
        raise
    except Exception as exc:
        log.exception("Unhandled error for '%s': %s", fname, exc)
        raise HTTPException(500, f"Internal server error: {exc}") from exc
    finally:
        try:
            if tmp_path.exists():
                tmp_path.unlink()
        except OSError:
            pass


@app.get("/api/health")
async def health():
    """Liveness probe used by start_bot.sh."""
    return {
        "status":    "ok",
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "memory":    MEMORY_FILE.exists(),
    }


@app.exception_handler(404)
async def _404(_req: Request, _exc: Exception):
    return JSONResponse({"status": "error", "detail": "Route not found."}, status_code=404)


@app.exception_handler(500)
async def _500(_req: Request, exc: Exception):
    log.exception("500: %s", exc)
    return JSONResponse({"status": "error", "detail": "Internal server error."}, status_code=500)


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    uvicorn.run(
        "main:app",
        host="0.0.0.0",       # Required for GitHub Codespaces
        port=8000,
        reload=True,
        log_level="info",
    )