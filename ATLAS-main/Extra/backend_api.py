# run with: uvicorn backend_api:app --host 127.0.0.1 --port 8000 --reload
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from pathlib import Path
import tempfile, shutil, os

# your existing modules
from ras_module import build_ras
from tiv_module import build_tiv

app = FastAPI(title="RAS/TIV Alloc API")

def _save_upload(u: UploadFile) -> Path:
    if not u.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Upload must be .xlsx")
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    with tmp:
        shutil.copyfileobj(u.file, tmp)
    return Path(tmp.name)

@app.post("/build")
def build(mode: str, file: UploadFile = File(...)):
    """
    POST /build?mode=ras   or   /build?mode=tiv
    form-data: file=<xlsx>
    """
    mode = (mode or "").lower().strip()
    if mode not in {"ras", "tiv"}:
        raise HTTPException(400, "mode must be 'ras' or 'tiv'")

    src = _save_upload(file)
    try:
        out_path = build_ras(str(src)) if mode == "ras" else build_tiv(str(src))
        filename = Path(out_path).name
        return FileResponse(path=out_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    finally:
        try: os.remove(src)
        except Exception: pass
