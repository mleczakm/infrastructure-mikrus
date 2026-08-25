import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, Request
from fastapi.responses import JSONResponse, Response

EXTRACTOR = Path(__file__).parent / "egib_extractor.py"
XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SUBPROCESS_TIMEOUT_SECONDS = 150

app = FastAPI()


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/process")
async def process(request: Request, background_tasks: BackgroundTasks):
    body = await request.body()
    if not body:
        return JSONResponse(status_code=400, content={"error": {"message": "Empty request body"}})

    filename = request.headers.get("x-original-filename", "document.pdf")
    suffix = Path(filename).suffix.lower()
    if suffix not in (".pdf", ".html", ".htm"):
        suffix = ".html" if body.lstrip().startswith(b"<") else ".pdf"

    workdir = Path(tempfile.mkdtemp(prefix="egib_"))
    background_tasks.add_task(shutil.rmtree, workdir, True)

    input_path = workdir / f"input{suffix}"
    input_path.write_bytes(body)
    output_base = workdir / "parcels"

    try:
        result = subprocess.run(
            [sys.executable, str(EXTRACTOR), str(input_path), "--output", str(output_base)],
            capture_output=True,
            text=True,
            timeout=SUBPROCESS_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        return JSONResponse(status_code=504, content={"error": {"message": "Processing timed out"}})

    if result.returncode != 0:
        message = (result.stderr or result.stdout or "unknown error").strip().splitlines()[-1]
        return JSONResponse(status_code=500, content={"error": {"message": message}})

    xlsx_path = output_base.with_suffix(".xlsx")
    if not xlsx_path.exists():
        return JSONResponse(status_code=500, content={"error": {"message": "Extractor produced no Excel output"}})

    return Response(
        content=xlsx_path.read_bytes(),
        media_type=XLSX_MEDIA_TYPE,
        headers={"Content-Disposition": 'attachment; filename="parcels.xlsx"'},
    )
