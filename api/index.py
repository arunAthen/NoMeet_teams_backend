import sys
import traceback
from pathlib import Path
from fastapi import FastAPI
from fastapi.responses import JSONResponse

# Add repo root to sys.path for Vercel Serverless environment
root_dir = Path(__file__).resolve().parent.parent
backend_dir = root_dir / "backend"

if backend_dir.exists() and str(backend_dir) not in sys.path:
    sys.path.insert(0, str(backend_dir))
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

try:
    from app.main import app
except Exception as e:
    err_tb = traceback.format_exc()
    app = FastAPI(title="Diagnostic Fallback")

    @app.api_route("/{path:path}", methods=["GET", "POST", "PUT", "DELETE"])
    def catch_all(path: str = ""):
        return JSONResponse(
            status_code=500,
            content={
                "status": "error",
                "message": "Import or initialization failed in app.main",
                "exception_type": type(e).__name__,
                "exception": str(e),
                "traceback": err_tb.splitlines()
            }
        )
