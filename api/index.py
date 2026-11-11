import sys
from pathlib import Path

# Add repo root to sys.path for Vercel Serverless environment
root_dir = Path(__file__).resolve().parent.parent
backend_dir = root_dir / "backend"

if backend_dir.exists() and str(backend_dir) not in sys.path:
    sys.path.insert(0, str(backend_dir))
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

from app.main import app
