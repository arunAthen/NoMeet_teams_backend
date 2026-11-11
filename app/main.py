import os
from pathlib import Path
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from app.core.config import settings
from app.core.security_headers import TeamsSecurityHeadersMiddleware
from app.db.session import init_db
from app.api.v1.teams_routes import router as teams_router
from app.integrations.teams.graph_delta_sync import sync_sharepoint_files_delta

app = FastAPI(
    title=settings.PROJECT_NAME,
    version=settings.VERSION,
    description="Enterprise Microsoft Teams Delivery Intelligence Agent"
)

# Enable Teams iframe security headers
app.add_middleware(TeamsSecurityHeadersMiddleware)

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount API Routers
app.include_router(teams_router, prefix=f"{settings.API_V1_STR}/teams", tags=["Microsoft Teams"])

@app.get("/")
def root():
    return {
        "status": "online",
        "app": settings.PROJECT_NAME,
        "version": settings.VERSION,
        "health": "/health",
        "docs": "/docs",
        "cockpit": "/teams/tab"
    }

@app.get("/health")
def health_check():
    return {
        "status": "healthy",
        "app": settings.PROJECT_NAME,
        "database": settings.DATABASE_PATH.exists()
    }

# Mount Teams Personal Tab UI (supports bundled static or separated frontend)
static_candidates = [
    settings.BASE_DIR / "app" / "static" / "teams",
    settings.BASE_DIR / "backend" / "app" / "static" / "teams",
    settings.BASE_DIR / "frontend",
]
tab_dir = next((c for c in static_candidates if (c / "index.html").exists()), None)
if tab_dir and tab_dir.exists():
    app.mount("/teams/tab", StaticFiles(directory=str(tab_dir), html=True), name="teams_tab")

@app.on_event("startup")
def on_startup():
    print("[Teams 365 Agent]: Initializing database schema...")
    try:
        init_db()
    except Exception as e:
        print(f"[Startup Database Warning]: {e}")
    try:
        stats = sync_sharepoint_files_delta()
        print(f"[Teams 365 Agent]: Sync complete: {stats}")
    except Exception as e:
        print(f"[Startup Sync Warning]: {e}")
