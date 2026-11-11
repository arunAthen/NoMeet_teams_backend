import os
from pathlib import Path
from typing import List
from pydantic import BaseModel
from dotenv import load_dotenv

def _find_repo_root(p: Path) -> Path:
    for parent in [p] + list(p.parents):
        if (parent / "requirements.txt").exists() or (parent / "vercel.json").exists() or (parent / ".git").exists():
            return parent
    return p.parents[2]

BASE_DIR = _find_repo_root(Path(__file__).resolve())
env_file = BASE_DIR / ".env"
if env_file.exists():
    load_dotenv(env_file, override=True)

is_vercel = os.getenv("VERCEL") == "1" or os.getenv("VERCEL_ENV") is not None

class Settings(BaseModel):
    PROJECT_NAME: str = "Teams 365 Delivery Intelligence Agent"
    VERSION: str = "1.0.0"
    API_V1_STR: str = "/api/v1"
    BASE_DIR: Path = BASE_DIR
    DATABASE_PATH: Path = (Path("/tmp") / "teams_365_agent.db") if is_vercel else (BASE_DIR / os.getenv("DATABASE_PATH", "data/teams_365_agent.db"))
    DATA_SYNC_DIR: Path = (Path("/tmp") / "files_sync") if is_vercel else (BASE_DIR / os.getenv("DATA_SYNC_DIR", "data/files_sync"))
    SEED_DATA_DIR: Path = (Path("/tmp") / "seed_data") if is_vercel else (BASE_DIR / os.getenv("SEED_DATA_DIR", "data/seed_data"))
    PUBLIC_BASE_URL: str = os.getenv("PUBLIC_BASE_URL", "https://localhost:8000")

    GROQ_API_KEY: str = os.getenv("GROQ_API_KEY", "")
    GROQ_MODEL: str = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")
    GOOGLE_API_KEY: str = os.getenv("GOOGLE_API_KEY", "")
    GEMINI_MODEL: str = os.getenv("GEMINI_MODEL", "gemini-1.5-flash")
    ENABLE_EXTERNAL_LLM: bool = os.getenv("ENABLE_EXTERNAL_LLM", "true").lower() in ("true", "1")

    # Azure Entra ID App Credentials
    AZURE_CLIENT_ID: str = os.getenv("AZURE_CLIENT_ID", "")
    AZURE_CLIENT_SECRET: str = os.getenv("AZURE_CLIENT_SECRET", "")
    AZURE_TENANT_ID: str = os.getenv("AZURE_TENANT_ID", "")

    # Ring-fenced Group Chats Whitelist
    AUTHORIZED_GROUP_CHAT_IDS_RAW: str = os.getenv("AUTHORIZED_GROUP_CHAT_IDS", "")

    @property
    def AUTHORIZED_GROUP_CHAT_IDS(self) -> List[str]:
        if not self.AUTHORIZED_GROUP_CHAT_IDS_RAW:
            return []
        return [cid.strip() for cid in self.AUTHORIZED_GROUP_CHAT_IDS_RAW.split(",") if cid.strip()]

settings = Settings()
