import sqlite3
from pathlib import Path
from sqlmodel import SQLModel, create_engine, Session
from app.core.config import settings

DB_FILE = settings.DATABASE_PATH
try:
    DB_FILE.parent.mkdir(parents=True, exist_ok=True)
    # Probe writability
    test_probe = DB_FILE.parent / ".write_probe"
    test_probe.touch()
    test_probe.unlink()
except (OSError, PermissionError):
    # Serverless fallback for Vercel / AWS Lambda
    DB_FILE = Path("/tmp") / "teams_365_agent.db"
    settings.DATABASE_PATH = DB_FILE

DATABASE_URL = f"sqlite:///{DB_FILE}"

engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False, "timeout": 60},
    echo=False
)

def init_db():
    # Configure SQLite pragmas safely
    try:
        with sqlite3.connect(DB_FILE) as conn:
            try:
                conn.execute("PRAGMA journal_mode=WAL;")
            except Exception:
                conn.execute("PRAGMA journal_mode=DELETE;")
            conn.execute("PRAGMA synchronous=NORMAL;")
            conn.execute("PRAGMA busy_timeout=60000;")
            conn.commit()

            # Virtual FTS5 tables for sub-millisecond BM25 search
            conn.execute("""
                CREATE VIRTUAL TABLE IF NOT EXISTS specs_fts USING fts5(
                    doc_title, category, content, tokenize='porter unicode61'
                );
            """)
            conn.execute("""
                CREATE VIRTUAL TABLE IF NOT EXISTS teams_fts USING fts5(
                    message_id UNINDEXED, sender, timestamp UNINDEXED, content, tokenize='porter unicode61'
                );
            """)
            conn.execute("""
                CREATE VIRTUAL TABLE IF NOT EXISTS items_fts USING fts5(
                    ticket_id UNINDEXED, title, description, assignee, status, module, tokenize='porter unicode61'
                );
            """)
            conn.commit()

        # Create relational tables
        from app.db.schema import ProjectItem, TeamsMessage, ChatSession, ChatMessage, DataSource, AuditEvent
        SQLModel.metadata.create_all(engine)
    except Exception as e:
        print(f"[Database init warning]: {e}")
