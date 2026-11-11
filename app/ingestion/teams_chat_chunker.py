import json
import re
from typing import List, Dict, Any, Tuple
import sqlite3
from app.core.config import settings

DB_FILE = settings.DATABASE_PATH

def upsert_teams_message(msg: Dict[str, Any], project_id: str = "teams_delivery") -> None:
    """Atomic UPSERT on message_id: handles live edits and creations seamlessly."""
    with sqlite3.connect(DB_FILE) as conn:
        conn.execute("""
            INSERT INTO teams_messages (
                project_id, message_id, thread_id, sender_name, sender_email,
                timestamp, content, has_attachments, is_technical_signal, has_blocker, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
            ON CONFLICT(message_id) DO UPDATE SET
                content = excluded.content,
                timestamp = excluded.timestamp,
                has_blocker = excluded.has_blocker,
                is_technical_signal = excluded.is_technical_signal,
                updated_at = CURRENT_TIMESTAMP;
        """, (
            project_id,
            msg["message_id"],
            msg.get("thread_id") or msg["message_id"],
            msg.get("sender_name", "Team Member"),
            msg.get("sender_email"),
            msg.get("timestamp", ""),
            msg["content"],
            1 if msg.get("has_attachments") else 0,
            1 if msg.get("is_technical_signal", True) else 0,
            1 if msg.get("has_blocker") else 0
        ))
        
        # Resync virtual FTS5 search index
        conn.execute("DELETE FROM teams_fts WHERE message_id = ?", (msg["message_id"],))
        conn.execute("""
            INSERT INTO teams_fts (message_id, sender, timestamp, content)
            VALUES (?, ?, ?, ?)
        """, (msg["message_id"], msg.get("sender_name", "Team Member"), msg.get("timestamp", ""), msg["content"]))
        conn.commit()

def process_teams_chat_json(file_path: str, project_id: str = "teams_delivery") -> int:
    """Ingests exported or streamed Teams chat JSON preserving conversation threads."""
    with open(file_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    raw_messages = data.get("messages", [])
    count = 0
    for m in raw_messages:
        mid = str(m.get("id"))
        if not mid:
            continue
        text = m.get("text") or re.sub(r'<[^>]+>', ' ', m.get("contentHtml", ""))
        text = re.sub(r'\s+', ' ', text).strip()
        if not text:
            continue

        reply_to = m.get("replyTo")
        parent_id = str(reply_to.get("id")) if isinstance(reply_to, dict) else (str(reply_to).strip() if reply_to else mid)
        
        lower = text.lower()
        has_blk = any(b in lower for b in ["blocker", "blocked", "waiting for", "stuck", "cannot proceed", "failing"])
        
        upsert_teams_message({
            "message_id": mid,
            "thread_id": parent_id,
            "sender_name": m.get("author", "Team Member"),
            "timestamp": m.get("timestamp", ""),
            "content": text,
            "has_attachments": bool(m.get("attachments")),
            "is_technical_signal": True,
            "has_blocker": has_blk
        }, project_id=project_id)
        count += 1

    return count
