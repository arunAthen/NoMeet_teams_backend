import sqlite3
import uuid
import json
from datetime import datetime, timezone
from typing import Dict, Any, Optional, List
from fastapi import APIRouter, Request, HTTPException
from pydantic import BaseModel
from app.core.config import settings
from app.integrations.teams.bot_adapter import handle_teams_bot_activity
from app.integrations.teams.webhook_listener import handle_graph_chat_webhook
from app.integrations.teams.graph_delta_sync import sync_sharepoint_files_delta
from app.agents.delivery_agent import run_agent_query

router = APIRouter()
DB_FILE = settings.DATABASE_PATH

class QueryRequest(BaseModel):
    query: str
    session_id: Optional[str] = None
    user_name: Optional[str] = "Team Member"
    group_id: Optional[str] = "all"

class CreateSessionRequest(BaseModel):
    title: Optional[str] = "New Conversation"
    user_name: Optional[str] = "Team Member"
    group_id: Optional[str] = "all"

# ==========================================
# 0. Multi-Group / Workspace Discovery
# ==========================================

@router.get("/groups")
def list_available_groups():
    """Dynamically returns all distinct group workspaces from the database and configuration."""
    groups = []
    seen = set()

    with sqlite3.connect(DB_FILE) as conn:
        c = conn.cursor()
        try:
            c.execute("SELECT DISTINCT project_id FROM project_items WHERE project_id IS NOT NULL")
            for r in c.fetchall():
                gid = r[0]
                if gid and gid not in seen:
                    seen.add(gid)
                    groups.append({"id": gid, "name": gid.replace("_", " ").title()})
        except Exception:
            pass

        try:
            c.execute("SELECT DISTINCT project_id FROM teams_messages WHERE project_id IS NOT NULL")
            for r in c.fetchall():
                gid = r[0]
                if gid and gid not in seen:
                    seen.add(gid)
                    groups.append({"id": gid, "name": gid.replace("_", " ").title()})
        except Exception:
            pass

    for gid in settings.AUTHORIZED_GROUP_CHAT_IDS:
        if gid and gid not in seen:
            seen.add(gid)
            short_label = gid.split("@")[0].replace("19:", "")[:16]
            groups.append({"id": gid, "name": f"Teams: {short_label}"})

    if not groups:
        groups.append({"id": "default", "name": "General Workspace"})

    return {"groups": groups}

# ==========================================
# 1. Chat Sessions Management (Group-Scoped)
# ==========================================

@router.get("/sessions")
def list_chat_sessions(group_id: Optional[str] = None, limit: int = 25):
    """Retrieves past conversational sessions, filtered strictly by group workspace if specified."""
    sessions = []
    with sqlite3.connect(DB_FILE) as conn:
        c = conn.cursor()
        if group_id and group_id != "all":
            c.execute("""
                SELECT s.session_id, s.title, s.user_name, s.created_at, s.updated_at,
                       (SELECT content FROM chat_messages WHERE session_id = s.session_id ORDER BY id DESC LIMIT 1) as last_msg
                FROM chat_sessions s
                WHERE s.project_id = ?
                ORDER BY s.updated_at DESC
                LIMIT ?
            """, (group_id, limit))
        else:
            c.execute("""
                SELECT s.session_id, s.title, s.user_name, s.created_at, s.updated_at,
                       (SELECT content FROM chat_messages WHERE session_id = s.session_id ORDER BY id DESC LIMIT 1) as last_msg
                FROM chat_sessions s
                ORDER BY s.updated_at DESC
                LIMIT ?
            """, (limit,))
        for r in c.fetchall():
            sessions.append({
                "session_id": r[0],
                "title": r[1] or "Conversation",
                "user_name": r[2] or "Team Member",
                "created_at": r[3],
                "updated_at": r[4],
                "last_message": r[5] or "No messages yet"
            })
    return sessions

@router.post("/sessions")
def create_chat_session(req: CreateSessionRequest):
    """Creates a new isolated conversation session tagged with the active group workspace."""
    session_id = str(uuid.uuid4())
    now = datetime.now(timezone.utc).isoformat()
    gid = req.group_id or "all"
    with sqlite3.connect(DB_FILE) as conn:
        c = conn.cursor()
        c.execute("""
            INSERT INTO chat_sessions (session_id, project_id, title, user_name, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (session_id, gid, req.title, req.user_name, now, now))
        c.execute("""
            INSERT INTO chat_messages (session_id, role, content, created_at)
            VALUES (?, 'assistant', ?, ?)
        """, (session_id, "Hello! I am your NoMeet Delivery & Operations Copilot. How can I assist you with this workspace's tasks, bugs, or specifications today?", now))
        conn.commit()
    return {"session_id": session_id, "title": req.title, "group_id": gid}

@router.get("/sessions/{session_id}")
def get_session_history(session_id: str):
    """Retrieves all chronological messages for a specific session."""
    messages = []
    with sqlite3.connect(DB_FILE) as conn:
        c = conn.cursor()
        c.execute("""
            SELECT id, role, content, evidence, eval_scores, created_at
            FROM chat_messages
            WHERE session_id = ?
            ORDER BY id ASC
        """, (session_id,))
        for r in c.fetchall():
            messages.append({
                "id": r[0],
                "role": r[1],
                "content": r[2],
                "evidence": json.loads(r[3]) if r[3] else None,
                "eval_scores": json.loads(r[4]) if r[4] else None,
                "created_at": r[5]
            })
    return {"session_id": session_id, "messages": messages}

@router.delete("/sessions/{session_id}")
def delete_chat_session(session_id: str):
    """Deletes a chat session and its conversation messages."""
    with sqlite3.connect(DB_FILE) as conn:
        c = conn.cursor()
        c.execute("DELETE FROM chat_messages WHERE session_id = ?", (session_id,))
        c.execute("DELETE FROM chat_sessions WHERE session_id = ?", (session_id,))
        conn.commit()
    return {"status": "deleted", "session_id": session_id}

# ==========================================
# 2. Synchronous Query Interface (Group-Scoped)
# ==========================================

@router.post("/query")
async def execute_query(req: QueryRequest):
    """Direct query endpoint for Teams Cockpit, strictly scoped to the active workspace."""
    now = datetime.now(timezone.utc).isoformat()
    session_id = req.session_id
    gid = req.group_id or "all"

    if not session_id:
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            if gid != "all":
                c.execute("SELECT session_id FROM chat_sessions WHERE project_id = ? ORDER BY updated_at DESC LIMIT 1", (gid,))
            else:
                c.execute("SELECT session_id FROM chat_sessions ORDER BY updated_at DESC LIMIT 1")
            row = c.fetchone()
            if row:
                session_id = row[0]
            else:
                session_id = str(uuid.uuid4())
                c.execute("INSERT INTO chat_sessions (session_id, project_id, title, created_at, updated_at) VALUES (?, ?, ?, ?, ?)", (session_id, gid, req.query[:35] + "...", now, now))
                conn.commit()

    # Update session title if first query
    with sqlite3.connect(DB_FILE) as conn:
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM chat_messages WHERE session_id = ? AND role = 'user'", (session_id,))
        if c.fetchone()[0] == 0:
            short_title = req.query[:35] + ("..." if len(req.query) > 35 else "")
            c.execute("UPDATE chat_sessions SET title = ?, updated_at = ? WHERE session_id = ?", (short_title, now, session_id))
            conn.commit()

    res = run_agent_query(user_query=req.query, session_id=session_id, user_name=req.user_name, group_id=gid)
    return res

# ==========================================
# 3. Teams Cloud Webhooks & Bot
# ==========================================

@router.post("/messages")
async def teams_bot_messages_webhook(request: Request):
    """Receives 1:1 Bot Framework activities from Microsoft Teams."""
    try:
        activity = await request.json()
        resp = await handle_teams_bot_activity(activity)
        return resp
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/webhook")
async def teams_graph_change_webhook(request: Request):
    """Receives Microsoft Graph change notifications (message edits/creations)."""
    val_token = request.query_params.get("validationToken")
    if val_token:
        from fastapi.responses import PlainTextResponse
        return PlainTextResponse(val_token, status_code=200)

    payload = await request.json()
    result = await handle_graph_chat_webhook(payload)
    return result

@router.post("/sync/trigger")
async def trigger_manual_sync():
    """Triggers recursive crawl and sync over project files and Excel sheets."""
    stats = sync_sharepoint_files_delta()
    return {"status": "completed", "stats": stats}

@router.get("/overview")
async def get_project_overview(group_id: Optional[str] = None):
    """Returns high-level project health filtered strictly by the group workspace."""
    with sqlite3.connect(DB_FILE) as conn:
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        
        if group_id and group_id != "all":
            c.execute("SELECT status, count(*) as count FROM project_items WHERE item_type = 'change_request' AND project_id = ? GROUP BY status", (group_id,))
            crs = {r["status"]: r["count"] for r in c.fetchall()}
            
            c.execute("SELECT status, count(*) as count FROM project_items WHERE item_type = 'bug' AND project_id = ? GROUP BY status", (group_id,))
            bugs = {r["status"]: r["count"] for r in c.fetchall()}
            
            c.execute("SELECT count(*) FROM teams_messages WHERE is_technical_signal = 1 AND project_id = ?", (group_id,))
            total_msgs = c.fetchone()[0]

            c.execute("SELECT count(*) FROM specs_fts")
            total_specs = c.fetchone()[0]
        else:
            c.execute("SELECT status, count(*) as count FROM project_items WHERE item_type = 'change_request' GROUP BY status")
            crs = {r["status"]: r["count"] for r in c.fetchall()}
            
            c.execute("SELECT status, count(*) as count FROM project_items WHERE item_type = 'bug' GROUP BY status")
            bugs = {r["status"]: r["count"] for r in c.fetchall()}
            
            c.execute("SELECT count(*) FROM teams_messages WHERE is_technical_signal = 1")
            total_msgs = c.fetchone()[0]

            c.execute("SELECT count(*) FROM specs_fts")
            total_specs = c.fetchone()[0]

    return {
        "group_id": group_id or "all",
        "change_requests": {"total": sum(crs.values()), "breakdown": crs},
        "bugs": {"total": sum(bugs.values()), "breakdown": bugs},
        "indexed_communications": total_msgs,
        "indexed_specifications": total_specs,
        "status": "synchronized"
    }
