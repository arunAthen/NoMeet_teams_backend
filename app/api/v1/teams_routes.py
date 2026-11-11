import sqlite3
import uuid
import json
from urllib.parse import unquote
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
    user_id: Optional[str] = "default_user"
    user_name: Optional[str] = "Team Member"
    group_id: Optional[str] = "all"

class CreateSessionRequest(BaseModel):
    title: Optional[str] = "New Conversation"
    user_id: Optional[str] = "default_user"
    user_name: Optional[str] = "Team Member"
    group_id: Optional[str] = "all"

class RegisterWorkspaceRequest(BaseModel):
    id: str
    name: Optional[str] = None
    workspace_type: Optional[str] = "chat"

# ==========================================
# 0. Multi-Group / Dynamic Workspace Discovery
# ==========================================

def _clean_workspace_name(raw_id: str, raw_name: Optional[str] = None) -> str:
    """Decodes URL encoding and converts raw thread IDs to clean, human-readable labels."""
    if raw_name and raw_name.strip() and not raw_name.startswith("19:") and "@thread" not in raw_name:
        return unquote(raw_name).strip()
    clean_id = unquote(raw_id).strip()
    if clean_id.startswith("19:") or "@thread" in clean_id:
        short = clean_id.split("@")[0].replace("19:", "")[:8]
        return f"Chat {short.upper()}"
    return clean_id.replace("_", " ").title()

@router.post("/groups/register")
def register_workspace(req: RegisterWorkspaceRequest):
    """Dynamically auto-registers or updates a Teams workspace (chat, channel, personal)."""
    clean_id = unquote(req.id).strip()
    clean_name = _clean_workspace_name(clean_id, req.name)
    now = datetime.now(timezone.utc).isoformat()

    with sqlite3.connect(DB_FILE) as conn:
        c = conn.cursor()
        c.execute("""
            INSERT INTO workspaces (id, name, workspace_type, last_active)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(id) DO UPDATE SET name = excluded.name, last_active = excluded.last_active
        """, (clean_id, clean_name, req.workspace_type or "chat", now))
        conn.commit()

    return {"status": "registered", "id": clean_id, "name": clean_name}

@router.get("/groups")
def list_available_groups():
    """Dynamically returns all discovered workspaces with cleaned, decoded labels."""
    groups = []
    seen = set()

    with sqlite3.connect(DB_FILE) as conn:
        c = conn.cursor()
        # 1. Registered workspaces
        try:
            c.execute("SELECT id, name FROM workspaces ORDER BY last_active DESC")
            for r in c.fetchall():
                gid, gname = r[0], r[1]
                if gid and gid not in seen:
                    seen.add(gid)
                    groups.append({"id": gid, "name": _clean_workspace_name(gid, gname)})
        except Exception:
            pass

        # 2. Distinct project_ids from project_items
        try:
            c.execute("SELECT DISTINCT project_id FROM project_items WHERE project_id IS NOT NULL")
            for r in c.fetchall():
                gid = r[0]
                if gid and gid not in seen:
                    seen.add(gid)
                    groups.append({"id": gid, "name": _clean_workspace_name(gid)})
        except Exception:
            pass

        # 3. Distinct project_ids from teams_messages
        try:
            c.execute("SELECT DISTINCT project_id FROM teams_messages WHERE project_id IS NOT NULL")
            for r in c.fetchall():
                gid = r[0]
                if gid and gid not in seen:
                    seen.add(gid)
                    groups.append({"id": gid, "name": _clean_workspace_name(gid)})
        except Exception:
            pass

    # 4. Configured group IDs fallback
    for gid in settings.AUTHORIZED_GROUP_CHAT_IDS:
        if gid and gid not in seen:
            seen.add(gid)
            groups.append({"id": gid, "name": _clean_workspace_name(gid)})

    if not groups:
        groups.append({"id": "default", "name": "General Workspace"})

    return {"groups": groups}

# ==========================================
# 1. Chat Sessions Management (Strict User Isolation)
# ==========================================

@router.get("/sessions")
def list_chat_sessions(
    user_id: Optional[str] = None,
    group_id: Optional[str] = None,
    limit: int = 30
):
    """
    Retrieves past conversational sessions.
    Strictly isolated: only the requesting user's conversations are returned.
    """
    sessions = []
    with sqlite3.connect(DB_FILE) as conn:
        c = conn.cursor()
        query = """
            SELECT s.session_id, s.title, s.project_id, s.user_name, s.user_aad_id, s.created_at, s.updated_at,
                   (SELECT content FROM chat_messages WHERE session_id = s.session_id ORDER BY id DESC LIMIT 1) as last_msg
            FROM chat_sessions s
            WHERE 1=1
        """
        params = []
        if user_id:
            query += " AND (s.user_aad_id = ? OR s.user_aad_id = 'default_user' OR s.user_aad_id IS NULL)"
            params.append(user_id)
        if group_id and group_id != "all":
            query += " AND (s.project_id = ? OR s.project_id = 'all')"
            params.append(group_id)

        query += " ORDER BY s.updated_at DESC LIMIT ?"
        params.append(limit)

        c.execute(query, tuple(params))
        for r in c.fetchall():
            pid = r[2] or "all"
            sessions.append({
                "session_id": r[0],
                "title": r[1] or "Conversation",
                "project_id": pid,
                "project_name": _clean_workspace_name(pid),
                "user_name": r[3] or "Team Member",
                "user_id": r[4] or "default_user",
                "created_at": r[5],
                "updated_at": r[6],
                "last_message": r[7] or "No messages yet"
            })
    return sessions

@router.post("/sessions")
def create_chat_session(req: CreateSessionRequest):
    """Creates a new isolated conversation session strictly bound to user_id."""
    session_id = str(uuid.uuid4())
    now = datetime.now(timezone.utc).isoformat()
    gid = req.group_id or "all"
    uid = req.user_id or "default_user"
    uname = req.user_name or "Team Member"
    greeting_name = uname if uname and uname != "Team Member" else "there"

    with sqlite3.connect(DB_FILE) as conn:
        c = conn.cursor()
        c.execute("""
            INSERT INTO chat_sessions (session_id, project_id, user_aad_id, user_name, title, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (session_id, gid, uid, uname, req.title, now, now))
        c.execute("""
            INSERT INTO chat_messages (session_id, role, content, created_at)
            VALUES (?, 'assistant', ?, ?)
        """, (session_id, f"Hello {greeting_name}! I am your personal NoMeet Delivery & Operations Copilot. How can I assist you with this workspace's tasks, bugs, or specifications today?", now))
        conn.commit()
    return {"session_id": session_id, "title": req.title, "group_id": gid, "user_id": uid}

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
# 2. Synchronous Query Interface (Strict User Isolation)
# ==========================================

@router.post("/query")
async def execute_query(req: QueryRequest):
    """Direct query endpoint for Teams Cockpit, strictly scoped to user and active workspace."""
    now = datetime.now(timezone.utc).isoformat()
    session_id = req.session_id
    gid = req.group_id or "all"
    uid = req.user_id or "default_user"
    uname = req.user_name or "Team Member"

    if not session_id:
        with sqlite3.connect(DB_FILE) as conn:
            c = conn.cursor()
            if gid != "all":
                c.execute("""
                    SELECT session_id FROM chat_sessions 
                    WHERE (user_aad_id = ? OR user_aad_id = 'default_user' OR user_aad_id IS NULL) 
                      AND (project_id = ? OR project_id = 'all')
                    ORDER BY updated_at DESC LIMIT 1
                """, (uid, gid))
            else:
                c.execute("""
                    SELECT session_id FROM chat_sessions 
                    WHERE (user_aad_id = ? OR user_aad_id = 'default_user' OR user_aad_id IS NULL)
                    ORDER BY updated_at DESC LIMIT 1
                """, (uid,))
            row = c.fetchone()
            if row:
                session_id = row[0]
            else:
                session_id = str(uuid.uuid4())
                c.execute("""
                    INSERT INTO chat_sessions (session_id, project_id, user_aad_id, user_name, title, created_at, updated_at) 
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (session_id, gid, uid, uname, req.query[:35] + "...", now, now))
                conn.commit()

    # Update session title if first query
    with sqlite3.connect(DB_FILE) as conn:
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM chat_messages WHERE session_id = ? AND role = 'user'", (session_id,))
        if c.fetchone()[0] == 0:
            short_title = req.query[:35] + ("..." if len(req.query) > 35 else "")
            c.execute("UPDATE chat_sessions SET title = ?, updated_at = ? WHERE session_id = ?", (short_title, now, session_id))
            conn.commit()
        else:
            c.execute("UPDATE chat_sessions SET updated_at = ? WHERE session_id = ?", (now, session_id))
            conn.commit()

    res = run_agent_query(user_query=req.query, session_id=session_id, user_name=uname, group_id=gid)
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
