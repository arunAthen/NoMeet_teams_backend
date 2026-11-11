import json
import sqlite3
import re
from typing import Dict, Any, List
from app.core.config import settings

DB_FILE = settings.DATABASE_PATH

DATABASE_SCHEMA_DESCRIPTION = """
Available Tables:
1. `project_items`: ticket_id, title, description, status, priority, assignee, module, item_type ('task', 'bug', 'change_request')
2. `teams_messages`: message_id, thread_id, sender_name, timestamp, content, has_blocker, is_technical_signal
3. `specs_fts`: doc_title, category, content (BM25 full-text search)
"""

def execute_sql_query(query: str) -> str:
    """Execute a safe, read-only SELECT query against the SQLite database."""
    clean_q = query.strip().rstrip(';')
    forbidden = ["insert", "update", "delete", "drop", "alter", "create", "truncate", "replace"]
    tokens = re.findall(r'\b[a-zA-Z]+\b', clean_q.lower())
    if any(f in tokens for f in forbidden):
        return json.dumps({"error": "Only read-only SELECT queries are permitted."})
    if not clean_q.lower().startswith("select") and not clean_q.lower().startswith("with"):
        return json.dumps({"error": "Query must start with SELECT or WITH."})

    try:
        with sqlite3.connect(DB_FILE) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(clean_q)
            rows = cursor.fetchall()
            results = [dict(r) for r in rows[:20]]  # Free tier compact ceiling
            return json.dumps({"row_count": len(rows), "data": results}, default=str)
    except Exception as e:
        return json.dumps({"error": f"SQL error: {str(e)}"})

def search_project_records(keywords: str, limit: int = 5) -> str:
    """Full-text BM25 search across indexed specifications, UI mockups, and Teams discussions."""
    try:
        clean_kw = re.sub(r'[^a-zA-Z0-9 ]', ' ', keywords).strip()
        terms = [t for t in clean_kw.split() if len(t) > 2]
        if not terms:
            return json.dumps({"records": []})
        fts_query = " OR ".join(f'"{t}"*' for t in terms[:4])

        results = []
        with sqlite3.connect(DB_FILE) as conn:
            conn.row_factory = sqlite3.Row
            c = conn.cursor()
            # Search specs
            try:
                c.execute("SELECT doc_title, snippet(specs_fts, 2, '<b>', '</b>', '...', 25) as snippet FROM specs_fts WHERE specs_fts MATCH ? LIMIT ?", (fts_query, limit))
                for r in c.fetchall():
                    results.append({"type": "spec", "citation": f"[Doc: {r['doc_title']}]", "text": r["snippet"]})
            except Exception:
                pass
            # Search teams messages
            try:
                c.execute("SELECT sender, timestamp, content FROM teams_fts WHERE teams_fts MATCH ? LIMIT ?", (fts_query, limit))
                for r in c.fetchall():
                    results.append({"type": "teams", "citation": f"[Teams: {r['sender']} @ {r['timestamp'][:10]}]", "text": r["content"][:250]})
            except Exception:
                pass

        return json.dumps({"total": len(results), "records": results[:limit]})
    except Exception as e:
        return json.dumps({"error": f"Search error: {str(e)}"})

AGENT_TOOLS_DEFINITIONS = [
    {
        "type": "function",
        "function": {
            "name": "execute_sql_query",
            "description": f"Execute a read-only SQL query against SQLite to count items, aggregate metrics, or filter by status or assignee. Schema:\n{DATABASE_SCHEMA_DESCRIPTION}",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "Valid SQLite SELECT query."}
                },
                "required": ["query"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "search_project_records",
            "description": "Full-text search across specifications, UI mockups, and stakeholder communications. Use for qualitative questions, locating tabs, buttons, or design rules.",
            "parameters": {
                "type": "object",
                "properties": {
                    "keywords": {"type": "string", "description": "Search keywords."}
                },
                "required": ["keywords"]
            }
        }
    }
]

TOOL_DISPATCHER = {
    "execute_sql_query": execute_sql_query,
    "search_project_records": search_project_records
}
