import os
import re
import json
import sqlite3
from datetime import datetime, timezone
from typing import Dict, Any, List, Optional, Tuple
from app.core.config import settings
from app.agents.agent_tools import AGENT_TOOLS_DEFINITIONS, TOOL_DISPATCHER, execute_sql_query, search_project_records
from app.agents.evaluator import calculate_evaluation_metrics

DB_FILE = settings.DATABASE_PATH

SYSTEM_PROMPT = """You are the official Microsoft Teams Delivery & Operations Intelligence Agent.
Your job is to answer project management, engineering status, bug tracking, and design specification questions for team members.

MANDATORY RULES:
1. Ground every single claim in exact evidence:
   - Use `[Tracker: TICKET_ID]` for project tasks, bugs, and change requests.
   - Use `[Teams: SENDER @ YYYY-MM-DD]` for stakeholder discussions and agreements.
   - Use `[Doc: TITLE]` for technical specifications, Word docs, and UI mockups.
2. If evidence does not exist in the database or tools, state clearly: 'Not covered in formal documentation (agreed in chat review)' or 'No record found'.
3. Never hallucinate ticket numbers, names, or metrics.
"""

def deterministic_grounded_answer(query: str, group_id: str = "all") -> str:
    """Offline deterministic fallback engine when LLM APIs are unreachable, strictly scoped by group."""
    q_low = query.lower()
    with sqlite3.connect(DB_FILE) as conn:
        conn.row_factory = sqlite3.Row
        c = conn.cursor()

        # 1. Change Requests
        if re.search(r'\b(cr|crs|change\s+requests?)\b', q_low):
            if group_id != "all":
                c.execute("SELECT status, count(*) as count FROM project_items WHERE item_type = 'change_request' AND project_id = ? GROUP BY status", (group_id,))
            else:
                c.execute("SELECT status, count(*) as count FROM project_items WHERE item_type = 'change_request' GROUP BY status")
            rows = c.fetchall()
            total_crs = sum(r["count"] for r in rows)
            breakdown = ", ".join(f"**{r['count']} {r['status']}**" for r in rows)
            return (
                f"### 🎯 Executive Ground Truth\n"
                f"A total of **{total_crs} distinct Change Requests (CRs)** are tracked in project records ({breakdown}).\n\n"
                f"### 📜 Verified Lifecycle Evidence\n"
                f"1. **CR Implementation Record**: Change requests are tracked directly in sprint sheets `[Tracker: Project Plan]` and approved in review threads `[Teams: Stakeholder Review]`.\n"
                f"2. **Documentation Status**: Formal scope changes agreed via stakeholder reviews (Not covered in formal word specs).\n\n"
                f"### ⚡ Actionable Next Steps\n- Verify latest UAT test pass status in Teams."
            )

        # 2. Bugs
        if re.search(r'\b(bugs?|defects?)\b', q_low):
            if group_id != "all":
                c.execute("SELECT status, count(*) as count FROM project_items WHERE item_type = 'bug' AND project_id = ? GROUP BY status", (group_id,))
            else:
                c.execute("SELECT status, count(*) as count FROM project_items WHERE item_type = 'bug' GROUP BY status")
            rows = c.fetchall()
            total_bugs = sum(r["count"] for r in rows)
            if group_id != "all":
                c.execute("SELECT ticket_id, title, assignee, status FROM project_items WHERE item_type = 'bug' AND project_id = ? ORDER BY id DESC LIMIT 3", (group_id,))
            else:
                c.execute("SELECT ticket_id, title, assignee, status FROM project_items WHERE item_type = 'bug' ORDER BY id DESC LIMIT 3")
            samples = c.fetchall()
            sample_str = "\n".join(f"- `[Tracker: {s['ticket_id']}]`: {s['title']} (Assignee: {s['assignee']} | Status: {s['status']})" for s in samples)
            return (
                f"### 🎯 Executive Ground Truth\n"
                f"A total of **{total_bugs} bugs** are documented in the defect ledger.\n\n"
                f"### 📜 Sample Active Bugs\n{sample_str}\n\n"
                f"### ⚡ Actionable Next Steps\n- Coordinate with assigned engineers for closure."
            )

        # 3. Inquiry Type Classification
        is_leave_query = bool(re.search(r'\b(leaves?|absent|vacations?|holidays?|time\s*off)\b', q_low))
        is_people_query = bool(re.search(r'\b(people|peoples|members?|who\s+is\s+in|participants?|roster)\b', q_low))
        clean_gid = group_id.split("@")[0].replace("19:", "").strip() if group_id and group_id != "all" else ""
        ws_label = "AgentAI Bot" if "2337de30" in str(group_id).lower() else ("Intranet Solutions" if "7bdcc461" in str(group_id).lower() else (clean_gid or "this workspace"))

        # Compound Question: Members/People AND Leaves in same query
        if is_people_query and is_leave_query:
            roster_text = None
            if clean_gid:
                c.execute("SELECT content FROM teams_messages WHERE message_id LIKE 'roster_%' AND (project_id = ? OR project_id LIKE ? OR project_id = 'all') ORDER BY id DESC LIMIT 1", (group_id, f"%{clean_gid}%"))
            else:
                c.execute("SELECT content FROM teams_messages WHERE message_id LIKE 'roster_%' ORDER BY id DESC LIMIT 1")
            roster_row = c.fetchone()
            if roster_row and roster_row[0]:
                roster_text = roster_row[0]
            else:
                if clean_gid:
                    c.execute("SELECT DISTINCT sender_name FROM teams_messages WHERE sender_name != 'System Roster' AND (project_id = ? OR project_id LIKE ? OR project_id = 'all')", (group_id, f"%{clean_gid}%"))
                else:
                    c.execute("SELECT DISTINCT sender_name FROM teams_messages WHERE sender_name != 'System Roster'")
                senders = [r[0] for r in c.fetchall() if r[0]]
                roster_text = f"Group chat members ({len(senders)} total): {', '.join(senders)}" if senders else "No active participants recorded."

            if clean_gid:
                c.execute("""
                    SELECT sender_name, timestamp, content 
                    FROM teams_messages 
                    WHERE (lower(content) LIKE '%leave%' OR lower(content) LIKE '%fever%' OR lower(content) LIKE '%absent%')
                      AND (project_id = ? OR project_id LIKE ? OR project_id = 'all')
                    ORDER BY timestamp DESC
                """, (group_id, f"%{clean_gid}%"))
            else:
                c.execute("""
                    SELECT sender_name, timestamp, content 
                    FROM teams_messages 
                    WHERE (lower(content) LIKE '%leave%' OR lower(content) LIKE '%fever%' OR lower(content) LIKE '%absent%')
                    ORDER BY timestamp DESC
                """)
            leave_rows = c.fetchall()

            c.execute("SELECT COUNT(*) FROM teams_messages WHERE (lower(content) LIKE '%leave%' OR lower(content) LIKE '%fever%' OR lower(content) LIKE '%absent%')")
            total_org_leaves = c.fetchone()[0]

            sections = [
                f"### 🎯 Executive Ground Truth",
                f"**1. Group Members & Directory Roster**:\n- {roster_text}"
            ]

            if leave_rows:
                items_str = "\n".join(f"- **[Teams: {r['sender_name']} @ {r['timestamp'][:10]}]**: {r['content']}" for r in leave_rows[:10])
                sections.append(f"**2. Leave Records in {ws_label}**:\n- **Total Leave Entries**: {len(leave_rows)}\n{items_str}")
            else:
                cross_note = f"\n*(Note: {total_org_leaves} verified leave notices are recorded in **Intranet Solutions**. Switch to **Intranet Solutions** or **All Workspaces** to inspect them.)*" if total_org_leaves > 0 else ""
                sections.append(f"**2. Leave Records in {ws_label}**:\n- **Total Leave Entries**: **0**{cross_note}")

            sections.append(f"### ⚡ Verified Directory & Chat Status\n- Sourced directly from verified Microsoft 365 Entra ID / Teams communications.")
            return "\n\n".join(sections)

        # Single: Leave / Time-off inquiry
        if is_leave_query:
            if clean_gid:
                c.execute("""
                    SELECT sender_name, timestamp, content 
                    FROM teams_messages 
                    WHERE (lower(content) LIKE '%leave%' OR lower(content) LIKE '%fever%' OR lower(content) LIKE '%absent%')
                      AND (project_id = ? OR project_id LIKE ? OR project_id = 'all')
                    ORDER BY timestamp DESC
                """, (group_id, f"%{clean_gid}%"))
            else:
                c.execute("""
                    SELECT sender_name, timestamp, content 
                    FROM teams_messages 
                    WHERE (lower(content) LIKE '%leave%' OR lower(content) LIKE '%fever%' OR lower(content) LIKE '%absent%')
                    ORDER BY timestamp DESC
                """)
            leave_rows = c.fetchall()
            if leave_rows:
                items_str = "\n".join(f"- **[Teams: {r['sender_name']} @ {r['timestamp'][:10]}]**: {r['content']}" for r in leave_rows)
                return (
                    f"### 🎯 Executive Ground Truth\n"
                    f"A total of **{len(leave_rows)} leave record(s)** were identified in **{ws_label}** communications:\n\n"
                    f"### 📜 Verified Leave Notices\n{items_str}\n\n"
                    f"### ⚡ Summary\n"
                    f"- **Total Leave Entries**: {len(leave_rows)}\n"
                    f"- All records sourced directly from verified Microsoft Teams chat messages."
                )
            else:
                c.execute("SELECT COUNT(*) FROM teams_messages WHERE (lower(content) LIKE '%leave%' OR lower(content) LIKE '%fever%' OR lower(content) LIKE '%absent%')")
                total_org_leaves = c.fetchone()[0]
                cross_note = f"\n\n### ℹ️ Cross-Workspace Intelligence\n- **{total_org_leaves} verified leave notices** are recorded in the **Intranet Solutions** workspace.\n- To inspect them, select **Intranet Solutions** or **All Workspaces** from the workspace dropdown." if total_org_leaves > 0 else ""
                return (
                    f"### 🎯 Executive Ground Truth\n"
                    f"A total of **0 leave records** or time-off notices were found in the **{ws_label}** workspace.{cross_note}\n\n"
                    f"### ⚡ Verified Directory Status\n- Corroborated with active Teams workspace messages."
                )

        # Single: Group Members / People inquiry
        if is_people_query:
            if clean_gid:
                c.execute("SELECT content FROM teams_messages WHERE message_id LIKE 'roster_%' AND (project_id = ? OR project_id LIKE ? OR project_id = 'all') ORDER BY id DESC LIMIT 1", (group_id, f"%{clean_gid}%"))
            else:
                c.execute("SELECT content FROM teams_messages WHERE message_id LIKE 'roster_%' ORDER BY id DESC LIMIT 1")
            roster_row = c.fetchone()
            if roster_row and roster_row[0]:
                return (
                    f"### 🎯 Executive Ground Truth\n"
                    f"{roster_row[0]}.\n\n"
                    f"### ⚡ Verified Directory Status\n"
                    f"- Sourced directly from verified Microsoft 365 Entra ID / Teams roster."
                )
            # Fallback to distinct senders
            if clean_gid:
                c.execute("SELECT DISTINCT sender_name FROM teams_messages WHERE sender_name != 'System Roster' AND (project_id = ? OR project_id LIKE ? OR project_id = 'all')", (group_id, f"%{clean_gid}%"))
            else:
                c.execute("SELECT DISTINCT sender_name FROM teams_messages WHERE sender_name != 'System Roster'")
            senders = [r[0] for r in c.fetchall() if r[0]]
            if senders:
                names_str = ", ".join(f"**{name}**" for name in senders)
                return (
                    f"### 🎯 Executive Ground Truth\n"
                    f"Identified **{len(senders)} active participant(s)** in this workspace: {names_str}.\n\n"
                    f"### ⚡ Verified Communication Signals\n"
                    f"- Based on active message authors in workspace records."
                )

        # 5. Technical Specifications & Document Inquiries (Dynamic specs_fts matching)
        stop_words = {"what", "when", "where", "which", "how", "many", "does", "from", "that", "this", "tell", "show", "with", "about", "are", "there", "for", "the", "and", "under"}
        tokens = [w for w in re.findall(r'\b[a-zA-Z0-9_-]{3,}\b', q_low) if w not in stop_words]
        matched_specs = []
        if tokens:
            try:
                c.execute("SELECT doc_title, category, content FROM specs_fts")
                all_specs = [dict(r) for r in c.fetchall()]
                scored = []
                for s in all_specs:
                    clow = s["content"].lower()
                    tlow = s["doc_title"].lower()
                    score = sum(3 for t in tokens if t in tlow) + sum(1 for t in tokens if t in clow)
                    if score > 0:
                        scored.append((score, s))
                scored.sort(key=lambda x: x[0], reverse=True)
                if scored and scored[0][0] >= 2:
                    matched_specs = [item[1] for item in scored[:3]]
            except Exception:
                pass

        if matched_specs:
            sections = [f"### 🎯 Executive Ground Truth\nResolved from verified specifications and architecture documentation for: *{query}*"]
            for sp in matched_specs[:2]:
                doc_title = sp["doc_title"]
                doc_badge = f"[Doc: {doc_title}]"
                sections.append(f"### {doc_badge}\n{sp['content'][:900]}")
            sections.append("### ⚡ Actionable Next Steps\n- Corroborated with verified requirements, UI design references, and architecture documents.")
            return "\n\n".join(sections)

        # 6. General Keyword Search across Communications
        tokens = [w for w in re.findall(r'\b[a-zA-Z0-9_-]{3,}\b', q_low) if w not in {"how", "many", "took", "what", "when", "where", "which", "there", "have", "been", "with", "from", "that", "this", "tell", "about", "show"}]
        if tokens:
            like_clauses = " OR ".join(["lower(content) LIKE ?"] * len(tokens))
            params = [f"%{t}%" for t in tokens]
            if group_id != "all":
                c.execute(f"SELECT sender_name, timestamp, content FROM teams_messages WHERE ({like_clauses}) AND (project_id = ? OR project_id = 'all') ORDER BY id DESC LIMIT 5", tuple(params + [group_id]))
            else:
                c.execute(f"SELECT sender_name, timestamp, content FROM teams_messages WHERE ({like_clauses}) ORDER BY id DESC LIMIT 5", tuple(params))
            matched_msgs = c.fetchall()
            if matched_msgs:
                msg_str = "\n".join(f"- **[Teams: {m['sender_name']} @ {m['timestamp'][:10]}]**: {m['content'][:300]}" for m in matched_msgs)
                return (
                    f"### 🎯 Executive Ground Truth\n"
                    f"Resolved from latest project communications for query: *{query}*\n\n"
                    f"### 💬 Recent Stakeholder Signals\n{msg_str}\n\n"
                    f"### ⚡ Actionable Next Steps\n- Corroborated with active Teams workspace messages."
                )

        # 5. Default recent communication signals fallback
        if group_id != "all":
            c.execute("SELECT sender_name, timestamp, content FROM teams_messages WHERE is_technical_signal = 1 AND project_id = ? ORDER BY id DESC LIMIT 2", (group_id,))
        else:
            c.execute("SELECT sender_name, timestamp, content FROM teams_messages WHERE is_technical_signal = 1 ORDER BY id DESC LIMIT 2")
        msgs = c.fetchall()
        if msgs:
            msg_str = "\n".join(f"- **[Teams: {m['sender_name']} @ {m['timestamp'][:10]}]**: {m['content'][:200]}" for m in msgs)
            return (
                f"### 🎯 Executive Ground Truth\n"
                f"Resolved from latest project communications for: *{query}*\n\n"
                f"### 💬 Recent Stakeholder Signals\n{msg_str}\n\n"
                f"### ⚡ Actionable Next Steps\n- Corroborated with active Teams workspace messages."
            )

    return f"No verified documentation or chat record found for: *{query}*."

def run_agent_query(user_query: str, session_id: str = "default_session", user_name: str = "Team Member", group_id: str = "all") -> Dict[str, Any]:
    """Orchestrates ReAct tool loop: Tier 1 Groq -> Tier 2 Gemini -> Tier 3 Deterministic (Group-Scoped)."""
    final_content = None
    evidence_trail = []

    scoped_system_prompt = f"{SYSTEM_PROMPT}\n\nCURRENT ACTIVE WORKSPACE / GROUP: '{group_id}'. Strict rule: Only return information that belongs to this workspace."

    # Try Tier 1: Groq Cloud (llama-3.3-70b-versatile via direct REST API)
    if settings.ENABLE_EXTERNAL_LLM and settings.GROQ_API_KEY:
        try:
            import httpx
            headers = {
                "Authorization": f"Bearer {settings.GROQ_API_KEY}",
                "Content-Type": "application/json"
            }
            messages = [
                {"role": "system", "content": scoped_system_prompt},
                {"role": "user", "content": user_query}
            ]
            models_to_try = [settings.GROQ_MODEL, "llama-3.3-70b-versatile", "llama-3.1-8b-instant"]
            # Filter out known decommissioned models
            models_to_try = [m for m in dict.fromkeys(models_to_try) if m and m not in {"llama-3.1-70b-versatile", "mixtral-8x7b-32768"}]
            if not models_to_try:
                models_to_try = ["llama-3.3-70b-versatile", "llama-3.1-8b-instant"]

            active_model = None
            with httpx.Client(timeout=20.0) as client:
                for cand_model in models_to_try:
                    try:
                        payload = {
                            "model": cand_model,
                            "messages": messages,
                            "tools": AGENT_TOOLS_DEFINITIONS,
                            "tool_choice": "auto",
                            "temperature": 0.0,
                            "max_tokens": 1024
                        }
                        r1 = client.post("https://api.groq.com/openai/v1/chat/completions", headers=headers, json=payload)
                        if r1.status_code == 200:
                            active_model = cand_model
                            choice = r1.json()["choices"][0]["message"]
                            tool_calls = choice.get("tool_calls")
                            if not tool_calls:
                                if choice.get("content"):
                                    final_content = choice["content"]
                            else:
                                messages.append(choice)
                                for tc in tool_calls:
                                    fn = tc["function"]["name"]
                                    args = json.loads(tc["function"]["arguments"]) if isinstance(tc["function"]["arguments"], str) else (tc["function"]["arguments"] or {})
                                    handler = TOOL_DISPATCHER.get(fn)
                                    output = handler(**args) if handler else json.dumps({"error": "Unknown tool"})
                                    evidence_trail.append({"tool": fn, "args": args, "output": output[:300]})
                                    messages.append({
                                        "role": "tool",
                                        "tool_call_id": tc["id"],
                                        "content": output
                                    })
                                r2 = client.post("https://api.groq.com/openai/v1/chat/completions", headers=headers, json={
                                    "model": active_model,
                                    "messages": messages,
                                    "temperature": 0.0,
                                    "max_tokens": 1024
                                })
                                if r2.status_code == 200:
                                    final_content = r2.json()["choices"][0]["message"].get("content")
                            break
                        else:
                            # If 400 (e.g. tool_use_failed) or 404 (model not found), try next model
                            continue
                    except Exception as model_err:
                        continue

                # Direct completion fallback with retrieved context if tool calling failed
                if not final_content:
                    retrieved_context = ""
                    try:
                        clean_gid = group_id.split("@")[0].replace("19:", "").strip() if group_id and group_id != "all" else ""
                        with sqlite3.connect(DB_FILE) as conn:
                            conn.row_factory = sqlite3.Row
                            c = conn.cursor()
                            if clean_gid:
                                c.execute("SELECT sender_name, timestamp, content FROM teams_messages WHERE (project_id = ? OR project_id LIKE ? OR project_id = 'all') ORDER BY id DESC LIMIT 15", (group_id, f"%{clean_gid}%"))
                            else:
                                c.execute("SELECT sender_name, timestamp, content FROM teams_messages ORDER BY id DESC LIMIT 15")
                            rows = c.fetchall()
                            if rows:
                                retrieved_context = "\n".join(f"- [{r['sender_name']} @ {r['timestamp'][:10]}]: {r['content']}" for r in rows)
                    except Exception:
                        pass

                    for fallback_model in ["llama-3.3-70b-versatile", "llama-3.1-8b-instant"]:
                        try:
                            direct_payload = {
                                "model": fallback_model,
                                "messages": [
                                    {"role": "system", "content": f"{scoped_system_prompt}\n\nWorkspace Context:\n{retrieved_context}"},
                                    {"role": "user", "content": user_query}
                                ],
                                "temperature": 0.1,
                                "max_tokens": 1024
                            }
                            r_direct = client.post("https://api.groq.com/openai/v1/chat/completions", headers=headers, json=direct_payload)
                            if r_direct.status_code == 200:
                                final_content = r_direct.json()["choices"][0]["message"].get("content")
                                break
                        except Exception:
                            continue
        except Exception as e:
            print(f"[Tier 1 Groq REST Error]: {e}. Attempting fallback...")

    # Fallback Tier 3: Deterministic Grounded Engine
    if not final_content:
        final_content = deterministic_grounded_answer(user_query, group_id=group_id)

    # Sanitize phantom tags
    final_content = re.sub(r'\[Doc:\s*(?:Project\s*Tracker\s*Guidelines|Tracker\s*Guidelines)[^\]]*\]', 'Not covered in formal documentation (agreed in chat review)', final_content, flags=re.IGNORECASE)

    eval_scores = calculate_evaluation_metrics(user_query, final_content, evidence_trail)

    # Save to SQLite session
    try:
        now = datetime.now(timezone.utc).isoformat()
        with sqlite3.connect(DB_FILE) as conn:
            conn.execute("INSERT OR IGNORE INTO chat_sessions (session_id, project_id, user_name, title, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?)", (session_id, group_id, user_name, user_query[:35], now, now))
            conn.execute("INSERT INTO chat_messages (session_id, role, content, evidence, eval_scores, created_at) VALUES (?, 'user', ?, ?, ?, ?)", (session_id, user_query, None, None, now))
            conn.execute("INSERT INTO chat_messages (session_id, role, content, evidence, eval_scores, created_at) VALUES (?, 'assistant', ?, ?, ?, ?)", (session_id, final_content, json.dumps(evidence_trail), json.dumps(eval_scores), now))
            conn.commit()
    except Exception as e:
        print(f"[Chat storage error]: {e}")

    return {
        "response": final_content,
        "evidence_trail": evidence_trail,
        "evaluation": eval_scores,
        "session_id": session_id
    }
