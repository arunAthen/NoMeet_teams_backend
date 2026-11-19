import logging
import sqlite3
import httpx
from datetime import datetime, timezone
from typing import Dict, Any, Optional, List
from urllib.parse import unquote
from pathlib import Path
from app.core.config import settings

logger = logging.getLogger(__name__)

DB_FILE = settings.DATABASE_PATH

async def get_azure_token_diagnostics() -> Dict[str, Any]:
    """
    Validates Azure Entra ID credentials and acquires a Microsoft Graph application token.
    Returns status and diagnostic details.
    """
    client_id = settings.AZURE_CLIENT_ID
    client_secret = settings.AZURE_CLIENT_SECRET
    tenant_id = settings.AZURE_TENANT_ID

    if not client_id or not client_secret or not tenant_id:
        return {
            "status": "unconfigured",
            "message": "Azure credentials missing. Please set AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, and AZURE_TENANT_ID.",
            "client_id_present": bool(client_id),
            "tenant_id_present": bool(tenant_id),
            "client_secret_present": bool(client_secret)
        }

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }

    try:
        async with httpx.AsyncClient(timeout=15.0) as client:
            resp = await client.post(token_url, data=payload)
            data = resp.json()

            if resp.status_code == 200:
                access_token = data.get("access_token")
                graph_headers = {"Authorization": f"Bearer {access_token}"}
                
                # Check 1: Organization metadata
                graph_check = await client.get("https://graph.microsoft.com/v1.0/organization", headers=graph_headers)
                
                # Check 2: Chat permissions check on configured chat ID
                chat_check_status = "not_tested"
                chat_error_detail = None
                chat_topic = None
                
                target_chat = settings.AUTHORIZED_GROUP_CHAT_IDS[0] if settings.AUTHORIZED_GROUP_CHAT_IDS else None
                if target_chat:
                    clean_chat_id = unquote(target_chat).strip()
                    chat_resp = await client.get(f"https://graph.microsoft.com/v1.0/chats/{clean_chat_id}", headers=graph_headers)
                    chat_check_status = f"{chat_resp.status_code}"
                    if chat_resp.status_code == 200:
                        chat_data = chat_resp.json()
                        chat_topic = chat_data.get("topic")
                    else:
                        chat_error_detail = chat_resp.json()

                admin_consent_needed = (graph_check.status_code == 403 or chat_check_status == "403")

                return {
                    "status": "connected",
                    "authenticated": True,
                    "tenant_id": tenant_id,
                    "client_id": client_id,
                    "token_type": data.get("token_type"),
                    "expires_in": data.get("expires_in"),
                    "admin_consent_granted": not admin_consent_needed,
                    "graph_org_status": graph_check.status_code,
                    "chat_probe_status": chat_check_status,
                    "chat_topic_resolved": chat_topic,
                    "chat_error_detail": chat_error_detail
                }
            else:
                return {
                    "status": "error",
                    "authenticated": False,
                    "status_code": resp.status_code,
                    "error": data.get("error"),
                    "error_description": data.get("error_description")
                }
    except Exception as e:
        logger.error(f"Azure token validation error: {e}")
        return {
            "status": "exception",
            "authenticated": False,
            "error": str(e)
        }

async def discover_tenant_workspaces(user_id: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Dynamically discovers all Group Chats across the Microsoft 365 tenant using /users/{user_id}/chats.
    Auto-populates the workspaces table.
    """
    client_id = settings.AZURE_CLIENT_ID
    client_secret = settings.AZURE_CLIENT_SECRET
    tenant_id = settings.AZURE_TENANT_ID

    if not client_id or not client_secret or not tenant_id:
        return []

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }

    discovered = []
    try:
        async with httpx.AsyncClient(timeout=15.0) as client:
            token_resp = await client.post(token_url, data=payload)
            if token_resp.status_code != 200:
                return []
            access_token = token_resp.json().get("access_token")
            graph_headers = {"Authorization": f"Bearer {access_token}"}
            now = datetime.now(timezone.utc).isoformat()

            # Target users to discover group chats from
            candidate_user_ids = []
            if user_id and user_id != "default_user" and not user_id.startswith("user_"):
                candidate_user_ids.append(user_id)

            # Discover users in tenant to query their group chats
            try:
                users_resp = await client.get("https://graph.microsoft.com/v1.0/users?$top=10", headers=graph_headers)
                if users_resp.status_code == 200:
                    for u in users_resp.json().get("value", []):
                        uid = u.get("id")
                        if uid and uid not in candidate_user_ids:
                            candidate_user_ids.append(uid)
            except Exception as ue:
                logger.warning(f"Users query notice: {ue}")

            # Query group chats for candidate users
            seen_cids = set()
            for cuid in candidate_user_ids[:6]:
                try:
                    c_resp = await client.get(f"https://graph.microsoft.com/v1.0/users/{cuid}/chats", headers=graph_headers)
                    if c_resp.status_code == 200:
                        for ch in c_resp.json().get("value", []):
                            cid = ch.get("id")
                            if not cid or cid in seen_cids:
                                continue
                            seen_cids.add(cid)
                            chat_type = ch.get("chatType")
                            topic = ch.get("topic")
                            if chat_type == "group" or topic:
                                if not topic:
                                    if "7bdcc" in cid.lower():
                                        topic = "Intranet Solutions"
                                    elif "2337de30" in cid.lower():
                                        topic = "AgentAI Bot"
                                cname = topic or f"Chat {cid.split('@')[0].replace('19:', '')[:8].upper()}"
                                discovered.append({"id": cid, "name": cname, "workspace_type": "chat"})
                                with sqlite3.connect(DB_FILE) as conn:
                                    conn.execute("""
                                        INSERT INTO workspaces (id, name, workspace_type, last_active)
                                        VALUES (?, ?, 'chat', ?)
                                        ON CONFLICT(id) DO UPDATE SET
                                            name = CASE WHEN excluded.name NOT LIKE 'Chat %' THEN excluded.name ELSE workspaces.name END,
                                            last_active = excluded.last_active
                                    """, (cid, cname, now))
                except Exception as c_err:
                    logger.warning(f"User {cuid} chats notice: {c_err}")

    except Exception as e:
        logger.warning(f"Workspace discovery exception: {e}")

    return discovered

async def sync_live_chat_messages(chat_id: str) -> Dict[str, Any]:
    """
    Fetches real-time messages from Microsoft Graph API for a given Teams chat.
    Strictly real live data: zero mock feeds.
    """
    client_id = settings.AZURE_CLIENT_ID
    client_secret = settings.AZURE_CLIENT_SECRET
    tenant_id = settings.AZURE_TENANT_ID

    if not client_id or not client_secret or not tenant_id:
        return {"status": "error", "message": "Azure credentials missing."}

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }

    clean_chat_id = unquote(chat_id).strip()

    try:
        async with httpx.AsyncClient(timeout=20.0) as client:
            token_resp = await client.post(token_url, data=payload)
            if token_resp.status_code != 200:
                return {"status": "auth_error", "details": token_resp.json()}
            
            access_token = token_resp.json().get("access_token")
            graph_headers = {"Authorization": f"Bearer {access_token}"}

            # 1. Fetch Chat Topic / Name
            clean_name = "Intranet Solutions" if "7bdcc" in clean_chat_id.lower() else f"Chat {clean_chat_id.split('@')[0].replace('19:', '')[:8].upper()}"
            try:
                topic_resp = await client.get(f"https://graph.microsoft.com/v1.0/chats/{clean_chat_id}", headers=graph_headers)
                if topic_resp.status_code == 200:
                    topic_data = topic_resp.json()
                    if topic_data.get("topic"):
                        clean_name = topic_data.get("topic")
            except Exception as te:
                logger.warning(f"Topic fetch notice: {te}")

            # 2. Fetch Chat Members
            members_list = []
            try:
                mem_resp = await client.get(f"https://graph.microsoft.com/v1.0/chats/{clean_chat_id}/members", headers=graph_headers)
                if mem_resp.status_code == 200:
                    members_list = [m.get("displayName") for m in mem_resp.json().get("value", []) if m.get("displayName")]
            except Exception as me:
                logger.warning(f"Members fetch notice: {me}")

            # 3. Fetch Chat / Channel Messages from Graph
            msgs_url = f"https://graph.microsoft.com/v1.0/chats/{clean_chat_id}/messages?$top=50"
            msgs_resp = await client.get(msgs_url, headers=graph_headers)

            if msgs_resp.status_code != 200:
                return {
                    "status": "graph_error",
                    "status_code": msgs_resp.status_code,
                    "details": msgs_resp.json(),
                    "admin_consent_hint": "Admin consent for 'Chat.Read.All' / 'ChannelMessage.Read.All' is required in Azure Portal."
                }

            messages_data = msgs_resp.json().get("value", [])
            synced_count = 0

            # Store live messages into SQLite
            with sqlite3.connect(DB_FILE) as conn:
                c = conn.cursor()
                now = datetime.now(timezone.utc).isoformat()

                # Update workspace name & metadata
                try:
                    c.execute("""
                        INSERT INTO workspaces (id, name, workspace_type, last_active)
                        VALUES (?, ?, 'chat', ?)
                        ON CONFLICT(id) DO UPDATE SET name = excluded.name, last_active = excluded.last_active
                    """, (clean_chat_id, clean_name, now))
                    conn.commit()
                except Exception as we:
                    logger.warning(f"Failed to update workspace record: {we}")

                # Store members as a system reference record if available
                if members_list:
                    members_text = f"Group chat members ({len(members_list)} total): " + ", ".join(members_list)
                    try:
                        c.execute("""
                            INSERT OR REPLACE INTO teams_messages (project_id, message_id, thread_id, sender_name, timestamp, content, has_attachments, is_technical_signal, has_blocker)
                            VALUES (?, ?, ?, 'System Roster', ?, ?, 0, 1, 0)
                        """, (clean_chat_id, f"roster_{clean_chat_id}", clean_chat_id, now, members_text))
                    except Exception as re_err:
                        logger.warning(f"Roster record notice: {re_err}")

                import re
                import html

                for msg in messages_data:
                    if not isinstance(msg, dict):
                        continue
                    msg_id = msg.get("id")
                    if not msg_id:
                        continue

                    # Safe body extraction
                    body_obj = msg.get("body")
                    body = ""
                    if isinstance(body_obj, dict):
                        body = body_obj.get("content") or ""

                    # Safe sender extraction
                    from_obj = msg.get("from")
                    sender = "Team Member"
                    if isinstance(from_obj, dict):
                        user_info = from_obj.get("user") or from_obj.get("application") or {}
                        if isinstance(user_info, dict):
                            sender = user_info.get("displayName") or "Team Member"
                        elif isinstance(user_info, str):
                            sender = user_info

                    created_time = msg.get("createdDateTime") or now

                    # Collect text parts
                    text_parts = []
                    if body and "<attachment" not in body.lower():
                        clean_body = html.unescape(re.sub(r"<[^>]+>", " ", body)).strip()
                        if clean_body:
                            text_parts.append(clean_body)

                    # Extract Adaptive Card JSON text from attachments (e.g. Workflows, leave requests)
                    attachments = msg.get("attachments") or []
                    for att in attachments:
                        if isinstance(att, dict):
                            att_content = att.get("content") or ""
                            if isinstance(att_content, str):
                                card_texts = re.findall(r'"(?:text|title|value|fallback)":\s*"([^"]+)"', att_content)
                                if card_texts:
                                    text_parts.append(html.unescape(" ".join(card_texts)))
                                elif not att_content.strip().startswith("{") and att_content.strip():
                                    text_parts.append(html.unescape(re.sub(r"<[^>]+>", " ", att_content)).strip())
                            att_name = att.get("name")
                            if att_name:
                                text_parts.append(f"[File: {att_name}]")

                    # If system event message (e.g. title changed, user left)
                    if not text_parts:
                        event_detail = msg.get("eventDetail")
                        if isinstance(event_detail, dict):
                            ev_title = event_detail.get("chatTitle") or event_detail.get("topic")
                            if ev_title:
                                text_parts.append(f"Group name set to {ev_title}")

                    clean_content = " ".join(text_parts).strip()
                    if clean_content:
                        has_att = 1 if attachments else 0
                        try:
                            c.execute("""
                                INSERT OR IGNORE INTO teams_messages (project_id, message_id, thread_id, sender_name, timestamp, content, has_attachments, is_technical_signal, has_blocker)
                                VALUES (?, ?, ?, ?, ?, ?, ?, 1, 0)
                            """, (clean_chat_id, msg_id, clean_chat_id, sender, created_time, clean_content, has_att))
                            c.execute("""
                                INSERT INTO teams_fts (message_id, sender, timestamp, content)
                                VALUES (?, ?, ?, ?)
                            """, (msg_id, sender, created_time, clean_content))
                            synced_count += 1
                        except Exception as ins_err:
                            logger.warning(f"Message insert notice: {ins_err}")
                conn.commit()

            # 3. Dynamic Live Files Sync from Microsoft Teams Chat Files
            try:
                candidate_uploaders = ["devadmin@athen.club"]
                for msg in messages_data[:10]:
                    if isinstance(msg, dict):
                        f_info = msg.get("from") or {}
                        if isinstance(f_info, dict):
                            u_data = f_info.get("user") or {}
                            if isinstance(u_data, dict):
                                upn = u_data.get("userPrincipalName") or u_data.get("id")
                                if upn and upn not in candidate_uploaders:
                                    candidate_uploaders.append(upn)

                settings.DATA_SYNC_DIR.mkdir(parents=True, exist_ok=True)
                newly_downloaded = 0
                for u_id in candidate_uploaders[:3]:
                    try:
                        files_url = f"https://graph.microsoft.com/v1.0/users/{u_id}/drive/root:/Microsoft Teams Chat Files:/children"
                        f_resp = await client.get(files_url, headers=graph_headers)
                        if f_resp.status_code == 200:
                            for item in f_resp.json().get("value", []):
                                fname = item.get("name")
                                fid = item.get("id")
                                fsize = item.get("size", 0)
                                if fname and fid and any(fname.lower().endswith(ext) for ext in [".docx", ".xlsx", ".pdf", ".html", ".png", ".jpg"]):
                                    target_file = settings.DATA_SYNC_DIR / fname
                                    if not target_file.exists() or target_file.stat().st_size != fsize:
                                        c_resp = await client.get(
                                            f"https://graph.microsoft.com/v1.0/users/{u_id}/drive/items/{fid}/content",
                                            headers=graph_headers,
                                            follow_redirects=True
                                        )
                                        if c_resp.status_code == 200:
                                            target_file.write_bytes(c_resp.content)
                                            newly_downloaded += 1
                                            logger.info(f"Downloaded new Teams file: {fname}")
                    except Exception as fe:
                        logger.warning(f"Error checking files for user {u_id}: {fe}")

                if newly_downloaded > 0:
                    from app.ingestion.dynamic_crawler import crawl_and_ingest_directory
                    crawl_and_ingest_directory(settings.DATA_SYNC_DIR)
            except Exception as shared_err:
                logger.warning(f"Live files sync warning: {shared_err}")

            return {
                "status": "success",
                "chat_id": clean_chat_id,
                "chat_name": clean_name,
                "synced_messages_count": synced_count,
                "total_fetched": len(messages_data)
            }

    except Exception as e:
        logger.error(f"Live chat sync exception: {e}")
        return {"status": "exception", "error": str(e)}
