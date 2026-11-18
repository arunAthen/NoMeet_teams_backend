import logging
import sqlite3
import httpx
from datetime import datetime, timezone
from typing import Dict, Any, Optional, List
from urllib.parse import unquote
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

async def discover_tenant_workspaces() -> List[Dict[str, Any]]:
    """
    Dynamically discovers all Group Chats and Team Channels across the Microsoft 365 tenant.
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

            # 1. Discover all Group Chats
            try:
                chats_resp = await client.get("https://graph.microsoft.com/v1.0/chats?$select=id,topic,chatType&$top=50", headers=graph_headers)
                if chats_resp.status_code == 200:
                    for ch in chats_resp.json().get("value", []):
                        cid = ch.get("id")
                        topic = ch.get("topic")
                        if not topic and "7bdcc" in cid.lower():
                            topic = "Intranet Solutions"
                        cname = topic or f"Chat {cid.split('@')[0].replace('19:', '')[:8].upper()}"
                        discovered.append({"id": cid, "name": cname, "workspace_type": "chat"})
                        with sqlite3.connect(DB_FILE) as conn:
                            conn.execute("""
                                INSERT INTO workspaces (id, name, workspace_type, last_active)
                                VALUES (?, ?, 'chat', ?)
                                ON CONFLICT(id) DO UPDATE SET name = excluded.name, last_active = excluded.last_active
                            """, (cid, cname, now))
            except Exception as ce:
                logger.warning(f"Chats discovery notice: {ce}")

            # 2. Discover all Teams & Channels
            try:
                teams_resp = await client.get("https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName&$top=20", headers=graph_headers)
                if teams_resp.status_code == 200:
                    for t in teams_resp.json().get("value", []):
                        tid = t.get("id")
                        tname = t.get("displayName")
                        chans_resp = await client.get(f"https://graph.microsoft.com/v1.0/teams/{tid}/channels?$select=id,displayName&$top=20", headers=graph_headers)
                        if chans_resp.status_code == 200:
                            for ch in chans_resp.json().get("value", []):
                                ch_id = ch.get("id")
                                ch_name = ch.get("displayName")
                                full_label = f"{tname} — {ch_name}"
                                discovered.append({"id": ch_id, "name": full_label, "workspace_type": "channel"})
                                with sqlite3.connect(DB_FILE) as conn:
                                    conn.execute("""
                                        INSERT INTO workspaces (id, name, workspace_type, last_active)
                                        VALUES (?, ?, 'channel', ?)
                                        ON CONFLICT(id) DO UPDATE SET name = excluded.name, last_active = excluded.last_active
                                    """, (ch_id, full_label, now))
            except Exception as te:
                logger.warning(f"Teams channels discovery notice: {te}")

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
                    c.execute("""
                        INSERT OR REPLACE INTO teams_messages (project_id, message_id, thread_id, sender_name, timestamp, content, is_technical_signal)
                        VALUES (?, ?, ?, 'System Roster', ?, ?, 1)
                    """, (clean_chat_id, f"roster_{clean_chat_id}", clean_chat_id, now, members_text))

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
                        c.execute("""
                            INSERT OR IGNORE INTO teams_messages (project_id, message_id, thread_id, sender_name, timestamp, content, is_technical_signal)
                            VALUES (?, ?, ?, ?, ?, ?, 1)
                        """, (clean_chat_id, msg_id, clean_chat_id, sender, created_time, clean_content))
                        try:
                            c.execute("""
                                INSERT INTO teams_fts (message_id, sender, timestamp, content)
                                VALUES (?, ?, ?, ?)
                            """, (msg_id, sender, created_time, clean_content))
                        except Exception:
                            pass
                        synced_count += 1
                conn.commit()

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
