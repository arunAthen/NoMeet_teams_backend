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
            topic_resp = await client.get(f"https://graph.microsoft.com/v1.0/chats/{clean_chat_id}", headers=graph_headers)
            clean_name = f"Chat {clean_chat_id.split('@')[0].replace('19:', '')[:8].upper()}"
            if topic_resp.status_code == 200:
                topic_data = topic_resp.json()
                if topic_data.get("topic"):
                    clean_name = topic_data.get("topic")

            # 2. Fetch Chat / Channel Messages from Graph
            if "thread.tacv2" in clean_chat_id:
                # Team Channel message endpoint
                msgs_url = f"https://graph.microsoft.com/v1.0/chats/{clean_chat_id}/messages?$top=50"
            else:
                # Group Chat message endpoint
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
                # Update workspace name
                now = datetime.now(timezone.utc).isoformat()
                c.execute("""
                    INSERT INTO workspaces (id, name, workspace_type, last_active)
                    VALUES (?, ?, 'chat', ?)
                    ON CONFLICT(id) DO UPDATE SET name = excluded.name, last_active = excluded.last_active
                """, (clean_chat_id, clean_name, now))

                for msg in messages_data:
                    msg_id = msg.get("id")
                    body = msg.get("body", {}).get("content", "")
                    sender = msg.get("from", {}).get("user", {}).get("displayName", "Team Member")
                    created_time = msg.get("createdDateTime", now)

                    if not body or "<attachment" in body.lower():
                        continue

                    # Clean html tags from Graph message body
                    import re
                    clean_content = re.sub(r"<[^>]+>", " ", body).strip()
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
