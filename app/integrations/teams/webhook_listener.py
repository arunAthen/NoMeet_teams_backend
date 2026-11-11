import logging
from typing import Dict, Any
from app.ingestion.teams_chat_chunker import upsert_teams_message

logger = logging.getLogger(__name__)

async def handle_graph_chat_webhook(payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    Receives live Microsoft Graph change notifications for created, updated (edited), or deleted chat messages.
    """
    notifications = payload.get("value", [])
    processed_count = 0

    for notif in notifications:
        change_type = notif.get("changeType")  # 'created', 'updated', 'deleted'
        resource_data = notif.get("resourceData", {})
        message_id = resource_data.get("id")

        if not message_id:
            continue

        if change_type in ("created", "updated"):
            # If notification contains encrypted/delta content or requires fetch:
            content = resource_data.get("body", {}).get("content", "")
            sender = resource_data.get("from", {}).get("user", {}).get("displayName", "Team Member")
            timestamp = resource_data.get("createdDateTime", "")
            
            if content:
                upsert_teams_message({
                    "message_id": str(message_id),
                    "thread_id": str(resource_data.get("chatId", message_id)),
                    "sender_name": sender,
                    "timestamp": timestamp,
                    "content": content,
                    "is_technical_signal": True
                })
                processed_count += 1
                logger.info(f"Processed live {change_type} event for message {message_id}")

    return {"status": "success", "processed": processed_count}
