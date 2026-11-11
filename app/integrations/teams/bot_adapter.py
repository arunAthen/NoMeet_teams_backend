import json
from typing import Dict, Any
from app.agents.delivery_agent import run_agent_query
from app.integrations.teams.adaptive_cards import create_teams_evidence_card

async def handle_teams_bot_activity(activity: Dict[str, Any]) -> Dict[str, Any]:
    """
    Handles 1:1 Private Bot activities from Microsoft Teams.
    Each user gets a private isolated session keyed by their Microsoft Entra ID.
    """
    activity_type = activity.get("type", "message")
    if activity_type != "message":
        return {"status": "ignored", "reason": f"Activity type {activity_type} not processed"}

    # Extract user identity from Teams Activity
    from_user = activity.get("from", {})
    user_aad_id = from_user.get("aadObjectId") or from_user.get("id", "teams_user")
    user_name = from_user.get("name", "Team Member")
    user_query = activity.get("text", "").strip()

    if not user_query:
        return {"status": "empty"}

    # Isolated session per team member (Zero group chat pollution)
    session_id = f"teams_{user_aad_id}"

    # Execute ReAct agent query
    result = run_agent_query(user_query=user_query, session_id=session_id, user_name=user_name)

    # Wrap in Teams Adaptive Card
    card_payload = create_teams_evidence_card(
        title="Delivery Intelligence",
        response_text=result["response"],
        evidence_trail=result.get("evidence_trail", [])
    )

    return {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": card_payload
            }
        ],
        "session_id": session_id,
        "evaluation": result.get("evaluation", {})
    }
