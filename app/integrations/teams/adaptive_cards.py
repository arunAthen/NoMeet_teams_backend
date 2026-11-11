import json
from typing import Dict, Any, List

def create_teams_evidence_card(title: str, response_text: str, evidence_trail: List[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Generates an official Microsoft Teams Adaptive Card v1.5 with citations and action buttons."""
    facts = []
    if evidence_trail:
        for ev in evidence_trail[:3]:
            tool_name = ev.get("tool", "Evidence")
            facts.append({"title": tool_name, "value": str(ev.get("output", ""))[:90] + "..."})

    card = {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.5",
        "body": [
            {
                "type": "Container",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "⚡ M365 Delivery & Operations Intelligence",
                        "weight": "Bolder",
                        "size": "Medium",
                        "color": "Accent"
                    },
                    {
                        "type": "TextBlock",
                        "text": response_text,
                        "wrap": True,
                        "spacing": "Medium"
                    }
                ]
            }
        ]
    }
    if facts:
        card["body"].append({
            "type": "FactSet",
            "facts": facts,
            "spacing": "Medium"
        })
    return card
