import logging
import httpx
from typing import Dict, Any, Optional
from app.core.config import settings

logger = logging.getLogger(__name__)

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
                # Test a basic Graph call to verify permissions
                graph_headers = {"Authorization": f"Bearer {access_token}"}
                graph_check = await client.get("https://graph.microsoft.com/v1.0/organization", headers=graph_headers)
                
                return {
                    "status": "connected",
                    "authenticated": True,
                    "tenant_id": tenant_id,
                    "client_id": client_id,
                    "token_type": data.get("token_type"),
                    "expires_in": data.get("expires_in"),
                    "graph_api_status": "accessible" if graph_check.status_code == 200 else f"limited_permissions ({graph_check.status_code})"
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
