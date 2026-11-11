from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import Response

class TeamsSecurityHeadersMiddleware(BaseHTTPMiddleware):
    """Allows Microsoft Teams client to embed the personal tab inside secure iframes."""
    async def dispatch(self, request: Request, call_next):
        response: Response = await call_next(request)
        # Allow embedding in Teams clients across desktop, web, and mobile
        response.headers["Content-Security-Policy"] = (
            "frame-ancestors 'self' teams.microsoft.com *.teams.microsoft.com "
            "*.skype.com *.sharepoint.com *.office.com *.microsoft365.com;"
        )
        # Remove restrictive legacy header if present
        if "X-Frame-Options" in response.headers:
            del response.headers["X-Frame-Options"]
        return response
