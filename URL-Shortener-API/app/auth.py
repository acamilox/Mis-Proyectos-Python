import time
from collections import defaultdict

from fastapi import Request, Response
from starlette.middleware.base import BaseHTTPMiddleware, RequestResponseEndpoint

from .config import settings

EXEMPT_PATHS = {"/", "/docs", "/redoc", "/openapi.json"}

# Store: {ip: [timestamp, ...]}
_rate_limit_store: dict[str, list[float]] = defaultdict(list)


class AuthRateLimitMiddleware(BaseHTTPMiddleware):
    async def dispatch(
        self, request: Request, call_next: RequestResponseEndpoint
    ) -> Response:
        path = request.url.path

        # Rate limiting (skip for exempt paths)
        if path not in EXEMPT_PATHS:
            client_ip = request.client.host if request.client else "unknown"
            now = time.time()
            window = 60.0

            # Clean old entries
            _rate_limit_store[client_ip] = [
                t for t in _rate_limit_store[client_ip] if now - t < window
            ]

            if len(_rate_limit_store[client_ip]) >= settings.RATE_LIMIT:
                return Response(
                    content="Rate limit exceeded",
                    status_code=429,
                    media_type="text/plain",
                )
            _rate_limit_store[client_ip].append(now)

        # API Key validation (skip for exempt paths)
        if path not in EXEMPT_PATHS:
            api_key = request.headers.get("X-API-Key")
            if api_key != settings.API_KEY:
                return Response(
                    content="Invalid or missing API key",
                    status_code=401,
                    media_type="text/plain",
                )

        return await call_next(request)
