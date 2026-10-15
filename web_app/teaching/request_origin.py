"""Cookie writes accept the site's origin, including a trusted TLS proxy."""
import ipaddress
from urllib.parse import urlsplit


def same_origin(request):
    if request.headers.get("sec-fetch-site") == "cross-site":
        return False
    origin = request.headers.get("origin")
    if not origin:
        return True
    try:
        parsed = urlsplit(origin)
        if parsed.scheme not in {"http", "https"} or parsed.username is not None or parsed.password is not None:
            return False
        if parsed.path not in {"", "/"} or parsed.query or parsed.fragment:
            return False
        parsed.port  # Reject malformed ports.
    except ValueError:
        return False
    scheme = request.url.scheme
    forwarded = request.headers.get("x-forwarded-proto", "").strip().lower()
    if forwarded in {"http", "https"} and request.client:
        try:
            peer = ipaddress.ip_address(request.client.host)
        except ValueError:
            peer = None
        if peer is not None:
            from web_app.app import _trusted_proxies
            if any(peer in network for network in _trusted_proxies()):
                scheme = forwarded
    return parsed.scheme == scheme and parsed.netloc.lower() == request.headers.get("host", "").lower()
