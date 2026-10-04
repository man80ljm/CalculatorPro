"""很小的内存限流：登录（按 IP 和用户名）、注册（按 IP）和 AI 接口。"""
import threading
import time
from collections import defaultdict


class RateLimiter:
    def __init__(self, max_calls: int, window_seconds: float):
        self.max_calls = max_calls
        self.window_seconds = window_seconds
        self._hits: dict[str, list[float]] = defaultdict(list)
        self._lock = threading.Lock()

    def allow(self, key: str) -> bool:
        now = time.monotonic()
        with self._lock:
            recent = [stamp for stamp in self._hits[key] if now - stamp < self.window_seconds]
            if len(recent) >= self.max_calls:
                self._hits[key] = recent
                return False
            recent.append(now)
            self._hits[key] = recent
            return True

    def reset(self) -> None:
        with self._lock:
            self._hits.clear()


login_limiter = RateLimiter(max_calls=8, window_seconds=60)
login_user_limiter = RateLimiter(max_calls=8, window_seconds=60)
register_limiter = RateLimiter(max_calls=8, window_seconds=60)
ai_limiter = RateLimiter(max_calls=6, window_seconds=60)


def allow_login(ip: str, username: str) -> bool:
    """同一次尝试同时计入 IP 和用户名，任一超限都拒绝。"""
    user_key = (username or "").strip().casefold() or "-"
    ip_ok = login_limiter.allow(f"ip:{ip or 'unknown'}")
    user_ok = login_user_limiter.allow(f"user:{user_key}")
    return ip_ok and user_ok


def reset_limiters() -> None:
    login_limiter.reset()
    login_user_limiter.reset()
    register_limiter.reset()
    ai_limiter.reset()
