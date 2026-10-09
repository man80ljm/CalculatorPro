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
        self._next_sweep = 0.0

    def allow(self, key: str) -> bool:
        now = time.monotonic()
        with self._lock:
            # 限流记录自身也要有上限，避免大量不同用户名/IP累积内存。
            if now >= self._next_sweep:
                stale = [name for name, hits in self._hits.items() if not hits or now - hits[-1] >= self.window_seconds]
                for name in stale:
                    self._hits.pop(name, None)
                self._next_sweep = now + 1
            if key not in self._hits and len(self._hits) >= 10000:
                return False
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
            self._next_sweep = 0.0


login_limiter = RateLimiter(max_calls=120, window_seconds=60)
login_user_limiter = RateLimiter(max_calls=8, window_seconds=60)
register_limiter = RateLimiter(max_calls=8, window_seconds=60)
# 防刷用，不是 DeepSeek 容量。真正同时打到模型的上限在 deepseek_pool
# （DEEPSEEK_ACCOUNT_CONCURRENCY，默认每个账号 5）。按大约 20–30 人同时使用留出余量。
ai_limiter = RateLimiter(max_calls=30, window_seconds=60)
# 建课读取按登录用户另计，和报告分开，避免误点刷额度。
syllabus_limiter = RateLimiter(max_calls=20, window_seconds=60)
operation_limiter = RateLimiter(max_calls=120, window_seconds=60)


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
    syllabus_limiter.reset()
    operation_limiter.reset()
