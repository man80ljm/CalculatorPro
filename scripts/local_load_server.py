"""仅由本地压测启动：独立库、合成 AI、禁止外部 HTTP。"""
from __future__ import annotations
import argparse
import json
import os
from pathlib import Path
import secrets
import sys
import threading
import time

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--directory", required=True)
    parser.add_argument("--port", type=int, required=True)
    parser.add_argument("--ai-delay", type=float, default=3)
    args = parser.parse_args()
    directory = Path(args.directory).resolve()
    local_root = (ROOT / ".local").resolve()
    if local_root not in directory.parents:
        parser.error("测试数据必须放在项目 .local 子目录")
    directory.mkdir(parents=True, exist_ok=True)
    for key in list(os.environ):
        if key.startswith(("DEEPSEEK", "CALC_DEEPSEEK")):
            os.environ.pop(key, None)
    os.environ.update({
        "SECRET_KEY": secrets.token_hex(48), "ALLOW_SQLITE": "1",
        "DATABASE_URL": "sqlite:///" + (directory / "app.db").as_posix(),
        "UPLOAD_DIR": str(directory / "uploads"), "COOKIE_SECURE": "false",
        "TRUSTED_PROXIES": "", "DEEPSEEK_API_KEY": secrets.token_urlsafe(32),
        "DEEPSEEK_BASE_URL": "http://127.0.0.1:1",
        "DEEPSEEK_ACCOUNT_CONCURRENCY": "5",
    })
    import requests
    import uvicorn
    from web_app.limiter import register_limiter, login_limiter
    # 只有批量准备账号需要放宽同一 IP 的注册/登录限流；报告限流和并发池保持原值。
    register_limiter.max_calls = login_limiter.max_calls = 1000
    lock = threading.Lock()
    stats = {"calls": 0, "active": 0, "peak_active": 0, "external_http_attempts": 0}

    class SimulatedResponse:
        status_code = 200
        def raise_for_status(self):
            pass
        def json(self):
            content = {"overall": "合成总体分析", "objectives": [
                {"analysis": "合成目标分析", "improvement": "合成改进措施"}
                for _ in range(2)
            ]}
            return {"choices": [{"message": {"content": json.dumps(content, ensure_ascii=False)}}]}

    def fake_post(*_args, **_kwargs):
        with lock:
            stats["calls"] += 1
            stats["active"] += 1
            stats["peak_active"] = max(stats["peak_active"], stats["active"])
        try:
            time.sleep(max(0, args.ai_delay))
            return SimulatedResponse()
        finally:
            with lock:
                stats["active"] -= 1

    def deny_http(*_args, **_kwargs):
        with lock:
            stats["external_http_attempts"] += 1
        raise RuntimeError("隔离压测禁止外部 HTTP")

    requests.sessions.Session.request = deny_http
    requests.post = fake_post
    from web_app.app import app
    @app.get("/__load_metrics")
    async def metrics():
        with lock:
            return dict(stats)
    uvicorn.run(app, host="127.0.0.1", port=args.port, log_level="warning")


if __name__ == "__main__":
    main()
