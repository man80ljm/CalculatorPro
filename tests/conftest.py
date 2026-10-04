"""测试默认走本机 sqlite，避免误连部署用的 Postgres。"""
import os
from pathlib import Path

_root = Path("/tmp/calculatorpro-pytest-bootstrap")
_root.mkdir(parents=True, exist_ok=True)
os.environ["SECRET_KEY"] = "test-secret-key"
os.environ["ALLOW_SQLITE"] = "1"
os.environ["DATABASE_URL"] = "sqlite:///" + (_root / "bootstrap.db").as_posix()
os.environ["UPLOAD_DIR"] = str(_root / "uploads")
os.environ["COOKIE_SECURE"] = "false"
os.environ["DEEPSEEK_API_KEY"] = ""

import web_app.app  # noqa: E402,F401
