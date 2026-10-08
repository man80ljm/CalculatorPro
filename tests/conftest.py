"""测试默认走本机 sqlite，避免误连部署用的 Postgres。"""
import os
from pathlib import Path

import pytest

_root = Path("/tmp/calculatorpro-pytest-bootstrap")
_root.mkdir(parents=True, exist_ok=True)
# 导入期会建一次库。清掉旧文件，避免停在 v2 之前的表结构上。
_bootstrap_db = _root / "bootstrap.db"
if _bootstrap_db.exists():
    _bootstrap_db.unlink()
os.environ["SECRET_KEY"] = "test-secret-key"
os.environ["ALLOW_SQLITE"] = "1"
os.environ["DATABASE_URL"] = "sqlite:///" + (_root / "bootstrap.db").as_posix()
os.environ["UPLOAD_DIR"] = str(_root / "uploads")
os.environ["COOKIE_SECURE"] = "false"
os.environ["DEEPSEEK_API_KEY"] = ""
for _name in (
    "DEEPSEEK_KEYS_A",
    "DEEPSEEK_KEYS_B",
    "DEEPSEEK_KEYS_A_FILE",
    "DEEPSEEK_KEYS_B_FILE",
    "DEEPSEEK_ACCOUNT_CONCURRENCY",
    "DEEPSEEK_QUEUE_REPORT_SECONDS",
    "DEEPSEEK_QUEUE_SYLLABUS_SECONDS",
):
    os.environ.pop(_name, None)

import web_app.app  # noqa: E402,F401


@pytest.fixture(autouse=True)
def _reset_deepseek_pool_between_tests():
    from web_app.deepseek_pool import reset_pool

    reset_pool()
    yield
    reset_pool()
