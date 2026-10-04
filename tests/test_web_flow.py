"""网页版：账号、课程隔离、关系表，以及模板到导出的全流程。AI 调用打桩。"""
import io
import sys
import zipfile
from pathlib import Path
from urllib.parse import unquote

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from web_app.limiter import login_limiter, login_user_limiter, register_limiter, reset_limiters  # noqa: E402

PASSWORD = "test-pass"

SAMPLE_RELATION = {
    "objectives_count": 2,
    "links": [
        {
            "name": "平时考核",
            "ratio": 0.3,
            "methods": [
                {
                    "name": "平时作业",
                    "supports": {"课程目标1": 0.5, "课程目标2": 0.5},
                    "subtotal": 1.0,
                }
            ],
        },
        {
            "name": "期末考核",
            "ratio": 0.7,
            "methods": [
                {
                    "name": "期末考试",
                    "supports": {"课程目标1": 0.4, "课程目标2": 0.6},
                    "subtotal": 1.0,
                }
            ],
        },
    ],
    "objectives_total_weights": {"课程目标1": 0.43, "课程目标2": 0.57},
    "total_sum": 1.0,
}


def _settings(mode="forward", **extra):
    data = {
        "mode": mode,
        "student_count": 3,
        "course_open_info": {
            "year_start": "2024",
            "year_end": "2025",
            "semester": "1",
            "term": "1",
            "course_name": "测试课程",
            "department": "计算机学院",
            "teacher": "王老师",
        },
        "course_basic_info": {
            "course_name": "测试课程",
            "credits": "3",
            "hours": "48",
            "course_type": "专业课",
            "course_code": "CS101",
            "school_year_term": "2024-2025-1",
            "college": "计算机学院",
            "teacher": "王老师",
            "major": "软件工程",
            "class_name": "软工2201",
            "student_count": "2",
            "exam_count": "2",
        },
        "ratios": {"usual": 0.3, "midterm": 0, "final": 0.7},
        "relation_payload": SAMPLE_RELATION,
        "grad_req_map": [
            {"objective": "课程目标1", "requirement": "毕业要求1", "indicator": "指标点1.1"},
            {"objective": "课程目标2", "requirement": "毕业要求2", "indicator": "指标点2.2"},
        ],
        "course_description": "这是一门用于测试的程序设计课程。",
        "objective_requirements": ["能完成基本程序设计", "能分析简单问题"],
        "spread_mode": "中跨度（7-13分）",
        "distribution": "标准正态",
        "noise_config": None,
        "report_style": "专业",
        "word_limit": 80,
    }
    data.update(extra)
    return data


@pytest.fixture
def app(tmp_path, monkeypatch):
    monkeypatch.setenv("SECRET_KEY", "test-secret-key")
    db_path = (tmp_path / "app.db").resolve()
    monkeypatch.setenv("DATABASE_URL", "sqlite:///" + db_path.as_posix())
    monkeypatch.setenv("UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setenv("COOKIE_SECURE", "false")
    monkeypatch.setenv("DEEPSEEK_API_KEY", "")
    reset_limiters()
    login_limiter.max_calls = 8
    login_user_limiter.max_calls = 8
    register_limiter.max_calls = 8
    from web_app.app import create_app

    return create_app()


@pytest.fixture
def client(app):
    with TestClient(app) as test_client:
        created = test_client.post("/api/register", json={"username": "teacher", "password": PASSWORD})
        assert created.status_code == 200, created.text
        logged = test_client.post("/api/login", json={"username": "teacher", "password": PASSWORD})
        assert logged.status_code == 200, logged.text
        assert "cp_session" in logged.cookies
        yield test_client


def _course(client, settings=None, name="测试课程"):
    created = client.post("/api/courses", json={"name": name})
    assert created.status_code == 200, created.text
    course_id = created.json()["id"]
    if settings is not None:
        saved = client.patch(f"/api/courses/{course_id}", json={"settings": settings})
        assert saved.status_code == 200, saved.text
    return course_id


def test_healthz_is_public_and_pages_require_login(app):
    with TestClient(app) as anonymous:
        assert anonymous.get("/healthz").json() == {"status": "ok"}
        page = anonymous.get("/", follow_redirects=False)
        assert page.status_code == 303
        assert page.headers["location"] == "/login"
        assert anonymous.get("/api/courses").status_code == 401
        assert anonymous.post("/api/password", json={"current_password": "x", "new_password": "yyyyyyyy"}).status_code == 401
        assert anonymous.get("/static/app.js", follow_redirects=False).status_code == 303
        login = anonymous.get("/login")
        assert login.status_code == 200
        assert "登录" in login.text
        assert "注册" in login.text
        register = anonymous.get("/register")
        assert register.status_code == 200
        assert "api_key" not in login.text.lower()
        assert "deepseek" not in login.text.lower()


def test_secret_key_required(monkeypatch, tmp_path):
    monkeypatch.setenv("SECRET_KEY", "")
    monkeypatch.setenv("DATABASE_URL", "sqlite:///" + (tmp_path / "missing-key.db").resolve().as_posix())
    monkeypatch.setenv("UPLOAD_DIR", str(tmp_path / "uploads"))
    from web_app.app import create_app

    with pytest.raises(RuntimeError, match="SECRET_KEY"):
        create_app()


def test_register_duplicate_and_weak_password(app):
    with TestClient(app) as anonymous:
        weak = anonymous.post("/api/register", json={"username": "new-teacher", "password": "short"})
        assert weak.status_code == 400
        assert "8" in weak.json()["detail"]
        created = anonymous.post("/api/register", json={"username": "New.Teacher@School.edu", "password": PASSWORD})
        assert created.status_code == 200, created.text
        duplicate = anonymous.post("/api/register", json={"username": "new.teacher@school.edu", "password": PASSWORD})
        assert duplicate.status_code == 409


def test_login_wrong_password_logout_and_cookie_flags(app):
    with TestClient(app) as anonymous:
        anonymous.post("/api/register", json={"username": "teacher", "password": PASSWORD})
        bad = anonymous.post("/api/login", json={"username": "teacher", "password": "wrong-password"})
        assert bad.status_code == 401
        missing = anonymous.post("/api/login", json={"username": "ghost-user", "password": PASSWORD})
        assert missing.status_code == 401
        logged = anonymous.post("/api/login", json={"username": "teacher", "password": PASSWORD})
        assert logged.status_code == 200
        cookie = logged.headers["set-cookie"].lower()
        assert "httponly" in cookie
        assert "samesite=lax" in cookie
        assert "secure" not in cookie
        token = logged.cookies.get("cp_session")
        assert anonymous.get("/api/courses").status_code == 200
        assert anonymous.post("/api/logout").status_code == 200
        assert anonymous.get("/api/courses").status_code == 401
    with TestClient(app) as replay:
        replay.cookies.set("cp_session", token)
        assert replay.get("/api/courses").status_code == 401


def test_cookie_secure_flag(app, monkeypatch):
    monkeypatch.setenv("COOKIE_SECURE", "true")
    with TestClient(app) as anonymous:
        anonymous.post("/api/register", json={"username": "secure-user", "password": PASSWORD})
        logged = anonymous.post("/api/login", json={"username": "secure-user", "password": PASSWORD})
        assert "secure" in logged.headers["set-cookie"].lower()


def test_password_change(client):
    wrong = client.post("/api/password", json={"current_password": "not-the-pass", "new_password": "brand-new-pass"})
    assert wrong.status_code == 400
    short = client.post("/api/password", json={"current_password": PASSWORD, "new_password": "short"})
    assert short.status_code == 400
    changed = client.post("/api/password", json={"current_password": PASSWORD, "new_password": "brand-new-pass"})
    assert changed.status_code == 200, changed.text
    assert client.post("/api/logout").status_code == 200
    client.cookies.clear()
    old = client.post("/api/login", json={"username": "teacher", "password": PASSWORD})
    assert old.status_code == 401
    fresh = client.post("/api/login", json={"username": "teacher", "password": "brand-new-pass"})
    assert fresh.status_code == 200
    assert client.get("/api/me").json()["username"] == "teacher"


def test_login_rate_limit(app):
    with TestClient(app) as anonymous:
        statuses = [
            anonymous.post("/api/login", json={"username": "teacher", "password": "bad-password"}).status_code
            for _ in range(9)
        ]
    assert statuses[:8] == [401] * 8
    assert statuses[8] == 429


def test_login_rate_limit_per_username(app):
    login_limiter.max_calls = 100
    login_user_limiter.max_calls = 3
    try:
        with TestClient(app) as anonymous:
            one = [
                anonymous.post("/api/login", json={"username": "same-person", "password": "bad-password"}).status_code
                for _ in range(4)
            ]
            other = anonymous.post("/api/login", json={"username": "other-person", "password": "bad-password"}).status_code
        assert one == [401, 401, 401, 429]
        assert other == 401
    finally:
        login_limiter.max_calls = 8
        login_user_limiter.max_calls = 8


def test_register_rate_limit(app):
    register_limiter.max_calls = 2
    try:
        with TestClient(app) as anonymous:
            statuses = [
                anonymous.post("/api/register", json={"username": f"user{index}", "password": PASSWORD}).status_code
                for index in range(3)
            ]
        assert statuses == [200, 200, 429]
    finally:
        register_limiter.max_calls = 8


def test_home_has_relation_grid(client):
    page = client.get("/")
    assert page.status_code == 200
    assert "添加行" in page.text
    assert "删除选中行" in page.text
    assert "relation_grid.js" in page.text
    script = client.get("/static/relation_grid.js")
    assert script.status_code == 200
    assert "function parsePastedTable" in script.text


def test_course_crud_and_pasted_relation_grid(client):
    course_id = _course(client, name="高等数学")
    listed = client.get("/api/courses").json()["courses"]
    assert any(item["id"] == course_id and item["name"] == "高等数学" for item in listed)
    renamed = client.patch(f"/api/courses/{course_id}", json={"name": "线性代数"})
    assert renamed.status_code == 200
    assert renamed.json()["name"] == "线性代数"
    pasted = (
        "考核环节\t占比\t考核方式\t课程目标1\t课程目标2\t小计\r\n"
        "平时考核\t0.3\t平时作业\t50%\t50%\t100%\r\n"
        "期末考核\t0.7\t期末考试\t40%\t60%\t100%\r\n"
    )
    settings = _settings()
    settings.pop("relation_payload")
    settings["relation_grid"] = pasted
    saved = client.patch(f"/api/courses/{course_id}", json={"settings": settings})
    assert saved.status_code == 200, saved.text
    stored = saved.json()["settings"]
    assert stored["relation_grid"][1][0] == "平时考核"
    assert stored["relation_grid"][2][3] == "40%"
    assert stored["relation_payload"]["links"][0]["ratio"] == pytest.approx(0.3)
    template = client.post(f"/api/courses/{course_id}/template")
    assert template.status_code == 200, template.text
    assert "正向成绩模板" in unquote(template.headers["content-disposition"])
    assert client.delete(f"/api/courses/{course_id}").status_code == 200
    assert client.get(f"/api/courses/{course_id}").status_code == 404


def test_user_isolation(app, tmp_path):
    with TestClient(app) as alice, TestClient(app) as bob:
        assert alice.post("/api/register", json={"username": "alice", "password": "alice-pass"}).status_code == 200
        assert bob.post("/api/register", json={"username": "bob", "password": "bob-password"}).status_code == 200
        assert alice.post("/api/login", json={"username": "alice", "password": "alice-pass"}).status_code == 200
        assert bob.post("/api/login", json={"username": "bob", "password": "bob-password"}).status_code == 200
        created = alice.post("/api/courses", json={"name": "仅甲可见", "settings": _settings()})
        assert created.status_code == 200, created.text
        course_id = created.json()["id"]
        uploaded = alice.post(
            f"/api/courses/{course_id}/files",
            data={"kind": "grade"},
            files={"file": ("../../etc/passwd.xlsx", b"PK\x03\x04secret", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        )
        assert uploaded.status_code == 200, uploaded.text
        file_id = uploaded.json()["id"]
        assert uploaded.json()["original_name"].endswith("passwd.xlsx")
        assert bob.get("/api/courses").json()["courses"] == []
        assert bob.get(f"/api/courses/{course_id}").status_code == 404
        assert bob.patch(f"/api/courses/{course_id}", json={"name": "偷看"}).status_code == 404
        assert bob.delete(f"/api/courses/{course_id}").status_code == 404
        assert bob.get(f"/api/courses/{course_id}/files").status_code == 404
        assert bob.get(f"/api/courses/{course_id}/files/{file_id}").status_code == 404
        assert bob.post(f"/api/courses/{course_id}/calculate").status_code == 404
        assert bob.post(f"/api/courses/{course_id}/template").status_code == 404
        own = alice.get(f"/api/courses/{course_id}/files/{file_id}")
        assert own.status_code == 200
        assert own.content.startswith(b"PK")
        root = tmp_path / "uploads"
        stored = [path for path in root.rglob("*") if path.is_file() and path.name != ".write-probe"]
        assert stored
        assert all(path.parent.parent.parent == root for path in stored)
        assert all("passwd" not in path.name for path in stored)
        assert alice.get(f"/api/courses/{course_id}").status_code == 200


def test_web_stack_does_not_import_pyqt():
    assert "PyQt6" not in sys.modules
    import web_app.service  # noqa: F401

    assert "PyQt6" not in sys.modules
    assert "PyQt6.QtWidgets" not in sys.modules


def _cell_value(sheet, row, col):
    value = sheet.cell(row, col).value
    if value is not None:
        return value
    for merged in sheet.merged_cells.ranges:
        if merged.min_row <= row <= merged.max_row and merged.min_col <= col <= merged.max_col:
            return sheet.cell(merged.min_row, merged.min_col).value
    return None


def _fill_forward_template(template_bytes: bytes) -> bytes:
    workbook = load_workbook(io.BytesIO(template_bytes))
    sheet = workbook.active
    headers = [_cell_value(sheet, 2, col) for col in range(1, sheet.max_column + 1)]
    assert headers[0] == "姓名"
    assert "平时作业" in headers
    assert "期末考试" in headers
    usual_col = headers.index("平时作业") + 1
    final_col = headers.index("期末考试") + 1
    rows = [("张三", 86, 92), ("李四", 74, 63)]
    for offset, (name, usual, final) in enumerate(rows):
        row = 3 + offset
        sheet.cell(row, 1, name)
        sheet.cell(row, usual_col, usual)
        sheet.cell(row, final_col, final)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def test_forward_template_calculate_export_and_mocked_ai(client, monkeypatch):
    outputs = ROOT / "outputs"
    before = {path.name for path in outputs.iterdir()} if outputs.exists() else set()
    course_id = _course(client, _settings())
    template = client.post(f"/api/courses/{course_id}/template")
    assert template.status_code == 200, template.text
    assert "正向成绩模板" in unquote(template.headers["content-disposition"])
    excel_bytes = _fill_forward_template(template.content)
    uploaded = client.post(
        f"/api/courses/{course_id}/files",
        data={"kind": "grade"},
        files={"file": ("grades.xlsx", excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert uploaded.status_code == 200, uploaded.text

    calculated = client.post(f"/api/courses/{course_id}/calculate")
    assert calculated.status_code == 200, calculated.text
    body = calculated.json()
    assert body["course_name"] == "测试课程"
    assert body["student_count"] == 2
    assert body["average_score"] > 0
    assert "总达成度" in body["achievement"]
    assert any(name.endswith(".xlsx") for name in body["files"])
    assert any("课程成绩统计表" in name for name in body["files"])
    assert any("评价结果" in name for name in body["files"])

    exported = client.post(f"/api/courses/{course_id}/export")
    assert exported.status_code == 200, exported.text
    assert exported.headers["content-type"].startswith("application/zip")
    archive = zipfile.ZipFile(io.BytesIO(exported.content))
    names = archive.namelist()
    assert any(name.endswith(".xlsx") for name in names)
    assert any(name.endswith("2.课程成绩统计表.docx") for name in names)
    assert any(name.endswith("5.基于考核结果的课程目标达成情况评价结果表.docx") for name in names)
    assert any(name.endswith("1.课程基本信息表.docx") for name in names)
    assert any(name.endswith("4.课程考核与课程目标对应关系表.docx") for name in names)
    detail_name = next(name for name in names if name.endswith("成绩明细.xlsx"))
    detail = load_workbook(io.BytesIO(archive.read(detail_name)))
    assert detail.active["A2"].value == "张三"
    after = {path.name for path in outputs.iterdir()} if outputs.exists() else set()
    assert after == before

    listing = client.get(f"/api/courses/{course_id}/files").json()["files"]
    assert any(item["kind"] == "grade" for item in listing)
    assert any(item["kind"] == "output" and item["original_name"].endswith("成绩明细.xlsx") for item in listing)

    monkeypatch.setenv("DEEPSEEK_API_KEY", "")
    missing = client.post(f"/api/courses/{course_id}/ai-report")
    assert missing.status_code == 400
    assert "DEEPSEEK_API_KEY" in missing.json()["detail"]
    status = client.get("/api/ai-status").json()
    assert status["enabled"] is False

    calls = []

    class _Response:
        text = ""

        def raise_for_status(self):
            return None

        def json(self):
            return {
                "choices": [
                    {
                        "message": {
                            "content": "本轮课程目标达成情况总体平稳，考核覆盖了主要能力点，后续将补充针对性练习。"
                        }
                    }
                ]
            }

    def fake_post(url, headers=None, json=None, timeout=None):
        calls.append({"url": url, "headers": headers, "json": json, "timeout": timeout})
        return _Response()

    monkeypatch.setenv("DEEPSEEK_API_KEY", "sk-test-secret")
    monkeypatch.setattr("core_app.ai_report.requests.post", fake_post)
    report = client.post(f"/api/courses/{course_id}/ai-report")
    assert report.status_code == 200, report.text
    assert calls, "AI 接口没有被调用"
    assert "api.deepseek.com" in calls[0]["url"]
    assert calls[0]["json"]["model"] == "deepseek-chat"
    assert calls[0]["headers"]["Authorization"] == "Bearer sk-test-secret"
    report_zip = zipfile.ZipFile(io.BytesIO(report.content))
    report_names = report_zip.namelist()
    assert any(name.startswith("6.课程目标达成情况分析") for name in report_names)
    assert any("达成度报告" in name for name in report_names)
    assert b"sk-test-secret" not in report.content
    enabled = client.get("/api/ai-status").json()
    assert enabled["enabled"] is True
    assert "sk-test-secret" not in enabled["message"]

    saved_report = client.get(f"/api/courses/{course_id}/files").json()["files"]
    assert any(item["kind"] == "report" for item in saved_report)
    detail_file = next(item for item in saved_report if item["original_name"].endswith("成绩明细.xlsx"))
    assert client.post("/api/logout").status_code == 200
    client.cookies.clear()
    again = client.post("/api/login", json={"username": "teacher", "password": PASSWORD})
    assert again.status_code == 200, again.text
    downloaded = client.get(f"/api/courses/{course_id}/files/{detail_file['id']}")
    assert downloaded.status_code == 200, downloaded.text
    reopened = load_workbook(io.BytesIO(downloaded.content))
    assert reopened.active["A2"].value == "张三"


def test_reverse_export(client):
    course_id = _course(client, _settings(mode="reverse"), name="逆向课程")
    template = client.post(f"/api/courses/{course_id}/template")
    assert template.status_code == 200, template.text
    workbook = load_workbook(io.BytesIO(template.content))
    sheet = workbook.active
    headers = [sheet.cell(1, col).value for col in range(1, sheet.max_column + 1)]
    assert headers[0] == "姓名"
    sheet.cell(2, 1, "赵六")
    sheet.cell(2, headers.index("平时考核") + 1, 80)
    sheet.cell(2, headers.index("期末考核") + 1, 88)
    sheet.cell(3, 1, "钱七")
    sheet.cell(3, headers.index("平时考核") + 1, 70)
    sheet.cell(3, headers.index("期末考核") + 1, 76)
    buffer = io.BytesIO()
    workbook.save(buffer)
    uploaded = client.post(
        f"/api/courses/{course_id}/files",
        data={"kind": "grade"},
        files={"file": ("reverse.xlsx", buffer.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert uploaded.status_code == 200, uploaded.text
    exported = client.post(f"/api/courses/{course_id}/export")
    assert exported.status_code == 200, exported.text
    names = zipfile.ZipFile(io.BytesIO(exported.content)).namelist()
    assert any("逆向" in name and name.endswith(".xlsx") for name in names)
    assert any(name.endswith(".docx") for name in names)
    listed = client.get(f"/api/courses/{course_id}/files").json()["files"]
    assert any(item["kind"] == "output" for item in listed)


def test_outputs_override_is_isolated():
    import tempfile
    import threading
    import time

    from utils import get_outputs_dir, override_outputs_dir

    results = {}

    def worker(index):
        with tempfile.TemporaryDirectory() as directory:
            with override_outputs_dir(directory):
                results[index] = get_outputs_dir()
                time.sleep(0.05)
                assert get_outputs_dir() == results[index]

    threads = [threading.Thread(target=worker, args=(index,)) for index in range(8)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()
    assert len(set(results.values())) == 8


def test_upload_limit(client):
    course_id = _course(client, name="大文件")
    huge = b"PK" + b"x" * (10 * 1024 * 1024)
    response = client.post(
        f"/api/courses/{course_id}/files",
        data={"kind": "grade"},
        files={"file": ("grades.xlsx", huge, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 400
    assert "10MB" in response.json()["detail"]
    rejected = client.post(
        f"/api/courses/{course_id}/files",
        data={"kind": "output"},
        files={"file": ("grades.xlsx", b"PK\x03\x04ok", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert rejected.status_code == 400


def test_password_change_revokes_other_sessions(app):
    with TestClient(app) as first, TestClient(app) as second:
        first.post("/api/register", json={"username": "multi-device", "password": PASSWORD})
        assert first.post("/api/login", json={"username": "multi-device", "password": PASSWORD}).status_code == 200
        assert second.post("/api/login", json={"username": "multi-device", "password": PASSWORD}).status_code == 200
        assert second.get("/api/me").status_code == 200
        changed = first.post("/api/password", json={"current_password": PASSWORD, "new_password": "another-pass-1"})
        assert changed.status_code == 200, changed.text
        assert first.get("/api/me").status_code == 200
        assert second.get("/api/me").status_code == 401


def test_client_key_only_trusts_forwarded_from_proxies(monkeypatch):
    from starlette.requests import Request

    from web_app.app import _client_key

    def make(host, xff=None):
        headers = [(b"x-forwarded-for", xff.encode())] if xff else []
        return Request({"type": "http", "client": (host, 1234), "headers": headers})

    monkeypatch.delenv("TRUSTED_PROXIES", raising=False)
    # 来自 Docker 网关（受信）时，取代理追加的最右一项，忽略客户端伪造的前缀
    assert _client_key(make("172.18.0.1", "6.6.6.6, 203.0.113.9")) == "203.0.113.9"
    # 直连的外部地址不能伪造 X-Forwarded-For
    assert _client_key(make("198.51.100.7", "1.2.3.4")) == "198.51.100.7"
    monkeypatch.setenv("TRUSTED_PROXIES", "10.0.0.0/8")
    assert _client_key(make("172.18.0.1", "203.0.113.9")) == "172.18.0.1"


def test_database_url_required_unless_sqlite_opt_in(monkeypatch, tmp_path):
    from web_app.app import create_app
    from web_app.db import database_url

    monkeypatch.setenv("UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.delenv("DATABASE_URL", raising=False)
    monkeypatch.delenv("ALLOW_SQLITE", raising=False)
    with pytest.raises(RuntimeError, match="DATABASE_URL is required"):
        create_app()
    monkeypatch.setenv("DATABASE_URL", "sqlite:///" + (tmp_path / "x.db").as_posix())
    with pytest.raises(RuntimeError, match="ALLOW_SQLITE"):
        create_app()
    monkeypatch.setenv("ALLOW_SQLITE", "1")
    assert database_url().startswith("sqlite:///")
    monkeypatch.delenv("DATABASE_URL")
    assert database_url() == "sqlite:///./calculatorpro.db"
    monkeypatch.setenv("ALLOW_SQLITE", "0")
    monkeypatch.setenv("DATABASE_URL", "postgresql+psycopg://u:p@db:5432/x")
    assert database_url().startswith("postgresql+psycopg://")
