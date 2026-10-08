"""只读运维后台：独立登录、统计，以及没有写接口。"""
import pytest
from fastapi.testclient import TestClient

from web_app.limiter import login_limiter, login_user_limiter, register_limiter, reset_limiters
from web_app.report_jobs import reset_report_jobs

ADMIN_PASSWORD = "test-admin-pass-9"
TEACHER_PASSWORD = "teacher-pass-1"
_WRITE_METHODS = {"DELETE", "PUT", "PATCH"}
_BANNED_PATH = ("impersonate", "password", "delete", "settings")


@pytest.fixture
def app(tmp_path, monkeypatch):
    monkeypatch.setenv("SECRET_KEY", "test-secret-key")
    db_path = (tmp_path / "admin.db").resolve()
    monkeypatch.setenv("DATABASE_URL", "sqlite:///" + db_path.as_posix())
    monkeypatch.setenv("UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setenv("COOKIE_SECURE", "false")
    monkeypatch.setenv("DEEPSEEK_API_KEY", "")
    monkeypatch.setenv("ADMIN_PASSWORD", ADMIN_PASSWORD)
    monkeypatch.delenv("ADMIN_PASSWORD_HASH", raising=False)
    monkeypatch.delenv("ADMIN_USERNAME", raising=False)
    monkeypatch.delenv("ADMIN_HOST", raising=False)
    reset_limiters()
    reset_report_jobs()
    login_limiter.max_calls = 20
    login_user_limiter.max_calls = 20
    register_limiter.max_calls = 20
    from web_app.app import create_app

    application = create_app()
    yield application
    reset_report_jobs()


def _admin_routes(application):
    found = []
    for route in application.routes:
        path = getattr(route, "path", "") or ""
        if path == "/admin" or path.startswith("/admin/"):
            methods = set(getattr(route, "methods", None) or ())
            found.append((path, methods))
    return found


def test_unconfigured_admin_is_closed(app, monkeypatch):
    monkeypatch.delenv("ADMIN_PASSWORD", raising=False)
    monkeypatch.delenv("ADMIN_PASSWORD_HASH", raising=False)
    with TestClient(app) as client:
        login = client.post("/admin/api/login", json={"username": "admin", "password": ADMIN_PASSWORD})
        assert login.status_code == 503
        assert login.json()["detail"] == "未配置管理员密码"
        assert "cp_admin_session" not in login.cookies
        for path in ("/admin/api/overview", "/admin/api/users", "/admin/api/jobs"):
            denied = client.get(path)
            assert denied.status_code == 503
            assert denied.json()["detail"] == "未配置管理员密码"
        page = client.get("/admin/", follow_redirects=False)
        assert page.status_code == 200
        assert "未配置管理员密码" in page.text
        assert "ADMIN_PASSWORD" in page.text
        assert "ADMIN_PASSWORD_HASH" in page.text


def test_wrong_password_and_teacher_password_rejected(app):
    with TestClient(app) as client:
        created = client.post("/api/register", json={"username": "admin", "password": TEACHER_PASSWORD})
        assert created.status_code == 200, created.text
        wrong = client.post("/admin/api/login", json={"username": "admin", "password": "not-the-admin-pass"})
        assert wrong.status_code == 401
        assert "cp_admin_session" not in wrong.cookies
        teacher_pass = client.post("/admin/api/login", json={"username": "admin", "password": TEACHER_PASSWORD})
        assert teacher_pass.status_code == 401
        teacher = client.post("/api/login", json={"username": "admin", "password": ADMIN_PASSWORD})
        assert teacher.status_code == 401
        teacher_ok = client.post("/api/login", json={"username": "admin", "password": TEACHER_PASSWORD})
        assert teacher_ok.status_code == 200
        assert client.get("/api/me").status_code == 200
        for path in ("/admin/api/overview", "/admin/api/users", "/admin/api/jobs"):
            denied = client.get(path)
            assert denied.status_code == 401


def test_admin_login_cookie_and_reads(app, monkeypatch):
    with TestClient(app) as client:
        logged = client.post("/admin/api/login", json={"username": "admin", "password": ADMIN_PASSWORD})
        assert logged.status_code == 200, logged.text
        header = logged.headers["set-cookie"].lower()
        assert "cp_admin_session=" in header
        assert "httponly" in header
        assert "samesite=lax" in header
        assert "path=/" in header
        assert "secure" not in header
        assert client.get("/api/courses").status_code == 401
        assert client.get("/api/me").status_code == 401
        token = logged.cookies.get("cp_admin_session")
        client.cookies.set("cp_session", token)
        assert client.get("/api/me").status_code == 401
        assert client.get("/admin/api/overview").status_code == 200
        assert client.get("/admin/api/users").status_code == 200
        assert client.get("/admin/api/jobs").status_code == 200
        logged_out = client.post("/admin/api/logout")
        assert logged_out.status_code == 200
        assert client.get("/admin/api/overview").status_code == 401

    monkeypatch.setenv("COOKIE_SECURE", "true")
    with TestClient(app) as client:
        logged = client.post("/admin/api/login", json={"username": "Admin", "password": ADMIN_PASSWORD})
        assert logged.status_code == 200, logged.text
        assert "secure" in logged.headers["set-cookie"].lower()


def test_password_hash_overrides_plaintext(app, monkeypatch):
    from web_app.auth import hash_password

    monkeypatch.setenv("ADMIN_PASSWORD", "plain-not-used-99")
    monkeypatch.setenv("ADMIN_PASSWORD_HASH", hash_password(ADMIN_PASSWORD))
    with TestClient(app) as client:
        plain = client.post("/admin/api/login", json={"username": "admin", "password": "plain-not-used-99"})
        assert plain.status_code == 401
        hashed = client.post("/admin/api/login", json={"username": "admin", "password": ADMIN_PASSWORD})
        assert hashed.status_code == 200, hashed.text


def test_non_ascii_admin_password_is_not_a_server_error(app, monkeypatch):
    monkeypatch.setenv("ADMIN_USERNAME", "运维")
    monkeypatch.setenv("ADMIN_PASSWORD", "运维密码-9")
    with TestClient(app) as client:
        wrong = client.post("/admin/api/login", json={"username": "运维", "password": "不对的口令"})
        assert wrong.status_code == 401
        logged = client.post("/admin/api/login", json={"username": "运维", "password": "运维密码-9"})
        assert logged.status_code == 200, logged.text
        assert client.get("/admin/api/overview").status_code == 200


def test_custom_username_and_password_change_invalidates_cookie(app, monkeypatch):
    monkeypatch.setenv("ADMIN_USERNAME", "ops")
    with TestClient(app) as client:
        assert client.post("/admin/api/login", json={"username": "admin", "password": ADMIN_PASSWORD}).status_code == 401
        assert client.post("/admin/api/login", json={"username": "Ops", "password": ADMIN_PASSWORD}).status_code == 200
        assert client.get("/admin/api/users").status_code == 200
        monkeypatch.setenv("ADMIN_PASSWORD", "another-admin-pass")
        assert client.get("/admin/api/users").status_code == 401
        again = client.post("/admin/api/login", json={"username": "ops", "password": "another-admin-pass"})
        assert again.status_code == 200


def test_overview_users_and_jobs_match_seeded_data(app):
    from web_app.db import CourseFile, session_scope
    from web_app.report_jobs import ReportJob, _JOBS, _LOCK, _succeed

    with TestClient(app) as client:
        created_a = client.post("/api/register", json={"username": "teacher.a@school.edu", "password": TEACHER_PASSWORD})
        assert created_a.status_code == 200, created_a.text
        assert client.post("/api/login", json={"username": "teacher.a@school.edu", "password": TEACHER_PASSWORD}).status_code == 200
        teacher_a = client.get("/api/me").json()
        course_response = client.post("/api/courses", json={"name": "高等数学"})
        assert course_response.status_code == 200, course_response.text
        course = course_response.json()
        with session_scope() as db:
            db.add(
                CourseFile(
                    user_id=teacher_a["id"],
                    course_id=course["id"],
                    term_id=course["current_term_id"],
                    kind="report",
                    stored_name="report.zip",
                    original_name="报告.zip",
                    size=12,
                )
            )

        created_b = client.post("/api/register", json={"username": "teacher.b@school.edu", "password": TEACHER_PASSWORD})
        assert created_b.status_code == 200, created_b.text
        assert client.post("/api/login", json={"username": "teacher.b@school.edu", "password": TEACHER_PASSWORD}).status_code == 200
        teacher_b = client.get("/api/me").json()

        done = ReportJob(teacher_a["id"], course["id"], course["current_term_id"])
        done.content = b"zip-bytes-should-not-leak"
        with _LOCK:
            _JOBS[done.id] = done
        _succeed(done, {"hidden": True}, "报告.zip", done.content)
        queued = ReportJob(teacher_a["id"], course["id"], course["current_term_id"])
        queued.stage = "queued"
        queued.content = b"still-hidden"
        with _LOCK:
            _JOBS[queued.id] = queued

        client.cookies.clear()
        logged = client.post("/admin/api/login", json={"username": "admin", "password": ADMIN_PASSWORD})
        assert logged.status_code == 200, logged.text
        assert client.get("/api/courses").status_code == 401

        overview = client.get("/admin/api/overview")
        assert overview.status_code == 200, overview.text
        body = overview.json()
        assert body["teacher_count"] == 2
        assert body["course_count"] == 1
        assert body["active_teachers_week"] == 2
        assert body["reports_today"] == {"success": 1, "fail": 0}
        assert body["reports_week"] == {"success": 1, "fail": 0}
        assert body["report_queued"] == 1
        assert body["deepseek_queue"] == 0
        assert body["queue_depth"] == body["deepseek_queue"] + body["report_queued"]

        users = client.get("/admin/api/users")
        assert users.status_code == 200, users.text
        assert "password" not in users.text.lower()
        listed = users.json()["users"]
        assert {row["username"] for row in listed} == {"teacher.a@school.edu", "teacher.b@school.edu"}
        row_a = next(row for row in listed if row["username"] == "teacher.a@school.edu")
        row_b = next(row for row in listed if row["username"] == "teacher.b@school.edu")
        assert row_a["course_count"] == 1
        assert row_a["term_count"] == 1
        assert row_a["created_at"]
        assert row_a["last_login_at"]
        assert row_a["last_report_at"]
        assert row_b["course_count"] == 0
        assert row_b["term_count"] == 0
        assert row_b["last_login_at"]
        assert row_b["last_report_at"] is None
        assert teacher_b["id"] != teacher_a["id"]

        jobs = client.get("/admin/api/jobs")
        assert jobs.status_code == 200, jobs.text
        rows = jobs.json()["jobs"]
        assert "zip-bytes-should-not-leak" not in jobs.text
        assert "still-hidden" not in jobs.text
        for row in rows:
            assert "content" not in row
            assert "password" not in row
            assert "summary" not in row
        memory_queued = [row for row in rows if row["source"] == "memory" and row["status"] == "queued"]
        assert len(memory_queued) == 1
        assert memory_queued[0]["username"] == "teacher.a@school.edu"
        assert memory_queued[0]["course_name"] == "高等数学"
        events = [row for row in rows if row["source"] == "event" and row["status"] == "success"]
        assert len(events) == 1
        assert events[0]["username"] == "teacher.a@school.edu"
        assert events[0]["course_name"] == "高等数学"
        assert events[0]["error"] == ""
        assert events[0]["duration_ms"] >= 0

        reset_report_jobs()
        again = client.get("/admin/api/overview").json()
        assert again["reports_today"]["success"] == 1
        assert again["reports_week"]["fail"] == 0
        assert again["report_queued"] == 0
        remaining = client.get("/admin/api/jobs").json()["jobs"]
        assert any(row["source"] == "event" and row["status"] == "success" for row in remaining)
        assert not any(row["source"] == "memory" for row in remaining)


def test_failed_job_records_short_error(app):
    from web_app.report_jobs import ReportJob, _fail

    job = ReportJob(1, 2, 3)
    _fail(job, "ai", RuntimeError("坏了" * 200))
    with TestClient(app) as client:
        assert client.post("/admin/api/login", json={"username": "admin", "password": ADMIN_PASSWORD}).status_code == 200
        overview = client.get("/admin/api/overview").json()
        assert overview["reports_today"]["fail"] == 1
        assert overview["reports_today"]["success"] == 0
        rows = client.get("/admin/api/jobs").json()["jobs"]
        event = next(row for row in rows if row["source"] == "event")
        assert event["status"] == "fail"
        assert event["error"].startswith("坏了")
        assert len(event["error"]) <= 180
        assert "坏了" * 200 not in event["error"]


def test_no_admin_write_routes(app):
    paths = _admin_routes(app)
    assert paths
    post_paths = set()
    for path, methods in paths:
        lowered = path.lower()
        for banned in _BANNED_PATH:
            assert banned not in lowered
        assert not (_WRITE_METHODS & methods)
        if "POST" in methods:
            post_paths.add(path)
    assert post_paths == {"/admin/api/login", "/admin/api/logout"}

    with TestClient(app) as client:
        assert client.post("/admin/api/login", json={"username": "admin", "password": ADMIN_PASSWORD}).status_code == 200
        for path in (
            "/admin/api/users",
            "/admin/api/jobs",
            "/admin/api/overview",
            "/admin/api/password",
            "/admin/api/impersonate",
            "/admin/api/settings",
            "/admin/users",
            "/admin/jobs",
        ):
            for method in ("DELETE", "PUT", "PATCH"):
                response = client.request(method, path)
                assert response.status_code in {404, 405}, (method, path, response.status_code)
        for path in ("/admin/api/overview", "/admin/api/users", "/admin/api/jobs", "/admin/api/password", "/admin/api/impersonate"):
            response = client.post(path, json={})
            assert response.status_code in {404, 405}, (path, response.status_code)


def test_admin_pages_static_and_host_gate(app, monkeypatch):
    with TestClient(app) as client:
        slash = client.get("/admin", follow_redirects=False)
        assert slash.status_code == 307
        assert slash.headers["location"] == "/admin/"
        home = client.get("/admin/", follow_redirects=False)
        assert home.status_code == 303
        assert home.headers["location"] == "login"
        login = client.get("/admin/login")
        assert login.status_code == 200
        assert "运维" in login.text
        assert "static/admin.css" in login.text
        assert "只读" in login.text
        css = client.get("/admin/static/admin.css")
        assert css.status_code == 200
        assert ".shell" in css.text
        script = client.get("/admin/static/admin.js")
        assert script.status_code == 200
        assert client.get("/admin/static/db.py").status_code == 404
        assert client.post("/admin/api/login", json={"username": "admin", "password": ADMIN_PASSWORD}).status_code == 200
        page = client.get("/admin/")
        assert page.status_code == 200
        assert "概览" in page.text
        assert "本周活跃" in page.text
        assert client.get("/admin/users").status_code == 200
        assert "最近登录" in client.get("/admin/users").text
        assert "不能改密码" in client.get("/admin/users").text
        jobs = client.get("/admin/jobs")
        assert jobs.status_code == 200
        assert "不能重新触发" in jobs.text

    monkeypatch.setenv("ADMIN_HOST", "admin.calc.geekhuang.com")
    with TestClient(app) as client:
        teacher = client.get("/", follow_redirects=False)
        assert teacher.status_code == 303
        assert teacher.headers["location"] == "/login"
        admin_root = client.get("/", headers={"host": "admin.calc.geekhuang.com"}, follow_redirects=False)
        assert admin_root.status_code == 303
        assert admin_root.headers["location"] == "login"
        login = client.get("/login", headers={"host": "admin.calc.geekhuang.com"})
        assert login.status_code == 200
        assert "运维" in login.text
        health = client.get("/healthz", headers={"host": "admin.calc.geekhuang.com"})
        assert health.status_code == 200
        assert health.json()["status"] == "ok"


def test_admin_narrow_screen_reads_users_and_jobs(app):
    with TestClient(app) as client:
        css = client.get("/admin/static/admin.css")
        assert css.status_code == 200
        assert "data-cards" in css.text
        assert "@media" in css.text
        assert "max-width: 720px" in css.text
        script = client.get("/admin/static/admin.js")
        assert script.status_code == 200
        assert "data-cards" in script.text
        assert client.post("/admin/api/login", json={"username": "admin", "password": ADMIN_PASSWORD}).status_code == 200
        for path in ("/admin/users", "/admin/jobs"):
            page = client.get(path)
            assert page.status_code == 200
            assert 'name="viewport"' in page.text
            assert "width=device-width" in page.text


def test_queue_depth_counts_waiting_tasks():
    from web_app.deepseek_pool import _Waiter, get_pool, queue_depth, reset_pool

    reset_pool()
    try:
        assert queue_depth() == 0
        pool = get_pool()
        pool.waiters.append(_Waiter("report", None))
        assert queue_depth() == 1
    finally:
        reset_pool()
