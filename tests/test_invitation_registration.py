"""正式邀请码注册：权限、事务、额度和并发。"""
from concurrent.futures import ThreadPoolExecutor
import re
import threading

import pytest
from fastapi.testclient import TestClient
from sqlalchemy import func, select

from web_app.db import InvitationCode, InvitationRegistration, User, session_scope
from web_app.invitations import InvitationError
from web_app.limiter import register_limiter, reset_limiters

PASSWORD = "teacher-invite-test-pass"
ADMIN_PASSWORD = "admin-invite-test-pass"


@pytest.fixture
def app(tmp_path, monkeypatch):
    monkeypatch.setenv("SECRET_KEY", "invite-test-secret")
    monkeypatch.setenv("ALLOW_SQLITE", "1")
    monkeypatch.setenv("DATABASE_URL", "sqlite:///" + (tmp_path / "invites.db").as_posix())
    monkeypatch.setenv("UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setenv("COOKIE_SECURE", "false")
    monkeypatch.setenv("ADMIN_PASSWORD", ADMIN_PASSWORD)
    monkeypatch.setenv("INVITE_USER_DEFAULT_USES", "1")
    for name in ("ADMIN_PASSWORD_HASH", "ADMIN_USERNAME", "ADMIN_HOST"):
        monkeypatch.delenv(name, raising=False)
    reset_limiters()
    monkeypatch.setattr(register_limiter, "max_calls", 100)
    from web_app.app import create_app
    return create_app()


@pytest.fixture
def admin(app):
    with TestClient(app) as client:
        assert client.post("/admin/api/login", json={"username": "admin", "password": ADMIN_PASSWORD}).status_code == 200
        csrf = client.get("/admin/api/invites").json()["csrf_token"]
        client.headers["X-Admin-CSRF"] = csrf
        yield client


def issue(admin, limit=1):
    response = admin.post("/admin/api/invites", json={"max_uses": limit})
    assert response.status_code == 201, response.text
    return response.json()["invitation"]


def register(client, name, code):
    return client.post("/api/register", json={"username": name, "password": PASSWORD, "invite_code": code})


def counts():
    with session_scope() as db:
        return (db.scalar(select(func.count(User.id))), db.scalar(select(func.count(InvitationRegistration.user_id))))


def test_empty_database_has_no_public_demo_or_bypass(app, admin):
    assert admin.get("/admin/api/invites").json()["invites"] == []
    with TestClient(app) as client:
        response = client.post("/api/register", json={"username": "missing", "password": PASSWORD})
        assert response.status_code == 400
        assert "邀请码" in response.json()["detail"]
        assert register(client, "demo-code", "DEMO26").status_code == 400
        assert counts() == (0, 0)


@pytest.mark.parametrize("value", ["", None, 123456, True, [], {}, "ABC", "CP-DEMO-2026", "ABC$12", "ＡＢＣ１２３"])
def test_registration_requires_well_formed_code(app, value):
    with TestClient(app) as client:
        assert register(client, "bad-invite", value).status_code == 400
        assert counts() == (0, 0)


def test_register_and_personal_invitation_chain(app, admin):
    invitation = issue(admin, 2)
    assert re.fullmatch(r"[A-Z0-9]{6}", invitation["code"])
    assert re.search("[A-Z]", invitation["code"]) and re.search("[0-9]", invitation["code"])
    with TestClient(app) as client:
        created = register(client, "Teacher.One", " " + invitation["code"].lower() + " ")
        assert created.status_code == 200, created.text
        own = created.json()["invitation"]
        assert own["max_uses"] == 1 and own["remaining"] == 1
        assert re.fullmatch(r"[A-Z0-9]{6}", own["code"])
        assert client.post("/api/login", json={"username": "teacher.one", "password": PASSWORD}).status_code == 200
        me = client.get("/api/me").json()
        assert me["invitation"]["code"] == own["code"]
        assert own["owner_user_id"] == me["id"]
        second = register(client, "teacher.two", own["code"].lower())
        assert second.status_code == 200, second.text
        assert second.json()["invitation"]["code"] != own["code"]
        assert register(client, "teacher.three", own["code"]).status_code == 400
    users = admin.get("/admin/api/users").json()["users"]
    assert next(row for row in users if row["username"] == "teacher.two")["invited_by"] == "teacher.one"
    first = next(row for row in users if row["username"] == "teacher.one")
    assert first["invited_by"] == "管理员"
    assert first["invitation"]["used_count"] == 1
    assert first["invitation"]["status"] == "exhausted"


def test_success_only_consumes_quota_and_admin_can_extend_it(app, admin):
    invitation = issue(admin, 2)
    with TestClient(app) as client:
        assert register(client, "duplicate.user", invitation["code"]).status_code == 200
        assert register(client, "Duplicate.User", invitation["code"]).status_code == 409
        weak = client.post("/api/register", json={"username": "weak-user", "password": "short", "invite_code": invitation["code"]})
        assert weak.status_code == 400
        assert register(client, "second.user", invitation["code"]).status_code == 200
        assert register(client, "third.user", invitation["code"]).status_code == 400
        below = admin.patch(f'/admin/api/invites/{invitation["id"]}', json={"max_uses": 1})
        assert below.status_code == 400
        assert "已使用" in below.json()["detail"]
        extended = admin.patch(f'/admin/api/invites/{invitation["id"]}', json={"max_uses": 3})
        assert extended.status_code == 200
        assert extended.json()["invitation"]["remaining"] == 1
        assert register(client, "third.user", invitation["code"]).status_code == 200
    assert counts() == (3, 3)


def test_disabled_and_deleted_codes_keep_accounts_and_history(app, admin):
    invitation = issue(admin, 3)
    with TestClient(app) as client:
        first = register(client, "history.user", invitation["code"])
        assert first.status_code == 200
        own = first.json()["invitation"]
        assert admin.patch(f'/admin/api/invites/{invitation["id"]}', json={"enabled": False}).status_code == 200
        assert register(client, "disabled.user", invitation["code"]).status_code == 400
        assert admin.patch(f'/admin/api/invites/{invitation["id"]}', json={"enabled": True}).status_code == 200
        assert register(client, "history.child", own["code"]).status_code == 200
        assert admin.delete(f'/admin/api/invites/{invitation["id"]}').status_code == 200
        assert admin.delete(f'/admin/api/invites/{own["id"]}').status_code == 200
        assert register(client, "deleted.user", invitation["code"]).status_code == 400
        assert admin.patch(f'/admin/api/invites/{own["id"]}', json={"enabled": True}).status_code == 404
        assert client.post("/api/login", json={"username": "history.user", "password": PASSWORD}).status_code == 200
        assert client.get("/api/me").json()["invitation"]["status"] == "deleted"
    listed = admin.get("/admin/api/invites").json()["invites"]
    assert invitation["id"] not in {row["id"] for row in listed}
    users = admin.get("/admin/api/users").json()["users"]
    assert next(row for row in users if row["username"] == "history.child")["invited_by"] == "history.user"
    assert counts() == (2, 2)
    with session_scope() as db:
        assert db.get(InvitationCode, invitation["id"]).used_count == 1
        assert db.get(InvitationCode, own["id"]).used_count == 1


@pytest.mark.parametrize("value", [-1, 1.5, True, "5", 1_000_001])
def test_admin_limit_validation(admin, value):
    assert admin.post("/admin/api/invites", json={"max_uses": value}).status_code == 400


def test_zero_and_unlimited_quota(admin, app):
    zero, unlimited = issue(admin, 0), issue(admin, None)
    assert zero["status"] == "exhausted"
    assert unlimited["max_uses"] is None and unlimited["remaining"] is None
    with TestClient(app) as client:
        assert register(client, "zero-user", zero["code"]).status_code == 400
        for i in range(4):
            assert register(client, f"unlimited-{i}", unlimited["code"]).status_code == 200
    refreshed = next(row for row in admin.get("/admin/api/invites").json()["invites"] if row["id"] == unlimited["id"])
    assert refreshed["used_count"] == 4 and refreshed["status"] == "active"


def test_admin_permissions_and_csrf(app, admin, monkeypatch):
    invitation = issue(admin, 5)
    path = f'/admin/api/invites/{invitation["id"]}'
    with TestClient(app) as anonymous:
        assert anonymous.get("/admin/api/invites").status_code == 401
        assert anonymous.post("/admin/api/invites", json={"max_uses": 5}).status_code == 401
        assert anonymous.patch(path, json={"max_uses": 5}).status_code == 401
        assert anonymous.delete(path).status_code == 401
        assert register(anonymous, "teacher-only", invitation["code"]).status_code == 200
        assert anonymous.post("/api/login", json={"username": "teacher-only", "password": PASSWORD}).status_code == 200
        assert anonymous.get("/admin/api/invites").status_code == 401
        assert anonymous.delete(path, headers={"X-Admin-CSRF": admin.headers["X-Admin-CSRF"]}).status_code == 401
    assert admin.post("/admin/api/invites", json={"max_uses": 1}, headers={"X-Admin-CSRF": ""}).status_code == 403
    assert admin.patch(path, json={"max_uses": 1}, headers={"Origin": "https://evil.example"}).status_code == 403
    assert admin.delete(path, headers={"Sec-Fetch-Site": "cross-site"}).status_code == 403
    assert admin.patch(path, json={"code": "HACK12", "used_count": 0}).status_code == 400
    assert admin.post("/admin/api/invites", content='{"max_uses":2}', headers={"Content-Type": "text/plain"}).status_code == 400
    assert admin.get("/admin/api/invites").headers["cache-control"] == "no-store"
    monkeypatch.delenv("ADMIN_PASSWORD")
    assert admin.get("/admin/api/invites").status_code == 503
    assert admin.delete(path).status_code == 503


def test_host_based_admin_paths_work(app, admin, monkeypatch):
    monkeypatch.setenv("ADMIN_HOST", "admin.calc.geekhuang.com")
    headers = {"Host": "admin.calc.geekhuang.com", "Origin": "http://admin.calc.geekhuang.com"}
    assert admin.get("/invites", headers=headers).status_code == 200
    result = admin.post("/api/invites", json={"max_uses": 2}, headers=headers)
    assert result.status_code == 201, result.text
    assert admin.get("/api/invites", headers=headers).status_code == 200


def test_registration_failure_rolls_back_account_quota_and_personal_code(app, admin, monkeypatch):
    invitation = issue(admin, 1)
    import web_app.auth as auth
    def fail(*args, **kwargs):
        raise InvitationError("模拟个人码生成失败", status=503)
    monkeypatch.setattr(auth, "new_invitation", fail)
    with TestClient(app) as client:
        assert register(client, "rollback.user", invitation["code"]).status_code == 503
    assert counts() == (0, 0)
    with session_scope() as db:
        assert db.get(InvitationCode, invitation["id"]).used_count == 0
        assert db.scalar(select(func.count(InvitationCode.id))) == 1


def test_random_collision_retries_and_deleted_values_remain_reserved(admin, monkeypatch):
    invitation = issue(admin, 1)
    assert admin.delete(f'/admin/api/invites/{invitation["id"]}').status_code == 200
    import web_app.invitations as invites
    codes = iter([invitation["code"], "FIX123"])
    monkeypatch.setattr(invites, "_random_code", lambda: next(codes))
    new = issue(admin, 2)
    assert new["code"] == "FIX123"


def test_concurrent_registrations_cannot_overdraw_last_slots(app, admin):
    invitation = issue(admin, 3)
    from web_app.auth import register_user
    barrier = threading.Barrier(12)
    def attempt(index):
        barrier.wait(timeout=15)
        try:
            return register_user(f"parallel-{index}", PASSWORD, invitation["code"])
        except InvitationError:
            return None
    with ThreadPoolExecutor(max_workers=12) as pool:
        results = list(pool.map(attempt, range(12)))
    assert sum(item is not None for item in results) == 3
    assert counts() == (3, 3)
    with session_scope() as db:
        assert db.get(InvitationCode, invitation["id"]).used_count == 3
        assert db.scalar(select(func.count(InvitationCode.id))) == 4


def test_existing_users_receive_one_stable_code_and_deleted_code_is_not_reissued(app, admin):
    from web_app.auth import hash_password
    from web_app.invitations import personal_invitation
    with session_scope() as db:
        user = User(username="existing.user", password_hash=hash_password(PASSWORD))
        db.add(user)
        db.flush()
        user_id = user.id
    with ThreadPoolExecutor(max_workers=4) as pool:
        rows = list(pool.map(lambda _: personal_invitation(user_id), range(4)))
    assert len({row["code"] for row in rows}) == 1
    code = rows[0]
    with TestClient(app) as client:
        assert client.post("/api/login", json={"username": "existing.user", "password": PASSWORD}).status_code == 200
        assert client.get("/api/me").json()["invitation"]["code"] == code["code"]
        assert admin.delete(f'/admin/api/invites/{code["id"]}').status_code == 200
        assert client.get("/api/me").json()["invitation"]["status"] == "deleted"
    with session_scope() as db:
        assert db.scalar(select(func.count(InvitationCode.id))) == 1


def test_personal_limit_configuration_and_schema_reinitialization(app, admin, monkeypatch):
    monkeypatch.setenv("INVITE_USER_DEFAULT_USES", "5")
    invitation = issue(admin, 2)
    with TestClient(app) as client:
        result = register(client, "quota-five", invitation["code"])
        assert result.json()["invitation"]["max_uses"] == 5
        from web_app.db import init_db
        init_db()
        assert counts() == (1, 1)
        assert register(client, "after-restart", invitation["code"]).status_code == 200


def test_default_admin_and_personal_codes_allow_five_registrations(app, admin, monkeypatch):
    monkeypatch.delenv("INVITE_USER_DEFAULT_USES", raising=False)
    created = admin.post("/admin/api/invites", json={})
    assert created.status_code == 201
    invitation = created.json()["invitation"]
    assert invitation["max_uses"] == 5 and invitation["remaining"] == 5
    with TestClient(app) as client:
        for index in range(5):
            response = register(client, f"default-five-{index}", invitation["code"])
            assert response.status_code == 200, response.text
            assert response.json()["invitation"]["max_uses"] == 5
        assert register(client, "default-sixth", invitation["code"]).status_code == 400
        adjusted = admin.patch(f'/admin/api/invites/{invitation["id"]}', json={"max_uses": 6})
        assert adjusted.status_code == 200
        assert register(client, "default-sixth", invitation["code"]).status_code == 200
