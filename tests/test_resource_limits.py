"""资源预算拒绝异常文件；正常文件与计算规则仍走原流程。"""
import io
import zipfile
import pytest
from openpyxl import Workbook
from web_app.resource_limits import admit_request, release_request, validate_document
from web_app.service import ServiceError


def test_request_budget_per_teacher_and_total():
    acquired = []
    try:
        for user in range(5):
            for _ in range(4):
                assert admit_request(user)
                acquired.append(user)
            assert not admit_request(user)
        assert not admit_request(99)
        release_request(acquired.pop())
        assert admit_request(99)
        acquired.append(99)
    finally:
        for user in acquired:
            release_request(user)


def test_compressed_document_checks_expanded_size(monkeypatch):
    monkeypatch.setenv("DOCUMENT_EXPANDED_MB", "1")
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/media/oversized.png", b"0" * (2 * 1024 * 1024))
    assert len(buffer.getvalue()) < 10000
    with pytest.raises(ServiceError) as caught:
        validate_document(buffer.getvalue(), "syllabus.docx")
    assert caught.value.status == 413


def test_sparse_excel_does_not_expand_millions_of_empty_cells():
    workbook = Workbook()
    workbook.active["A1048576"] = "异常远处的单元格"
    buffer = io.BytesIO()
    workbook.save(buffer)
    with pytest.raises(ServiceError) as caught:
        validate_document(buffer.getvalue(), "grade.xlsx")
    assert caught.value.status == 413


def test_invalid_excel_rejected_and_normal_excel_accepted():
    with pytest.raises(ServiceError):
        validate_document(b"PK\x03\x04fake", "grade.xlsx")
    workbook = Workbook()
    workbook.active.append(["学号", "姓名", "成绩"])
    workbook.active.append(["T001", "合成学生", 88])
    buffer = io.BytesIO()
    workbook.save(buffer)
    validate_document(buffer.getvalue(), "grade.xlsx")


def test_limiter_bounds_distinct_keys_and_expires_old_records(monkeypatch):
    from web_app.limiter import RateLimiter
    limiter = RateLimiter(1, 60)
    monkeypatch.setattr("web_app.limiter.time.monotonic", lambda: 10)
    limiter._hits.update({str(i): [10] for i in range(10000)})
    assert not limiter.allow("new-key")
    assert len(limiter._hits) == 10000
    monkeypatch.setattr("web_app.limiter.time.monotonic", lambda: 71)
    assert limiter.allow("new-key")
    assert len(limiter._hits) == 1


def test_ai_total_limit_across_multiple_accounts(monkeypatch):
    from web_app.deepseek_pool import reserve, reset_pool, account_loads
    monkeypatch.setenv("DEEPSEEK_KEYS_A", "sk-test-account-a")
    monkeypatch.setenv("DEEPSEEK_KEYS_B", "sk-test-account-b")
    monkeypatch.setenv("DEEPSEEK_ACCOUNT_CONCURRENCY", "5")
    monkeypatch.setenv("DEEPSEEK_GLOBAL_CONCURRENCY", "5")
    reset_pool()
    held = [reserve("report").wait(1) for _ in range(5)]
    waiting = reserve("report")
    try:
        assert sum(account_loads().values()) == 5
        held.pop().release()
        lease = waiting.wait(1)
        assert sum(account_loads().values()) == 5
        lease.release()
    finally:
        waiting.cancel()
        for lease in held:
            lease.release()


def test_json_stream_limit_does_not_trust_content_length(tmp_path, monkeypatch):
    from fastapi.testclient import TestClient
    from web_app.app import create_app
    monkeypatch.setenv("DATABASE_URL", "sqlite:///" + (tmp_path / "limited.db").as_posix())
    monkeypatch.setenv("UPLOAD_DIR", str(tmp_path / "uploads"))
    with TestClient(create_app()) as client:
        chunks = iter([b'{"username":"', b"x" * (1024 * 1024), b'"}'])
        response = client.post("/api/register", content=chunks, headers={"Content-Type": "application/json"})
        assert response.status_code == 413


def test_abnormal_objective_count_rejected_before_template_allocation():
    from scripts.local_load_test import settings
    from web_app.service import build_template
    value = settings()
    value["student_count"] = 2
    value["relation_payload"]["objectives_count"] = 1000000
    with pytest.raises(ServiceError, match="最多支持 30"):
        build_template(value)
    value["relation_payload"]["objectives_count"] = 2
    assert build_template(value)[1].startswith(b"PK")
