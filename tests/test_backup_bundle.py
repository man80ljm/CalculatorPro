import os
import sqlite3
import pytest
from cryptography.exceptions import InvalidTag
from scripts.backup_bundle import make_bundle, verify_bundle, upload_qiniu


@pytest.fixture
def source(tmp_path):
    database = tmp_path / "database.dump"
    with sqlite3.connect(database) as db:
        db.execute("CREATE TABLE example (name TEXT)")
        db.execute("INSERT INTO example VALUES (?)", ("合成课程",))
    uploads, code = tmp_path / "uploads", tmp_path / "code"
    uploads.mkdir(); code.mkdir()
    (uploads / "grade.xlsx").write_bytes(b"synthetic-file")
    (code / "app.py").write_text("print('synthetic')", encoding="utf-8")
    key = tmp_path / "key.bin"
    key.write_bytes(os.urandom(32)); os.chmod(key, 0o600)
    return database, uploads, code, key


def test_encrypted_backup_roundtrip_and_database_restore(source, tmp_path):
    database, uploads, code, key = source
    package = make_bundle(database, uploads, code, tmp_path / "backups", key, "sqlite")
    assert b"synthetic-file" not in package.read_bytes()
    target = tmp_path / "restored"
    manifest = verify_bundle(package, key, target)
    assert len(manifest["files"]) == 3
    assert (target / "uploads/grade.xlsx").read_bytes() == b"synthetic-file"
    with sqlite3.connect(target / "database.dump") as db:
        assert db.execute("SELECT name FROM example").fetchone()[0] == "合成课程"
    with pytest.raises(ValueError):
        verify_bundle(package, key, target)


def test_corrupt_backup_is_rejected_before_restore(source, tmp_path):
    database, uploads, code, key = source
    package = make_bundle(database, uploads, code, tmp_path / "backups", key, "sqlite")
    data = bytearray(package.read_bytes())
    data[30] ^= 1
    package.write_bytes(data)
    target = tmp_path / "must-not-exist"
    with pytest.raises(InvalidTag):
        verify_bundle(package, key, target)
    assert not target.exists()


def test_missing_dump_and_nested_destination_cannot_succeed(source, tmp_path):
    database, uploads, code, key = source
    with pytest.raises(ValueError):
        make_bundle(tmp_path / "missing", uploads, code, tmp_path / "backups", key)
    with pytest.raises(ValueError):
        make_bundle(database, uploads, code, uploads / "recursive", key, "sqlite")


def test_symlink_external_file_is_excluded(source, tmp_path, monkeypatch):
    database, uploads, code, key = source
    try:
        (uploads / "key-link").symlink_to(key)
    except OSError:
        # Windows 普通账号无符号链接权限时，模拟该入口而不增加跳过用例。
        (uploads / "key-link").write_bytes(key.read_bytes())
        from pathlib import Path
        original = Path.is_symlink
        monkeypatch.setattr(Path, "is_symlink", lambda path: path.name == "key-link" or original(path))
    package = make_bundle(database, uploads, code, tmp_path / "backups", key, "sqlite")
    assert "uploads/key-link" not in verify_bundle(package, key)["files"]


def test_qiniu_upload_verifies_remote_hash_and_preserves_local_on_failure(source, tmp_path, monkeypatch):
    import qiniu
    from types import SimpleNamespace
    database, uploads, code, key = source
    package = make_bundle(database, uploads, code, tmp_path / "backups", key, "sqlite")
    for name, value in {"QINIU_ACCESS_KEY": "test-access", "QINIU_SECRET_KEY": "test-secret", "QINIU_BACKUP_BUCKET": "test-private", "QINIU_BACKUP_PRIVATE": "1"}.items():
        monkeypatch.setenv(name, value)
    calls = []
    def put(token, object_key, file_path, **kwargs):
        assert kwargs["version"] == "v2"
        calls.append(object_key)
        return {"key": object_key}, SimpleNamespace(status_code=200)
    monkeypatch.setattr(qiniu, "put_file", put)
    monkeypatch.setattr(qiniu.BucketManager, "stat", lambda *args: ({"hash": qiniu.etag(str(package)), "fsize": package.stat().st_size}, SimpleNamespace(status_code=200)))
    objects = upload_qiniu(package)
    assert objects == calls
    assert all(name.startswith("calculatorpro-backups/v1/") for name in objects)
    monkeypatch.setattr(qiniu.BucketManager, "stat", lambda *args: ({"hash": "wrong", "fsize": 0}, SimpleNamespace(status_code=200)))
    with pytest.raises(RuntimeError, match="校验失败"):
        upload_qiniu(package)
    assert package.is_file()
