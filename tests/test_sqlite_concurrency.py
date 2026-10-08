"""文件SQLite的事务隔离：一个老师的提交/回滚不能影响另一个老师。"""
from concurrent.futures import ThreadPoolExecutor
import threading

import pytest
from sqlalchemy import text
from web_app.db import get_engine


@pytest.fixture
def engine(tmp_path, monkeypatch):
    monkeypatch.setenv("ALLOW_SQLITE", "1")
    monkeypatch.setenv("DATABASE_URL", "sqlite:///" + (tmp_path / "concurrency.db").as_posix())
    engine = get_engine()
    with engine.begin() as connection:
        connection.execute(text("CREATE TABLE probes (id INTEGER PRIMARY KEY, owner INTEGER)"))
    yield engine
    engine.dispose()


def test_other_transaction_cannot_read_or_commit_uncommitted_work(engine):
    with engine.connect() as left:
        transaction = left.begin()
        left.execute(text("INSERT INTO probes VALUES (1, 101)"))
        with engine.begin() as right:
            assert right.execute(text("SELECT COUNT(*) FROM probes")).scalar_one() == 0
        transaction.rollback()
    with engine.connect() as check:
        assert check.execute(text("SELECT COUNT(*) FROM probes")).scalar_one() == 0


def test_twenty_concurrent_writers_keep_independent_commits_and_rollbacks(engine):
    barrier = threading.Barrier(20)
    def write(owner):
        barrier.wait(timeout=10)
        try:
            with engine.begin() as connection:
                for offset in range(3):
                    connection.execute(text("INSERT INTO probes VALUES (:id, :owner)"),
                                       {"id": owner * 3 + offset, "owner": owner})
                if owner % 4 == 0:
                    raise ValueError("模拟本人的事务失败")
        except ValueError:
            pass
    with ThreadPoolExecutor(max_workers=20) as executor:
        list(executor.map(write, range(20)))
    with engine.connect() as check:
        rows = check.execute(text("SELECT id, owner FROM probes")).all()
    assert set(rows) == {(owner * 3 + offset, owner) for owner in range(20) if owner % 4 for offset in range(3)}


@pytest.mark.parametrize("url", ["sqlite://", "sqlite:///:memory:"])
def test_memory_database_remains_available_across_sequential_sessions(monkeypatch, url):
    monkeypatch.setenv("ALLOW_SQLITE", "1")
    monkeypatch.setenv("DATABASE_URL", url)
    engine = get_engine()
    try:
        with engine.begin() as one:
            one.execute(text("CREATE TABLE memory_probe (id INTEGER)"))
            one.execute(text("INSERT INTO memory_probe VALUES (7)"))
        with engine.connect() as two:
            assert two.execute(text("SELECT id FROM memory_probe")).scalar_one() == 7
    finally:
        engine.dispose()
