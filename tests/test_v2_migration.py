"""v2 学期表：旧库迁移、重复执行、退回，以及启动时不自动改表。"""
import json
from pathlib import Path

import pytest
from sqlalchemy import create_engine, inspect, text

from sqlalchemy.orm import Session

from web_app.db import Base, Course, Term, init_db
from web_app.migrate_v2 import (
    apply_migration,
    downgrade,
    downgrade_course_term_fields,
    main,
    plan,
    preview_term_fields,
    upgrade_course_term_fields,
)
from web_app.terms import NEW_TERM_LABEL


def _legacy(path: Path):
    engine = create_engine("sqlite:///" + path.as_posix())
    settings = {
        "mode": "reverse",
        "student_count": 12,
        "spread_mode": "中跨度（7-13分）",
        "course_open_info": {"year_start": "2024", "year_end": "2025", "semester": "1", "teacher": "王老师"},
        "course_basic_info": {
            "school_year_term": "2024-2025学年第1学期",
            "teacher": "王老师",
            "class_name": "软工2201",
            "exam_count": "11",
        },
    }
    with engine.begin() as conn:
        conn.execute(
            text(
                """
                CREATE TABLE courses (
                  id INTEGER PRIMARY KEY,
                  user_id INTEGER,
                  name VARCHAR(120),
                  settings_json TEXT,
                  created_at TIMESTAMP,
                  updated_at TIMESTAMP
                )
                """
            )
        )
        conn.execute(
            text(
                """
                CREATE TABLE course_files (
                  id INTEGER PRIMARY KEY,
                  user_id INTEGER,
                  course_id INTEGER,
                  kind VARCHAR(32),
                  stored_name VARCHAR(80),
                  original_name VARCHAR(255),
                  size INTEGER,
                  created_at TIMESTAMP
                )
                """
            )
        )
        conn.execute(
            text("INSERT INTO courses (id, user_id, name, settings_json) VALUES (1, 7, '旧课', :settings)"),
            {"settings": json.dumps(settings, ensure_ascii=False)},
        )
        conn.execute(
            text(
                "INSERT INTO course_files (id, user_id, course_id, kind, stored_name, original_name, size) "
                "VALUES (3, 7, 1, 'grade', 'a.xlsx', '成绩.xlsx', 10)"
            )
        )
    return engine


def test_dry_run_apply_twice_and_downgrade(tmp_path):
    db_path = tmp_path / "legacy.db"
    engine = _legacy(db_path)
    preview = plan(engine)
    assert preview["terms_to_create"] == 1
    assert preview["files_to_attach"] == 1
    assert preview["schema_missing"] is True
    assert "terms" not in inspect(engine).get_table_names()

    backup = tmp_path / "backups"
    url = "sqlite:///" + db_path.as_posix()
    assert main(["--dry-run", "--database-url", url]) == 0
    assert main(["--backup-dir", str(backup), "--database-url", url]) == 0
    saved = list(backup.glob("*.json"))
    assert saved and "旧课" in saved[0].read_text(encoding="utf-8")

    engine.dispose()
    engine = create_engine(url)
    with engine.connect() as conn:
        terms = conn.execute(text("SELECT * FROM terms")).mappings().all()
        files = conn.execute(text("SELECT term_id FROM course_files")).all()
    assert len(terms) == 1
    assert terms[0]["teacher"] == "王老师"
    assert terms[0]["class_name"] == "软工2201"
    assert terms[0]["student_count"] == 12
    assert terms[0]["is_current"] == 1
    assert "reverse" in terms[0]["settings_json"]
    assert files[0][0] == terms[0]["id"]

    again = apply_migration(engine)
    assert again["created"] == 0
    with engine.connect() as conn:
        assert conn.execute(text("SELECT COUNT(*) FROM terms")).scalar() == 1

    with engine.begin() as conn:
        conn.execute(text("UPDATE terms SET teacher = '李老师' WHERE course_id = 1"))
    assert downgrade(engine)["courses_restored"] == 1
    with engine.connect() as conn:
        names = {row[0] for row in conn.execute(text("SELECT name FROM sqlite_master WHERE type='table'"))}
        assert "terms" not in names
        cols = [row[1] for row in conn.execute(text("PRAGMA table_info(course_files)"))]
        assert "term_id" not in cols
        stored = json.loads(conn.execute(text("SELECT settings_json FROM courses WHERE id = 1")).scalar())
    assert stored["course_open_info"]["teacher"] == "李老师"
    assert stored["course_basic_info"]["class_name"] == "软工2201"
    assert stored["mode"] == "reverse"
    engine.dispose()


def test_startup_refuses_legacy_database(tmp_path, monkeypatch):
    db_path = tmp_path / "old.db"
    engine = _legacy(db_path)
    engine.dispose()
    monkeypatch.setenv("ALLOW_SQLITE", "1")
    monkeypatch.setenv("DATABASE_URL", "sqlite:///" + db_path.as_posix())
    monkeypatch.setenv("SECRET_KEY", "test-secret-key")
    with pytest.raises(RuntimeError, match="migrate_v2"):
        init_db()


def _nested_keys(value, found: set[str] | None = None) -> set[str]:
    found = set() if found is None else found
    if isinstance(value, dict):
        for key, item in value.items():
            found.add(key)
            _nested_keys(item, found)
    return found


def test_course_term_fields_move_onto_empty_terms(tmp_path, capsys):
    """空学期补上课程层身份，已填学期保持原值，然后课程 settings 去掉这些键。

    第二次 upgrade 不再改数据。downgrade_course_term_fields 只把当前学期写回课程层，
    不删除 terms。整库退回仍是 downgrade()，那个会删学期表；这里不调用它，也不读 /data。
    """
    db_path = tmp_path / "fields.db"
    engine = create_engine("sqlite:///" + db_path.as_posix())
    Base.metadata.create_all(engine)
    settings = {
        "mode": "reverse",
        "student_count": 12,
        "course_description": "稳定简介",
        "course_basic_info": {
            "course_name": "创意手作",
            "school_year_term": "2024-2025学年第1学期",
            "teacher": "王老师",
            "class_name": "软工2201",
            "major": "软件工程",
            "student_count": "12",
            "exam_count": "11",
        },
        "course_open_info": {
            "course_name": "创意手作",
            "department": "设计学院",
            "year_start": "2024",
            "year_end": "2025",
            "semester": "1",
            "teacher": "王老师",
        },
    }
    with Session(engine) as db:
        course = Course(user_id=1, name="创意手作", settings_json=json.dumps(settings, ensure_ascii=False))
        db.add(course)
        db.flush()
        db.add(
            Term(
                user_id=1,
                course_id=course.id,
                label=NEW_TERM_LABEL,
                year_start="",
                year_end="",
                semester="",
                school_year_term="",
                teacher="",
                class_name="",
                student_count=0,
                exam_count=0,
                settings_json="{}",
                is_current=0,
            )
        )
        db.add(
            Term(
                user_id=1,
                course_id=course.id,
                label="2023-2024学年第2学期",
                year_start="2023",
                year_end="2024",
                semester="2",
                school_year_term="2023-2024学年第2学期",
                teacher="李老师",
                class_name="数媒2301",
                student_count=8,
                exam_count=7,
                settings_json=json.dumps({"major": "数字媒体", "mode": "forward"}, ensure_ascii=False),
                is_current=1,
            )
        )
        db.commit()
        course_id = course.id

    url = "sqlite:///" + db_path.as_posix()
    assert main(["--dry-run", "--term-fields", "--database-url", url]) == 0
    preview = json.loads(capsys.readouterr().out)
    assert preview == {"courses_with_identity": 1, "terms_missing_identity": 1}
    assert preview_term_fields(engine) == preview

    first = upgrade_course_term_fields(engine)
    assert first["terms_filled"] == 1
    assert first["courses_stripped"] == 1
    with engine.connect() as conn:
        terms = {
            row["id"]: dict(row)
            for row in conn.execute(text("SELECT * FROM terms ORDER BY id")).mappings()
        }
        stored = json.loads(conn.execute(text("SELECT settings_json FROM courses WHERE id = :id"), {"id": course_id}).scalar())
    empty, filled = terms[1], terms[2]
    assert empty["teacher"] == "王老师"
    assert empty["class_name"] == "软工2201"
    assert empty["year_start"] == "2024"
    assert empty["year_end"] == "2025"
    assert empty["semester"] == "1"
    assert empty["school_year_term"] == "2024-2025学年第1学期"
    assert empty["label"] == "2024-2025学年第1学期"
    assert empty["student_count"] == 12
    assert empty["exam_count"] == 11
    assert json.loads(empty["settings_json"])["major"] == "软件工程"
    assert filled["teacher"] == "李老师"
    assert filled["class_name"] == "数媒2301"
    assert filled["year_start"] == "2023"
    assert filled["semester"] == "2"
    assert filled["student_count"] == 8
    assert filled["exam_count"] == 7
    assert json.loads(filled["settings_json"])["major"] == "数字媒体"
    forbidden = {
        "teacher",
        "major",
        "class_name",
        "year_start",
        "year_end",
        "semester",
        "term",
        "school_year_term",
        "student_count",
        "exam_count",
    }
    assert forbidden.isdisjoint(_nested_keys(stored))
    assert stored["course_basic_info"]["course_name"] == "创意手作"
    assert stored["course_open_info"]["department"] == "设计学院"
    assert stored["course_description"] == "稳定简介"
    assert stored["mode"] == "reverse"

    with engine.connect() as conn:
        before = [dict(row) for row in conn.execute(text("SELECT * FROM terms ORDER BY id")).mappings()]
        before_settings = conn.execute(text("SELECT settings_json FROM courses WHERE id = :id"), {"id": course_id}).scalar()
    again = upgrade_course_term_fields(engine)
    assert again["terms_filled"] == 0
    assert again["courses_stripped"] == 0
    with engine.connect() as conn:
        after = [dict(row) for row in conn.execute(text("SELECT * FROM terms ORDER BY id")).mappings()]
        after_settings = conn.execute(text("SELECT settings_json FROM courses WHERE id = :id"), {"id": course_id}).scalar()
    assert after == before
    assert after_settings == before_settings

    assert downgrade_course_term_fields(engine)["courses_restored"] == 1
    with engine.connect() as conn:
        names = {row[0] for row in conn.execute(text("SELECT name FROM sqlite_master WHERE type='table'"))}
        restored = json.loads(conn.execute(text("SELECT settings_json FROM courses WHERE id = :id"), {"id": course_id}).scalar())
        still = {row["id"]: dict(row) for row in conn.execute(text("SELECT * FROM terms")).mappings()}
    assert "terms" in names
    assert still[1]["teacher"] == "王老师"
    assert still[1]["class_name"] == "软工2201"
    assert still[2]["teacher"] == "李老师"
    assert restored["course_open_info"]["teacher"] == "李老师"
    assert restored["course_open_info"]["year_start"] == "2023"
    assert restored["course_open_info"]["year_end"] == "2024"
    assert restored["course_open_info"]["semester"] == "2"
    assert restored["course_basic_info"]["class_name"] == "数媒2301"
    assert restored["course_basic_info"]["major"] == "数字媒体"
    assert restored["student_count"] == 8
    assert restored["course_open_info"]["department"] == "设计学院"
    engine.dispose()
