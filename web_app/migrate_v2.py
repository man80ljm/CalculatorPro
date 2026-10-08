"""v2 学期表迁移。可重复执行，可预览，可退回。

启动服务时如果发现旧库缺表，会拒绝启动并提示运行本命令，不会在打开页面时改数据。

用法：
  python -m web_app.migrate_v2 --dry-run
  python -m web_app.migrate_v2 --backup-dir ./backups
  python -m web_app.migrate_v2 --downgrade --backup-dir ./backups
  python -m web_app.migrate_v2 --term-fields --backup-dir ./backups
  python -m web_app.migrate_v2 --term-fields --downgrade --backup-dir ./backups

正式库备份还可以用：
  docker compose exec -T postgres pg_dump -U "$POSTGRES_USER" "$POSTGRES_DB" -t courses -t course_files -t terms | gzip > backups/v2-courses.sql.gz
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

from sqlalchemy import create_engine, inspect, text
from sqlalchemy.engine import Engine

from web_app.db import Term, database_url, sqlite_allowed
from web_app.terms import parse_settings, settings_from_term_row, term_label, without_term_identity


def _engine(url: str) -> Engine:
    if url.startswith("sqlite"):
        return create_engine(url, connect_args={"check_same_thread": False})
    return create_engine(url, pool_pre_ping=True)


def _tables(engine: Engine) -> set[str]:
    return set(inspect(engine).get_table_names())


def _columns(engine: Engine, table: str) -> set[str]:
    if table not in _tables(engine):
        return set()
    return {column["name"] for column in inspect(engine).get_columns(table)}


def _rows(engine: Engine, sql: str) -> list[dict]:
    with engine.connect() as conn:
        result = conn.execute(text(sql))
        return [dict(row._mapping) for row in result]


def _jsonable(value):
    if isinstance(value, datetime):
        return value.isoformat(timespec="seconds")
    return value


def backup_tables(engine: Engine, directory: Path) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    payload = {}
    for table in ("courses", "course_files", "terms"):
        if table not in _tables(engine):
            payload[table] = []
            continue
        rows = _rows(engine, f"SELECT * FROM {table}")
        payload[table] = [{key: _jsonable(val) for key, val in row.items()} for row in rows]
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    path = directory / f"v2-migrate-{stamp}.json"
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _ensure_schema(engine: Engine) -> None:
    Term.__table__.create(engine, checkfirst=True)
    if "course_files" in _tables(engine) and "term_id" not in _columns(engine, "course_files"):
        with engine.begin() as conn:
            conn.execute(text("ALTER TABLE course_files ADD COLUMN term_id INTEGER"))


def _course_rows(engine: Engine) -> list[dict]:
    if "courses" not in _tables(engine):
        return []
    return _rows(engine, "SELECT id, user_id, settings_json FROM courses ORDER BY id")


def plan(engine: Engine) -> dict:
    courses = _course_rows(engine)
    has_terms = "terms" in _tables(engine)
    existing = set()
    if has_terms:
        existing = {row["course_id"] for row in _rows(engine, "SELECT course_id FROM terms")}
    pending = [row["id"] for row in courses if row["id"] not in existing]
    files = 0
    if "course_files" in _tables(engine) and "term_id" in _columns(engine, "course_files"):
        files = len(_rows(engine, "SELECT id FROM course_files WHERE term_id IS NULL"))
    elif "course_files" in _tables(engine):
        files = len(_rows(engine, "SELECT id FROM course_files"))
    return {
        "courses": len(courses),
        "terms_to_create": len(pending),
        "files_to_attach": files,
        "schema_missing": "terms" not in _tables(engine) or "term_id" not in _columns(engine, "course_files"),
    }


def apply_migration(engine: Engine) -> dict:
    summary = plan(engine)
    _ensure_schema(engine)
    now = datetime.now(timezone.utc).replace(tzinfo=None).isoformat(timespec="seconds")
    created = 0
    attached = 0
    with engine.begin() as conn:
        courses = conn.execute(text("SELECT id, user_id, settings_json FROM courses ORDER BY id")).mappings().all()
        have = {row["course_id"] for row in conn.execute(text("SELECT course_id FROM terms")).mappings()}
        for course in courses:
            if course["id"] in have:
                continue
            settings = parse_settings(course["settings_json"])
            basic = settings.get("course_basic_info") if isinstance(settings.get("course_basic_info"), dict) else {}
            open_info = settings.get("course_open_info") if isinstance(settings.get("course_open_info"), dict) else {}
            year_start = str(open_info.get("year_start") or "")[:16]
            year_end = str(open_info.get("year_end") or "")[:16]
            semester = str(open_info.get("semester") or open_info.get("term") or "")[:8]
            school_year_term = str(basic.get("school_year_term") or "")[:80]
            teacher = str(open_info.get("teacher") or basic.get("teacher") or "")[:80]
            class_name = str(basic.get("class_name") or "")[:120]
            try:
                student_count = int(float(settings.get("student_count") or basic.get("student_count") or 0))
            except (TypeError, ValueError):
                student_count = 0
            try:
                exam_count = int(float(basic.get("exam_count") or 0))
            except (TypeError, ValueError):
                exam_count = 0
            if school_year_term:
                label = school_year_term
            elif year_start and year_end and semester:
                label = f"{year_start}-{year_end}学年第{semester}学期"
            else:
                label = "新学期（请填写学年学期）"
            blob = {}
            for key in ("mode", "spread_mode", "distribution", "noise_config", "report_style", "word_limit", "student_count"):
                if key in settings:
                    blob[key] = settings.get(key)
            if "mode" not in blob:
                blob["mode"] = "forward"
            conn.execute(
                text(
                    """
                    INSERT INTO terms (
                      user_id, course_id, label, year_start, year_end, semester, school_year_term,
                      teacher, class_name, student_count, exam_count, settings_json, is_current,
                      created_at, updated_at
                    ) VALUES (
                      :user_id, :course_id, :label, :year_start, :year_end, :semester, :school_year_term,
                      :teacher, :class_name, :student_count, :exam_count, :settings_json, 1,
                      :created_at, :updated_at
                    )
                    """
                ),
                {
                    "user_id": course["user_id"],
                    "course_id": course["id"],
                    "label": label[:160],
                    "year_start": year_start,
                    "year_end": year_end,
                    "semester": semester,
                    "school_year_term": school_year_term,
                    "teacher": teacher,
                    "class_name": class_name,
                    "student_count": max(student_count, 0),
                    "exam_count": max(exam_count, 0),
                    "settings_json": json.dumps(blob, ensure_ascii=False),
                    "created_at": now,
                    "updated_at": now,
                },
            )
            created += 1
        term_of = {
            row["course_id"]: row["id"]
            for row in conn.execute(
                text("SELECT id, course_id FROM terms WHERE is_current = 1")
            ).mappings()
        }
        pending_files = conn.execute(text("SELECT id, course_id FROM course_files WHERE term_id IS NULL")).mappings().all()
        for item in pending_files:
            term_id = term_of.get(item["course_id"])
            if term_id is None:
                continue
            conn.execute(text("UPDATE course_files SET term_id = :term_id WHERE id = :id"), {"term_id": term_id, "id": item["id"]})
            attached += 1
    summary["created"] = created
    summary["attached"] = attached
    return summary


def _drop_term_column(engine: Engine) -> None:
    if "term_id" not in _columns(engine, "course_files"):
        return
    dialect = engine.dialect.name
    with engine.begin() as conn:
        if dialect == "postgresql":
            conn.execute(text("ALTER TABLE course_files DROP COLUMN IF EXISTS term_id"))
            return
        try:
            conn.execute(text("ALTER TABLE course_files DROP COLUMN term_id"))
        except Exception:
            conn.execute(text("ALTER TABLE course_files RENAME TO course_files_old"))
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
                text(
                    """
                    INSERT INTO course_files (id, user_id, course_id, kind, stored_name, original_name, size, created_at)
                    SELECT id, user_id, course_id, kind, stored_name, original_name, size, created_at FROM course_files_old
                    """
                )
            )
            conn.execute(text("DROP TABLE course_files_old"))


def _positive_count(value) -> int:
    try:
        number = int(float(str(value).strip()))
    except (TypeError, ValueError):
        return 0
    return number if number > 0 else 0


def _donor_identity(settings: dict) -> dict:
    """课程层里还能挪到学期上的身份。空值不参与填充。"""
    basic = settings.get("course_basic_info") if isinstance(settings.get("course_basic_info"), dict) else {}
    opened = settings.get("course_open_info") if isinstance(settings.get("course_open_info"), dict) else {}
    year_start = str(opened.get("year_start") or "").strip()[:16]
    year_end = str(opened.get("year_end") or "").strip()[:16]
    semester = str(opened.get("semester") or opened.get("term") or "").strip()[:8]
    school = str(basic.get("school_year_term") or "").strip()[:80]
    if not school and year_start and year_end and semester:
        school = f"{year_start}-{year_end}学年第{semester}学期"[:80]
    student = _positive_count(settings.get("student_count")) or _positive_count(basic.get("student_count"))
    return {
        "year_start": year_start,
        "year_end": year_end,
        "semester": semester,
        "school_year_term": school,
        "teacher": str(opened.get("teacher") or basic.get("teacher") or "").strip()[:80],
        "class_name": str(basic.get("class_name") or "").strip()[:120],
        "major": str(basic.get("major") or "").strip()[:120],
        "student_count": student,
        "exam_count": _positive_count(basic.get("exam_count")),
    }


def _identity_present(donor: dict) -> bool:
    return any(donor.values())


def upgrade_course_term_fields(engine: Engine) -> dict:
    """把课程层遗留的学期身份拷到缺这些字段的学期上，再从课程 settings 删掉。

    学期上已有的非空字段不覆盖。上课人数、考核人数为 0 视为缺失。
    上课专业写在学期 settings_json 的 major 里。没有学期的课先不动，留给建学期迁移读取。
    重复执行不再改数据。本函数不在启动时调用，也不读取生产库路径。
    """
    summary = {"courses_seen": 0, "terms_filled": 0, "courses_stripped": 0}
    if "courses" not in _tables(engine) or "terms" not in _tables(engine):
        return summary
    now = datetime.now(timezone.utc).replace(tzinfo=None).isoformat(timespec="seconds")
    text_keys = ("year_start", "year_end", "semester", "school_year_term", "teacher", "class_name")
    with engine.begin() as conn:
        courses = conn.execute(text("SELECT id, settings_json FROM courses ORDER BY id")).mappings().all()
        summary["courses_seen"] = len(courses)
        for course in courses:
            settings = parse_settings(course["settings_json"])
            donor = _donor_identity(settings)
            terms = conn.execute(
                text("SELECT * FROM terms WHERE course_id = :course_id ORDER BY id"),
                {"course_id": course["id"]},
            ).mappings().all()
            if not terms:
                continue
            for term in terms:
                current = {key: str(term[key] or "") for key in text_keys}
                blob = parse_settings(term["settings_json"])
                changed = False
                for key in text_keys:
                    if not current[key].strip() and donor[key]:
                        current[key] = donor[key]
                        changed = True
                student_count = int(term["student_count"] or 0)
                exam_count = int(term["exam_count"] or 0)
                if student_count <= 0 and donor["student_count"]:
                    student_count = donor["student_count"]
                    changed = True
                if exam_count <= 0 and donor["exam_count"]:
                    exam_count = donor["exam_count"]
                    changed = True
                if not str(blob.get("major") or "").strip() and donor["major"]:
                    blob["major"] = donor["major"]
                    changed = True
                if not current["school_year_term"] and current["year_start"] and current["year_end"] and current["semester"]:
                    current["school_year_term"] = f"{current['year_start']}-{current['year_end']}学年第{current['semester']}学期"[:80]
                    changed = True
                if not changed:
                    continue
                label = term_label(
                    current["year_start"],
                    current["year_end"],
                    current["semester"],
                    current["school_year_term"],
                )
                conn.execute(
                    text(
                        """
                        UPDATE terms SET
                          year_start = :year_start,
                          year_end = :year_end,
                          semester = :semester,
                          school_year_term = :school_year_term,
                          teacher = :teacher,
                          class_name = :class_name,
                          student_count = :student_count,
                          exam_count = :exam_count,
                          settings_json = :settings_json,
                          label = :label,
                          updated_at = :updated_at
                        WHERE id = :id
                        """
                    ),
                    {
                        **current,
                        "student_count": student_count,
                        "exam_count": exam_count,
                        "settings_json": json.dumps(blob, ensure_ascii=False),
                        "label": label[:160],
                        "updated_at": now,
                        "id": term["id"],
                    },
                )
                summary["terms_filled"] += 1
            cleaned = without_term_identity(settings)
            if cleaned != settings:
                conn.execute(
                    text("UPDATE courses SET settings_json = :settings WHERE id = :id"),
                    {"settings": json.dumps(cleaned, ensure_ascii=False), "id": course["id"]},
                )
                summary["courses_stripped"] += 1
    return summary


def downgrade_course_term_fields(engine: Engine) -> dict:
    """把当前学期的身份写回课程 settings。不删除 terms 表。

    与 downgrade() 不同：那个会把学期字段写回后删掉学期表。这里只恢复当前学期
    （is_current 优先，否则 id 最小）到课程层，其他学期保持原样。
    """
    if "courses" not in _tables(engine) or "terms" not in _tables(engine):
        return {"courses_restored": 0}
    restored = 0
    with engine.begin() as conn:
        courses = conn.execute(text("SELECT id, settings_json FROM courses ORDER BY id")).mappings().all()
        for course in courses:
            term = conn.execute(
                text(
                    """
                    SELECT * FROM terms
                    WHERE course_id = :course_id
                    ORDER BY is_current DESC, id ASC
                    LIMIT 1
                    """
                ),
                {"course_id": course["id"]},
            ).mappings().first()
            if term is None:
                continue
            row = dict(term)
            row["course_settings"] = course["settings_json"]
            row["term_settings"] = term["settings_json"]
            merged = settings_from_term_row(row)
            conn.execute(
                text("UPDATE courses SET settings_json = :settings WHERE id = :id"),
                {"settings": json.dumps(merged, ensure_ascii=False), "id": course["id"]},
            )
            restored += 1
    return {"courses_restored": restored}


def preview_term_fields(engine: Engine) -> dict:
    """只统计还将挪动的课程，不写库。"""
    if "courses" not in _tables(engine):
        return {"courses_with_identity": 0, "terms_missing_identity": 0}
    courses_with_identity = 0
    missing = 0
    has_terms = "terms" in _tables(engine)
    for course in _course_rows(engine):
        donor = _donor_identity(parse_settings(course["settings_json"]))
        if not _identity_present(donor):
            continue
        courses_with_identity += 1
        if not has_terms:
            continue
        with engine.connect() as conn:
            term_rows = conn.execute(
                text("SELECT * FROM terms WHERE course_id = :course_id"),
                {"course_id": course["id"]},
            ).mappings()
            for term in term_rows:
                blob = parse_settings(term.get("settings_json"))
                gaps = [
                    not str(term.get(key) or "").strip() and donor[key]
                    for key in ("year_start", "year_end", "semester", "school_year_term", "teacher", "class_name")
                ]
                if int(term.get("student_count") or 0) <= 0 and donor["student_count"]:
                    gaps.append(True)
                if int(term.get("exam_count") or 0) <= 0 and donor["exam_count"]:
                    gaps.append(True)
                if not str(blob.get("major") or "").strip() and donor["major"]:
                    gaps.append(True)
                if any(gaps):
                    missing += 1
    return {"courses_with_identity": courses_with_identity, "terms_missing_identity": missing}


def downgrade(engine: Engine) -> dict:
    """把学期字段写回课程设置，然后去掉 terms 和 term_id。文件仍留在原目录。"""
    if "terms" not in _tables(engine):
        _drop_term_column(engine)
        return {"courses_restored": 0}
    restored = 0
    with engine.begin() as conn:
        courses = conn.execute(text("SELECT id, settings_json FROM courses")).mappings().all()
        for course in courses:
            term = conn.execute(
                text(
                    """
                    SELECT * FROM terms
                    WHERE course_id = :course_id
                    ORDER BY is_current DESC, id ASC
                    LIMIT 1
                    """
                ),
                {"course_id": course["id"]},
            ).mappings().first()
            if term is None:
                continue
            row = dict(term)
            row["course_settings"] = course["settings_json"]
            row["term_settings"] = term["settings_json"]
            merged = settings_from_term_row(row)
            conn.execute(
                text("UPDATE courses SET settings_json = :settings WHERE id = :id"),
                {"settings": json.dumps(merged, ensure_ascii=False), "id": course["id"]},
            )
            restored += 1
        conn.execute(text("DROP TABLE terms"))
    _drop_term_column(engine)
    return {"courses_restored": restored}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="CalculatorPro v2 学期表迁移")
    parser.add_argument("--dry-run", action="store_true", help="只打印将影响的课程和文件数量")
    parser.add_argument("--downgrade", action="store_true", help="退回：学期字段写回课程设置并删除学期表")
    parser.add_argument(
        "--term-fields",
        action="store_true",
        help="把课程层遗留的学期字段拷到缺字段的学期上再删除；与 --downgrade 合用时只写回当前学期，不删表",
    )
    parser.add_argument("--backup-dir", default="", help="执行前把 courses、course_files、terms 导出为 JSON")
    parser.add_argument("--database-url", default="", help="默认读取 DATABASE_URL")
    args = parser.parse_args(argv)
    url = (args.database_url or os.environ.get("DATABASE_URL") or "").strip()
    if not url:
        if sqlite_allowed():
            url = database_url()
        else:
            print("需要 DATABASE_URL。SQLite 仅在 ALLOW_SQLITE=1 时允许。", file=sys.stderr)
            return 2
    engine = _engine(url)
    try:
        if args.dry_run and args.term_fields:
            print(json.dumps(preview_term_fields(engine), ensure_ascii=False))
            return 0
        if args.dry_run and args.downgrade:
            print(json.dumps({"downgrade_preview": "terms" in _tables(engine)}, ensure_ascii=False))
            return 0
        if args.dry_run:
            print(json.dumps(plan(engine), ensure_ascii=False))
            return 0
        if not args.backup_dir:
            print("执行迁移或退回必须提供 --backup-dir。", file=sys.stderr)
            return 2
        backup = backup_tables(engine, Path(args.backup_dir))
        print(json.dumps({"backup": str(backup)}, ensure_ascii=False))
        if args.term_fields and args.downgrade:
            print(json.dumps(downgrade_course_term_fields(engine), ensure_ascii=False))
        elif args.term_fields:
            print(json.dumps(upgrade_course_term_fields(engine), ensure_ascii=False))
        elif args.downgrade:
            print(json.dumps(downgrade(engine), ensure_ascii=False))
        else:
            print(json.dumps(apply_migration(engine), ensure_ascii=False))
        return 0
    finally:
        engine.dispose()


if __name__ == "__main__":
    raise SystemExit(main())
