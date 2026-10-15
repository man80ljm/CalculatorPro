"""Administrator-owned staff directory: local PDF parsing, review, indexed lookup."""
import base64
import hashlib
import io
import json
import math
import re
import unicodedata
import uuid
from contextlib import contextmanager
from pathlib import Path

from sqlalchemy import func, or_, select, text
from sqlalchemy.orm import Session, load_only
from web_app.db import TeacherImport, TeacherProfile, utcnow
from web_app.resource_limits import validate_document, computation_slot
from .api_settings import ApiError

MAX_BYTES = 5 * 1024 * 1024
FIELDS = ("name", "college", "department", "teaching_unit", "group_name", "mobile", "office_phone", "notes")
MOBILE = re.compile(r"(?<!\d)(?:\+?86[-\s]*)?1[3-9]\d(?:[-\s]*\d){8}(?!\d)")


def key(value):
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(value or ""))).casefold()


def dump(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def clean(value):
    if not isinstance(value, dict) or set(value) - set(FIELDS):
        raise ApiError("教师资料字段不正确。")
    result = {}
    for field in FIELDS:
        item = value.get(field, "")
        maximum = 1000 if field == "notes" else 100 if field == "office_phone" else 60 if field == "mobile" else 160
        if not isinstance(item, str) or len(item) > maximum or re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", item):
            raise ApiError("教师资料格式不正确或内容过长。")
        result[field] = item.strip()
    if not result["name"] or not key(result["name"]): raise ApiError("教师姓名不能为空。")
    if result["mobile"] and not MOBILE.fullmatch(result["mobile"]):
        raise ApiError("手机号码应为有效的 11 位号码，可包含空格、连字符或 +86。")
    result["teaching_unit"] = result["teaching_unit"].removesuffix("系")
    return result


def parse_pdf(filename, raw):
    if not isinstance(filename, str) or Path(filename).suffix.lower() != ".pdf": raise ApiError("请选择 PDF 教师联系表。")
    if not raw or len(raw) > MAX_BYTES: raise ApiError("教师联系表不能超过 5MB。", 413)
    if not raw.startswith(b"%PDF-"): raise ApiError("文件内容不是有效的 PDF。")
    validate_document(raw, filename)
    rows, groups, warnings = [], [], []
    try:
        import pdfplumber
        with computation_slot(), pdfplumber.open(io.BytesIO(raw)) as pdf:
            if not 1 <= len(pdf.pages) <= 20: raise ApiError("每份联系表支持 1–20 页。")
            all_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
            compact = re.sub(r"\s+", "", unicodedata.normalize("NFKC", all_text))
            college_match = re.search(r"([\u3400-\u9fff]+学院)(?:教职工|教师|职工)", compact)
            college = college_match[1] if college_match else ""
            date_match = re.search(r"(20\d{2})年(1[0-2]|0?[1-9])月", compact)
            source_date = f"{date_match[1]}-{int(date_match[2]):02d}" if date_match else ""
            total_match = re.search(r"(?:联系表|通讯录|信息表)[(（](\d+)人[)）]", compact)
            expected = int(total_match[1]) if total_match else None
            for page_number, page in enumerate(pdf.pages, 1):
                words = page.extract_words()
                repeated_header = sum(key(word["text"]) == "姓名" for word in words) >= 2
                boxes = [(0, page.width/2), (page.width/2, page.width)] if repeated_header else [(0, page.width)]
                for left, right in boxes:
                    current = None
                    for line in (page.crop((left, 0, right, page.height)).extract_text() or "").splitlines():
                        normalized = unicodedata.normalize("NFKC", line).strip()
                        group_line = re.sub(r"\s+", "", normalized)
                        heading = re.fullmatch(r"(.+?(?:系|教研室)|院领导|办公室|.+部门|.+中心)(?:[(（](\d+)人[)）])?", group_line)
                        if heading:
                            current = {"name": heading[1], "expected": int(heading[2]) if heading[2] else None, "actual": 0, "page": page_number}
                            groups.append(current); continue
                        phone = MOBILE.search(normalized)
                        if not phone: continue
                        name = re.sub(r"\s+", "", normalized[:phone.start()])
                        if not re.fullmatch(r"[\u3400-\u9fffA-Za-z·]{2,80}", name):
                            warnings.append(f"第 {page_number} 页有一条姓名无法确定，请对照原文件补充。"); continue
                        group = current["name"] if current else ""
                        if not group: warnings.append(f"第 {page_number} 页的 {name} 未确定所属分组。")
                        department = group if group.endswith("系") else ""
                        unit = group.removesuffix("系") if department else group if group.endswith("教研室") else ""
                        fields = clean({"name": name, "college": college, "department": department, "teaching_unit": unit,
                                        "group_name": group, "mobile": phone.group(), "office_phone": normalized[phone.end():].strip()})
                        rows.append({"fields": fields, "source_page": page_number, "source_group": group, "source_text": normalized})
                        if current: current["actual"] += 1
                        if len(rows) > 2000: raise ApiError("一次最多导入 2000 条教师资料。", 413)
            if not rows: raise ApiError("未读取到教师记录。请使用可复制文字的联系表 PDF，或手动新增教师。")
            for group in groups:
                if group["expected"] is not None and group["expected"] != group["actual"]:
                    warnings.append(f"{group['name']} 标注 {group['expected']} 人，读取到 {group['actual']} 人，请核对。")
            if expected is not None and expected != len(rows): warnings.append(f"文件标注 {expected} 人，读取到 {len(rows)} 人，请核对。")
            return {"rows": rows, "original_rows": rows, "groups": groups, "warnings": list(dict.fromkeys(warnings)),
                    "expected_count": expected, "source_date": source_date, "pages": len(pdf.pages)}
    except ApiError: raise
    except Exception: raise ApiError("无法读取教师联系表，请检查 PDF 是否损坏或加密。") from None


def _public(row, sources):
    source = sources.get(row.source_import_id)
    return {"id": row.id, "revision": row.revision, "active": bool(row.active),
            **{field: getattr(row, field) for field in FIELDS},
            "source_filename": source.filename if source else "手动维护", "source_date": source.source_date if source else "",
            "source_page": row.source_page, "source_group": row.source_group,
            "updated_at": row.updated_at.isoformat(timespec="seconds")+"Z"}


def teacher_info(db, name, college=""):
    candidates = list(db.scalars(select(TeacherProfile).where(TeacherProfile.name_key == key(name), TeacherProfile.active == 1))) if name else []
    if college: candidates = [row for row in candidates if key(row.college) == key(college)]
    result = {"status": "not_found", "college": "", "department": "", "teaching_unit": "", "unit_source": ""}
    if len(candidates) > 1: return {**result, "status": "ambiguous"}
    if not candidates: return result
    row = candidates[0]
    unit = (row.teaching_unit or row.department).removesuffix("系")
    return {"status": "ready" if unit else "missing_unit", "college": row.college, "department": row.department,
            "teaching_unit": unit, "unit_source": "teaching_unit" if row.teaching_unit else "department" if row.department else ""}


class TeacherDirectory:
    def __init__(self, engine): self.engine = engine

    @contextmanager
    def connect(self, write=False):
        with Session(self.engine, expire_on_commit=False) as db:
            try:
                if write:
                    if self.engine.dialect.name == "sqlite": db.execute(text("BEGIN IMMEDIATE"))
                    elif self.engine.dialect.name == "postgresql": db.execute(text("SELECT pg_advisory_xact_lock(1735160213)"))
                yield db
                db.commit()
            except Exception:
                db.rollback(); raise

    def fingerprint(self, db):
        return hashlib.sha256(dump([(row.id, row.revision) for row in db.scalars(select(TeacherProfile).order_by(TeacherProfile.id))]).encode()).hexdigest()

    def matches(self, db, fields):
        return [row for row in db.scalars(select(TeacherProfile).where(TeacherProfile.name_key == key(fields["name"])))
                if key(row.college) == key(fields["college"])]

    def preview(self, db, batch):
        value = json.loads(batch.payload); rows = []; counts = {"new": 0, "update": 0, "skip": 0, "ambiguous": 0}
        sources = {row.id: row for row in db.scalars(select(TeacherImport).options(load_only(TeacherImport.id, TeacherImport.filename, TeacherImport.source_date, TeacherImport.status, TeacherImport.saved_at, TeacherImport.created_at)).order_by(TeacherImport.created_at.desc()))}
        identities = [(key(item["fields"]["name"]), key(item["fields"]["college"])) for item in value["rows"]]
        for index, item in enumerate(value["rows"]):
            candidates = self.matches(db, item["fields"])
            repeated = identities.count(identities[index]) > 1
            action = item.get("action") or ("ambiguous" if repeated or len(candidates)>1 else "update" if candidates else "new")
            target = item.get("target_id") or (candidates[0].id if action == "update" and len(candidates) == 1 else None)
            fields = dict(item["fields"])
            # A PDF grouped by departments does not overwrite a manually maintained teaching unit.
            old = next((row for row in candidates if row.id == target), None)
            if old and not item.get("unit_edited") and old.teaching_unit and old.teaching_unit != old.department.removesuffix("系") and key(old.department) == key(fields["department"]):
                fields["teaching_unit"] = old.teaching_unit
            if old and not item.get("notes_edited") and not fields["notes"]: fields["notes"] = old.notes
            counts[action] += 1
            rows.append({**item, "fields": fields, "index": index, "action": action, "target_id": target,
                         "candidates": [_public(row, sources) for row in candidates]})
        latest = db.scalar(select(TeacherImport).where(TeacherImport.status == "imported").order_by(TeacherImport.saved_at.desc()))
        warnings = list(value["warnings"])
        if latest and batch.source_date and latest.source_date and batch.source_date < latest.source_date:
            warnings.append("本文件资料日期早于最近导入，请核对后决定是否更新。")
        return {"id": batch.id, "revision": batch.revision, "filename": batch.filename, "status": batch.status,
                "source_date": batch.source_date, "rows": rows, "counts": counts, "groups": value["groups"],
                "warnings": warnings, "expected_count": value["expected_count"], "is_stale": batch.base_revision != self.fingerprint(db)}

    def snapshot(self, query="", department="", status="active", page=1):
        if len(query) > 200 or len(department) > 160 or status not in {"active", "inactive", "all"}: raise ApiError("查询条件不正确。")
        with self.connect() as db:
            statement = select(TeacherProfile)
            if status != "all": statement = statement.where(TeacherProfile.active == (1 if status == "active" else 0))
            if query:
                value = "%" + query.replace("/", "//").replace("%", "/%").replace("_", "/_") + "%"
                normalized = "%" + key(query).replace("/", "//").replace("%", "/%").replace("_", "/_") + "%"
                statement = statement.where(or_(TeacherProfile.name_key.like(normalized, escape="/"),
                    *(getattr(TeacherProfile, field).like(value, escape="/") for field in ("college", "department", "teaching_unit", "group_name", "mobile", "office_phone"))))
            if department: statement = statement.where(or_(TeacherProfile.department == department, TeacherProfile.group_name == department))
            total = db.scalar(select(func.count()).select_from(statement.subquery()))
            pages = max(1, math.ceil(total/20)); page = min(max(1, page), pages)
            all_rows = list(db.scalars(statement.order_by(TeacherProfile.department, TeacherProfile.name, TeacherProfile.id).offset((page-1)*20).limit(20)))
            sources = {row.id: row for row in db.scalars(select(TeacherImport).options(load_only(TeacherImport.id, TeacherImport.filename, TeacherImport.source_date, TeacherImport.status, TeacherImport.saved_at, TeacherImport.created_at)).order_by(TeacherImport.created_at.desc()))}
            departments = sorted(set(department or group for department, group in db.execute(select(TeacherProfile.department, TeacherProfile.group_name).distinct()) if department or group))
            pending = [{"id": row.id, "filename": row.filename, "source_date": row.source_date} for row in sources.values() if row.status == "preview"][:5]
            latest = max((row for row in sources.values() if row.saved_at), key=lambda row: row.saved_at, default=None)
            return {"records": [_public(row, sources) for row in all_rows], "total": total,
                    "page": page, "pages": pages, "departments": departments, "pending": pending,
                    "latest_import": {"filename": latest.filename, "source_date": latest.source_date} if latest else None}

    def upload(self, filename, encoded):
        if not isinstance(filename, str) or Path(filename).suffix.lower() != ".pdf": raise ApiError("请选择 PDF 教师联系表。")
        if not isinstance(encoded, str) or len(encoded) > (MAX_BYTES+2)//3*4: raise ApiError("教师联系表不能超过 5MB。", 413)
        try: raw = base64.b64decode(encoded, validate=True)
        except (ValueError, TypeError): raise ApiError("PDF 文件内容不正确。") from None
        digest = hashlib.sha256(raw).hexdigest()
        with self.connect() as db:
            previous = db.scalar(select(TeacherImport).where(TeacherImport.digest == digest))
            if previous: return self.preview(db, previous)
        value = parse_pdf(filename, raw)
        with self.connect(True) as db:
            previous = db.scalar(select(TeacherImport).where(TeacherImport.digest == digest))
            if previous: return self.preview(db, previous)
            batch = TeacherImport(id=uuid.uuid4().hex, filename=Path(filename.replace(chr(92), "/")).name[:200], digest=digest,
                payload=dump(value), source_date=value["source_date"], base_revision=self.fingerprint(db))
            db.add(batch); db.flush(); return self.preview(db, batch)

    def read_import(self, identity):
        with self.connect() as db:
            batch = db.get(TeacherImport, identity)
            if batch is None: raise ApiError("导入记录不存在。", 404)
            return self.preview(db, batch)

    def review(self, data):
        with self.connect(True) as db:
            batch = db.get(TeacherImport, str(data.get("id") or ""))
            if batch is None or batch.status != "preview": raise ApiError("这份导入记录已保存或不存在，请刷新列表。", 409)
            if data.get("revision") != batch.revision: raise ApiError("预览已在其他窗口修改，请重新打开。", 409)
            if data.get("refresh") is True:
                batch.base_revision = self.fingerprint(db)
            elif batch.base_revision != self.fingerprint(db): raise ApiError("教师库已更新，请重新匹配后再确认导入。", 409)
            value = json.loads(batch.payload)
            if "index" in data:
                index = data["index"]
                if type(index) is not int or not 0 <= index < len(value["rows"]): raise ApiError("导入记录编号不正确。")
                item = value["rows"][index]; fields = clean(data.get("fields"))
                action = data.get("action", "auto")
                if action not in {"auto", "new", "update", "skip"}: raise ApiError("请选择新增、更新或跳过。")
                target = data.get("target_id")
                candidates = self.matches(db, fields)
                if action == "update" and target not in {row.id for row in candidates}: raise ApiError("请选择同名教师记录进行更新。")
                if fields["teaching_unit"] != item["fields"]["teaching_unit"]: item["unit_edited"] = True
                if fields["notes"] != item["fields"]["notes"]: item["notes_edited"] = True
                item["fields"] = fields; item["action"] = action if action != "auto" else None; item["target_id"] = target if action == "update" else None
            batch.payload = dump(value); batch.revision += 1
            db.flush(); return self.preview(db, batch)

    def confirm(self, data):
        with self.connect(True) as db:
            batch = db.get(TeacherImport, str(data.get("id") or ""))
            if batch is None: raise ApiError("导入记录不存在。", 404)
            if batch.status == "imported": return {"saved": True, "duplicate": True}
            if data.get("revision") != batch.revision or batch.base_revision != self.fingerprint(db): raise ApiError("教师库或预览已更新，请重新匹配后再确认导入。", 409)
            preview = self.preview(db, batch)
            if preview["counts"]["ambiguous"]: raise ApiError("有同名教师待确认，请选择更新对象或新增独立记录。")
            touched = set()
            for item in preview["rows"]:
                if item["action"] == "skip": continue
                fields = clean(item["fields"])
                if item["action"] == "update":
                    row = db.get(TeacherProfile, item["target_id"])
                    if row is None or row.id in touched: raise ApiError("多条导入记录指向同一教师，请核对后再保存。")
                    touched.add(row.id); row.revision += 1
                else:
                    row = TeacherProfile(id=uuid.uuid4().hex, name=fields["name"], name_key=key(fields["name"]))
                    db.add(row)
                for field in FIELDS: setattr(row, field, fields[field])
                row.name_key = key(row.name); row.active = 1
                row.source_import_id = batch.id; row.source_page = item["source_page"]; row.source_group = item["source_group"]
            batch.status = "imported"; batch.saved_at = utcnow(); db.flush()
            return {"saved": True, "counts": preview["counts"]}

    def save_record(self, data):
        fields = clean(data.get("fields"))
        with self.connect(True) as db:
            if data.get("id"):
                row = db.get(TeacherProfile, str(data["id"]))
                if row is None: raise ApiError("教师资料不存在。", 404)
                if data.get("revision") != row.revision: raise ApiError("资料已在其他窗口更新，请重新打开。", 409)
                row.revision += 1
            else:
                if self.matches(db, fields) and data.get("allow_duplicate") is not True: raise ApiError("已有同名教师，请编辑已有记录；确为同名人员时勾选新增独立记录。", 409)
                row = TeacherProfile(id=uuid.uuid4().hex, name=fields["name"], name_key=key(fields["name"]))
                db.add(row)
            for field in FIELDS: setattr(row, field, fields[field])
            row.name_key = key(row.name); db.flush()
            return {"saved": True, "id": row.id}

    def set_status(self, data):
        if type(data.get("active")) is not bool: raise ApiError("教师状态不正确。")
        with self.connect(True) as db:
            row = db.get(TeacherProfile, str(data.get("id") or ""))
            if row is None: raise ApiError("教师资料不存在。", 404)
            if data.get("revision") != row.revision: raise ApiError("资料已在其他窗口更新，请刷新后重试。", 409)
            row.active = int(data["active"]); row.revision += 1
            return {"saved": True}
