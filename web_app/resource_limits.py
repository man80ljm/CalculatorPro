"""限量执行与输入检查。参数控制资源预算，不改变计算规则。"""
from __future__ import annotations
from contextlib import contextmanager
import io
import os
import threading
import zipfile
import re
from functools import wraps
from xml.etree import ElementTree


def setting(name, default, minimum=1, maximum=10000):
    try:
        value = int(os.environ.get(name, str(default)))
    except ValueError:
        value = default
    return min(max(value, minimum), maximum)


_compute = threading.BoundedSemaphore(2)
_requests_lock = threading.Lock()
_requests: dict[int, int] = {}


def admit_request(user_id: int) -> bool:
    with _requests_lock:
        if sum(_requests.values()) >= 20 or _requests.get(user_id, 0) >= 4:
            return False
        _requests[user_id] = _requests.get(user_id, 0) + 1
        return True


def release_request(user_id: int):
    with _requests_lock:
        count = _requests.get(user_id, 0) - 1
        if count > 0:
            _requests[user_id] = count
        else:
            _requests.pop(user_id, None)


def limited_computation(function):
    @wraps(function)
    def wrapped(*args, **kwargs):
        with computation_slot():
            return function(*args, **kwargs)
    return wrapped


@contextmanager
def computation_slot():
    _compute.acquire()
    try:
        yield
    finally:
        _compute.release()


def validate_document(data: bytes, filename: str = ""):
    from web_app.service import ServiceError
    if not zipfile.is_zipfile(io.BytesIO(data)):
        if str(filename).lower().endswith((".xlsx", ".docx")):
            label = "这份 Word 文件" if str(filename).lower().endswith(".docx") else "文件"
            raise ServiceError(f"无法读取{label}：内容与格式不符，请重新另存为 Excel 或 Word 文件后导入。")
        return
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            entries = archive.infolist()
            limit = setting("DOCUMENT_EXPANDED_MB", 100, maximum=500) * 1024 * 1024
            if len(entries) > 2000 or sum(item.file_size for item in entries) > limit:
                raise ServiceError("文件内容过大，请删除多余图片、空白表格或附件后再导入。", status=413)
            if any(item.flag_bits & 1 for item in entries):
                raise ServiceError("请先去掉文件密码，再导入。")
            # 不加载工作簿就先数实际单元格，不能相信可伪造的 dimension 属性。
            cells = 0
            for item in entries:
                if item.filename.startswith("xl/worksheets/") and item.filename.endswith(".xml"):
                    with archive.open(item) as handle:
                        for _, element in ElementTree.iterparse(handle, events=("end",)):
                            if element.tag.endswith("}c"):
                                reference = re.fullmatch(r"([A-Z]+)(\d+)", element.attrib.get("r", ""))
                                if reference:
                                    column = 0
                                    for letter in reference[1]:
                                        column = column * 26 + ord(letter) - 64
                                    if int(reference[2]) > 5000 or column > 200:
                                        raise ServiceError("表格超过 5000 行或 200 列，请删除多余空白行列后再导入。", status=413)
                                cells += 1
                                if cells > 250000:
                                    raise ServiceError("表格内容过多，请删除多余空白行列后再导入。", status=413)
                            element.clear()
    except (zipfile.BadZipFile, OSError, RuntimeError, ElementTree.ParseError) as exc:
        raise ServiceError("文件无法读取，请重新另存后导入。") from exc
