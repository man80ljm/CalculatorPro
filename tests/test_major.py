"""上课专业从班级推导：全部例子和空值、纯数字、无前缀。"""
from web_app.syllabus.major import derive_major_from_class, resolve_majors

EXAMPLES = {
    "24数字媒体": "数字媒体",
    "2024级数字媒体": "数字媒体",
    "24级数字媒体": "数字媒体",
    "数字媒体（1）班": "数字媒体",
    "数字媒体(1)班": "数字媒体",
    "数字媒体1班": "数字媒体",
    "24数字媒体2班": "数字媒体",
    "23美术教育A": "美术教育",
    "23美术教育B": "美术教育",
    "数字媒体": "数字媒体",
}


def test_examples_and_edges():
    for raw, expected in EXAMPLES.items():
        assert derive_major_from_class(raw) == expected, raw
    assert derive_major_from_class("") is None
    assert derive_major_from_class("   ") is None
    assert derive_major_from_class("24") is None
    assert derive_major_from_class("2024") is None
    assert derive_major_from_class("2024级") is None
    assert derive_major_from_class(None) is None


def test_same_major_and_conflicting_majors():
    same = resolve_majors(["24数字媒体", "24数字媒体2班", "数字媒体（1）班"])
    assert same["status"] == "已填"
    assert same["value"] == "数字媒体"
    assert same["source"] == "成绩登记表·班级（推导）"
    split = resolve_majors(["24数字媒体", "23美术教育A"])
    assert split["status"] == "需手填"
    assert split["value"] == ""
    assert {item["value"] for item in split["candidates"]} == {"数字媒体", "美术教育"}
    empty = resolve_majors(["", "24", "2024级"])
    assert empty["status"] == "需手填"
    assert empty["candidates"] == []
