"""报告一次返回全部段落：最多两次请求，解析失败时用原文作总体情况。"""

import json

from core_app.ai_report import generate_report_answers, report_max_tokens


class _Processor:
    def __init__(self):
        self.api_key = "sk-test"
        self.course_description = "一门练习课"
        self.objective_requirements = ["会做", "会讲"]
        self.previous_achievement_data = {"课程目标1": 0.6, "课程目标2": 0.5, "课程总目标": 0.55}
        self.current_achievement = {"课程目标1": 0.8, "课程目标2": 0.7, "总达成度": 0.75}


class _Response:
    def __init__(self, content):
        self._content = content

    def raise_for_status(self):
        return None

    def json(self):
        return {"choices": [{"message": {"content": self._content}}]}


def test_max_tokens_scales_with_segments():
    assert report_max_tokens(2, 200) == min(8000, max(1500, 5 * 200 * 3))
    assert report_max_tokens(5, 800) == 8000
    assert report_max_tokens(0, 10) == 1500


def test_one_call_fills_every_section(monkeypatch):
    payload = {
        "overall": "总体一段",
        "objectives": [
            {"analysis": "分析1", "improvement": "改进1"},
            {"analysis": "分析2", "improvement": "改进2"},
        ],
    }
    calls = []

    def fake_post(url, headers=None, json=None, timeout=None):
        calls.append({"url": url, "json": json})
        return _Response(json_dumps(payload))

    monkeypatch.setenv("DEEPSEEK_MODEL", "deepseek-flash")
    monkeypatch.setattr("core_app.ai_report.requests.post", fake_post)
    answers = generate_report_answers(_Processor(), 2, "专业", 200)
    assert len(calls) == 1
    assert calls[0]["json"]["model"] == "deepseek-flash"
    assert calls[0]["json"]["max_tokens"] == 3000
    user_prompt = calls[0]["json"]["messages"][1]["content"]
    assert "24000000001" not in user_prompt
    assert "张三" not in user_prompt
    assert answers == ["总体一段", "分析1", "改进1", "分析2", "改进2"]


def _capture_prompt(monkeypatch, processor):
    calls = []

    def fake_post(url, headers=None, json=None, timeout=None):
        calls.append(json)
        return _Response(json_dumps({"overall": "总体一段", "objectives": [{"analysis": "分析", "improvement": "改进"}, {"analysis": "分析2", "improvement": "改进2"}]}))

    monkeypatch.setattr("core_app.ai_report.requests.post", fake_post)
    generate_report_answers(processor, 2, "专业", 200)
    return calls[0]["messages"][1]["content"]


def test_missing_previous_year_is_not_treated_as_zero(monkeypatch):
    processor = _Processor()
    processor.previous_achievement_data = None
    prompt = _capture_prompt(monkeypatch, processor)
    assert "上一学年达成度: 0" not in prompt
    assert "无数据" in prompt
    assert "不要做同比" in prompt


def test_recorded_previous_year_includes_the_real_number(monkeypatch):
    prompt = _capture_prompt(monkeypatch, _Processor())
    assert "课程目标1上一学年达成度: 0.6" in prompt
    assert "0.55" in prompt
    assert "不要做同比" not in prompt


def test_previous_zero_stays_zero(monkeypatch):
    processor = _Processor()
    processor.previous_achievement_data = {"课程目标1": 0, "课程目标2": 0, "课程总目标": 0}
    prompt = _capture_prompt(monkeypatch, processor)
    assert "课程目标1上一学年达成度: 0" in prompt
    assert "无数据" not in prompt


def test_missing_previous_file_does_not_fill_zeros():
    from core_app.ai_report import AIReportMixin

    class _Box(AIReportMixin):
        def __init__(self):
            self.relation_payload = {"objectives": ["a"]}
            self.objective_requirements = ["会做"]
            self.status_label = None
            self.previous_achievement_data = {"课程目标1": 0}

    box = _Box()
    box.load_previous_achievement("")
    assert box.previous_achievement_data is None
    box.load_previous_achievement("/tmp/calculatorpro-missing-previous.xlsx")
    assert box.previous_achievement_data is None


def test_eval_docx_uses_dash_without_previous_year(tmp_path, monkeypatch):
    from docx import Document

    from core_app.word_exports import WordExportMixin

    monkeypatch.setattr("core_app.word_exports.get_outputs_dir", lambda: str(tmp_path))

    class _Export(WordExportMixin):
        pass

    _Export()._export_eval_result_docx(
        [{"name": "平时考核", "ratio": 1, "methods": [{"name": "作业", "supports": {"课程目标1": 1}}]}],
        ["课程目标1"],
        {"作业": 80},
        None,
        0.8,
        0.8,
        0,
    )
    path = tmp_path / "5.基于考核结果的课程目标达成情况评价结果表.docx"
    document = Document(path)
    text = "\n".join(cell.text for table in document.tables for row in table.rows for cell in row.cells)
    assert "—" in text
    assert "0.000" not in text


def test_bad_json_retries_once_then_keeps_raw_text(monkeypatch):
    calls = []

    def fake_post(url, headers=None, json=None, timeout=None):
        calls.append(json["model"])
        return _Response("这不是 JSON，只是一段总体说明。")

    monkeypatch.setenv("DEEPSEEK_MODEL", "custom-model")
    monkeypatch.setattr("core_app.ai_report.requests.post", fake_post)
    answers = generate_report_answers(_Processor(), 2, "专业", 200)
    assert calls == ["custom-model", "custom-model"]
    assert answers[0] == "这不是 JSON，只是一段总体说明。"
    assert answers[1:] == ["", "", "", ""]


def json_dumps(payload):
    return json.dumps(payload, ensure_ascii=False)
