"""网页报告：一次请求返回全部段落，由 core_app 解析后填进原模板。"""

from core_app.ai_report import generate_report_answers


def generate_answers(processor, num_objectives: int, report_style: str, word_limit: int) -> list[str]:
    return generate_report_answers(processor, num_objectives, report_style, word_limit)
