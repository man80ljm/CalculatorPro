"""按桌面端 GenerateReportThread 的提示词生成表6，不依赖 PyQt。"""


def generate_answers(processor, num_objectives: int, report_style: str, word_limit: int) -> list[str]:
    prev_data = processor.previous_achievement_data or {}
    current_data = getattr(processor, "current_achievement", None) or {}

    questions = [
        "概述本次评价工作开展的总体情况，可以从教学目标、教学内容、教学方法、教学手段、教学策略、教学资源、教学环境等环节进行实质性分析。"
    ]
    for i in range(1, num_objectives + 1):
        questions.append(f"课程目标{i}达成情况分析")
        questions.append(f"课程目标{i}存在问题及改进措施")

    context = f"课程简介: {processor.course_description}\n"
    for i, req in enumerate(processor.objective_requirements or [], 1):
        context += f"课程目标{i}要求: {req}\n"
    for i in range(1, num_objectives + 1):
        prev_score = prev_data.get(f"课程目标{i}", 0)
        current_score = current_data.get(f"课程目标{i}", 0)
        context += f"课程目标{i}上一学年达成度: {prev_score}\n"
        context += f"课程目标{i}本学年达成度: {current_score}\n"

    prev_total = prev_data.get("课程总目标", 0)
    current_total = current_data.get("总达成度", 0)
    context += f"课程目标达成值（本学年）: {current_total}\n"
    context += "课程目标达成期望值: 0.7\n"
    context += f"上一轮教学课程目标达成值: {prev_total}\n"

    safe_limit = word_limit if word_limit and word_limit > 0 else 200
    min_chars = max(50, int(safe_limit * 0.8))
    answers = []
    for question in questions:
        prompt = (
            f"{context}\n问题: {question}\n"
            f"请以{report_style}风格回答，用一段话表述，不分点阐述，"
            f"不要使用Markdown或标题符号（如###、**等）。"
            f"字数尽量接近{safe_limit}字，不少于{min_chars}字。"
        )
        answers.append(processor.call_deepseek_api(prompt))
    return answers
