"""按北京时间计算下载当日的学年学期；学校模板第二行 A2 为学期标题。"""
from datetime import date, datetime, timedelta, timezone

CHINA_TIMEZONE = timezone(timedelta(hours=8))
TERM_CELL = "A2"


def current_academic_term(today: date | datetime | None = None) -> dict:
    if today is None:
        today = datetime.now(CHINA_TIMEZONE).date()
    elif isinstance(today, datetime):
        if today.tzinfo is not None:
            today = today.astimezone(CHINA_TIMEZONE)
        today = today.date()
    start = today.year if today.month >= 9 else today.year - 1
    semester = 1 if today.month >= 9 or today.month == 1 else 2
    return {
        "year_start": str(start),
        "year_end": str(start + 1),
        "semester": str(semester),
        "label": f"{start}-{start + 1}学年第{semester}学期",
        "template_text": f"{start}-{start + 1} 学年第 {semester} 学期",
    }
