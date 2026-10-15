"""编辑快照、字数提示、原模板渲染与参考页数检查。"""
import copy
import hashlib
import json
import re
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from web_app.teaching_calendar.academic_term import current_academic_term, CHINA_TIMEZONE
from . import engine
from .errors import LessonPlanError

META = ("课程名称","学年起","学年止","学期","教研室","授课对象","授课教师")
EVAL = ("平时比例","期末比例","平时项1","平时项2","评价方式","考查能力","评价说明","平时项1考查","平时项2考查","期末说明")


def text(value, label, limit=6000):
    if not isinstance(value,str) or len(value)>limit or re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]",value):
        raise LessonPlanError(label+"格式不正确或内容过长。")
    return value


def lines(value,label):
    if not isinstance(value,list) or len(value)>200: raise LessonPlanError(label+"必须为段落列表。")
    return [text(item,label) for item in value]


def initial_context(outline):
    term=current_academic_term()
    context={"课程名称":outline["课程名称"],"学年起":term["year_start"],"学年止":term["year_end"],"学期":term["semester"],
             "教研室":"","授课对象":"","授课教师":"","学习资料":outline["学习资料"],"学习网站":outline["学习网站"],
             "评价":engine.build_eval(outline,None),"lessons":[]}
    for item in outline["lessons"]:
        context["lessons"].append({"课次":int(engine.pick(item,"课次")),"学时":engine.pick(item,"学时"),
            "授课题目":engine.pick(item,"章目","章节"),"授课类型":engine.teaching_type(engine.pick(item,"教学方式")),**engine.rule_lesson(item)})
    context["_outline"] = outline
    return normalize_context(context)


def normalize_context(value):
    if not isinstance(value,dict): raise LessonPlanError("教案内容格式不正确。")
    try: size=len(json.dumps(value,ensure_ascii=False,allow_nan=False).encode())
    except (ValueError,TypeError): raise LessonPlanError("教案内容格式不正确。") from None
    if size>1024*1024: raise LessonPlanError("教案内容不能超过 1MB。",413)
    result={key:text(value.get(key,""),key,2000 if key=="授课对象" else 200) for key in META}
    if not re.fullmatch(r"\d{4}",result["学年起"]) or not re.fullmatch(r"\d{4}",result["学年止"]) or int(result["学年止"])!=int(result["学年起"])+1 or result["学期"] not in {"1","2"}:
        raise LessonPlanError("请填写相邻两个年份和第 1 或第 2 学期。")
    if not result["课程名称"].strip(): raise LessonPlanError("请填写课程名称。")
    raw=value.get("lessons")
    if not isinstance(raw,list) or not 1<=len(raw)<=100: raise LessonPlanError("教案应包含 1–100 个课次。")
    lessons=[]; seen=set()
    for item in raw:
        if not isinstance(item,dict) or type(item.get("课次")) is not int or item["课次"]<1 or item["课次"] in seen:
            raise LessonPlanError("课次必须为不重复的正整数。")
        seen.add(item["课次"])
        lesson={"课次":item["课次"],**{k:text(item.get(k,""),k,500) for k in ("学时","授课题目","授课类型")}}
        lesson.update({k:lines(item.get(k,[]),k) for k in engine.LIMITS}); lessons.append(lesson)
    result["lessons"]=sorted(lessons,key=lambda row:row["课次"])
    for key in ("学习资料","学习网站"): result[key]=lines(value.get(key,[]),key)
    evaluation=value.get("评价",{})
    if not isinstance(evaluation,dict): raise LessonPlanError("课程评价方式格式不正确。")
    result["评价"]={key:text(evaluation.get(key,""),key,1000) for key in EVAL}
    if "_outline" in value:
        from web_app.syllabus.compatibility import normalize_outline
        result["_outline"] = normalize_outline(value["_outline"])
    today=datetime.now(CHINA_TIMEZONE).date(); result.update(年=today.year,月=today.month,日=today.day)
    return result


def counts(context):
    return {str(row["课次"]):{key:{"count":engine.hanzi_len(row[key]),"min":low,"max":high,
                                   "in_range":low<=engine.hanzi_len(row[key])<=high}
                            for key,(low,high) in engine.LIMITS.items()} for row in context["lessons"]}


def template_digest():
    return hashlib.sha256(Path(engine.DEFAULT_TEMPLATE).read_bytes()).hexdigest()


def make_document(context, *, check=True):
    context=normalize_context(copy.deepcopy(context)); pages={}; notices=[]
    with TemporaryDirectory(prefix="calculatorpro-lesson-plan-") as directory:
        root=Path(directory); output=root/"lesson-plan.docx"
        engine.render(engine.DEFAULT_TEMPLATE,context,str(output))
        content=output.read_bytes()
        if check:
            try:
                pdf=engine.to_pdf(str(output),str(root))
                pages=engine.lesson_pages(pdf,len(context["lessons"]))
                if set(pages)!={row["课次"] for row in context["lessons"]}: notices.append("部分课次无法定位页数，请人工核对 Word。")
            except Exception:
                notices.append("本次页数检查未完成，可以重新检查；不影响 Word 导出。")
    notices.append("页数检查基于服务器字体，供参考；请在 Word 中核对最终排版。")
    return content,pages,notices
