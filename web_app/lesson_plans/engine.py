"""复用用户的大纲读取、教案生成和固定 Word 模板渲染函数。

网站入口使用持久化后台任务；首次字数不符仅重写一次，超页时由用户手动压缩。
API 密钥仅来自后端环境，页数检查通过 LibreOffice headless。
"""
import copy
import datetime
import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from .errors import LessonPlanError
from concurrent.futures import ThreadPoolExecutor

import docx
from docx.oxml.ns import qn
from docx.shared import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = str(Path(__file__).with_name("templates") / "教案模板_占位符.docx")

DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"
DEEPSEEK_MODEL = "deepseek-chat"

# 每栏字数范围（不含序号和标点外的空白），保证每个课次正好一页
LIMITS = {"教学目标": (120, 180), "教学内容": (220, 300), "教学手段与方法": (100, 150), "思考讨论作业": (50, 80)}


class OutlineError(LessonPlanError):
    pass


# =====================================================================
# 1. 读取大纲
# =====================================================================
def cell_text(tc):
    lines = []
    for p in tc.iter(qn("w:p")):
        t = "".join(x.text or "" for x in p.iter(qn("w:t"))).strip()
        if t:
            lines.append(t)
    return "\n".join(lines)


def table_rows(tbl):
    return [[cell_text(tc) for tc in tr.findall(qn("w:tc"))] for tr in tbl.iter(qn("w:tr"))]


def read_outline(path):
    d = docx.Document(path)
    body = d.element.body
    tables = [table_rows(t) for t in body.iter(qn("w:tbl"))]
    paras = [p.text.strip() for p in d.paragraphs]

    info = {}
    for rows in tables:
        for r in rows:
            for i in range(0, len(r) - 1):
                k = r[i].replace(" ", "")
                if k in ("课程中文名称", "课程负责人", "教材选用", "学时", "考核方式") and k not in info:
                    info[k] = r[i + 1].strip()

    course = info.get("课程中文名称")
    if not course:
        m = re.search(r"《(.+?)》", "".join(paras[:3]))
        course = m.group(1) if m else ""

    # 课次表
    lessons = []
    for rows in tables:
        header = [h.replace("\n", "") for h in rows[0]] if rows else []
        if "课次" in header and "学时" in header:
            for r in rows[1:]:
                item = {header[i]: (r[i].strip() if i < len(r) else "") for i in range(len(header))}
                if item.get("课次", "").isdigit():
                    lessons.append(item)
            break
    if not lessons:
        raise OutlineError("没有在大纲里找到包含「课次」「学时」的教学内容表格。")

    # 课程目标、课程简介（作为 AI 的背景）
    goals = [p for p in paras if re.match(r"课程目标\s*\d", p)]
    intro = []
    on = False
    for p in paras:
        if re.match(r"^[二2][、.．]\s*课程简介", p):
            on = True
            continue
        if on and re.match(r"^[三3][、.．]", p):
            break
        if on and p:
            intro.append(p)

    # 参考资料
    sections = {"参考书目": [], "参考文献": [], "网络资源": []}
    cur = None
    started = False
    for p in paras:
        if re.match(r"^[六6][、.．]\s*参考资料", p):
            started = True
            continue
        if not started:
            continue
        hit = next((k for k in sections if k in p and re.match(r"^[（(]", p)), None)
        if hit:
            cur = hit
            continue
        if re.match(r"^(制定人|审订人|审批人)", p.replace(" ", "")) or re.match(r"^\d{4}\s*年", p):
            break
        if cur and p and p != "无":
            sections[cur].append(p)
    clean = lambda s: re.sub(r"^\s*[\[［]?\d+[\]］]?\s*[．.、]?\s*", "", s).strip().rstrip("．.")
    materials = [clean(x) for x in sections["参考书目"] + sections["参考文献"] if clean(x) and clean(x) != "无"]
    sites = []
    for x in sections["网络资源"]:
        x = clean(x)
        if sites and (sites[-1].endswith(("网址：", "网址:", "：", ":")) or re.match(r"^(https?://|www\.)", x)):
            sites[-1] = sites[-1] + x
        elif x and x != "无":
            sites.append(x)

    # 考核方式表
    assess = {"平时比例": "", "期末比例": "", "平时项": [], "期末项": [], "原表": ""}
    for rows in tables:
        flat = " ".join(" ".join(r) for r in rows)
        if "考核" in flat and "课程目标" in flat and ("平时" in flat or "期末" in flat):
            assess["原表"] = "\n".join(" | ".join(c.replace("\n", "") for c in r) for r in rows)
            part = None
            for r in rows:
                first = r[0].replace("\n", "") if r else ""
                if part and any(w in "".join(r[:2]) for w in ("总权重", "合计")) and not re.search(r"平时|期末|期中", first):
                    break
                m = re.search(r"(平时|期中|期末)\D*(\d+)\s*%", first)
                if m:
                    part = m.group(1)
                    if part in ("平时", "期末"):
                        assess[part + "比例"] = m.group(2)
                if part in ("平时", "期末") and len(r) > 1 and r[1].strip() and "考核方式" not in r[1]:
                    name = re.sub(r"完成$", "", r[1].strip())
                    if name != "无":
                        assess[part + "项"].append(name)
            break

    return {
        "课程名称": course,
        "课程负责人": info.get("课程负责人", ""),
        "课程简介": "".join(intro),
        "课程目标": goals,
        "lessons": lessons,
        "学习资料": materials,
        "学习网站": sites,
        "考核": assess,
    }


def pick(item, *keys):
    for k in keys:
        for name, v in item.items():
            if k in name:
                return v.strip()
    return ""


def teaching_type(method):
    practice = re.search(r"上机|实践|实操|实验|练习|项目|展示|实训|操作", method)
    lecture = re.search(r"讲授|演示|理论|讲解|案例|示范", method)
    if practice and lecture:
        return "讲授、实践"
    if practice:
        return "实践"
    return "讲授"


# =====================================================================
# 2. AI 生成
# =====================================================================
LESSON_PROMPT = """你是有经验的高校教师，正在根据课程教学大纲撰写某一次课的教案。请回答以下预设问题，严格依据大纲内容，不编造大纲之外的知识点，语言规范、具体、可操作，面向本科生。

预设问题与字数要求（字数指汉字数，必须落在范围内）：
1.【教学目标或要求】本次课学生在知识、能力、素养三方面应达到什么要求？分 3 条，以“掌握/理解/了解/能够/培养”等动词开头；有课程思政点时，素养目标要体现它。共 {教学目标} 字。
2.【教学内容】本次课讲授哪些内容？重点是什么？难点是什么、如何突破？
   分三部分：基本内容（按 1.2.3. 列出知识点，可带一句说明）；教学重点（1–2 条）；教学难点（1–2 条，每条后用“突破方法：…”说明）。三部分合计 {教学内容} 字。
3.【教学手段与方法】本次课具体采用哪些教学方法和手段？每种方法用在哪个环节、怎么做？结合大纲的“教学方式”，写 3 条左右，要具体到环节和操作。共 {教学手段与方法} 字。
4.【思考题、讨论题、作业】各出 1 道，与本次课内容紧密相关，作业要可提交、可评价。共 {思考讨论作业} 字。

严格输出 JSON，不要输出其他内容：
{{"教学目标":["1.…","2.…","3.…"],
 "基本内容":["1.…","2.…"],
 "教学重点":["…"],
 "教学难点":["…突破方法：…"],
 "教学手段与方法":["1.…","2.…","3.…"],
 "思考题":"…","讨论题":"…","作业":"…"}}"""

EVAL_PROMPT = """你是高校教务老师。根据课程教学大纲中的考核方式表，补全“课程评价方式”中的空白处。数字比例已由程序填好，你只写文字部分，每处 10–40 字，语句要能直接嵌入下面的句子：
本课程采用【评价方式】的评价方式，全面考查学生在课程学习中的【考查能力】能力。课程评价【评价说明】。
平时考核占课程总成绩的X%，主要包括{平时项1}和{平时项2}完成两个方面。{平时项1}重点考查学生的【平时项1考查】；{平时项2}重点考查学生【平时项2考查】。
期末考核占课程总成绩的Y%，主要【期末说明】。
注意：【期末说明】要说明期末考核的形式和各部分构成（依据考核表），以动词开头，如“以……形式进行，其中……”。
严格输出 JSON：{{"评价方式":"…","考查能力":"…","评价说明":"…","平时项1考查":"…","平时项2考查":"…","期末说明":"…"}}"""


def deepseek(system, user, api_key, temperature=0.5):
    from .client import complete_json
    return complete_json(system, user, temperature=temperature)


def hanzi_len(lines):
    if isinstance(lines, str):
        lines = [lines]
    return len(re.sub(r"\s", "", "".join(lines)))


def lesson_brief(outline, item):
    return json.dumps({
        "课程名称": outline["课程名称"],
        "课程简介": outline["课程简介"][:400],
        "课程目标": outline["课程目标"],
        "本次课": {
            "课次": pick(item, "课次"), "学时": pick(item, "学时"), "章目名称": pick(item, "章目", "章节"),
            "教学目标及教学内容": pick(item, "教学目标", "教学内容"), "教学方式": pick(item, "教学方式"),
            "支撑课程目标": pick(item, "支撑"), "课程思政": pick(item, "思政"),
        },
    }, ensure_ascii=False)


def ai_to_fields(a):
    """把 AI 返回的 JSON 整理成模板四栏（每栏是段落列表）。"""
    content = ["基本内容："] + list(a.get("基本内容", []))
    content += ["教学重点：" + "；".join(a.get("教学重点", []))]
    content += ["教学难点：" + "；".join(a.get("教学难点", []))]
    return {
        "教学目标": list(a.get("教学目标", [])),
        "教学内容": content,
        "教学手段与方法": list(a.get("教学手段与方法", [])),
        "思考讨论作业": [f"思考题：{a.get('思考题', '')}", f"讨论题：{a.get('讨论题', '')}", f"作业：{a.get('作业', '')}"],
    }


def gen_lesson(outline, item, api_key, extra=""):
    system = LESSON_PROMPT.format(**{k: f"{lo}–{hi}" for k, (lo, hi) in LIMITS.items()})
    user = lesson_brief(outline, item) + (("\n\n" + extra) if extra else "")
    fields = ai_to_fields(deepseek(system, user, api_key))
    # 字数不达标时，带着反馈再写一次
    bad = [f"{k}现在 {hanzi_len(v)} 字，应为 {LIMITS[k][0]}–{LIMITS[k][1]} 字"
           for k, v in fields.items() if not LIMITS[k][0] <= hanzi_len(v) <= LIMITS[k][1]]
    if bad:
        fields = ai_to_fields(deepseek(system, user + "\n\n上一版字数不符合要求：" + "；".join(bad) + "。请调整后重新输出。", api_key))
    return fields


def rule_lesson(item):
    """不用 AI：直接把大纲原文拆开放进去（仅用于测试排版）。"""
    detail = pick(item, "教学目标", "教学内容")
    g = re.search(r"教学目标[:：](.*?)(?=教学内容[:：]|$)", detail, re.S)
    c = re.search(r"教学内容[:：](.*)", detail, re.S)
    return {
        "教学目标": [g.group(1).strip() if g else ""],
        "教学内容": ["基本内容：" + (c.group(1).strip() if c else detail), "教学重点：", "教学难点："],
        "教学手段与方法": [pick(item, "教学方式")],
        "思考讨论作业": ["思考题：", "讨论题：", "作业："],
    }


def build_eval(outline, api_key, *, notices=None):
    a = outline["考核"]
    p1 = a["平时项"][0] if a["平时项"] else "课堂参与"
    p2 = a["平时项"][1] if len(a["平时项"]) > 1 else "作业"
    ev = {"平时比例": a["平时比例"], "期末比例": a["期末比例"], "平时项1": p1, "平时项2": p2,
          "评价方式": "过程性评价与终结性评价相结合", "考查能力": "知识掌握、实践应用与创新",
          "评价说明": "注重学习过程与作品质量的综合考查", "平时项1考查": "课堂出勤与参与讨论情况",
          "平时项2考查": "完成课程作业的质量与规范性", "期末说明": "以" + "、".join(a["期末项"] or ["设计作品"]) + "形式进行"}
    if api_key:
        try:
            ai = deepseek(EVAL_PROMPT.format(平时项1=p1, 平时项2=p2),
                          "考核方式表：\n" + a["原表"], api_key, temperature=0.3)
            for k in ("评价方式", "考查能力", "评价说明", "平时项1考查", "平时项2考查", "期末说明"):
                if str(ai.get(k, "")).strip():
                    ev[k] = str(ai[k]).strip().rstrip("。；")
        except Exception as e:
            if notices is not None:
                notices.append("课程评价方式 AI 生成失败，已使用原程序的默认句式，请核对。")
    return ev


# =====================================================================
# 3. 生成 Word 与页数检查
# =====================================================================
def render(template, ctx, out):
    from docxtpl import DocxTemplate
    tpl = DocxTemplate(template)
    tpl.render(ctx, autoescape=True)
    tpl.save(out)
    # Normalize the whole lesson table, including its title, number, type and hours.
    # Keep the source template and the cover's formatting intact.
    document = docx.Document(out)
    for table in document.tables:
        if len(table.rows) != 6 or "教学目标或要求" not in table.rows[2].cells[0].paragraphs[0].text:
            continue
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if not run.text:
                            continue
                        run.font.name = "仿宋_GB2312"
                        run.font.size = Pt(12)
                        fonts = run._element.get_or_add_rPr().rFonts
                        fonts.set(qn("w:eastAsia"), "仿宋_GB2312")
                        fonts.set(qn("w:cs"), "仿宋_GB2312")
    document.save(out)


def to_pdf(docx_path, outdir):
    from .office import convert_pdf
    return str(convert_pdf(Path(docx_path), Path(outdir)))


def lesson_pages(pdf, n):
    """返回 {课次: 占用页数}。依据每页顶部的「第 N 课次」标题定位。"""
    from pypdf import PdfReader
    pages = [p.extract_text() or "" for p in PdfReader(pdf).pages]
    starts = {}
    end = len(pages)
    for i, t in enumerate(pages):
        head = re.sub(r"\s", "", t)[:60]
        m = re.search(r"第(\d+)课次", head)
        if m and int(m.group(1)) not in starts:
            starts[int(m.group(1))] = i
        if "课程学习资料及网站" in re.sub(r"\s", "", t) and i > 0 and starts:
            end = min(end, i)
    result = {}
    keys = sorted(starts)
    for j, k in enumerate(keys):
        nxt = starts[keys[j + 1]] if j + 1 < len(keys) else end
        result[k] = nxt - starts[k]
    return result


