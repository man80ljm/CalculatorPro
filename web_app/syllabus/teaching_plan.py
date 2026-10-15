"""Recover teaching arrangements from a course lesson plan without changing legacy rules."""
import re
from decimal import Decimal
from web_app.service import ServiceError


def to_markdown(document):
    from web_app.lesson_plans.engine import table_rows
    tables = [table_rows(table._tbl) for table in document.tables]
    lessons = []
    materials, sites = [], []
    for rows in tables:
        labels = " ".join(cell for row in rows[:2] for cell in row)
        if "授课题目" in labels and "授课课次" in labels and "授课学时" in labels:
            fields = {}
            for row in rows[:2]:
                for index in range(0,len(row)-1,2):
                    fields[re.sub(r"\s+","",row[index])] = row[index+1].strip()
            number = re.search(r"\d+",fields.get("授课课次",""))
            hours = fields.get("授课学时","").strip()
            if not number or not re.fullmatch(r"\d+(?:\.\d+)?",hours):
                raise ServiceError("教案中的课次或授课学时无法识别，请核对后上传。")
            detail = {}
            for row in rows[2:]:
                for cell in row:
                    label,sep,content = cell.partition("：")
                    if not sep: label,sep,content = cell.partition(":")
                    if sep: detail[label.strip()] = content.strip()
            goal = next((v for k,v in detail.items() if k.startswith("教学目标")),"")
            content = next((v for k,v in detail.items() if k.startswith("教学内容")),"")
            methods = next((v for k,v in detail.items() if k.startswith("教学手段")),"")
            lessons.append([str(int(number[0])),hours,fields.get("授课题目（教学章节或主题）",rows[0][-1]),
                "教学目标："+goal+"\n教学内容："+content,methods,"",""])
        for row in rows:
            if len(row)>1 and row[0].strip().rstrip("：:") in ("学习资料","学习网站"):
                dest = materials if "资料" in row[0] else sites
                dest.extend(re.sub(r"^\s*\[\d+\]\s*","",line).strip() for line in row[1].splitlines() if line.strip())
    if not lessons: return None
    if len(lessons)>100 or len({row[0] for row in lessons})!=len(lessons):
        raise ServiceError("教案课次应为不重复的正整数，最多支持 100 课次。")
    lessons.sort(key=lambda row:int(row[0]))
    paragraphs = [p.text.strip() for p in document.paragraphs if p.text.strip()]
    cover = "\n".join(paragraphs[:8])
    title = re.search(r"《([^》]+)》",cover)
    course = title[1] if title else ""
    total = sum(Decimal(row[1]) for row in lessons)
    source = "\n".join(paragraphs)
    scores = re.findall(r"(平时|期中|期末)考核占课程总成绩的\s*(\d+(?:\.\d+)?)\s*%",source)
    def table(number,rows):
        return "\n".join([f"[表格{number}]"]+["| "+" | ".join(" ".join(cell.split()).replace("|","／") for cell in row)+" |" for row in rows]+[f"[/表格{number}]"])
    parts = ["《"+course+"》课程教学资料",table(1,[["字段","内容"],["课程中文名称",course],["学时",str(total)]]),
        "二、课程简介","教学安排包括："+"；".join(row[2] for row in lessons),
        "六、教学内容及学时分配",table(2,[["课次","学时","章目名称","教学目标及教学内容","教学方式","支撑课程目标","课程思政"],*lessons]),
        "七、考核形式与成绩评定"]
    if scores:
        parts.append(table(3,[["项目","说明"],["成绩评定"," + ".join(name+"考核（"+ratio+"%）" for name,ratio in scores)]]))
    # Keep the original evaluation prose available to the model; it contains method details.
    evaluation = source.partition("课程评价方式")[2].strip()
    if evaluation: parts.append(evaluation)
    parts += ["八、参考资料","（一）参考书目",*materials,"（二）网络资源",*sites]
    return "\n".join(parts)
