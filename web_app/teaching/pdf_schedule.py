"""Read a PDF locally, then ask the configured provider for structured records."""
import base64
import hashlib
import io
import json
import re
import threading
import time
import unicodedata
import urllib.error
import urllib.request
import uuid
from datetime import date, datetime, timezone
from .api_settings import ApiError, BASE_URL, NoRedirect
from .schedule_db import ScheduleDatabase
from .teacher_assignments import split_teacher_records

def now():
    return datetime.now(timezone.utc).isoformat()

def extract_pdf(raw):
    if len(raw) > 8 * 1024 * 1024 or not raw.startswith(b"%PDF-"):
        raise ApiError("请选择 8 MB 以内的有效 PDF 文件。")
    try:
        import pdfplumber
    except ImportError:
        raise ApiError("服务缺少 PDF 解析组件，请联系管理员。",503) from None
    try:
        with pdfplumber.open(io.BytesIO(raw)) as pdf:
            if not 1 <= len(pdf.pages) <= 15:
                raise ApiError("每份 PDF 支持 1—15 页。")
            pages=[]
            for number,page in enumerate(pdf.pages,1):
                text=page.extract_text() or ""
                tables=[]
                for table in page.find_tables()[:8]:
                    rows=table.extract()
                    boxes=[cell for row in table.rows for cell in row.cells if cell]
                    xs=sorted(set(round(x,2) for box in boxes for x in (box[0],box[2])))
                    ys=sorted(set(round(y,2) for box in boxes for y in (box[1],box[3])))
                    cells=[]
                    for ri,row in enumerate(table.rows):
                        for ci,box in enumerate(row.cells):
                            value=rows[ri][ci]
                            if box and value:
                                x0,y0,x1,y1=[round(v,2) for v in box]
                                cells.append({"row":ys.index(y0)+1,"column":xs.index(x0)+1,
                                    "rowSpan":ys.index(y1)-ys.index(y0),"columnSpan":xs.index(x1)-xs.index(x0),
                                    "text":value})
                    # Week numbers and class labels are read from their physical cells,
                    # independently of the model's interpretation of merged grid columns.
                    physical=[(box,rows[ri][ci]) for ri,row in enumerate(table.rows)
                              for ci,box in enumerate(row.cells) if box and rows[ri][ci]]
                    week_headers=[(box,int(value.strip())) for box,value in physical if re.fullmatch(r"(?:[1-9]|[12][0-9])",value.strip())]
                    bands={}
                    for box,week in week_headers: bands.setdefault(round(box[1],1),[]).append((box,week))
                    headers=max(bands.values(),key=len,default=[])
                    if len(headers)<3 or sorted(week for _,week in headers)!=list(range(1,len(headers)+1)): headers=[]
                    first=min((box[0] for box,_ in headers),default=0)
                    header_bottom=max((box[3] for box,_ in headers),default=0)
                    label_box=next((box for box,value in physical if ("班级" in value or value.strip().lower() in {"class","classes"}) and abs(box[2]-first)<2),None)
                    class_rows=[(box,value.replace("\n","")) for box,value in physical if label_box
                        and abs(box[0]-label_box[0])<2 and abs(box[2]-label_box[2])<2 and box[1]>=header_bottom
                        and len(value.strip())<=100 and not value.strip().isdigit()]
                    legend_text=unicodedata.normalize("NFKC","".join(value for box,value in physical if box[2]<=first and re.search(r"A\s*[:：]",value)))
                    legends={match.group(1):re.sub(r"\s+","",match.group(2)) for match in re.finditer(r"([A-Z])\s*:\s*(.*?)(?=[A-Z]\s*:|$)",legend_text,re.S)}
                    for cell_index,cell in enumerate(cells):
                        box=next(box for box,value in physical if value==cell["text"] and ys.index(round(box[1],2))+1==cell["row"] and xs.index(round(box[0],2))+1==cell["column"])
                        classes=[value for label,value in class_rows if min(box[3],label[3])-max(box[1],label[1])>1]
                        cell["id"]="p"+str(number)+"-t"+str(len(tables)+1)+"-c"+str(cell_index+1)
                        if classes and box[0]>=first-2:
                            cell["classNames"]=classes
                            cell["coveredWeeks"]=[week for header,week in headers if box[0]-1<=(header[0]+header[2])/2<=box[2]+1]
                        cell["legends"]=legends
                    tables.append({"rows":len(ys)-1,"columns":len(xs)-1,"cells":cells})
                pages.append({"page":number,"text":text,"tables":tables})
            if sum(len(p["text"]) for p in pages)<30:
                raise ApiError("此 PDF 没有可读取的文字层。当前支持文字型 PDF；扫描版需接入图像识别。")
            if len(json.dumps(pages,ensure_ascii=False))>180000:
                raise ApiError("PDF 内容超过读取范围，请使用院系课表。")
            return pages
    except ApiError:
        raise
    except Exception:
        raise ApiError("PDF 无法解析，请检查文件是否损坏或已加密。") from None

def text(value, limit=300):
    return str(value or "").strip()[:limit]

def items(value, limit=80):
    if not isinstance(value,list): return []
    return list(dict.fromkeys(text(v,160) for v in value[:limit] if isinstance(v,str) and v.strip()))

def normalized(value):
    return re.sub(r"[\s\W_]+","",unicodedata.normalize("NFKC",str(value)))

def source_cell(value,pages):
    if not pages: return None
    cells=[cell for page in pages for table in page.get("tables",[]) for cell in table["cells"]]
    identifier=value.get("sourceCell")
    if identifier:
        exact=next((cell for cell in cells if cell["id"]==identifier),None)
        if exact: return exact
    name=normalized(value.get("courseName",""))
    if not name: return None
    candidates=[]
    for cell in cells:
        content=normalized(cell["text"])
        if name not in content: continue
        teacher_hits=sum(normalized(t) in content for t in items(value.get("teachers")))
        class_hits=sum(any(normalized(c) in normalized(actual) or normalized(actual) in normalized(c)
                           for actual in cell.get("classNames",[])) for c in items(value.get("classes")))
        score=teacher_hits+class_hits*3+(2 if cell.get("classNames") else 0)
        candidates.append((score,cell))
    candidates.sort(key=lambda item:item[0],reverse=True)
    if len(candidates)==1 or (candidates and candidates[0][0]>(candidates[1][0] if len(candidates)>1 else -1)):
        return candidates[0][1]
    return None

def class_signature(value):
    value=unicodedata.normalize("NFKC",str(value))
    for full,short in [("美术教育","美教"),("环境设计","环艺"),("视觉传达","视觉"),("数字媒体","数媒")]:
        value=value.replace(full,short)
    return normalized(value.replace("班",""))

def course_in_source(course,source,anchor):
    if normalized(course) in source: return True
    direction=re.fullmatch(r"(.*?)[（(]([^()（）]*方向)[）)]",course)
    return bool(anchor and direction and normalized(direction[1]) in normalized(anchor["text"]) and normalized(direction[2]) in normalized(anchor["text"]))

def expand_named_class_assignments(records,pages):
    result=[]
    for row in records:
        anchor=source_cell(row,pages)
        marker=re.search(r"研习\s*[【\[]([^】\]]+)[】\]]",anchor["text"],re.S) if anchor and row["courseName"]=="研习" else None
        assigned=[]
        if marker and len(row["teachers"])>1:
            for teacher in row["teachers"]:
                match=re.search(re.escape(teacher)+r"\s*([A-Z])(?:班)?(?:[、和]*(实验班|小学班))?",marker[1])
                if not match: break
                classes=[c for c in row["classes"] if re.search(match[1]+r"(?:班)?$",c)]
                if match[2]: classes.extend(c for c in row["classes"] if match[2].replace("班","") in c)
                if not classes: break
                assigned.append(dict(row,teachers=[teacher],classes=classes,issues=[s for s in row["issues"] if s!="多位教师的授课分工需核对"]))
            if len(assigned)==len(row["teachers"]) and {c for part in assigned for c in part["classes"]}==set(row["classes"]):
                result.extend(assigned);continue
        result.append(row)
    return result

def normalize_result(payload,pages=None):
    if not isinstance(payload,dict) or not isinstance(payload.get("records"),list):
        raise ApiError("API 返回的课表结构不完整，请重新识别。",502)
    raw_term=payload.get("term") if isinstance(payload.get("term"),dict) else {}
    academic=text(raw_term.get("academicYear"),20).replace("—","-")
    if not re.fullmatch(r"20\d{2}-20\d{2}",academic) or int(academic[5:])!=int(academic[:4])+1:
        raise ApiError("未能确定学年，请核对 PDF 标题后重新识别。",502)
    semester=raw_term.get("semester")
    if isinstance(semester,bool) or semester not in (1,2): raise ApiError("未能确定学期，请核对 PDF 标题后重新识别。",502)
    start=text(raw_term.get("startDate",raw_term.get("start")),10)
    try:
        start = start if date.fromisoformat(start).weekday()==0 else ""
    except ValueError: start=""
    count=raw_term.get("weeks")
    if not isinstance(count,int) or isinstance(count,bool) or not 1<=count<=30: count=30
    term={"id":academic[:4]+"-"+str(semester),"academicYear":academic,"semester":semester,
          "label":academic.replace("-","—")+"学年第"+str(semester)+"学期","start":start,"weeks":count,"skip":[]}
    if not 1<=len(payload["records"])<=1000: raise ApiError("没有识别到课程，或记录数量超过支持数量。",502)
    records=[];seen=set()
    all_text="\n".join(page["text"] for page in pages) if pages else None
    for value in payload["records"]:
        if not isinstance(value,dict): raise ApiError("部分课程记录格式不正确。",502)
        course=text(value.get("courseName"),160)
        if not course: raise ApiError("识别结果包含没有课程名称的记录，请重试。",502)
        if re.fullmatch(r"\d{2}级.*选课[①②③④ⅠⅡⅢIV1234]*",course): continue
        anchor=source_cell(value,pages)
        teachers=items(value.get("teachers")); classes=items(value.get("classes"))
        weeks=value.get("weeks") or []
        if anchor and anchor.get("classNames"):
            if not classes or not set(classes).issubset(set(anchor["classNames"])): classes=anchor["classNames"]
            covered=anchor.get("coveredWeeks",[])
            if covered and (not isinstance(weeks,list) or not weeks or not set(weeks).issubset(set(covered))): weeks=covered
        if not isinstance(weeks,list): weeks=[]
        weeks=sorted(set(w for w in weeks if isinstance(w,int) and not isinstance(w,bool) and 1<=w<=count))
        automatic={"课程名称需与原文复核","教师姓名需与原文复核","班级名称需与原文复核",
                   "教学周次未确定","教学班级未确定","任课教师未确定","学期起始日期未确定","多位教师的授课分工需核对"}
        issues=[issue for issue in items(value.get("issues"),15) if issue not in automatic]
        if weeks and anchor and anchor.get("coveredWeeks"):
            issues=[issue for issue in issues if "周次" not in issue and "教学周" not in issue]
        pattern=text(value.get("timePattern"),1200)
        if anchor: pattern=anchor.get("legends",{}).get(pattern,pattern)
        if not teachers: issues.append("任课教师未确定")
        if not classes: issues.append("教学班级未确定")
        if not weeks: issues.append("教学周次未确定")
        if len(teachers)>1 and not value.get("reviewed"): issues.append("多位教师的授课分工需核对")
        if not start: issues.append("学期起始日期未确定")
        credits=value.get("credits")
        if not isinstance(credits,(int,float)) or isinstance(credits,bool) or not 0<credits<=30: credits=None
        hours=value.get("totalHours")
        if not isinstance(hours,int) or isinstance(hours,bool) or not 1<=hours<=500: hours=None
        if pages and "学时" not in all_text: hours=None
        page=value.get("sourcePage")
        if not isinstance(page,int) or isinstance(page,bool) or page<1 or (pages and page>len(pages)): page=None;issues.append("原文页码未确定")
        if pages and page:
            page_source=pages[page-1]["text"]+" ".join(cell["text"] for table in pages[page-1].get("tables",[]) for cell in table["cells"])
            source=normalized(page_source)
            if not course_in_source(course,source,anchor): issues.append("课程名称需与原文复核")
            if any(normalized(t) not in source for t in teachers): issues.append("教师姓名需与原文复核")
            # Physical class labels remain authoritative when the text layer omits them.
            anchored_classes={class_signature(c) for c in anchor.get("classNames",[])} if anchor else set()
            class_source=class_signature(page_source)
            if any(class_signature(c) not in anchored_classes and class_signature(c) not in class_source for c in classes): issues.append("班级名称需与原文复核")
        row={"courseName":course,"teachers":teachers,"classes":classes,"weeks":weeks,"credits":credits,
             "totalHours":hours,"timePattern":pattern,"room":text(value.get("room"),160),
             "sourcePage":page,"sourceText":text(anchor["text"] if anchor else value.get("sourceText"),700),"sourceCell":anchor["id"] if anchor else text(value.get("sourceCell"),60),"issues":list(dict.fromkeys(issues)),
             "outlineId":text(value.get("outlineId"),100),"reviewed":bool(value.get("reviewed")) if pages is None else False}
        key=json.dumps([course,teachers,classes,weeks,row["timePattern"],row["room"]],ensure_ascii=False)
        if key in seen: continue
        seen.add(key);records.append(row)
    from .source_grounding import reconcile_source_records
    records=reconcile_source_records(records,pages)
    records=split_teacher_records(expand_named_class_assignments(records,pages),pages)
    if not records: raise ApiError("没有提取到独立课程，请核对课表内容后重试。",502)
    notes=items(payload.get("notes"),20)
    if all(row["weeks"] for row in records):
        notes=[note for note in notes if "未明确周次" not in note and "周次未明确" not in note]
    if pages and any(row.get("sourceCell") for row in records):
        notes.append("课程块的周次和班级已按 PDF 实际表格范围读取；选修名单与多教师分工仍需核对。")
    return {"term":term,"records":records,"notes":list(dict.fromkeys(notes))}

def recognize(pages,config,opener=None,progress=None,cache=None):
    from .schedule_recognition import recognize_batches
    if opener is None:
        from .provider import ProviderOpener
        opener = ProviderOpener(config)
    return recognize_batches(pages,config,opener,progress,cache)

class PdfSchedule:
    def __init__(self,repository,settings):
        self.db=repository
        self.db.recover_interrupted()
        self.settings=settings
        self.lock=threading.Lock()
        self.active=set()
    def start(self,body):
        name=text(body.get("filename"),200).replace("\\","/").split("/")[-1]
        if not name.lower().endswith(".pdf"): raise ApiError("请选择 PDF 课表。")
        encoded=body.get("pdf")
        if not isinstance(encoded,str) or len(encoded)>12*1024*1024: raise ApiError("PDF 文件过大。",413)
        try: raw=base64.b64decode(encoded,validate=True)
        except Exception: raise ApiError("PDF 上传内容无效。") from None
        if len(raw)>8*1024*1024 or not raw.startswith(b"%PDF-"): raise ApiError("请选择 8 MB 以内的有效 PDF 文件。")
        digest=hashlib.sha256(raw).hexdigest()
        existing=self.db.find_source(digest)
        if existing: return existing
        config=self.settings.resolve({})
        with self.lock:
            if len(self.active)>=1: raise ApiError("已有一份 PDF 正在识别，请稍后再上传。",429)
            identifier=uuid.uuid4().hex
            self.db.create_draft(identifier,name,digest,config["model"])
            self.active.add(identifier)
        worker=threading.Thread(target=self.process,args=(identifier,raw,config),daemon=True)
        worker.start()
        return {"jobId":identifier,"reused":False}
    def process(self,identifier,raw,config):
        try:
            pages=extract_pdf(raw)
            self.db.update_draft(identifier,"recognizing")
            def progress(value):
                self.db.update_draft(identifier,"recognizing",{"progress":value,"pages":len(pages)})
            result,usage=recognize(pages,config,progress=progress,cache=self.db)
            # Usage contains no credentials and is kept only as provider metadata.
            safe_usage={k:v for k,v in usage.items() if k in {"prompt_tokens","completion_tokens","total_tokens"} and isinstance(v,int)}
            self.db.update_draft(identifier,"ready",dict(result,pages=len(pages),usage=safe_usage))
        except ApiError as exc:
            self.db.update_draft(identifier,"error",error=str(exc))
        except Exception:
            self.db.update_draft(identifier,"error",error="PDF 识别失败，未保存课程。")
        finally:
            with self.lock: self.active.discard(identifier)
    def job(self,identifier):
        return self.db.draft(identifier)
    def save(self,body):
        identifier=text(body.get("jobId"),50)
        job=self.db.draft(identifier)
        if job["status"]=="saved": return self.db.snapshot()
        if job["status"]!="ready": raise ApiError("识别结果尚未准备好。")
        result=normalize_result(body.get("result"))
        # Retain the original page count and never trust browser-provided API metadata.
        if result["term"]["id"]!=job["result"]["term"]["id"]: raise ApiError("请保留识别到的学年学期；不同学期应分别导入。")
        for row in result["records"]:
            if row["sourcePage"] and row["sourcePage"]>job["result"]["pages"]: raise ApiError("原文页码超出 PDF 页数。")
        self.db.save_draft(identifier,result)
        return self.db.snapshot()
    def edit(self,body):
        self.db.edit_record(text(body.get("id"),60),body.get("record"))
        return self.db.snapshot()
