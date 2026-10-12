"""Reconcile explicit course/class/time tuples against their physical PDF cells."""
import re,unicodedata
from .teacher_assignments import signature,assigned_rows,course_section,read_weeks

def compact_text(value):return re.sub(r'\s+','',unicodedata.normalize('NFKC',value))
def groups(value):return re.findall(r'\(([^()]*)\)',compact_text(value))
def teachers_in(value):
 return list(dict.fromkeys(g for g in groups(value) if re.fullmatch(r'[\u4e00-\u9fff]{2,4}',g) and g not in {'国画方向','油画方向','水彩方向','综合方向','版画方向','周末'}))
def blank_row(name,cell,page):
 return dict(courseName=name,teachers=[],classes=[],weeks=[],credits=None,totalHours=None,timePattern='',room='',sourcePage=page,sourceCell=cell['id'],sourceText=cell['text'],issues=[],outlineId='',reviewed=False)
def numbered_courses(cell,page):
 if not cell.get('classNames') or not cell.get('coveredWeeks'):return []
 source=cell['text'];starts=list(re.finditer(r'(?m)^\s*(\d+)\s*[.、]\s*',source))
 if not starts:return []
 header=source[:starts[0].start()];credit=re.search(r'(\d+(?:\.\d+)?)\s*学分',header);code=re.search(r'学分\s*([A-H])(?:\s|$)',header)
 result=[]
 for index,start in enumerate(starts):
  line=source[start.end():starts[index+1].start() if index+1<len(starts) else len(source)].strip();plain=compact_text(line);gs=groups(line)
  if not gs or not re.fullmatch(r'[\u4e00-\u9fff]{2,4}',gs[-1]):continue
  name=re.split(r'[（(]',re.sub(r'\s+','',line),1)[0]
  direction=next((g for g in gs if g.endswith('方向')),None)
  if direction:name+='（'+direction+'）'
  row=blank_row(name,cell,page);row.update(teachers=[gs[-1]],classes=list(cell.get('classNames',[])),weeks=list(cell.get('coveredWeeks',[])),credits=float(credit[1]) if credit else None,room=gs[-2] if len(gs)>=2 and gs[-2]!=direction and gs[-2]!='周末' else '',timePattern=cell.get('legends',{}).get(code[1],code[1]) if code else '',sourceText=header.strip()+'\n'+start[1]+'.'+line)
  if '周末' in gs:
   row['timePattern']='周末（原PDF未注明具体节次）';row['issues'].append('周末课程的具体节次未确定')
  for restriction in re.finditer(r'(实验班|书法班)(仅能选|只能选|不能选)(\d+(?:[-—]\d+)?)课程',header):
   ends=re.split('[-—]',restriction[3]);allowed=int(ends[0])<=int(start[1])<=int(ends[-1]);exclude=(restriction[2] in {'仅能选','只能选'} and not allowed) or (restriction[2]=='不能选' and allowed)
   if exclude:row['classes']=[c for c in row['classes'] if ('实验' if restriction[1]=='实验班' else '书法') not in c]
  year=re.search(r'(\d{2})级',header)
  if year and any(not c.startswith(year[1]) for c in row['classes']):row['issues'].append('原PDF选课标题年级与合并班级范围不一致，需核对选课范围')
  row['issues'].append('原PDF仅列可选班级范围，实际选课名单需核对')
  result.append(row)
 return result

def quoted_courses(cell,page,choices,existing):
 if cell.get('classNames') or '周' not in cell['text']:return []
 titles=re.findall(r'《([^》]+)》',cell['text']);result=[]
 for name in titles:
  section=course_section(cell['text'],name);ts=teachers_in(section)
  if not ts:continue
  old=[r for r in existing if signature(r['courseName'])==signature(name)]
  base=blank_row(name,cell,page)
  if old:base.update({k:v for k,v in old[0].items() if k not in {'id','courseName','sourceText','teachers','classes','weeks','timePattern','room','issues'}})
  base.update(teachers=ts,issues=[],sourceText=section.strip())
  credit=re.search(r'(\d+(?:\.\d+)?)\s*学分',section[:section.find('\n') if '\n' in section else len(section)])
  if credit:base['credits']=float(credit[1])
  parts=assigned_rows(base,cell['text'],choices)
  if parts:
   # A conflicting cohort is explicit in the source; keep it visible instead of inferring a correction.
   if name=='小学美术教学设计' and re.search(r'23美教实验班\s*《小学美术教学设计》',cell['text']):
    for part in parts:
     if any(c.startswith('22') for c in part['classes']):part['issues'].append('原PDF标题为23级实验班，授课行写22实验班，年级需确认')
   result.extend(parts)
  else:
   # Some elective notes name a teacher and timetable but no concrete class.
   for teacher in ts:
    compatible=[r for r in old if r['teachers']==[teacher]]
    row=dict(base,teachers=[teacher],weeks=read_weeks(section))
    if compatible:row['classes']=compatible[0]['classes'];row['reviewed']=compatible[0].get('reviewed',False)
    gs=groups(section);rooms=[g for g in gs if re.search(r'\d',g) and not re.search(r'周|节|学分|学时',g)]
    body=compact_text(section.split('》',1)[1]).lstrip(':')
    for g in gs:
     if g in rooms or g in ts:body=body.replace('('+g+')','')
    body=re.sub(r'^\([^)]*学分\)','',body).lstrip(':')
    simple=re.fullmatch(r'\d+(?:-\d+)?周\(([^()]*)\)',body)
    row['timePattern']=simple[1] if simple else body;row['room']='、'.join(rooms)
    if not row['classes']:row['issues'].append('教学班级未确定')
    if not row['weeks']:row['issues'].append('教学周次未确定')
    result.append(row)
 return result

def time_slots(row):
    pattern=compact_text(row['timePattern']);result=set()
    week_pattern=r'(\d{1,2}(?:[-—–~～]\d{1,2})?(?:[、,;；]\d{1,2}(?:[-—–~～]\d{1,2})?)*)周\(([^()]*)\)'
    phases=list(re.finditer(week_pattern,pattern))
    values=[(read_weeks(m[1]+'周'),m[2]) for m in phases] if phases else [(row['weeks'],pattern)]
    for weeks,time in values:
        for match in re.finditer(r'周([一二三四五六日天])第?(\d{1,2})(?:[-—–](\d{1,2}))?节',time):
            for week in weeks:
                for period in range(int(match[2]),int(match[3] or match[2])+1):result.add((week,match[1].replace('天','日'),period))
    return result

def preserve_phases(expected,existing):
    result=[]
    for row in expected:
        candidates=[r for r in existing if signature(r['courseName'])==signature(row['courseName']) and r['teachers']==row['teachers'] and {signature(c) for c in r['classes']}=={signature(c) for c in row['classes']}]
        expected_slots=time_slots(row)
        if len(candidates)>1 and expected_slots and set().union(*(time_slots(r) for r in candidates))==expected_slots:
            for candidate in candidates:
                result.append(dict(row,weeks=candidate['weeks'],timePattern=candidate['timePattern'],reviewed=candidate.get('reviewed',False)))
        else:result.append(row)
    return result

def grounded_groups(records,pages):
 """Return only source blocks with complete, explicit teacher assignments."""
 choices=list(dict.fromkeys([c for p in pages for table in p.get('tables',[]) for cell in table['cells'] for c in cell.get('classNames',[])]+[c for r in records for c in r['classes']]))
 result={}
 for page in pages:
  for table in page.get('tables',[]):
   for cell in table['cells']:
    existing=[r for r in records if r.get('sourceCell')==cell['id']]
    expected=numbered_courses(cell,page['page']) or quoted_courses(cell,page['page'],choices,existing)
    if not expected and existing and cell.get('classNames') and '周末' in groups(cell['text']):
     expected=[]
     for record in existing:
      row=dict(record,issues=list(record.get('issues',[])))
      if '周末' not in row['timePattern']:
       row['timePattern']+='；原文另注“周末”（实际时段待确认）'
      if '原PDF同时标工作日时段和周末，实际授课时间需确认' not in row['issues']:
       row['issues'].append('原PDF同时标工作日时段和周末，实际授课时间需确认')
      expected.append(row)
    if expected:result[cell['id']]=preserve_phases(expected,existing)
 return result

def reconcile_source_records(records,pages):
 if not pages:return records
 blocks=grounded_groups(records,pages);result=[dict(r,issues=list(r.get('issues',[]))) for r in records if r.get('sourceCell') not in blocks]
 cells={c['id']:c for p in pages for table in p.get('tables',[]) for c in table['cells']}
 for row in result:
  cell=cells.get(row.get('sourceCell'))
  if cell and not row.get('reviewed') and row['courseName'] in {'实习','毕业论文（创作）'}:
   relevant=cell['text'].split('研习')[0]
   listed=[name for group in re.findall(r'【([^】]+)】',relevant,re.S) for name in re.split(r'[、,，]',compact_text(group)) if name]
   if len(listed)>1:
    issue='原PDF只列'+('实习教师' if row['courseName']=='实习' else '毕业论文导师')+'名单，未明确教师与班级或学生分工'
    if issue not in row['issues']:row['issues'].append(issue)
 for cell,expected in blocks.items():
  existing=[r for r in records if r.get('sourceCell')==cell]
  for row in expected:
   old=next((r for r in existing if signature(r['courseName'])==signature(row['courseName']) and r['teachers']==row['teachers'] and {signature(c) for c in r['classes']}=={signature(c) for c in row['classes']}),None)
   if old:
    row['outlineId']=old.get('outlineId','');row['reviewed']=old.get('reviewed',False)
   result.append(row)
 return result