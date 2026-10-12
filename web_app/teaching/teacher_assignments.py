"""Keep each teacher's classes and teaching times in an independent record."""
import re
import unicodedata


def signature(value):
    value=unicodedata.normalize('NFKC',str(value))
    for full,short in [('美术教育','美教'),('环境设计','环艺'),('视觉传达','视觉'),('数字媒体','数媒')]:
        value=value.replace(full,short)
    return re.sub(r'[\s\W_]+','',value.replace('班',''))


def course_section(source,name):
    headings=list(re.finditer(r'《([^》]+)》',source))
    for index,heading in enumerate(headings):
        if signature(heading[1])==signature(name):
            return source[heading.start():headings[index+1].start() if index+1<len(headings) else len(source)]
    return '' if headings else source


def resolve_classes(label,choices):
    result=[];year='';prefix=''
    for part in re.split(r'[、,，]',label):
        part=re.sub(r'\s+','',part)
        if not part: continue
        match=re.match(r'^(\d{2})',part)
        if match:
            year=match[1]
            explicit=re.match(r'^\d{2}((?:美术教育|美教|环境设计|环艺|视觉传达|视觉|数字媒体|数媒))',part)
            if explicit: prefix=explicit[1]
        elif year:
            short=part.replace('班','')
            part=year+(prefix if prefix and re.fullmatch(r'[A-Z]+|实验|小学',short) else '')+part
        exact=[c for c in choices if signature(c)==signature(part)]
        if not exact and re.fullmatch(r'\d{2}(?:实验|小学)班?',part):
            exact=[c for c in choices if c.startswith(part[:2]) and part[2:].replace('班','') in c]
        if exact:
            result.extend(exact);continue
        letters=re.fullmatch(r'(.*?)([A-Z]{2,6})(?:班)?',part)
        expanded=[]
        if letters:
            for letter in letters[2]:
                found=[c for c in choices if signature(c)==signature(letters[1]+letter)]
                if not found: break
                expanded.extend(found)
            else:
                result.extend(expanded);continue
        result.append(part)
    return list(dict.fromkeys(result))


def read_weeks(source):
    result=set()
    for match in re.finditer(r'(\d{1,2}(?:[-—–~～]\d{1,2})?(?:[、,;；]\d{1,2}(?:[-—–~～]\d{1,2})?)*)周',source):
        for part in re.split(r'[、,;；]',match[1]):
            ends=re.split(r'[-—–~～]',part)
            start=int(ends[0]);end=int(ends[-1])
            if not 1<=start<=end<=30: return []
            result.update(range(start,end+1))
    return sorted(result)


def assigned_rows(row,source,choices):
    section=course_section(source,row['courseName'])
    lines=list(re.finditer(r'(?m)^\s*(\d{2}[^:\n：]{1,100})[:：]\s*',section))
    if not lines: return None
    header=section[:lines[0].start()]
    defaults=[t for t in row['teachers'] if re.search(r'[（(]\s*'+re.escape(t)+r'\s*[)）]',header)]
    result=[]
    for index,line in enumerate(lines):
        raw=section[line.start():lines[index+1].start() if index+1<len(lines) else len(section)].strip()
        body=section[line.end():lines[index+1].start() if index+1<len(lines) else len(section)]
        raw=re.split(r'(?m)^\s*\d{2}(?:级|美教|美术|书法|环艺|环境|视觉|数媒|数字)[^:：\n]*$',raw,1)[0].strip()
        body=re.split(r'(?m)^\s*\d{2}(?:级|美教|美术|书法|环艺|环境|视觉|数媒|数字)[^:：\n]*$',body,1)[0]
        body=re.sub(r'\s+','',unicodedata.normalize('NFKC',body))
        groups=list(re.finditer(r'\(([^()]*)\)',body))
        teachers=[t for t in row['teachers'] if any(t in re.split(r'[、,，]',g[1]) for g in groups)] or defaults
        weeks=read_weeks(body)
        if not teachers or not weeks: return None
        classes=resolve_classes(line[1],choices)
        rooms=[g[1] for g in groups if re.search(r'\d',g[1]) and not re.search(r'周|节|学分|学时',g[1])]
        excluded={g[0] for g in groups if g[1] in rooms or any(t in re.split(r'[、,，]',g[1]) for t in teachers)}
        pattern=body
        for group in excluded: pattern=pattern.replace(group,'')
        pattern=pattern.rstrip(';；')
        simple=re.fullmatch(r'\d{1,2}(?:[-—–~～]\d{1,2})?周\(([^()]*)\)',pattern)
        if simple: pattern=simple[1]
        clean_issues=[i for i in row.get('issues',[]) if not ('教师' in i and any(s in i for s in ('班级','分工','方向'))) and i not in {'教学班级未确定','教学周次未确定','班级名称需与原文复核','任课教师未确定'}]
        for teacher in teachers:
            result.append(dict(row,teachers=[teacher],classes=classes,weeks=weeks,timePattern=pattern,room='、'.join(dict.fromkeys(rooms)),sourceText=header.strip()+'\n'+raw,issues=clean_issues))
    return result


def split_teacher_records(records,pages=None):
    cells={c['id']:c for p in pages or [] for table in p.get('tables',[]) for c in table['cells']}
    choices=list(dict.fromkeys(c for cell in cells.values() for c in cell.get('classNames',[])))
    result=[]
    for row in records:
        if len(row['teachers'])<=1:
            result.append(row);continue
        anchor=cells.get(row.get('sourceCell'))
        source=anchor['text'] if anchor else row.get('sourceText','')
        parts=assigned_rows(row,source,choices or row['classes'])
        if parts:
            result.extend(parts);continue
        for teacher in row['teachers']:
            issues=list(row.get('issues',[]))
            if not row.get('reviewed') and '多位教师的授课分工需核对' not in issues:
                issues.append('多位教师的授课分工需核对')
            result.append(dict(row,teachers=[teacher],issues=issues))
    return result