"""课程考核与课程目标对应关系：粘贴解析与表格校验。

`parse_pasted_table` 与浏览器里的 `parsePastedTable`（static/relation_grid.js）保持同一规则：
按 Excel 的 TSV 规则：制表符分列、未加引号的换行分行，引号单元格可含换行/制表符/""，兼容 \\r\\n，末尾换行不额外产生空行，允许每行列数不同。
"""
from __future__ import annotations

MAX_ROWS = 300
MAX_COLS = 40
MAX_CELL = 2000

def parse_pasted_table(text: str) -> list[list[str]]:
    """把从 Excel 复制的文本（TSV）拆成行和单元格，规则同 Excel：

    - 列以制表符分隔，未加引号的换行分行；``\\r\\n``、``\\r`` 都视为 ``\\n``
    - 以 ``"`` 开头的单元格是引号单元格，内部可含换行（Alt+Enter）、制表符，``""`` 表示一个 ``"``
    - 引号没有闭合时按普通文本处理
    - 文本结尾的换行不会多出一行；中间的空行保留；各行的列数可以不同
    """
    if text is None:
        return []
    if not isinstance(text, str):
        raise TypeError("text must be a string")
    s = text.replace("\r\n", "\n").replace("\r", "\n")
    if s == "":
        return []
    n = len(s)

    def plain_end(start: int) -> int:
        k = start
        while k < n and s[k] not in "\t\n":
            k += 1
        return k

    rows: list[list[str]] = []
    row: list[str] = []
    i = 0
    while True:
        if i < n and s[i] == '"':
            j = i + 1
            buf: list[str] = []
            closed = False
            while j < n:
                if s[j] == '"':
                    if j + 1 < n and s[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    closed = True
                    j += 1
                    break
                buf.append(s[j])
                j += 1
            if closed:
                k = plain_end(j)
                field = "".join(buf) + s[j:k]
            else:
                k = plain_end(i)
                field = s[i:k]
        else:
            k = plain_end(i)
            field = s[i:k]
        i = k
        row.append(field)
        if i >= n:
            rows.append(row)
            break
        if s[i] == "\t":
            i += 1
            continue
        rows.append(row)
        row = []
        i += 1
        if i >= n:
            break
    return rows


def coerce_relation_grid(value) -> list[list[str]]:
    """校验提交上来的关系表。字符串按粘贴文本解析，列表则逐格转成字符串。"""
    if isinstance(value, str):
        rows = parse_pasted_table(value)
    elif isinstance(value, list):
        rows = []
        for row in value:
            if not isinstance(row, list):
                raise ValueError("对应关系表的每一行都必须是单元格列表")
            rows.append(["" if cell is None else str(cell) for cell in row])
    else:
        raise ValueError("对应关系表格式不正确")
    if len(rows) > MAX_ROWS:
        raise ValueError("对应关系表行数过多")
    for row in rows:
        if len(row) > MAX_COLS:
            raise ValueError("对应关系表列数过多")
        for cell in row:
            if len(cell) > MAX_CELL:
                raise ValueError("对应关系表单元格过长")
    return rows


def _parse_portion(text: str) -> float:
    raw = (text or "").strip().replace("％", "%")
    if not raw:
        raise ValueError("数字不能为空")
    percent = raw.endswith("%")
    if percent:
        raw = raw[:-1].strip()
    try:
        value = float(raw)
    except ValueError as exc:
        raise ValueError(f"不是有效数字：{text}") from exc
    if percent or value > 1:
        value = value / 100.0
    return value


def _cell(row: list[str], index: int | None) -> str:
    if index is None or index < 0 or index >= len(row):
        return ""
    return row[index].strip()


def _layout(rows: list[list[str]]):
    if not rows:
        raise ValueError("请先填写课程考核与课程目标对应关系")
    first = [_cell(rows[0], i) for i in range(len(rows[0]))]
    if first and first[0] == "考核环节":
        header = first
        data = rows[1:]

        def index_of(name: str) -> int | None:
            return header.index(name) if name in header else None

        link_i = index_of("考核环节")
        ratio_i = index_of("占比")
        method_i = index_of("考核方式")
        if method_i is None:
            method_i = 2 if ratio_i is not None else 1
        obj_i = [i for i, name in enumerate(header) if name.startswith("课程目标")]
        if not obj_i:
            raise ValueError("对应关系表缺少课程目标列")
        return data, link_i, ratio_i, method_i, obj_i
    width = max((len(row) for row in rows), default=0)
    # 无表头时列顺序固定为：考核环节、占比、考核方式、课程目标…
    if width < 4:
        raise ValueError("对应关系表至少需要考核环节、占比、考核方式和一个课程目标")
    return rows, 0, 1, 2, list(range(3, width))


def relation_payload_from_grid(rows: list[list[str]]) -> dict:
    """把关系表网格转成计算模块使用的 relation_payload。占比为 0 的环节不写入。"""
    data, link_i, ratio_i, method_i, obj_i = _layout(rows)
    groups: list[dict] = []
    current: dict | None = None

    for row in data:
        interesting = [link_i, ratio_i, method_i, *obj_i]
        if not any(_cell(row, i) for i in interesting):
            continue
        link_name = _cell(row, link_i)
        if link_name:
            if current is None or link_name != current["name"]:
                current = {"name": link_name, "ratio_text": _cell(row, ratio_i), "methods": []}
                groups.append(current)
            elif _cell(row, ratio_i) and not current["ratio_text"]:
                current["ratio_text"] = _cell(row, ratio_i)
        elif current is None:
            raise ValueError("请填写考核环节")
        elif _cell(row, ratio_i) and not current["ratio_text"]:
            current["ratio_text"] = _cell(row, ratio_i)
        assert current is not None
        method_name = _cell(row, method_i)
        weight_text = [_cell(row, i) for i in obj_i]
        if not method_name and not any(weight_text):
            continue
        if not method_name:
            raise ValueError("考核方式名称不能为空")
        weights = [0.0 if not text else _parse_portion(text) for text in weight_text]
        current["methods"].append({"name": method_name, "weights": weights})

    links = []
    for group in groups:
        if not group["ratio_text"]:
            raise ValueError(f"{group['name']} 缺少占比")
        ratio = _parse_portion(group["ratio_text"])
        if ratio < -1e-9 or ratio > 1 + 1e-9:
            raise ValueError(f"{group['name']} 的占比需要在 0 到 1 之间")
        if ratio <= 1e-9:
            continue
        if not group["methods"]:
            raise ValueError(f"{group['name']} 的占比大于 0，请至少填写一种考核方式")
        methods = []
        for method in group["methods"]:
            supports = {}
            subtotal = 0.0
            for index, weight in enumerate(method["weights"]):
                supports[f"课程目标{index + 1}"] = round(float(weight), 6)
                subtotal += float(weight)
            methods.append(
                {
                    "name": method["name"],
                    "supports": supports,
                    "subtotal": round(subtotal, 6),
                }
            )
        links.append({"name": group["name"], "ratio": round(ratio, 6), "methods": methods})

    if not links:
        raise ValueError("请先填写课程考核与课程目标对应关系")

    totals = {}
    total_sum = 0.0
    for index in range(len(obj_i)):
        key = f"课程目标{index + 1}"
        total = 0.0
        for link in links:
            for method in link["methods"]:
                total += float(method["supports"].get(key, 0.0)) * float(link["ratio"])
        totals[key] = round(total, 6)
        total_sum += totals[key]
    return {
        "objectives_count": len(obj_i),
        "links": links,
        "objectives_total_weights": totals,
        "total_sum": round(total_sum, 6),
    }


def ratios_from_payload(payload: dict) -> dict:
    usual = midterm = final = 0.0
    for link in payload.get("links") or []:
        name = str(link.get("name") or "")
        ratio = float(link.get("ratio") or 0)
        if "平时" in name:
            usual += ratio
        elif "期中" in name:
            midterm += ratio
        elif "期末" in name:
            final += ratio
    return {"usual": round(usual, 6), "midterm": round(midterm, 6), "final": round(final, 6)}


def absorb_relation_grid(settings: dict, strict: bool) -> dict:
    """若设置里带了关系表网格，先校验再尽量生成 relation_payload。

    strict 为 False 时，网格结构必须合法，但内容不完整仍可保存草稿。
    strict 为 True 时，网格必须能转成计算用的关系表。
    """
    if not isinstance(settings, dict):
        raise ValueError("课程设置格式不正确")
    grid = settings.get("relation_grid")
    if grid in (None, ""):
        return settings
    updated = dict(settings)
    updated["relation_grid"] = coerce_relation_grid(grid)
    try:
        payload = relation_payload_from_grid(updated["relation_grid"])
    except ValueError:
        if strict:
            raise
        return updated
    updated["relation_payload"] = payload
    updated["ratios"] = ratios_from_payload(payload)
    return updated
