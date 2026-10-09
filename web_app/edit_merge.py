"""同字段冲突交给老师选择；互不影响的字段自动合并。"""
from copy import deepcopy

_MISSING = object()


def merge_edits(base, local, saved, path=""):
    if local == base:
        return deepcopy(saved), deepcopy(saved), []
    if saved == base or local == saved:
        return deepcopy(local), deepcopy(local), []
    if all(isinstance(item, dict) for item in (base, local, saved)):
        mine, theirs, conflicts = {}, {}, []
        for key in base.keys() | local.keys() | saved.keys():
            b, l, s = (item.get(key, _MISSING) for item in (base, local, saved))
            label = f"{path}.{key}" if path else key
            if l is _MISSING:
                # 页面未提交的后台元数据始终沿用已保存的值。
                if s is not _MISSING:
                    mine[key] = theirs[key] = deepcopy(s)
                continue
            if b is _MISSING or s is _MISSING:
                if s is _MISSING or l == s:
                    mine[key] = theirs[key] = deepcopy(l)
                else:
                    mine[key], theirs[key] = deepcopy(l), deepcopy(s)
                    conflicts.append({"field": label, "local": l, "saved": s})
                continue
            a, z, problems = merge_edits(b, l, s, label)
            mine[key], theirs[key] = a, z
            conflicts.extend(problems)
        return mine, theirs, conflicts
    return deepcopy(local), deepcopy(saved), [{"field": path, "local": local, "saved": saved}]
