/*
 * 与 web_app/relation_grid.py 的 parse_pasted_table 保持同一规则（Excel 复制出的 TSV）：
 * - 制表符分列，未加引号的换行分行；\r\n、\r 都视为 \n
 * - 以 " 开头的单元格是引号单元格，内部可含换行（Alt+Enter）、制表符，"" 表示一个 "
 * - 引号没有闭合时，按普通文本处理（" 保留为字面字符）
 * - 结尾的换行不会多出一行；中间空行保留；各行列数可以不同
 */
function parsePastedTable(text) {
  if (text == null) return [];
  const s = String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  if (s === "") return [];
  const n = s.length;
  const rows = [];
  let row = [];
  let i = 0;
  const plainEnd = (from) => {
    let k = from;
    while (k < n && s[k] !== "\t" && s[k] !== "\n") k += 1;
    return k;
  };
  for (;;) {
    let field;
    if (s[i] === '"') {
      let j = i + 1;
      let buf = "";
      let closed = false;
      while (j < n) {
        if (s[j] === '"') {
          if (s[j + 1] === '"') { buf += '"'; j += 2; continue; }
          closed = true; j += 1; break;
        }
        buf += s[j]; j += 1;
      }
      if (closed) {
        const k = plainEnd(j);
        field = buf + s.slice(j, k);
        i = k;
      } else {
        const k = plainEnd(i);
        field = s.slice(i, k);
        i = k;
      }
    } else {
      const k = plainEnd(i);
      field = s.slice(i, k);
      i = k;
    }
    row.push(field);
    if (i >= n) { rows.push(row); break; }
    if (s[i] === "\t") { i += 1; continue; }
    rows.push(row);
    row = [];
    i += 1;
    if (i >= n) break;
  }
  return rows;
}

if (typeof module !== "undefined" && module.exports) module.exports = { parsePastedTable };
