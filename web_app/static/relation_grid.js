/* 与 web_app/relation_grid.py 的 parse_pasted_table 保持同一规则。 */
function parsePastedTable(text) {
  if (text == null) return [];
  const raw = String(text);
  if (raw === "") return [];
  const normalized = raw.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalized.split("\n");
  if (normalized.endsWith("\n")) lines.pop();
  return lines.map((line) => line.split("\t"));
}
