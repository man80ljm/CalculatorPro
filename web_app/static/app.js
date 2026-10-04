const COURSE_KEY = "calculatorpro.currentCourseId";

const OPEN_FIELDS = [
  ["year_start", "学年起"],
  ["year_end", "学年止"],
  ["semester", "学期"],
  ["course_name", "课程名称"],
  ["department", "开课部门"],
  ["teacher", "授课教师"],
];

const BASIC_FIELDS = [
  ["course_name", "课程名称"],
  ["credits", "学分"],
  ["hours", "学时"],
  ["course_type", "课程性质"],
  ["course_code", "课程代码"],
  ["school_year_term", "学年学期"],
  ["college", "开课学院"],
  ["teacher", "任课教师"],
  ["major", "上课专业"],
  ["class_name", "上课班级"],
  ["student_count", "上课人数"],
  ["exam_count", "考核人数"],
];

const SPREAD_OPTIONS = ["大跨度（14-23分）", "中跨度（7-13分）", "小跨度（2-6分）"];
const DIST_OPTIONS = ["标准正态", "高分倾向", "低分倾向", "两极分化", "档位打分", "完全随机"];
const STYLE_OPTIONS = ["专业", "口语", "简洁", "详细", "幽默"];

function defaultState() {
  return {
    mode: "forward",
    studentCount: 40,
    courseOpen: { year_start: "", year_end: "", semester: "1", course_name: "", department: "", teacher: "" },
    courseBasic: {
      course_name: "", credits: "", hours: "", course_type: "", course_code: "",
      school_year_term: "", college: "", teacher: "", major: "", class_name: "",
      student_count: "", exam_count: "",
    },
    objectivesCount: 2,
    grid: [],
    gradReq: [{ requirement: "", indicator: "" }, { requirement: "", indicator: "" }],
    courseDescription: "",
    objectiveRequirements: ["", ""],
    spreadMode: "中跨度（7-13分）",
    distribution: "标准正态",
    noiseEnabled: false,
    noiseRatio: 0.1,
    severityMode: "random",
    noiseAllowed: [],
    reportStyle: "专业",
    wordLimit: 200,
  };
}

let state = defaultState();
let courseId = null;
let courseFiles = [];
let focusCell = { row: 0, col: 0 };
const selectedRows = new Set();

function clampCount(value) {
  const number = Number(value) || 2;
  return Math.max(1, Math.min(8, Math.round(number)));
}

function el(tag, attrs, children) {
  const node = document.createElement(tag);
  Object.entries(attrs || {}).forEach(([key, value]) => {
    if (key === "class") node.className = value;
    else if (key === "text") node.textContent = value;
    else node.setAttribute(key, value);
  });
  (children || []).forEach((child) => node.append(child));
  return node;
}

function fillSelect(select, options, current) {
  select.innerHTML = "";
  options.forEach((option) => {
    const node = el("option", { value: option, text: option });
    if (option === current) node.selected = true;
    select.append(node);
  });
}

function gridHeaders() {
  const headers = ["考核环节", "占比", "考核方式"];
  for (let i = 0; i < state.objectivesCount; i += 1) headers.push(`课程目标${i + 1}`);
  headers.push("小计");
  return headers;
}

function blankRow() {
  return gridHeaders().map(() => "");
}

function defaultRows() {
  const count = state.objectivesCount;
  function build(link, ratio, method, weights) {
    const row = blankRow();
    row[0] = link;
    row[1] = ratio;
    row[2] = method;
    for (let i = 0; i < count; i += 1) row[3 + i] = weights[i] || "0%";
    return row;
  }
  return [
    build("平时考核", "0.3", "平时作业", ["50%", "50%"]),
    build("期中考核", "0", "期中测验", ["50%", "50%"]),
    build("期末考核", "0.7", "期末考试", ["40%", "60%"]),
  ];
}

function parsePortion(text) {
  let raw = String(text ?? "").trim().replace(/％/g, "%");
  if (!raw) return 0;
  const percent = raw.endsWith("%");
  if (percent) raw = raw.slice(0, -1).trim();
  const value = Number(raw);
  if (!Number.isFinite(value)) return NaN;
  if (percent || value > 1) return value / 100;
  return value;
}

function formatPercent(fraction) {
  if (!Number.isFinite(fraction)) return "";
  const text = (fraction * 100).toFixed(2).replace(/\.?0+$/, "");
  return `${text}%`;
}

function subtotalText(row) {
  let sum = 0;
  for (let i = 0; i < state.objectivesCount; i += 1) {
    const value = parsePortion(row[3 + i]);
    if (Number.isFinite(value)) sum += value;
  }
  return formatPercent(sum);
}

function iterGridRows() {
  const groups = [];
  let current = null;
  state.grid.forEach((row) => {
    const linkName = String(row[0] || "").trim();
    const ratioText = String(row[1] || "").trim();
    const methodName = String(row[2] || "").trim();
    const weights = [];
    for (let i = 0; i < state.objectivesCount; i += 1) weights.push(String(row[3 + i] || "").trim());
    if (!linkName && !ratioText && !methodName && weights.every((item) => !item)) return;
    if (!current || (linkName && linkName !== current.name)) {
      current = { name: linkName, ratioText, methods: [] };
      groups.push(current);
    } else if (ratioText && !current.ratioText) {
      current.ratioText = ratioText;
    }
    if (methodName || weights.some(Boolean)) current.methods.push({ name: methodName, weights });
  });
  return groups;
}

function buildRelationPayload() {
  const links = [];
  iterGridRows().forEach((group) => {
    const ratio = parsePortion(group.ratioText);
    if (!Number.isFinite(ratio) || ratio <= 0) return;
    const methods = group.methods.map((method) => {
      const supports = {};
      let subtotal = 0;
      method.weights.forEach((text, index) => {
        const parsed = parsePortion(text);
        const value = Number.isFinite(parsed) ? parsed : 0;
        supports[`课程目标${index + 1}`] = Math.round(value * 1e6) / 1e6;
        subtotal += value;
      });
      return {
        name: method.name.trim(),
        supports,
        subtotal: Math.round(subtotal * 1e6) / 1e6,
      };
    });
    links.push({ name: group.name, ratio, methods });
  });
  const totals = {};
  let totalSum = 0;
  for (let i = 0; i < state.objectivesCount; i += 1) {
    const key = `课程目标${i + 1}`;
    let total = 0;
    links.forEach((link) => {
      link.methods.forEach((method) => {
        total += (method.supports[key] || 0) * link.ratio;
      });
    });
    totals[key] = Math.round(total * 1e6) / 1e6;
    totalSum += totals[key];
  }
  return {
    objectives_count: state.objectivesCount,
    links,
    objectives_total_weights: totals,
    total_sum: Math.round(totalSum * 1e6) / 1e6,
  };
}

function gridForSave() {
  const header = gridHeaders();
  const rows = state.grid.map((row) => header.map((_, index) => {
    if (index === header.length - 1) return subtotalText(row);
    return String(row[index] ?? "");
  }));
  return [header, ...rows];
}

function loadGrid(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    state.grid = defaultRows();
    return;
  }
  const copied = rows.map((row) => (Array.isArray(row) ? row.map((cell) => String(cell ?? "")) : []));
  let body = copied;
  if (String(copied[0][0] || "").trim() === "考核环节") {
    const objCount = copied[0].filter((name) => String(name).trim().startsWith("课程目标")).length;
    if (objCount) state.objectivesCount = clampCount(objCount);
    body = copied.slice(1);
  }
  state.grid = body.map((row) => {
    const next = blankRow();
    const limit = next.length - 1;
    for (let i = 0; i < limit && i < row.length; i += 1) next[i] = row[i];
    return next;
  });
  if (!state.grid.length) state.grid = defaultRows();
}

function gridFromPayload(payload) {
  const count = clampCount(payload.objectives_count || state.objectivesCount || 2);
  state.objectivesCount = count;
  const rows = [];
  (payload.links || []).forEach((link) => {
    const methods = link.methods || [];
    if (!methods.length) return;
    methods.forEach((method, index) => {
      const row = blankRow();
      if (index === 0) {
        row[0] = link.name || "";
        row[1] = String(link.ratio ?? "");
      }
      row[2] = method.name || "";
      for (let i = 0; i < count; i += 1) {
        const value = (method.supports || {})[`课程目标${i + 1}`];
        row[3 + i] = value == null || value === "" ? "" : formatPercent(Number(value));
      }
      rows.push(row);
    });
  });
  state.grid = rows.length ? rows : defaultRows();
}

function refitRows(previousCount) {
  const oldRows = state.grid.map((row) => row.slice());
  state.grid = oldRows.map((row) => {
    const next = blankRow();
    next[0] = row[0] || "";
    next[1] = row[1] || "";
    next[2] = row[2] || "";
    const copy = Math.min(previousCount, state.objectivesCount);
    for (let i = 0; i < copy; i += 1) next[3 + i] = row[3 + i] || "";
    return next;
  });
  if (!state.grid.length) state.grid = defaultRows();
}

function applyPaste(startRow, startCol, text) {
  const block = parsePastedTable(text);
  if (!block.length) return;
  while (state.grid.length < startRow + block.length) state.grid.push(blankRow());
  const editableUntil = gridHeaders().length - 1;
  block.forEach((pasted, rOffset) => {
    pasted.forEach((value, cOffset) => {
      const col = startCol + cOffset;
      if (col < 0 || col >= editableUntil) return;
      state.grid[startRow + rOffset][col] = String(value).trim();
    });
  });
  focusCell = { row: startRow, col: startCol };
  renderGrid();
  updateRatioSum();
  renderNoiseTargets();
  updateSummaries();
}

function renderScalarFields(container, spec, bag) {
  container.innerHTML = "";
  spec.forEach(([key, label]) => {
    const input = el("input", { type: "text", value: bag[key] || "" });
    input.addEventListener("input", () => {
      bag[key] = input.value;
    });
    if (bag === state.courseBasic) input.dataset.basic = key;
    container.append(el("label", { class: "field", text: label }, [input]));
  });
}

function renderGrid() {
  const table = document.getElementById("relationGrid");
  const headers = gridHeaders();
  table.innerHTML = "";
  const head = document.createElement("thead");
  const headRow = document.createElement("tr");
  headRow.append(el("th", { text: "" }));
  headers.forEach((name) => headRow.append(el("th", { text: name })));
  head.append(headRow);
  table.append(head);
  const body = document.createElement("tbody");
  state.grid.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    if (selectedRows.has(rowIndex)) tr.classList.add("row-selected");
    const pick = el("input", { type: "checkbox" });
    pick.checked = selectedRows.has(rowIndex);
    pick.addEventListener("change", () => {
      if (pick.checked) selectedRows.add(rowIndex);
      else selectedRows.delete(rowIndex);
      tr.classList.toggle("row-selected", pick.checked);
    });
    tr.append(el("td", { class: "pick" }, [pick]));
    headers.forEach((header, col) => {
      if (header === "小计") {
        tr.append(el("td", { class: "readonly", "data-subtotal": "1", text: subtotalText(row) }));
        return;
      }
      // textarea 而不是 input：单元格可以保留 Excel 里 Alt+Enter 的换行
      const input = el("textarea", { rows: "1", class: "cell" });
      input.value = row[col] || "";
      const fit = () => {
        input.rows = Math.max(1, String(input.value).split("\n").length);
      };
      fit();
      input.addEventListener("focus", () => {
        focusCell = { row: rowIndex, col };
      });
      input.addEventListener("input", () => {
        state.grid[rowIndex][col] = input.value;
        fit();
        const total = tr.querySelector("[data-subtotal]");
        if (total) total.textContent = subtotalText(state.grid[rowIndex]);
        updateRatioSum();
        if (col <= 2) renderNoiseTargets();
      });
      input.addEventListener("paste", (event) => {
        const text = event.clipboardData ? event.clipboardData.getData("text/plain") : "";
        if (text.includes("\t") || text.includes("\n") || text.includes("\r")) {
          event.preventDefault();
          applyPaste(rowIndex, col, text);
        }
      });
      const cell = el("td", {}, [input]);
      if (focusCell.row === rowIndex && focusCell.col === col) cell.classList.add("focused");
      tr.append(cell);
    });
    body.append(tr);
  });
  table.append(body);
  const focused = table.querySelector("td.focused textarea");
  if (focused) focused.focus();
}

function updateRatioSum() {
  let sum = 0;
  iterGridRows().forEach((group) => {
    const ratio = parsePortion(group.ratioText);
    if (Number.isFinite(ratio)) sum += ratio;
  });
  const note = document.getElementById("ratioSum");
  const ok = Math.abs(sum - 1) < 0.001;
  note.textContent = `当前占比合计 ${sum.toFixed(2)}${ok ? "" : "，需要等于 1.00"}`;
  note.classList.toggle("bad", !ok);
}

function renderGrad() {
  while (state.gradReq.length < state.objectivesCount) state.gradReq.push({ requirement: "", indicator: "" });
  while (state.objectiveRequirements.length < state.objectivesCount) state.objectiveRequirements.push("");
  const root = document.getElementById("gradReq");
  root.innerHTML = "";
  for (let i = 0; i < state.objectivesCount; i += 1) {
    const requirement = el("input", { type: "text", value: state.gradReq[i].requirement || "" });
    const indicator = el("input", { type: "text", value: state.gradReq[i].indicator || "" });
    requirement.addEventListener("input", () => { state.gradReq[i].requirement = requirement.value; });
    indicator.addEventListener("input", () => { state.gradReq[i].indicator = indicator.value; });
    root.append(el("div", { class: "fields" }, [
      el("p", { class: "span-2", text: `课程目标${i + 1}` }),
      el("label", { class: "field", text: "支撑的毕业要求" }, [requirement]),
      el("label", { class: "field", text: "支撑的毕业要求指标点" }, [indicator]),
    ]));
  }
  const reqRoot = document.getElementById("requirements");
  reqRoot.innerHTML = "";
  for (let i = 0; i < state.objectivesCount; i += 1) {
    const area = el("textarea");
    area.value = state.objectiveRequirements[i] || "";
    area.addEventListener("input", () => { state.objectiveRequirements[i] = area.value; });
    reqRoot.append(el("label", { text: `课程目标${i + 1}要求` }, [area]));
  }
}

function renderNoiseTargets() {
  const root = document.getElementById("noiseTargets");
  root.innerHTML = "";
  const names = [];
  iterGridRows().forEach((group) => {
    const ratio = parsePortion(group.ratioText);
    if (!Number.isFinite(ratio) || ratio <= 0) return;
    group.methods.forEach((method) => {
      const name = method.name.trim();
      if (name && !names.includes(name)) names.push(name);
    });
  });
  if (!state.noiseAllowed.length) state.noiseAllowed = names.slice();
  names.forEach((name) => {
    const box = el("input", { type: "checkbox" });
    box.checked = state.noiseAllowed.includes(name);
    box.addEventListener("change", () => {
      if (box.checked && !state.noiseAllowed.includes(name)) state.noiseAllowed.push(name);
      else state.noiseAllowed = state.noiseAllowed.filter((item) => item !== name);
    });
    const label = el("label");
    label.append(box, document.createTextNode(name));
    root.append(label);
  });
}

function collectSettings() {
  const basic = { ...state.courseBasic };
  if (!basic.course_name) basic.course_name = state.courseOpen.course_name || "";
  const openInfo = { ...state.courseOpen, term: state.courseOpen.semester || "1" };
  const payload = buildRelationPayload();
  const ratios = { usual: 0, midterm: 0, final: 0 };
  iterGridRows().forEach((group) => {
    const ratio = parsePortion(group.ratioText);
    if (!Number.isFinite(ratio)) return;
    if (group.name.includes("平时")) ratios.usual += ratio;
    else if (group.name.includes("期中")) ratios.midterm += ratio;
    else if (group.name.includes("期末")) ratios.final += ratio;
  });
  return {
    mode: state.mode,
    student_count: Number(state.studentCount) || 1,
    course_open_info: openInfo,
    course_basic_info: basic,
    ratios,
    relation_grid: gridForSave(),
    relation_payload: payload,
    grad_req_map: state.gradReq.slice(0, state.objectivesCount).map((row, index) => ({
      objective: `课程目标${index + 1}`,
      requirement: row.requirement || "",
      indicator: row.indicator || "",
    })),
    course_description: state.courseDescription,
    objective_requirements: state.objectiveRequirements.slice(0, state.objectivesCount),
    spread_mode: state.spreadMode,
    distribution: state.distribution,
    noise_config: state.mode === "reverse" && state.noiseEnabled ? {
      noise_ratio: Number(state.noiseRatio) || 0,
      severity_mode: state.severityMode,
      allowed_items: state.noiseAllowed.slice(),
    } : null,
    report_style: state.reportStyle,
    word_limit: Number(state.wordLimit) || 200,
  };
}

function fillBag(bag, incoming) {
  const source = incoming && typeof incoming === "object" ? incoming : {};
  Object.keys(bag).forEach((key) => {
    bag[key] = source[key] == null ? (key === "semester" ? "1" : "") : source[key];
  });
}

function applyServerSettings(settings) {
  const data = settings || {};
  state.mode = data.mode === "reverse" ? "reverse" : "forward";
  state.studentCount = Number(data.student_count) || state.studentCount || 40;
  fillBag(state.courseOpen, data.course_open_info);
  fillBag(state.courseBasic, data.course_basic_info);
  state.courseDescription = data.course_description || "";
  state.spreadMode = data.spread_mode || state.spreadMode;
  state.distribution = data.distribution || state.distribution;
  state.reportStyle = data.report_style || state.reportStyle;
  state.wordLimit = Number(data.word_limit) || 200;
  if (Array.isArray(data.objective_requirements) && data.objective_requirements.length) {
    state.objectiveRequirements = data.objective_requirements.map((item) => String(item ?? ""));
  }
  if (Array.isArray(data.grad_req_map) && data.grad_req_map.length) {
    state.gradReq = data.grad_req_map.map((row) => ({
      requirement: (row && row.requirement) || "",
      indicator: (row && row.indicator) || "",
    }));
  }
  if (data.noise_config && typeof data.noise_config === "object") {
    state.noiseEnabled = true;
    state.noiseRatio = Number(data.noise_config.noise_ratio);
    if (!Number.isFinite(state.noiseRatio)) state.noiseRatio = 0.1;
    state.severityMode = data.noise_config.severity_mode || "random";
    state.noiseAllowed = Array.isArray(data.noise_config.allowed_items) ? data.noise_config.allowed_items.slice() : [];
  } else {
    state.noiseEnabled = false;
    state.noiseAllowed = [];
  }
  if (Array.isArray(data.relation_grid) && data.relation_grid.length) loadGrid(data.relation_grid);
  else if (data.relation_payload && data.relation_payload.links) gridFromPayload(data.relation_payload);
  else {
    state.objectivesCount = clampCount(state.objectivesCount || 2);
    state.grid = defaultRows();
  }
  selectedRows.clear();
}

function validateBeforeRun(needsFile) {
  if (!courseId) return "请先新建并选择课程文件夹";
  const groups = iterGridRows();
  let ratioSum = 0;
  let positive = 0;
  for (const group of groups) {
    const ratio = parsePortion(group.ratioText);
    if (!group.ratioText) return `${group.name || "考核环节"} 缺少占比`;
    if (!Number.isFinite(ratio)) return `${group.name} 的占比不是有效数字`;
    ratioSum += ratio;
    if (ratio <= 0) continue;
    positive += 1;
    if (!group.methods.length) return `${group.name} 请至少填写一种考核方式`;
    let weightSum = 0;
    for (const method of group.methods) {
      if (!method.name.trim()) return `${group.name} 有未命名的考核方式`;
      for (const text of method.weights) {
        if (!text) continue;
        const value = parsePortion(text);
        if (!Number.isFinite(value)) return `${method.name} 的目标权重不是有效数字`;
        weightSum += value;
      }
    }
    if (Math.abs(weightSum - 1) > 0.02) {
      return `${group.name} 的目标权重合计应为 100，当前约为 ${Math.round(weightSum * 100)}`;
    }
  }
  if (!positive) return "请至少保留一个占比大于 0 的考核环节";
  if (Math.abs(ratioSum - 1) > 0.001) return "各考核环节占比之和必须等于 1";
  if (needsFile && !courseFiles.some((file) => file.kind === "grade")) return "请先导入成绩 Excel";
  return "";
}

function setBusy(busy) {
  ["templateBtn", "calcBtn", "exportBtn", "aiBtn", "saveCourseBtn"].forEach((id) => {
    document.getElementById(id).disabled = busy;
  });
}

let statusTimer = null;

/* 状态提示（保存成功、出错、进度）。计算结果在“计算结果”面板里单独显示。 */
function showResult(text, isError) {
  const node = document.getElementById("result");
  node.textContent = text;
  node.classList.toggle("error", Boolean(isError));
  clearTimeout(statusTimer);
  if (text && !isError && !/…$/.test(text)) {
    statusTimer = setTimeout(() => { node.textContent = ""; }, 6000);
  }
}

async function errorMessage(response) {
  try {
    const data = await response.json();
    if (typeof data.detail === "string") return data.detail;
  } catch (err) {
    /* 下载失败时响应体不是 JSON */
  }
  return `请求失败（${response.status}）`;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function filenameFromDisposition(header, fallback) {
  if (!header) return fallback;
  const star = header.match(/filename\*=UTF-8''([^;]+)/i);
  if (star) return decodeURIComponent(star[1]);
  const plain = header.match(/filename="?([^"]+)"?/i);
  return plain ? plain[1] : fallback;
}

async function api(path, options) {
  const response = await fetch(path, options);
  if (response.status === 401) {
    location.href = "/login";
    throw new Error("未登录");
  }
  return response;
}

function syncMode() {
  const reverse = state.mode === "reverse";
  document.getElementById("modeForward").setAttribute("aria-pressed", String(!reverse));
  document.getElementById("modeReverse").setAttribute("aria-pressed", String(reverse));
  document.getElementById("modeForward").classList.toggle("secondary", reverse);
  document.getElementById("modeReverse").classList.toggle("secondary", !reverse);
  ["spreadMode", "distribution", "noiseEnabled", "noiseRatio", "severityMode"].forEach((id) => {
    document.getElementById(id).disabled = !reverse;
  });
  document.getElementById("noiseBox").style.opacity = reverse ? "1" : "0.45";
}

function syncControls() {
  document.getElementById("objCount").value = state.objectivesCount;
  document.getElementById("description").value = state.courseDescription;
  document.getElementById("studentCount").value = state.studentCount;
  fillSelect(document.getElementById("spreadMode"), SPREAD_OPTIONS, state.spreadMode);
  fillSelect(document.getElementById("distribution"), DIST_OPTIONS, state.distribution);
  fillSelect(document.getElementById("reportStyle"), STYLE_OPTIONS, state.reportStyle);
  document.getElementById("wordLimit").value = state.wordLimit;
  document.getElementById("noiseEnabled").checked = state.noiseEnabled;
  document.getElementById("noiseRatio").value = state.noiseRatio;
  document.getElementById("severityMode").value = state.severityMode;
  syncMode();
}

function renderAll() {
  renderScalarFields(document.getElementById("openFields"), OPEN_FIELDS, state.courseOpen);
  renderScalarFields(document.getElementById("basicFields"), BASIC_FIELDS, state.courseBasic);
  syncControls();
  renderGrid();
  renderGrad();
  updateRatioSum();
  renderNoiseTargets();
}

function renderFiles() {
  const root = document.getElementById("savedFiles");
  root.innerHTML = "";
  const grade = courseFiles.find((file) => file.kind === "grade");
  const previous = courseFiles.find((file) => file.kind === "previous");
  document.getElementById("fileName").textContent = grade ? `已保存：${grade.original_name}` : "尚未导入成绩表";
  document.getElementById("prevName").textContent = previous ? `已保存：${previous.original_name}` : "不导入则按 0 对比";
  const more = document.getElementById("filesMore");
  if (!courseFiles.length) {
    root.append(el("li", { class: "hint", text: "还没有文件" }));
    more.hidden = true;
    updateSummaries();
    return;
  }
  const visible = showAllFiles ? courseFiles : courseFiles.slice(0, FILES_PREVIEW);
  visible.forEach((file) => {
    const button = el("button", { type: "button", class: "secondary", text: "下载" });
    button.addEventListener("click", () => downloadSaved(file));
    const label = el("span", { text: `${file.original_name}（${KIND_LABELS[file.kind] || file.kind}）` });
    root.append(el("li", {}, [label, button]));
  });
  more.hidden = courseFiles.length <= FILES_PREVIEW;
  more.textContent = showAllFiles ? "收起" : `显示全部（${courseFiles.length}）`;
  updateSummaries();
}

async function downloadSaved(file) {
  try {
    const response = await api(`/api/courses/${courseId}/files/${file.id}`);
    if (!response.ok) return showResult(await errorMessage(response), true);
    const blob = await response.blob();
    const name = filenameFromDisposition(response.headers.get("Content-Disposition"), file.original_name || "download");
    downloadBlob(blob, name);
  } catch (err) {
    if (err.message !== "未登录") showResult("下载失败", true);
  }
}

async function refreshFiles() {
  if (!courseId) {
    courseFiles = [];
    renderFiles();
    return;
  }
  const response = await api(`/api/courses/${courseId}/files`);
  if (!response.ok) return;
  const data = await response.json();
  courseFiles = data.files || [];
  renderFiles();
}

async function saveCourse() {
  if (!courseId) throw new Error("请先新建课程文件夹");
  const name = document.getElementById("courseNameInput").value.trim();
  const response = await api(`/api/courses/${courseId}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, settings: collectSettings() }),
  });
  if (!response.ok) {
    const message = await errorMessage(response);
    showResult(message, true);
    throw new Error(message);
  }
  const data = await response.json();
  const option = document.querySelector(`#courseSelect option[value="${courseId}"]`);
  if (option) option.textContent = data.name;
  showResult("已保存到课程文件夹。", false);
  return data;
}

async function openCourse(id) {
  const response = await api(`/api/courses/${id}`);
  if (!response.ok) {
    showResult(await errorMessage(response), true);
    return;
  }
  const data = await response.json();
  courseId = data.id;
  localStorage.setItem(COURSE_KEY, String(courseId));
  document.getElementById("courseSelect").value = String(courseId);
  document.getElementById("courseNameInput").value = data.name || "";
  state = defaultState();
  applyServerSettings(data.settings || {});
  courseFiles = data.files || [];
  showAllFiles = false;
  renderAll();
  renderFiles();
  afterCourseLoaded();
}

async function loadCourses(preferId) {
  const response = await api("/api/courses");
  const data = await response.json();
  const courses = data.courses || [];
  const select = document.getElementById("courseSelect");
  select.innerHTML = "";
  if (!courses.length) {
    courseId = null;
    courseFiles = [];
    select.append(el("option", { value: "", text: "尚未创建课程" }));
    renderFiles();
    afterCourseLoaded();
    showResult("请先新建课程文件夹。", false);
    return;
  }
  courses.forEach((course) => {
    select.append(el("option", { value: String(course.id), text: course.name }));
  });
  const stored = Number(localStorage.getItem(COURSE_KEY));
  const wanted = preferId || stored;
  const chosen = courses.some((course) => course.id === wanted) ? wanted : courses[0].id;
  await openCourse(chosen);
}

async function uploadKind(input, kind) {
  const file = input.files && input.files[0];
  if (!file) return;
  if (!courseId) {
    showResult("请先新建并选择课程文件夹", true);
    input.value = "";
    return;
  }
  try {
    await saveCourse();
    const body = new FormData();
    body.append("kind", kind);
    body.append("file", file);
    const response = await api(`/api/courses/${courseId}/files`, { method: "POST", body });
    if (!response.ok) return showResult(await errorMessage(response), true);
    await refreshFiles();
    showResult(kind === "grade" ? `已保存成绩表：${file.name}` : `已保存上一学年达成度表：${file.name}`, false);
  } catch (err) {
    if (err.message !== "未登录" && err.message !== "请先新建课程文件夹") showResult(err.message || "上传失败", true);
  }
}

async function postCourse(path) {
  await saveCourse();
  const body = new FormData();
  body.append("settings", JSON.stringify(collectSettings()));
  return api(`/api/courses/${courseId}${path}`, { method: "POST", body });
}

function bindOnce() {
  document.getElementById("objCount").addEventListener("change", () => {
    const previous = state.objectivesCount;
    state.objectivesCount = clampCount(document.getElementById("objCount").value);
    document.getElementById("objCount").value = state.objectivesCount;
    refitRows(previous);
    renderGrid();
    renderGrad();
    updateRatioSum();
    renderNoiseTargets();
  });
  document.getElementById("description").addEventListener("input", (event) => {
    state.courseDescription = event.target.value;
  });
  document.getElementById("studentCount").addEventListener("input", (event) => {
    state.studentCount = Number(event.target.value);
  });
  document.getElementById("spreadMode").addEventListener("change", (event) => { state.spreadMode = event.target.value; });
  document.getElementById("distribution").addEventListener("change", (event) => { state.distribution = event.target.value; });
  document.getElementById("reportStyle").addEventListener("change", (event) => { state.reportStyle = event.target.value; });
  document.getElementById("wordLimit").addEventListener("input", (event) => { state.wordLimit = Number(event.target.value); });
  document.getElementById("noiseEnabled").addEventListener("change", (event) => { state.noiseEnabled = event.target.checked; });
  document.getElementById("noiseRatio").addEventListener("input", (event) => { state.noiseRatio = Number(event.target.value); });
  document.getElementById("severityMode").addEventListener("change", (event) => { state.severityMode = event.target.value; });
  document.getElementById("modeForward").addEventListener("click", () => { state.mode = "forward"; syncMode(); });
  document.getElementById("modeReverse").addEventListener("click", () => { state.mode = "reverse"; syncMode(); });
  document.getElementById("addRowBtn").addEventListener("click", () => {
    state.grid.push(blankRow());
    renderGrid();
  });
  document.getElementById("removeRowsBtn").addEventListener("click", () => {
    let indexes = [...selectedRows];
    if (!indexes.length) indexes = [focusCell.row];
    const drop = new Set(indexes);
    state.grid = state.grid.filter((_, index) => !drop.has(index));
    if (!state.grid.length) state.grid.push(blankRow());
    selectedRows.clear();
    focusCell = { row: 0, col: 0 };
    renderGrid();
    updateRatioSum();
    renderNoiseTargets();
  });
  document.getElementById("gradeFile").addEventListener("change", () => {
    uploadKind(document.getElementById("gradeFile"), "grade");
  });
  document.getElementById("prevFile").addEventListener("change", () => {
    uploadKind(document.getElementById("prevFile"), "previous");
  });
  document.getElementById("courseSelect").addEventListener("change", async (event) => {
    const next = Number(event.target.value);
    if (!next || next === courseId) return;
    const previous = courseId;
    try {
      if (previous) await saveCourse();
      await openCourse(next);
    } catch (err) {
      if (previous) event.target.value = String(previous);
    }
  });
  document.getElementById("newCourseBtn").addEventListener("click", async () => {
    const name = document.getElementById("courseNameInput").value.trim() || "未命名课程";
    try {
      if (courseId) await saveCourse();
      const response = await api("/api/courses", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name }),
      });
      if (!response.ok) return showResult(await errorMessage(response), true);
      const data = await response.json();
      await loadCourses(data.id);
      showResult(`已新建课程文件夹：${data.name}`, false);
    } catch (err) {
      if (err.message !== "未登录") showResult(err.message || "新建失败", true);
    }
  });
  document.getElementById("saveCourseBtn").addEventListener("click", async () => {
    try {
      await saveCourse();
    } catch (err) {
      /* saveCourse 已显示原因 */
    }
  });
  document.getElementById("deleteCourseBtn").addEventListener("click", async () => {
    if (!courseId) return;
    const name = document.getElementById("courseNameInput").value.trim() || "当前课程";
    if (!window.confirm(`删除课程文件夹「${name}」及其文件？`)) return;
    const response = await api(`/api/courses/${courseId}`, { method: "DELETE" });
    if (!response.ok) return showResult(await errorMessage(response), true);
    courseId = null;
    localStorage.removeItem(COURSE_KEY);
    state = defaultState();
    state.grid = defaultRows();
    renderAll();
    await loadCourses();
    showResult("课程文件夹已删除。", false);
  });
  document.getElementById("passwordBtn").addEventListener("click", () => {
    const panel = document.getElementById("passwordPanel");
    panel.hidden = !panel.hidden;
  });
  document.getElementById("changePasswordBtn").addEventListener("click", async () => {
    const message = document.getElementById("passwordMsg");
    message.textContent = "";
    try {
      const response = await api("/api/password", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          current_password: document.getElementById("currentPassword").value,
          new_password: document.getElementById("newPassword").value,
        }),
      });
      if (!response.ok) {
        message.textContent = await errorMessage(response);
        return;
      }
      document.getElementById("currentPassword").value = "";
      document.getElementById("newPassword").value = "";
      message.textContent = "密码已更新。";
    } catch (err) {
      if (err.message !== "未登录") message.textContent = "修改失败";
    }
  });
  document.getElementById("logoutBtn").addEventListener("click", async () => {
    await fetch("/api/logout", { method: "POST" });
    clearStoredResults();
    location.href = "/login";
  });
}

async function loadAiStatus() {
  const banner = document.getElementById("aiBanner");
  try {
    const response = await api("/api/ai-status");
    const data = await response.json();
    banner.textContent = data.message || "";
    banner.classList.toggle("ok", Boolean(data.enabled));
  } catch (err) {
    banner.textContent = "无法确认 AI 报告状态";
  }
}

async function loadMe() {
  try {
    const response = await api("/api/me");
    if (!response.ok) return;
    const data = await response.json();
    document.getElementById("who").textContent = data.username || "";
  } catch (err) {
    /* 未登录时 api() 会跳转 */
  }
}

document.getElementById("templateBtn").addEventListener("click", async () => {
  const problem = validateBeforeRun(false);
  if (problem) return showResult(problem, true);
  setBusy(true);
  showResult("正在生成模板…", false);
  try {
    const response = await postCourse("/template");
    if (!response.ok) return showResult(await errorMessage(response), true);
    const blob = await response.blob();
    const name = filenameFromDisposition(response.headers.get("Content-Disposition"), "成绩模板.xlsx");
    downloadBlob(blob, name);
    await refreshFiles();
    showResult(`模板已下载：${name}`, false);
  } catch (err) {
    if (err.message !== "未登录") showResult(err.message || "模板下载失败", true);
  } finally {
    setBusy(false);
  }
});

document.getElementById("calcBtn").addEventListener("click", async () => {
  const problem = validateBeforeRun(true);
  if (problem) return showResult(problem, true);
  setBusy(true);
  showResult("正在计算…", false);
  try {
    const response = await postCourse("/calculate");
    if (!response.ok) return showResult(await errorMessage(response), true);
    const data = await response.json();
    await refreshFiles();
    storeResult(data);
    renderResultView();
    setPanelOpen("result", true);
    savePanelState();
    document.querySelector('details[data-panel="result"]').scrollIntoView({ behavior: "smooth", block: "start" });
    showResult("计算完成，结果见“计算结果”。", false);
  } catch (err) {
    if (err.message !== "未登录") showResult(err.message || "计算失败", true);
  } finally {
    setBusy(false);
  }
});

async function downloadZip(path, fallbackName, pendingText) {
  const problem = validateBeforeRun(true);
  if (problem) return showResult(problem, true);
  setBusy(true);
  showResult(pendingText, false);
  try {
    const response = await postCourse(path);
    if (!response.ok) return showResult(await errorMessage(response), true);
    const blob = await response.blob();
    const name = filenameFromDisposition(response.headers.get("Content-Disposition"), fallbackName);
    downloadBlob(blob, name);
    await refreshFiles();
    showResult(`已下载：${name}`, false);
  } catch (err) {
    if (err.message !== "未登录") showResult(err.message || "下载失败", true);
  } finally {
    setBusy(false);
  }
}

document.getElementById("exportBtn").addEventListener("click", () => {
  downloadZip("/export", "统计表.zip", "正在导出统计表…");
});

document.getElementById("aiBtn").addEventListener("click", () => {
  downloadZip("/ai-report", "AI分析报告.zip", "正在生成 AI 分析报告…");
});

/* ---------- 折叠面板：每门课程记住展开状态，收起时显示一行摘要 ---------- */
const PANEL_PREFIX = "calculatorpro.panels.";
const RESULT_PREFIX = "calculatorpro.result.";
const FILES_PREVIEW = 6;
const KIND_LABELS = { grade: "成绩表", previous: "上一学年", template: "模板", output: "导出", report: "AI 报告" };
let showAllFiles = false;
let summaryTimer = null;

function panelNodes() {
  return Array.from(document.querySelectorAll("details.panel"));
}

function setPanelOpen(key, open) {
  const node = document.querySelector(`details.panel[data-panel="${key}"]`);
  if (node) node.open = Boolean(open);
}

function readJson(key) {
  try {
    return JSON.parse(localStorage.getItem(key) || "null");
  } catch (err) {
    return null;
  }
}

function savePanelState() {
  if (!courseId) return;
  const map = {};
  panelNodes().forEach((node) => { map[node.dataset.panel] = node.open; });
  localStorage.setItem(PANEL_PREFIX + courseId, JSON.stringify(map));
}

function relationReady() {
  try {
    const payload = buildRelationPayload();
    const sum = payload.links.reduce((acc, link) => acc + link.ratio, 0);
    return payload.links.length > 0 && Math.abs(sum - 1) < 0.001;
  } catch (err) {
    return false;
  }
}

/* 默认只展开当前这一步需要的面板 */
function defaultOpenPanels() {
  const open = new Set();
  if (!courseId) return open;
  const named = String(state.courseBasic.course_name || state.courseOpen.course_name || "").trim();
  if (!named) open.add("basic");
  if (!relationReady()) {
    open.add("relation");
    return open;
  }
  open.add("run");
  if (loadStoredResult()) open.add("result");
  return open;
}

function applyPanelState() {
  const saved = courseId ? readJson(PANEL_PREFIX + courseId) : null;
  const defaults = defaultOpenPanels();
  panelNodes().forEach((node) => {
    const key = node.dataset.panel;
    node.open = saved && Object.prototype.hasOwnProperty.call(saved, key) ? Boolean(saved[key]) : defaults.has(key);
  });
}

function storeResult(data) {
  if (!courseId) return;
  const keep = {
    mode: data.mode,
    course_name: data.course_name,
    student_count: data.student_count,
    average_score: data.average_score,
    achievement: data.achievement || {},
    files: data.files || [],
    at: new Date().toISOString(),
  };
  localStorage.setItem(RESULT_PREFIX + courseId, JSON.stringify(keep));
}

function loadStoredResult() {
  return courseId ? readJson(RESULT_PREFIX + courseId) : null;
}

function clearStoredResults() {
  Object.keys(localStorage)
    .filter((key) => key.startsWith(RESULT_PREFIX))
    .forEach((key) => localStorage.removeItem(key));
}

function totalAchievement(achievement) {
  const entries = Object.entries(achievement || {});
  const total = entries.find(([key]) => key.includes("总"));
  return { total: total ? total[1] : null, parts: entries.filter(([key]) => !key.includes("总")) };
}

function stat(label, value, strong) {
  return el("div", { class: strong ? "stat stat-main" : "stat" }, [
    el("span", { class: "stat-k", text: label }),
    el("span", { class: "stat-v", text: value == null || value === "" ? "—" : String(value) }),
  ]);
}

function renderResultView() {
  const root = document.getElementById("resultView");
  root.innerHTML = "";
  const data = loadStoredResult();
  if (!data) {
    root.append(el("p", { class: "hint", text: "尚未计算。导入成绩表后点“计算达成度”。" }));
    updateSummaries();
    return;
  }
  const { total, parts } = totalAchievement(data.achievement);
  const when = data.at ? new Date(data.at).toLocaleString("zh-CN", { hour12: false }) : "";
  root.append(el("div", { class: "stats" }, [
    stat("总达成度", total, true),
    stat("学生人数", data.student_count),
    stat("平均分", data.average_score),
    stat("模式", data.mode === "reverse" ? "逆向" : "正向"),
  ]));
  root.append(el("p", { class: "hint", text: `${data.course_name || ""} · 计算于 ${when}` }));

  const table = el("table", { class: "kv" });
  parts.forEach(([key, value]) => {
    table.append(el("tr", {}, [el("th", { text: key }), el("td", { text: String(value) })]));
  });
  root.append(el("details", { class: "sub", "data-sub": "objectives" }, [
    el("summary", { text: `各课程目标达成度（${parts.length}）` }),
    table,
  ]));

  const detailFile = courseFiles.find((file) => file.kind === "output" && /成绩明细\.xlsx$/.test(file.original_name || ""));
  const studentBox = el("div", { class: "sub-body" }, [
    el("p", { class: "hint", text: "每位学生的各考核方式得分与课程目标得分在“成绩明细.xlsx”里。" }),
  ]);
  if (detailFile) {
    const button = el("button", { type: "button", class: "secondary", text: `下载 ${detailFile.original_name}` });
    button.addEventListener("click", () => downloadSaved(detailFile));
    studentBox.append(button);
  }
  root.append(el("details", { class: "sub", "data-sub": "students" }, [
    el("summary", { text: `学生成绩明细（${data.student_count || 0} 人）` }),
    studentBox,
  ]));

  const list = el("ul", { class: "plain-list" });
  (data.files || []).forEach((name) => list.append(el("li", { text: name })));
  root.append(el("details", { class: "sub", "data-sub": "outputs" }, [
    el("summary", { text: `本次生成的文件（${(data.files || []).length}）` }),
    list,
  ]));
  updateSummaries();
}

function joinBits(bits) {
  return bits.filter((bit) => bit && String(bit).trim()).join(" · ");
}

function shortText(text, size) {
  const clean = String(text || "").replace(/\s+/g, " ").trim();
  return clean.length > size ? `${clean.slice(0, size)}…` : clean;
}

function updateSummaries() {
  const set = (key, text) => {
    const node = document.querySelector(`[data-sum="${key}"]`);
    if (node) node.textContent = text || "未填写";
  };
  const o = state.courseOpen;
  const years = o.year_start || o.year_end ? `${o.year_start || "?"}-${o.year_end || "?"} 学年` : "";
  set("open", joinBits([years, o.semester ? `第${o.semester}学期` : "", o.department, o.teacher]));
  const b = state.courseBasic;
  set("basic", joinBits([
    b.course_name || o.course_name,
    b.credits ? `${b.credits} 学分` : "",
    b.hours ? `${b.hours} 学时` : "",
    b.student_count ? `${b.student_count} 人` : "",
  ]));
  const rows = state.grid.filter((row) => row.some((cell) => String(cell || "").trim())).length;
  let ratios = "";
  try {
    ratios = buildRelationPayload().links.map((link) => `${link.name} ${formatPercent(link.ratio)}`).join(" / ");
  } catch (err) {
    ratios = "";
  }
  set("relation", joinBits([`${state.objectivesCount} 个目标`, `${rows} 行`, ratios || "占比未填", relationReady() ? "" : "未完成"]));
  const filled = state.gradReq.slice(0, state.objectivesCount).filter((item) => (item.requirement || "").trim()).length;
  set("grad", `已填 ${filled}/${state.objectivesCount} 个目标`);
  set("intro", shortText(state.courseDescription, 28));
  const grade = courseFiles.find((file) => file.kind === "grade");
  const previous = courseFiles.find((file) => file.kind === "previous");
  set("run", joinBits([
    state.mode === "reverse" ? "逆向模式" : "正向模式",
    grade ? `成绩表：${grade.original_name}` : "未导入成绩表",
    previous ? "含上一学年" : "",
  ]));
  set("files", courseFiles.length ? `${courseFiles.length} 个文件` : "还没有文件");
  const result = loadStoredResult();
  if (result) {
    const { total } = totalAchievement(result.achievement);
    set("result", joinBits([`总达成度 ${total ?? "—"}`, `${result.student_count} 人`, `平均分 ${result.average_score}`]));
  } else {
    set("result", "尚未计算");
  }
}

function afterCourseLoaded() {
  renderResultView();
  applyPanelState();
  updateSummaries();
}

function bindPanels() {
  panelNodes().forEach((node) => {
    node.querySelector("summary").addEventListener("click", () => setTimeout(savePanelState, 0));
  });
  document.getElementById("expandAllBtn").addEventListener("click", () => {
    panelNodes().forEach((node) => { node.open = true; });
    savePanelState();
  });
  document.getElementById("collapseAllBtn").addEventListener("click", () => {
    panelNodes().forEach((node) => { node.open = false; });
    savePanelState();
  });
  document.getElementById("filesMore").addEventListener("click", () => {
    showAllFiles = !showAllFiles;
    renderFiles();
  });
  const refresh = () => {
    clearTimeout(summaryTimer);
    summaryTimer = setTimeout(updateSummaries, 150);
  };
  document.querySelector("main").addEventListener("input", refresh);
  document.querySelector("main").addEventListener("change", refresh);
  document.querySelector("main").addEventListener("click", (event) => {
    if (event.target.closest("#modeForward, #modeReverse, #addRowBtn, #removeRowsBtn")) refresh();
  });
}

state.grid = defaultRows();
bindOnce();
bindPanels();
renderAll();
afterCourseLoaded();
loadCourses();
loadAiStatus();
loadMe();
