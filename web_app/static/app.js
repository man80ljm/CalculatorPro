const COURSE_KEY = "calculatorpro.currentCourseId";

const OPEN_FIELDS = [
  ["course_name", "课程名称"],
  ["department", "开课部门"],
  ["year_start", "学年起"],
  ["year_end", "学年止"],
  ["semester", "学期"],
  ["teacher", "任课教师"],
  ["class_name", "上课班级"],
  ["major", "上课专业"],
];

const BASIC_FIELDS = [
  ["course_name", "课程名称"],
  ["credits", "学分"],
  ["hours", "学时"],
  ["course_type", "课程性质"],
  ["course_code", "课程代码"],
  ["college", "开课学院"],
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
    gradReq: [{ requirement: "", indicator: "", strength: "" }, { requirement: "", indicator: "", strength: "" }],
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
let terms = [];
let currentTermId = null;
let materialsRequest = 0;
let wizard = null;
let importedHeadcount = 0;
let lastRegisterFile = null;
let saveTimer = null;
let saveFlight = null;
let saveDirty = false;
let savePaused = false;
const SAVE_DELAY = 800;
const downloadedReportJobs = new Set();
let reportWatch = 0;
let focusCell = { row: 0, col: 0 };
const selectedRows = new Set();

function clampCount(value) {
  const number = Number(value) || 2;
  return Math.max(1, Math.min(8, Math.round(number)));
}

function fieldText(value) {
  if (value == null) return "";
  const kind = typeof value;
  if (kind === "string" || kind === "number" || kind === "boolean") return String(value);
  if (Array.isArray(value)) return value.map(fieldText).filter(Boolean).join("、");
  if (kind === "object") {
    if (Object.prototype.hasOwnProperty.call(value, "value")) return fieldText(value.value);
    if (Object.prototype.hasOwnProperty.call(value, "text")) return fieldText(value.text);
    return "";
  }
  return "";
}

function normalizeField(item) {
  if (item == null || item === "") return { value: "", status: "需手填", reason: "", note: "", candidates: [], source: "" };
  if (typeof item !== "object" || Array.isArray(item)) {
    const value = fieldText(item);
    return { value, status: value ? "已填" : "需手填", reason: "", note: "", candidates: [], source: "" };
  }
  const rawCandidates = item.candidates || item.candidate || [];
  const list = Array.isArray(rawCandidates)
    ? rawCandidates
    : String(rawCandidates || "").split(/[；;]/).map((part) => part.trim()).filter(Boolean);
  const value = fieldText(item);
  return {
    value,
    status: item.status || (value ? "已填" : "需手填"),
    reason: fieldText(item.reason),
    note: fieldText(item.note),
    source: fieldText(item.source),
    candidates: list.map((candidate) => {
      if (candidate == null || typeof candidate !== "object") {
        const text = fieldText(candidate);
        return { label: text, value: text };
      }
      const candidateValue = fieldText(candidate.value || candidate.label);
      return { label: fieldText(candidate.label || candidate.value), value: candidateValue };
    }).filter((candidate) => candidate.value || candidate.label),
  };
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
  while (state.gradReq.length < state.objectivesCount) state.gradReq.push({ requirement: "", indicator: "", strength: "" });
  while (state.objectiveRequirements.length < state.objectivesCount) state.objectiveRequirements.push("");
  const root = document.getElementById("gradReq");
  root.innerHTML = "";
  for (let i = 0; i < state.objectivesCount; i += 1) {
    const requirement = el("input", { type: "text", value: state.gradReq[i].requirement || "" });
    const indicator = el("input", { type: "text", value: state.gradReq[i].indicator || "" });
    const strength = el("input", { type: "text", value: state.gradReq[i].strength || "" });
    requirement.addEventListener("input", () => { state.gradReq[i].requirement = requirement.value; });
    indicator.addEventListener("input", () => { state.gradReq[i].indicator = indicator.value; });
    strength.addEventListener("input", () => { state.gradReq[i].strength = strength.value; });
    root.append(el("div", { class: "fields" }, [
      el("p", { class: "span-2", text: `课程目标${i + 1}` }),
      el("label", { class: "field", text: "支撑的毕业要求" }, [requirement]),
      el("label", { class: "field", text: "支撑的毕业要求指标点" }, [indicator]),
      el("label", { class: "field", text: "支撑强度 H/M/L" }, [strength]),
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
  const openInfo = { ...state.courseOpen, term: state.courseOpen.semester || "" };
  const payload = buildRelationPayload();
  const ratios = { usual: 0, midterm: 0, final: 0 };
  iterGridRows().forEach((group) => {
    const ratio = parsePortion(group.ratioText);
    if (!Number.isFinite(ratio)) return;
    if (group.name.includes("平时")) ratios.usual += ratio;
    else if (group.name.includes("期中")) ratios.midterm += ratio;
    else if (group.name.includes("期末")) ratios.final += ratio;
  });
  const parsedCount = Number(state.studentCount);
  const basicCount = Number(state.courseBasic.student_count);
  const studentCount = parsedCount > 0 ? parsedCount : (basicCount > 0 ? basicCount : 0);
  return {
    mode: state.mode,
    student_count: studentCount,
    course_open_info: openInfo,
    course_basic_info: basic,
    ratios,
    relation_grid: gridForSave(),
    relation_payload: payload,
    grad_req_map: state.gradReq.slice(0, state.objectivesCount).map((row, index) => ({
      objective: `课程目标${index + 1}`,
      requirement: row.requirement || "",
      indicator: row.indicator || "",
      strength: row.strength || "",
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
    bag[key] = source[key] == null ? "" : fieldText(source[key]);
  });
}

function openingFields() {
  const fields = OPEN_FIELDS.slice();
  const students = String(state.courseBasic.student_count || "").trim();
  const exams = String(state.courseBasic.exam_count || "").trim();
  if (students && students !== "0") fields.push(["student_count", "上课人数"]);
  if (exams && exams !== "0") fields.push(["exam_count", "考核人数"]);
  return fields;
}

function openingValue(key) {
  if (key === "class_name" || key === "major" || key === "student_count" || key === "exam_count") {
    return state.courseBasic[key] || "";
  }
  if (key === "teacher") return state.courseOpen.teacher || state.courseBasic.teacher || "";
  return state.courseOpen[key] || "";
}

function writeOpeningValue(key, value) {
  if (key === "class_name" || key === "major" || key === "exam_count") {
    state.courseBasic[key] = value;
    return;
  }
  if (key === "student_count") {
    state.courseBasic.student_count = value;
    const count = Number(value);
    state.studentCount = count > 0 ? count : "";
    return;
  }
  if (key === "teacher") {
    state.courseOpen.teacher = value;
    state.courseBasic.teacher = value;
    return;
  }
  state.courseOpen[key] = value;
  if (key === "year_start" || key === "year_end" || key === "semester") {
    const { year_start: start, year_end: end, semester } = state.courseOpen;
    state.courseBasic.school_year_term = start && end && semester ? `${start}-${end}学年第${semester}学期` : "";
  }
}

function renderOpeningFields() {
  const root = document.getElementById("openFields");
  root.innerHTML = "";
  openingFields().forEach(([key, label]) => {
    const input = el("input", { type: "text", value: openingValue(key) });
    input.addEventListener("input", () => writeOpeningValue(key, input.value));
    root.append(el("label", { class: "field", text: label }, [input]));
  });
}

function applyServerSettings(settings) {
  const data = settings || {};
  state.mode = data.mode === "reverse" ? "reverse" : "forward";
  state.studentCount = Number(data.student_count) || state.studentCount || 40;
  fillBag(state.courseOpen, data.course_open_info);
  fillBag(state.courseBasic, data.course_basic_info);
  state.courseDescription = fieldText(data.course_description);
  state.spreadMode = data.spread_mode || state.spreadMode;
  state.distribution = data.distribution || state.distribution;
  state.reportStyle = data.report_style || state.reportStyle;
  state.wordLimit = Number(data.word_limit) || 200;
  if (Array.isArray(data.objective_requirements) && data.objective_requirements.length) {
    state.objectiveRequirements = data.objective_requirements.map((item) => fieldText(item));
  }
  if (Array.isArray(data.grad_req_map) && data.grad_req_map.length) {
    state.gradReq = data.grad_req_map.map((row) => ({
      requirement: fieldText(row && row.requirement),
      indicator: fieldText(row && row.indicator),
      strength: fieldText(row && row.strength),
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

let actionBusy = false;

function syncIdleActions() {
  const hasCourse = Boolean(courseId);
  ["deleteCourseBtn"].forEach((id) => {
    const node = document.getElementById(id);
    if (node) node.disabled = !hasCourse;
  });
  const registerInput = document.getElementById("registerFile");
  if (registerInput) registerInput.disabled = !hasCourse;
  const reportBtn = document.getElementById("reportBtn");
  const hasGrade = courseFiles.some((file) => file.kind === "grade");
  if (reportBtn) reportBtn.disabled = actionBusy || !hasGrade;
  const reportHint = document.getElementById("reportHint");
  if (reportHint) {
    reportHint.textContent = actionBusy ? "报告正在生成，请稍候…" : !hasCourse
      ? "请先新建或选择课程，再导入成绩登记表。" : !hasGrade ? "请先导入本学期的成绩登记表。" : "";
    reportHint.hidden = !reportHint.textContent;
  }
}

function setBusy(busy) {
  actionBusy = busy;
  ["templateBtn", "calcBtn", "exportBtn"].forEach((id) => {
    const node = document.getElementById(id);
    if (node) node.disabled = busy;
  });
  syncIdleActions();
}

let statusTimer = null;

/* 状态提示（保存成功、出错、进度）。计算结果在“计算结果”面板里单独显示。 */
function showResult(text, isError, sticky) {
  const node = document.getElementById("result");
  node.textContent = text;
  node.classList.toggle("error", Boolean(isError));
  clearTimeout(statusTimer);
  if (text && !isError && !sticky && !/…$/.test(text)) {
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
  renderOpeningFields();
  renderScalarFields(document.getElementById("basicFields"), BASIC_FIELDS, state.courseBasic);
  syncControls();
  renderGrid();
  renderGrad();
  updateRatioSum();
  renderNoiseTargets();
}

function renderFiles() {
  const grade = courseFiles.find((file) => file.kind === "grade");
  const previous = courseFiles.find((file) => file.kind === "previous");
  document.getElementById("fileName").textContent = grade ? `已保存：${grade.original_name}` : "尚未导入成绩表";
  document.getElementById("prevName").textContent = previous
    ? `已保存：${previous.original_name}`
    : "导入新学期成绩时会自动读取上一学期达成度；也可手动上传覆盖。没有上一学期或不曾计算则为 —";
  updateSummaries();
  if (document.getElementById("dlgFiles").open) loadMaterials();
}

function materialDate(value) {
  return value ? new Date(value).toLocaleString("zh-CN", { hour12: false }) : "生成时间未记录";
}

function materialDownload(file, text = "下载") {
  const button = el("button", { type: "button", class: text === "下载本学期资料包" ? "" : "secondary small", text });
  button.addEventListener("click", async () => {
    button.disabled = true;
    try { await downloadSaved(file); } finally { button.disabled = false; }
  });
  return button;
}

function materialFileGroups(files) {
  const root = el("div", { class: "material-groups" });
  const groups = [
    ["报告文档", (file) => /\.docx$/i.test(file.original_name)],
    ["统计表格", (file) => /\.xlsx$/i.test(file.original_name)],
    ["其他文件", (file) => !/\.(docx|xlsx)$/i.test(file.original_name)],
  ];
  groups.forEach(([title, matches]) => {
    const members = files.filter(matches).sort((a, b) => a.original_name.localeCompare(b.original_name, "zh-CN", { numeric: true }));
    if (!members.length) return;
    const list = el("ul", { class: "file-list" });
    members.forEach((file) => list.append(el("li", {}, [
      el("div", { class: "material-file-name" }, [
        el("span", { text: file.original_name }),
        el("small", { class: "hint", text: materialDate(file.created_at) }),
      ]), materialDownload(file),
    ])));
    root.append(el("p", { class: "dlg-section", text: title }), list);
  });
  return root;
}

function materialFold(title, body) {
  return el("details", { class: "material-fold" }, [el("summary", { text: title }), body]);
}

function renderMaterials(data) {
  const root = document.getElementById("savedFiles");
  root.replaceChildren();
  document.getElementById("dlgFilesTitle").textContent = `${data.course_name} · 课程资料`;
  if (!data.terms.length) {
    root.append(el("p", { class: "hint", text: "还没有学期资料，请先导入成绩登记表。" }));
    return;
  }
  data.terms.forEach((term) => {
    const card = el("section", { class: "material-term", "data-term-id": String(term.id) });
    const heading = el("div", { class: "material-heading" }, [el("h3", { text: term.label })]);
    if (Number(term.id) === Number(currentTermId)) heading.append(el("span", { class: "material-badge", text: "当前学期" }));
    card.append(heading);
    const detail = [term.class_name, term.student_count ? `上课 ${term.student_count} 人` : ""].filter(Boolean).join(" · ");
    if (detail) card.append(el("p", { class: "hint", text: detail }));
    const latest = term.versions[0];
    if (latest) {
      const label = latest.kind === "report" ? "AI 分析报告与统计资料" : "统计资料";
      card.append(el("p", { class: "material-version-note", text: `最近生成：${materialDate(latest.archive.created_at)} · ${label}` }));
      card.append(el("div", { class: "material-package" }, [
        el("span", { text: term.download_name }), materialDownload(latest.archive, "下载本学期资料包"),
      ]));
      if (latest.files.length) card.append(materialFold(`查看单个文件（${latest.files.length}）`, materialFileGroups(latest.files)));
      if (term.versions.length > 1) {
        const history = el("div", { class: "material-history" });
        term.versions.slice(1).forEach((version, index) => {
          const item = el("div", { class: "material-history-item" }, [
            el("div", { class: "material-package" }, [
              el("span", { text: `第 ${term.versions.length - index - 1} 版 · ${materialDate(version.archive.created_at)} · ${version.kind === "report" ? "含 AI 分析报告" : "统计资料"}` }),
              materialDownload(version.archive, "下载这一版"),
            ]),
          ]);
          if (version.files.length) item.append(materialFold("查看这一版的单个文件", materialFileGroups(version.files)));
          history.append(item);
        });
        card.append(materialFold(`查看历史版本（${term.versions.length - 1}）`, history));
      }
    } else {
      card.append(el("p", { class: "hint", text: "本学期尚未生成资料包，完成计算或生成报告后即可下载。" }));
    }
    if (term.source_files.length) card.append(materialFold("查看导入资料和模板", materialFileGroups(term.source_files)));
    if (term.other_files.length) card.append(materialFold(`其他已保存文件（${term.other_files.length}）`, materialFileGroups(term.other_files)));
    root.append(card);
  });
}

async function loadMaterials() {
  const wantedCourse = courseId;
  const requestId = ++materialsRequest;
  const root = document.getElementById("savedFiles");
  root.replaceChildren(el("p", { class: "hint", text: "正在读取各学期资料…" }));
  try {
    const response = await api(`/api/courses/${wantedCourse}/materials`);
    if (!response.ok) throw new Error(await errorMessage(response));
    const data = await response.json();
    if (requestId !== materialsRequest || courseId !== wantedCourse) return;
    renderMaterials(data);
  } catch (err) {
    if (requestId !== materialsRequest || courseId !== wantedCourse) return;
    const retry = el("button", { type: "button", class: "secondary", text: "重新读取" });
    retry.addEventListener("click", loadMaterials);
    root.replaceChildren(el("p", { class: "hint", text: err.message === "未登录" ? "请重新登录" : "资料读取失败，请重试。" }), retry);
  }
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

function setSaveStatus(state, text) {
  const node = document.getElementById("saveStatus");
  if (!node) return;
  node.dataset.state = state;
  if (text != null) node.textContent = text;
}

function scheduleSave() {
  if (savePaused || !courseId) return;
  saveDirty = true;
  setSaveStatus("saving", "保存中…");
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    flushSave().catch(() => {});
  }, SAVE_DELAY);
}

function editTarget(node) {
  if (!node || !node.closest) return false;
  if (node.closest("#passwordPanel, #dlgWizard, #dlgConfirm, #dlgReport, #dlgRegister, #dlgImportTerm")) return false;
  if (node.id === "courseSelect" || node.id === "termSelect" || node.id === "saveStatus") return false;
  if ((node.type || "") === "file") return false;
  return Boolean(node.closest("main"));
}

async function flushSave() {
  clearTimeout(saveTimer);
  saveTimer = null;
  if (saveFlight) {
    try {
      await saveFlight;
    } catch (err) {
      /* 上一次失败已经写在保存状态上 */
    }
  }
  if (savePaused || !courseId || !saveDirty) return null;
  saveDirty = false;
  const courseAtStart = courseId;
  const name = document.getElementById("courseNameInput").value.trim();
  let settings;
  try {
    settings = collectSettings();
  } catch (err) {
    saveDirty = true;
    setSaveStatus("error", "保存失败（点此重试）");
    showResult(err.message || "保存失败", true);
    throw err;
  }
  setSaveStatus("saving", "保存中…");
  const flight = api(`/api/courses/${courseAtStart}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, settings }),
  });
  saveFlight = flight;
  let result = null;
  try {
    const response = await flight;
    if (!response.ok) throw new Error(await errorMessage(response));
    const data = await response.json();
    const option = document.querySelector(`#courseSelect option[value="${courseAtStart}"]`);
    if (option && data.name) option.textContent = data.name;
    if (!saveDirty && courseId === courseAtStart) setSaveStatus("saved", "已保存");
    result = data;
  } catch (err) {
    if (courseId === courseAtStart) {
      saveDirty = true;
      setSaveStatus("error", "保存失败（点此重试）");
      showResult(err.message || "保存失败", true);
    }
    throw err;
  } finally {
    if (saveFlight === flight) saveFlight = null;
  }
  if (saveDirty && courseId === courseAtStart) return flushSave();
  return result;
}

function flushSaveSync() {
  clearTimeout(saveTimer);
  saveTimer = null;
  if (!courseId || !saveDirty) return true;
  const name = document.getElementById("courseNameInput").value.trim();
  try {
    const xhr = new XMLHttpRequest();
    xhr.open("PATCH", `/api/courses/${courseId}`, false);
    xhr.setRequestHeader("Content-Type", "application/json");
    xhr.send(JSON.stringify({ name, settings: collectSettings() }));
    if (xhr.status >= 200 && xhr.status < 300) {
      saveDirty = false;
      setSaveStatus("saved", "已保存");
      return true;
    }
  } catch (err) {
    /* 部分浏览器会拦截页面关闭时的同步请求 */
  }
  setSaveStatus("error", "保存失败（点此重试）");
  return false;
}

async function saveCourse() {
  if (!courseId) throw new Error("请先新建课程文件夹");
  saveDirty = true;
  return flushSave();
}

function overlayTerm(term) {
  if (!term) return;
  state.courseOpen.year_start = term.year_start || "";
  state.courseOpen.year_end = term.year_end || "";
  state.courseOpen.semester = term.semester || "";
  state.courseOpen.teacher = term.teacher || "";
  state.courseBasic.teacher = term.teacher || "";
  state.courseBasic.major = term.major || "";
  state.courseBasic.school_year_term = term.school_year_term || "";
  state.courseBasic.class_name = term.class_name || "";
  state.courseBasic.student_count = term.student_count ? String(term.student_count) : "";
  state.courseBasic.exam_count = term.exam_count ? String(term.exam_count) : "";
  state.studentCount = term.student_count ? term.student_count : "";
  if (term.mode) state.mode = term.mode === "reverse" ? "reverse" : "forward";
  if (term.spread_mode) state.spreadMode = term.spread_mode;
  if (term.distribution) state.distribution = term.distribution;
  if (term.report_style) state.reportStyle = term.report_style;
  if (term.word_limit) state.wordLimit = term.word_limit;
  if (term.noise_config && typeof term.noise_config === "object") {
    state.noiseEnabled = true;
    state.noiseRatio = Number(term.noise_config.noise_ratio);
    if (!Number.isFinite(state.noiseRatio)) state.noiseRatio = 0.1;
    state.severityMode = term.noise_config.severity_mode || "random";
    state.noiseAllowed = Array.isArray(term.noise_config.allowed_items) ? term.noise_config.allowed_items.slice() : [];
  }
}

function renderTerms() {
  const select = document.getElementById("termSelect");
  select.innerHTML = "";
  select.size = 1;
  select.removeAttribute("size");
  if (!terms.length) {
    select.append(el("option", { value: "", text: courseId ? "尚未导入学期" : "未选课程" }));
    return;
  }
  terms.forEach((term) => {
    const text = term.class_name ? `${term.label} · ${term.class_name}` : term.label;
    select.append(el("option", { value: String(term.id), text }));
  });
  if (currentTermId) select.value = String(currentTermId);
}

function applyCoursePayload(data) {
  savePaused = true;
  clearTimeout(saveTimer);
  saveTimer = null;
  saveDirty = false;
  try {
  courseId = data.id;
  localStorage.setItem(COURSE_KEY, String(courseId));
  document.getElementById("courseSelect").value = String(courseId);
  document.getElementById("courseNameInput").value = data.name || "";
  state = defaultState();
  applyServerSettings(data.settings || {});
  overlayTerm(data.current_term);
  if (!data.current_term) state.studentCount = "";
  terms = data.terms || [];
  currentTermId = data.current_term_id || (data.current_term && data.current_term.id) || null;
  courseFiles = data.files || [];
  importedHeadcount = 0;
  const registerBanner = document.getElementById("registerBanner");
  if (registerBanner) {
    registerBanner.hidden = true;
    registerBanner.textContent = "";
  }
  renderTerms();
  renderAll();
  renderFiles();
  setSaveStatus("saved", "已保存");
  } finally {
    savePaused = false;
  }
  afterCourseLoaded();
}

async function openCourse(id) {
  const response = await api(`/api/courses/${id}`);
  if (!response.ok) {
    showResult(await errorMessage(response), true);
    return;
  }
  applyCoursePayload(await response.json());
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
    terms = [];
    currentTermId = null;
    select.append(el("option", { value: "", text: "尚未创建课程" }));
    renderTerms();
    renderFiles();
    setSaveStatus("idle", "");
    afterCourseLoaded();
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
    if (kind === "grade") markCurrentResultStale();
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
    state.courseBasic.student_count = String(event.target.value || "").trim();
  });
  document.getElementById("spreadMode").addEventListener("change", (event) => { state.spreadMode = event.target.value; });
  document.getElementById("distribution").addEventListener("change", (event) => { state.distribution = event.target.value; });
  document.getElementById("reportStyle").addEventListener("change", (event) => { state.reportStyle = event.target.value; });
  document.getElementById("wordLimit").addEventListener("input", (event) => { state.wordLimit = Number(event.target.value); });
  document.getElementById("noiseEnabled").addEventListener("change", (event) => { state.noiseEnabled = event.target.checked; });
  document.getElementById("noiseRatio").addEventListener("input", (event) => { state.noiseRatio = Number(event.target.value); });
  document.getElementById("severityMode").addEventListener("change", (event) => { state.severityMode = event.target.value; });
  document.getElementById("modeForward").addEventListener("click", () => { state.mode = "forward"; syncMode(); scheduleSave(); });
  document.getElementById("modeReverse").addEventListener("click", () => { state.mode = "reverse"; syncMode(); scheduleSave(); });
  document.getElementById("addRowBtn").addEventListener("click", () => {
    state.grid.push(blankRow());
    renderGrid();
    scheduleSave();
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
    scheduleSave();
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
  document.getElementById("saveStatus").addEventListener("click", () => {
    if (document.getElementById("saveStatus").dataset.state !== "error") return;
    flushSave().catch(() => {});
  });
  window.addEventListener("beforeunload", (event) => {
    if (!saveDirty && !saveTimer) return;
    saveDirty = true;
    if (!flushSaveSync()) {
      event.preventDefault();
      event.returnValue = "课程修改还没保存";
    }
  });
  document.getElementById("deleteCourseBtn").addEventListener("click", async () => {
    if (!courseId) return;
    const name = document.getElementById("courseNameInput").value.trim() || "当前课程";
    if (!window.confirm(`删除课程文件夹「${name}」及其文件？`)) return;
    clearTimeout(saveTimer);
    saveTimer = null;
    const pending = saveDirty;
    saveDirty = false;
    const response = await api(`/api/courses/${courseId}`, { method: "DELETE" });
    if (!response.ok) {
      saveDirty = pending;
      return showResult(await errorMessage(response), true);
    }
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
    updateSummaries();
    renderResultView();
    showResult("计算完成。", false);
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

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function formatReportError(data) {
  const stage = data.failed_stage || data.stage || "";
  const names = {
    calculate: "计算达成度",
    tables: "生成统计表",
    ai: "AI 撰写分析",
    package: "打包",
  };
  const name = names[stage] || "生成报告";
  let advice = "请稍后重试。";
  if (stage === "ai") advice = "统计表已生成，可在高级选项里只导出统计表，或稍后重试。";
  else if (stage === "calculate" || stage === "tables") advice = "请检查成绩文件和关系表。";
  return `${name}失败：${data.error || "未知错误"}\n${advice}`;
}

async function downloadReportJob(jobId) {
  const response = await api(`/api/report-jobs/${jobId}/download`);
  if (!response.ok) {
    showResult(await errorMessage(response), true);
    return;
  }
  const blob = await response.blob();
  const name = filenameFromDisposition(response.headers.get("Content-Disposition"), "AI分析报告.zip");
  downloadBlob(blob, name);
}

async function watchReportJob(jobId) {
  const token = ++reportWatch;
  const dlg = document.getElementById("dlgReport");
  const stageNode = document.getElementById("reportStage");
  const bar = document.getElementById("reportProgress");
  const errorNode = document.getElementById("reportError");
  const download = document.getElementById("reportDownload");
  while (token === reportWatch) {
    let data;
    try {
      const response = await api(`/api/report-jobs/${jobId}`);
      if (!response.ok) {
        stageNode.textContent = "无法获取进度";
        errorNode.hidden = false;
        errorNode.textContent = await errorMessage(response);
        if (!dlg.open) openDialog("dlgReport");
        return;
      }
      data = await response.json();
    } catch (err) {
      if (err.message === "未登录") return;
      if (token !== reportWatch) return;
      await delay(700);
      continue;
    }
    if (token !== reportWatch) return;
    stageNode.textContent = data.stage_label || "正在生成报告…";
    bar.value = Number(data.percent) || 0;
    if (data.error) {
      errorNode.hidden = false;
      errorNode.textContent = formatReportError(data);
      download.hidden = true;
      if (!dlg.open) openDialog("dlgReport");
      showResult(errorNode.textContent, true, true);
      return;
    }
    if (data.done) {
      stageNode.textContent = data.stage_label || "打包完成";
      bar.value = 100;
      errorNode.hidden = true;
      download.hidden = false;
      download.onclick = () => downloadReportJob(jobId);
      if (!dlg.open) openDialog("dlgReport");
      await refreshFiles();
      if (data.summary) storeResult({ ...data.summary, reported: true });
      updateSummaries();
      showResult("报告已生成。", false);
      if (!downloadedReportJobs.has(jobId)) {
        downloadedReportJobs.add(jobId);
        await downloadReportJob(jobId);
      }
      return;
    }
    await delay(700);
  }
}

async function generateReport() {
  const problem = validateBeforeRun(true);
  if (problem) return showResult(problem, true);
  const stageNode = document.getElementById("reportStage");
  const bar = document.getElementById("reportProgress");
  const errorNode = document.getElementById("reportError");
  const download = document.getElementById("reportDownload");
  stageNode.textContent = "正在计算达成度…";
  bar.value = 0;
  errorNode.hidden = true;
  errorNode.textContent = "";
  download.hidden = true;
  if (!document.getElementById("dlgReport").open) openDialog("dlgReport");
  setBusy(true);
  try {
    const response = await postCourse("/report-jobs");
    if (!response.ok) {
      const message = await errorMessage(response);
      errorNode.hidden = false;
      errorNode.textContent = message;
      showResult(message, true);
      return;
    }
    const data = await response.json();
    if (data.stage_label) stageNode.textContent = data.stage_label;
    if (data.percent) bar.value = Number(data.percent) || 0;
    await watchReportJob(data.job_id);
  } catch (err) {
    if (err.message !== "未登录") {
      errorNode.hidden = false;
      errorNode.textContent = err.message || "生成失败";
      showResult(err.message || "生成失败", true);
    }
  } finally {
    setBusy(false);
  }
}

document.getElementById("reportBtn").addEventListener("click", () => {
  generateReport();
});

/* ---------- 主界面按钮 + 弹窗（对应桌面版 ui_app 的主窗口和各对话框） ---------- */
const RESULT_PREFIX = "calculatorpro.result.";
const RATIO_LINKS = [["平时考核", "平时"], ["期中考核", "期中"], ["期末考核", "期末"]];
let summaryTimer = null;
let dialogSnapshot = null;
let aiEnabled = false;

function readJson(key) {
  try {
    return JSON.parse(localStorage.getItem(key) || "null");
  } catch (err) {
    return null;
  }
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

function resultStorageKey(termId) {
  return `${RESULT_PREFIX}${courseId}.${termId}`;
}

function currentGradeId() {
  const grade = courseFiles.find((file) => file.kind === "grade");
  return grade ? Number(grade.id) : 0;
}

function storeResult(data) {
  const termId = Number(data.term_id || currentTermId);
  if (!courseId || !termId) return;
  localStorage.removeItem(RESULT_PREFIX + courseId);
  localStorage.setItem(resultStorageKey(termId), JSON.stringify({
    mode: data.mode,
    course_name: data.course_name,
    student_count: data.student_count,
    average_score: data.average_score,
    achievement: data.achievement || {},
    files: data.files || [],
    term_id: termId,
    term_label: data.term_label || "",
    class_name: data.class_name || "",
    warnings: data.warnings || [],
    grade_id: currentGradeId(),
    reported: Boolean(data.reported),
    stale: false,
    at: new Date().toISOString(),
  }));
}

function loadStoredResult() {
  if (!courseId || !currentTermId) return null;
  localStorage.removeItem(RESULT_PREFIX + courseId);
  const data = readJson(resultStorageKey(currentTermId));
  if (!data || Number(data.term_id) !== Number(currentTermId)) return null;
  const gradeId = currentGradeId();
  if (data.stale || (Number(data.grade_id || 0) && Number(data.grade_id) !== gradeId)) {
    return { ...data, stale: true };
  }
  return data;
}

function markCurrentResultStale() {
  if (!courseId || !currentTermId) return;
  const data = readJson(resultStorageKey(currentTermId));
  if (!data) return;
  data.stale = true;
  localStorage.setItem(resultStorageKey(currentTermId), JSON.stringify(data));
}

function resultNote(data) {
  if (!courseId) return "请先新建或选择课程文件夹。";
  if (data && data.stale) return "成绩已更换，请重新计算";
  return "本学期尚未计算";
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

/* 主界面的结果卡片：总达成度 + 关键数字，详情在弹窗里 */
function renderResultCard() {
  const card = document.getElementById("resultCard");
  card.innerHTML = "";
  const data = loadStoredResult();
  const filesBtn = el("button", { type: "button", class: "ghost small", text: "查看各学期资料" });
  filesBtn.addEventListener("click", () => openDialog("dlgFiles"));
  card.dataset.state = !courseId ? "none" : data && data.stale ? "stale" : data ? "ready" : "empty";
  card.dataset.termId = currentTermId ? String(currentTermId) : "";
  if (!data || data.stale) {
    card.append(el("div", { class: "result-empty" }, [
      el("span", { id: "resultTermNote", class: "hint", text: resultNote(data) }),
      filesBtn,
    ]));
    syncOpenResult();
    return;
  }
  const { total } = totalAchievement(data.achievement);
  (data.warnings || []).forEach((item) => card.append(el("p", { class: "headcount-alert", text: item })));
  const when = data.at ? new Date(data.at).toLocaleString("zh-CN", { hour12: false }) : "";
  const termBit = data.term_label ? `${data.term_label} · ` : "";
  const detailBtn = el("button", { type: "button", class: "secondary small", text: "查看详情" });
  detailBtn.addEventListener("click", () => openDialog("dlgResult"));
  card.append(
    el("div", { class: "stats" }, [
      stat("总达成度", total, true),
      stat("学生人数", data.student_count),
      stat("平均分", data.average_score),
      stat("模式", data.mode === "reverse" ? "逆向" : "正向"),
    ]),
    el("div", { class: "result-foot" }, [
      el("span", { class: "hint", text: `${termBit}${data.course_name || ""} · 计算于 ${when}` }),
      el("span", { class: "row-actions" }, [detailBtn, filesBtn]),
    ]),
  );
  syncOpenResult();
}

function syncOpenResult() {
  const dlg = document.getElementById("dlgResult");
  if (dlg && dlg.open) renderResultView();
}

/* 结果详情弹窗：各课程目标、学生明细下载、生成的文件 */
function renderResultView() {
  const root = document.getElementById("resultView");
  root.innerHTML = "";
  const data = loadStoredResult();
  if (!data || data.stale) {
    root.append(el("p", { class: "hint", text: resultNote(data) }));
    return;
  }
  const { total, parts } = totalAchievement(data.achievement);
  root.append(el("div", { class: "stats" }, [
    stat("总达成度", total, true),
    stat("学生人数", data.student_count),
    stat("平均分", data.average_score),
    stat("模式", data.mode === "reverse" ? "逆向" : "正向"),
  ]));
  const table = el("table", { class: "kv" });
  parts.forEach(([key, value]) => {
    table.append(el("tr", {}, [el("th", { text: key }), el("td", { text: String(value) })]));
  });
  (data.warnings || []).forEach((item) => root.append(el("p", { class: "headcount-alert", text: item })));
  root.append(el("p", { class: "hint", text: `学期：${data.term_label || "—"}　班级：${data.class_name || "—"}` }));
  root.append(el("p", { class: "dlg-section", text: `各课程目标达成度（${parts.length}）` }), table);
  root.append(el("p", { class: "dlg-section", text: `学生成绩明细（${data.student_count || 0} 人）` }));
  root.append(el("p", { class: "hint", text: "每位学生的各考核方式得分与课程目标得分在“成绩明细.xlsx”里。" }));
  const detailFile = courseFiles.find((file) => file.kind === "output" && /成绩明细\.xlsx$/.test(file.original_name || ""));
  if (detailFile) {
    const button = el("button", { type: "button", class: "secondary", text: `下载 ${detailFile.original_name}` });
    button.addEventListener("click", () => downloadSaved(detailFile));
    root.append(button);
  }
  const list = el("ul", { class: "plain-list" });
  (data.files || []).forEach((name) => list.append(el("li", { text: name })));
  root.append(el("p", { class: "dlg-section", text: `本次生成的文件（${(data.files || []).length}）` }), list);
}

/* 成绩占比弹窗：三项占比写回对应关系表的“占比”列 */
function currentRatios() {
  const found = {};
  iterGridRows().forEach((group) => {
    const ratio = parsePortion(group.ratioText);
    if (Number.isFinite(ratio)) found[group.name] = ratio;
  });
  return found;
}

function renderRatioFields() {
  const root = document.getElementById("ratioFields");
  root.innerHTML = "";
  const found = currentRatios();
  RATIO_LINKS.forEach(([name, label]) => {
    const value = found[name];
    const input = el("input", { type: "number", min: "0", max: "1", step: "0.05", "data-link": name });
    input.value = Number.isFinite(value) ? String(value) : "0";
    input.addEventListener("input", updateRatioDialogSum);
    root.append(el("label", { class: "field", text: `${label}占比（${name}）` }, [input]));
  });
  updateRatioDialogSum();
}

function updateRatioDialogSum() {
  let sum = 0;
  document.querySelectorAll("#ratioFields input").forEach((input) => { sum += Number(input.value) || 0; });
  const ok = Math.abs(sum - 1) < 0.001;
  const note = document.getElementById("ratioDlgSum");
  note.textContent = `合计 ${sum.toFixed(2)}${ok ? "" : "，需要等于 1.00"}`;
  note.classList.toggle("bad", !ok);
  return ok;
}

function applyRatioFields() {
  if (!updateRatioDialogSum()) throw new Error("三项占比之和必须等于 1.0");
  document.querySelectorAll("#ratioFields input").forEach((input) => {
    const name = input.dataset.link;
    const value = String(Number(input.value) || 0);
    const index = state.grid.findIndex((row) => String(row[0] || "").trim() === name);
    if (index >= 0) {
      state.grid[index][1] = value;
    } else if (Number(value) > 0) {
      const row = blankRow();
      row[0] = name;
      row[1] = value;
      state.grid.push(row);
    }
  });
  renderGrid();
  updateRatioSum();
  renderNoiseTargets();
}

/* 弹窗：打开时拍快照；取消/Esc/× 时若有改动先确认，再恢复快照；保存则写服务器后关闭 */
function snapshot() {
  return JSON.stringify({ state, objCount: state.objectivesCount, ratios: Array.from(document.querySelectorAll("#ratioFields input")).map((i) => i.value) });
}

function openDialog(id) {
  const dlg = document.getElementById(id);
  if (!dlg) return;
  if (dlg.hasAttribute("data-needs-course") && !courseId) return showResult("请先新建或选择课程文件夹", true);
  if (id === "dlgRatio") renderRatioFields();
  if (id === "dlgResult") renderResultView();
  if (id === "dlgFiles") loadMaterials();
  dialogSnapshot = snapshot();
  dlg.showModal();
  const first = dlg.querySelector(".dlg-body input, .dlg-body textarea, .dlg-body select, .dlg-body button");
  if (first) first.focus();
}

function dialogDirty() {
  return dialogSnapshot !== null && dialogSnapshot !== snapshot();
}

function closeDialog(dlg, force) {
  const savable = Boolean(dlg.querySelector("[data-save]"));
  if (!force && savable && dialogDirty()) {
    if (!window.confirm("有未保存的修改，确定放弃吗？")) return false;
    const saved = JSON.parse(dialogSnapshot).state;
    state = saved;
    renderAll();
    saveDirty = true;
    scheduleSave();
  }
  dialogSnapshot = null;
  dlg.close();
  updateSummaries();
  return true;
}

async function saveDialog(dlg) {
  try {
    if (dlg.id === "dlgRatio") applyRatioFields();
    if (!courseId) throw new Error("请先新建课程文件夹");
    const saved = await saveCourse();
    if (dlg.id === "dlgOpen" && saved && saved.id) applyCoursePayload(saved);
    closeDialog(dlg, true);
  } catch (err) {
    if (err.message !== "未登录") showResult(err.message || "保存失败", true);
  }
}

function bindDialogs() {
  document.querySelectorAll("[data-open]").forEach((button) => {
    button.addEventListener("click", () => openDialog(button.dataset.open));
  });
  document.querySelectorAll("dialog.dlg").forEach((dlg) => {
    dlg.setAttribute("closedby", "closerequest");
    const form = dlg.querySelector("form");
    if (form) form.addEventListener("submit", (event) => event.preventDefault());
    dlg.addEventListener("cancel", (event) => {
      event.preventDefault();
      closeDialog(dlg, false);
    });
    dlg.querySelectorAll("[data-close]").forEach((button) => {
      button.addEventListener("click", () => closeDialog(dlg, false));
    });
    const save = dlg.querySelector("[data-save]");
    if (save) save.addEventListener("click", () => saveDialog(dlg));
  });
  const refresh = () => {
    clearTimeout(summaryTimer);
    summaryTimer = setTimeout(updateSummaries, 150);
  };
  const onField = (event) => {
    refresh();
    if (editTarget(event.target)) scheduleSave();
  };
  document.addEventListener("input", onField);
  document.addEventListener("change", onField);
  document.querySelector("main").addEventListener("click", (event) => {
    if (event.target.closest("#modeForward, #modeReverse, #addRowBtn, #removeRowsBtn")) refresh();
  });
}

function filledState(values) {
  const filled = values.filter((value) => String(value || "").trim()).length;
  if (!filled) return ["todo", "未填"];
  if (filled === values.length) return ["done", "✓ 已填"];
  return ["part", `部分 ${filled}/${values.length}`];
}

/* 主界面按钮上的完成状态 */
function updateSummaries() {
  const set = (key, [tone, text]) => {
    const node = document.querySelector(`[data-status="${key}"]`);
    if (!node) return;
    node.textContent = text;
    node.dataset.tone = tone;
    const card = node.closest(".top-btn");
    if (card) card.dataset.tone = tone;
  };
  set("open", filledState(openingFields().map(([key]) => openingValue(key))));
  set("basic", filledState(BASIC_FIELDS.map(([key]) => state.courseBasic[key])));
  const found = currentRatios();
  const ratioValues = RATIO_LINKS.map(([name]) => found[name] || 0);
  const ratioSum = ratioValues.reduce((a, b) => a + b, 0);
  set("ratio", Math.abs(ratioSum - 1) < 0.001
    ? ["done", `✓ ${ratioValues.map((v) => Math.round(v * 100)).join("/")}`]
    : ratioSum > 0 ? ["part", `合计 ${ratioSum.toFixed(2)}`] : ["todo", "未填"]);
  const rows = state.grid.filter((row) => row.some((cell) => String(cell || "").trim())).length;
  set("relation", relationReady() ? ["done", `✓ ${state.objectivesCount} 个目标 · ${rows} 行`] : rows ? ["part", `未完成 · ${rows} 行`] : ["todo", "未填"]);
  const gradValues = [];
  state.gradReq.slice(0, state.objectivesCount).forEach((item) => {
    gradValues.push(item.requirement, item.indicator, item.strength);
  });
  set("grad", filledState(gradValues));
  const previous = courseFiles.find((file) => file.kind === "previous");
  const settingsState = filledState([state.courseDescription, ...state.objectiveRequirements.slice(0, state.objectivesCount)]);
  if (previous) settingsState[1] += " · 含上一学年";
  set("settings", settingsState);
  const term = terms.find((item) => item.id === currentTermId);
  set("term", term ? [term.teacher ? "done" : "part", term.label] : ["todo", "未选学期"]);
  const grade = courseFiles.find((file) => file.kind === "grade");
  const fileNode = document.getElementById("fileName");
  fileNode.textContent = grade ? `✓ ${grade.original_name}` : "未导入";
  fileNode.dataset.tone = grade ? "done" : "todo";
  syncIdleActions();
  const registerHint = document.getElementById("registerHint");
  if (registerHint) {
    if (grade) {
      const count = importedHeadcount || (term && term.student_count) || 0;
      const modeLabel = state.mode === "reverse" ? "逆向" : "正向";
      registerHint.textContent = count ? `已导入 ${count} 人 · ${modeLabel}` : `已导入 · ${modeLabel}`;
      registerHint.dataset.tone = "done";
    } else {
      registerHint.textContent = "xlsx / xls / pdf";
      registerHint.dataset.tone = "todo";
    }
  }
  const noiseBtn = document.getElementById("noiseBtn");
  noiseBtn.textContent = state.noiseEnabled ? `已配置 ${Math.round((Number(state.noiseRatio) || 0) * 100)}%` : "无";
  noiseBtn.disabled = state.mode !== "reverse";
  renderResultCard();
}

function afterCourseLoaded() {
  updateSummaries();
}

const WIZARD_LABELS = {
  course_name: "课程名称",
  credits: "学分",
  hours: "学时",
  course_type: "课程性质",
  course_code: "课程代码",
  college: "开课学院",
  major: "上课专业",
  school_year_term: "学年学期",
  year_start: "学年起",
  year_end: "学年止",
  semester: "学期",
  teacher: "任课教师",
  class_name: "上课班级",
  student_count: "上课人数",
  exam_count: "考核人数",
};

function appendFieldNotes(wrap, item) {
  if (item.status === "需手填") {
    wrap.append(el("span", { class: "wizard-reason", text: item.reason ? `需手填：${item.reason}` : "需手填" }));
  } else if (item.source) {
    wrap.append(el("span", { class: "hint", text: `来源：${item.source}` }));
  }
  if (item.status !== "需手填" && item.note) {
    wrap.append(el("span", { class: "hint", text: `附注：${item.note}` }));
  } else if (item.status !== "需手填" && item.reason) {
    wrap.append(el("span", { class: "hint", text: item.reason }));
  }
  (item.candidates || []).forEach((candidate) => {
    const button = el("button", {
      type: "button",
      class: "candidate-btn",
      text: candidate.label || candidate.value || "",
    });
    button.addEventListener("click", () => {
      const input = wrap.querySelector("input, textarea");
      if (!input) return;
      input.value = candidate.value || "";
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });
    wrap.append(button);
  });
}

function wizardField(key) {
  if (!wizard.fields[key]) wizard.fields[key] = { value: "", status: "需手填", reason: "", candidates: [], source: "" };
  const item = normalizeField(wizard.fields[key]);
  wizard.fields[key] = item;
  const wrap = el("label", { class: item.status === "需手填" ? "field need-fill" : "field" });
  wrap.append(document.createTextNode(WIZARD_LABELS[key] || key));
  const input = el("input", { type: "text", value: item.value });
  input.addEventListener("input", () => {
    wizard.fields[key].value = input.value;
  });
  wrap.append(input);
  appendFieldNotes(wrap, item);
  return wrap;
}

function renderWizardFields() {
  const courseRoot = document.getElementById("wizardCourse");
  courseRoot.innerHTML = "";
  const courseKeys = (wizard.groups.course || []).filter((key) => !(wizard.groups.term || []).includes(key));
  const typeIndex = courseKeys.indexOf("course_type");
  if (typeIndex > 0) {
    courseKeys.splice(typeIndex, 1);
    courseKeys.unshift("course_type");
  }
  courseKeys.forEach((key) => courseRoot.append(wizardField(key)));
  const description = normalizeField(wizard.descriptionField);
  wizard.descriptionField = description;
  wizard.description = description.value;
  const descriptionBox = document.getElementById("wizardDescription").closest("label");
  descriptionBox.className = description.status === "需手填" ? "field need-fill" : "field";
  descriptionBox.querySelectorAll(".wizard-reason, .hint, .candidate-btn").forEach((node) => node.remove());
  const descriptionInput = document.getElementById("wizardDescription");
  descriptionInput.value = description.value;
  appendFieldNotes(descriptionBox, description);
  const objectives = document.getElementById("wizardObjectives");
  objectives.innerHTML = "";
  (wizard.objectives || []).forEach((text, index) => {
    const item = normalizeField(typeof text === "object" ? text : { value: fieldText(text), status: fieldText(text) ? "已填" : "需手填" });
    wizard.objectives[index] = item.value;
    const area = el("textarea");
    area.value = item.value;
    area.addEventListener("input", () => { wizard.objectives[index] = area.value; });
    const box = el("label", { class: item.status === "需手填" ? "field need-fill" : "field", text: `课程目标${index + 1}` }, [area]);
    appendFieldNotes(box, item);
    objectives.append(box);
  });
  const grad = document.getElementById("wizardGrad");
  grad.innerHTML = "";
  (wizard.grad || []).forEach((item, index) => {
    const box = el("div", { class: "field" });
    box.append(el("p", { class: "dlg-section", text: `课程目标${index + 1}` }));
    [
      ["requirement", "毕业要求"],
      ["indicator", "指标点"],
      ["strength", "支撑强度 H/M/L"],
    ].forEach(([key, label]) => {
      const side = item[key] && typeof item[key] === "object" ? item[key] : gradSide(item[key]);
      wizard.grad[index][key] = side;
      const wrap = el("label", { class: side.status === "需手填" ? "field need-fill" : "field" });
      wrap.append(document.createTextNode(label));
      const input = el("input", { type: "text", value: side.value || "" });
      input.addEventListener("input", () => {
        side.value = input.value;
        side.status = input.value.trim() ? "已填" : "需手填";
        wrap.classList.toggle("need-fill", side.status === "需手填");
      });
      wrap.append(input);
      appendFieldNotes(wrap, side);
      box.append(wrap);
    });
    grad.append(box);
  });
  renderWizardRelation();
}

function renderWizardRelation() {
  const root = document.getElementById("wizardRelation");
  root.innerHTML = "";
  const grid = (wizard.relation && wizard.relation.grid) || [];
  if (!grid.length) {
    root.append(el("p", { class: "hint", text: "没有读到关系表。" }));
  } else {
    const table = el("table", { class: "sheet" });
    grid.forEach((row, rowIndex) => {
      const tr = el("tr");
      row.forEach((cell, colIndex) => {
        const input = el("input", { type: "text", value: fieldText(cell) });
        input.addEventListener("input", () => {
          wizard.relation.grid[rowIndex][colIndex] = input.value;
          wizard.relation.ok = false;
          document.getElementById("wizardCreate").disabled = true;
          document.getElementById("wizardRelationMsg").textContent = "关系表已改动，请重新校验。";
        });
        tr.append(el("td", {}, [input]));
      });
      table.append(tr);
    });
    root.append(table);
  }
  const message = document.getElementById("wizardRelationMsg");
  const errors = (wizard.relation && wizard.relation.errors) || [];
  message.textContent = wizard.relation && wizard.relation.ok
    ? "关系表已通过程序校验。"
    : (errors.join("；") || "关系表未通过校验。");
  message.classList.toggle("bad", !(wizard.relation && wizard.relation.ok));
  document.getElementById("wizardCreate").disabled = !(wizard.relation && wizard.relation.ok);
}

function gradSide(raw) {
  if (raw && typeof raw === "object" && !Array.isArray(raw)) return normalizeField(raw);
  const value = fieldText(raw);
  return { value, status: value ? "已填" : "需手填", reason: "", note: "", source: "", candidates: [] };
}

function showWizardDraft(draft) {
  const fields = draft.fields || {};
  wizard = {
    fields,
    groups: draft.groups || {
      course: ["course_type", "course_name", "credits", "hours", "course_code", "college"],
      term: ["school_year_term", "year_start", "year_end", "semester", "teacher", "major", "class_name", "student_count", "exam_count"],
    },
    descriptionField: draft.course_description || "",
    description: fieldText(draft.course_description),
    objectives: (draft.objectives || []).map((item) => fieldText(item)),
    grad: (draft.grad_req_map || []).map((item) => {
      const strengthValue = fieldText(item.strength);
      return {
        requirement: gradSide(item.requirement),
        indicator: gradSide(item.indicator),
        strength: gradSide({
          value: strengthValue,
          status: strengthValue ? "已填" : "需手填",
          source: fieldText(item.strength_source),
          reason: strengthValue ? "" : fieldText(item.strength_reason),
        }),
      };
    }),
    relation: draft.relation || { ok: false, errors: [], grid: [] },
  };
  document.getElementById("wizardUpload").hidden = true;
  document.getElementById("wizardConfirm").hidden = false;
  renderWizardFields();
}

function resetWizard() {
  wizard = null;
  document.getElementById("wizardUpload").hidden = false;
  document.getElementById("wizardConfirm").hidden = true;
  document.getElementById("wizardCreate").disabled = true;
  document.getElementById("wizardWait").hidden = true;
  document.getElementById("wizardError").textContent = "";
  const syllabus = document.getElementById("syllabusFile");
  if (syllabus) syllabus.value = "";
}

function percentText(value) {
  if (value == null || value === "") return "—";
  return `${Math.round(Number(value) * 1000) / 10}%`;
}

function registerSummary(data) {
  const modeText = data.mode === "reverse" ? "逆向" : "正向";
  const reason = data.reason || (data.detection && data.detection.reason) || "";
  return `识别为${modeText}，共 ${data.student_count} 人。${reason}`;
}

function showRegisterBanner(data) {
  const banner = document.getElementById("registerBanner");
  if (!banner) return;
  banner.hidden = false;
  banner.textContent = registerSummary(data);
}

function syncReimportButtons(mode) {
  const forward = document.getElementById("reimportForward");
  const reverse = document.getElementById("reimportReverse");
  if (!forward || !reverse) return;
  const isForward = mode === "forward";
  const isReverse = mode === "reverse";
  forward.hidden = isForward;
  reverse.hidden = isReverse;
  forward.disabled = isForward;
  reverse.disabled = isReverse;
  forward.classList.toggle("secondary", mode !== "reverse");
  reverse.classList.toggle("secondary", mode !== "forward");
}

function showRegisterResult(data) {
  syncReimportButtons(data.mode);
  const root = document.getElementById("registerResult");
  root.innerHTML = "";
  const percents = data.percents || {};
  root.append(el("p", { id: "registerDetect", class: "headcount-alert", text: registerSummary(data) }));
  root.append(el("p", { text: `识别 ${data.student_count} 人，实考 ${data.exam_count || data.student_count} 人。` }));
  root.append(el("p", { text: `班级：${(data.classes || []).join("、") || "未识别"}` }));
  root.append(el("p", { text: `平时 ${percentText(percents.usual)}，期中 ${percentText(percents.midterm)}，期末 ${percentText(percents.final)}。` }));
  (data.warnings || []).forEach((item) => {
    const loud = String(item).includes("人数不一致");
    root.append(el("p", { class: loud ? "headcount-alert" : "wizard-reason", text: item }));
  });
  (data.conflicts || []).forEach((item) => root.append(el("p", { class: "wizard-reason", text: item })));
  root.append(el("p", { class: "hint", text: "姓名和学号只留在本机这一学期的成绩文件里。" }));
  const dlg = document.getElementById("dlgRegister");
  if (!dlg.open) openDialog("dlgRegister");
}

document.getElementById("syllabusBtn").addEventListener("click", () => {
  resetWizard();
  openDialog("dlgWizard");
});

document.getElementById("syllabusReadBtn").addEventListener("click", async () => {
  const file = document.getElementById("syllabusFile").files[0];
  if (!file) {
    document.getElementById("wizardError").textContent = "请先选择教学大纲。";
    return;
  }
  const body = new FormData();
  body.append("syllabus", file);
  const wait = document.getElementById("wizardWait");
  const error = document.getElementById("wizardError");
  wait.hidden = false;
  error.textContent = "";
  document.getElementById("syllabusReadBtn").disabled = true;
  try {
    const response = await api("/api/syllabus/extract", { method: "POST", body });
    if (!response.ok) {
      error.textContent = await errorMessage(response);
      return;
    }
    showWizardDraft(await response.json());
  } catch (err) {
    if (err.message !== "未登录") error.textContent = err.message || "读取失败";
  } finally {
    wait.hidden = true;
    document.getElementById("syllabusReadBtn").disabled = false;
  }
});

document.getElementById("wizardRecheck").addEventListener("click", async () => {
  if (!wizard) return;
  const response = await api("/api/syllabus/validate-relation", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ relation_grid: wizard.relation.grid }),
  });
  if (!response.ok) {
    document.getElementById("wizardRelationMsg").textContent = await errorMessage(response);
    document.getElementById("wizardCreate").disabled = true;
    return;
  }
  const data = await response.json();
  wizard.relation.ok = Boolean(data.ok);
  wizard.relation.errors = data.errors || [];
  renderWizardRelation();
});

document.getElementById("wizardCreate").addEventListener("click", async () => {
  if (!wizard || !wizard.relation.ok) return;
  if (courseId) {
    try {
      await flushSave();
    } catch (err) {
      return;
    }
  }
  const fields = {};
  Object.entries(wizard.fields).forEach(([key, item]) => {
    fields[key] = (item && item.value) || "";
  });
  const payload = {
    fields,
    objectives: wizard.objectives,
    description: document.getElementById("wizardDescription").value,
    grad_req_map: wizard.grad,
    relation_grid: wizard.relation.grid,
  };
  const response = await api("/api/courses/from-syllabus", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    document.getElementById("wizardRelationMsg").textContent = await errorMessage(response);
    return;
  }
  const created = await response.json();
  closeDialog(document.getElementById("dlgWizard"), true);
  await loadCourses(created.id);
  showResult("已按大纲建课。", false);
});

document.getElementById("termSelect").addEventListener("change", async () => {
  const select = document.getElementById("termSelect");
  const next = Number(select.value);
  if (!courseId || !next || next === currentTermId) return;
  try {
    await flushSave();
  } catch (err) {
    select.value = String(currentTermId || "");
    return;
  }
  const response = await api(`/api/courses/${courseId}/terms/${next}/select`, { method: "POST" });
  if (!response.ok) {
    renderTerms();
    return showResult(await errorMessage(response), true);
  }
  applyCoursePayload(await response.json());
  showResult("已切换学期。", false);
});

function askConfirm(message) {
  const dlg = document.getElementById("dlgConfirm");
  const yes = document.getElementById("confirmYes");
  const no = document.getElementById("confirmNo");
  const close = document.getElementById("confirmNoX");
  document.getElementById("confirmMessage").textContent = message;
  return new Promise((resolve) => {
    const finish = (ok) => {
      yes.removeEventListener("click", onYes);
      no.removeEventListener("click", onNo);
      close.removeEventListener("click", onNo);
      dlg.removeEventListener("cancel", onCancel);
      if (dlg.open) dlg.close();
      resolve(ok);
    };
    const onYes = () => finish(true);
    const onNo = () => finish(false);
    const onCancel = (event) => {
      event.preventDefault();
      finish(false);
    };
    yes.addEventListener("click", onYes);
    no.addEventListener("click", onNo);
    close.addEventListener("click", onNo);
    dlg.addEventListener("cancel", onCancel);
    dlg.showModal();
  });
}

async function postRegister(file, extra) {
  const body = new FormData();
  const options = extra || {};
  body.append("file", file);
  if (options.confirm) body.append("confirm", "1");
  if (options.mode) body.append("mode", options.mode);
  ["year_start", "year_end", "semester", "class_name"].forEach((key) => {
    if (options[key]) body.append(key, options[key]);
  });
  return api(`/api/courses/${courseId}/grade-register`, { method: "POST", body });
}

function askImportPrompt(data) {
  const dlg = document.getElementById("dlgImportTerm");
  const replace = data.code === "replace";
  const needs = data.needs || [];
  const needTerm = needs.includes("term");
  const needClass = needs.includes("class");
  document.getElementById("importPromptText").textContent = data.detail || "";
  document.getElementById("importTermFields").hidden = replace;
  document.getElementById("importYearStartWrap").hidden = !needTerm;
  document.getElementById("importYearEndWrap").hidden = !needTerm;
  document.getElementById("importSemesterWrap").hidden = !needTerm;
  document.getElementById("importClassWrap").hidden = !needClass;
  const select = document.getElementById("importClassSelect");
  select.innerHTML = "";
  if (needClass) {
    select.append(el("option", { value: "", text: "请选择班级" }));
    (data.classes || []).forEach((name) => select.append(el("option", { value: name, text: name })));
  }
  document.getElementById("importFillConfirm").hidden = replace;
  document.getElementById("importReplaceYes").hidden = !replace;
  const yes = document.getElementById("importReplaceYes");
  const fill = document.getElementById("importFillConfirm");
  const no = document.getElementById("importFillCancel");
  const close = document.getElementById("importFillCancelX");
  return new Promise((resolve) => {
    let settled = false;
    const finish = (value) => {
      if (settled) return;
      settled = true;
      yes.removeEventListener("click", onYes);
      fill.removeEventListener("click", onFill);
      no.removeEventListener("click", onNo);
      close.removeEventListener("click", onNo);
      dlg.removeEventListener("cancel", onCancel);
      dlg.removeEventListener("close", onClose);
      if (dlg.open) dlg.close();
      resolve(value);
    };
    const onYes = () => finish(true);
    const onFill = () => {
      const result = {};
      if (needTerm) {
        result.year_start = document.getElementById("importYearStart").value.trim();
        result.year_end = document.getElementById("importYearEnd").value.trim();
        result.semester = document.getElementById("importSemester").value.trim();
        if (!result.year_start || !result.year_end || !result.semester) {
          document.getElementById("importPromptText").textContent = "请填写学年起、学年止和学期。";
          return;
        }
      }
      if (needClass) {
        result.class_name = select.value;
        if (!result.class_name) {
          document.getElementById("importPromptText").textContent = "请选择一个班级。";
          return;
        }
      }
      finish(result);
    };
    const onNo = () => finish(null);
    const onCancel = (event) => {
      event.preventDefault();
      finish(null);
    };
    const onClose = () => finish(null);
    yes.addEventListener("click", onYes);
    fill.addEventListener("click", onFill);
    no.addEventListener("click", onNo);
    close.addEventListener("click", onNo);
    dlg.addEventListener("cancel", onCancel);
    dlg.addEventListener("close", onClose);
    dlg.showModal();
  });
}

async function importRegister(file, mode, extra) {
  if (!courseId) {
    showResult("请先新建并选择课程文件夹", true);
    return;
  }
  try {
    await flushSave();
  } catch (err) {
    return;
  }
  lastRegisterFile = file;
  const options = { ...(extra || {}), mode: mode || "auto" };
  showResult("正在导入成绩登记表…", false);
  try {
    let response = await postRegister(file, options);
    if (response.status === 409) {
      let data = {};
      try {
        data = await response.json();
      } catch (err) {
        data = {};
      }
      showResult("", false);
      if (data.code === "need_input") {
        const filled = await askImportPrompt(data);
        if (!filled) {
          showResult("已取消导入。", false);
          return;
        }
        await importRegister(file, mode, { ...options, ...filled });
        return;
      }
      if (data.code === "replace") {
        const agreed = await askImportPrompt(data);
        if (!agreed) {
          showResult("已取消导入。", false);
          return;
        }
        await importRegister(file, mode, { ...options, confirm: true });
        return;
      }
      const agreed = await askConfirm(typeof data.detail === "string" ? data.detail : "确定要导入吗？");
      if (!agreed) {
        showResult("已取消导入。", false);
        return;
      }
      await importRegister(file, mode, { ...options, confirm: true });
      return;
    }
    if (!response.ok) {
      showResult(await errorMessage(response), true);
      return;
    }
    const data = await response.json();
    markCurrentResultStale();
    await openCourse(courseId);
    importedHeadcount = Number(data.student_count) || 0;
    updateSummaries();
    showRegisterBanner(data);
    showRegisterResult(data);
    showResult(registerSummary(data), false);
  } catch (err) {
    if (err.message !== "未登录") showResult(err.message || "导入失败", true);
  }
}

document.getElementById("registerFile").addEventListener("change", async () => {
  const input = document.getElementById("registerFile");
  const file = input.files && input.files[0];
  input.value = "";
  if (!file) return;
  await importRegister(file, "auto");
});

document.getElementById("reimportForward").addEventListener("click", () => {
  if (!lastRegisterFile) return showResult("请重新选择成绩登记表。", true);
  importRegister(lastRegisterFile, "forward");
});

document.getElementById("reimportReverse").addEventListener("click", () => {
  if (!lastRegisterFile) return showResult("请重新选择成绩登记表。", true);
  importRegister(lastRegisterFile, "reverse");
});

state.grid = defaultRows();
bindOnce();
bindDialogs();
renderAll();
afterCourseLoaded();
loadCourses();
loadAiStatus();
loadMe();
