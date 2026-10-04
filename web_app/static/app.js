const STORAGE_KEY = "calculatorpro.web.v1";

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
    links: [
      { name: "平时考核", ratio: 0.3, methods: [{ name: "平时作业", weights: [50, 50] }] },
      { name: "期中考核", ratio: 0, methods: [{ name: "期中测验", weights: [50, 50] }] },
      { name: "期末考核", ratio: 0.7, methods: [{ name: "期末考试", weights: [40, 60] }] },
    ],
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

function loadState() {
  const base = defaultState();
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return base;
    const saved = JSON.parse(raw);
    return {
      ...base,
      ...saved,
      courseOpen: { ...base.courseOpen, ...(saved.courseOpen || {}) },
      courseBasic: { ...base.courseBasic, ...(saved.courseBasic || {}) },
      links: Array.isArray(saved.links) && saved.links.length === 3 ? saved.links : base.links,
    };
  } catch (err) {
    return base;
  }
}

let state = loadState();
state.objectivesCount = clampCount(state.objectivesCount);

function clampCount(value) {
  const number = Number(value) || 2;
  return Math.max(1, Math.min(8, Math.round(number)));
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
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

function renderScalarFields(container, spec, bag, onChange) {
  container.innerHTML = "";
  spec.forEach(([key, label]) => {
    const input = el("input", { type: "text", value: bag[key] || "" });
    input.addEventListener("input", () => {
      bag[key] = input.value;
      if (key === "course_name" && bag === state.courseOpen && !state.courseBasic.course_name) {
        const basic = document.querySelector('[data-basic="course_name"]');
        if (basic && !basic.dataset.touched) basic.value = input.value;
      }
      onChange();
    });
    if (bag === state.courseBasic) input.dataset.basic = key;
    const field = el("label", { class: "field", text: label }, [input]);
    container.append(field);
  });
}

function resizeWeights(weights) {
  const next = (weights || []).slice(0, state.objectivesCount);
  while (next.length < state.objectivesCount) next.push(0);
  return next;
}

function renderLinks() {
  const root = document.getElementById("links");
  root.innerHTML = "";
  state.links.forEach((link, linkIndex) => {
    const ratio = el("input", { type: "number", min: "0", max: "1", step: "0.01", value: String(link.ratio) });
    ratio.addEventListener("input", () => {
      link.ratio = Number(ratio.value);
      updateRatioSum();
      saveState();
      renderNoiseTargets();
    });
    const head = el("div", { class: "link-head" }, [
      el("h3", { text: link.name }),
      el("label", { class: "field", text: "环节占比" }, [ratio]),
    ]);
    const card = el("div", { class: "link-card" }, [head]);
    link.methods.forEach((method, methodIndex) => {
      method.weights = resizeWeights(method.weights);
      const name = el("input", { type: "text", value: method.name || "" });
      name.addEventListener("input", () => {
        method.name = name.value;
        saveState();
        renderNoiseTargets();
      });
      const row = el("div", { class: "method" });
      row.style.setProperty("--objs", String(state.objectivesCount));
      row.append(el("label", { class: "field", text: "考核方式" }, [name]));
      method.weights.forEach((weight, objIndex) => {
        const input = el("input", { type: "number", min: "0", max: "100", step: "1", value: String(weight) });
        input.addEventListener("input", () => {
          method.weights[objIndex] = Number(input.value);
          saveState();
        });
        row.append(el("label", { class: "field", text: `目标${objIndex + 1}` }, [input]));
      });
      if (link.methods.length > 1) {
        const remove = el("button", { class: "ghost", type: "button", text: "删除" });
        remove.addEventListener("click", () => {
          link.methods.splice(methodIndex, 1);
          saveState();
          renderLinks();
        });
        row.append(remove);
      }
      card.append(row);
    });
    const add = el("button", { class: "secondary", type: "button", text: "添加考核方式" });
    add.style.marginTop = "8px";
    add.addEventListener("click", () => {
      link.methods.push({ name: "", weights: Array(state.objectivesCount).fill(0) });
      saveState();
      renderLinks();
    });
    card.append(add);
    root.append(card);
  });
  updateRatioSum();
  renderNoiseTargets();
}

function updateRatioSum() {
  const sum = state.links.reduce((total, link) => total + (Number(link.ratio) || 0), 0);
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
    requirement.addEventListener("input", () => { state.gradReq[i].requirement = requirement.value; saveState(); });
    indicator.addEventListener("input", () => { state.gradReq[i].indicator = indicator.value; saveState(); });
    const row = el("div", { class: "fields" }, [
      el("p", { class: "span-2", text: `课程目标${i + 1}` }),
      el("label", { class: "field", text: "支撑的毕业要求" }, [requirement]),
      el("label", { class: "field", text: "支撑的毕业要求指标点" }, [indicator]),
    ]);
    root.append(row);
  }
  const reqRoot = document.getElementById("requirements");
  reqRoot.innerHTML = "";
  for (let i = 0; i < state.objectivesCount; i += 1) {
    const area = el("textarea");
    area.value = state.objectiveRequirements[i] || "";
    area.addEventListener("input", () => { state.objectiveRequirements[i] = area.value; saveState(); });
    reqRoot.append(el("label", { text: `课程目标${i + 1}要求` }, [area]));
  }
}

function renderNoiseTargets() {
  const root = document.getElementById("noiseTargets");
  root.innerHTML = "";
  const names = [];
  state.links.forEach((link) => {
    if ((Number(link.ratio) || 0) <= 0) return;
    link.methods.forEach((method) => {
      const name = (method.name || "").trim();
      if (name && !names.includes(name)) names.push(name);
    });
  });
  if (!state.noiseAllowed.length) state.noiseAllowed = names.slice();
  names.forEach((name) => {
    const box = el("input", { type: "checkbox" });
    box.checked = state.noiseAllowed.includes(name);
    box.addEventListener("change", () => {
      if (box.checked) state.noiseAllowed.push(name);
      else state.noiseAllowed = state.noiseAllowed.filter((item) => item !== name);
      saveState();
    });
    root.append(el("label", { text: name }, [box]));
  });
}

function buildRelationPayload() {
  const links = [];
  state.links.forEach((link) => {
    const ratio = Number(link.ratio) || 0;
    if (ratio <= 0) return;
    const methods = link.methods.map((method) => {
      const supports = {};
      let subtotal = 0;
      resizeWeights(method.weights).forEach((weight, index) => {
        const value = (Number(weight) || 0) / 100;
        supports[`课程目标${index + 1}`] = Math.round(value * 1e6) / 1e6;
        subtotal += value;
      });
      return {
        name: (method.name || "").trim(),
        supports,
        subtotal: Math.round(subtotal * 1e6) / 1e6,
      };
    });
    links.push({ name: link.name, ratio, methods });
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

function collectSettings() {
  const basic = { ...state.courseBasic };
  if (!basic.course_name) basic.course_name = state.courseOpen.course_name || "";
  const openInfo = { ...state.courseOpen, term: state.courseOpen.semester || "1" };
  return {
    mode: state.mode,
    student_count: Number(state.studentCount) || 1,
    course_open_info: openInfo,
    course_basic_info: basic,
    ratios: {
      usual: Number(state.links[0].ratio) || 0,
      midterm: Number(state.links[1].ratio) || 0,
      final: Number(state.links[2].ratio) || 0,
    },
    relation_payload: buildRelationPayload(),
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

function validateBeforeRun(needsFile) {
  const sum = state.links.reduce((total, link) => total + (Number(link.ratio) || 0), 0);
  if (Math.abs(sum - 1) > 0.001) return "平时、期中、期末占比之和必须等于 1";
  const payload = buildRelationPayload();
  if (!payload.links.length) return "请至少保留一个占比大于 0 的考核环节";
  for (const link of payload.links) {
    let weightSum = 0;
    for (const method of link.methods) {
      if (!method.name) return `${link.name} 有未命名的考核方式`;
      weightSum += method.subtotal;
    }
    if (Math.abs(weightSum - 1) > 0.02) {
      return `${link.name} 的目标权重合计应为 100，当前约为 ${Math.round(weightSum * 100)}`;
    }
  }
  if (needsFile && !document.getElementById("gradeFile").files[0]) return "请先导入成绩 Excel";
  return "";
}

function setBusy(busy) {
  ["templateBtn", "calcBtn", "exportBtn", "aiBtn"].forEach((id) => {
    document.getElementById(id).disabled = busy;
  });
}

function showResult(text, isError) {
  const node = document.getElementById("result");
  node.textContent = text;
  node.classList.toggle("error", Boolean(isError));
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

async function postForm(path, withFile) {
  const body = new FormData();
  body.append("settings", JSON.stringify(collectSettings()));
  if (withFile) {
    body.append("file", document.getElementById("gradeFile").files[0]);
    const previous = document.getElementById("prevFile").files[0];
    if (previous) body.append("previous", previous);
  }
  const response = await fetch(path, { method: "POST", body });
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

function bindStatic() {
  renderScalarFields(document.getElementById("openFields"), OPEN_FIELDS, state.courseOpen, saveState);
  renderScalarFields(document.getElementById("basicFields"), BASIC_FIELDS, state.courseBasic, saveState);
  const basicName = document.querySelector('[data-basic="course_name"]');
  if (basicName) {
    basicName.addEventListener("input", () => { basicName.dataset.touched = "1"; });
  }
  document.getElementById("objCount").value = state.objectivesCount;
  document.getElementById("objCount").addEventListener("change", () => {
    state.objectivesCount = clampCount(document.getElementById("objCount").value);
    document.getElementById("objCount").value = state.objectivesCount;
    state.links.forEach((link) => link.methods.forEach((method) => {
      method.weights = resizeWeights(method.weights);
    }));
    saveState();
    renderLinks();
    renderGrad();
  });
  document.getElementById("description").value = state.courseDescription;
  document.getElementById("description").addEventListener("input", (event) => {
    state.courseDescription = event.target.value;
    saveState();
  });
  document.getElementById("studentCount").value = state.studentCount;
  document.getElementById("studentCount").addEventListener("input", (event) => {
    state.studentCount = Number(event.target.value);
    saveState();
  });
  fillSelect(document.getElementById("spreadMode"), SPREAD_OPTIONS, state.spreadMode);
  fillSelect(document.getElementById("distribution"), DIST_OPTIONS, state.distribution);
  fillSelect(document.getElementById("reportStyle"), STYLE_OPTIONS, state.reportStyle);
  document.getElementById("spreadMode").addEventListener("change", (event) => { state.spreadMode = event.target.value; saveState(); });
  document.getElementById("distribution").addEventListener("change", (event) => { state.distribution = event.target.value; saveState(); });
  document.getElementById("reportStyle").addEventListener("change", (event) => { state.reportStyle = event.target.value; saveState(); });
  document.getElementById("wordLimit").value = state.wordLimit;
  document.getElementById("wordLimit").addEventListener("input", (event) => { state.wordLimit = Number(event.target.value); saveState(); });
  document.getElementById("noiseEnabled").checked = state.noiseEnabled;
  document.getElementById("noiseRatio").value = state.noiseRatio;
  document.getElementById("severityMode").value = state.severityMode;
  document.getElementById("noiseEnabled").addEventListener("change", (event) => { state.noiseEnabled = event.target.checked; saveState(); });
  document.getElementById("noiseRatio").addEventListener("input", (event) => { state.noiseRatio = Number(event.target.value); saveState(); });
  document.getElementById("severityMode").addEventListener("change", (event) => { state.severityMode = event.target.value; saveState(); });
  document.getElementById("modeForward").addEventListener("click", () => { state.mode = "forward"; saveState(); syncMode(); });
  document.getElementById("modeReverse").addEventListener("click", () => { state.mode = "reverse"; saveState(); syncMode(); });
  document.getElementById("gradeFile").addEventListener("change", () => {
    const file = document.getElementById("gradeFile").files[0];
    document.getElementById("fileName").textContent = file ? file.name : "尚未选择文件";
  });
  document.getElementById("prevFile").addEventListener("change", () => {
    const file = document.getElementById("prevFile").files[0];
    document.getElementById("prevName").textContent = file ? file.name : "不上传则按 0 对比";
  });
  renderLinks();
  renderGrad();
  syncMode();
}

async function loadAiStatus() {
  const banner = document.getElementById("aiBanner");
  try {
    const response = await fetch("/api/ai-status");
    if (response.status === 401) {
      location.href = "/login";
      return;
    }
    const data = await response.json();
    banner.textContent = data.message || "";
    banner.classList.toggle("ok", Boolean(data.enabled));
  } catch (err) {
    banner.textContent = "无法确认 AI 报告状态";
  }
}

document.getElementById("templateBtn").addEventListener("click", async () => {
  const problem = validateBeforeRun(false);
  if (problem) return showResult(problem, true);
  setBusy(true);
  showResult("正在生成模板…", false);
  try {
    const response = await postForm("/api/template", false);
    if (!response.ok) return showResult(await errorMessage(response), true);
    const blob = await response.blob();
    const name = filenameFromDisposition(response.headers.get("Content-Disposition"), "成绩模板.xlsx");
    downloadBlob(blob, name);
    showResult(`模板已下载：${name}`, false);
  } catch (err) {
    if (err.message !== "未登录") showResult("模板下载失败", true);
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
    const response = await postForm("/api/calculate", true);
    if (!response.ok) return showResult(await errorMessage(response), true);
    const data = await response.json();
    const lines = [
      `${data.course_name} · ${data.mode === "reverse" ? "逆向" : "正向"}`,
      `学生 ${data.student_count} 人，平均分 ${data.average_score}`,
    ];
    Object.entries(data.achievement || {}).forEach(([key, value]) => {
      lines.push(`${key}：${value}`);
    });
    if (data.files && data.files.length) lines.push(`将导出：${data.files.join("、")}`);
    showResult(lines.join("\n"), false);
  } catch (err) {
    if (err.message !== "未登录") showResult("计算失败", true);
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
    const response = await postForm(path, true);
    if (!response.ok) return showResult(await errorMessage(response), true);
    const blob = await response.blob();
    const name = filenameFromDisposition(response.headers.get("Content-Disposition"), fallbackName);
    downloadBlob(blob, name);
    showResult(`已下载：${name}`, false);
  } catch (err) {
    if (err.message !== "未登录") showResult("下载失败", true);
  } finally {
    setBusy(false);
  }
}

document.getElementById("exportBtn").addEventListener("click", () => {
  downloadZip("/api/export", "统计表.zip", "正在导出统计表…");
});

document.getElementById("aiBtn").addEventListener("click", () => {
  downloadZip("/api/ai-report", "AI分析报告.zip", "正在生成 AI 分析报告…");
});

document.getElementById("clearBtn").addEventListener("click", () => {
  localStorage.removeItem(STORAGE_KEY);
  state = defaultState();
  bindStatic();
  showResult("已清空本机保存的课程设置。", false);
});

document.getElementById("logoutBtn").addEventListener("click", async () => {
  await fetch("/api/logout", { method: "POST" });
  location.href = "/login";
});

bindStatic();
loadAiStatus();
