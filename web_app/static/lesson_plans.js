(() => {
  "use strict";
  const byId = (id) => document.getElementById(id);
  const state = { plan: null, index: 0, source: null, course: null, watch: 0, busy: false, meta: {}, teacherChoice: null, assignmentRequest: 0, assignmentLoading: false, metadataEdit: null, manualMode: false, teacherLookup: 0 };
  const fields = ["课程名称", "学年起", "学年止", "学期", "教研室", "授课对象", "授课教师"];
  const labels = { "教学目标": "教学目标或要求", "教学内容": "教学内容", "教学手段与方法": "教学手段与方法", "思考讨论作业": "思考题、讨论题、作业" };
  const alert = (message = "") => { byId("lpError").textContent = message; byId("lpError").hidden = !message; };
  const paragraphs = (text) => text.split(/\r?\n/).filter((line) => line.trim());
  const characterCount = (text) => Array.from(text.replace(/\s/g, "")).length;
  const jsonOptions = (body) => ({ method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  async function request(path, options) {
    const response = await api(path, options);
    if (!response.ok) throw new Error(await errorMessage(response));
    return response.json();
  }
  function inputField(label, value, change, multiline = false) {
    const wrapper = document.createElement("label"); wrapper.className = "field";
    const title = document.createElement("span"); title.textContent = label;
    const input = document.createElement(multiline ? "textarea" : "input");
    input.value = value; input.addEventListener("input", () => { change(input.value); hideDownload(); });
    wrapper.append(title, input); return { wrapper, input };
  }
  function hideDownload() { byId("lpDownload").hidden = true; byId("lpDownloadHint").hidden = true; }
  function controls(busy) {
    state.busy = busy;
    ["lpReadOutline", "lpGenerate", "lpSave", "lpCheck", "lpExport", "lpCompress", "lpUseManual", "lpUseTimetable"].forEach((id) => {
      const needsChoice = state.assignmentLoading || (state.plan && (state.plan.teaching_assignment?.status === "pending" || (state.plan.teaching_assignment?.status === "selection_required" && !state.manualMode)));
      byId(id).disabled = busy || (id !== "lpReadOutline" && !state.plan) || sessionStopped || (needsChoice && ["lpGenerate", "lpCheck", "lpExport", "lpCompress"].includes(id));
    });
    byId("lpEditor").querySelectorAll("input, textarea, select").forEach((input) => { input.disabled = busy; });
    byId("lpOutlineFile").disabled = busy;
  }
  function counter(node, key, value) {
    const [low, high] = state.plan.limits[key]; const count = characterCount(value);
    node.textContent = `${count} 字 · 建议 ${low}–${high} 字`;
    node.className = "lp-counter" + (count < low || count > high ? " bad" : "");
  }
  function lessonView() {
    if (!state.plan) return;
    const plan = state.plan, row = plan.context.lessons[state.index];
    byId("lpLessonSelect").value = String(state.index);
    byId("lpPrevious").disabled = state.index === 0;
    byId("lpNext").disabled = state.index === plan.context.lessons.length - 1;
    const container = byId("lpLessonFields"); container.replaceChildren();
    Object.keys(labels).forEach((key) => {
      const number = document.createElement("span");
      const editor = inputField(labels[key], row[key].join("\n"), (value) => {
        row[key] = paragraphs(value); counter(number, key, value); markPagesOld();
      }, true);
      editor.input.disabled = state.busy;
      counter(number, key, editor.input.value); editor.wrapper.append(number); container.append(editor.wrapper);
    });
    pageStatus();
  }
  function markPagesOld() { state.plan.pages_current = false; pageStatus(); }
  function pageStatus() {
    const plan = state.plan, row = plan.context.lessons[state.index];
    const pages = plan.pages_current ? plan.pages[String(row["课次"])] : null;
    const status = byId("lpPageStatus");
    status.textContent = pages ? `参考页数检查：第 ${row["课次"]} 课次 ${pages} 页${pages > 1 ? "，可压缩或手动调整后重新检查。" : "。"}` : "页数尚未检查或内容已修改，请重新检查；服务器字体下的页数仅供参考。";
    status.classList.toggle("over", Boolean(pages && pages > 1));
  }
  const metadataFields = { "授课教师": "teacher", "授课对象": "audience", "教研室": "teaching_unit" };
  function manualEdit(key, value) {
    if (!state.plan || !metadataFields[key]) return;
    state.metadataEdit ||= { mode: "manual", revision: state.plan.teaching_assignment?.override_revision || 0, fields: {} };
    state.metadataEdit.fields[metadataFields[key]] = value;
    state.manualMode = true;
  }
  function metadataControls() {
    const assignment = state.plan?.teaching_assignment;
    Object.entries(metadataFields).forEach(([key, field]) => {
      if (state.meta[key]) state.meta[key].readOnly = !(state.manualMode || assignment?.editable?.[field]);
    });
    byId("lpUseManual").hidden = state.manualMode || Object.keys(metadataFields).every((key) => assignment?.editable?.[metadataFields[key]]);
    byId("lpUseTimetable").hidden = !assignment?.has_manual;
  }
  async function teacherHints(value) {
    const token = ++state.teacherLookup;
    try {
      const result = await request("/api/teaching/teachers?q=" + encodeURIComponent(value.trim()) + "&college=" + encodeURIComponent(state.plan.teaching_assignment?.college || ""));
      if (token !== state.teacherLookup || state.plan?.context["授课教师"] !== value) return;
      const list = byId("lpTeacherNames"); list.replaceChildren();
      result.teachers.forEach((item) => {
        const option = document.createElement("option"); option.value = item.name;
        option.label = [item.college, item.teaching_unit].filter(Boolean).join(" · "); list.append(option);
      });
      const matches = result.teachers.filter((item) => item.name === value.trim());
      const unitEdited = Object.hasOwn(state.metadataEdit?.fields || {}, "teaching_unit") || state.plan.teaching_assignment?.manual_fields?.includes("teaching_unit");
      if (matches.length === 1 && !unitEdited) {
        state.plan.context["教研室"] = matches[0].teaching_unit;
        state.meta["教研室"].value = matches[0].teaching_unit;
      }
    } catch (_) { /* Free-text entry remains available if the directory cannot be queried. */ }
  }
  byId("lpUseManual").addEventListener("click", () => {
    state.manualMode = true; metadataControls(); controls(false);
    byId("lpTeachingHint").textContent = "可以手动填写；保存后只用于当前课程、当前学期的教学日历和教案。";
  });
  byId("lpUseTimetable").addEventListener("click", async () => {
    if (!state.plan || state.busy) return;
    state.metadataEdit = { mode: "timetable", revision: state.plan.teaching_assignment?.override_revision || 0 };
    alert(); controls(true);
    try { await save(); } catch (err) { alert(err.message); }
    finally { controls(false); }
  });
  function teachingView(assignment) {
    if (!assignment) return;
    state.plan.teaching_assignment = assignment;
    state.plan.context["授课教师"] = assignment.teacher;
    state.plan.context["授课对象"] = assignment.classes.join("、");
    state.plan.context["教研室"] = assignment.teacher_info?.teaching_unit || "";
    ["授课教师", "授课对象", "教研室"].forEach((key) => { if (state.meta[key]) state.meta[key].value = state.plan.context[key]; });
    metadataControls();
    const info = assignment.teacher_info;
    const unitHint = !assignment.teacher ? "" : info?.unit_source === "manual" ? " 教研室使用已保存的手填信息。" : info?.status === "ambiguous" ? " 教师库有同名记录，可手动填写教研室。" : !info?.teaching_unit ? " 教研室尚未维护，可手动填写。" : info.unit_source === "department" ? ` 教研室按所属系填写：${info.teaching_unit}。` : " 教研室已从教师信息自动填入。";
    const optionalHint = ["missing", "review_required", "manual"].includes(assignment.status) ? " 封面这三项可留空，下载后在 Word 中补填。" : "";
    byId("lpTeachingHint").textContent = assignment.message + unitHint + optionalHint;
    const courseWrap = byId("lpCourseMatchWrap"), courseSelect = byId("lpCourseMatch");
    const courseMatches = assignment.course_matches || [];
    const matchedCourse = assignment.matched_course_name || "";
    courseWrap.hidden = !courseMatches.length && !matchedCourse;
    courseSelect.replaceChildren();
    if (!courseWrap.hidden) {
      const empty = document.createElement("option"); empty.value = "";
      empty.textContent = matchedCourse ? "取消课程对应，按原课程名查找" : "请选择课表中的对应课程";
      courseSelect.append(empty);
      [...new Set([...courseMatches, ...(matchedCourse ? [matchedCourse] : [])])].forEach((name) => {
        const option = document.createElement("option"); option.value = name; option.textContent = name;
        courseSelect.append(option);
      });
      courseSelect.value = matchedCourse;
    }
    const wrap = byId("lpTeacherChoiceWrap"), select = byId("lpTeacherChoice");
    wrap.hidden = assignment.options.length <= 1 || Boolean(assignment.manual_fields?.includes("teacher"));
    select.replaceChildren();
    if (!wrap.hidden) {
      const empty = document.createElement("option"); empty.value = ""; empty.textContent = "请选择本学期的任课教师"; select.append(empty);
      assignment.options.forEach((item) => {
        const option = document.createElement("option"); option.value = item.teacher;
        option.textContent = `${item.teacher}（${item.classes.join("、")}）`; select.append(option);
      });
      select.value = assignment.teacher;
    }
  }
  function pendingTeaching() {
    state.assignmentRequest++; state.assignmentLoading = false;
    state.teacherChoice = null; state.metadataEdit = null; state.manualMode = false; state.teacherLookup++;
    teachingView({ status: "pending", teacher: "", classes: [], options: [], message: "请完整填写学年、学期，系统将重新匹配后台课表。" });
    controls(state.busy);
  }
  async function refreshTeaching() {
    if (!state.plan || state.busy) return;
    const identity = state.plan.id, token = ++state.assignmentRequest;
    const context = state.plan.context;
    if (!/^\d{4}$/.test(context["学年起"]) || !/^\d{4}$/.test(context["学年止"]) || Number(context["学年止"]) !== Number(context["学年起"]) + 1 || !["1", "2"].includes(context["学期"])) {
      pendingTeaching(); return;
    }
    state.assignmentLoading = true; controls(false);
    try {
      const query = new URLSearchParams({ year_start: context["学年起"], year_end: context["学年止"], semester: context["学期"] });
      const result = await request(`/api/lesson-plans/${identity}/teaching-assignment?${query}`);
      if (token !== state.assignmentRequest || identity !== state.plan?.id) return;
      state.teacherChoice = null; teachingView(result); hideDownload(); markPagesOld(); alert();
    } catch (err) {
      if (token !== state.assignmentRequest || identity !== state.plan?.id) return;
      pendingTeaching(); alert(err.message || "课表查询失败，请重新填写学年、学期后重试。");
    } finally {
      if (token === state.assignmentRequest) { state.assignmentLoading = false; controls(false); }
    }
  }
  byId("lpCourseMatch").addEventListener("change", async (event) => {
    if (!state.plan || state.busy) return;
    state.teacherChoice = null;
    state.metadataEdit = { mode: "course_match",
      revision: state.plan.teaching_assignment?.override_revision || 0,
      course_name: event.target.value };
    alert(); controls(true);
    try { await save(); } catch (err) { alert(err.message || "课程对应关系保存失败，请重试。"); }
    finally { controls(false); }
  });
  byId("lpTeacherChoice").addEventListener("change", (event) => {
    const teacher = event.target.value;
    const assignment = state.plan?.teaching_assignment;
    const item = assignment?.options.find((option) => option.teacher === teacher);
    state.teacherChoice = item ? teacher : null;
    teachingView({ ...assignment, teacher: item?.teacher || "", classes: item?.classes || [], teacher_info: item?.teacher_info, status: item ? "ready" : "selection_required", message: item ? "已选择课表中的任课教师；保存后仅在本课程、本学期沿用。新学期需重新确认。" : "请选择本学期的任课教师。" });
    hideDownload(); markPagesOld(); controls(false);
  });
  function render(plan) {
    state.plan = plan; state.index = Math.min(state.index, plan.context.lessons.length - 1);
    byId("lpEditor").hidden = false;
    state.teacherChoice = null; state.meta = {}; state.assignmentRequest++; state.assignmentLoading = false;
    state.metadataEdit = null; state.manualMode = Boolean(plan.teaching_assignment?.has_manual); state.teacherLookup++;
    const meta = byId("lpMetaFields"); meta.replaceChildren();
    fields.forEach((key) => {
      const periodField = ["学年起", "学年止", "学期"].includes(key);
      const editor = inputField(key, plan.context[key], (value) => { plan.context[key] = value; manualEdit(key, value); markPagesOld(); if (periodField) pendingTeaching(); });
      if (key === "授课教师") {
        editor.input.setAttribute?.("list", "lpTeacherNames");
        editor.input.addEventListener("input", () => teacherHints(editor.input.value));
      }
      if (key === "学年起" || key === "学年止") editor.input.inputMode = "numeric";
      editor.input.readOnly = ["授课教师", "授课对象", "教研室"].includes(key);
      if (periodField) editor.input.addEventListener("change", refreshTeaching);
      state.meta[key] = editor.input; meta.append(editor.wrapper);
    });
    byId("lpTeacherChoiceWrap").hidden = true; byId("lpTeachingHint").textContent = "";
    teachingView(plan.teaching_assignment);
    const select = byId("lpLessonSelect"); select.replaceChildren();
    plan.context.lessons.forEach((row, index) => {
      const option = document.createElement("option"); option.value = String(index);
      option.textContent = `第 ${row["课次"]} 课次 · ${row["授课题目"]}（${row["学时"]}学时）`; select.append(option);
    });
    byId("lpMaterials").value = plan.context["学习资料"].join("\n");
    byId("lpWebsites").value = plan.context["学习网站"].join("\n");
    const evaluation = byId("lpEvaluation"); evaluation.replaceChildren();
    Object.keys(plan.context["评价"]).forEach((key) => {
      evaluation.append(inputField(key, plan.context["评价"][key], (value) => { plan.context["评价"][key] = value; markPagesOld(); }, true).wrapper);
    });
    byId("lpWarnings").textContent = plan.warnings.join("\n");
    const download = byId("lpDownload"); download.hidden = !plan.download_url;
    if (plan.download_url) { download.href = plan.download_url; download.download = `${plan.context["课程名称"]}_课程教案.docx`; }
    byId("lpDownloadHint").hidden = !plan.download_url;
    lessonView(); controls(Boolean(plan.job));
  }
  async function save() {
    const response = await api(`/api/lesson-plans/${state.plan.id}`, {
      method: "PATCH", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ revision: state.plan.revision, context: state.plan.context, ...(state.teacherChoice ? { teacher_choice: state.teacherChoice } : {}), ...(state.metadataEdit ? { teaching_metadata: state.metadataEdit } : {}) }),
    });
    if (!response.ok) throw new Error(await errorMessage(response));
    render(await response.json()); return state.plan;
  }
  async function watch(job) {
    const token = ++state.watch, identity = state.plan.id, started = Date.now();
    byId("lpTaskPanel").hidden = false; controls(true);
    while (token === state.watch) {
      let current;
      try { current = await request(`/api/lesson-plan-jobs/${job.job_id}`); }
      catch (err) {
        if (token !== state.watch || !state.plan || identity !== state.plan.id) return;
        if (err.message === "未登录") { alert("登录已失效，请重新登录。任务结果已保存。"); controls(false); return; }
        byId("lpTaskStage").textContent = "暂时无法连接，正在重新查询进度…";
        await delay(2500); continue;
      }
      if (token !== state.watch || !state.plan || identity !== state.plan.id) return;
      byId("lpTaskStage").textContent = current.stage_label;
      byId("lpProgress").value = current.percent || 0;
      const completed = current.action === "generate" || current.action === "compress" ? `已处理 ${current.completed || 0}/${current.total || state.plan.context.lessons.length} 课次 · ` : "";
      byId("lpTaskCounts").textContent = `${completed}已请求 ${current.api_calls || 0} 次 API · 本次查看已等待 ${Math.floor((Date.now() - started) / 1000)} 秒`;
      if (current.status === "fail") {
        try { render(await request(`/api/lesson-plans/${identity}`)); } catch (_) { /* 进度错误仍可显示。 */ }
        alert(current.error || "生成失败，请重试。已完成课次已保存。"); controls(false); return;
      }
      if (current.done) {
        render(await request(`/api/lesson-plans/${identity}`)); controls(false);
        if (current.action === "export" && state.plan.download_url) {
          try { byId("lpDownload").click(); } catch (_) { /* 手动下载链接继续可用。 */ }
        }
        return;
      }
      await delay(1500);
    }
  }
  async function start(action) {
    if (!state.plan || state.busy || state.assignmentLoading) return;
    if (state.plan.teaching_assignment?.status === "pending" || (state.plan.teaching_assignment?.status === "selection_required" && !state.manualMode)) { alert("请先确定本学期的任课教师和授课班级。"); return; }
    alert(); hideDownload(); controls(true);
    try {
      await save(); controls(true);
      const number = state.plan.context.lessons[state.index]["课次"];
      const suffix = action === "compress" ? `lessons/${number}/compress` : action === "check" ? "check-pages" : action;
      const job = await request(`/api/lesson-plans/${state.plan.id}/${suffix}`, jsonOptions({ revision: state.plan.revision }));
      await watch(job);
    } catch (err) { alert(err.message || "教案处理失败。"); controls(false); }
  }
  async function open() {
    const opening = ++state.watch;
    state.assignmentRequest++; state.assignmentLoading = false; state.teacherChoice = null;
    state.course = courseId || null; state.plan = null; state.source = null; state.index = 0;
    alert(); byId("lpEditor").hidden = true; byId("lpTaskPanel").hidden = true; hideDownload(); controls(true);
    byId("lpSourcePanel").hidden = true; byId("lpLoading").hidden = false; byId("lpOutlineFile").value = "";
    openDialog("dlgLessonPlan");
    try {
      let plan = null;
      if (state.course) {
        const [saved, latest] = await Promise.all([
          request(`/api/courses/${state.course}/lesson-plan-outline`), request(`/api/courses/${state.course}/lesson-plans/current`),
        ]);
        if (opening !== state.watch) return;
        state.source = saved.source;
        if (latest.plan && (!saved.source || (latest.plan.source_id === saved.source.id && (latest.plan.outline_version || "") === (saved.source.outline_version || "")) || latest.plan.job)) {
          plan = latest.plan;
        } else if (saved.source) {
          plan = await request("/api/lesson-plans", jsonOptions({ source_id: saved.source.id, reuse: true }));
          if (opening !== state.watch) return;
        }
        if (!saved.source) {
          byId("lpSourceHint").textContent = saved.message;
          byId("lpSourcePanel").hidden = false;
        }
      } else {
        byId("lpSourceHint").textContent = "可以直接上传生成教案；选择课程后，原大纲会与课程关联。";
        byId("lpSourcePanel").hidden = false;
      }
      byId("lpLoading").hidden = true;
      if (plan) render(plan);
      controls(Boolean(plan && plan.job));
      if (plan && plan.job) watch(plan.job);
    } catch (err) {
      if (opening !== state.watch) return;
      byId("lpLoading").hidden = true;
      alert(err.message || "无法读取教案资料。"); controls(false);
    }
  }
  byId("lessonPlanBtn").addEventListener("click", open);
  byId("lpReadOutline").addEventListener("click", async () => {
    if (state.busy) return;
    const opening = ++state.watch;
    alert(); controls(true);
    try {
      const file = byId("lpOutlineFile").files[0];
      if (!file || !file.name.toLowerCase().endsWith(".docx")) throw new Error("请选择 .docx 教学大纲。");
      if (file.size > 5 * 1024 * 1024) throw new Error("教学大纲不能超过 5MB。");
      const body = new FormData(); body.append("outline", file); if (state.course) body.append("course_id", state.course);
      const plan = await request("/api/lesson-plans", { method: "POST", body });
      if (opening !== state.watch) return;
      state.index = 0; render(plan); byId("lpTaskPanel").hidden = true; byId("lpSourcePanel").hidden = true;
    } catch (err) { if (opening === state.watch) alert(err.message || "大纲读取失败。"); }
    finally { if (opening === state.watch) controls(false); }
  });
  byId("lpGenerate").addEventListener("click", () => start("generate"));
  byId("lpCompress").addEventListener("click", () => start("compress"));
  byId("lpCheck").addEventListener("click", () => start("check"));
  byId("lpExport").addEventListener("click", () => start("export"));
  byId("lpSave").addEventListener("click", async () => {
    alert(); controls(true); try { await save(); byId("lpWarnings").textContent = "修改已保存。"; } catch (err) { alert(err.message); } finally { controls(false); }
  });
  byId("lpLessonSelect").addEventListener("change", (event) => { state.index = Number(event.target.value); lessonView(); });
  byId("lpPrevious").addEventListener("click", () => { if (state.index > 0) { state.index--; lessonView(); } });
  byId("lpNext").addEventListener("click", () => { if (state.plan && state.index < state.plan.context.lessons.length - 1) { state.index++; lessonView(); } });
  byId("lpMaterials").addEventListener("input", (event) => { if (state.plan) { state.plan.context["学习资料"] = paragraphs(event.target.value); hideDownload(); markPagesOld(); } });
  byId("lpWebsites").addEventListener("input", (event) => { if (state.plan) { state.plan.context["学习网站"] = paragraphs(event.target.value); hideDownload(); markPagesOld(); } });
  window.saveLessonPlanOutline = async (identity, file) => {
    if (file.size > 5 * 1024 * 1024) throw new Error("原始 Word 大纲超过 5MB。");
    const body = new FormData(); body.append("outline", file);
    return request(`/api/courses/${identity}/lesson-plan-outline`, { method: "POST", body });
  };
})();
