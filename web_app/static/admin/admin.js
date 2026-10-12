(function () {
  const page = document.body.dataset.page || "";

  function $(id) {
    return document.getElementById(id);
  }

  async function api(path, options) {
    const opts = options || {};
    const headers = { Accept: "application/json" };
    if (opts.body) headers["Content-Type"] = "application/json";
    if (opts.csrf) headers["X-Admin-CSRF"] = opts.csrf;
    const response = await fetch(path, {
      method: opts.method || "GET",
      headers: headers,
      body: opts.body,
      credentials: "same-origin",
    });
    let data = {};
    const text = await response.text();
    if (text) {
      try {
        data = JSON.parse(text);
      } catch (err) {
        data = {};
      }
    }
    return { response: response, data: data };
  }

  function formatTime(value) {
    if (!value) return "—";
    const date = new Date(String(value).endsWith("Z") ? value : value + "Z");
    if (Number.isNaN(date.getTime())) return String(value).replace("T", " ");
    try {
      return new Intl.DateTimeFormat("zh-CN", {
        timeZone: "Asia/Shanghai",
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
        hour12: false,
      }).format(date).replace(/\//g, "-");
    } catch (err) {
      return String(value).replace("T", " ");
    }
  }

  function formatDuration(ms) {
    const value = Number(ms);
    if (!Number.isFinite(value) || value < 0) return "—";
    if (value < 1000) return value + " 毫秒";
    const seconds = value / 1000;
    if (seconds < 60) return (seconds < 10 ? seconds.toFixed(1) : String(Math.round(seconds))) + " 秒";
    const minutes = Math.floor(seconds / 60);
    const rest = Math.round(seconds % 60);
    return minutes + " 分 " + rest + " 秒";
  }

  const STATUS_LABEL = {
    queued: "排队",
    running: "进行中",
    success: "成功",
    fail: "失败",
    active: "可使用",
    disabled: "已停用",
    exhausted: "次数已用完",
    deleted: "已删除",
  };

  function cell(text) {
    const td = document.createElement("td");
    td.textContent = text == null || text === "" ? "—" : String(text);
    return td;
  }

  function statusCell(status) {
    const td = document.createElement("td");
    const span = document.createElement("span");
    span.className = "pill " + (status || "");
    span.textContent = STATUS_LABEL[status] || status || "—";
    td.appendChild(span);
    return td;
  }

  function showBanner(message) {
    const banner = $("error");
    if (!banner) return;
    if (!message) {
      banner.hidden = true;
      banner.textContent = "";
      return;
    }
    banner.hidden = false;
    banner.textContent = message;
  }

  async function loadJson(path) {
    const result = await api(path);
    if (result.response.status === 401) {
      location.href = "login";
      return null;
    }
    if (!result.response.ok) {
      showBanner((result.data && result.data.detail) || "加载失败");
      return null;
    }
    showBanner("");
    return result.data;
  }

  function metric(label, value, hint) {
    const article = document.createElement("article");
    article.className = "metric";
    const name = document.createElement("span");
    name.textContent = label;
    const number = document.createElement("b");
    number.textContent = value == null ? "—" : String(value);
    const small = document.createElement("small");
    small.textContent = hint || "";
    article.append(name, number, small);
    return article;
  }

  function fillOverview(data) {
    const host = $("cards");
    host.replaceChildren(
      metric("教师", data.teacher_count, "已注册"),
      metric("课程", data.course_count, "全部课程"),
      metric("本周活跃", data.active_teachers_week, "登录、改课或新文件"),
      metric("今日成功", data.reports_today.success, "报告"),
      metric("今日失败", data.reports_today.fail, "报告"),
      metric("本周成功", data.reports_week.success, "报告"),
      metric("本周失败", data.reports_week.fail, "报告"),
      metric("排队深度", data.queue_depth, "DeepSeek " + data.deepseek_queue + " · 报告排队 " + data.report_queued)
    );
  }

  function table(columns, rows, emptyText) {
    const wrap = document.createElement("div");
    wrap.className = "table-wrap";
    if (!rows.length) {
      const p = document.createElement("p");
      p.className = "empty";
      p.textContent = emptyText;
      wrap.appendChild(p);
      return wrap;
    }
    const el = document.createElement("table");
    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    columns.forEach(function (column) {
      const th = document.createElement("th");
      th.textContent = column.label;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    const tbody = document.createElement("tbody");
    rows.forEach(function (row) {
      const tr = document.createElement("tr");
      columns.forEach(function (column) {
        tr.appendChild(column.cell(row));
      });
      tbody.appendChild(tr);
    });
    const cards = document.createElement("div");
    cards.className = "data-cards";
    rows.forEach(function (row) {
      const card = document.createElement("article");
      card.className = "data-card";
      const title = document.createElement("h3");
      title.className = "data-card-title";
      const fields = document.createElement("dl");
      fields.className = "data-fields";
      columns.forEach(function (column, index) {
        const cardCell = column.cell(row);
        if (index === 0) {
          while (cardCell.firstChild) title.appendChild(cardCell.firstChild);
          return;
        }
        const field = document.createElement("div");
        field.className = "data-field";
        const dt = document.createElement("dt");
        dt.textContent = column.label;
        const dd = document.createElement("dd");
        while (cardCell.firstChild) dd.appendChild(cardCell.firstChild);
        field.append(dt, dd);
        fields.appendChild(field);
      });
      card.append(title, fields);
      cards.appendChild(card);
    });
    el.append(thead, tbody);
    wrap.append(el, cards);
    return wrap;
  }

  const userColumns = [
    { label: "邮箱 / 用户名", cell: function (row) { return cell(row.username); } },
    { label: "专属邀请码", cell: function (row) { return cell(row.invitation ? row.invitation.code + (row.invitation.status === "deleted" ? "（已删除）" : "") : "首次登录时生成"); } },
    { label: "邀请来源", cell: function (row) { return cell(row.invited_by || "历史账号"); } },
    { label: "注册时间", cell: function (row) { return cell(formatTime(row.created_at)); } },
    { label: "课程", cell: function (row) { return cell(row.course_count); } },
    { label: "学期", cell: function (row) { return cell(row.term_count); } },
    { label: "最近登录", cell: function (row) { return cell(formatTime(row.last_login_at)); } },
    { label: "最近报告", cell: function (row) { return cell(formatTime(row.last_report_at)); } },
  ];

  const jobColumns = [
    { label: "教师", cell: function (row) { return cell(row.username); } },
    { label: "课程", cell: function (row) { return cell(row.course_name); } },
    { label: "状态", cell: function (row) { return statusCell(row.status); } },
    { label: "耗时", cell: function (row) { return cell(formatDuration(row.duration_ms)); } },
    { label: "错误", cell: function (row) { return cell(row.error); } },
    { label: "时间", cell: function (row) { return cell(formatTime(row.created_at)); } },
  ];

  function fillJobs(data) {
    const jobs = data.jobs || [];
    const live = jobs.filter(function (row) { return row.source === "memory"; });
    const history = jobs.filter(function (row) { return row.source === "event"; });
    const host = $("jobs");
    host.replaceChildren();
    const liveTitle = document.createElement("h2");
    liveTitle.className = "section-title";
    liveTitle.textContent = "当前进程";
    const historyTitle = document.createElement("h2");
    historyTitle.className = "section-title";
    historyTitle.textContent = "历史记录";
    host.append(
      liveTitle,
      table(jobColumns, live, "内存里没有报告任务"),
      historyTitle,
      table(jobColumns, history, "还没有成功或失败的记录")
    );
  }

  let invitationRows = [];
  let invitationCsrf = "";
  let invitationBusy = false;
  let editingInvitation = null;
  let deletingInvitation = null;

  const invitationColumns = [
    { label: "邀请码", cell: function (row) { const td = cell(row.code); td.className = "invite-code"; return td; } },
    { label: "归属", cell: function (row) { return cell(row.owner_user_id === null ? "管理员新增" : row.owner); } },
    { label: "注册次数", cell: function (row) { return cell(row.used_count + " / " + (row.max_uses === null ? "不限" : row.max_uses) + (row.remaining === null ? "" : " · 剩余 " + row.remaining)); } },
    { label: "状态", cell: function (row) { return statusCell(row.status); } },
    { label: "操作", cell: function (row) {
      const td = document.createElement("td");
      const actions = document.createElement("div");
      actions.className = "actions";
      [["copy", "复制"], ["edit", "次数"], ["toggle", row.enabled ? "停用" : "启用"], ["delete", "删除"]].forEach(function (item) {
        const button = document.createElement("button");
        button.type = "button";
        button.className = item[0] === "delete" ? "danger" : "ghost";
        button.dataset.action = item[0];
        button.dataset.id = row.id;
        button.textContent = item[1];
        actions.appendChild(button);
      });
      td.appendChild(actions);
      return td;
    } },
  ];

  function invitationNotice(message) {
    $("notice").textContent = message;
    $("notice").hidden = !message;
  }

  async function loadInvitations() {
    const data = await loadJson("api/invites");
    if (!data) return false;
    invitationRows = data.invites || [];
    invitationCsrf = data.csrf_token;
    $("invites").replaceChildren(table(invitationColumns, invitationRows, "还没有邀请码，点击「新增邀请码」即可创建。"));
    return true;
  }

  async function writeInvitation(path, method, body) {
    if (!invitationCsrf) throw new Error("请刷新列表后重试。");
    const result = await api(path, { method: method, csrf: invitationCsrf, body: body === undefined ? undefined : JSON.stringify(body) });
    if (result.response.status === 401) {
      location.href = "login";
      throw new Error("登录已过期，请重新登录。");
    }
    if (!result.response.ok) throw new Error(result.data.detail || "操作失败，请重试。");
    return result.data;
  }

  function setInvitationBusy(value) {
    invitationBusy = value;
    document.querySelectorAll('#invites button, #new-invite, #refresh, #invite-dialog button, #delete-dialog button').forEach(function (button) { button.disabled = value; });
  }

  function openInvitationEditor(row) {
    editingInvitation = row || null;
    $("invite-dialog-title").textContent = row ? "调整注册次数" : "新增邀请码";
    $("invite-owner").textContent = row ? row.code + " · " + row.owner : "管理员新增";
    $("max-uses").min = row ? row.used_count : 0;
    $("max-uses").value = row && row.max_uses !== null ? row.max_uses : Math.max(5, row ? row.used_count : 0);
    $("unlimited").checked = !!row && row.max_uses === null;
    $("max-uses").disabled = $("unlimited").checked;
    $("limit-hint").textContent = row ? "已注册 " + row.used_count + " 人。总次数不能小于已使用次数；增加额度后可以继续注册。" : "6 位邀请码自动生成。填写允许注册的总人数，也可选择不限次数。";
    $("invite-error").textContent = "";
    $("save-invite").textContent = row ? "保存次数" : "生成邀请码";
    $("invite-dialog").showModal();
  }

  function bindInvitations() {
    $("new-invite").addEventListener("click", function () { openInvitationEditor(null); });
    $("unlimited").addEventListener("change", function () { $("max-uses").disabled = this.checked; });
    document.querySelectorAll("dialog [data-close]").forEach(function (button) {
      button.addEventListener("click", function () { if (!invitationBusy) button.closest("dialog").close(); });
    });
    document.querySelectorAll("dialog").forEach(function (dialog) {
      dialog.addEventListener("cancel", function (event) { if (invitationBusy) event.preventDefault(); });
    });
    $("invite-form").addEventListener("submit", async function (event) {
      event.preventDefault();
      if (invitationBusy) return;
      const maxUses = $("unlimited").checked ? null : Number($("max-uses").value);
      $("invite-error").textContent = "";
      setInvitationBusy(true);
      try {
        const result = await writeInvitation(editingInvitation ? "api/invites/" + editingInvitation.id : "api/invites", editingInvitation ? "PATCH" : "POST", { max_uses: maxUses });
        $("invite-dialog").close();
        invitationNotice(editingInvitation ? "注册次数已更新。" : "已生成邀请码 " + result.invitation.code + "，可在列表中复制。" );
        if (!await loadInvitations()) showBanner("已保存，但列表刷新失败，请点击刷新。");
      } catch (err) {
        if ($("invite-dialog").open) $("invite-error").textContent = err.message || "无法连接服务器，请重试。";
        else showBanner("已保存，请刷新列表确认。");
      } finally { setInvitationBusy(false); }
    });
    $("invites").addEventListener("click", async function (event) {
      const button = event.target.closest("button[data-action]");
      if (!button || invitationBusy) return;
      const row = invitationRows.find(function (item) { return item.id === Number(button.dataset.id); });
      if (!row) return;
      const action = button.dataset.action;
      if (action === "edit") { openInvitationEditor(row); return; }
      if (action === "delete") {
        deletingInvitation = row;
        $("delete-message").textContent = "确定删除 " + row.code + "（" + row.owner + "）吗？";
        $("delete-error").textContent = "";
        $("delete-dialog").showModal();
        return;
      }
      if (action === "copy") {
        try { await navigator.clipboard.writeText(row.code); invitationNotice("已复制邀请码 " + row.code + "。" ); }
        catch (err) { invitationNotice("请手动复制邀请码：" + row.code); }
        return;
      }
      if (action === "toggle") {
        setInvitationBusy(true);
        try {
          await writeInvitation("api/invites/" + row.id, "PATCH", { enabled: !row.enabled });
          invitationNotice(row.enabled ? "邀请码已停用。" : "邀请码已启用。" );
          if (!await loadInvitations()) showBanner("已保存，但列表刷新失败，请点击刷新。");
        } catch (err) { showBanner(err.message || "无法连接服务器，请重试。"); }
        finally { setInvitationBusy(false); }
      }
    });
    $("confirm-delete").addEventListener("click", async function () {
      if (invitationBusy || !deletingInvitation) return;
      setInvitationBusy(true);
      try {
        await writeInvitation("api/invites/" + deletingInvitation.id, "DELETE");
        $("delete-dialog").close();
        invitationNotice("邀请码已删除，已有账号和邀请记录保留。" );
        if (!await loadInvitations()) showBanner("已删除，但列表刷新失败，请点击刷新。");
      } catch (err) {
        if ($("delete-dialog").open) $("delete-error").textContent = err.message || "无法连接服务器，请重试。";
        else showBanner("已删除，请刷新列表确认。");
      } finally { setInvitationBusy(false); }
    });
  }

  async function load() {
    if (page === "overview") {
      const data = await loadJson("api/overview");
      if (data) fillOverview(data);
    } else if (page === "users") {
      const data = await loadJson("api/users");
      if (data) $("users").replaceChildren(table(userColumns, data.users || [], "还没有教师"));
    } else if (page === "jobs") {
      const data = await loadJson("api/jobs");
      if (data) fillJobs(data);
    } else if (page === "invites") {
      await loadInvitations();
    }
  }

  if (page === "login") {
    const form = $("login-form");
    form.addEventListener("submit", async function (event) {
      event.preventDefault();
      const button = form.querySelector("button");
      const error = $("error");
      error.textContent = "";
      button.disabled = true;
      const result = await api("api/login", {
        method: "POST",
        body: JSON.stringify({
          username: form.username.value,
          password: form.password.value,
        }),
      });
      button.disabled = false;
      if (result.response.ok) {
        location.href = "./";
        return;
      }
      error.textContent = (result.data && result.data.detail) || "用户名或密码不正确";
    });
    return;
  }

  const logout = $("logout");
  if (logout) {
    logout.addEventListener("click", async function () {
      await api("api/logout", { method: "POST" });
      location.href = "login";
    });
  }

  const refresh = $("refresh");
  if (refresh) refresh.addEventListener("click", function () { load().catch(function () { showBanner("无法连接服务器，请稍后刷新。"); }); });
  if (page === "invites") bindInvitations();

  if (page === "overview" || page === "users" || page === "jobs" || page === "invites") {
    load().catch(function () { showBanner("无法连接服务器，请稍后刷新。"); });
    if (page === "overview" || page === "jobs") {
      window.setInterval(function () {
        if (!document.hidden) load();
      }, 20000);
    }
  }
})();
