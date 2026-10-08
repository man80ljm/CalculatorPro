(function () {
  const page = document.body.dataset.page || "";

  function $(id) {
    return document.getElementById(id);
  }

  async function api(path, options) {
    const opts = options || {};
    const headers = { Accept: "application/json" };
    if (opts.body) headers["Content-Type"] = "application/json";
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
  if (refresh) refresh.addEventListener("click", load);

  if (page === "overview" || page === "users" || page === "jobs") {
    load();
    if (page === "overview" || page === "jobs") {
      window.setInterval(function () {
        if (!document.hidden) load();
      }, 20000);
    }
  }
})();
