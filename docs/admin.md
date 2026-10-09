# 只读运维后台

后台给运维看教师数、课程数、报告成败和排队情况。不能改密码、删课、代登录、改设置或重新生成报告。

页面在本服务的 `/admin`。`/admin` 会转到 `/admin/`。目标子域是 `admin.calc.geekhuang.com`，由 nginx 反代到这个前缀，或按 Host 分流。

## 环境变量

| 变量 | 说明 |
| --- | --- |
| `ADMIN_PASSWORD` | 明文密码。只写在部署环境里，不要写进仓库、日志或页面。 |
| `ADMIN_PASSWORD_HASH` | argon2 哈希。**设置后优先于明文**，`ADMIN_PASSWORD` 会被忽略。 |
| `ADMIN_USERNAME` | 登录名，默认 `admin`。大小写不敏感。 |
| `ADMIN_HOST` | 可选。例如 `admin.calc.geekhuang.com`。设置后，Host 匹配时根路径也进后台；路径以 `/admin` 开头时，其他 Host 也能打开（方便本机）。未设置时只认 `/admin`。 |
| `COOKIE_SECURE` | `true` 时运维 Cookie 带 Secure，和教师 Cookie 一样。只在 HTTPS 后面打开。 |
| `SECRET_KEY` | 给运维 Cookie 签名。盐是 `cp-admin-session-v1`，和教师 `cp_session` 的盐 `cp-session-v2` 不同。 |

两个密码变量都空着时，登录和数据接口返回 **503**，`detail` 为「未配置管理员密码」。页面只显示配置提示，不会放行。

生成哈希（把口令换成你自己的，输出贴到环境变量，不要提交）：

```bash
python -c "from web_app.auth import hash_password; print(hash_password('replace-with-your-password'))"
```

改密码或改用户名后，旧的运维 Cookie 失效，需要重新登录。

## 登录

- 接口：`POST /admin/api/login`，JSON `{"username","password"}`。
- 成功后发 HttpOnly Cookie `cp_admin_session`，`SameSite=Lax`，`Path=/`，有效期 12 小时。
- `POST /admin/api/logout` 清除这枚 Cookie。
- 教师 `cp_session` 不能访问 `/admin/api/*`。运维 Cookie 也不能当教师登录用。
- 后台不查 `users` 表。教师的用户名和密码即使碰巧相同也不能进后台，除非你把同一口令配进了上面的环境变量。

`Path=/` 是为了子域反代：浏览器上的路径是 `/api/overview`，不是 `/admin/api/overview`。Cookie 若限定在 `/admin`，这种反代就带不上。教师站看到这枚 Cookie 会忽略。

## 只读接口

都要运维 Cookie。未登录是 401。

| 方法 | 路径 | 内容 |
| --- | --- | --- |
| GET | `/admin/api/overview` | 教师数、课程数、本周活跃教师、今日/本周报告成功和失败、排队深度 |
| GET | `/admin/api/users` | 用户名、注册时间、课程数、学期数、最近登录、最近报告 |
| GET | `/admin/api/jobs` | 数据库里的报告任务与成功/失败记录，以及兼容的内存调试记录 |

页面：`/admin/` 概览，`/admin/users`，`/admin/jobs`，`/admin/login`。静态文件在 `/admin/static/`。HTML 里用相对路径（`static/admin.css`、`api/overview`），子域反代和 `/admin` 前缀两种接法都能用。

今日、本周按北京时间（UTC+8），周一 0 点起算。库里的时间仍是 UTC。

本周活跃教师：本周内新建过 `user_sessions`，或 `courses.updated_at`、`course_files.created_at` 落在本周的用户，按 `user_id` 去重。

排队深度 = 数据库队列中等待的报告数 + DeepSeek 池中等待的大纲任务数；报告在两边等待时只计入一次。两个原始等待数也会分开显示。

报告成功或失败时，worker 在更新任务终态的同一事务中追加一行 `report_job_events`（`user_id`、`course_id`、`term_id`、`status`、`error`、`duration_ms`、`created_at`）。成功终态还和课程资料发布处于同一事务，避免重启后重复发布。内存只缓存较小的运行状态，完成后约 30 分钟清理；数据库记录保留，后台任务列表最近 100 条。进程启动时 `init_db()` → `create_all` 补上新增的队列表和会话表，不改已有表。若启动被 v2 迁移提示拦住，先按 README 完成原来的迁移，再启动。

除登录和退出外，后台没有 POST / PUT / PATCH / DELETE。

## nginx

子域把站点根转到应用的 `/admin/`。请保留 `proxy_pass` 末尾的斜杠。

```nginx
# admin.calc.geekhuang.com -> 同一 uvicorn，只暴露 /admin
server {
  server_name admin.calc.geekhuang.com;
  location / {
    proxy_pass http://127.0.0.1:8000/admin/;  # 或 proxy_pass .../admin 按你实现调整
    proxy_set_header Host $host;
    proxy_set_header X-Forwarded-Proto $scheme;
    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
  }
}
```

若希望浏览器地址保持 `/admin` 前缀（和主站同一个域名）：

```nginx
location /admin/ {
  proxy_pass http://127.0.0.1:8000/admin/;
  proxy_set_header Host $host;
  proxy_set_header X-Forwarded-Proto $scheme;
  proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
}
```

另一种是不改路径，只按 Host 分流，并设置 `ADMIN_HOST=admin.calc.geekhuang.com`。应用看到这个 Host 且路径不是 `/admin` 时，会在内部加上 `/admin` 前缀。`/healthz` 保持原样。

```nginx
server {
  server_name admin.calc.geekhuang.com;
  location / {
    proxy_pass http://127.0.0.1:8000;
    proxy_set_header Host $host;
    proxy_set_header X-Forwarded-Proto $scheme;
    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
  }
}
```
