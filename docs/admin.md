# 运维后台与邀请码管理

后台提供教师数、课程数、报告成败和排队情况统计，以及邀请码新增、次数调整、启停和删除、PDF 课表管理和识别 API 设置。课程和报告功能保持只读，不能改密码、删课、代登录、重新生成报告。

页面在本服务的 `/admin`。`/admin` 会转到 `/admin/`。目标子域是 `admin.calc.geekhuang.com`，由 nginx 反代到这个前缀，或按 Host 分流。

## 环境变量

| 变量 | 说明 |
| --- | --- |
| `ADMIN_PASSWORD` | 明文密码。只写在部署环境里，不要写进仓库、日志或页面。 |
| `ADMIN_PASSWORD_HASH` | argon2 哈希。**设置后优先于明文**，`ADMIN_PASSWORD` 会被忽略。 |
| `ADMIN_USERNAME` | 登录名，默认 `admin`。大小写不敏感。 |
| `ADMIN_HOST` | 可选。例如 `admin.calc.geekhuang.com`。设置后，Host 匹配时根路径也进后台；路径以 `/admin` 开头时，其他 Host 也能打开（方便本机）。未设置时只认 `/admin`。 |
| `COOKIE_SECURE` | `true` 时运维 Cookie 带 Secure，和教师 Cookie 一样。只在 HTTPS 后面打开。 |
| `INVITE_USER_DEFAULT_USES` | 新生成的用户专属码可注册总人数，默认 `5`，范围 `0` 到 `1000000`，或 `unlimited`。管理员可逐码调整，修改默认值不会改变已有码。 |
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
- 管理员认证不查 `users` 表。教师的用户名和密码即使碰巧相同也不能进后台，除非你把同一口令配进了上面的环境变量。

`Path=/` 是为了子域反代：浏览器上的路径是 `/api/overview`，不是 `/admin/api/overview`。Cookie 若限定在 `/admin`，这种反代就带不上。教师站看到这枚 Cookie 会忽略。

## 只读接口

都要运维 Cookie。未登录是 401。

| 方法 | 路径 | 内容 |
| --- | --- | --- |
| GET | `/admin/api/overview` | 教师数、课程数、本周活跃教师、今日/本周报告成功和失败、排队深度 |
| GET | `/admin/api/users` | 用户名、专属邀请码、邀请来源、注册时间、课程数、学期数、最近登录、最近报告 |
| GET | `/admin/api/jobs` | 报告任务及按用户分页的历史记录；可选 `user_id`、`page`（每页 10 条） |

任务页在历史记录上方提供用户下拉框，默认显示全部用户最近 10 条，可切换用户及翻页查看更早记录。筛选在数据库查询前执行，不受原先最近 100 条的限制；实时任务保持展示全部用户。`GET /admin/api/jobs` 返回 `jobs` 和 `history`，后者包含用户列表、所选 `user_id`、页码、每页条数、总数与总页数。空结果返回第 1 页，超出总页数会回到最后一页。刷新及自动刷新保持用户和页码，切换用户回到第 1 页。

手机端（宽度不超过 720 px）顶部右侧“菜单”按钮展开导航抽屉，退出入口收在抽屉底部。支持关闭按钮、点击遮罩、Esc 和浏览器返回关闭；电脑端保留左侧导航。设置页只展示 API 配置，邀请码在独立“邀请码”页面管理。

页面：`/admin/` 概览，`/admin/users`，`/admin/jobs`，`/admin/invites` 邀请码管理，`/admin/login`。静态文件在 `/admin/static/`。HTML 里用相对路径（`static/admin.css`、`api/overview`），子域反代和 `/admin` 前缀两种接法都能用。

今日、本周按北京时间（UTC+8），周一 0 点起算。库里的时间仍是 UTC。

本周活跃教师：本周内新建过 `user_sessions`，或 `courses.updated_at`、`course_files.created_at` 落在本周的用户，按 `user_id` 去重。

排队深度 = 数据库队列中等待的报告数 + DeepSeek 池中等待的大纲任务数；报告在两边等待时只计入一次。两个原始等待数也会分开显示。

报告成功或失败时，worker 在更新任务终态的同一事务中追加一行 `report_job_events`（`user_id`、`course_id`、`term_id`、`status`、`error`、`duration_ms`、`created_at`）。成功终态还和课程资料发布处于同一事务，避免重启后重复发布。内存只缓存较小的运行状态，完成后约 30 分钟清理；数据库记录保留，后台任务列表最近 100 条。进程启动时 `init_db()` → `create_all` 补上新增的队列表和会话表，不改已有表。若启动被 v2 迁移提示拦住，先按 README 完成原来的迁移，再启动。

## 邀请码与正式注册

正式注册 `POST /api/register` 的 JSON 必须含 `username`、`password` 和 `invite_code`。邀请码是服务端用 `secrets` 生成的 6 位大写字母与数字混合码；输入时忽略大小写和首尾空白。无效、停用、删除或次数已满的码均不能注册，旧的无邀请码页面也不能绕过服务端校验。

注册成功会消耗邀请来源码的一次额度，生成该用户自己的专属码，并保存邀请关系。以上与账号创建在同一数据库事务中完成，重复用户名、密码不合要求或生成失败不会扣次。PostgreSQL 使用行锁，SQLite 使用写锁，多个请求争抢最后一个名额不会超发。注册响应和已登录的 `GET /api/me` 包含个人码、使用次数和状态；用户可从「我的邀请码」复制。

| 方法 | 路径 | 请求 / 返回 |
| --- | --- | --- |
| GET | `/admin/api/invites` | 未删除的邀请码列表和 `csrf_token` |
| POST | `/admin/api/invites` | `{"max_uses": 5}`，创建管理员码，返回 201；省略时为 5 |
| PATCH | `/admin/api/invites/{id}` | `{"max_uses": 5}` 或 `{"enabled": false}`，可同时修改 |
| DELETE | `/admin/api/invites/{id}` | 删除邀请码，已注册账号和邀请记录保留 |

上述写接口均要求独立的管理员 Cookie，并在 `X-Admin-CSRF` 头中携带列表接口返回的、与当前会话绑定的令牌。POST/PATCH 只接受 JSON，拒绝跨来源写请求。教师 Cookie 无管理权限。后台写接口仅包含登录/退出、邀请码管理、课表识别/保存/修正和识别 API 设置。

`max_uses` 表示**可注册总次数**，不是额外次数。例如已用 2 次、总次数 5 次，剩余 3 次；总次数不能降到 2 以下。`0` 表示没有可用名额，`null` 表示不限次数。停用不会抹掉已用次数，重新启用时仍受额度限制。删除后码不可恢复、不可重复生成，个人码被删除后也不会自行补发。

启动时 `init_db()` 会新建 `invitation_codes`、`invitation_registrations` 两张表，不修改现有用户表。已有账号继续使用原密码登录，并在首次加载个人信息时补发一个专属邀请码。历史账号标为「历史账号」，不会伪造之前不存在的邀请关系。数据库不预置任何公开邀请码；上线后管理员先创建首批码，再向用户发放。原型的浏览器本地邀请码不会导入正式数据库。

发布前备份数据库和上传文件，保留旧镜像与代码，设置独立管理员密码，按现有部署流程重启应用。随后检查管理员创建码 → 注册 → 个人码邀请下一人 → 原码额度已满被拒绝；原有账号登录及课程数据也应核对。此版本已于北京时间 2026-10-12 发布，默认 5 次；线上验收与备份见 [邀请码上线记录](invitations-deploy-20261012.md)。

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


## 课表管理与识别设置

页面 `/admin/schedule` 提供文字层 PDF 上传（8 MB、15 页以内）、后台识别、进度查询、核对入库、学期/课程/教师/班级查询和记录修正。关闭弹窗不会终止识别，草稿和已完成批次保存到正式数据库。同一 PDF 不会重复入库，重复上传会复用草稿或提示已入库。重启中断的任务会明确提示重新选择 PDF，已完成批次可复用；完整原 PDF 不长期保存。扫描 PDF 明确拒绝，不以示例内容代替结果。

页面 `/admin/settings` 提供课表识别使用的 API 来源、模型、连接测试和保存。默认沿用已配置的服务器 DeepSeek 单密钥或 A/B 多密钥池，也可填写独立密钥。设置仅作用于课表识别，不覆盖原报告服务的环境配置。独立密钥通过 Fernet 加密保存至数据库，密钥材料从服务器 `SECRET_KEY` 经独立用途 HMAC 派生，不能公开或随意更换这项服务器配置。页面不回显密钥，留空保留之前保存的独立密钥。连接测试只请求官方模型列表，验证认证与模型可用性；不验证余额或教学内容生成。接口固定为官方 HTTPS，不允许任意地址或重定向。

| 方法 | 路径 | 内容 |
| --- | --- | --- |
| GET / POST | `/admin/api/settings` | 读取安全元数据 / 保存 `source`、`model`、`key` |
| POST | `/admin/api/settings/test` | 测试所选 API 连接（每分钟最多 5 次） |
| GET | `/admin/api/schedule` | 已保存课表、开课记录及待处理草稿 |
| POST | `/admin/api/schedule/analyze` | `filename` + base64 `pdf`，返回 202 任务 ID（每分钟最多 3 次，同一服务同时识别 1 份） |
| GET | `/admin/api/schedule/jobs/{id}` | 识别进度或待核对结果 |
| POST | `/admin/api/schedule/save` | `jobId` + 已核对的 `result`，事务入库，重复保存幂等 |
| POST | `/admin/api/schedule/record` | `id` + `record`，核对修正，多个教师按单教师记录拆分 |

以上接口全部要求管理员 Cookie；写入还要求与当前会话绑定的 `X-Admin-CSRF` 和 JSON 正文。教师账号无法读取后台课表或 API 配置。新增 `admin_api_configuration`、`schedule_documents`、`schedule_drafts`、`schedule_records`、`schedule_recognition_chunks` 五张表，由 `init_db()` 兼容性创建，不改已有表。识别保存失败整笔回滚，原有课表保留。原文引用、待核对状态及学分/学时分别保留，未标明的学时不会推算。

识别使用服务器配置时共用现有 DeepSeek 账号租约、密钥轮换和全站并发上限；独立 API 密钥也计入全站并发上限。本次用户明确选择先上线功能，本地 2 份已入库 PDF 的 512 条记录不迁移，仍保留在独立原型数据库。

课表管理和设置已于北京时间 2026-10-12 以 `060073a` 上线。真实连接、PDF 流程、手机操作及资料保全结果见 [课表与设置上线记录](teaching-admin-deploy-20261012.md)。
