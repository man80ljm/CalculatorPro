# 2026-10-09 本地稳定性改动与验证

本批改动仅在 `D:\HtmlProject\calculatorpro` 开发，尚未部署。桌面版和达成度计算公式未修改。本地试用地址为 `http://127.0.0.1:18090`，启动前使用 SQLite 快照备份本地数据库，备份位于忽略的 `.local/pre-stability/`。

功能分别提交：`c93a3af`（登录与保存保护）、`4e2e045`（保留学期字段及重新登录草稿）、`13d2150`（持久队列与资源保护）、`4ca367b`（加密备份工具与服务器模板）。

## 已完成

| 改动 | 对应问题与老师看到的行为 |
| --- | --- |
| 同账号保留一个有效登录 | 新登录替换旧登录。旧页停止保存、保留草稿并显示重新登录入口；已提交的报告继续执行。 |
| 课程保存版本校验与三方合并 | 不同字段尽量自动合并；同一字段有不同内容时才让老师选择，防止旧窗口覆盖新修改。 |
| 持久化报告队列与输入快照 | 每账号1个执行、3个等待；全站100个等待、最多5个报告线程。重复点击复用当前任务，刷新后自动找回进度。 |
| 重启接续与AI检查点 | 90秒没有心跳的任务自动接续；已落盘的分析直接复用。成功状态、资料和统计事件同一事务发布，避免重复保存。 |
| 计算与输入资源限制 | 计算/解析/打包最多2个处理槽，AI默认全站5个；重操作每账号4个、全站20个同时受理。增加压缩文件、行列、单元格、页数和异常设置检查。 |
| 文件空间与下载 | 每账号默认2GB资料空间、磁盘至少保留200MB；完成的ZIP从磁盘下载，避免在内存中长期保留整包。 |
| 部署参数模板 | 保留网页容器2GB内存，准备Uvicorn并发100、连接等待128、空闲10秒、进程/线程256，以及Nginx基础限流示例。线上未启用本批参数。 |
| 加密异地备份工具 | 数据库、上传资料、代码和配置打成加密包，校验后才上传七牛私有空间。准备夜间任务模板与恢复说明，尚未配置真实云端任务。 |

## 改动文件

以下列出本批所有源代码、配置、测试及文档；`.local/` 中的合成数据、截图和测试结果不纳入Git。

| 类别 | 文件 |
| --- | --- |
| 登录、保存与数据库 | `web_app/auth.py`、`web_app/db.py`、`web_app/edit_merge.py`、`web_app/app.py` |
| 报告与资源保护 | `web_app/report_jobs.py`、`web_app/resource_limits.py`、`web_app/deepseek_pool.py`、`web_app/service.py`、`web_app/limiter.py`、`web_app/storage.py`、`web_app/grade_register.py`、`web_app/syllabus/extract_text.py`、`web_app/v2_routes.py`、`web_app/admin_stats.py` |
| 网页 | `web_app/static/app.js`、`web_app/static/index.html`、`web_app/static/style.css` |
| 部署配置 | `.env.example`、`Dockerfile`、`docker-compose.yml`、`deployment/nginx-resource-limits.conf.example` |
| 备份 | `scripts/backup_bundle.py`、`scripts/nightly_backup.sh`、`scripts/backup_db.sh`、`requirements-backup.txt`、`deployment/backup.env.example`、`deployment/calculatorpro-backup.service`、`deployment/calculatorpro-backup.timer`、`.gitattributes` |
| 验证 | `tests/conftest.py`、`tests/test_edit_sessions.py`、`tests/test_web_flow.py`、`tests/test_syllabus_rules.py`、`tests/test_durable_reports.py`、`tests/test_resource_limits.py`、`tests/test_backup_bundle.py`、`scripts/local_load_test.py`、`requirements-test.txt` |
| 文档 | `README.md`、`docs/admin.md`、`docs/backup.md`、本文件 |

## 验证结果

- 完整测试：**229 passed、1 skipped、1条依赖弃用警告**，约95秒；唯一跳过项为未配置的真实样例。XML：`.local/stability-browser/pytest-final.xml`。
- 20人混合正向/逆向流程：20份报告成功，80次账号隔离检查通过，模拟AI最高并发5次，外部HTTP请求0次。每课35名合成学生，模拟AI等待3秒；工作流程共34.852秒。结果：`.local/load-test/20261009-114606-4268c7/results.json`。
- 计算P95为8.232秒、统计表导出P95为9.681秒、报告完成（含排队）P95为16.293秒。P95表示95%的请求在此时间内完成。这些数字来自本机、单进程、SQLite、模拟AI，不能推算真实DeepSeek或生产容量。
- 浏览器：保存冲突选择、不同设备登录替换与草稿保留、重新登录恢复草稿、刷新找回报告且不主动弹窗、报告完成后下载全部通过，无脚本异常。记录：`.local/stability-browser/browser-results.json`。
- 实际进程恢复：AI答案保存后、发布资料前强制终止独立服务；重启后自动接续同一任务，第二个进程AI调用0次，成功事件和资料批次各1次，ZIP完整。验收为缩短等待使用5秒租约，生产默认90秒。记录：`.local/stability-browser/restart-results.json`。
- 计算口径：相同合成正向/逆向成绩表分别在基线 `6f1844d` 和当前版本运行，达成度及表5每个单元格完全一致。记录：`.local/stability-browser/math-before.json`、`math-after.json`。
- 备份：合成SQLite加密/解密与恢复通过；篡改包被拒绝；七牛上传/远端哈希校验使用模拟SDK。没有真实上传或真实凭据。
- JavaScript语法、Git差异空白检查、两个Shell备份脚本语法检查通过。

## 界面截图

截图均使用合成账号和课程：

- `.local/stability-browser/edit-conflict.png`：同一字段冲突时保留双方选择。
- `.local/stability-browser/session-preserved.png`：旧登录停止保存并保留草稿。
- `.local/stability-browser/report-resumed.png`：刷新后继续显示报告进度，不弹窗打断。

## 上线前仍需完成

1. 在独立PostgreSQL/生产容器环境验证新会话锁、报告队列、资源上限和原数据兼容性，再执行部署、回滚和页面验收。当前没有连接或修改线上服务器。
2. 配置七牛私有空间、备份凭据、服务器外保管的加密密钥、云端保留策略，实测一次上传和PostgreSQL完整恢复后启用定时任务。备份说明见 `docs/backup.md`。
3. 真实AI调用速度和长时间运行容量未在本轮测量。AI返回到答案落盘之间退出仍可能重复调用，不能保证付费接口恰好执行一次。

当前资源槽属于单进程，正式部署继续使用一个Uvicorn worker；暂不增加生产worker数。夜间备份为保证数据库与资料一致，会短暂停止网页后备份，重启网页再上传；模板尚未启用。
