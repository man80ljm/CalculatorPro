# CalculatorPro

一个基于 PyQt6 的课程目标达成度计算与报告生成工具，支持“正向计算”与“逆向推分”两种模式，并可生成 Excel/Word 报表与 AI 分析报告。

## 功能概览

- 正向模式：导入“正向模板”成绩，计算课程目标达成度并输出明细与统计。
- 逆向模式：导入“逆向模板”总分，按分布/噪声参数反推方法级分数并输出明细与统计。
- 统计与报告：生成表2（课程成绩统计）、表5（达成度评价结果）、表6（AI 改进报告）及最终拼接报告。
- 支持分数跨度、分布模式、噪声注入配置。

## 环境要求

- Python 3.10+
- Windows（推荐）

## 依赖安装

```bash
pip install -r requirements.txt
```

如需使用镜像：

```bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

## 使用步骤

### 1) 先配置“课程考核与课程目标对应关系”
- 打开“课程考核与课程目标对应关系”并设置目标数、考核环节与方法。

### 2) 正向模式（导入方法级成绩）
1. 选择“正向模式”
2. 点击“模板下载”，生成正向模板
3. 填入学生成绩并保存
4. 点击“导入文件”选择该模板
5. 点击“导出结果”生成报表

### 3) 逆向模式（导入环节总分）
1. 选择“逆向模式”
2. 点击“模板下载”（必须先完成关系表设置）
3. 填入环节总分并保存
4. 点击“导入文件”选择该模板
5. 选择分数跨度 / 分布模式 / 噪声注入
6. 点击“导出结果”生成报表

### 4) 生成 AI 报告（表6）
- 在“设置”中填写 API Key
- 点击“生成报告”

## 输出文件说明（outputs 目录）

正向模式：
- `{课程名}成绩明细.xlsx`（成绩明细 + 课程成绩统计 + 达成度评价结果）
- `2.课程成绩统计表.docx`
- `5.基于考核结果的课程目标达成情况评价结果表.docx`
- `6.课程目标达成情况分析、存在问题及改进措施表.docx`
- 最终拼接报告（基于 `report_template.docx`）

逆向模式：
- `{课程名}成绩明细（逆向）.xlsx`（成绩明细 + 课程成绩统计 + 达成度评价结果）
- `{课程名}正向成绩表（逆向生成）.xlsx`
- `{课程名}课程目标达成情况评价结果（逆向）.xlsx`
- `2.课程成绩统计表.docx`
- `5.基于考核结果的课程目标达成情况评价结果表.docx`
- `6.课程目标达成情况分析、存在问题及改进措施表.docx`
- 最终拼接报告（基于 `report_template.docx`）

## 打包为 EXE（PyInstaller）

在项目根目录执行：

```bash
pyinstaller --onedir -w --clean --noconfirm --icon=calculator.ico --add-data "calculator.ico;." --add-data "report_template.docx;." --collect-all PyQt6 --collect-submodules docx --collect-submodules openpyxl --collect-submodules pandas --collect-submodules numpy main.py
```

说明：
- `--onedir` 打包为目录模式
- `-w` 去掉控制台
- 必须带入 `calculator.ico` 与 `report_template.docx`
- 打包后模板会被放到 `_internal`（或 onefile 临时目录），程序会从资源目录读取

## 常见问题

- 模板导入提示不匹配：确认当前模式与模板类型一致。
- 逆向模板下载失败：请先完成“课程考核与课程目标对应关系”设置。
- 文件被占用：关闭已打开的 Excel/Word 文件后重试。

## 网页版

桌面端 PyQt6 程序保持原样，仍然用 `main.py` 启动，配置仍写在本机 `%APPDATA%/CalculatorApp/config.json`（旧模块 `utils_app` 使用 `GradeAnalysisSystem`）。网页版单独放在 `web_app/`，复用 `core_app` 的计算、Excel/Word 导出和 DeepSeek 报告，不导入 PyQt。

教师自行注册账号（用户名或邮箱 + 至少 8 位密码）。登录后按课程建立文件夹：开课信息、占比、课程目标对应关系、成绩表、上一学年达成度表，以及生成的模板和报表都保存在该文件夹里，换一台电脑登录仍能查看和下载。计算仍在当次请求的临时目录中进行，结束后把输入和结果写入课程目录。DeepSeek 密钥只放在服务器环境变量里，页面不提供密钥输入框。登录和注册按 IP（登录另按用户名）限流；修改密码需要输入当前密码，修改后其他设备上的登录会话会失效。

课程考核与课程目标对应关系是一张可编辑表格：格子可以直接改，也能从 Excel 粘贴，按 Excel 的复制格式解析：制表符分列、换行分行；单元格内用 Alt+Enter 换行、含制表符或双引号时，Excel 会给该格加引号（内部 `""` 表示一个 `"`），这种格子会完整保留为一格，不会被拆开。兼容 Windows 换行和行尾换行。可以添加行、删除选中行。粘贴从当前格子开始，行数不够时自动加行。

### 环境变量

| 变量 | 说明 |
| --- | --- |
| `SECRET_KEY` | 签名会话 Cookie 的密钥。未设置时进程直接退出。 |
| `POSTGRES_USER` / `POSTGRES_PASSWORD` / `POSTGRES_DB` | Postgres 账号，供 compose 里的数据库容器使用。 |
| `DATABASE_URL` | 例如 `postgresql+psycopg://calculatorpro:密码@postgres:5432/calculatorpro`。**必填**：未设置时服务拒绝启动。本地试用可写成 `sqlite:///./calculatorpro.db`，但必须同时设置 `ALLOW_SQLITE=1`。 |
| `ALLOW_SQLITE` | 设为 `1` 才允许使用 SQLite（测试和本地试用）；未设置 `DATABASE_URL` 时也会退回到 `./calculatorpro.db`。部署时不要设置。 |
| `DEEPSEEK_API_KEY` | 只配这一把时，当作单账号单钥，仍走下面的排队和故障转移。留空且没有账号池时，AI 分析报告和从大纲建课会提示未配置，模板、计算和 xlsx/docx 导出仍可用。 |
| `DEEPSEEK_KEYS_A` / `DEEPSEEK_KEYS_B` | 两个账号的密钥，逗号分隔，例如 `sk-example-a1,sk-example-a2`。也可以写成密钥文件路径：每行一把，空行和 `#` 注释会忽略。 |
| `DEEPSEEK_KEYS_A_FILE` / `DEEPSEEK_KEYS_B_FILE` | 密钥文件路径。设置后优先于同账号的 `DEEPSEEK_KEYS_A` / `DEEPSEEK_KEYS_B`。 |
| `DEEPSEEK_ACCOUNT_CONCURRENCY` | 每个账号同时进行的 AI 任务数，默认 5。这是打到 DeepSeek 的容量；两个账号都满时，新的报告和大纲读取排队。报告比大纲先拿到空位。 |
| `DEEPSEEK_QUEUE_REPORT_SECONDS` | 排队时报告的估计耗时，默认 30。`queue_eta_seconds` = `queue_position` × 这个值。 |
| `DEEPSEEK_QUEUE_SYLLABUS_SECONDS` | 大纲读取的估计耗时，默认 15。 |
| `DEEPSEEK_MODEL` | 模型名，默认 `deepseek-flash`。 |
| `DEEPSEEK_BASE_URL` | 接口根地址，默认 `https://api.deepseek.com`。 |
| `COOKIE_SECURE` | `true` 时会话 Cookie 带 Secure。只在 HTTPS 反向代理后打开，默认 `false`。教师 Cookie 和运维 Cookie 都看这个值。 |
| `ADMIN_PASSWORD` | 运维后台明文密码，只放在服务器环境里。与教师账号无关。不设密码时后台不放行。 |
| `ADMIN_PASSWORD_HASH` | 运维密码的 argon2 哈希。设置后优先于 `ADMIN_PASSWORD`，明文会被忽略。 |
| `ADMIN_USERNAME` | 运维登录名，默认 `admin`。 |
| `ADMIN_HOST` | 可选。设成运维子域后，该 Host 的根路径也进后台；`/admin` 在其他 Host 上仍然可用。 |
| `UPLOAD_DIR` | 课程文件目录。容器里是 `/data/uploads`，对应宿主机 `/data/uploads/calculatorpro`。 |
| `MAX_UPLOAD_MB` | 单个 Excel 上限，默认 10。 |
| `TRUSTED_PROXIES` | 受信反向代理网段（逗号分隔），只有来自这些地址的请求才读取 `X-Forwarded-For` 最右一项作为客户端 IP 用于限流。默认 `127.0.0.1/32,::1/128,172.16.0.0/12`（本机 + Docker 默认网段）。 |

示例见 `.env.example`。把真实值写在 `.env`，不要提交。`.env` 已在 `.gitignore` 中。

只读运维后台在 `/admin`（`/admin/` 同样可用），不使用教师 `users` 表。未设置 `ADMIN_PASSWORD` 或 `ADMIN_PASSWORD_HASH` 时，登录和数据接口返回 503。环境变量、Cookie 和 nginx 示例见 `docs/admin.md`。

### 用 Docker 运行

在项目根目录准备目录和权限。应用容器以 uid `10001` 运行，必须能写入上传目录：

```bash
sudo mkdir -p /data/uploads/calculatorpro /data/databases/calculatorpro/postgres
sudo chown -R 10001:10001 /data/uploads/calculatorpro
cp .env.example .env
# 编辑 .env：SECRET_KEY、数据库口令、DATABASE_URL。DEEPSEEK_API_KEY 可留空。
docker compose up -d --build
```

浏览器打开 <http://127.0.0.1:18090> 。网页端口只绑定在 `127.0.0.1`，应用内存限制约 2GB。Postgres 16 不向宿主机发布端口，内存限制约 512MB，数据在 `/data/databases/calculatorpro/postgres`。应用会等数据库健康检查通过后再启动。`/healthz` 返回 `{"status":"ok"}` 时表示进程和数据库都可用。上传的 Excel 默认不超过 10MB。

### 本地直接运行

```bash
pip install -r requirements-web.txt
```

Windows PowerShell：

```powershell
$env:SECRET_KEY="请换成一长串随机字符"
$env:DATABASE_URL="sqlite:///./calculatorpro.db"
$env:ALLOW_SQLITE="1"
$env:UPLOAD_DIR="./data/uploads"
$env:DEEPSEEK_API_KEY=""
uvicorn web_app.app:app --host 127.0.0.1 --port 18090
```

Linux / macOS：

```bash
export SECRET_KEY="请换成一长串随机字符"
export DATABASE_URL="sqlite:///./calculatorpro.db"
export ALLOW_SQLITE=1
export UPLOAD_DIR="./data/uploads"
export DEEPSEEK_API_KEY=""
uvicorn web_app.app:app --host 127.0.0.1 --port 18090
```

### 页面上的流程

1. 注册或登录自己的账号。退出后该会话立即失效。登录后可以修改密码，需要填写当前密码。
2. 新建课程文件夹，填写开课信息、课程基本信息、考核与课程目标对应关系、毕业要求，以及课程简介和各目标要求，然后保存。
3. 选择正向或逆向模式。逆向模式可设置分数跨度、分布模式和噪声注入。
4. 下载成绩模板，填好后导入；学校的成绩登记表支持 Excel（.xlsx、.xls）、Word（.docx，包含学生成绩表格）和带文字的 PDF。旧版 .doc 请先在 Word 中另存为 .docx。文件记在当前课程下。
5. 「计算达成度」查看人数、平均分和各目标达成度。
6. 「导出 xlsx / docx」下载统计表压缩包（成绩明细、表 1 至表 5），并留在课程文件夹里。
7. 「生成 AI 分析报告」在密钥已配置时下载表 6 和拼接后的总报告。上一学年达成度表可以先导入。没有上学年数据时不做同比，评价表该列显示「—」。有数据时按实际数值对比。
8. 之后在任意电脑登录，打开同一课程，点击「查看各学期资料」。每个学期默认显示最新资料包，单个文件、导入资料和历史版本可展开查看；查找或下载旧学期的资料不会切换正在编辑的学期。

资料包统一使用「课程名_2026-2027学年_第1学期.zip」的简短名称，班级和生成时间显示在页面上。每次成功计算、导出或生成报告保存一个独立版本；新版资料包包含本次生成的文档，以及「导入资料」中的本次成绩表和上一轮达成度表（如有）。旧压缩包保留原有内容，下载时使用简短名称，旧单文件仍能下载。生成批次保存在 `file_batches` 表中，启动时只补建新表，不修改既有表结构。

主要接口（除注册、登录和健康检查外，都需要登录 Cookie）：

- `POST /api/register`、`POST /api/login`、`POST /api/logout`、`POST /api/password`
- `GET/POST /api/courses`，`GET/PATCH/DELETE /api/courses/{id}`
- `GET/POST /api/courses/{id}/files`，`GET /api/courses/{id}/files/{file_id}`
- `GET /api/courses/{id}/materials`：查看本课程所有学期的资料、生成版本与下载入口
- `POST /api/courses/{id}/template`：按课程设置生成模板并保存
- `POST /api/courses/{id}/calculate`：用已保存或随请求上传的成绩计算
- `POST /api/courses/{id}/export`：导出 xlsx/docx 压缩包并保存
- `POST /api/courses/{id}/ai-report`：生成 AI 报告压缩包并保存
- `POST /api/courses/{id}/report-jobs`：开始生成报告。同一用户、同一课程、同一学期已有进行中的任务时返回原来的任务
- `GET /api/report-jobs/{job_id}`：进度。账号都在忙时 `stage` 为 `queued`，`stage_label` 为「前面还有 N 人，大约再等 X 秒」，`queue_position` 从 1 计（1 表示下一个开始），`queue_eta_seconds` 等于队列位置乘以该任务的估计秒数。没有排队时这两个字段为 `null`。轮到之后依次是 `calculate`、`tables`、`ai`、`package`
- `GET /api/ai-status`：是否已配置密钥（不返回密钥本身）
- `GET /healthz`：健康检查，同时确认数据库可连接

登录、注册和 AI 接口有频率限制。登录同时按 IP 和用户名计数。报告接口每个客户端大约 30 次/分钟，大纲读取每个用户大约 20 次/分钟，用来防刷。真正同时打到 DeepSeek 的数量由每个账号的并发上限控制，默认 5。访问别人的课程或文件会得到 404。

### 备份

数据库和上传文件要一起留：

```bash
mkdir -p backups
docker compose exec -T postgres pg_dump -U "$POSTGRES_USER" "$POSTGRES_DB" | gzip > "backups/calculatorpro-$(date +%Y%m%d-%H%M%S).sql.gz"
```

也可以直接运行 `sh scripts/backup_db.sh`（它会读取 `.env`，用 `docker compose exec -T postgres pg_dump ... | gzip` 写到 `backups/`）。同时把宿主机目录 `/data/uploads/calculatorpro` 拷走，里面是各教师的成绩表和导出文件。

### v2 迁移说明

v2 为一门课增加学期记录：课程仍只建一次（新学期采用当前大纲，已建立的学期保存各自的课程目标、关系表和简介），每个学期另记学年学期、任课教师、班级、人数，以及该学期的成绩和报告。文件仍放在原来的课程文件夹 `UPLOAD_DIR/<用户>/<课程>/<文件名>`，不按学期搬目录；学期只在数据库的 `course_files.term_id` 上关联，因此旧路径还能打开。

程序启动时如果发现已有课程表、但还没有 `terms` 或 `course_files.term_id`，会直接退出并提示运行迁移，**不会在老师打开页面时自动改数据**。新装的空库会在启动时建好新表，不必先迁移。

迁移可以重复执行：已经有学期的课不会再生成第二条默认学期。执行前会把 `courses`、`course_files`、`terms` 导出成 JSON。退回时把学期字段写回课程设置，去掉学期表和 `term_id`，文件仍留在原处。

先做数据库备份（只导出这几张表即可；整库备份见上文）：

```bash
mkdir -p backups
docker compose exec -T postgres pg_dump -U "$POSTGRES_USER" "$POSTGRES_DB" -t courses -t course_files -t terms | gzip > "backups/v2-pre-migrate-$(date +%Y%m%d-%H%M%S).sql.gz"
```

然后在项目目录、带好与运行中服务相同的 `DATABASE_URL`：

```bash
python -m web_app.migrate_v2 --dry-run
python -m web_app.migrate_v2 --backup-dir ./backups
python -m web_app.migrate_v2 --downgrade --backup-dir ./backups
```

`--dry-run` 只打印将新建的学期数和待挂上的文件数。后两条都会先写 JSON 备份。SQLite 试用库同样适用，但必须 `ALLOW_SQLITE=1`。

从大纲建课走 `POST /api/syllabus/extract` 和 `POST /api/courses/from-syllabus`。每个登录用户每 60 秒最多 20 次读取，与报告接口的每 60 秒 30 次分开计数。大纲读取会等到有 DeepSeek 空位再调用模型；超过本次读取的等待时间仍没有空位时返回 503，`detail` 是「前面还有 N 人，大约再等 X 秒」。成绩登记表导入走 `POST /api/courses/{id}/grade-register`，不调用 AI。分析报告一次请求写完全部段落，模型由 `DEEPSEEK_MODEL` 决定，默认 `deepseek-flash`。

## 网页版未覆盖的桌面细节

- API Key 仍只在桌面「设置」里填写和测试连接；网页从环境变量读取，不提供密钥输入框。
- 桌面的关系表仍是原来的 PyQt 弹窗。网页用一张可粘贴的表格完成同一份数据，两边都交给 `core_app` 计算。
- 桌面打包（PyInstaller）和本机 `outputs/` 目录只属于桌面程序。

## 同名课程和大纲更新

同一账号建课时，同名课程会提示打开已有课程；从大纲建课还可以选择「用新大纲更新」。课程代码相同也视为同一门课；同名但代码明确不同的课程可以分别建立，课程列表会显示代码以便区分。

更新前保存旧大纲和各学期设置快照。更新后，新导入的学期使用新版；旧学期仍使用当时的目标、比例和对应关系，成绩、报告和上一轮达成度记录保留。若新旧学期大纲版本不同，自动读取上一轮达成度后会提示核对目标对应关系。已有学期始终在本学期快照上修改，不会覆盖新版默认设置。

启动时增加 `course_versions` 表，不改已有表结构。已有课程第一次保存或更新大纲时建立第1版和学期快照。更新接口为 `POST /api/courses/{id}/syllabus-version`，提交格式与从大纲建课相同；`expected_version` 用于防止旧窗口覆盖其他窗口刚更新的大纲。

## 本地多用户压测

用项目虚拟环境运行，不需要配置 AI 密钥：

```powershell
.\.venv\Scripts\python.exe scripts\local_load_test.py --users 5 10 20
# 20人混合使用正向、逆向
.\.venv\Scripts\python.exe scripts\local_load_test.py --users 20 --mode mixed
# 多轮重复操作；轮次之间默认等60秒，保留程序原有的报告限流
.\.venv\Scripts\python.exe scripts\local_load_test.py --users 20 --mode mixed --rounds 3
```

脚本只启动 `127.0.0.1` 上的独立测试服务；每档并发建立一个新的 SQLite 库和上传目录，放在 `.local/load-test/<时间>/`。不读取 `.env`，不接受线上地址，不操作18090端口的日常测试库。老师账号、密码和35名学生成绩全部自动合成；AI HTTP 调用替换为模拟响应（默认等待3秒），外部 HTTP 传输被阻止，不消费真实额度。仅测试服务批量注册/登录账号的同IP限流放宽，报告的30次/分钟限流和每账号5个并发槽仍保留。

每位老师同时导入成绩，随后与单人基准比较计算结果，再同时计算、导出、生成报告并下载。验证混合班整表人数、达成度、表5、本人原始成绩、压缩包完整性、后台成功事件，以及其他账号访问课程/成绩/报告均返回404。结果写入 `results.json`，包含各接口调用数量、失败次数、P50/P95/最长响应时间、排队人数和模拟AI并发数；非零退出码表示发现失败。

### 2026-10-09 本机结果

本机、单进程 Uvicorn、独立 SQLite，每课35人、单轮正向流程，AI模拟等待3秒。以下报告时间从提交报告计起，包含排队和文档生成；P95表示95%的请求在此时间内完成。

| 同时使用人数 | 计算P95 | 统计表导出P95 | 报告完成P95 | 排队人数 | 失败 |
| --- | --- | --- | --- | --- | --- |
| 5 | 1.91秒 | 1.88秒 | 4.79秒 | 0 | 0 |
| 10 | 3.56秒 | 4.27秒 | 8.60秒 | 5 | 0 |
| 20 | 9.21秒 | 8.14秒 | 17.45秒 | 15 | 0 |

35份报告均成功，模拟AI最高并发5；140次跨账号访问检查全部返回404。另测20人正向/逆向混合流程，20份报告全部成功，达成度和表5与各自单人结果一致。完整pytest为208 passed、1 skipped（未配置真实样例）、1条依赖弃用警告。

首次20人测试发现两个报告保存失败，定位到SQLite文件库使用StaticPool共用连接的事务冲突；已改为文件库默认连接池，内存库保留原共享连接。补充未提交数据不可被其他事务读取/提交、20个并发事务独立提交/回滚的回归测试后，以上各档复测通过。配置依据：[SQLAlchemy SQLite线程与连接池说明](https://docs.sqlalchemy.org/en/20/dialects/sqlite.html#threading-pooling-behavior)。

这是本机的短时流程压测。真实DeepSeek响应时间、额度/限流、部署机器的资源和PostgreSQL容量没有在本次测试中测量，不能据此承诺线上20人性能。长时间稳定性可增加轮数继续观察。该轮压测对应的旧版本把报告任务保存在单进程内存；本地新版的队列见下文。

## 本地稳定性开发（尚未部署）

新登录自动替换同账号的旧会话，旧页停止写入并保留课程填写草稿。课程写接口要求 `X-Course-Revision`（读取课程返回的 `edit_revision`），旧页不能静默覆盖新内容；普通保存带 `base` 时自动合并不同字段，相同字段保留双方选择。会话切换与写入在数据库内按用户串行验证；正在处理的报告归账号所有，不因会话切换取消。

报告输入快照保存在上传目录的 `.jobs/<任务号>`，任务和进度保存在新增 `report_tasks` 表。重复点击复用本学期正在执行的任务；每账号最多 1 个执行中、3 个等待，全站最多 100 个等待。最多 5 个报告工作线程，计算、文档解析和打包共用 2 个处理槽，AI 默认全站合计最多 5 个。等待任务不占工作线程，不在内存中保留输入字节；完成后直接从课程资料读取压缩包，临时输入与 AI 检查点在终态满一天后清理。

刷新后网页自动找回当前学期任务，不主动弹窗打断填写。后台每次领取任务取得租约，90 秒内没有心跳则自动接续；已落盘的 AI 答案会复用。若进程恰好在 API 返回与答案落盘之间退出，仍可能再次调用 API，不能保证外部接口恰好调用一次。报告使用提交时的成绩和设置，生成期间的新修改用于下一次生成。

重操作每账号最多 4 个、全站最多 20 个同时受理；繁忙时网页自动等 3 秒再试，最多重试 3 次。单账号写操作最多 120 次/分钟，进度查询不消耗该额度。上传除原来的单文件大小限制外，增加解压后默认 100 MB、Excel 5000 行/200 列/25 万单元格、PDF 100 页的检查；默认每账号资料空间 2 GB，磁盘保留至少 200 MB。参数见 `.env.example`，不改变达成度公式、期望值或表 5。

畸形设置最多允许 30 个课程目标、30 个考核环节、150 个考核方式；超出时在生成前拒绝。已保存资料和报告改用磁盘流式下载，避免同时下载时在内存里复制整包。Compose 还准备了 100 个连接/任务的 Uvicorn 并发上限、128 个等待连接、10 秒空闲连接超时和 256 个进程/线程上限，仍保留原 2 GB 内存配置；这些部署参数尚未在线上启用。

当前计算槽和 AI 池仍属于进程内资源，部署继续使用一个 Uvicorn worker。`REPORT_WORKER_MODE=external` 仅用于专项测试或另行配置后台处理进程，不能直接据此增加生产 worker 数；正式发布前还需在 PostgreSQL 和服务器容器里验证这批改动。

完整本地测试依赖：`pip install -r requirements-test.txt`，然后 `pytest -q`。新测试覆盖登录替换、保存冲突、学期身份保持、队列限量、过期租约接续、复用 AI 答案和异常文件限制。备份工具及服务器夜间任务的准备说明见 [docs/backup.md](docs/backup.md)。七牛真实凭据未配置时，不会上传云端；这批改动尚未部署。

本地新版最终验证：**229 passed、1 skipped**，与本轮接手时的208 passed相比新增21项通过用例。2026-10-09再次执行20人正向/逆向混合流程，每课35名合成学生、AI模拟等待3秒，20份报告成功、80次跨账号隔离检查通过、无外网请求。计算P95为8.23秒、统计表导出P95为9.68秒、报告完成P95为16.29秒；这些是本机短时模拟结果，不代表线上真实AI性能。实际强制终止隔离服务并重启后，原任务自动接续，已落盘的AI答案复用，成功事件与资料批次各保存一次。同一份正向/逆向成绩表在改动前后，达成度和表5所有单元格一致。详细改动文件与验证记录见 [docs/stability-20261009.md](docs/stability-20261009.md)。

## 2026-10-09 上线记录


功能版本 `088d25e` 已部署到 `https://calc.geekhuang.com/`，正式目录为 `/data/projects/calculatorpro`。镜像复用原有生产依赖，只覆盖已验证的代码；正式Compose中上传目录和只读密钥目录挂载保留，`.env`沿用服务器原配置。发布前保存数据库、上传文件、旧代码和旧镜像，发布后核对原有记录和上传文件完整。

发布目录和备份在 `/data/projects/calculatorpro-releases/088d25e-20261009/`，旧镜像标签为 `calculatorpro-web:rollback-20261009`。`deployment-result.json`记录镜像ID、备份位置及数据核对结果。生产库没有导入本地账号、课程或密钥；线上浏览器验收使用的临时合成账号与课程已清理。

独立PostgreSQL16预发布环境中使用20个合成账号，每课35人，混合正向/逆向，AI模拟等待3秒；验证计算、导出、报告、跨账号隔离及大纲版本。使用同一成绩表分别在旧镜像和新镜像中计算、导出，达成度与表5一致。正式HTTPS登录后检查新版资源、同名提示、Word格式提示和弹窗外侧误触/Esc行为，页面无脚本异常。

预发布曾在1GB限制下因内存上限退出，即使报告任务已完成也存在风险。因此网页容器上限改为2GB，Docker Compose默认配置同步更新；PostgreSQL仍为512MB。正式服务通过Docker update平滑调整并在Compose中持久保存。2GB配置下20人复测全部完成，应用峰值约1.12GiB；任务完成后的连续健康检查正常，PostgreSQL保存20条成功报告事件、无失败事件。真实AI速度和长时间运行容量仍需另测。
