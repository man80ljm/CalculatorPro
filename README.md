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

课程考核与课程目标对应关系是一张可编辑表格：格子可以直接改，也能从 Excel 粘贴（制表符分列、换行分行，兼容 Windows 换行和行尾换行）。可以添加行、删除选中行。粘贴从当前格子开始，行数不够时自动加行。

### 环境变量

| 变量 | 说明 |
| --- | --- |
| `SECRET_KEY` | 签名会话 Cookie 的密钥。未设置时进程直接退出。 |
| `POSTGRES_USER` / `POSTGRES_PASSWORD` / `POSTGRES_DB` | Postgres 账号，供 compose 里的数据库容器使用。 |
| `DATABASE_URL` | 例如 `postgresql+psycopg://calculatorpro:密码@postgres:5432/calculatorpro`。本地试用可写成 `sqlite:///./calculatorpro.db`。 |
| `DEEPSEEK_API_KEY` | 服务器调用 `https://api.deepseek.com`、模型 `deepseek-chat` 的密钥。留空时 AI 分析报告会提示未配置，模板、计算和 xlsx/docx 导出仍可用。 |
| `COOKIE_SECURE` | `true` 时会话 Cookie 带 Secure。只在 HTTPS 反向代理后打开，默认 `false`。 |
| `UPLOAD_DIR` | 课程文件目录。容器里是 `/data/uploads`，对应宿主机 `/data/uploads/calculatorpro`。 |
| `MAX_UPLOAD_MB` | 单个 Excel 上限，默认 10。 |
| `TRUSTED_PROXIES` | 受信反向代理网段（逗号分隔），只有来自这些地址的请求才读取 `X-Forwarded-For` 最右一项作为客户端 IP 用于限流。默认 `127.0.0.1/32,::1/128,172.16.0.0/12`（本机 + Docker 默认网段）。 |

示例见 `.env.example`。把真实值写在 `.env`，不要提交。`.env` 已在 `.gitignore` 中。

### 用 Docker 运行

在项目根目录准备目录和权限。应用容器以 uid `10001` 运行，必须能写入上传目录：

```bash
sudo mkdir -p /data/uploads/calculatorpro /data/databases/calculatorpro/postgres
sudo chown -R 10001:10001 /data/uploads/calculatorpro
cp .env.example .env
# 编辑 .env：SECRET_KEY、数据库口令、DATABASE_URL。DEEPSEEK_API_KEY 可留空。
docker compose up -d --build
```

浏览器打开 <http://127.0.0.1:18090> 。网页端口只绑定在 `127.0.0.1`，应用内存限制约 1GB。Postgres 16 不向宿主机发布端口，内存限制约 512MB，数据在 `/data/databases/calculatorpro/postgres`。应用会等数据库健康检查通过后再启动。`/healthz` 返回 `{"status":"ok"}` 时表示进程和数据库都可用。上传的 Excel 默认不超过 10MB。

### 本地直接运行

```bash
pip install -r requirements-web.txt
```

Windows PowerShell：

```powershell
$env:SECRET_KEY="请换成一长串随机字符"
$env:DATABASE_URL="sqlite:///./calculatorpro.db"
$env:UPLOAD_DIR="./data/uploads"
$env:DEEPSEEK_API_KEY=""
uvicorn web_app.app:app --host 127.0.0.1 --port 18090
```

Linux / macOS：

```bash
export SECRET_KEY="请换成一长串随机字符"
export DATABASE_URL="sqlite:///./calculatorpro.db"
export UPLOAD_DIR="./data/uploads"
export DEEPSEEK_API_KEY=""
uvicorn web_app.app:app --host 127.0.0.1 --port 18090
```

### 页面上的流程

1. 注册或登录自己的账号。退出后该会话立即失效。登录后可以修改密码，需要填写当前密码。
2. 新建课程文件夹，填写开课信息、课程基本信息、考核与课程目标对应关系、毕业要求，以及课程简介和各目标要求，然后保存。
3. 选择正向或逆向模式。逆向模式可设置分数跨度、分布模式和噪声注入。
4. 下载成绩模板，填好后导入。文件记在当前课程下。
5. 「计算达成度」查看人数、平均分和各目标达成度。
6. 「导出 xlsx / docx」下载统计表压缩包（成绩明细、表 1 至表 5），并留在课程文件夹里。
7. 「生成 AI 分析报告」在密钥已配置时下载表 6 和拼接后的总报告。上一学年达成度表可以先导入，不导入则按 0 对比。
8. 之后在任意电脑登录，打开同一课程，即可再次下载这些文件。

主要接口（除注册、登录和健康检查外，都需要登录 Cookie）：

- `POST /api/register`、`POST /api/login`、`POST /api/logout`、`POST /api/password`
- `GET/POST /api/courses`，`GET/PATCH/DELETE /api/courses/{id}`
- `GET/POST /api/courses/{id}/files`，`GET /api/courses/{id}/files/{file_id}`
- `POST /api/courses/{id}/template`：按课程设置生成模板并保存
- `POST /api/courses/{id}/calculate`：用已保存或随请求上传的成绩计算
- `POST /api/courses/{id}/export`：导出 xlsx/docx 压缩包并保存
- `POST /api/courses/{id}/ai-report`：生成 AI 报告压缩包并保存
- `GET /api/ai-status`：是否已配置密钥（不返回密钥本身）
- `GET /healthz`：健康检查，同时确认数据库可连接

登录、注册和 AI 接口有频率限制。登录同时按 IP 和用户名计数。访问别人的课程或文件会得到 404。

### 备份

数据库和上传文件要一起留：

```bash
mkdir -p backups
docker compose exec -T postgres pg_dump -U "$POSTGRES_USER" "$POSTGRES_DB" | gzip > "backups/calculatorpro-$(date +%Y%m%d-%H%M%S).sql.gz"
```

也可以直接运行 `sh scripts/backup_db.sh`（它会读取 `.env`，用 `docker compose exec -T postgres pg_dump ... | gzip` 写到 `backups/`）。同时把宿主机目录 `/data/uploads/calculatorpro` 拷走，里面是各教师的成绩表和导出文件。

### 网页版未覆盖的桌面细节

- API Key 仍只在桌面「设置」里填写和测试连接；网页从环境变量读取，不提供密钥输入框。
- 桌面的关系表仍是原来的 PyQt 弹窗。网页用一张可粘贴的表格完成同一份数据，两边都交给 `core_app` 计算。
- 桌面打包（PyInstaller）和本机 `outputs/` 目录只属于桌面程序。
