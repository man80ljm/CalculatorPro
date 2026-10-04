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

每位教师的课程信息、占比和对应关系保存在自己的浏览器（localStorage），随请求提交。服务器不保存个人配置，也不接收 API Key。每次上传都会在独立临时目录里计算，响应结束后删除，避免同时使用时互相覆盖 `outputs/`。

### 环境变量

| 变量 | 说明 |
| --- | --- |
| `APP_PASSWORD` | 全站共用登录密码。保护所有页面和接口（`/healthz` 除外）。未设置时无法登录。 |
| `DEEPSEEK_API_KEY` | 服务器调用 `https://api.deepseek.com`、模型 `deepseek-chat` 的密钥。留空时 AI 分析报告会提示未配置，模板、计算和 xlsx/docx 导出仍可用。不要写进前端。 |

示例见 `.env.example`。把真实值写在 `.env`，不要提交。

### 用 Docker 运行

在项目根目录：

```bash
cp .env.example .env
# 编辑 .env，至少填上 APP_PASSWORD
docker compose up -d --build
```

浏览器打开 <http://127.0.0.1:18090> 。端口只绑定在 `127.0.0.1`，内存限制约 1GB，进程退出后自动重启。健康检查地址是 `/healthz`。上传的 Excel 不超过 10MB。

### 本地直接运行

```bash
pip install -r requirements-web.txt
```

Windows PowerShell：

```powershell
$env:APP_PASSWORD="请改成密码"
$env:DEEPSEEK_API_KEY=""   # 可选
uvicorn web_app.app:app --host 127.0.0.1 --port 18090
```

Linux / macOS：

```bash
export APP_PASSWORD="请改成密码"
export DEEPSEEK_API_KEY=""
uvicorn web_app.app:app --host 127.0.0.1 --port 18090
```

### 页面上的流程

1. 用 `APP_PASSWORD` 登录。
2. 填写开课信息、课程基本信息、成绩占比、考核与课程目标对应关系、毕业要求对应关系，以及课程简介和各目标要求。这些内容保存在本机浏览器。
3. 选择正向或逆向模式。逆向模式可设置分数跨度、分布模式和噪声注入。
4. 下载成绩模板，填好后导入。
5. 「计算达成度」查看人数、平均分和各目标达成度。
6. 「导出 xlsx / docx」下载统计表压缩包（成绩明细、表 1 至表 5）。
7. 「生成 AI 分析报告」在密钥已配置时下载表 6 和拼接后的总报告。上一学年达成度表可以随请求一起上传，不上传则按 0 对比。

主要接口（均需登录 Cookie）：

- `POST /api/template`：下载正向或逆向模板
- `POST /api/calculate`：上传成绩并计算
- `POST /api/export`：导出 xlsx/docx 压缩包
- `POST /api/ai-report`：生成 AI 报告压缩包
- `GET /api/ai-status`：是否已配置密钥（不返回密钥本身）
- `GET /healthz`：健康检查，无需登录

登录和 AI 接口有简单的频率限制。

### 网页版未覆盖的桌面细节

- API Key 仍只在桌面「设置」里填写和测试连接；网页不提供密钥输入框。
- 关系表在网页上是表单，不是桌面那个可从 Excel 粘贴的表格弹窗。
- 桌面打包（PyInstaller）和本机 `outputs/` 目录只属于桌面程序。
