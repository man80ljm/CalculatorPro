成绩计算器
一个基于 PyQt6 的课程目标达成度计算工具，支持成绩逆向推算、生成详细成绩报告，并通过 DeepSeek API 生成 AI 驱动的改进分析报告。
功能

成绩处理：根据平时、期中、期末成绩，计算课程目标的加权分数。
灵活配置：通过用户友好的界面设置课程目标数量、权重及成绩占比。
成绩分布：支持正态分布、左偏态分布、右偏态分布和均匀分布。
Excel 输出：生成详细成绩单和目标达成度分析表（Excel 格式）。
AI 分析：调用 DeepSeek API 生成基于课程数据的持续改进报告。
历史数据对比：加载并对比上一学年的达成度数据。

前提条件

Python 3.8 或更高版本
DeepSeek API Key（用于生成 AI 分析报告）

安装

克隆仓库
git clone <repository-url>
cd grade-calculator


创建虚拟环境
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate


安装依赖使用以下国内镜像源加速下载（任选其一）：

清华大学镜像
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple


阿里云镜像
pip install -r requirements.txt -i http://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com

直接安装requirements.txt
pip install -r requirements.txt --index-url https://mirrors.aliyun.com/pypi/simple/

依赖包列于 requirements.txt：
PyQt6==6.5.2
pandas==2.0.3
numpy==1.24.4
openpyxl==3.1.2
requests==2.31.0



使用方法

运行程序
python main.py


输入数据

输入课程名称和目标数量。
设置各目标的权重系数（总和必须为 1）。
设置平时、期中、期末成绩的占比（总和必须为 1）。
选择分数跨度和分布模式。
导入包含学生成绩的 Excel 文件（需包含列：学生姓名、平时成绩、期中成绩、期末成绩、总和）。


配置设置

点击“设置”按钮，输入课程简介、目标要求和 DeepSeek API Key。
可选：导入上一学年达成度分析表（Excel 格式）。


生成报告

点击“导出结果”处理成绩并生成 Excel 报告。
点击“生成 AI 分析报告”生成持续改进报告（需配置 API Key）。



文件结构

main.py：程序入口。
ui.py：PyQt6 实现的图形界面。
core.py：成绩处理和 API 交互的核心逻辑。
utils.py：成绩标准化和 Excel 格式化工具函数。
config.json：存储 API Key 和课程设置（不纳入 git 跟踪）。
calculator.ico：程序图标（需确保路径正确）。

注意事项

确保 calculator.ico 文件位于 D:\calculator\calculator.ico 或项目根目录。
DeepSeek API 调用可能需要 VPN 或代理访问。
导入的成绩单和达成度表需符合指定格式（见“设置”窗口提示）。

打包为可执行文件
使用 PyInstaller 打包程序为单个 EXE 文件：
pyinstaller --onefile --icon=calculator.ico --add-data "calculator.ico;." main.py


确保 calculator.ico 在项目根目录或指定路径。
Windows 用户需将 ; 替换为 : 在 --add-data 参数中。

移除被跟踪的信息
# 移除单个文件
git rm --cached config.json
git rm --cached main.spec

# 移除文件夹及其下的所有文件（使用 -r 表示递归）
git rm -r --cached __pycache__
git rm -r --cached backup
git rm -r --cached dist
git rm -r --cached build


打包指令：
pyinstaller -F -w --icon=calculator.ico --add-data "calculator.ico;." main.py

Linux/MacOS：
pyinstaller -F -w --icon=calculator.ico --add-data "calculator.ico:." main.py