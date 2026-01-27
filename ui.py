import sys

import os

import json

import pandas as pd

from PyQt6.QtGui import QIcon, QDoubleValidator, QIntValidator  # 导入 QIcon 类

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 

                            QHBoxLayout, QLabel, QLineEdit, QPushButton, 

                            QFileDialog, QMessageBox, QGridLayout, QDialog, 

                            QTextEdit, QComboBox, QDialogButtonBox,QProgressBar)

from PyQt6.QtGui import QDoubleValidator

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer

from core import GradeProcessor

from relation_table import RelationTableSetupDialog, RelationTableEditorDialog



class GenerateReportThread(QThread):

    """用于在后台线程中生成AI分析报告"""

    finished = pyqtSignal()

    error = pyqtSignal(str)

    progress = pyqtSignal(str)

    progress_value = pyqtSignal(int)  # 新增进度值信号



    def __init__(self, processor, num_objectives, current_achievement, report_style):

        super().__init__()

        self.processor = processor

        self.num_objectives = num_objectives

        self.current_achievement = current_achievement

        self.report_style = report_style  # 新增报告风格参数



    def run(self):

        try:

            questions = ["针对上一年度存在问题的改进情况"]

            for i in range(1, 6):

                questions.append(f"课程目标{i}达成情况分析")

                questions.append(f"该课程目标{i}达成情况存在问题分析及改进措施")

            total_questions = len(questions)

            self.progress.emit("正在生成AI分析报告...")

            self.progress_value.emit(0)



            context = f"课程简介: {self.processor.course_description}\n"

            for i, req in enumerate(self.processor.objective_requirements, 1):

                context += f"课程目标{i}要求: {req}\n"

            for i in range(1, 6):

                prev_score = self.processor.previous_achievement_data.get(f'课程目标{i}', 0)

                current_score = self.current_achievement.get(f'课程目标{i}', 0)

                context += f"课程目标{i}上一年度达成度: {prev_score}\n"

                context += f"课程目标{i}本年度达成度: {current_score}\n"

            prev_total = self.processor.previous_achievement_data.get('课程总目标', 0)

            current_total = self.current_achievement.get('总达成度', 0)

            context += f"课程总目标上一年度达成度: {prev_total}\n"

            context += f"课程总目标本年度达成度: {current_total}\n"



            answers = []

            course_name = self.processor.course_name_input.text()

            for i, question in enumerate(questions):

                self.progress.emit(f"正在处理第 {i+1}/{total_questions} 个问题...")

                self.progress_value.emit(i + 1)

                QApplication.processEvents()

                if "课程目标" in question and int(question.split('课程目标')[1][0]) > self.num_objectives:

                    answers.append("无")

                    continue

                # 将报告风格添加到提示中

                prompt = f"{context}\n问题: {question}\n请以{self.report_style}的风格回答。"

                answer = self.processor.call_deepseek_api(prompt)

                answers.append(answer)



            self.processor.generate_improvement_report(self.current_achievement, course_name, self.num_objectives, answers=answers)

            self.progress_value.emit(total_questions)

            self.finished.emit()

        except Exception as e:

            self.error.emit(f"生成AI分析报告失败：{str(e)}")



class TestApiThread(QThread):

    """用于异步测试 DeepSeek API 连接的线程"""

    result = pyqtSignal(str)  # 信号，用于传递测试结果



    def __init__(self, processor, api_key):

        super().__init__()

        self.processor = processor

        self.api_key = api_key



    def run(self):

        # 在线程中执行 API 连接测试

        result = self.processor.test_deepseek_api(self.api_key)

        self.result.emit(result)







class SettingsWindow(QDialog):

    def __init__(self, parent=None):

        super().__init__(parent)

        self.setWindowTitle('设置')

        self.setFixedWidth(500)

        self.initUI()

    

    def initUI(self):

        layout = QVBoxLayout()

        layout.setSpacing(10)

        layout.setContentsMargins(10, 10, 10, 10)

        

        # 课程简介

        self.description_label = QLabel('课程简介：')

        self.description_input = QTextEdit()

        self.description_input.setFixedHeight(80)

        # 加载已保存的课程简介

        self.description_input.setText(self.parent().course_description)

        layout.addWidget(self.description_label)

        layout.addWidget(self.description_input)

        

        # 课程目标要求（动态）

        self.objectives_layout = QVBoxLayout()

        self.objective_inputs = []

        layout.addLayout(self.objectives_layout)

        

        # 导入上一年达成度表

        self.import_prev_btn = QPushButton('导入上一学年达成度表')

        self.import_prev_btn.clicked.connect(self.import_previous_achievement)

        layout.addWidget(self.import_prev_btn)

        

        # 显示导入文件的路径

        self.file_path_label = QLabel('')

        self.file_path_label.setStyleSheet('font-size: 12px; color: #666666;')

        layout.addWidget(self.file_path_label)

        # 加载已保存的路径

        if self.parent().previous_achievement_file:

            self.file_path_label.setText(self.parent().previous_achievement_file)

        

        # API Key 和检测按钮

        api_layout = QHBoxLayout()

        self.api_key_label = QLabel('API KEY:')

        self.api_key_input = QLineEdit()

        # 加载已保存的 API Key

        self.api_key_input.setText(self.parent().api_key)

        self.test_api_btn = QPushButton('检测')

        self.test_api_btn.clicked.connect(self.test_api_connection)

        api_layout.addWidget(self.api_key_label)

        api_layout.addWidget(self.api_key_input)

        api_layout.addWidget(self.test_api_btn)

        layout.addLayout(api_layout)

        

        # 保存和清空按钮

        button_layout = QHBoxLayout()

        self.save_btn = QPushButton('保存')

        self.save_btn.clicked.connect(self.save_settings)

        self.clear_btn = QPushButton('清空')

        self.clear_btn.clicked.connect(self.clear_settings)

        button_layout.addWidget(self.save_btn)

        button_layout.addWidget(self.clear_btn)

        layout.addLayout(button_layout)

        

        self.setLayout(layout)

    

    def update_objective_inputs(self, num_objectives):

        # 清除现有输入框

        for input_field in self.objective_inputs:

            input_field.setParent(None)

        self.objective_inputs.clear()

        while self.objectives_layout.count():

            item = self.objectives_layout.takeAt(0)

            if item.widget():

                item.widget().deleteLater()

        

        # 生成新输入框

        parent_objective_requirements = self.parent().objective_requirements

        for i in range(num_objectives):

            label = QLabel(f'课程目标{i+1}要求')

            input_field = QLineEdit()

            input_field.setPlaceholderText(f'请输入目标{i+1}要求')

            # 加载已保存的目标要求

            if i < len(parent_objective_requirements):

                input_field.setText(parent_objective_requirements[i])

            self.objective_inputs.append(input_field)

            self.objectives_layout.addWidget(label)

            self.objectives_layout.addWidget(input_field)

        

        self.adjustSize()  # 自适应高度

    

    def import_previous_achievement(self):

        # 提示用户文件格式要求

        QMessageBox.information(

            self,

            '提示',

            '请确保导入的上一学年达成度表包含以下列：\n'

            '- "课程目标"（如：课程目标1, 课程目标2, ..., 课程总目标）\n'

            '- "上一年度达成度"（达成度数值）\n'

            '或者直接导入程序生成的“课程目标达成度分析表.xlsx”文件。'

        )

        file_name, _ = QFileDialog.getOpenFileName(self, "选择上一学年达成度表", "", "Excel Files (*.xlsx)")

        if file_name:

            # 验证文件格式

            try:

                df = pd.read_excel(file_name)

                # 检查是否是程序生成的达成度分析表（包含“考核环节”列）

                if '考核环节' in df.columns:

                    # 进一步验证“课程分目标达成度\n(M)”和“课程总目标达成度”行是否存在

                    if not ('课程分目标达成度\n(M)' in df['考核环节'].values and '课程总目标达成度' in df['考核环节'].values):

                        QMessageBox.warning(self, '错误', '文件格式错误：不是有效的程序生成的“课程目标达成度分析表.xlsx”文件。')

                        return

                else:

                    # 检查是否包含必需列

                    required_columns = ['课程目标', '上一年度达成度']

                    missing_columns = [col for col in required_columns if col not in df.columns]

                    if missing_columns:

                        QMessageBox.warning(self, '错误', f"文件缺少以下必需列：{', '.join(missing_columns)}。请确保文件包含：{', '.join(required_columns)}。")

                        return

                

                # 如果验证通过，保存文件路径

                self.parent().previous_achievement_file = file_name

                self.file_path_label.setText(file_name)

                QMessageBox.information(self, '成功', f'已选择文件: {os.path.basename(file_name)}')

            except Exception as e:

                QMessageBox.warning(self, '错误', f'无法读取文件：{str(e)}')

                self.file_path_label.clear()

                self.parent().previous_achievement_file = ""

    

    def save_settings(self):

        """保存设置到配置文件"""

        self.parent().course_description = self.description_input.toPlainText()

        self.parent().objective_requirements = [input_field.text() for input_field in self.objective_inputs]

        self.parent().api_key = self.api_key_input.text()

        self.parent().save_config()

        QMessageBox.information(self, '成功', '设置已保存')

    

    def clear_settings(self):

        """清空设置"""

        self.description_input.clear()

        for input_field in self.objective_inputs:

            input_field.clear()

        self.file_path_label.clear()

        # 清空父对象中的数据（不包括 API Key）

        self.parent().course_description = ""

        self.parent().objective_requirements = []

        self.parent().previous_achievement_file = ""



        # 保存清空后的配置到 config.json（API Key 保持不变）

        self.parent().save_config()



        QMessageBox.information(self, '成功', '设置已清空')

    

    def test_api_connection(self):

        """测试 DeepSeek API 连接"""

        api_key = self.api_key_input.text().strip()

        if not api_key:

            QMessageBox.warning(self, '警告', '请先输入 API Key')

            return

        

        # 确保 processor 总是被正确初始化

        parent = self.parent()

        if parent.processor is None:

            # 如果 input_file 为空，传入空字符串，因为 test_deepseek_api 不需要 input_file

            input_file = parent.input_file if parent.input_file is not None else ""

            parent.processor = GradeProcessor(

                parent.course_name_input,

                parent.num_objectives_input,

                parent.weight_inputs,

                parent.usual_ratio_input,

                parent.midterm_ratio_input,

                parent.final_ratio_input,

                parent.status_label,

                input_file,

                course_description=parent.course_description,

                objective_requirements=parent.objective_requirements

            )

        

        # 创建一个模态弹窗，显示“连接中”

        dialog = QMessageBox(self)

        dialog.setWindowTitle("测试连接")

        dialog.setText("连接中...")

        # 设置最小宽度

        dialog.setMinimumWidth(400)  # 设置宽度为 300 像素，可以根据需要调整

        # 设置样式以确保文本和按钮居中

        dialog.setStyleSheet("""

            QMessageBox {

                min-width: 300px;

            }

            QMessageBox QLabel {

                padding: 10px;

                text-align: center;

            }

            QMessageBox QDialogButtonBox {

                alignment: center;

            }

            QMessageBox QPushButton {

                min-width: 80px;

            }

        """)

        dialog.show()

        

        # 创建线程来测试 API 连接

        self.test_thread = TestApiThread(parent.processor, api_key)

        self.test_thread.result.connect(lambda result: self.on_test_api_finished(dialog, result))

        self.test_thread.start()

    

    def on_test_api_finished(self, dialog, result):

        """处理 API 连接测试结果"""

        if result == "连接成功":

            dialog.setText("连接成功")

        else:

            dialog.setText(f"连接失败：{result}")

        

        # 确保文本居中

        label = dialog.findChild(QLabel, "qt_msgbox_label")

        if label:

            label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        

        # 添加“确定”按钮，允许用户关闭弹窗

        dialog.setStandardButtons(QMessageBox.StandardButton.Ok)

        button_box = dialog.findChild(QDialogButtonBox)

        if button_box:

            button_box.setCenterButtons(True)

        dialog.exec()



class GradeAnalysisApp(QMainWindow):

    def __init__(self):

        super().__init__()

        self.input_file = None

        self.previous_achievement_file = None

        self.course_description = ""

        self.objective_requirements = []

        self.api_key = ""

        self.num_objectives = 0

        self.processor = None

        self.current_achievement = {}

        self.forward_input_file = None

        self.reverse_input_file = None

        # 加载保存的配置

        self.load_config()

        self.initUI()

    

    def load_config(self):

        """加载配置文件中的 API Key 和课程设置，启动时创建 config.json"""

        # 使用用户的 AppData 目录存储 config.json

        config_dir = os.path.join(os.getenv('APPDATA') or os.path.expanduser('~'), 'CalculatorApp')

        if not os.path.exists(config_dir):

            os.makedirs(config_dir)  # 创建目录

        config_file = os.path.join(config_dir, 'config.json')



        # 如果 config.json 不存在，创建并初始化

        if not os.path.exists(config_file):

            # 初始化所有字段为空

            config = {

                'api_key': '',

                'course_description': '',

                'objective_requirements': [],

                'previous_achievement_file': ''

            }

            try:

                with open(config_file, 'w', encoding='utf-8') as f:

                    json.dump(config, f, indent=4)

            except Exception as e:

                print(f"创建配置文件失败: {str(e)}")

        else:

            # 如果文件存在，加载配置

            try:

                with open(config_file, 'r', encoding='utf-8') as f:

                    config = json.load(f)

                    self.api_key = config.get('api_key', '')

                    self.course_description = config.get('course_description', '')

                    self.objective_requirements = config.get('objective_requirements', [])

                    self.previous_achievement_file = config.get('previous_achievement_file', '')

            except Exception as e:

                print(f"加载配置文件失败: {str(e)}")

    

    def save_config(self):

        """保存 API Key 和课程设置到配置文件"""

        # 使用用户的 AppData 目录存储 config.json

        config_dir = os.path.join(os.getenv('APPDATA') or os.path.expanduser('~'), 'CalculatorApp')

        if not os.path.exists(config_dir):

            os.makedirs(config_dir)  # 创建目录

        config_file = os.path.join(config_dir, 'config.json')



        config = {

            'api_key': self.api_key,

            'course_description': self.course_description,

            'objective_requirements': self.objective_requirements,

            'previous_achievement_file': self.previous_achievement_file

        }

        try:

            with open(config_file, 'w', encoding='utf-8') as f:

                json.dump(config, f, indent=4)

        except Exception as e:

            QMessageBox.critical(self, '错误', f"无法保存配置文件 {config_file}：{str(e)}")

            print(f"保存配置文件失败: {str(e)}")

    

    def initUI(self):

        import os

        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), 'calculator.ico')))

        self.setWindowTitle('Scores Calculator')

        # 移除 setMinimumSize 的高度限制，仅设置宽度

        self.setMinimumWidth(800)

        self.setStyleSheet("""

            QMainWindow {

                background-color: #E6E6E6;

            }

            QLabel {

                font-size: 14px;

                font-weight: bold;

            }

            QLineEdit {

                padding: 5px;

                background-color: white;

                border: 1px solid #CCCCCC;

                border-radius: 3px;

                min-height: 25px;

            }

            QTextEdit {

                padding: 5px;

                background-color: white;

                border: 1px solid #CCCCCC;

                border-radius: 3px;

            }

            QPushButton {

                padding: 8px 15px;

                background-color: #808080;

                color: white;

                border: none;

                border-radius: 5px;

                font-weight: bold;

                min-height: 35px;

            }

            QPushButton:hover {

                background-color: #666666;

            }

            QComboBox {

                padding: 5px;

                background-color: white;

                border: 1px solid #CCCCCC;

                border-radius: 3px;

                min-height: 25px;

            }

        """)

        

        central_widget = QWidget()

        self.setCentralWidget(central_widget)



        layout = QVBoxLayout(central_widget)

        layout.setSpacing(10)  # 减少间距

        layout.setContentsMargins(20, 20, 20, 20)  # 减少边距

        

        title_label = QLabel('课程目标达成度评价计算')

        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title_label.setStyleSheet("""

            font-size: 24px;

            color: black;

            padding: 10px;

        """)

        layout.addWidget(title_label)

        

        course_layout = QHBoxLayout()

        course_label = QLabel('课程名称:')

        self.course_name_input = QLineEdit()

        course_layout.addWidget(course_label)

        course_layout.addWidget(self.course_name_input)

        course_layout.addStretch()

        layout.addLayout(course_layout)

        

        objectives_layout = QVBoxLayout()

        num_label = QLabel('课程目标数量')

        self.num_objectives_input = QLineEdit()

        self.num_objectives_input.setFixedWidth(150)

        # 添加整数验证器，限制输入为数字

        int_validator = QIntValidator(0, 15, self)

        self.num_objectives_input.setValidator(int_validator)

        num_layout = QHBoxLayout()

        num_layout.addWidget(num_label)

        num_layout.addWidget(self.num_objectives_input)

        num_layout.addStretch()

        objectives_layout.addLayout(num_layout)

        

        weights_label = QLabel('课程目标权重系数(总和为1)')

        objectives_layout.addWidget(weights_label)

        

        self.weights_container = QGridLayout()

        self.weights_container.setSpacing(10)

        self.weights_container.setVerticalSpacing(10)

        self.weight_inputs = []

        objectives_layout.addLayout(self.weights_container)

        

        layout.addLayout(objectives_layout)

        

        ratios_layout = QHBoxLayout()

        ratios_layout.setSpacing(20)

        usual_layout = QVBoxLayout()

        usual_label = QLabel('平时成绩占比')

        self.usual_ratio_input = QLineEdit()

        # 添加浮点数验证器，限制 0-1

        ratio_validator = QDoubleValidator(0.0, 1.0, 2, self)

        self.usual_ratio_input.setValidator(ratio_validator)

        self.usual_ratio_input.textEdited.connect(self.validate_ratio_input)

        usual_layout.addWidget(usual_label)

        usual_layout.addWidget(self.usual_ratio_input)

        ratios_layout.addLayout(usual_layout)

        ratios_layout.addStretch()

        midterm_layout = QVBoxLayout()

        midterm_label = QLabel('期中成绩占比')

        self.midterm_ratio_input = QLineEdit()

        midterm_layout.addWidget(midterm_label)

        midterm_layout.addWidget(self.midterm_ratio_input)

        ratios_layout.addLayout(midterm_layout)

        ratios_layout.addStretch()

        final_layout = QVBoxLayout()

        final_label = QLabel('期末成绩占比')

        self.final_ratio_input = QLineEdit()

        final_layout.addWidget(final_label)

        final_layout.addWidget(self.final_ratio_input)

        ratios_layout.addLayout(final_layout)

        layout.addLayout(ratios_layout)



        relation_layout = QHBoxLayout()

        self.relation_btn = QPushButton('课程考核与课程目标对应关系表')

        self.relation_btn.clicked.connect(self.open_relation_table)

        relation_layout.addWidget(self.relation_btn)

        relation_layout.addStretch()

        layout.addLayout(relation_layout)



        mode_entry_layout = QHBoxLayout()

        self.forward_btn = QPushButton('正向模式')

        self.reverse_btn = QPushButton('逆向模式')

        self.forward_btn.clicked.connect(self.open_forward_mode)

        self.reverse_btn.clicked.connect(self.open_reverse_mode)

        mode_entry_layout.addWidget(self.forward_btn)

        mode_entry_layout.addWidget(self.reverse_btn)

        mode_entry_layout.addStretch()

        layout.addLayout(mode_entry_layout)

        

        mode_layout = QHBoxLayout()

        mode_layout.setSpacing(20)  # 增加间距

        spread_label = QLabel('分数跨度:')

        self.spread_combo = QComboBox()

        self.spread_combo.addItems(['大跨度 (12-23分)', '中跨度 (7-13分)', '小跨度 (2-8分)'])

        dist_label = QLabel('分布模式:')

        self.dist_combo = QComboBox()

        self.dist_combo.addItems(['正态分布', '左偏态分布', '右偏态分布', '均匀分布'])

        style_label = QLabel('报告风格:')  # 新增报告风格下拉框

        self.style_combo = QComboBox()

        self.style_combo.addItems(['专业', '口语', '简洁', '详细', '幽默'])

        mode_layout.addWidget(spread_label)

        mode_layout.addWidget(self.spread_combo)

        mode_layout.addWidget(dist_label)

        mode_layout.addWidget(self.dist_combo)

        mode_layout.addWidget(style_label)

        mode_layout.addWidget(self.style_combo)

        mode_layout.addStretch()

        layout.addLayout(mode_layout)



        layout.addSpacing(5)  # 减少额外间距

        

        self.buttons_layout = QHBoxLayout()

        self.buttons_layout.setSpacing(10)

        self.import_btn = QPushButton('导入文件')

        self.settings_btn = QPushButton('设置')

        self.export_btn = QPushButton('导出结果')

        self.ai_report_btn = QPushButton('生成AI分析报告')

        self.buttons_layout.addWidget(self.import_btn)

        self.buttons_layout.addWidget(self.settings_btn)

        self.buttons_layout.addWidget(self.export_btn)

        self.buttons_layout.addWidget(self.ai_report_btn)

        layout.addLayout(self.buttons_layout)

        

        self.status_label = QLabel('')

        self.status_label.setStyleSheet('color: #666666;')

        layout.addWidget(self.status_label)



        self.progress_bar = QProgressBar()

        self.progress_bar.setVisible(False)

        self.progress_bar.setTextVisible(True)

        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.progress_bar.setStyleSheet("""

            QProgressBar {

                border: 1px solid #CCCCCC;

                border-radius: 5px;

                text-align: center;

            }

            QProgressBar::chunk {

                background-color: #808080;

            }

        """)

        layout.addWidget(self.progress_bar)



        # 移除 stretch，完全依赖 adjust_window_height 控制高度

        

        self.num_objectives_input.textChanged.connect(self.update_weight_inputs)

        self.num_objectives_input.textChanged.connect(self.update_num_objectives)

        self.import_btn.clicked.connect(self.select_file)

        self.settings_btn.clicked.connect(self.open_settings_window)

        self.export_btn.clicked.connect(self.start_analysis)

        self.ai_report_btn.clicked.connect(self.start_generate_ai_report)

        

        self.usual_ratio_input.setText('0.2')

        self.midterm_ratio_input.setText('0.3')

        self.final_ratio_input.setText('0.5')



        self.usual_ratio_input.setFixedWidth(150)

        self.midterm_ratio_input.setFixedWidth(150)

        self.final_ratio_input.setFixedWidth(150)

        

        self.setTabOrder(self.num_objectives_input, self.usual_ratio_input)

        self.setTabOrder(self.usual_ratio_input, self.midterm_ratio_input)

        self.setTabOrder(self.midterm_ratio_input, self.final_ratio_input)

        

        self.adjust_button_widths()

        self.adjust_window_height()  # 初始调整高度

    

    def validate_ratio_input(self):

        """实时验证成绩占比输入，限制 0-1 且必须是数字"""

        sender = self.sender()  # 获取发出信号的输入框

        text = sender.text()

        

        try:

            value = float(text)

            if value < 0 or value > 1:

                sender.setText("")  # 清空非法输入

        except ValueError:

            sender.setText("")  # 清空非数字输入



    def adjust_window_height(self):

        """动态调整窗口高度以适应权重输入框和进度条"""

        margin = 20

        spacing = 10

        global_margin = 10



        try:

            num_objectives = int(self.num_objectives_input.text()) if self.num_objectives_input.text() else 0

        except ValueError:

            num_objectives = 0



        columns_per_row = 5

        weight_rows = (num_objectives + columns_per_row - 1) // columns_per_row if num_objectives > 0 else 0

        weight_input_height = 30

        weight_label_height = 20

        weight_row_spacing = 10

        weight_margin = 5



        weight_total_height = (

            weight_label_height +

            weight_rows * weight_input_height +

            (weight_rows - 1) * weight_row_spacing +

            weight_margin

        ) if weight_rows > 0 else 0



        title_height = 40

        course_name_height = 30

        num_objectives_height = 30

        ratios_height = 80

        mode_height = 30

        buttons_height = 40

        status_height = 20

        progress_bar_height = 25 if self.progress_bar.isVisible() else 0



        num_spacings = 8

        total_spacing = num_spacings * spacing

        extra_spacing = 5



        total_height = (

            margin * 2 +

            title_height +

            course_name_height +

            num_objectives_height +

            weight_label_height +

            weight_total_height +

            ratios_height +

            mode_height +

            buttons_height +

            status_height +

            progress_bar_height +

            total_spacing +

            extra_spacing +

            global_margin

        )



        print(f"Calculated window height: {total_height}, weight_rows: {weight_rows}, progress_bar_visible: {self.progress_bar.isVisible()}")



        # screen_height = QApplication.primaryScreen().availableGeometry().height()

        # max_height = int(screen_height * 0.8)

        # calculated_height = min(int(total_height), max_height)

        

        # 移除 max_height 限制，确保窗口高度完全适应内容

        final_height = int(total_height)

        # 动态设置最小高度

        min_height = 400 if num_objectives == 0 else 450

        final_height = max(final_height, min_height)



        # 移除 setMinimumHeight 限制，直接调整高度

        self.resize(self.width(), final_height)



    def resizeEvent(self, event):

        """重写resizeEvent以实现响应式布局"""

        super().resizeEvent(event)

        self.adjust_button_widths()



    def adjust_button_widths(self):

        """调整按钮宽度以实现平均分布"""

        window_width = self.width() - 60

        button_count = 4

        spacing = 10

        total_spacing = spacing * (button_count - 1)

        button_width = (window_width - total_spacing) // button_count

        

        self.import_btn.setMinimumWidth(button_width)

        self.settings_btn.setMinimumWidth(button_width)

        self.export_btn.setMinimumWidth(button_width)

        self.ai_report_btn.setMinimumWidth(button_width)



    def update_weight_inputs(self):

        for widget in self.weight_inputs:

            widget.setParent(None)

            widget.deleteLater()

        self.weight_inputs.clear()

        

        while self.weights_container.count():

            item = self.weights_container.takeAt(0)

            if item.widget():

                item.widget().deleteLater()

        

        try:

            num_objectives = int(self.num_objectives_input.text())

            if num_objectives <= 0 or num_objectives > 15:

                self.adjust_window_height()  # 直接调整高度

                return

                

            columns_per_row = 5

            validator = QDoubleValidator(0.0, 1.0, 2)

            for i in range(num_objectives):

                weight_input = QLineEdit()

                weight_input.setFixedWidth(80)

                weight_input.setPlaceholderText(f'权重{i+1}')

                weight_input.setValidator(validator)

                weight_input.textEdited.connect(self.validate_weights_sum)

                self.weight_inputs.append(weight_input)

                row = i // columns_per_row

                col = i % columns_per_row

                self.weights_container.addWidget(weight_input, row, col)

            

            if self.weight_inputs:

                self.setTabOrder(self.num_objectives_input, self.weight_inputs[0])

                for i in range(len(self.weight_inputs) - 1):

                    self.setTabOrder(self.weight_inputs[i], self.weight_inputs[i + 1])

                self.setTabOrder(self.weight_inputs[-1], self.usual_ratio_input)

        except ValueError:

            # 如果输入不是数字，清空输入框

            self.num_objectives_input.setText("")

        

        self.adjust_window_height()  # 直接调整高度，移除防抖



    def validate_weights_sum(self):

        """实时校验权重总和，仅在超过1时提示，并检查未填框"""

        try:

            weights = []

            empty_fields = False

            for input_field in self.weight_inputs:

                text = input_field.text()

                if text:

                    weight = float(text)

                    weights.append(weight)

                else:

                    weights.append(0.0)

                    empty_fields = True

            

            total = sum(weights)

            if total > 1.0:

                QMessageBox.warning(self, '警告', '总权重系数为1')

            

            if abs(total - 1.0) < 0.0001 and empty_fields:

                QMessageBox.warning(self, '警告', '请填写所有权重系数，未填写的请输入0')

        except ValueError:

            pass



    def check_empty_fields(self):

        """检查是否有未填框"""

        empty_fields = False

        for input_field in self.weight_inputs:

            if not input_field.text():

                empty_fields = True

                break

        return empty_fields



    def update_num_objectives(self):

        try:

            self.num_objectives = int(self.num_objectives_input.text())

            if hasattr(self, 'settings_window'):

                self.settings_window.update_objective_inputs(self.num_objectives)

        except ValueError:

            self.num_objectives = 0

    

    def select_file(self):

        file_name, _ = QFileDialog.getOpenFileName(

            self,

            "选择成绩单文件",

            "",

            "Excel Files (*.xlsx)"

        )

        if file_name:

            self.input_file = file_name

            self.status_label.setText(f"已选择文件: {os.path.basename(file_name)}")

    

    def open_settings_window(self):

        self.settings_window = SettingsWindow(self)

        self.settings_window.update_objective_inputs(self.num_objectives)

        self.settings_window.exec()

        # 保存数据

        self.course_description = self.settings_window.description_input.toPlainText()

        self.objective_requirements = [input_field.text() for input_field in self.settings_window.objective_inputs]

        self.api_key = self.settings_window.api_key_input.text()

        # 保存 API Key 到配置文件

        self.save_config()

    

    def open_forward_mode(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "选择正向模式成绩模板",
            "",
            "Excel Files (*.xlsx)"
        )
        if file_name:
            self.forward_input_file = file_name
            self.status_label.setText(f"正向模式模板已选择: {os.path.basename(file_name)}")


    def open_reverse_mode(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "选择逆向模式成绩模板",
            "",
            "Excel Files (*.xlsx)"
        )
        if file_name:
            self.reverse_input_file = file_name
            self.status_label.setText(f"逆向模式模板已选择: {os.path.basename(file_name)}")


    def open_relation_table(self):

        setup = RelationTableSetupDialog(self, default_objectives=self.num_objectives)

        if setup.exec():

            values = setup.result_values

            if values:

                try:

                    usual = float(self.usual_ratio_input.text().strip())

                    midterm = float(self.midterm_ratio_input.text().strip())

                    final = float(self.final_ratio_input.text().strip())

                except Exception:
                    QMessageBox.warning(self, '提示', '请先填写平时/期中/期末占比（小数形式）')
                    return
                if abs((usual + midterm + final) - 1.0) != 0:
                    QMessageBox.warning(self, '提示', '平时/期中/期末占比之和必须等于1')
                    return
                counts = values[1:]
                ratios = [usual, midterm, final]
                names = ['平时考核', '期中考核', '期末考核']
                for idx, count in enumerate(counts):
                    if count == 0 and ratios[idx] > 0:
                        QMessageBox.warning(self, '提示', f'{names[idx]}数量为0时，占比必须为0')
                        return
                dialog = RelationTableEditorDialog(self, *values)

                dialog.exec()



    def validate_inputs(self):

        try:

            # 检查课程名称是否为空

            if not self.course_name_input.text():

                raise ValueError("请输入课程名称")

            

            # 检查课程目标数量

            num_objectives_text = self.num_objectives_input.text().strip()

            if not num_objectives_text:

                raise ValueError("课程目标数量不能为空")

            try:

                num_objectives = int(num_objectives_text)

            except ValueError:

                raise ValueError("课程目标数量必须为数字")

            if num_objectives <= 0 or num_objectives > 15:

                raise ValueError("课程目标数量必须在 1 到 15 之间")

            

            # 检查权重系数

            weights = []

            for i, input_field in enumerate(self.weight_inputs, 1):

                text = input_field.text().strip()

                if not text:

                    raise ValueError(f"权重{i}不能为空，请输入 0 或有效数字")

                try:

                    weight = float(text)

                except ValueError:

                    raise ValueError(f"权重{i}必须为数字")

                if weight < 0 or weight > 1:

                    raise ValueError(f"权重{i}必须在 0 到 1 之间")

                weights.append(weight)

            

            # 检查权重总和

            if abs(sum(weights) - 1) > 0.0001:

                raise ValueError("总权重系数必须为 1")

            

            # 检查成绩占比

            usual_ratio_text = self.usual_ratio_input.text().strip()

            midterm_ratio_text = self.midterm_ratio_input.text().strip()

            final_ratio_text = self.final_ratio_input.text().strip()



            if not usual_ratio_text:

                raise ValueError("平时成绩占比不能为空")

            if not midterm_ratio_text:

                raise ValueError("期中成绩占比不能为空")

            if not final_ratio_text:

                raise ValueError("期末成绩占比不能为空")



            try:

                usual_ratio = float(usual_ratio_text)

                midterm_ratio = float(midterm_ratio_text)

                final_ratio = float(final_ratio_text)

            except ValueError:

                raise ValueError("成绩占比必须为数字")



            if any(r < 0 or r > 1 for r in [usual_ratio, midterm_ratio, final_ratio]):

                raise ValueError("成绩占比必须在 0 到 1 之间")



            # 检查成绩占比总和

            if abs(usual_ratio + midterm_ratio + final_ratio - 1) > 0.0001:

                raise ValueError("成绩占比之和必须等于 1")



            return num_objectives, weights, usual_ratio, midterm_ratio, final_ratio



        except ValueError as e:

            QMessageBox.warning(self, '输入错误', str(e))

            return None



    def start_analysis(self):

        if not self.input_file:

            QMessageBox.warning(self, '错误', '请先选择成绩单文件')

            return

        

        # 提前检查目标权重系数总和和成绩占比总和

        weights_sum_error = False

        ratios_sum_error = False



        # 检查目标权重系数总和

        try:

            weights = []

            for input_field in self.weight_inputs:

                text = input_field.text()

                if not text:

                    weights.append(0.0)

                else:

                    weight = float(text)

                    weights.append(weight)

            

            weights_sum = sum(weights)

            if abs(weights_sum - 1) > 0.0001:

                weights_sum_error = True

        except ValueError:

            weights_sum_error = True



        # 检查成绩占比总和

        try:

            usual_ratio_text = self.usual_ratio_input.text()

            midterm_ratio_text = self.midterm_ratio_input.text()

            final_ratio_text = self.final_ratio_input.text()

            

            if not usual_ratio_text or not midterm_ratio_text or not final_ratio_text:

                ratios_sum_error = True

            else:

                usual_ratio = float(usual_ratio_text)

                midterm_ratio = float(midterm_ratio_text)

                final_ratio = float(final_ratio_text)

                ratios_sum = usual_ratio + midterm_ratio + final_ratio

                if abs(ratios_sum - 1) > 0.0001:

                    ratios_sum_error = True

        except ValueError:

            ratios_sum_error = True



        # 根据检查结果显示提示

        if weights_sum_error and ratios_sum_error:

            QMessageBox.warning(self, '输入错误', '目标权重系数总和需为 1，成绩占比总和需为 1')

            return

        elif weights_sum_error:

            QMessageBox.warning(self, '输入错误', '目标权重系数总和需为 1')

            return

        elif ratios_sum_error:

            QMessageBox.warning(self, '输入错误', '成绩占比总和需为 1')

            return



        # 如果总和检查通过，继续进行其他验证

        result = self.validate_inputs()

        if result is None:

            return

            

        num_objectives, weights, usual_ratio, midterm_ratio, final_ratio = result

        

        spread_mode = {

            '大跨度 (12-23分)': 'large',

            '中跨度 (7-13分)': 'medium',

            '小跨度 (2-8分)': 'small'

        }[self.spread_combo.currentText()]

        distribution = {

            '正态分布': 'normal',

            '左偏态分布': 'left_skewed',

            '右偏态分布': 'right_skewed',

            '均匀分布': 'uniform'

        }[self.dist_combo.currentText()]

        

        self.processor = GradeProcessor(

            self.course_name_input,

            self.num_objectives_input,

            self.weight_inputs,

            self.usual_ratio_input,

            self.midterm_ratio_input,

            self.final_ratio_input,

            self.status_label,

            self.input_file,

            course_description=self.course_description,

            objective_requirements=self.objective_requirements

        )

        

         # 加载上一学年达成度表，添加异常处理

        try:

            if self.previous_achievement_file and os.path.exists(self.previous_achievement_file):

                self.processor.load_previous_achievement(self.previous_achievement_file)

            else:

                self.processor.load_previous_achievement(None)

        except ValueError as e:

            QMessageBox.critical(self, '错误', f'加载上一学年达成度表失败：{str(e)}\n请在设置中重新选择正确的文件。')

            self.status_label.setText("加载上一学年达成度表失败！")

            self.progress_bar.setVisible(False)

            self.adjust_window_height()

            return  # 停止执行后续步骤

               

        self.processor.store_api_key(self.api_key)

        

        try:

            # 读取输入文件

            try:

                df = pd.read_excel(self.input_file)

            except FileNotFoundError:

                QMessageBox.critical(self, '错误', '成绩单文件不存在，请重新选择')

                return

            except ValueError as e:

                QMessageBox.critical(self, '错误', f'成绩单文件格式错误：{str(e)}')

                return

            except PermissionError:

                QMessageBox.critical(self, '错误', '无权限读取成绩单文件，请检查文件权限')

                return

            except Exception as e:

                QMessageBox.critical(self, '错误', f'读取成绩单文件失败：{str(e)}')

                return



            self.progress_bar.setVisible(True)

            self.progress_bar.setMaximum(len(df))

            self.status_label.setText("正在处理数据...")



            # 检查课程名称是否包含非法字符

            course_name = self.course_name_input.text().strip()

            invalid_chars = '<>:"/\\|?*'

            if any(char in course_name for char in invalid_chars):

                QMessageBox.critical(self, '错误', f'课程名称包含非法字符：{invalid_chars}')

                return



            # 处理成绩数据

            try:

                overall_achievement = self.processor.process_grades(

                    num_objectives, weights, usual_ratio, midterm_ratio, final_ratio,

                    spread_mode, distribution,

                    progress_callback=lambda idx: self.progress_bar.setValue(idx + 1)

                )

            except PermissionError:

                QMessageBox.critical(self, '错误', '文件被占用，请关闭文件！')

                return

            except Exception as e:

                QMessageBox.critical(self, '错误', f'处理成绩数据失败：{str(e)}')

                return



            # 读取生成的分析表

            analysis_file = f"{os.path.dirname(self.input_file)}/{course_name}课程目标达成度分析表.xlsx"

            try:

                df = pd.read_excel(analysis_file)

                for i in range(1, num_objectives + 1):

                    m_row = df[df['考核环节'] == '课程分目标达成度\n(M)']

                    if not m_row.empty:

                        self.current_achievement[f'课程目标{i}'] = m_row[f'课程目标{i}'].iloc[0] if f'课程目标{i}' in m_row else 0

            except FileNotFoundError:

                QMessageBox.critical(self, '错误', f'分析表文件不存在：{analysis_file}')

                return

            except ValueError as e:

                QMessageBox.critical(self, '错误', f'分析表文件格式错误：{str(e)}')

                return

            except PermissionError:

                QMessageBox.critical(self, '错误', '无权限读取分析表文件，请检查文件权限')

                return

            except Exception as e:

                QMessageBox.critical(self, '错误', f'读取分析表文件失败：{str(e)}')

                return



            for i in range(1, num_objectives + 1):

                m_row = df[df['考核环节'] == '课程分目标达成度\n(M)']

                if not m_row.empty:

                    self.current_achievement[f'课程目标{i}'] = m_row[f'课程目标{i}'].iloc[0] if f'课程目标{i}' in m_row else 0

            self.current_achievement['总达成度'] = overall_achievement



            self.status_label.setText(f"处理完成！课程总目标达成度: {overall_achievement}")

            self.progress_bar.setVisible(False)

            self.adjust_window_height()

            QMessageBox.information(self, '成功', '分析报告已生成')



        except Exception as e:

            self.status_label.setText("处理失败！")

            self.progress_bar.setVisible(False)

            self.adjust_window_height()

            QMessageBox.critical(self, '错误', f'处理过程中发生错误：{str(e)}')



    def start_generate_ai_report(self):

        """在后台线程中生成AI分析报告"""

        if self.processor is None:

            QMessageBox.warning(self, '错误', '请先进行成绩分析')

            return

        

        if not self.course_description or not self.objective_requirements or not self.previous_achievement_file:

            QMessageBox.warning(self, '提示', '请完善设置中的内容')

            return

        

        # 获取用户选择的报告风格

        report_style = self.style_combo.currentText()



        self.report_thread = GenerateReportThread(self.processor, self.num_objectives, self.current_achievement, report_style=report_style)

        self.report_thread.finished.connect(self.on_generate_ai_report_finished)

        self.report_thread.error.connect(self.on_generate_ai_report_error)

        self.report_thread.progress.connect(self.update_status_label)

        self.report_thread.progress_value.connect(self.progress_bar.setValue)  # 新增信号连接

        self.ai_report_btn.setEnabled(False)

        self.progress_bar.setVisible(True)

        self.adjust_window_height()  # 显示进度条时调整高度

        self.progress_bar.setMaximum(11)  # 假设最多 11 个问题（1 + 5*2）

        self.report_thread.start()



    def on_generate_ai_report_finished(self):

        """生成报告完成时的回调"""

        self.ai_report_btn.setEnabled(True)

        self.status_label.setText("AI分析报告已生成")

        self.progress_bar.setVisible(False)

        self.adjust_window_height()  # 隐藏进度条时调整高度

        QMessageBox.information(self, '成功', 'AI分析报告已生成')



    def on_generate_ai_report_error(self, error_message):

        """生成报告失败时的回调"""

        self.ai_report_btn.setEnabled(True)

        self.status_label.setText("生成AI分析报告失败！")

        self.progress_bar.setVisible(False)

        self.adjust_window_height()  # 隐藏进度条时调整高度

        # 区分错误类型，显示不同的提示

        if "请先设置API Key" in error_message:

            QMessageBox.critical(self, '错误', '请填写API KEY')

        elif "API 调用失败" in error_message or "API 返回格式错误" in error_message:

            QMessageBox.critical(self, '错误', 'API KEY无效！')

        elif isinstance(error_message, PermissionError) or "[Errno 13] Permission denied" in str(error_message):

            QMessageBox.critical(self, '错误', '文件被占用，请关闭文件！')  # 优化提示

        else:

            QMessageBox.critical(self, '错误', error_message)



    def update_status_label(self, message):

        """更新状态标签"""

        self.status_label.setText(message)



if __name__ == '__main__':

    app = QApplication(sys.argv)

    window = GradeAnalysisApp()

    window.show()

    sys.exit(app.exec())

