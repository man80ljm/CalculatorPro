import sys
import os
import json
import pandas as pd
from PyQt6.QtGui import QIcon, QDoubleValidator, QIntValidator
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                            QHBoxLayout, QLabel, QLineEdit, QPushButton,
                            QFileDialog, QMessageBox, QGridLayout, QDialog,
                            QComboBox, QDialogButtonBox, QProgressBar,
                            QTabWidget, QFrame, QGroupBox, QSlider, QRadioButton,
                            QButtonGroup, QCheckBox, QSizePolicy)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from core import GradeProcessor
from io_app.excel_templates import create_forward_template, create_reverse_template
from relation_table import RelationTableSetupDialog, RelationTableEditorDialog
from ui_app.settings_dialog import SettingsDialog
from ui_app.course_open_dialog import CourseOpenDialog
from ui_app.course_basic_dialog import CourseBasicDialog
from ui_app.grad_req_dialog import GradRequirementDialog

# ==========================================
# 工具类：适配旧核心逻辑的 Mock 输入
# ==========================================

class MockInput:
    """兼容旧核心逻辑的 Mock 输入对象。
    提供 .text() 接口，避免直接依赖 UI 控件。
    """
    def __init__(self, text_value):
        self._text = str(text_value)

    def text(self):
        return self._text

    def setText(self, val):
        self._text = str(val)

class RatioDialog(QDialog):
    def __init__(self, parent=None, usual=0.2, midterm=0.3, final=0.5):
        super().__init__(parent)
        self.setWindowTitle("成绩占比设置")
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "..", "calculator.ico")))
        self.resize(180, 200)
        # QSS 分区：全局 / 主卡片 / Tab与输入 / 顶部按钮 / 控制面板 / 底部操作按钮
        self.setStyleSheet("""
            QLineEdit {
                background: #F8F9FB;
                border: none;
                border-radius: 0px;
                padding: 10px 12px;
                font-size: 13px;
                color: #2C3E50;
            }
            QLineEdit:focus {
                border: 1px solid #BDBDBD;
                background: #FFFFFF;
            }
        """)
        self.layout = QVBoxLayout(self)
        self.usual_input = self.add_row("平时考核:", usual)
        self.midterm_input = self.add_row("期中考核:", midterm)
        self.final_input = self.add_row("期末考核:", final)
        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        self.buttons.setStyleSheet("QPushButton{border-radius:10px; padding:6px 18px;}")
        ok_btn = self.buttons.button(QDialogButtonBox.StandardButton.Ok)
        if ok_btn:
            ok_btn.setMinimumWidth(90)
            ok_btn.setStyleSheet("background:#D9D9D9; border-radius:10px; padding:6px 18px;")
        self.buttons.accepted.connect(self.accept)
        self.layout.addWidget(self.buttons, alignment=Qt.AlignmentFlag.AlignHCenter)

    def add_row(self, label_text, value):
        layout = QHBoxLayout()
        label = QLabel(label_text)
        inp = QLineEdit(str(value))
        inp.setPlaceholderText("0-1")
        inp.setValidator(QDoubleValidator(0.0, 1.0, 2))
        layout.addWidget(label)
        layout.addWidget(inp)
        self.layout.addLayout(layout)
        return inp

    def get_values(self):
        return (self.usual_input.text(), self.midterm_input.text(), self.final_input.text())

# ==========================================
# 弹窗：噪声注入详情
class NoiseConfigDialog(QDialog):
    def __init__(self, parent=None, available_subjects=None):
        super().__init__(parent)
        self.setWindowTitle("噪声注入详情配置")
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "..", "calculator.ico")))
        self.setFixedWidth(450)
        self.available_subjects = available_subjects if available_subjects else ['平时作业', '期末考试']
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(15)
        # 1. 覆盖率设置
        group_rate = QGroupBox("1. 异常覆盖率 (Trigger Rate)")
        rate_layout = QVBoxLayout()
        h_layout = QHBoxLayout()
        self.slider = QSlider(Qt.Orientation.Horizontal)
        self.slider.setRange(0, 100) # 使用 0-100 对应 0.0-1.0 的百分比
        self.slider.setValue(10)     # 默认 0.1 (10%)
        self.label_rate = QLabel("0.10")
        self.label_rate.setStyleSheet("font-weight: bold; color: #333;")
        self.slider.valueChanged.connect(lambda v: self.label_rate.setText(f"{v/100:.2f}"))
        h_layout.addWidget(self.slider)
        h_layout.addWidget(self.label_rate)
        rate_layout.addLayout(h_layout)
        group_rate.setLayout(rate_layout)
        layout.addWidget(group_rate)
        # 2. 挂科严重程度选择
        group_mode = QGroupBox("2. 挂科严重程度 (Severity Mode)")
        mode_layout = QVBoxLayout()
        self.btn_group = QButtonGroup()
        self.rb_random = QRadioButton("随机分布 (40-59分) - [默认]")
        self.rb_near = QRadioButton("边缘挂科 (55-59分) - [模拟惜败]")
        self.rb_catastrophic = QRadioButton("严重缺失 (0-40分) - [模拟缺考]")
        self.rb_random.setChecked(True)
        self.btn_group.addButton(self.rb_random, 1)
        self.btn_group.addButton(self.rb_near, 2)
        self.btn_group.addButton(self.rb_catastrophic, 3)
        mode_layout.addWidget(self.rb_random)
        mode_layout.addWidget(self.rb_near)
        mode_layout.addWidget(self.rb_catastrophic)
        group_mode.setLayout(mode_layout)
        layout.addWidget(group_mode)
        # 3. 目标科目选择
        group_target = QGroupBox("3. 允许挂科的科目 (Target Preference)")
        target_layout = QGridLayout()
        self.checks = {}
        for i, subject in enumerate(self.available_subjects):
            cb = QCheckBox(subject)
            cb.setChecked(True) # 默认全选
            self.checks[subject] = cb
            target_layout.addWidget(cb, i // 2, i % 2) # 网格两列布局
        group_target.setLayout(target_layout)
        layout.addWidget(group_target)
        # 确认/取消按钮
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.setStyleSheet("QPushButton{border-radius:10px; padding:6px 16px;}")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def get_config(self):
        """返回噪声配置字典"""
        mode_map = {1: 'random', 2: 'near_miss', 3: 'catastrophic'}
        selected_mode = mode_map.get(self.btn_group.checkedId(), 'random')
        allowed = [subj for subj, cb in self.checks.items() if cb.isChecked()]
        return {
            "noise_ratio": self.slider.value() / 100.0,
            "severity_mode": selected_mode,
            "allowed_items": allowed
        }

# ==========================================
# 弹窗：模板下载
class TemplateDownloadDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("模板下载")
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "..", "calculator.ico")))
        self.setFixedWidth(300)
        self.student_count = 0
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.addWidget(QLabel("请输入学生人数:"))
        self.input_count = QLineEdit()
        self.input_count.setPlaceholderText("例如: 45")
        self.input_count.setValidator(QIntValidator(1, 500)) # 限制 1-500 人
        layout.addWidget(self.input_count)
        self.btn_download = QPushButton("点击下载")
        self.btn_download.setFixedHeight(35)
        self.btn_download.clicked.connect(self.validate_and_accept)
        layout.addWidget(self.btn_download)

    def validate_and_accept(self):
        text = self.input_count.text()
        if not text:
            QMessageBox.warning(self, "提示", "请输入学生人数")
            return
        self.student_count = int(text)
        self.accept()

# ==========================================
# 线程：AI 报告生成
class GenerateReportThread(QThread):
    """异步生成 AI 课程报告的线程"""
    finished = pyqtSignal()
    error = pyqtSignal(str)
    progress = pyqtSignal(str)
    progress_value = pyqtSignal(int)

    def __init__(self, processor, num_objectives, current_achievement, report_style):
        super().__init__()
        self.processor = processor
        self.num_objectives = num_objectives
        self.current_achievement = current_achievement
        self.report_style = report_style

    def run(self):
        try:
            questions = ["针对上一年度存在问题的改进情况"]
            for i in range(1, 6):
                questions.append(f"课程目标{i}达成情况分析")
                questions.append(f"课程目标{i}存在问题及改进措施")
            total_questions = len(questions)
            self.progress.emit("正在生成AI报告...")
            self.progress_value.emit(0)
            
            # 构造上下文
            context = f"课程简介: {self.processor.course_description}\n"
            for i, req in enumerate(self.processor.objective_requirements, 1):
                context += f"课程目标{i}要求: {req}\n"
            for i in range(1, 6):
                prev_score = self.processor.previous_achievement_data.get(f"课程目标{i}", 0)
                current_score = self.current_achievement.get(f"课程目标{i}", 0)
                context += f"课程目标{i}上一年度达成度: {prev_score}\n"
                context += f"课程目标{i}本年度达成度: {current_score}\n"

            prev_total = self.processor.previous_achievement_data.get("课程总目标", 0)
            current_total = self.current_achievement.get("总达成度", 0)
            context += f"课程总目标上一年度达成度: {prev_total}\n"
            context += f"课程总目标本年度达成度: {current_total}\n"

            answers = []
            course_name = "课程名称"
            try:
                if hasattr(self.processor, "course_name_input"):
                    course_name = self.processor.course_name_input.text() or course_name
            except Exception:
                pass

            for i, question in enumerate(questions):
                self.progress.emit(f"正在生成 {i+1}/{total_questions} 个问题...")
                self.progress_value.emit(i + 1)
                if "课程目标" in question and int(question.split("课程目标")[1][0]) > self.num_objectives:
                    answers.append("无")
                    continue
                prompt = f"{context}\n问题: {question}\n请以{self.report_style}风格回答，语言简洁。"
                answer = self.processor.call_deepseek_api(prompt)
                answers.append(answer)

            self.processor.generate_improvement_report(
                self.current_achievement,
                course_name,
                self.num_objectives,
                answers=answers,
            )
            self.progress_value.emit(total_questions)
            self.finished.emit()
        except Exception as e:
            self.error.emit(f"AI报告生成失败：{str(e)}")

class GradeAnalysisApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.input_file = None
        self.previous_achievement_file = None
        self.course_description = ""
        self.objective_requirements = []
        self.course_open_info = {}
        self.course_basic_info = {}
        self.grad_req_map = []
        self.api_key = ""
        self.num_objectives = 0
        self.processor = None
        self.current_achievement = {}
        self.forward_input_file = None
        self.reverse_input_file = None
        self.weight_data = [] # 存储权重数值
        self.weight_inputs = [] # 存储 MockInput 对象
        self.usual_ratio = 0.2
        self.midterm_ratio = 0.3
        self.final_ratio = 0.5
        self.noise_config = None
        self.relation_payload = None
        self.load_config()
        self.initUI()
        self.apply_styles()

    def load_config(self):
        config_dir = os.path.join(os.getenv('APPDATA') or os.path.expanduser('~'), 'CalculatorApp')
        if not os.path.exists(config_dir):
            os.makedirs(config_dir)
        config_file = os.path.join(config_dir, 'config.json')
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.api_key = config.get('api_key', '')
                    self.course_description = config.get('course_description', '')
                    self.objective_requirements = config.get('objective_requirements', [])
                    self.previous_achievement_file = config.get('previous_achievement_file', '')
                    self.relation_payload = config.get('relation_payload')
                    self.course_open_info = config.get('course_open_info', {})
                    self.course_basic_info = config.get('course_basic_info', {})
                    self.grad_req_map = config.get('grad_req_map', [])
                    if self.relation_payload and not self.num_objectives:
                        self.num_objectives = self.relation_payload.get("objectives_count", 0)
                    if 'ratios' in config:
                        self.usual_ratio = config['ratios'].get('usual', 0.2)
                        self.midterm_ratio = config['ratios'].get('midterm', 0.3)
                        self.final_ratio = config['ratios'].get('final', 0.5)
            except Exception as e:
                print(f"读取配置文件失败: {str(e)}")

    def _get_course_name(self):
        name = ''
        if isinstance(self.course_open_info, dict):
            name = self.course_open_info.get('course_name') or ''
        if not name and isinstance(self.course_basic_info, dict):
            name = self.course_basic_info.get('course_name') or ''
        name = str(name).strip()
        return name if name else '课程名称'

    def save_config(self):
        config_dir = os.path.join(os.getenv('APPDATA') or os.path.expanduser('~'), 'CalculatorApp')
        config_file = os.path.join(config_dir, 'config.json')
        config = {
            'api_key': self.api_key,
            'course_description': self.course_description,
            'objective_requirements': self.objective_requirements,
            'previous_achievement_file': self.previous_achievement_file,
            'relation_payload': self.relation_payload,
            'course_open_info': self.course_open_info,
            'course_basic_info': self.course_basic_info,
            'grad_req_map': self.grad_req_map,
            'ratios': {
                'usual': self.usual_ratio,
                'midterm': self.midterm_ratio,
                'final': self.final_ratio
            }
        }
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            print(f"保存配置文件失败: {str(e)}")

    def apply_styles(self):

        """应用全局 QSS 样式"""
        # QSS 分区：全局 / 主卡片 / Tab与输入 / 顶部按钮 / 控制面板 / 底部操作按钮
        self.setStyleSheet("""
            QMainWindow { background-color: #F5F5F5; }
            QWidget { font-family: "Microsoft YaHei", "SimHei"; font-size: 14px; color: #333333; }
            QTabWidget { background: #E6E6E6; }
            QTabWidget::pane { border: none; margin: 0px; padding: 0px; background: transparent; }
            QTabBar::tab {
                background: #F5F5F5;
                border: none;
                padding: 6px 14px;
                font-weight: bold;
            }
            /* 第一个 tab（正向模式）左上角圆角 */
            QTabBar::tab:first {
                border-top-left-radius: 10px;
            }

            /* 最后一个 tab（逆向模式）右上角圆角 */
            QTabBar::tab:last {
                border-top-right-radius: 10px;
            }
            /* 选中态背景 */
            QTabBar::tab:selected { background: #FFFFFF; }

                        #tab_forward, #tab_reverse { background: transparent; }

            QLineEdit, QComboBox {
                background-color: #FFFFFF;
                border: none;
                border-radius: 0px;
                padding: 6px 10px;
                min-height: 20px;
            }
            QComboBox::drop-down { border: none; }

            QPushButton#TopButton {
                background-color: #C4C4C4;
                color: #333333;
                border-radius: 12px;
                padding: 8px 20px;
                font-weight: bold;
                border: none;
            }

            #MainCard {
                background: #E6E6E6;
                border-bottom-left-radius: 0px;
                border-bottom-right-radius: 0px;
            }

            #ControlBar {
                background: #CFCFCF;
                border-radius: 0px;
            }

            #TabPanel {
                background: #CFCFCF;
                border-radius: 0px;
                border-top-right-radius: 12px;
                border-bottom-left-radius: 12px;
                border-bottom-right-radius: 12px;
              }
            QTabWidget::pane { background: transparent; }

            QPushButton#ActionButton {
                background-color: #EAEAEA;
                border: 1px solid #D6D6D6;
                border-radius: 10px;
                padding: 10px 0px;
                font-size: 14px;
                font-weight: bold;
                color: #333333;
            }

        """)

    def initUI(self):
        self.setWindowTitle('CalculatorPro')
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "..", "calculator.ico")))
        self.setMinimumWidth(885)
        self.setMinimumHeight(405)
        self.resize(885, 405)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        outer_layout = QVBoxLayout(central_widget)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)
        self.main_card = QFrame()
        self.main_card.setObjectName("MainCard")
        self.main_card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        card_layout = QVBoxLayout(self.main_card)
        card_layout.setSpacing(20)
        card_layout.setContentsMargins(30, 24, 30, 30)

        title = QLabel('课程目标达成情况评价及总结分析报告')
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 22px; font-weight: bold; margin-bottom: 10px;")
        card_layout.addWidget(title)
        top_btns_row1 = QHBoxLayout()
        top_btns_row2 = QHBoxLayout()

        self.course_open_btn = QPushButton('开课信息')
        self.course_basic_btn = QPushButton('课程基本信息')
        self.relation_btn = QPushButton('课程考核与课程目标对应关系')
        self.ratio_btn = QPushButton('成绩占比')
        self.grad_req_btn = QPushButton('课程目标与毕业要求的对应关系')
        self.settings_btn = QPushButton('设置')

        for btn in [self.course_open_btn, self.course_basic_btn, self.relation_btn, self.ratio_btn, self.grad_req_btn, self.settings_btn]:
            btn.setObjectName("TopButton")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)

        self.course_open_btn.clicked.connect(self.open_course_open_dialog)
        self.course_basic_btn.clicked.connect(self.open_course_basic_dialog)
        self.relation_btn.clicked.connect(self.open_relation_table)
        self.ratio_btn.clicked.connect(self.open_ratio_dialog)
        self.grad_req_btn.clicked.connect(self.open_grad_req_dialog)
        self.settings_btn.clicked.connect(self.open_settings_window)

        # Row 1: 开课信息 / 课程基本信息 / 课程考核与课程目标对应关系
        top_btns_row1.addWidget(self.course_open_btn)
        top_btns_row1.addWidget(self.course_basic_btn)
        top_btns_row1.addWidget(self.relation_btn)
        top_btns_row1.addStretch()

        # Row 2: 成绩占比 / 课程目标与毕业要求的对应关系 / 设置
        top_btns_row2.addWidget(self.ratio_btn)
        top_btns_row2.addWidget(self.grad_req_btn)
        top_btns_row2.addWidget(self.settings_btn)
        top_btns_row2.addStretch()

        card_layout.addLayout(top_btns_row1)
        card_layout.addLayout(top_btns_row2)

        self.tabs = QTabWidget()
        self.tab_forward = QWidget()
        self.tab_reverse = QWidget()
        self.tab_forward.setObjectName("tab_forward")
        self.tab_reverse.setObjectName("tab_reverse")
        self.tabs.addTab(self.tab_forward, "正向模式")
        self.tabs.addTab(self.tab_reverse, "逆向模式")
        self.tabs.currentChanged.connect(self.on_tab_changed)
        self.tabs.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        card_layout.addWidget(self.tabs)
        self.control_bar = QFrame()
        self.control_bar.setObjectName("ControlBar")
        self.control_bar.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        control_layout = QHBoxLayout(self.control_bar)
        control_layout.setContentsMargins(10, 6, 10, 4)
        self.combo_spread = QComboBox()
        self.combo_spread.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        self.combo_spread.setMinimumContentsLength(0)
        self.combo_spread.addItems(["大跨度（14-23分）", "中跨度（7-13分）", "小跨度（2-6分）"] )
        self.combo_dist = QComboBox()
        self.combo_dist.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        self.combo_dist.setMinimumContentsLength(0)
        self.combo_dist.addItems(["标准正态", "高分倾向", "低分倾向", "两极分化", "档位打分", "完全随机"] )
        self.combo_noise = QComboBox()
        self.combo_noise.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        self.combo_noise.setMinimumContentsLength(0)
        self.combo_noise.addItems(["无", "详情..."])
        self.combo_noise.currentIndexChanged.connect(self.on_noise_changed)
        self.combo_style = QComboBox()
        self.combo_style.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        self.combo_style.setMinimumContentsLength(0)
        self.combo_style.addItems(["专业", "口语", "简洁", "详细", "幽默"] )
        self.word_count_input = QLineEdit()
        self.word_count_input.setValidator(QIntValidator(1, 9999))
        self.word_count_input.setText("200")
        self.word_count_input.setFixedWidth(70)
        control_layout.addWidget(QLabel("分数跨度:"))
        control_layout.addWidget(self.combo_spread)
        control_layout.addWidget(QLabel("分布模式:"))
        control_layout.addWidget(self.combo_dist)
        control_layout.addWidget(QLabel("噪声注入:"))
        control_layout.addWidget(self.combo_noise)
        control_layout.addWidget(QLabel("报告风格:"))
        control_layout.addWidget(self.combo_style)
        control_layout.addWidget(QLabel("字数:"))
        control_layout.addWidget(self.word_count_input)
        action_row = QHBoxLayout()
        action_row.setSpacing(16)
        action_row.setContentsMargins(16, 8, 16, 8)
        self.download_btn = QPushButton("模板下载")
        self.import_btn = QPushButton("导入文件")
        self.export_btn = QPushButton("导出结果")
        self.ai_report_btn = QPushButton("生成报告")
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.status_label = QLabel("")
        self.download_btn.clicked.connect(self.open_template_download)
        self.import_btn.clicked.connect(self.select_file)
        self.export_btn.clicked.connect(self.start_analysis)
        self.ai_report_btn.clicked.connect(self.start_generate_ai_report)

        for btn in [self.download_btn, self.import_btn, self.export_btn, self.ai_report_btn]:
            btn.setObjectName("ActionButton")
            btn.setMinimumHeight(44)
            btn.setMinimumWidth(180)
            btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            action_row.addWidget(btn)

        self.tab_panel = QFrame()
        self.tab_panel.setObjectName("TabPanel")
        self.tab_panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.tab_panel.setFixedHeight(180)
        self.tab_panel.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        tab_panel_layout = QVBoxLayout(self.tab_panel)
        tab_panel_layout.setContentsMargins(12, 8, 12, 20)
        tab_panel_layout.setSpacing(4)
        tab_panel_layout.addWidget(self.control_bar, alignment=Qt.AlignmentFlag.AlignHCenter)
        tab_panel_layout.addLayout(action_row)
        card_layout.addWidget(self.progress_bar)
        card_layout.addWidget(self.status_label)
        outer_layout.addWidget(self.main_card)
        self.fwd_layout = QVBoxLayout(self.tab_forward)
        self.fwd_layout.setContentsMargins(0, 0, 0, 0)
        self.fwd_layout.setSpacing(0)
        self.rev_layout = QVBoxLayout(self.tab_reverse)
        self.rev_layout.setContentsMargins(0, 0, 0, 0)
        self.rev_layout.setSpacing(0)
        self.fwd_layout.addWidget(self.tab_panel)
        self._sync_tabs_height()
        self.on_tab_changed(0)

    def _sync_tabs_height(self):
        """同步 Tab 容器的高度"""
        tabbar_h = self.tabs.tabBar().sizeHint().height()
        panel_h = self.tab_panel.height()
        self.tabs.setFixedHeight(tabbar_h + panel_h)

    def on_tab_changed(self, index):
        """切换正向/逆向模式时，动态调整面板位置及 UI 状态"""
        is_reverse = (index == 1)
        current_layout = self.rev_layout if is_reverse else self.fwd_layout
        if self.tab_panel.parent() is not (self.tab_reverse if is_reverse else self.tab_forward):
            self.tab_panel.setParent(None)
            current_layout.addWidget(self.tab_panel)
        
        # 只有逆向模式才启用分数分布调节
        self.combo_spread.setEnabled(is_reverse)
        self.combo_dist.setEnabled(is_reverse)
        self.combo_noise.setEnabled(is_reverse)

        opacity = "1.0" if is_reverse else "0.5"
        bg_color = "#FFFFFF" if is_reverse else "#E0E0E0"
        style = f"background-color: {bg_color}; opacity: {opacity};"
        self.combo_spread.setStyleSheet(style)
        self.combo_dist.setStyleSheet(style)
        self.combo_noise.setStyleSheet(style)

    def on_noise_changed(self, index):
        if self.combo_noise.currentText() == "详情...":
            subjects = self._get_relation_subjects() or ['平时作业', '期末考试']
            dialog = NoiseConfigDialog(self, available_subjects=subjects)
            if dialog.exec():
                self.noise_config = dialog.get_config()
                self.status_label.setText(
                    f"噪声配置已更新：覆盖率{int(self.noise_config['noise_ratio']*100)}%"
                )
            self.combo_noise.setCurrentIndex(0)

    def _get_relation_subjects(self):
        """从对应关系载荷中提取科目名称"""
        payload = self.relation_payload or {}
        subjects = []
        for link in payload.get("links", []):
            for method in link.get("methods", []):
                name = (method.get("name") or "").strip()
                if name:
                    subjects.append(name)
        seen = set()
        unique = []
        for name in subjects:
            if name not in seen:
                unique.append(name)
                seen.add(name)
        return unique

    def open_ratio_dialog(self):
        """打开成绩占比对话框"""
        dialog = RatioDialog(self, self.usual_ratio, self.midterm_ratio, self.final_ratio)
        if dialog.exec():
            u, m, f = dialog.get_values()
            try:
                total = float(u) + float(m) + float(f)
            except Exception:
                QMessageBox.warning(self, "提示", "请输入有效的数字")
                return
            if abs(total - 1.0) > 1e-6:
                QMessageBox.warning(self, "提示", "平时/期中/期末占比之和必须等于1")
                return
            self.usual_ratio = float(u)
            self.midterm_ratio = float(m)
            self.final_ratio = float(f)
            self.save_config()
            self.status_label.clear()

    def open_settings_window(self):
        """打开 AI 与课程详情设置窗口"""
        dialog = SettingsDialog(
            self,
            api_key=self.api_key,
            description=self.course_description,
            objective_requirements=self.objective_requirements,
            objectives_count=self.num_objectives,
            previous_achievement_file=self.previous_achievement_file,
        )
        if dialog.exec():
            self.api_key = dialog.api_key_value
            self.course_description = dialog.description_value
            self.objective_requirements = dialog.objective_requirements
            self.previous_achievement_file = dialog.previous_achievement_file
            self.save_config()

    def open_course_open_dialog(self):
        data = dict(self.course_open_info or {})
        if not data.get('course_name'):
            data['course_name'] = self._get_course_name()
        dialog = CourseOpenDialog(self, data)
        if dialog.exec():
            self.course_open_info = dialog.get_data()
            self.save_config()

    def open_course_basic_dialog(self):
        data = dict(self.course_basic_info or {})
        if not data.get('course_name'):
            data['course_name'] = self._get_course_name()
        dialog = CourseBasicDialog(self, data)
        if dialog.exec():
            self.course_basic_info = dialog.get_data()
            self.save_config()

    def open_grad_req_dialog(self):
        obj_count = self.num_objectives
        if not obj_count and self.relation_payload:
            obj_count = int(self.relation_payload.get('objectives_count', 0) or 0)
        if not obj_count:
            QMessageBox.warning(self, '??', '???????????????????????????')
            return
        dialog = GradRequirementDialog(self, obj_count, self.grad_req_map or [])
        if dialog.exec():
            self.grad_req_map = dialog.get_data()
            self.save_config()


    def open_relation_table(self):
        """打开课程目标对应关系编辑器"""    
        if self.usual_ratio in ("", None) or self.midterm_ratio in ("", None) or self.final_ratio in ("", None):
             QMessageBox.warning(self, '提示', '请先在“成绩占比”中设置比例')
             return

        default_objectives = self.num_objectives
        default_counts = None
        if self.relation_payload:
            default_objectives = self.relation_payload.get("objectives_count", default_objectives)
            links = self.relation_payload.get("links", [])
            if len(links) == 3:
                default_counts = [
                    len(links[0].get("methods", [])),
                    len(links[1].get("methods", [])),
                    len(links[2].get("methods", [])),
                ]
        setup = RelationTableSetupDialog(self, default_objectives=default_objectives, default_counts=default_counts)

        if setup.exec():
            values = setup.result_values
            if values:
                self.num_objectives = values[0]
                dialog = RelationTableEditorDialog(self, *values, existing_payload=self.relation_payload)
                if dialog.exec():
                     pass
                # 更新 Mock 输入
                self.weight_inputs = []
                avg_weight = 1.0 / self.num_objectives if self.num_objectives > 0 else 0
                for _ in range(self.num_objectives):
                    self.weight_inputs.append(MockInput(str(avg_weight)))

    def set_relation_payload(self, payload):
        self.relation_payload = payload
        self.save_config()

    def open_template_download(self):
        """下载模板并保存到 outputs 目录"""
        dialog = TemplateDownloadDialog(self)
        if not dialog.exec():
            return
        count = dialog.student_count
        base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        try:
            if self.tabs.currentIndex() == 0:
                if not self.relation_payload:
                    QMessageBox.warning(
                        self,
                        "提示",
                        "请先填写“课程考核与课程目标对应关系”。",
                    )
                    return
                outputs_dir = os.path.join(base_dir, "outputs")
                os.makedirs(outputs_dir, exist_ok=True)
                relation_json_path = os.path.join(outputs_dir, "relation_table.json")
                with open(relation_json_path, "w", encoding="utf-8") as f:
                    json.dump(self.relation_payload, f, ensure_ascii=False, indent=2)
                output_path = create_forward_template(base_dir, count, relation_json_path)
            else:
                output_path = create_reverse_template(base_dir, count)

            QMessageBox.information(
                self,
                "下载成功",
                f"模板已生成：{output_path}",
            )
        except Exception as exc:
            QMessageBox.critical(
                self,
                "错误",
                f"模板生成失败：{str(exc)}",
            )

    def select_file(self):
        """选择 Excel 成绩单"""
        file_name, _ = QFileDialog.getOpenFileName(self, "选择成绩单文件", "", "Excel Files (*.xlsx)")
        if file_name:
            self.input_file = file_name
            self.status_label.setText(f"已选择文件: {os.path.basename(file_name)}")

    def start_analysis(self):
        """开始处理成绩数据"""
        if not self.input_file:
            QMessageBox.warning(self, '错误', '请先选择成绩单文件')
            return
        
        # 映射 UI 选项到代码枚举
        mock_usual = MockInput(str(self.usual_ratio))
        mock_midterm = MockInput(str(self.midterm_ratio))
        mock_final = MockInput(str(self.final_ratio))
        mock_num_obj = MockInput(str(self.num_objectives))

        if not self.weight_inputs and self.num_objectives > 0:
             avg = 1.0 / self.num_objectives
             self.weight_inputs = [MockInput(str(avg)) for _ in range(self.num_objectives)]

        spread_mode_map = {'大跨度（14-23分）': 'large', '中跨度（7-13分）': 'medium', '小跨度（2-6分）': 'small'}
        dist_mode_map = {
            '标准正态': 'normal',
            '高分倾向': 'left_skewed',
            '低分倾向': 'right_skewed',
            '两极分化': 'bimodal',
            '档位打分': 'discrete',
            '完全随机': 'uniform',
        }
        s_mode = spread_mode_map.get(self.combo_spread.currentText(), 'medium')
        d_mode = dist_mode_map.get(self.combo_dist.currentText(), 'normal')

        try:
            self.processor = GradeProcessor(
                MockInput(self._get_course_name()),
                mock_num_obj,        
                self.weight_inputs,  
                mock_usual,
                mock_midterm,
                mock_final,
                self.status_label,
                self.input_file,
                course_description=self.course_description,
                objective_requirements=self.objective_requirements,
                relation_payload=self.relation_payload
            )

            if hasattr(self.processor, 'set_noise_config') and self.noise_config:
                self.processor.set_noise_config(self.noise_config)
            if self.tabs.currentIndex() == 0:
                overall = self.processor.process_forward_grades(
                    spread_mode=s_mode,
                    distribution=d_mode,
                )
            else:
                overall = self.processor.process_reverse_grades(
                    spread_mode=s_mode,
                    distribution=d_mode,
                )
            self.status_label.setText(f"处理完成，总达成度: {overall}")
            QMessageBox.information(self, "成功", "成绩处理已完成，请导出结果。")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"成绩处理失败: {str(e)}")

    def start_generate_ai_report(self):
        if not self.processor:
             QMessageBox.warning(self, "提示", "请先执行成绩处理")
             return
        report_style = self.combo_style.currentText()
        self.report_thread = GenerateReportThread(self.processor, self.num_objectives, self.current_achievement, report_style)
        self.report_thread.finished.connect(self.on_report_finished)
        self.report_thread.progress.connect(self.status_label.setText)
        self.report_thread.start()

    def on_report_finished(self):
        self.status_label.setText("AI报告生成完成")
        QMessageBox.information(self, "成功", "AI报告已生成。")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = GradeAnalysisApp()
    window.show()
    sys.exit(app.exec())
