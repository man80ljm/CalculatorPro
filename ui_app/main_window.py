import sys

import os

import json

import pandas as pd

from PyQt6.QtGui import QIcon, QDoubleValidator, QIntValidator, QAction

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                            QFileDialog, QMessageBox, QGridLayout, QDialog, 
                            QTextEdit, QComboBox, QDialogButtonBox, QProgressBar,
                            QTabWidget, QFrame, QGroupBox, QSlider, QRadioButton, 
                            QButtonGroup, QCheckBox, QSizePolicy)
from PyQt6.QtCore import Qt, QThread, pyqtSignal



# ������ԭ�е��߼�ģ��

from core import GradeProcessor

from relation_table import RelationTableSetupDialog, RelationTableEditorDialog
from ui_app.settings_dialog import SettingsDialog


# ==========================================

# �����ࣺ���ڼ��ݾ��߼��� Mock ����

# ==========================================

class MockInput:

    """

    ��ƭ core.py �ĸ����ࡣ

    ��Ϊ�½����Ƴ�����ҳ���Ȩ������򣬵� core.py ��������Ҫ��ȡ widget.text()��

    ��������װ���ݣ��� core.py ��Ϊ�����ڲ��� UI �ؼ���

    """

    def __init__(self, text_value):

        self._text = str(text_value)

    def text(self):

        return self._text

    def setText(self, val):

        self._text = str(val)



# ==========================================

# �����ࣺ�ɼ�ռ������ (RatioDialog)

# ==========================================

class RatioDialog(QDialog):

    def __init__(self, parent=None, usual=0.2, midterm=0.3, final=0.5):

        super().__init__(parent)

        self.setWindowTitle("成绩占比设置")
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "..", "calculator.ico")))

        # 替换 setFixedSize
        self.resize(180, 200)
        self.setStyleSheet("""
            QLineEdit {
                background: #F2F2F2;
                border: 1px solid #D0D0D0;
                border-radius: 6px;
                padding: 6px 8px;
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

# �����ࣺ�������� (NoiseConfigDialog)

# ==========================================

class NoiseConfigDialog(QDialog):

    def __init__(self, parent=None, available_subjects=None):

        super().__init__(parent)

        self.setWindowTitle("噪声注入详情配置")
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "..", "calculator.ico")))

        self.setFixedWidth(450)

        # Ĭ�Ͻ��յĿ�Ŀ�����û�д����Ĭ��ֵ

        self.available_subjects = available_subjects if available_subjects else ['平时作业', '期末考试']

        self.init_ui()



    def init_ui(self):

        layout = QVBoxLayout()

        layout.setSpacing(15)



        # --- ���� 1: �쳣������ ---

        group_rate = QGroupBox("1. 异常覆盖率 (Trigger Rate)")

        rate_layout = QVBoxLayout()

        h_layout = QHBoxLayout()

        

        self.slider = QSlider(Qt.Orientation.Horizontal)

        self.slider.setRange(0, 100) # ʹ��0-100��Ӧ0.0-1.0

        self.slider.setValue(10)     # Ĭ�� 0.1

        

        self.label_rate = QLabel("0.10")

        self.label_rate.setStyleSheet("font-weight: bold; color: #333;")

        

        self.slider.valueChanged.connect(lambda v: self.label_rate.setText(f"{v/100:.2f}"))

        

        h_layout.addWidget(self.slider)

        h_layout.addWidget(self.label_rate)

        rate_layout.addLayout(h_layout)

        group_rate.setLayout(rate_layout)

        layout.addWidget(group_rate)



        # --- ���� 2: �ҿ����س̶� ---

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



        # --- ���� 3: ����ҿƵĿ�Ŀ ---

        group_target = QGroupBox("3. 允许挂科的科目 (Target Preference)")

        target_layout = QGridLayout()

        self.checks = {}

        

        for i, subject in enumerate(self.available_subjects):

            cb = QCheckBox(subject)

            cb.setChecked(True) # Ĭ��ȫѡ

            self.checks[subject] = cb

            target_layout.addWidget(cb, i // 2, i % 2) # �����Ų�

            

        group_target.setLayout(target_layout)

        layout.addWidget(group_target)



        # --- �ײ���ť ---

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.setStyleSheet("QPushButton{border-radius:10px; padding:6px 16px;}")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        

        self.setLayout(layout)



    def get_config(self):

        """���������ֵ�"""

        mode_map = {1: 'random', 2: 'near_miss', 3: 'catastrophic'}

        selected_mode = mode_map.get(self.btn_group.checkedId(), 'random')

        allowed = [subj for subj, cb in self.checks.items() if cb.isChecked()]

        

        return {

            "noise_ratio": self.slider.value() / 100.0,

            "severity_mode": selected_mode,

            "allowed_items": allowed

        }



# ==========================================

# �����ࣺģ������ (TemplateDownloadDialog)

# ==========================================

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

        # ����У�飬���� 1-500

        self.input_count.setValidator(QIntValidator(1, 500))

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

# �߳��ࣺ����ԭ�� (GenerateReportThread, TestApiThread)

# ==========================================

class GenerateReportThread(QThread):

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

            questions = ["�����һ��ȴ�������ĸĽ����"]

            for i in range(1, 6):

                questions.append(f"�γ�Ŀ��{i}����������")

                questions.append(f"�ÿγ�Ŀ��{i}��������������������Ľ���ʩ")

            total_questions = len(questions)

            self.progress.emit("��������AI��������...")

            self.progress_value.emit(0)



            context = f"�γ̼��: {self.processor.course_description}\n"

            for i, req in enumerate(self.processor.objective_requirements, 1):

                context += f"�γ�Ŀ��{i}Ҫ��: {req}\n"

            for i in range(1, 6):

                prev_score = self.processor.previous_achievement_data.get(f'�γ�Ŀ��{i}', 0)

                current_score = self.current_achievement.get(f'�γ�Ŀ��{i}', 0)

                context += f"�γ�Ŀ��{i}��һ��ȴ�ɶ�: {prev_score}\n"

                context += f"�γ�Ŀ��{i}����ȴ�ɶ�: {current_score}\n"

            

            # ��ȫ��ȡ�ܴ�ɶ�

            prev_total = self.processor.previous_achievement_data.get('�γ���Ŀ��', 0)

            current_total = self.current_achievement.get('�ܴ�ɶ�', 0)

            context += f"�γ���Ŀ����һ��ȴ�ɶ�: {prev_total}\n"

            context += f"�γ���Ŀ�걾��ȴ�ɶ�: {current_total}\n"



            answers = []

            # ע�⣺�˴���Ҫ���̴߳��� course_name�������� processor �л�ȡ

            # �����߳��в���ֱ�ӷ��� UI������ processor ���� course_name

            course_name = "�γ̱���" 

            

            for i, question in enumerate(questions):

                self.progress.emit(f"���ڴ���� {i+1}/{total_questions} ������...")

                self.progress_value.emit(i + 1)

                if "�γ�Ŀ��" in question and int(question.split('�γ�Ŀ��')[1][0]) > self.num_objectives:

                    answers.append("��")

                    continue

                prompt = f"{context}\n����: {question}\n����{self.report_style}�ķ��ش������������С�"

                answer = self.processor.call_deepseek_api(prompt)

                answers.append(answer)



            self.processor.generate_improvement_report(self.current_achievement, course_name, self.num_objectives, answers=answers)

            self.progress_value.emit(total_questions)

            self.finished.emit()

        except Exception as e:

            self.error.emit(f"����AI��������ʧ�ܣ�{str(e)}")



class TestApiThread(QThread):

    result = pyqtSignal(str)

    def __init__(self, processor, api_key):

        super().__init__()

        self.processor = processor

        self.api_key = api_key

    def run(self):

        result = self.processor.test_deepseek_api(self.api_key)

        self.result.emit(result)



# ==========================================

# �����ࣺ���ô��� (SettingsWindow - ����ԭ���߼�)

# ==========================================

class SettingsWindow(QDialog):

    def __init__(self, parent=None):

        super().__init__(parent)

        self.setWindowTitle('����')

        self.setFixedWidth(500)

        self.initUI()

    

    def initUI(self):

        layout = QVBoxLayout()

        layout.setSpacing(10)

        

        self.description_label = QLabel('�γ̼�飺')

        self.description_input = QTextEdit()

        self.description_input.setFixedHeight(80)

        if hasattr(self.parent(), 'course_description'):

            self.description_input.setText(self.parent().course_description)
        layout.addWidget(self.description_label)
        layout.addWidget(self.description_input)
        self.objectives_layout = QVBoxLayout()
        self.objective_inputs = []
        layout.addLayout(self.objectives_layout)      

        self.import_prev_btn = QPushButton('������һѧ���ɶȱ�')

        self.import_prev_btn.clicked.connect(self.import_previous_achievement)

        layout.addWidget(self.import_prev_btn)

        

        self.file_path_label = QLabel('')

        self.file_path_label.setStyleSheet('font-size: 12px; color: #666666;')

        layout.addWidget(self.file_path_label)

        if hasattr(self.parent(), 'previous_achievement_file') and self.parent().previous_achievement_file:

            self.file_path_label.setText(self.parent().previous_achievement_file)

        

        api_layout = QHBoxLayout()

        self.api_key_label = QLabel('API KEY:')

        self.api_key_input = QLineEdit()

        if hasattr(self.parent(), 'api_key'):

            self.api_key_input.setText(self.parent().api_key)

        self.test_api_btn = QPushButton('���')

        self.test_api_btn.clicked.connect(self.test_api_connection)

        api_layout.addWidget(self.api_key_label)

        api_layout.addWidget(self.api_key_input)

        api_layout.addWidget(self.test_api_btn)

        layout.addLayout(api_layout)

        

        button_layout = QHBoxLayout()

        self.save_btn = QPushButton('����')

        self.save_btn.clicked.connect(self.save_settings)

        self.clear_btn = QPushButton('���')

        self.clear_btn.clicked.connect(self.clear_settings)

        button_layout.addWidget(self.save_btn)

        button_layout.addWidget(self.clear_btn)

        layout.addLayout(button_layout)

        

        self.setLayout(layout)

    

    def update_objective_inputs(self, num_objectives):

        for input_field in self.objective_inputs:

            input_field.setParent(None)

        self.objective_inputs.clear()

        while self.objectives_layout.count():

            item = self.objectives_layout.takeAt(0)

            if item.widget():

                item.widget().deleteLater()

        

        parent_objective_requirements = getattr(self.parent(), 'objective_requirements', [])

        for i in range(num_objectives):

            label = QLabel(f'�γ�Ŀ��{i+1}Ҫ��')

            input_field = QLineEdit()

            input_field.setPlaceholderText(f'������Ŀ��{i+1}Ҫ��')

            if i < len(parent_objective_requirements):

                input_field.setText(parent_objective_requirements[i])

            self.objective_inputs.append(input_field)

            self.objectives_layout.addWidget(label)

            self.objectives_layout.addWidget(input_field)



    def import_previous_achievement(self):

        file_name, _ = QFileDialog.getOpenFileName(self, "ѡ����һѧ���ɶȱ�", "", "Excel Files (*.xlsx)")

        if file_name:

            self.parent().previous_achievement_file = file_name

            self.file_path_label.setText(file_name)



    def save_settings(self):

        self.parent().course_description = self.description_input.toPlainText()

        self.parent().objective_requirements = [input_field.text() for input_field in self.objective_inputs]

        self.parent().api_key = self.api_key_input.text()

        self.parent().save_config()

        QMessageBox.information(self, '�ɹ�', '�����ѱ���')



    def clear_settings(self):

        self.description_input.clear()

        for input_field in self.objective_inputs:

            input_field.clear()

        self.file_path_label.clear()

        self.parent().course_description = ""

        self.parent().objective_requirements = []

        self.parent().previous_achievement_file = ""

        self.parent().save_config()

        QMessageBox.information(self, '�ɹ�', '���������')



    def test_api_connection(self):

        # ��ģ������ processor �Ĳ���

        api_key = self.api_key_input.text().strip()

        if not api_key:

            QMessageBox.warning(self, '����', '�������� API Key')

            return

        QMessageBox.information(self, "��ʾ", "API ���Թ�����������������ʱ���� core ģ��")



# ==========================================

# �����ڣ�GradeAnalysisApp (UI �ع���)

# ==========================================

class GradeAnalysisApp(QMainWindow):

    def __init__(self):

        super().__init__()

        # ��ʼ�����ݱ���

        self.input_file = None

        self.previous_achievement_file = None

        self.course_description = ""

        self.objective_requirements = []

        self.api_key = ""

        self.num_objectives = 0  # ��Ҫ�����ٴ�UI��ȡ�����Ǵ洢�ڱ�����

        self.processor = None

        self.current_achievement = {}

        self.forward_input_file = None

        self.reverse_input_file = None

        

        # Ȩ������ (Mock)

        self.weight_data = [] # �洢Ȩ����ֵ

        self.weight_inputs = [] # �洢 MockInput ���󣬼��� core.py

        

        # ռ������

        self.usual_ratio = 0.2

        self.midterm_ratio = 0.3

        self.final_ratio = 0.5

        

        # ��������

        self.noise_config = None
        self.relation_payload = None



        self.load_config()

        self.initUI()

        self.apply_styles() # Ӧ����ʽ��



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
                    if self.relation_payload and not self.num_objectives:
                        self.num_objectives = self.relation_payload.get("objectives_count", 0)

                    # ���ر����ռ��

                    if 'ratios' in config:

                        self.usual_ratio = config['ratios'].get('usual', 0.2)

                        self.midterm_ratio = config['ratios'].get('midterm', 0.3)

                        self.final_ratio = config['ratios'].get('final', 0.5)

            except Exception as e:

                print(f"���������ļ�ʧ��: {str(e)}")



    def save_config(self):

        config_dir = os.path.join(os.getenv('APPDATA') or os.path.expanduser('~'), 'CalculatorApp')

        config_file = os.path.join(config_dir, 'config.json')

        config = {

            'api_key': self.api_key,

            'course_description': self.course_description,

            'objective_requirements': self.objective_requirements,

            'previous_achievement_file': self.previous_achievement_file,
            'relation_payload': self.relation_payload,

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

            print(f"���������ļ�ʧ��: {str(e)}")



    def apply_styles(self):

        """Ӧ���ִ����� QSS ��ʽ"""

        self.setStyleSheet("""

            /* ȫ�ֱ��� */

            QMainWindow { background-color: #F5F5F5; }

            QWidget { font-family: 'Microsoft YaHei', 'SimHei'; font-size: 14px; color: #333; }
            QPushButton { font-size: 14px; border-radius: 10px; }
            QDialogButtonBox QPushButton { border-radius: 10px; padding: 6px 16px; }

            

            /* ����Ƭ���� */

            #MainCard { background-color: #E6E6E6; border-radius: 0px; }
            QTabWidget { background: #E6E6E6; }
            QTabWidget::pane { border: none; margin: 0px; padding: 0px; background: #E6E6E6; }
            #tab_forward { background: #E6E6E6; }
            #tab_reverse { background: #E6E6E6; }



            /* ����� & ������ */

            QLineEdit, QComboBox {
                background-color: #FFFFFF;
                border: none;
                border-radius: 0px;
                padding: 6px 10px;
                min-height: 20px;
            }
            QComboBox::drop-down { border: none; }
            

            /* �������Ұ�ť */

            #TopButton {
                background-color: #C4C4C4; color: #333;
                border-radius: 10px; padding: 8px 20px;

                font-weight: bold; border: none;

            }

            #TopButton:hover { background-color: #B0B0B0; }



            /* ������ */

            #TabPanel { background-color: #D0D0D0; border-radius: 0px; }
            #ControlBar { background-color: transparent; border-radius: 0px; }
            

            /* �ײ�������ť */

            #ActionButton {
                background-color: #EAEAEA; border: 1px solid #D6D6D6;
                border-radius: 10px; padding: 10px 0;
                font-size: 14px; font-weight: bold; color: #333;
            }

            #ActionButton:hover { background-color: #FFFFFF; }

        """)



    def initUI(self):

        self.setWindowTitle('课程目标达成情况评价及总结分析报告')
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "..", "calculator.ico")))
        self.setMinimumWidth(850)

        self.setMinimumHeight(370)
        self.resize(850, 370)

        

        central_widget = QWidget()

        self.setCentralWidget(central_widget)

        outer_layout = QVBoxLayout(central_widget)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)


        # === 1. ����Ƭ���� ===

        self.main_card = QFrame()

        self.main_card.setObjectName("MainCard")

        self.main_card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        card_layout = QVBoxLayout(self.main_card)
        card_layout.setSpacing(20)

        card_layout.setContentsMargins(30, 24, 30, 30)


        # === 2. ���� ===

        title = QLabel('课程目标达成情况评价及总结分析报告')
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title.setStyleSheet("font-size: 22px; font-weight: bold; margin-bottom: 10px;")

        card_layout.addWidget(title)



        # === 3. �γ������� ===

        course_row = QHBoxLayout()

        course_label = QLabel('课程名称:')
        self.course_name_input = QLineEdit()

        self.course_name_input.setPlaceholderText('请输入课程名称...')
        course_row.addWidget(course_label)

        course_row.addWidget(self.course_name_input)

        card_layout.addLayout(course_row)



        # === 4. �������ܰ�ť ===

        top_btns_row = QHBoxLayout()

        self.ratio_btn = QPushButton('成绩占比')
        self.relation_btn = QPushButton('课程考核与课程目标对应关系')
        self.settings_btn = QPushButton('设置')
        

        for btn in [self.ratio_btn, self.relation_btn, self.settings_btn]:
            btn.setObjectName("TopButton")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            

        self.ratio_btn.clicked.connect(self.open_ratio_dialog)
        self.relation_btn.clicked.connect(self.open_relation_table)
        self.settings_btn.clicked.connect(self.open_settings_window)
        

        top_btns_row.addWidget(self.ratio_btn)
        top_btns_row.addWidget(self.relation_btn)
        top_btns_row.addWidget(self.settings_btn)
        top_btns_row.addStretch()

        card_layout.addLayout(top_btns_row)



        # === 5. ѡ� (����/����) ===

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



        # === 6. ������ (��ұ���) ===

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

        

        # === 7. �ײ�������ť ===

        action_row = QHBoxLayout()

        action_row.setSpacing(12)
        action_row.setContentsMargins(8, 0, 8, 0)

        

        self.download_btn = QPushButton("模板下载")

        self.import_btn = QPushButton("导入文件")

        self.export_btn = QPushButton("导出结果")

        self.ai_report_btn = QPushButton("生成报告")

        

        # ����������״̬��ǩ���ϵ��ײ�

        self.progress_bar = QProgressBar()

        self.progress_bar.setVisible(False)

        self.status_label = QLabel("")

        

        self.download_btn.clicked.connect(self.open_template_download)

        self.import_btn.clicked.connect(self.select_file)

        self.export_btn.clicked.connect(self.start_analysis)

        self.ai_report_btn.clicked.connect(self.start_generate_ai_report)



        for btn in [self.download_btn, self.import_btn, self.export_btn, self.ai_report_btn]:
            btn.setObjectName("ActionButton")
            btn.setMinimumHeight(40)
            btn.setFixedWidth(160)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            action_row.addWidget(btn)

            

        self.tab_panel = QFrame()
        self.tab_panel.setObjectName("TabPanel")
        self.tab_panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.tab_panel.setFixedHeight(140)
        tab_panel_layout = QVBoxLayout(self.tab_panel)
        tab_panel_layout.setContentsMargins(0, 0, 0, 0)

        tab_panel_layout.setSpacing(4)

        tab_panel_layout.addWidget(self.control_bar, alignment=Qt.AlignmentFlag.AlignHCenter)

        tab_panel_layout.addLayout(action_row)

        card_layout.addWidget(self.progress_bar)

        card_layout.addWidget(self.status_label)

        

        outer_layout.addWidget(self.main_card)

        

        # �������� Tab �ڲ��ӿղ��֣���ֹ���־���

        self.fwd_layout = QVBoxLayout(self.tab_forward)
        self.fwd_layout.setContentsMargins(0, 0, 0, 0)
        self.fwd_layout.setSpacing(0)

        self.rev_layout = QVBoxLayout(self.tab_reverse)
        self.rev_layout.setContentsMargins(0, 0, 0, 0)
        self.rev_layout.setSpacing(0)

        self.fwd_layout.addWidget(self.tab_panel)

        

        # ��ʼ�� Tab ״̬

        self._sync_tabs_height()
        self.on_tab_changed(0)

    def _sync_tabs_height(self):
        tabbar_h = self.tabs.tabBar().sizeHint().height()
        panel_h = self.tab_panel.height()
        self.tabs.setFixedHeight(tabbar_h + panel_h)



    # ================= �߼��������� =================



    def on_tab_changed(self, index):

        """�л�����/����ģʽʱ�����û�����ģ��ؼ�"""

        is_reverse = (index == 1)

        current_layout = self.rev_layout if is_reverse else self.fwd_layout

        if self.tab_panel.parent() is not (self.tab_reverse if is_reverse else self.tab_forward):

            self.tab_panel.setParent(None)

            current_layout.addWidget(self.tab_panel)

        # ����ģʽ�����ã�����ģʽ�½���

        self.combo_spread.setEnabled(is_reverse)

        self.combo_dist.setEnabled(is_reverse)

        self.combo_noise.setEnabled(is_reverse)

        

        # �Ӿ�����

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
        payload = self.relation_payload or {}
        subjects = []
        for link in payload.get("links", []):
            for method in link.get("methods", []):
                name = (method.get("name") or "").strip()
                if name:
                    subjects.append(name)
        # 去重保持顺序
        seen = set()
        unique = []
        for name in subjects:
            if name not in seen:
                unique.append(name)
                seen.add(name)
        return unique



    def open_ratio_dialog(self):

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



    def open_relation_table(self):

        # ������ԭ�е� RelationTable �߼�

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

                # values[0] ��Ŀ������

                self.num_objectives = values[0] 

                # ������Ա���Ȩ����Ϣ�� self.weight_data

                # ���� relation_table �Ƚϸ��ӣ����Ǽ������ᴦ��Ȩ�صı���

                # ����������Ҫ����������Ȩ�ز�����

                

                dialog = RelationTableEditorDialog(self, *values, existing_payload=self.relation_payload)

                if dialog.exec():

                     # ���� Dialog �������������õ�Ȩ�أ������� start_analysis ʱ��ȡ

                     pass

                

                # ��Ҫ������ MockInput �б���ƭ�� core.py

                # ���Ǽ���ÿ��Ŀ���Ӧһ��Ȩ������� (���������򻯣������û�ƽ������)

                # ʵ���� relation_table �Ѿ�������ϸȨ�ء�

                # Ϊ�˲��� core.py ������������ MockInput

                self.weight_inputs = []

                avg_weight = 1.0 / self.num_objectives if self.num_objectives > 0 else 0

                for _ in range(self.num_objectives):

                    self.weight_inputs.append(MockInput(str(avg_weight)))



    def set_relation_payload(self, payload):
        self.relation_payload = payload
        self.save_config()

    def open_template_download(self):

        dialog = TemplateDownloadDialog(self)

        if dialog.exec():

            count = dialog.student_count

            mode = "正向" if self.tabs.currentIndex() == 0 else "逆向"

            # TODO: ���� excel_handler ����ģ��

            QMessageBox.information(self, "下载成功", f"已生成{mode}模式模板，包含{count}名学生。")



    def select_file(self):

        file_name, _ = QFileDialog.getOpenFileName(self, "选择成绩单文件", "", "Excel Files (*.xlsx)")

        if file_name:

            self.input_file = file_name

            self.status_label.setText(f"已选择文件: {os.path.basename(file_name)}")



    def start_analysis(self):

        if not self.input_file:

            QMessageBox.warning(self, '错误', '请先选择成绩单文件')

            return

        

        # ���ݾ��߼������� MockInput ���󴫵ݸ� Processor

        # ����� Mock ��Ϊ�˷�ֹ core.py �����ʵ��������Ӧ�þ����� processor ��ȡ config

        mock_usual = MockInput(str(self.usual_ratio))

        mock_midterm = MockInput(str(self.midterm_ratio))

        mock_final = MockInput(str(self.final_ratio))

        mock_num_obj = MockInput(str(self.num_objectives))

        

        # ȷ��Ȩ���б��Ϊ��

        if not self.weight_inputs and self.num_objectives > 0:

             avg = 1.0 / self.num_objectives

             self.weight_inputs = [MockInput(str(avg)) for _ in range(self.num_objectives)]



        # ��ȡ�������

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

        

        # ʵ���� Processor (ע�����˳������� core.py һ��)

        # ���� core.py ǩ��: (course_name, num_obj, weights, usual_r, mid_r, final_r, status_label, input_file, ...)

        try:

            self.processor = GradeProcessor(

                self.course_name_input,

                mock_num_obj,        # ���� Mock ����

                self.weight_inputs,  # ���� Mock �����б�

                mock_usual,

                mock_midterm,

                mock_final,

                self.status_label,

                self.input_file,

                course_description=self.course_description,

                objective_requirements=self.objective_requirements

            )

            

            # ע���������� (��� processor ֧��)

            if hasattr(self.processor, 'set_noise_config') and self.noise_config:

                self.processor.set_noise_config(self.noise_config)

            

            # ��ʼ����

            # ע�⣺process_grades ������Ҫƥ�� core.py

            overall = self.processor.process_grades(

                self.num_objectives, 

                [float(w.text()) for w in self.weight_inputs], # ������ʵ float �б�

                self.usual_ratio,

                self.midterm_ratio,

                self.final_ratio,

                s_mode,

                d_mode,

                progress_callback=lambda idx: self.progress_bar.setValue(idx)

            )

            

            self.status_label.setText(f"处理完成，总达成度: {overall}")
            QMessageBox.information(self, "成功", "成绩处理已完成，请导出结果。")

            

        except Exception as e:
            QMessageBox.critical(self, "错误", f"成绩处理失败: {str(e)}")



    def start_generate_ai_report(self):

        # ����Ƿ��ѷ���

        if not self.processor:
             QMessageBox.warning(self, "提示", "请先执行成绩处理")
             return

             

        report_style = self.combo_style.currentText()

        # ����߳�

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



