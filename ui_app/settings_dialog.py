import os
import requests
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QFrame,
    QScrollArea,
    QWidget,
    QVBoxLayout,
)


class TestApiThread(QThread):
    result = pyqtSignal(str)

    def __init__(self, api_key: str):
        super().__init__()
        self.api_key = api_key

    def run(self):
        self.result.emit(test_deepseek_api(self.api_key))


def test_deepseek_api(api_key: str) -> str:
    url = "https://api.deepseek.com/v1/chat/completions"
    api_key = api_key.strip().strip("<").strip(">")
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": "deepseek-chat",
        "messages": [
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": "测试连接"},
        ],
        "temperature": 0.7,
        "top_p": 1,
        "max_tokens": 10,
        "stream": False,
    }
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=10)
        response.raise_for_status()
        return "连接成功"
    except requests.RequestException as exc:
        return f"连接失败: {exc}"


class SettingsDialog(QDialog):
    """设置弹窗：API Key、课程简介、目标要求、上一学年达成度"""

    def __init__(
        self,
        parent=None,
        api_key="",
        description="",
        objective_requirements=None,
        objectives_count=0,
        previous_achievement_file="",
    ):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.resize(650, 700)
        self.setMinimumSize(650, 500)
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "..", "calculator.ico")))

        self.api_key_value = api_key
        self.description_value = description
        self.objective_requirements = objective_requirements or []
        self.objectives_count = objectives_count
        self.previous_achievement_file = previous_achievement_file
        self.objective_inputs = []
        self._build_ui()
        # allow vertical resize

    def _build_ui(self):
        outer_layout = QVBoxLayout()
        outer_layout.setSpacing(0)
        outer_layout.setContentsMargins(0, 0, 0, 0)

        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)
        self.setStyleSheet("""
            QDialog { 
                background: #FFFFFF;
            }
            QLabel {
                font-size: 14px;
                color: #333333;
                font-weight: 600;
            }
            QLineEdit {
                background: #FFFFFF;
                border: 2px solid #C0C0C0;
                border-radius: 8px;
                padding: 8px 10px;
                font-size: 13px;
                color: #333333;
            }
            QLineEdit:focus {
                border: 2px solid #8A8A8A;
            }
            QTextEdit {
                background: #FFFFFF;
                border: 2px solid #C0C0C0;
                border-radius: 8px;
                padding: 8px 10px;
                font-size: 13px;
                color: #333333;
            }
            QTextEdit::viewport {
                border: 2px solid #C0C0C0;
                border-radius: 8px;
                background: #FFFFFF;
            }
            QTextEdit:focus {
                border: 2px solid #8A8A8A;
            }
            QPushButton {
                background: #8A8A8A;
                color: #FFFFFF;
                border: none;
                border-radius: 10px;
                padding: 10px 18px;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #7C7C7C;
            }
            QPushButton:pressed {
                background: #6E6E6E;
            }
        """)

        title = QLabel("设置")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        desc_label = QLabel("课程简介:")
        self.desc_input = QTextEdit()
        self.desc_input.setFixedHeight(140)
        self.desc_input.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.desc_input.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Plain)
        self.desc_input.setLineWidth(2)
        self.desc_input.setText(self.description_value)
        self.desc_input.setPlaceholderText("请输入课程简介...")
        self.desc_input.setStyleSheet("border: 2px solid #C0C0C0; border-radius: 8px;")
        layout.addWidget(desc_label)
        layout.addWidget(self.desc_input)

        self._build_objective_inputs(layout)

        self.import_prev_btn = QPushButton("导入上一学年达成度表")
        self.import_prev_btn.clicked.connect(self._import_previous_file)
        layout.addWidget(self.import_prev_btn)

        self.file_path_label = QLabel(self.previous_achievement_file or "未选择文件")
        layout.addWidget(self.file_path_label)

        api_layout = QHBoxLayout()
        api_label = QLabel("API KEY:")
        self.api_input = QLineEdit()
        self.api_input.setText(self.api_key_value)
        self.api_input.setPlaceholderText("请输入 DeepSeek API Key")
        self.test_btn = QPushButton("检测")
        self.test_btn.clicked.connect(self._test_api)
        api_layout.addWidget(api_label)
        api_layout.addWidget(self.api_input)
        api_layout.addWidget(self.test_btn)
        layout.addLayout(api_layout)

        btn_layout = QHBoxLayout()
        save_btn = QPushButton("保存")
        clear_btn = QPushButton("清空")
        save_btn.setFixedWidth(120)
        clear_btn.setFixedWidth(120)
        save_btn.clicked.connect(self._on_save)
        clear_btn.clicked.connect(self._on_clear)
        btn_layout.addStretch()
        btn_layout.addWidget(save_btn)
        btn_layout.addSpacing(214)
        btn_layout.addWidget(clear_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        content = QWidget()
        content.setLayout(layout)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setWidget(content)
        outer_layout.addWidget(scroll)
        self.setLayout(outer_layout)

    def _build_objective_inputs(self, layout):
        if self.objectives_count <= 0:
            return
        section = QLabel("课程目标要求:")
        section.setStyleSheet("font-weight: bold;")
        layout.addWidget(section)
        for idx in range(self.objectives_count):
            label = QLabel(f"课程目标{idx + 1}要求")
            input_box = QLineEdit()
            input_box.setPlaceholderText(f"请输入目标{idx + 1}要求")
            if idx < len(self.objective_requirements):
                input_box.setText(self.objective_requirements[idx])
            layout.addWidget(label)
            layout.addWidget(input_box)
            self.objective_inputs.append(input_box)

    def _import_previous_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "选择上一学年达成度表",
            "",
            "Excel Files (*.xlsx)",
        )
        if file_name:
            self.previous_achievement_file = file_name
            self.file_path_label.setText(file_name)

    def _test_api(self):
        api_key = self.api_input.text().strip()
        if not api_key:
            QMessageBox.warning(self, "提示", "请先输入API Key")
            return
        self.test_dialog = QMessageBox(self)
        self.test_dialog.setWindowTitle("测试连接")
        self.test_dialog.setText("连接中...")
        self.test_dialog.setStandardButtons(QMessageBox.StandardButton.NoButton)
        self.test_dialog.show()
        self.test_thread = TestApiThread(api_key)
        self.test_thread.result.connect(self._on_test_result)
        self.test_thread.start()

    def _on_test_result(self, result: str):
        self.test_dialog.setText(result)
        self.test_dialog.setStandardButtons(QMessageBox.StandardButton.Ok)
        self.test_dialog.exec()

    def _on_save(self):
        try:
            new_api = self.api_input.text().strip()
            if new_api:
                self.api_key_value = new_api
            self.description_value = self.desc_input.toPlainText().strip()
            self.objective_requirements = [item.text().strip() for item in self.objective_inputs]

            parent = self.parent()
            if parent:
                parent.api_key = self.api_key_value
                parent.course_description = self.description_value
                parent.objective_requirements = self.objective_requirements
                parent.previous_achievement_file = self.previous_achievement_file
                if hasattr(parent, "save_config"):
                    parent.save_config()

            QMessageBox.information(self, "提示", "保存成功")
            self.accept()
        except Exception as exc:
            QMessageBox.warning(self, "提示", f"保存失败：{exc}")

    def _on_clear(self):
        try:
            self.desc_input.clear()
            for item in self.objective_inputs:
                item.clear()
            self.file_path_label.setText("未选择文件")
            self.previous_achievement_file = ""

            parent = self.parent()
            if parent:
                if hasattr(parent, "usual_ratio"):
                    parent.usual_ratio = 0.0
                    parent.midterm_ratio = 0.0
                    parent.final_ratio = 0.0
                if hasattr(parent, "relation_payload"):
                    parent.relation_payload = None
                if hasattr(parent, "num_objectives"):
                    parent.num_objectives = 0
                if hasattr(parent, "course_open_info"):
                    parent.course_open_info = {}
                if hasattr(parent, "course_basic_info"):
                    parent.course_basic_info = {}
                if hasattr(parent, "grad_req_map"):
                    parent.grad_req_map = []
                parent.course_description = ""
                parent.objective_requirements = []
                parent.previous_achievement_file = ""
                if hasattr(parent, "save_config"):
                    parent.save_config()

            QMessageBox.information(self, "提示", "已清空")
        except Exception as exc:
            QMessageBox.warning(self, "提示", f"清空失败：{exc}")

