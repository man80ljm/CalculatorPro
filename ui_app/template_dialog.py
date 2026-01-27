import os
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIntValidator, QIcon
from PyQt6.QtWidgets import QDialog, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton, QVBoxLayout


class TemplateDownloadDialog(QDialog):
    """模板下载弹窗"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("模板下载")
        self.setFixedWidth(380)
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "..", "calculator.ico")))
        self.setStyleSheet("""
            QDialog { 
                background: #FFFFFF; 
                border-radius: 16px;
                border: 2px solid #E8ECF1;
            }
            QLabel {
                font-size: 14px;
                color: #2C3E50;
                font-weight: 500;
            }
            QLineEdit { 
                background: #F8F9FB; 
                border: 2px solid #E8ECF1; 
                border-radius: 8px; 
                padding: 10px 14px;
                font-size: 14px;
                color: #2C3E50;
            }
            QLineEdit:focus {
                border: 2px solid #007BFF;
                background: #FFFFFF;
            }
            QPushButton { 
                background: #007BFF; 
                border: none; 
                border-radius: 10px; 
                padding: 12px 32px;
                color: white;
                font-size: 14px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #0056B3;
            }
            QPushButton:pressed {
                background: #004085;
            }
        """)
        self.student_count = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(32, 32, 32, 32)

        # 标题
        title = QLabel("下载成绩模板")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("""
            font-size: 18px;
            font-weight: bold;
            color: #2C3E50;
            padding-bottom: 8px;
        """)
        layout.addWidget(title)

        # 学生人数输入
        row = QHBoxLayout()
        row.setSpacing(12)
        label = QLabel("学生人数：")
        label.setFixedWidth(100)
        self.input_box = QLineEdit()
        self.input_box.setValidator(QIntValidator(1, 9999, self))
        self.input_box.setFixedWidth(140)
        self.input_box.setPlaceholderText("1-9999")
        row.addWidget(label)
        row.addStretch()
        row.addWidget(self.input_box)
        layout.addLayout(row)

        # 提示信息
        hint = QLabel("💡 请输入班级学生总人数")
        hint.setStyleSheet("""
            font-size: 12px;
            color: #6C757D;
            padding: 8px;
            background: #F8F9FB;
            border-radius: 6px;
        """)
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(hint)

        # 下载按钮
        btn = QPushButton("📥 点击下载")
        btn.setFixedHeight(44)
        btn.clicked.connect(self._on_confirm)
        layout.addWidget(btn, alignment=Qt.AlignmentFlag.AlignCenter)

        self.setLayout(layout)

    def _on_confirm(self):
        text = self.input_box.text().strip()
        if not text:
            QMessageBox.warning(self, "提示", "请输入学生人数")
            return
        self.student_count = int(text)
        self.accept()