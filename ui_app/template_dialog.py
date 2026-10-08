import os
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIntValidator, QIcon
from PyQt6.QtWidgets import QDialog, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton, QVBoxLayout


class TemplateDownloadDialog(QDialog):
    """\u6a21\u677f\u4e0b\u8f7d\u5f39\u7a97"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("\u6a21\u677f\u4e0b\u8f7d")
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

        title = QLabel("\u4e0b\u8f7d\u6210\u7ee9\u6a21\u677f")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("""
            font-size: 18px;
            font-weight: bold;
            color: #2C3E50;
            padding-bottom: 8px;
        """)
        layout.addWidget(title)

        row = QHBoxLayout()
        row.setSpacing(12)
        label = QLabel("\u5b66\u751f\u4eba\u6570\uff1a")
        label.setFixedWidth(100)
        self.input_box = QLineEdit()
        self.input_box.setValidator(QIntValidator(1, 9999, self))
        self.input_box.setFixedWidth(140)
        self.input_box.setPlaceholderText("1-9999")
        row.addWidget(label)
        row.addStretch()
        row.addWidget(self.input_box)
        layout.addLayout(row)

        hint = QLabel("\u8bf7\u586b\u5199\u73ed\u7ea7\u5b66\u751f\u603b\u4eba\u6570")
        hint.setStyleSheet("""
            font-size: 12px;
            color: #6C757D;
            padding: 8px;
            background: #F8F9FB;
            border-radius: 6px;
        """)
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(hint)

        btn = QPushButton("\u70b9\u51fb\u4e0b\u8f7d")
        btn.setFixedHeight(44)
        btn.clicked.connect(self._on_confirm)
        layout.addWidget(btn, alignment=Qt.AlignmentFlag.AlignCenter)

        self.setLayout(layout)

    def _on_confirm(self):
        text = self.input_box.text().strip()
        if not text:
            QMessageBox.warning(self, "\u63d0\u793a", "\u8bf7\u8f93\u5165\u5b66\u751f\u4eba\u6570")
            return
        self.student_count = int(text)
        self.accept()
