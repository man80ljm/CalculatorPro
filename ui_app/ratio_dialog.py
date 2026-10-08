import os
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QDoubleValidator, QIcon
from PyQt6.QtWidgets import QDialog, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton, QVBoxLayout


class RatioDialog(QDialog):
    """成绩占比设置弹窗"""

    def __init__(self, parent=None, usual="", midterm="", final=""):
        super().__init__(parent)
        self.setWindowTitle("成绩占比")
        self.setFixedWidth(420)
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
                min-width: 120px;
            }
            QPushButton:hover {
                background: #0056B3;
            }
            QPushButton:pressed {
                background: #004085;
            }
        """)
        self.usual_value = None
        self.midterm_value = None
        self.final_value = None
        self._build_ui(usual, midterm, final)

    def _build_ui(self, usual, midterm, final):
        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(32, 32, 32, 32)

        # 标题
        title = QLabel("设置成绩占比")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("""
            font-size: 18px;
            font-weight: bold;
            color: #2C3E50;
            padding-bottom: 8px;
        """)
        layout.addWidget(title)

        validator = QDoubleValidator(0.0, 1.0, 4, self)

        self.usual_input = self._row(layout, "平时考核：", validator, usual)
        self.midterm_input = self._row(layout, "期中考核：", validator, midterm)
        self.final_input = self._row(layout, "期末考核：", validator, final)

        # 提示文本
        hint = QLabel("提示：三项占比之和必须等于 1.0")
        hint.setStyleSheet("""
            font-size: 12px;
            color: #6C757D;
            padding: 8px;
            background: #F8F9FB;
            border-radius: 6px;
        """)
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(hint)

        btn = QPushButton("确认")
        btn.clicked.connect(self._on_confirm)
        btn.setFixedHeight(44)
        layout.addWidget(btn, alignment=Qt.AlignmentFlag.AlignCenter)
        self.setLayout(layout)

    def _row(self, layout, label_text, validator, value):
        row = QHBoxLayout()
        row.setSpacing(12)
        label = QLabel(label_text)
        label.setFixedWidth(100)
        input_box = QLineEdit()
        input_box.setValidator(validator)
        input_box.setFixedWidth(160)
        input_box.setPlaceholderText("0.0 - 1.0")
        if value != "":
            input_box.setText(str(value))
        row.addWidget(label)
        row.addStretch()
        row.addWidget(input_box)
        layout.addLayout(row)
        return input_box

    def _on_confirm(self):
        try:
            usual_text = self.usual_input.text().strip()
            midterm_text = self.midterm_input.text().strip()
            final_text = self.final_input.text().strip()

            if usual_text == "" or midterm_text == "" or final_text == "":
                raise ValueError("占比不能为空(至少填写0)")

            usual = float(usual_text)
            midterm = float(midterm_text)
            final = float(final_text)
        except ValueError as exc:
            QMessageBox.warning(self, "提示", str(exc))
            return

        if abs((usual + midterm + final) - 1.0) != 0:
            QMessageBox.warning(self, "提示", "占比总和必须为1")
            return

        self.usual_value = usual
        self.midterm_value = midterm
        self.final_value = final
        self.accept()