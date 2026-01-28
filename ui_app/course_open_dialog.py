import os
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox

class CourseOpenDialog(QDialog):
    """开课信息"""
    def __init__(self, parent=None, data=None):
        super().__init__(parent)
        self.setWindowTitle("开课信息")
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "..", "calculator.ico")))
        self.data = data or {}
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("开课信息")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        # 学年与学期
        year_row = QHBoxLayout()
        year_row.addWidget(QLabel("学年起:"))
        self.year_start = QLineEdit()
        self.year_start.setPlaceholderText("如 2024")
        self.year_start.setText(self.data.get("year_start", ""))
        year_row.addWidget(self.year_start)
        year_row.addWidget(QLabel("学年止:"))
        self.year_end = QLineEdit()
        self.year_end.setPlaceholderText("如 2025")
        self.year_end.setText(self.data.get("year_end", ""))
        year_row.addWidget(self.year_end)
        year_row.addWidget(QLabel("学期:"))
        self.semester = QLineEdit()
        self.semester.setPlaceholderText("如 1")
        self.semester.setText(self.data.get("semester", ""))
        year_row.addWidget(self.semester)
        layout.addLayout(year_row)

        # 课程名称
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("课程名称:"))
        self.course_name = QLineEdit()
        self.course_name.setText(self.data.get("course_name", ""))
        row1.addWidget(self.course_name)
        layout.addLayout(row1)

        # 开课部门
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("开课部门:"))
        self.department = QLineEdit()
        self.department.setText(self.data.get("department", ""))
        row2.addWidget(self.department)
        layout.addLayout(row2)

        # 授课教师
        row3 = QHBoxLayout()
        row3.addWidget(QLabel("授课教师:"))
        self.teacher = QLineEdit()
        self.teacher.setText(self.data.get("teacher", ""))
        row3.addWidget(self.teacher)
        layout.addLayout(row3)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        save_btn = QPushButton("保存")
        save_btn.clicked.connect(self._on_save)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(save_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

    def _on_save(self):
        self.data = {
            "year_start": self.year_start.text().strip(),
            "year_end": self.year_end.text().strip(),
            "semester": self.semester.text().strip(),
            "course_name": self.course_name.text().strip(),
            "department": self.department.text().strip(),
            "teacher": self.teacher.text().strip(),
        }
        self.accept()

    def get_data(self):
        return self.data
