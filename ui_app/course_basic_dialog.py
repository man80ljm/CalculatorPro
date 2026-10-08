import os
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QGridLayout, QLabel, QLineEdit, QPushButton, QHBoxLayout

class CourseBasicDialog(QDialog):
    """课程基本信息"""
    def __init__(self, parent=None, data=None):
        super().__init__(parent)
        self.setWindowTitle("课程基本信息")
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "..", "calculator.ico")))
        self.data = data or {}
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("课程基本信息")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        grid = QGridLayout()
        grid.setHorizontalSpacing(16)
        grid.setVerticalSpacing(12)

        fields = [
            ("课程名称", "course_name"),
            ("学分", "credits"),
            ("学时", "hours"),
            ("课程性质", "course_type"),
            ("课程代码", "course_code"),
            ("学年学期", "school_year_term"),
            ("开课学院", "college"),
            ("任课教师", "teacher"),
            ("上课专业", "major"),
            ("上课班级", "class_name"),
            ("上课人数", "student_count"),
            ("考核人数", "exam_count"),
        ]

        self.inputs = {}
        for i, (label, key) in enumerate(fields):
            row = i // 2
            col = (i % 2) * 2
            grid.addWidget(QLabel(label + ":"), row, col)
            edit = QLineEdit()
            edit.setText(self.data.get(key, ""))
            grid.addWidget(edit, row, col + 1)
            self.inputs[key] = edit

        layout.addLayout(grid)

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
        self.data = {k: v.text().strip() for k, v in self.inputs.items()}
        self.accept()

    def get_data(self):
        return self.data
