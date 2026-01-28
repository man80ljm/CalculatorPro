import os
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon
from relation_table import PasteTableWidget
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QTableWidgetItem, QPushButton, QHBoxLayout, QLabel, QHeaderView, QSizePolicy

class GradRequirementDialog(QDialog):
    """课程目标与毕业要求对应关系"""
    def __init__(self, parent=None, objectives=0, data=None):
        super().__init__(parent)
        self.setWindowTitle("课程目标与毕业要求的对应关系")
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "..", "calculator.ico")))
        self.objectives = max(0, int(objectives))
        self.data = data or []
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("课程目标与毕业要求的对应关系")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)

        self.table = PasteTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["课程目标", "支撑的毕业要求", "支撑的毕业要求指标点"])
        self.table.setRowCount(self.objectives if self.objectives else max(1, len(self.data)))
        self.table.horizontalHeader().setMinimumSectionSize(80)
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        for row in range(self.table.rowCount()):
            obj_name = f"课程目标{row + 1}"
            item = QTableWidgetItem(obj_name)
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 0, item)

            if row < len(self.data):
                self.table.setItem(row, 1, QTableWidgetItem(self.data[row].get("requirement", "")))
                self.table.setItem(row, 2, QTableWidgetItem(self.data[row].get("indicator", "")))

        layout.addWidget(self.table)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        save_btn = QPushButton("保存")
        save_btn.clicked.connect(self._on_save)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(save_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

        self._apply_table_sizing()

    def _apply_table_sizing(self):
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()
        self.table.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Preferred)
        self.adjustSize()
        self.setMinimumSize(self.size())
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def _on_save(self):
        data = []
        for row in range(self.table.rowCount()):
            requirement = self.table.item(row, 1)
            indicator = self.table.item(row, 2)
            data.append({
                "objective": f"课程目标{row + 1}",
                "requirement": requirement.text().strip() if requirement else "",
                "indicator": indicator.text().strip() if indicator else "",
            })
        self.data = data
        self.accept()

    def get_data(self):
        return self.data
