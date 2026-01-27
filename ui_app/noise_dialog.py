import os
from typing import List

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QRadioButton,
    QSlider,
    QVBoxLayout,
)


class NoiseConfigDialog(QDialog):
    """噪声注入配置弹窗"""

    def __init__(self, available_subjects: List[str], parent=None):
        super().__init__(parent)
        self.available_subjects = available_subjects
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "..", "calculator.ico")))
        self._build_ui()

    def _build_ui(self):
        self.setWindowTitle("噪声注入详情")
        self.setFixedWidth(480)
        self.setStyleSheet("""
            QDialog { 
                background: #FFFFFF; 
                border: 3px solid #007BFF; 
                border-radius: 16px; 
            }
            QGroupBox { 
                border: 2px solid #E8ECF1;
                border-radius: 10px;
                padding: 16px;
                margin-top: 12px;
                font-size: 14px;
                font-weight: 600;
                color: #2C3E50;
                background: #F8F9FB;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 8px;
                color: #007BFF;
            }
            QCheckBox { 
                padding: 6px;
                font-size: 13px;
                color: #2C3E50;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 4px;
                border: 2px solid #CED4DA;
            }
            QCheckBox::indicator:checked {
                background: #007BFF;
                border: 2px solid #007BFF;
            }
            QRadioButton { 
                padding: 6px;
                font-size: 13px;
                color: #2C3E50;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border-radius: 9px;
                border: 2px solid #CED4DA;
            }
            QRadioButton::indicator:checked {
                background: #007BFF;
                border: 2px solid #007BFF;
            }
            QSlider::groove:horizontal {
                height: 6px;
                background: #E8ECF1;
                border-radius: 3px;
            }
            QSlider::handle:horizontal {
                background: #007BFF;
                width: 18px;
                height: 18px;
                margin: -6px 0;
                border-radius: 9px;
            }
            QSlider::handle:horizontal:hover {
                background: #0056B3;
            }
            QDialogButtonBox QPushButton { 
                background: #007BFF; 
                border: none; 
                border-radius: 8px; 
                padding: 10px 24px;
                color: white;
                font-weight: 500;
                min-width: 80px;
            }
            QDialogButtonBox QPushButton:hover {
                background: #0056B3;
            }
        """)
        layout = QVBoxLayout()
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        # 区域1：异常覆盖率
        group_rate = QGroupBox("异常覆盖率 (Trigger Rate)")
        rate_layout = QVBoxLayout()
        rate_layout.setSpacing(12)
        
        slider_container = QHBoxLayout()
        self.rate_label = QLabel("0.10")
        self.rate_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.rate_label.setStyleSheet("""
            font-size: 24px;
            font-weight: bold;
            color: #007BFF;
            padding: 8px;
            background: white;
            border-radius: 8px;
        """)
        
        self.rate_slider = QSlider(Qt.Orientation.Horizontal)
        self.rate_slider.setRange(0, 100)
        self.rate_slider.setValue(10)
        self.rate_slider.valueChanged.connect(self._on_rate_change)
        
        rate_layout.addWidget(self.rate_label)
        rate_layout.addWidget(self.rate_slider)
        group_rate.setLayout(rate_layout)

        # 区域2：挂科严重程度
        group_severity = QGroupBox("挂科严重程度 (Severity Mode)")
        severity_layout = QVBoxLayout()
        severity_layout.setSpacing(8)
        self.severity_group = QButtonGroup(self)
        self.severity_random = QRadioButton("随机分布 (40-59分) - [默认]")
        self.severity_near = QRadioButton("边缘挂科 (55-59分) - [模拟惜败]")
        self.severity_cat = QRadioButton("严重缺失 (0-40分) - [模拟缺考]")
        self.severity_random.setChecked(True)
        self.severity_group.addButton(self.severity_random)
        self.severity_group.addButton(self.severity_near)
        self.severity_group.addButton(self.severity_cat)
        severity_layout.addWidget(self.severity_random)
        severity_layout.addWidget(self.severity_near)
        severity_layout.addWidget(self.severity_cat)
        group_severity.setLayout(severity_layout)

        # 区域3：允许挂科的科目
        group_targets = QGroupBox("允许挂科的科目 (Target Preference)")
        targets_layout = QVBoxLayout()
        targets_layout.setSpacing(8)
        self.target_checks = []
        for name in self.available_subjects:
            chk = QCheckBox(name)
            chk.setChecked(True)
            self.target_checks.append(chk)
            targets_layout.addWidget(chk)
        group_targets.setLayout(targets_layout)

        # 底部按钮
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout.addWidget(group_rate)
        layout.addWidget(group_severity)
        layout.addWidget(group_targets)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def _on_rate_change(self, value):
        self.rate_label.setText(f"{value / 100:.2f}")

    def get_config(self):
        if self.severity_near.isChecked():
            severity = "near_miss"
        elif self.severity_cat.isChecked():
            severity = "catastrophic"
        else:
            severity = "random"

        allowed_items = [chk.text() for chk in self.target_checks if chk.isChecked()]
        return {
            "noise_ratio": self.rate_slider.value() / 100.0,
            "severity_mode": severity,
            "allowed_items": allowed_items,
        }