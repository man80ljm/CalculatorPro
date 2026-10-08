import sys

from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QProgressBar
from PyQt6.QtCore import Qt


class AppSplash(QWidget):
    def __init__(self):
        super().__init__(None, Qt.WindowType.SplashScreen | Qt.WindowType.FramelessWindowHint)
        self.setFixedSize(420, 220)
        self.setStyleSheet("background: #FFFFFF; border: 1px solid #D0D0D0;")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(12)

        title = QLabel("CalculatorPro")
        title.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        title.setStyleSheet("font-size: 20px; font-weight: 600; color: #202020;")

        self.status = QLabel("Starting...")
        self.status.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.status.setStyleSheet("font-size: 12px; color: #404040;")

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(10)
        self.progress.setStyleSheet(
            "QProgressBar {background: #F0F0F0; border: 1px solid #D0D0D0; border-radius: 5px;}"
            "QProgressBar::chunk {background: #4A90E2; border-radius: 5px;}"
        )

        layout.addWidget(title)
        layout.addStretch(1)
        layout.addWidget(self.status)
        layout.addWidget(self.progress)

    def update_status(self, text: str, value: int):
        self.status.setText(text)
        self.progress.setValue(value)
        QApplication.processEvents()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    splash = AppSplash()
    splash.show()
    splash.update_status("Starting...", 5)

    splash.update_status("Loading UI modules...", 30)
    from ui_app.main_window import GradeAnalysisApp

    splash.update_status("Initializing window...", 70)
    window = GradeAnalysisApp()
    window.show()
    splash.update_status("Finalizing...", 95)
    splash.close()
    sys.exit(app.exec())