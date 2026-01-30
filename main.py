import sys
from PyQt6.QtWidgets import QApplication
from ui_app.main_window import GradeAnalysisApp


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GradeAnalysisApp()
    window.show()
    sys.exit(app.exec())