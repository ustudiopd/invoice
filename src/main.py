import sys
from PyQt5.QtWidgets import QApplication
from src.ui.main_window import ExcelGPTViewer


if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = ExcelGPTViewer()
    viewer.show()
    sys.exit(app.exec_()) 