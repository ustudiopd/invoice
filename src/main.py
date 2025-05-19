import sys
from PyQt5.QtWidgets import QApplication
from ui.main_window import ExcelGPTViewer

def main():
    app = QApplication(sys.argv)
    viewer = ExcelGPTViewer()
    viewer.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main() 