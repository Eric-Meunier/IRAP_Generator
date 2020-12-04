import sys

from PyQt5.QtWidgets import QApplication

from src.report_generator import ReportGenerator

if __name__ == '__main__':
    app = QApplication(sys.argv)

    rg = ReportGenerator()
    rg.show()

    app.exec_()
