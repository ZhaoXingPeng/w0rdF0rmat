# -*- coding: utf-8 -*-
import sys
from PyQt6.QtWidgets import QApplication
from main_window import MainWindow

def run():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    import os
    sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
    run() 