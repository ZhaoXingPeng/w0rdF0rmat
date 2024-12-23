# -*- coding: utf-8 -*-
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from PyQt6.QtWidgets import QApplication
from src.gui.main_window import MainWindow  # 使用绝对导入

def run():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    run() 