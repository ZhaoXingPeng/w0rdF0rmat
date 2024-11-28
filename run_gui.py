# -*- coding: utf-8 -*-
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.gui.app import run

if __name__ == '__main__':
    run() 