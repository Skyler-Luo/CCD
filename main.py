# -*- coding: utf-8 -*-
"""
CCD GUI 启动入口
"""
import os
import sys

# 将项目根目录添加到 sys.path
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from src.ui.ui_main import launch_gui

if __name__ == '__main__':
    launch_gui()
