#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
STFMail - smartMail 工资条邮件群发工具
主程序入口

作者: Claude Code
版本: 1.0.0
"""

import sys
import os

# 添加项目根目录到路径
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from gui.main_window import MainWindow


def main():
    """主程序入口"""
    app = MainWindow()
    app.mainloop()


if __name__ == '__main__':
    main()
