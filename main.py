#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT2Manual - PPT转PDF手册工具
主程序入口文件
"""

import sys
import os
from pathlib import Path
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QTranslator, QLocale
from PySide6.QtGui import QIcon

# 添加项目根目录到Python路径
PROJECT_ROOT = Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT))

from gui.main_window import MainWindow


def setup_application():
    """设置应用程序"""
    app = QApplication(sys.argv)
    app.setApplicationName("PPT2Manual")
    app.setApplicationVersion("0.0.1-alpha")
    app.setOrganizationName("pyinglie")

    # 设置应用程序图标
    icon_path = PROJECT_ROOT / "resources" / "icon.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    # 设置样式
    app.setStyle("Fusion")

    return app


def main():
    """主函数"""
    try:
        # 创建应用程序
        app = setup_application()

        # 创建主窗口
        window = MainWindow()
        window.show()

        # 运行应用程序
        sys.exit(app.exec())

    except Exception as e:
        print(f"应用程序启动失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()