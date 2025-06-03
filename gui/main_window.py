#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主窗口界面
"""

import os
import sys
from pathlib import Path
from typing import List
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QSplitter,
    QGroupBox, QMessageBox, QApplication, QMenuBar, QMenu, QStatusBar
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtGui import QAction, QIcon

from .components import (
    FileDropWidget, FileListWidget, PathInputWidget,
    ProgressWidget, ControlWidget
)
from core.ppt_converter import PPTConverterThread


class MainWindow(QMainWindow):
    """主窗口类"""

    def __init__(self):
        super().__init__()
        self.converter_thread = None
        self.setup_ui()
        self.setup_connections()
        self.setup_default_values()

    def setup_ui(self):
        """设置用户界面"""
        self.setWindowTitle("PPT2Manual - PPT转PDF手册工具 v0.0.1-alpha")
        self.setMinimumSize(960, 600)
        self.resize(960, 600)

        # 创建中央widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # 创建分割器
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)

        # 左侧面板
        left_panel = self.create_left_panel()
        splitter.addWidget(left_panel)

        # 右侧面板
        right_panel = self.create_right_panel()
        splitter.addWidget(right_panel)

        # 设置分割器比例
        splitter.setSizes([400, 600])

        # 控制区域
        control_area = self.create_control_area()
        main_layout.addWidget(control_area)

        # 创建菜单栏
        self.create_menu_bar()

        # 创建状态栏
        self.create_status_bar()

    def create_left_panel(self) -> QWidget:
        """创建左侧面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)

        # 输入区域
        input_group = QGroupBox("文件输入")
        input_layout = QVBoxLayout(input_group)

        # 文件拖拽区域
        self.drop_widget = FileDropWidget()
        input_layout.addWidget(self.drop_widget)

        # 路径输入
        self.input_path_widget = PathInputWidget("输入路径:", "选择包含PPT文件的文件夹")
        self.input_path_widget.set_folder_mode(True)
        input_layout.addWidget(self.input_path_widget)

        layout.addWidget(input_group)

        # 输出设置区域
        output_group = QGroupBox("输出设置")
        output_layout = QVBoxLayout(output_group)

        # 输出路径
        self.output_path_widget = PathInputWidget("输出路径:", "选择PDF文件保存位置")
        self.output_path_widget.set_folder_mode(False)
        output_layout.addWidget(self.output_path_widget)

        layout.addWidget(output_group)

        # 添加弹性空间
        layout.addStretch()

        return panel

    def create_right_panel(self) -> QWidget:
        """创建右侧面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)

        # 文件列表区域
        files_group = QGroupBox("待处理文件 (可拖拽排序)")
        files_layout = QVBoxLayout(files_group)

        # 文件列表
        self.file_list = FileListWidget()
        files_layout.addWidget(self.file_list)

        layout.addWidget(files_group)

        # 进度显示区域
        progress_group = QGroupBox("处理进度")
        progress_layout = QVBoxLayout(progress_group)

        self.progress_widget = ProgressWidget()
        progress_layout.addWidget(self.progress_widget)

        layout.addWidget(progress_group)

        return panel

    def create_control_area(self) -> QWidget:
        """创建控制区域"""
        self.control_widget = ControlWidget()
        return self.control_widget

    def create_menu_bar(self):
        """创建菜单栏"""
        menubar = self.menuBar()

        # 文件菜单
        file_menu = menubar.addMenu("文件(&F)")

        # 打开文件夹
        open_folder_action = QAction("打开文件夹(&O)", self)
        open_folder_action.setShortcut("Ctrl+O")
        open_folder_action.triggered.connect(self.open_folder)
        file_menu.addAction(open_folder_action)

        file_menu.addSeparator()

        # 退出
        exit_action = QAction("退出(&X)", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # 帮助菜单
        help_menu = menubar.addMenu("帮助(&H)")

        # 关于
        about_action = QAction("关于(&A)", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def create_status_bar(self):
        """创建状态栏"""
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")

    def setup_connections(self):
        """设置信号连接"""
        # 文件拖拽
        self.drop_widget.files_dropped.connect(self.handle_dropped_files)

        # 路径改变
        self.input_path_widget.path_changed.connect(self.handle_input_path_changed)

        # 控制按钮
        self.control_widget.start_conversion.connect(self.start_conversion)
        self.control_widget.cancel_conversion.connect(self.cancel_conversion)

    def setup_default_values(self):
        """设置默认值"""
        # 设置默认输出文件名
        desktop = Path.home() / "Desktop"
        default_output = desktop / "合并手册.pdf"
        self.output_path_widget.set_path(str(default_output))

    def handle_dropped_files(self, files: List[str]):
        """处理拖拽的文件"""
        for file_path in files:
            path = Path(file_path)
            if path.is_file() and path.suffix.lower() in ['.ppt', '.pptx']:
                # 单个PPT文件
                self.file_list.add_file(str(path))
            elif path.is_dir():
                # 文件夹，扫描PPT文件
                self.scan_folder_for_ppt(str(path))
                self.input_path_widget.set_path(str(path))

        self.update_ui_state()

    def handle_input_path_changed(self, path: str):
        """处理输入路径改变"""
        if os.path.isdir(path):
            self.file_list.clear_files()
            self.scan_folder_for_ppt(path)
            self.update_ui_state()

    def scan_folder_for_ppt(self, folder_path: str):
        """扫描文件夹中的PPT文件"""
        try:
            folder = Path(folder_path)
            ppt_files = []

            # 扫描PPT文件
            for ext in ['*.ppt', '*.pptx']:
                ppt_files.extend(folder.glob(ext))

            # 按文件名排序
            ppt_files.sort(key=lambda x: x.name.lower())

            # 添加到列表
            for ppt_file in ppt_files:
                self.file_list.add_file(str(ppt_file))

            if ppt_files:
                self.progress_widget.add_status(f"发现 {len(ppt_files)} 个PPT文件")
            else:
                self.progress_widget.add_status("未发现PPT文件")

        except Exception as e:
            self.progress_widget.add_status(f"扫描文件夹失败: {e}")

    def update_ui_state(self):
        """更新UI状态"""
        has_files = self.file_list.count() > 0
        has_output = bool(self.output_path_widget.get_path())

        self.control_widget.set_enabled(has_files and has_output)

        # 更新状态栏
        if has_files:
            self.status_bar.showMessage(f"已加载 {self.file_list.count()} 个PPT文件")
        else:
            self.status_bar.showMessage("就绪")

    def start_conversion(self):
        """开始转换"""
        # 验证输入
        file_paths = self.file_list.get_file_paths()
        output_path = self.output_path_widget.get_path()

        if not file_paths:
            QMessageBox.warning(self, "警告", "请先添加PPT文件")
            return

        if not output_path:
            QMessageBox.warning(self, "警告", "请设置输出路径")
            return

        # 确认覆盖
        if os.path.exists(output_path):
            reply = QMessageBox.question(
                self, "确认",
                f"文件 {output_path} 已存在，是否覆盖？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        # 开始转换
        self.control_widget.set_converting(True)
        self.progress_widget.show_progress(True)
        self.progress_widget.clear_status()
        self.progress_widget.add_status("开始转换...")

        # 创建转换线程
        self.converter_thread = PPTConverterThread(file_paths, output_path)
        self.converter_thread.progress_updated.connect(self.progress_widget.set_progress)
        self.converter_thread.status_updated.connect(self.progress_widget.add_status)
        self.converter_thread.conversion_finished.connect(self.conversion_finished)
        self.converter_thread.error_occurred.connect(self.conversion_error)

        self.converter_thread.start()

    def cancel_conversion(self):
        """取消转换"""
        if self.converter_thread and self.converter_thread.isRunning():
            self.converter_thread.cancel()
            self.progress_widget.add_status("正在取消...")

    def conversion_finished(self, output_path: str):
        """转换完成"""
        self.control_widget.set_converting(False)
        self.progress_widget.show_progress(False)
        self.progress_widget.add_status("转换完成！")
        self.status_bar.showMessage("转换完成")

        # 询问是否打开文件
        reply = QMessageBox.question(
            self, "转换完成",
            f"PDF文件已保存到：\n{output_path}\n\n是否立即打开？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )

        if reply == QMessageBox.Yes:
            os.startfile(output_path)

    def conversion_error(self, error_message: str):
        """转换错误"""
        self.control_widget.set_converting(False)
        self.progress_widget.show_progress(False)
        self.progress_widget.add_status(f"转换失败: {error_message}")
        self.status_bar.showMessage("转换失败")

        QMessageBox.critical(self, "错误", f"转换失败：\n{error_message}")

    def open_folder(self):
        """打开文件夹"""
        from PySide6.QtWidgets import QFileDialog
        folder = QFileDialog.getExistingDirectory(self, "选择包含PPT文件的文件夹")
        if folder:
            self.input_path_widget.set_path(folder)

    def show_about(self):
        """显示关于对话框"""
        QMessageBox.about(
            self, "关于 PPT2Manual",
            """
            <h3>PPT2Manual v0.0.1-alpha</h3>
            <p>PPT转PDF手册工具</p>
            <p>作者: pyinglie</p>
            <p>功能: 批量将PPT文件转换为合并的PDF手册</p>
            <br>
            <p>支持功能：</p>
            <ul>
            <li>批量PPT/PPTX转换</li>
            <li>PDF合并与页码</li>
            <li>自动生成书签</li>
            <li>目录页生成</li>
            <li>中文完全支持</li>
            </ul>
            """
        )

    def closeEvent(self, event):
        """关闭事件"""
        if self.converter_thread and self.converter_thread.isRunning():
            reply = QMessageBox.question(
                self, "确认退出",
                "转换正在进行中，确定要退出吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )

            if reply == QMessageBox.Yes:
                self.converter_thread.cancel()
                self.converter_thread.wait(3000)  # 等待3秒
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()