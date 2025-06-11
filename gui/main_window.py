#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主窗口界面 - 支持混合文件导入
"""

import os
import sys
import logging
from pathlib import Path
from typing import List, Dict, Any

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QListWidgetItem, QLabel,
    QFileDialog, QMessageBox, QProgressBar, QTextEdit,
    QSplitter, QGroupBox, QCheckBox, QComboBox
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont, QIcon, QDrag, QPixmap

from core.ppt_converter import MixedFileConverterThread
from core.pdf_merger import PDFMerger


class FileListWidget(QListWidget):
    """文件列表控件 - 支持混合文件类型和拖拽排序"""

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setDragDropMode(QListWidget.InternalMove)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            # 检查是否包含支持的文件类型
            urls = event.mimeData().urls()
            supported_files = []
            for url in urls:
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                    if self._is_supported_file(file_path):
                        supported_files.append(file_path)

            if supported_files:
                event.acceptProposedAction()
            else:
                event.ignore()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            files = []
            for url in urls:
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                    if self._is_supported_file(file_path):
                        files.append(file_path)

            if files:
                # 发送信号给主窗口
                if hasattr(self, 'file_dropped_callback'):
                    self.file_dropped_callback(files)
                event.acceptProposedAction()
            else:
                event.ignore()
        else:
            super().dropEvent(event)
            # 通知主窗口列表顺序已改变
            if hasattr(self, 'order_changed_callback'):
                self.order_changed_callback()

    def _is_supported_file(self, file_path: str) -> bool:
        """检查是否为支持的文件类型"""
        file_ext = Path(file_path).suffix.lower()
        return file_ext in ['.ppt', '.pptx', '.pdf']


class MainWindow(QMainWindow):
    """主窗口 - 支持混合文件导入"""

    def __init__(self):
        super().__init__()
        self.file_list = []  # 存储文件信息的列表
        self.converter_thread = None
        self.pdf_merger = PDFMerger()

        self.init_ui()
        self.setup_connections()

    def init_ui(self):
        """初始化界面"""
        self.setWindowTitle("PPT2Manual - PPT/PDF混合批量合并工具 v0.0.1-alpha")
        self.setMinimumSize(900, 700)

        # 创建中央控件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout(central_widget)

        # 创建分割器
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)

        # 左侧：文件管理区域
        left_widget = self._create_file_management_area()
        splitter.addWidget(left_widget)

        # 右侧：操作和日志区域
        right_widget = self._create_operation_area()
        splitter.addWidget(right_widget)

        # 设置分割器比例
        splitter.setSizes([450, 450])

        # 状态栏
        self.statusBar().showMessage("就绪 - 支持PPT/PPTX和PDF文件混合导入")

    def _create_file_management_area(self):
        """创建文件管理区域"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # 文件列表组
        file_group = QGroupBox("文件列表 (支持拖拽排序)")
        file_layout = QVBoxLayout(file_group)

        # 说明标签
        info_label = QLabel(
            "支持的文件类型：\n"
            "• PPT/PPTX文件 - 将转换为8页布局PDF\n"
            "• PDF文件 - 直接合并到最终文档\n"
            "• 可拖拽文件到此区域或点击按钮添加\n"
            "• 支持拖拽调整文件顺序"
        )
        info_label.setStyleSheet("color: #666; padding: 8px; background-color: #f5f5f5; border-radius: 4px;")
        file_layout.addWidget(info_label)

        # 文件列表
        self.file_list_widget = FileListWidget()
        self.file_list_widget.file_dropped_callback = self.add_files_from_drop
        self.file_list_widget.order_changed_callback = self.update_file_order_from_ui
        file_layout.addWidget(self.file_list_widget)

        # 文件操作按钮
        file_buttons_layout = QHBoxLayout()

        self.add_ppt_btn = QPushButton("添加PPT文件")
        self.add_ppt_btn.setToolTip("添加PPT/PPTX文件，将转换为8页布局")
        file_buttons_layout.addWidget(self.add_ppt_btn)

        self.add_pdf_btn = QPushButton("添加PDF文件")
        self.add_pdf_btn.setToolTip("添加已转换的PDF文件，直接合并")
        file_buttons_layout.addWidget(self.add_pdf_btn)

        self.add_mixed_btn = QPushButton("混合添加")
        self.add_mixed_btn.setToolTip("同时选择PPT和PDF文件")
        file_buttons_layout.addWidget(self.add_mixed_btn)

        file_layout.addLayout(file_buttons_layout)

        # 列表管理按钮
        list_buttons_layout = QHBoxLayout()

        self.remove_btn = QPushButton("移除选中")
        list_buttons_layout.addWidget(self.remove_btn)

        self.clear_btn = QPushButton("清空列表")
        list_buttons_layout.addWidget(self.clear_btn)

        self.move_up_btn = QPushButton("上移")
        self.move_down_btn = QPushButton("下移")
        list_buttons_layout.addWidget(self.move_up_btn)
        list_buttons_layout.addWidget(self.move_down_btn)

        file_layout.addLayout(list_buttons_layout)

        layout.addWidget(file_group)

        # 输出设置组
        output_group = QGroupBox("输出设置")
        output_layout = QVBoxLayout(output_group)

        self.output_path_label = QLabel("输出路径：未选择")
        self.output_path_label.setWordWrap(True)
        output_layout.addWidget(self.output_path_label)

        self.select_output_btn = QPushButton("选择输出路径")
        output_layout.addWidget(self.select_output_btn)

        # 文件统计信息
        self.file_stats_label = QLabel("文件统计：无文件")
        self.file_stats_label.setStyleSheet("color: #666; padding: 4px;")
        output_layout.addWidget(self.file_stats_label)

        layout.addWidget(output_group)

        return widget

    def _create_operation_area(self):
        """创建操作区域"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # 转换控制组
        control_group = QGroupBox("转换控制")
        control_layout = QVBoxLayout(control_group)

        # 转换按钮
        self.convert_btn = QPushButton("开始转换")
        self.convert_btn.setMinimumHeight(45)
        self.convert_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        control_layout.addWidget(self.convert_btn)

        # 取消按钮
        self.cancel_btn = QPushButton("取消转换")
        self.cancel_btn.setEnabled(False)
        control_layout.addWidget(self.cancel_btn)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        control_layout.addWidget(self.progress_bar)

        # 状态标签
        self.status_label = QLabel("准备就绪")
        self.status_label.setStyleSheet("font-weight: bold; color: #333;")
        control_layout.addWidget(self.status_label)

        layout.addWidget(control_group)

        # 预览组
        preview_group = QGroupBox("文件预览")
        preview_layout = QVBoxLayout(preview_group)

        self.preview_text = QTextEdit()
        self.preview_text.setMaximumHeight(150)
        self.preview_text.setReadOnly(True)
        self.preview_text.setFont(QFont("Consolas", 9))
        self.preview_text.setPlainText("将显示合并后的文档结构...")
        preview_layout.addWidget(self.preview_text)

        layout.addWidget(preview_group)

        # 日志组
        log_group = QGroupBox("转换日志")
        log_layout = QVBoxLayout(log_group)

        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(200)
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        log_layout.addWidget(self.log_text)

        # 日志控制按钮
        log_buttons_layout = QHBoxLayout()

        self.clear_log_btn = QPushButton("清空日志")
        log_buttons_layout.addWidget(self.clear_log_btn)
        log_buttons_layout.addStretch()

        log_layout.addLayout(log_buttons_layout)

        layout.addWidget(log_group)

        return widget

    def setup_connections(self):
        """设置信号连接"""
        self.add_ppt_btn.clicked.connect(self.add_ppt_files)
        self.add_pdf_btn.clicked.connect(self.add_pdf_files)
        self.add_mixed_btn.clicked.connect(self.add_mixed_files)
        self.remove_btn.clicked.connect(self.remove_selected_files)
        self.clear_btn.clicked.connect(self.clear_file_list)
        self.move_up_btn.clicked.connect(self.move_file_up)
        self.move_down_btn.clicked.connect(self.move_file_down)
        self.select_output_btn.clicked.connect(self.select_output_path)
        self.convert_btn.clicked.connect(self.start_conversion)
        self.cancel_btn.clicked.connect(self.cancel_conversion)
        self.clear_log_btn.clicked.connect(self.clear_log)

        # 文件列表选择变化
        self.file_list_widget.itemSelectionChanged.connect(self.update_button_states)

    def add_ppt_files(self):
        """添加PPT文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择PPT文件",
            "",
            "PowerPoint文件 (*.ppt *.pptx);;所有文件 (*.*)"
        )

        if files:
            self.add_files_to_list(files, 'ppt')

    def add_pdf_files(self):
        """添加PDF文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择PDF文件",
            "",
            "PDF文件 (*.pdf);;所有文件 (*.*)"
        )

        if files:
            self.add_files_to_list(files, 'pdf')

    def add_mixed_files(self):
        """混合添加文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择PPT或PDF文件",
            "",
            "支持的文件 (*.ppt *.pptx *.pdf);;PowerPoint文件 (*.ppt *.pptx);;PDF文件 (*.pdf);;所有文件 (*.*)"
        )

        if files:
            ppt_files = []
            pdf_files = []

            for file_path in files:
                ext = Path(file_path).suffix.lower()
                if ext in ['.ppt', '.pptx']:
                    ppt_files.append(file_path)
                elif ext == '.pdf':
                    pdf_files.append(file_path)

            if ppt_files:
                self.add_files_to_list(ppt_files, 'ppt')
            if pdf_files:
                self.add_files_to_list(pdf_files, 'pdf')

    def add_files_from_drop(self, files: List[str]):
        """从拖拽添加文件"""
        ppt_files = []
        pdf_files = []

        for file_path in files:
            ext = Path(file_path).suffix.lower()
            if ext in ['.ppt', '.pptx']:
                ppt_files.append(file_path)
            elif ext == '.pdf':
                pdf_files.append(file_path)

        if ppt_files:
            self.add_files_to_list(ppt_files, 'ppt')
        if pdf_files:
            self.add_files_to_list(pdf_files, 'pdf')

    def add_files_to_list(self, files: List[str], file_type: str):
        """添加文件到列表"""
        added_count = 0

        for file_path in files:
            # 检查是否已存在
            if any(item['file'] == file_path for item in self.file_list):
                continue

            file_name = Path(file_path).name

            if file_type == 'pdf':
                # 获取PDF信息
                pdf_info = self.pdf_merger.get_pdf_info(file_path)
                display_name = f"📄 [PDF] {file_name} ({pdf_info['page_count']} 页)"

                file_info = {
                    'file': file_path,
                    'name': file_name,
                    'title': pdf_info['title'] if pdf_info['title'] else Path(file_path).stem,
                    'type': 'pdf',
                    'page_count': pdf_info['page_count'],
                    'order': len(self.file_list)
                }
            else:
                display_name = f"📊 [PPT] {file_name}"

                file_info = {
                    'file': file_path,
                    'name': file_name,
                    'title': Path(file_path).stem,
                    'type': 'ppt',
                    'order': len(self.file_list)
                }

            self.file_list.append(file_info)

            # 添加到界面列表
            item = QListWidgetItem(display_name)
            item.setToolTip(f"文件路径: {file_path}\n类型: {file_type.upper()}")
            self.file_list_widget.addItem(item)

            added_count += 1

        if added_count > 0:
            self.update_button_states()
            self.update_file_statistics()
            self.update_preview()
            self.log_message(f"已添加 {added_count} 个{file_type.upper()}文件")

    def update_file_order_from_ui(self):
        """从界面更新文件顺序"""
        # 根据界面列表的顺序重新排序文件列表
        new_file_list = []
        for i in range(self.file_list_widget.count()):
            item = self.file_list_widget.item(i)
            # 根据显示文本找到对应的文件信息
            for file_info in self.file_list:
                expected_name = f"📄 [PDF] {file_info['name']}" if file_info[
                                                                      'type'] == 'pdf' else f"📊 [PPT] {file_info['name']}"
                if file_info['type'] == 'pdf':
                    expected_name += f" ({file_info['page_count']} 页)"

                if item.text() == expected_name:
                    file_info['order'] = i
                    new_file_list.append(file_info)
                    break

        self.file_list = new_file_list
        self.update_preview()

    def remove_selected_files(self):
        """移除选中的文件"""
        current_row = self.file_list_widget.currentRow()
        if current_row >= 0:
            self.file_list_widget.takeItem(current_row)
            self.file_list.pop(current_row)

            # 重新排序
            for i, file_info in enumerate(self.file_list):
                file_info['order'] = i

            self.update_button_states()
            self.update_file_statistics()
            self.update_preview()
            self.log_message("已移除选中的文件")

    def clear_file_list(self):
        """清空文件列表"""
        if self.file_list:
            reply = QMessageBox.question(
                self,
                "确认清空",
                "确定要清空所有文件吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )

            if reply == QMessageBox.Yes:
                self.file_list.clear()
                self.file_list_widget.clear()
                self.update_button_states()
                self.update_file_statistics()
                self.update_preview()
                self.log_message("已清空文件列表")

    def move_file_up(self):
        """上移文件"""
        current_row = self.file_list_widget.currentRow()
        if current_row > 0:
            # 交换列表中的位置
            self.file_list[current_row], self.file_list[current_row - 1] = \
                self.file_list[current_row - 1], self.file_list[current_row]

            # 更新order
            self.file_list[current_row]['order'] = current_row
            self.file_list[current_row - 1]['order'] = current_row - 1

            # 更新界面
            self._refresh_file_list_display()
            self.file_list_widget.setCurrentRow(current_row - 1)
            self.update_preview()

    def move_file_down(self):
        """下移文件"""
        current_row = self.file_list_widget.currentRow()
        if current_row >= 0 and current_row < len(self.file_list) - 1:
            # 交换列表中的位置
            self.file_list[current_row], self.file_list[current_row + 1] = \
                self.file_list[current_row + 1], self.file_list[current_row]

            # 更新order
            self.file_list[current_row]['order'] = current_row
            self.file_list[current_row + 1]['order'] = current_row + 1

            # 更新界面
            self._refresh_file_list_display()
            self.file_list_widget.setCurrentRow(current_row + 1)
            self.update_preview()

    def _refresh_file_list_display(self):
        """刷新文件列表显示"""
        self.file_list_widget.clear()

        for file_info in self.file_list:
            if file_info['type'] == 'pdf':
                display_name = f"📄 [PDF] {file_info['name']} ({file_info['page_count']} 页)"
            else:
                display_name = f"📊 [PPT] {file_info['name']}"

            item = QListWidgetItem(display_name)
            item.setToolTip(f"文件路径: {file_info['file']}\n类型: {file_info['type'].upper()}")
            self.file_list_widget.addItem(item)

    def update_file_statistics(self):
        """更新文件统计信息"""
        if not self.file_list:
            self.file_stats_label.setText("文件统计：无文件")
            return

        ppt_count = sum(1 for item in self.file_list if item['type'] == 'ppt')
        pdf_count = sum(1 for item in self.file_list if item['type'] == 'pdf')

        total_pdf_pages = sum(item.get('page_count', 0) for item in self.file_list if item['type'] == 'pdf')

        stats_text = f"文件统计：{len(self.file_list)} 个文件 (PPT: {ppt_count}, PDF: {pdf_count}"
        if total_pdf_pages > 0:
            stats_text += f", PDF原始页数: {total_pdf_pages}"
        stats_text += ") - 全部转换为8页布局"

        self.file_stats_label.setText(stats_text)

    def update_preview(self):
        """更新文件预览"""
        if not self.file_list:
            self.preview_text.setPlainText("将显示合并后的文档结构...")
            return

        preview_text = "合并后的文档结构预览（统一8页布局）：\n\n"
        preview_text += "目录页\n"
        preview_text += "=" * 40 + "\n"

        for i, file_info in enumerate(self.file_list):
            file_type_str = "PDF-8页布局" if file_info['type'] == 'pdf' else "PPT-8页布局"
            preview_text += f"{i + 1}. [{file_type_str}] {file_info['title']}\n"

        preview_text += "\n注：所有文件都将转换为每页显示8张的统一布局格式"

        self.preview_text.setPlainText(preview_text)

    def select_output_path(self):
        """选择输出路径"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "选择输出PDF文件",
            "混合合并手册.pdf",
            "PDF文件 (*.pdf);;所有文件 (*.*)"
        )

        if file_path:
            self.output_path = file_path
            self.output_path_label.setText(f"输出路径：{file_path}")
            self.update_button_states()

    def update_button_states(self):
        """更新按钮状态"""
        has_files = len(self.file_list) > 0
        has_output = hasattr(self, 'output_path')
        has_selection = self.file_list_widget.currentRow() >= 0
        is_converting = self.converter_thread is not None and self.converter_thread.isRunning()

        # 文件操作按钮
        self.remove_btn.setEnabled(has_selection and not is_converting)
        self.clear_btn.setEnabled(has_files and not is_converting)
        self.move_up_btn.setEnabled(has_selection and self.file_list_widget.currentRow() > 0 and not is_converting)
        self.move_down_btn.setEnabled(
            has_selection and self.file_list_widget.currentRow() < len(self.file_list) - 1 and not is_converting)

        # 转换按钮
        self.convert_btn.setEnabled(has_files and has_output and not is_converting)
        self.cancel_btn.setEnabled(is_converting)

        # 添加文件按钮
        self.add_ppt_btn.setEnabled(not is_converting)
        self.add_pdf_btn.setEnabled(not is_converting)
        self.add_mixed_btn.setEnabled(not is_converting)
        self.select_output_btn.setEnabled(not is_converting)

    def start_conversion(self):
        """开始转换"""
        if not self.file_list:
            QMessageBox.warning(self, "警告", "请先添加要转换的文件！")
            return

        if not hasattr(self, 'output_path'):
            QMessageBox.warning(self, "警告", "请先选择输出路径！")
            return

        # 使用混合文件转换线程
        self.converter_thread = MixedFileConverterThread(self.file_list, self.output_path)
        self.converter_thread.progress_updated.connect(self.update_progress)
        self.converter_thread.status_updated.connect(self.update_status)
        self.converter_thread.conversion_finished.connect(self.conversion_finished)
        self.converter_thread.error_occurred.connect(self.conversion_error)

        self.converter_thread.start()

        self.progress_bar.setVisible(True)
        self.update_button_states()
        self.log_message("开始混合文件转换...")

    def cancel_conversion(self):
        """取消转换"""
        if self.converter_thread and self.converter_thread.isRunning():
            self.converter_thread.cancel()
            self.log_message("正在取消转换...")

    def update_progress(self, current, total):
        """更新进度"""
        self.progress_bar.setValue(int(current * 100 / total))

    def update_status(self, message):
        """更新状态"""
        self.status_label.setText(message)
        self.log_message(message)

    def conversion_finished(self, output_file):
        """转换完成"""
        self.progress_bar.setVisible(False)
        self.status_label.setText("转换完成")
        self.update_button_states()

        # 询问是否打开文件
        reply = QMessageBox.question(
            self,
            "转换完成",
            f"混合文件转换完成！\n输出文件：{output_file}\n\n是否现在打开文件？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )

        if reply == QMessageBox.Yes:
            os.startfile(output_file)

        self.log_message(f"转换完成：{output_file}")

    def conversion_error(self, error_message):
        """转换错误"""
        self.progress_bar.setVisible(False)
        self.status_label.setText("转换失败")
        self.update_button_states()

        QMessageBox.critical(self, "转换失败", f"转换过程中发生错误：\n{error_message}")
        self.log_message(f"错误：{error_message}")

    def log_message(self, message):
        """添加日志消息"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")

    def clear_log(self):
        """清空日志"""
        self.log_text.clear()