#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自定义UI组件
"""

import os
from pathlib import Path
from typing import List, Optional
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QListWidget, QListWidgetItem, QProgressBar, QTextEdit, QFileDialog,
    QFrame, QGroupBox, QSplitter, QMessageBox
)
from PySide6.QtCore import Qt, Signal, QMimeData, QUrl
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QFont


class FileDropWidget(QWidget):
    """支持拖拽的文件区域组件"""

    files_dropped = Signal(list)  # 文件拖拽信号

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setup_ui()

    def setup_ui(self):
        """设置UI"""
        layout = QVBoxLayout(self)

        # 拖拽提示标签
        self.drop_label = QLabel("将PPT文件或文件夹拖拽到此处")
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 10px;
                padding: 50px;
                font-size: 16px;
                color: #666;
                background-color: #f9f9f9;
            }
            QLabel:hover {
                border-color: #0078d4;
                color: #0078d4;
                background-color: #f0f8ff;
            }
        """)

        layout.addWidget(self.drop_label)

    def dragEnterEvent(self, event: QDragEnterEvent):
        """拖拽进入事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drop_label.setStyleSheet("""
                QLabel {
                    border: 2px dashed #0078d4;
                    border-radius: 10px;
                    padding: 50px;
                    font-size: 16px;
                    color: #0078d4;
                    background-color: #e6f3ff;
                }
            """)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        """拖拽离开事件"""
        self.drop_label.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 10px;
                padding: 50px;
                font-size: 16px;
                color: #666;
                background-color: #f9f9f9;
            }
        """)

    def dropEvent(self, event: QDropEvent):
        """拖拽放下事件"""
        files = []
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if os.path.exists(file_path):
                files.append(file_path)

        if files:
            self.files_dropped.emit(files)

        # 恢复样式
        self.dragLeaveEvent(event)
        event.acceptProposedAction()


class FileListWidget(QListWidget):
    """可拖拽排序的文件列表组件"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragDropMode(QListWidget.InternalMove)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSelectionMode(QListWidget.ExtendedSelection)

    def add_file(self, file_path: str):
        """添加文件到列表"""
        file_info = Path(file_path)
        if file_info.suffix.lower() in ['.ppt', '.pptx']:
            item = QListWidgetItem()
            item.setText(f"{file_info.name}")
            item.setData(Qt.UserRole, str(file_path))
            item.setToolTip(f"路径: {file_path}\n大小: {self.get_file_size(file_path)}")
            self.addItem(item)

    def get_file_size(self, file_path: str) -> str:
        """获取文件大小"""
        try:
            size = os.path.getsize(file_path)
            if size < 1024:
                return f"{size} B"
            elif size < 1024 * 1024:
                return f"{size / 1024:.1f} KB"
            else:
                return f"{size / (1024 * 1024):.1f} MB"
        except:
            return "未知大小"

    def get_file_paths(self) -> List[str]:
        """获取所有文件路径（按当前顺序）"""
        paths = []
        for i in range(self.count()):
            item = self.item(i)
            paths.append(item.data(Qt.UserRole))
        return paths

    def clear_files(self):
        """清空文件列表"""
        self.clear()


class PathInputWidget(QWidget):
    """路径输入组件"""

    path_changed = Signal(str)

    def __init__(self, label_text: str = "路径:", placeholder: str = "", parent=None):
        super().__init__(parent)
        self.setup_ui(label_text, placeholder)

    def setup_ui(self, label_text: str, placeholder: str):
        """设置UI"""
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 标签
        self.label = QLabel(label_text)
        self.label.setMinimumWidth(80)
        layout.addWidget(self.label)

        # 路径输入框
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText(placeholder)
        self.path_edit.textChanged.connect(self.path_changed.emit)
        layout.addWidget(self.path_edit)

        # 浏览按钮
        self.browse_btn = QPushButton("浏览...")
        self.browse_btn.setMaximumWidth(80)
        self.browse_btn.clicked.connect(self.browse_path)
        layout.addWidget(self.browse_btn)

    def browse_path(self):
        """浏览路径"""
        if hasattr(self, '_is_folder') and self._is_folder:
            path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        else:
            path, _ = QFileDialog.getSaveFileName(
                self, "保存文件", "", "PDF文件 (*.pdf)"
            )

        if path:
            self.path_edit.setText(path)

    def set_folder_mode(self, is_folder: bool = True):
        """设置为文件夹模式"""
        self._is_folder = is_folder

    def get_path(self) -> str:
        """获取路径"""
        return self.path_edit.text().strip()

    def set_path(self, path: str):
        """设置路径"""
        self.path_edit.setText(path)


class ProgressWidget(QWidget):
    """进度显示组件"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        """设置UI"""
        layout = QVBoxLayout(self)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # 状态文本
        self.status_text = QTextEdit()
        self.status_text.setMaximumHeight(100)
        self.status_text.setReadOnly(True)
        font = QFont("Consolas", 9)
        self.status_text.setFont(font)
        layout.addWidget(self.status_text)

    def show_progress(self, show: bool = True):
        """显示/隐藏进度条"""
        self.progress_bar.setVisible(show)

    def set_progress(self, value: int, maximum: int = 100):
        """设置进度"""
        self.progress_bar.setMaximum(maximum)
        self.progress_bar.setValue(value)

    def add_status(self, message: str):
        """添加状态信息"""
        self.status_text.append(message)
        self.status_text.ensureCursorVisible()

    def clear_status(self):
        """清空状态信息"""
        self.status_text.clear()


class ControlWidget(QWidget):
    """控制按钮组件"""

    start_conversion = Signal()
    cancel_conversion = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.is_converting = False

    def setup_ui(self):
        """设置UI"""
        layout = QHBoxLayout(self)

        # 添加弹性空间
        layout.addStretch()

        # 开始转换按钮
        self.start_btn = QPushButton("开始转换")
        self.start_btn.setMinimumHeight(35)
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 5px;
                font-size: 14px;
                font-weight: bold;
                padding: 8px 20px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.start_btn.clicked.connect(self._on_start_clicked)
        layout.addWidget(self.start_btn)

        # 取消按钮
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setMinimumHeight(35)
        self.cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #d73527;
                color: white;
                border: none;
                border-radius: 5px;
                font-size: 14px;
                padding: 8px 20px;
            }
            QPushButton:hover {
                background-color: #c42e21;
            }
            QPushButton:pressed {
                background-color: #a52714;
            }
        """)
        self.cancel_btn.clicked.connect(self._on_cancel_clicked)
        self.cancel_btn.setVisible(False)
        layout.addWidget(self.cancel_btn)

    def _on_start_clicked(self):
        """开始按钮点击"""
        self.start_conversion.emit()

    def _on_cancel_clicked(self):
        """取消按钮点击"""
        self.cancel_conversion.emit()

    def set_converting(self, converting: bool):
        """设置转换状态"""
        self.is_converting = converting
        self.start_btn.setVisible(not converting)
        self.cancel_btn.setVisible(converting)

    def set_enabled(self, enabled: bool):
        """设置按钮可用状态"""
        self.start_btn.setEnabled(enabled)