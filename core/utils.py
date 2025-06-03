#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工具函数模块
"""

import os
import logging
from pathlib import Path
from typing import List, Tuple
import platform


def setup_logging(log_level=logging.INFO):
    """设置日志"""
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),
        ]
    )


def ensure_directory(path: str) -> bool:
    """确保目录存在"""
    try:
        os.makedirs(path, exist_ok=True)
        return True
    except Exception as e:
        logging.error(f"创建目录失败 {path}: {e}")
        return False


def get_file_size_string(file_path: str) -> str:
    """获取文件大小字符串"""
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


def validate_ppt_file(file_path: str) -> bool:
    """验证PPT文件"""
    if not os.path.exists(file_path):
        return False

    file_ext = Path(file_path).suffix.lower()
    return file_ext in ['.ppt', '.pptx']


def get_available_filename(base_path: str) -> str:
    """获取可用的文件名（避免覆盖）"""
    if not os.path.exists(base_path):
        return base_path

    base = Path(base_path)
    name = base.stem
    ext = base.suffix
    parent = base.parent

    counter = 1
    while True:
        new_name = f"{name}_{counter}{ext}"
        new_path = parent / new_name
        if not new_path.exists():
            return str(new_path)
        counter += 1


def is_file_locked(file_path: str) -> bool:
    """检查文件是否被锁定"""
    try:
        with open(file_path, 'a'):
            pass
        return False
    except IOError:
        return True


def get_temp_directory() -> str:
    """获取临时目录"""
    import tempfile
    return tempfile.gettempdir()


def clean_filename(filename: str) -> str:
    """清理文件名，移除非法字符"""
    import re
    # 移除Windows文件名中的非法字符
    illegal_chars = r'[<>:"/\\|?*]'
    clean_name = re.sub(illegal_chars, '_', filename)
    return clean_name.strip()


def get_system_info() -> dict:
    """获取系统信息"""
    return {
        'platform': platform.platform(),
        'system': platform.system(),
        'release': platform.release(),
        'machine': platform.machine(),
        'processor': platform.processor(),
    }


def format_time_duration(seconds: float) -> str:
    """格式化时间持续时间"""
    if seconds < 60:
        return f"{seconds:.1f} 秒"
    elif seconds < 3600:
        minutes = seconds / 60
        return f"{minutes:.1f} 分钟"
    else:
        hours = seconds / 3600
        return f"{hours:.1f} 小时"


def get_unique_temp_filename(prefix: str = "temp", suffix: str = ".tmp") -> str:
    """获取唯一的临时文件名"""
    import tempfile
    import uuid

    temp_dir = get_temp_directory()
    unique_id = str(uuid.uuid4())[:8]
    filename = f"{prefix}_{unique_id}{suffix}"

    return os.path.join(temp_dir, filename)


class ProgressTracker:
    """进度跟踪器"""

    def __init__(self, total_steps: int):
        self.total_steps = total_steps
        self.current_step = 0
        self.step_weights = [1] * total_steps  # 默认每步权重相等

    def set_step_weights(self, weights: List[float]):
        """设置每步的权重"""
        if len(weights) == self.total_steps:
            self.step_weights = weights

    def get_progress_percentage(self) -> int:
        """获取进度百分比"""
        if self.total_steps == 0:
            return 100

        total_weight = sum(self.step_weights)
        completed_weight = sum(self.step_weights[:self.current_step])

        return int((completed_weight / total_weight) * 100)

    def next_step(self):
        """移动到下一步"""
        if self.current_step < self.total_steps:
            self.current_step += 1

    def reset(self):
        """重置进度"""
        self.current_step = 0