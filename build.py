#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT2Manual 构建脚本 - 修正版
"""

import os
import sys
import shutil
import subprocess
import logging
from pathlib import Path

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)

# 项目信息
PROJECT_NAME = "PPT2Manual"
VERSION = "v0.0.1-alpha"
MAIN_SCRIPT = "main.py"
ICON_FILE = "resources/icon.ico"
DIST_DIR = "dist"
BUILD_DIR = "build"

# 必需的依赖包
REQUIRED_PACKAGES = [
    "PySide6",
    "python-pptx",
    "PyMuPDF",
    "reportlab",
    "pywin32",  # 修正为正确的包名
    "PyInstaller",
    "Pillow"
]


def check_dependencies():
    """检查依赖包是否已安装"""
    logger.info("检查依赖...")

    missing_packages = []

    for package in REQUIRED_PACKAGES:
        try:
            # 使用更可靠的方式检查包是否安装
            if package == "python-pptx":
                # python-pptx 实际导入名为 pptx
                import pptx
                logger.info(f"✓ {package}")
            elif package == "PyMuPDF":
                # PyMuPDF 实际导入名为 fitz
                import fitz
                logger.info(f"✓ {package}")
            elif package == "PySide6":
                import PySide6
                logger.info(f"✓ {package}")
            elif package == "reportlab":
                import reportlab
                logger.info(f"✓ {package}")
            elif package == "pywin32":
                try:
                    import win32com.client
                    logger.info(f"✓ {package}")
                except ImportError:
                    # 尝试替代方案
                    try:
                        import comtypes
                        logger.info(f"✓ {package} (使用 comtypes 替代)")
                    except ImportError:
                        missing_packages.append(package)
                        logger.info(f"✗ {package}")
            elif package == "PyInstaller":
                import PyInstaller
                logger.info(f"✓ {package}")
            elif package == "Pillow":
                from PIL import Image
                logger.info(f"✓ {package}")
            else:
                # 对于其他包，使用通用导入方式
                __import__(package.lower().replace("-", "_"))
                logger.info(f"✓ {package}")

        except ImportError:
            missing_packages.append(package)
            logger.info(f"✗ {package}")
        except Exception as e:
            logger.warning(f"? {package} (检查时出错: {e})")

    if missing_packages:
        logger.error(f"\n缺少依赖包: {', '.join(missing_packages)}")
        logger.error("请运行: pip install -r requirements.txt")
        return False

    logger.info("✓ 所有依赖已满足")
    return True


def clean_build_dirs():
    """清理构建目录"""
    logger.info("清理构建目录...")

    dirs_to_clean = [DIST_DIR, BUILD_DIR, "__pycache__"]

    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            try:
                shutil.rmtree(dir_name)
                logger.info(f"已清理: {dir_name}")
            except Exception as e:
                logger.warning(f"清理 {dir_name} 失败: {e}")


def check_main_script():
    """检查主脚本文件"""
    if not os.path.exists(MAIN_SCRIPT):
        logger.error(f"主脚本文件不存在: {MAIN_SCRIPT}")
        return False

    logger.info(f"✓ 主脚本文件: {MAIN_SCRIPT}")
    return True


def check_icon():
    """检查图标文件"""
    if os.path.exists(ICON_FILE):
        logger.info(f"✓ 图标文件: {ICON_FILE}")
        return ICON_FILE
    else:
        logger.warning(f"图标文件不存在: {ICON_FILE}")
        return None


def create_spec_file():
    """创建 PyInstaller spec 文件"""
    logger.info("创建 PyInstaller 配置文件...")

    icon_option = f"icon='{ICON_FILE}'" if check_icon() else "icon=None"

    spec_content = f'''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['{MAIN_SCRIPT}'],
    pathex=[],
    binaries=[],
    datas=[
        ('resources', 'resources'),
        ('core', 'core'),
        ('gui', 'gui'),
    ],
    hiddenimports=[
        'PySide6.QtCore',
        'PySide6.QtWidgets', 
        'PySide6.QtGui',
        'win32com.client',
        'comtypes.client',
        'pptx',
        'fitz',
        'reportlab.pdfgen',
        'reportlab.lib',
        'PIL.Image',
    ],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='{PROJECT_NAME}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    {icon_option},
    version_file=None,
)
'''

    spec_file = f"{PROJECT_NAME}.spec"
    with open(spec_file, 'w', encoding='utf-8') as f:
        f.write(spec_content)

    logger.info(f"✓ 已创建 spec 文件: {spec_file}")
    return spec_file


def run_pyinstaller(spec_file):
    """运行 PyInstaller"""
    logger.info("开始 PyInstaller 构建...")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--clean",
        "--noconfirm",
        spec_file
    ]

    logger.info(f"执行命令: {' '.join(cmd)}")

    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        logger.info("✓ PyInstaller 构建成功")
        return True
    except subprocess.CalledProcessError as e:
        logger.error("PyInstaller 构建失败:")
        logger.error(f"返回码: {e.returncode}")
        logger.error(f"标准输出: {e.stdout}")
        logger.error(f"标准错误: {e.stderr}")
        return False


def verify_build():
    """验证构建结果"""
    exe_path = os.path.join(DIST_DIR, f"{PROJECT_NAME}.exe")

    if os.path.exists(exe_path):
        file_size = os.path.getsize(exe_path)
        logger.info(f"✓ 可执行文件已生成: {exe_path}")
        logger.info(f"  文件大小: {file_size / (1024 * 1024):.1f} MB")
        return True
    else:
        logger.error(f"可执行文件未找到: {exe_path}")
        return False


def show_usage_info():
    """显示使用说明"""
    exe_path = os.path.join(DIST_DIR, f"{PROJECT_NAME}.exe")

    logger.info("\n" + "=" * 50)
    logger.info("构建完成!")
    logger.info("=" * 50)
    logger.info(f"可执行文件位置: {exe_path}")
    logger.info(f"项目版本: {VERSION}")
    logger.info("\n使用说明:")
    logger.info("1. 双击运行 PPT2Manual.exe")
    logger.info("2. 添加要转换的 PPT 文件")
    logger.info("3. 选择输出路径")
    logger.info("4. 点击开始转换")
    logger.info("\n注意事项:")
    logger.info("- 确保已安装 Microsoft Office 或 LibreOffice")
    logger.info("- PPT文件路径不要包含特殊字符")
    logger.info("- 转换大文件时请耐心等待")


def main():
    """主构建流程"""
    logger.info(f"开始构建 {PROJECT_NAME} {VERSION}")
    logger.info("=" * 50)

    try:
        # 1. 清理构建目录
        clean_build_dirs()

        # 2. 检查依赖
        if not check_dependencies():
            sys.exit(1)

        # 3. 检查主脚本
        if not check_main_script():
            sys.exit(1)

        # 4. 创建 spec 文件
        spec_file = create_spec_file()

        # 5. 运行 PyInstaller
        if not run_pyinstaller(spec_file):
            sys.exit(1)

        # 6. 验证构建结果
        if not verify_build():
            sys.exit(1)

        # 7. 显示使用信息
        show_usage_info()

    except KeyboardInterrupt:
        logger.info("\n构建被用户中断")
        sys.exit(1)
    except Exception as e:
        logger.error(f"构建过程中发生错误: {e}")
        import traceback
        logger.error(traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    main()