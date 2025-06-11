#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT转换核心逻辑 - 修正类名引用
"""

import os
import sys
import time
import logging
from pathlib import Path
from typing import List, Optional, Tuple, Dict, Any
import tempfile
import traceback
import subprocess

from PySide6.QtCore import QThread, Signal

# 导入必要的库
try:
    import fitz  # PyMuPDF
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader
    from reportlab.lib.colors import black, lightgrey
    from PIL import Image
except ImportError as e:
    logging.error(f"导入PDF处理库失败: {e}")

from .utils import ProgressTracker, get_unique_temp_filename, clean_filename
from .pdf_merger import PDFMerger
from .slide_layout_generator import ChineseSlideLayoutGenerator  # 修正导入
from .pdf_layout_generator import PDFLayoutGenerator


class PPTConverter:
    """PPT转换器"""

    def __init__(self):
        self.temp_files = []

    def __del__(self):
        self.cleanup_temp_files()

    def cleanup_temp_files(self):
        """清理临时文件"""
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    if os.path.isdir(temp_file):
                        import shutil
                        shutil.rmtree(temp_file)
                    else:
                        os.remove(temp_file)
            except Exception as e:
                logging.warning(f"清理临时文件失败 {temp_file}: {e}")
        self.temp_files.clear()

    def convert_ppt_to_images(self, ppt_path: str) -> List[str]:
        """将PPT转换为图片列表"""
        try:
            logging.info(f"开始转换PPT: {ppt_path}")

            # 创建临时目录
            temp_dir = tempfile.mkdtemp(prefix="ppt_convert_")
            self.temp_files.append(temp_dir)

            # 方法1: 使用修正的COM接口
            image_files = self._convert_with_com_interface(ppt_path, temp_dir)

            if not image_files:
                # 方法2: 使用PowerShell COM调用
                logging.info("尝试使用PowerShell COM调用")
                image_files = self._convert_with_powershell_com(ppt_path, temp_dir)

            if not image_files:
                # 方法3: 通过PDF中间格式
                logging.info("尝试通过PDF中间格式转换")
                image_files = self._convert_via_pdf_intermediate(ppt_path, temp_dir)

            return image_files

        except Exception as e:
            logging.error(f"PPT转换失败: {e}")
            return []

    def _convert_with_com_interface(self, ppt_path: str, output_dir: str) -> List[str]:
        """使用修正的COM接口转换"""
        try:
            import win32com.client
            logging.info("使用 win32com.client 进行转换")

            # 创建PowerPoint应用程序实例
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = 1  # 设置为可见

            # 打开PPT文件
            presentation = powerpoint.Presentations.Open(
                os.path.abspath(ppt_path),
                ReadOnly=True,
                Untitled=True,
                WithWindow=False
            )

            image_files = []

            # 导出每张幻灯片
            for i in range(1, presentation.Slides.Count + 1):
                slide = presentation.Slides(i)
                image_filename = f"slide_{i:03d}.png"
                image_path = os.path.join(output_dir, image_filename)

                # 导出幻灯片为PNG
                slide.Export(image_path, "PNG", 1920, 1080)

                if os.path.exists(image_path):
                    image_files.append(image_path)
                    logging.debug(f"导出幻灯片 {i}: {image_path}")

            # 关闭演示文稿和应用程序
            presentation.Close()
            powerpoint.Quit()

            # 释放COM对象
            del presentation
            del powerpoint

            logging.info(f"COM接口转换完成，共 {len(image_files)} 张幻灯片")
            return image_files

        except ImportError:
            logging.error("win32com.client 未安装，请安装 pywin32")
            return []
        except Exception as e:
            logging.error(f"COM接口转换失败: {e}")
            try:
                # 尝试清理
                if 'presentation' in locals():
                    presentation.Close()
                if 'powerpoint' in locals():
                    powerpoint.Quit()
            except:
                pass
            return []

    def _convert_with_powershell_com(self, ppt_path: str, output_dir: str) -> List[str]:
        """使用PowerShell COM调用转换"""
        try:
            # PowerShell脚本内容
            ps_script = f'''
Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint

$ppt_path = "{ppt_path.replace(chr(92), chr(92) + chr(92))}"
$output_dir = "{output_dir.replace(chr(92), chr(92) + chr(92))}"

try {{
    $powerpoint = New-Object -ComObject PowerPoint.Application
    $powerpoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    $presentation = $powerpoint.Presentations.Open($ppt_path, $true, $true, $false)

    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {{
        $slide = $presentation.Slides.Item($i)
        $image_path = Join-Path $output_dir ("slide_" + $i.ToString("000") + ".png")
        $slide.Export($image_path, "PNG", 1920, 1080)
        Write-Host "Exported slide $i to $image_path"
    }}

    $presentation.Close()
    $powerpoint.Quit()

    Write-Host "Conversion completed successfully"
}} catch {{
    Write-Error "PowerShell conversion failed: $($_.Exception.Message)"
    exit 1
}} finally {{
    if ($presentation) {{ $presentation = $null }}
    if ($powerpoint) {{ $powerpoint = $null }}
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}}
'''

            # 创建临时PowerShell脚本文件
            ps_file = os.path.join(output_dir, "convert_ppt.ps1")
            with open(ps_file, 'w', encoding='utf-8') as f:
                f.write(ps_script)

            # 执行PowerShell脚本
            cmd = [
                "powershell.exe",
                "-ExecutionPolicy", "Bypass",
                "-NoProfile",
                "-File", ps_file
            ]

            logging.info("执行PowerShell COM转换...")
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)

            if result.returncode != 0:
                logging.error(f"PowerShell转换失败: {result.stderr}")
                return []

            # 收集生成的图片文件
            image_files = []
            for file in os.listdir(output_dir):
                if file.startswith("slide_") and file.endswith(".png"):
                    image_files.append(os.path.join(output_dir, file))

            image_files.sort()
            logging.info(f"PowerShell转换完成，共 {len(image_files)} 张幻灯片")
            return image_files

        except Exception as e:
            logging.error(f"PowerShell转换失败: {e}")
            return []

    def _convert_via_pdf_intermediate(self, ppt_path: str, output_dir: str) -> List[str]:
        """通过PDF中间格式转换"""
        try:
            # 先转换为PDF
            temp_pdf = os.path.join(output_dir, "temp.pdf")

            if self._convert_ppt_to_pdf_com(ppt_path, temp_pdf):
                # 再将PDF转换为图片
                return self._pdf_to_images(temp_pdf, output_dir)
            else:
                return []

        except Exception as e:
            logging.error(f"PDF中间格式转换失败: {e}")
            return []

    def _convert_ppt_to_pdf_com(self, ppt_path: str, pdf_path: str) -> bool:
        """使用COM将PPT转换为PDF"""
        try:
            import win32com.client

            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = 1

            presentation = powerpoint.Presentations.Open(
                os.path.abspath(ppt_path),
                ReadOnly=True
            )

            # 导出为PDF (32 = ppSaveAsPDF)
            presentation.SaveAs(os.path.abspath(pdf_path), 32)

            presentation.Close()
            powerpoint.Quit()

            # 清理COM对象
            del presentation
            del powerpoint

            success = os.path.exists(pdf_path)
            if success:
                logging.info(f"PPT转PDF成功: {pdf_path}")
            return success

        except ImportError:
            logging.error("win32com.client 未安装")
            return False
        except Exception as e:
            logging.error(f"PPT转PDF失败: {e}")
            return False

    def _pdf_to_images(self, pdf_path: str, output_dir: str) -> List[str]:
        """将PDF转换为高质量图片"""
        try:
            doc = fitz.open(pdf_path)
            image_files = []

            for page_num in range(doc.page_count):
                page = doc[page_num]

                # 使用高分辨率渲染
                zoom = 2.0  # 2倍缩放提高质量
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)

                # 保存为PNG
                image_filename = f"slide_{page_num + 1:03d}.png"
                image_path = os.path.join(output_dir, image_filename)
                pix.save(image_path)

                image_files.append(image_path)
                pix = None

            doc.close()
            logging.info(f"PDF转图片完成: {len(image_files)} 张")
            return image_files

        except Exception as e:
            logging.error(f"PDF转图片失败: {e}")
            return []


class MixedFileConverterMain:
    """混合文件转换器主类"""

    def __init__(self):
        self.ppt_converter = PPTConverter()
        self.slide_layout_generator = ChineseSlideLayoutGenerator()  # 修正类名
        self.pdf_layout_generator = PDFLayoutGenerator()

    def __del__(self):
        self.cleanup_temp_files()

    def cleanup_temp_files(self):
        """清理临时文件"""
        if hasattr(self, 'ppt_converter'):
            self.ppt_converter.cleanup_temp_files()

    def convert_ppt_to_layout_pdf(self, ppt_path: str, output_path: str) -> bool:
        """将PPT转换为布局PDF"""
        try:
            logging.info(f"开始转换PPT到布局PDF: {ppt_path}")

            # 将PPT转换为图片
            slide_images = self.ppt_converter.convert_ppt_to_images(ppt_path)

            if not slide_images:
                logging.error("没有成功转换的幻灯片图片")
                return False

            # 创建布局PDF
            title = Path(ppt_path).stem
            success = self.slide_layout_generator.create_layout_pdf(
                slide_images, output_path, title
            )

            if success:
                logging.info(f"PPT转换完成: {output_path}")

            return success

        except Exception as e:
            logging.error(f"PPT转换失败: {e}")
            return False

    def convert_pdf_to_layout_pdf(self, pdf_path: str, output_path: str) -> bool:
        """将PDF转换为8页布局PDF"""
        try:
            logging.info(f"开始转换PDF到布局PDF: {pdf_path}")

            title = Path(pdf_path).stem
            success = self.pdf_layout_generator.convert_pdf_to_layout(
                pdf_path, output_path, title
            )

            if success:
                logging.info(f"PDF布局转换完成: {output_path}")

            return success

        except Exception as e:
            logging.error(f"PDF布局转换失败: {e}")
            return False


class MixedFileConverterThread(QThread):
    """混合文件转换线程 - PDF也采用8页布局"""

    progress_updated = Signal(int, int)
    status_updated = Signal(str)
    conversion_finished = Signal(str)
    error_occurred = Signal(str)

    def __init__(self, file_list: List[Dict[str, Any]], output_path: str):
        super().__init__()
        self.file_list = file_list
        self.output_path = output_path
        self.is_cancelled = False
        self.converter = MixedFileConverterMain()
        self.merger = PDFMerger()

    def cancel(self):
        """取消转换"""
        self.is_cancelled = True

    def run(self):
        """运行混合文件转换"""
        try:
            self.status_updated.emit("初始化转换器...")

            if not self.file_list:
                self.error_occurred.emit("没有文件需要转换")
                return

            # 分离PPT和PDF文件
            ppt_files = [item for item in self.file_list if item['type'] == 'ppt']
            pdf_files = [item for item in self.file_list if item['type'] == 'pdf']

            logging.info(f"混合文件处理: {len(ppt_files)} 个PPT文件, {len(pdf_files)} 个PDF文件")

            # 创建进度跟踪器
            total_steps = len(ppt_files) + len(pdf_files) + 2  # 所有文件转换 + 合并 + 优化
            progress_tracker = ProgressTracker(total_steps)

            # 临时PDF文件列表
            temp_pdfs = []
            merge_info = []

            try:
                # 步骤1: 转换PPT文件为布局PDF
                for i, ppt_info in enumerate(ppt_files):
                    if self.is_cancelled:
                        self.status_updated.emit("转换已取消")
                        return

                    ppt_file = ppt_info['file']
                    file_name = Path(ppt_file).name
                    self.status_updated.emit(f"转换PPT文件 ({i + 1}/{len(ppt_files)}): {file_name}")

                    # 创建临时PDF文件
                    temp_pdf = get_unique_temp_filename("ppt_layout_", ".pdf")
                    temp_pdfs.append(temp_pdf)

                    # 转换PPT为布局PDF
                    success = self.converter.convert_ppt_to_layout_pdf(ppt_file, temp_pdf)
                    if not success:
                        self.error_occurred.emit(f"转换PPT失败: {file_name}")
                        return

                    # 添加到合并信息
                    merge_info.append({
                        'file': temp_pdf,
                        'title': ppt_info.get('title', Path(ppt_file).stem),
                        'order': ppt_info.get('order', 0),
                        'file_type': 'converted_ppt'
                    })

                    # 更新进度
                    progress_tracker.next_step()
                    progress = progress_tracker.get_progress_percentage()
                    self.progress_updated.emit(progress, 100)

                    self.status_updated.emit(f"完成PPT转换: {file_name}")
                    self.msleep(100)

                # 步骤2: 转换PDF文件为布局PDF
                for i, pdf_info in enumerate(pdf_files):
                    if self.is_cancelled:
                        self.status_updated.emit("转换已取消")
                        return

                    pdf_file = pdf_info['file']
                    file_name = Path(pdf_file).name
                    self.status_updated.emit(f"转换PDF布局 ({i + 1}/{len(pdf_files)}): {file_name}")

                    # 创建临时布局PDF文件
                    temp_pdf = get_unique_temp_filename("pdf_layout_", ".pdf")
                    temp_pdfs.append(temp_pdf)

                    # 转换PDF为布局PDF
                    success = self.converter.convert_pdf_to_layout_pdf(pdf_file, temp_pdf)
                    if not success:
                        self.error_occurred.emit(f"转换PDF布局失败: {file_name}")
                        return

                    # 添加到合并信息
                    merge_info.append({
                        'file': temp_pdf,
                        'title': pdf_info.get('title', Path(pdf_file).stem),
                        'order': pdf_info.get('order', 0),
                        'file_type': 'converted_pdf'
                    })

                    # 更新进度
                    progress_tracker.next_step()
                    progress = progress_tracker.get_progress_percentage()
                    self.progress_updated.emit(progress, 100)

                    self.status_updated.emit(f"完成PDF布局转换: {file_name}")
                    self.msleep(100)

                if self.is_cancelled:
                    return

                # 根据order字段排序
                merge_info.sort(key=lambda x: x.get('order', 0))

                # 步骤3: 合并所有布局PDF文件
                self.status_updated.emit("合并PDF文件...")

                success = self.merger.merge_pdfs_with_bookmarks(merge_info, self.output_path)

                if not success:
                    self.error_occurred.emit("PDF合并失败")
                    return

                progress_tracker.next_step()
                progress = progress_tracker.get_progress_percentage()
                self.progress_updated.emit(progress, 100)

                if self.is_cancelled:
                    return

                # 步骤4: 优化PDF
                self.status_updated.emit("优化PDF文件...")

                self.merger.optimize_pdf(self.output_path)

                progress_tracker.next_step()
                self.progress_updated.emit(100, 100)

                # 完成
                total_files = len(ppt_files) + len(pdf_files)
                self.status_updated.emit(f"转换完成！已处理 {total_files} 个文件 (统一8页布局)")
                self.conversion_finished.emit(self.output_path)

            finally:
                # 清理临时文件
                self.status_updated.emit("清理临时文件...")
                for temp_pdf in temp_pdfs:
                    try:
                        if os.path.exists(temp_pdf):
                            os.remove(temp_pdf)
                    except Exception as e:
                        logging.warning(f"清理临时文件失败: {e}")

                self.converter.cleanup_temp_files()

        except Exception as e:
            error_msg = f"转换过程中发生错误: {str(e)}"
            logging.error(error_msg)
            logging.error(traceback.format_exc())
            self.error_occurred.emit(error_msg)


# 保持向后兼容性
class PPTConverterThread(MixedFileConverterThread):
    """PPT转换线程 - 向后兼容版本"""

    def __init__(self, ppt_files: List[str], output_path: str):
        # 将旧格式转换为新格式
        file_list = []
        for i, ppt_file in enumerate(ppt_files):
            file_list.append({
                'file': ppt_file,
                'title': Path(ppt_file).stem,
                'type': 'ppt',
                'order': i
            })

        super().__init__(file_list, output_path)