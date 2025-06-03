#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT转换核心逻辑 - 修正 COM 接口版本
"""
import io
import os
import sys
import time
import logging
from pathlib import Path
from typing import List, Optional, Tuple
import tempfile
import traceback
import subprocess

from PySide6.QtCore import QThread, Signal
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

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


class PPTConverter:
    """PPT转换器 - 修正版"""

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
                slide.Export(image_path, "PNG", 1440, 960)  # 使用高分辨率

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


class PPTConverterMain:
    """主PPT转换器"""

    def __init__(self):
        self.ppt_converter = PPTConverter()
        self.layout_generator = SlideLayoutGenerator()

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
            success = self.layout_generator.create_layout_pdf(
                slide_images, output_path, title
            )

            if success:
                logging.info(f"PPT转换完成: {output_path}")

            return success

        except Exception as e:
            logging.error(f"PPT转换失败: {e}")
            return False


class PPTConverterThread(QThread):
    """PPT转换线程"""

    progress_updated = Signal(int, int)
    status_updated = Signal(str)
    conversion_finished = Signal(str)
    error_occurred = Signal(str)

    def __init__(self, ppt_files: List[str], output_path: str):
        super().__init__()
        self.ppt_files = ppt_files
        self.output_path = output_path
        self.is_cancelled = False
        self.converter = PPTConverterMain()
        self.merger = PDFMerger()

    def cancel(self):
        """取消转换"""
        self.is_cancelled = True

    def run(self):
        """运行转换"""
        try:
            self.status_updated.emit("初始化转换器...")

            if not self.ppt_files:
                self.error_occurred.emit("没有PPT文件需要转换")
                return

            # 创建进度跟踪器
            total_steps = len(self.ppt_files) + 2
            progress_tracker = ProgressTracker(total_steps)

            # 临时PDF文件列表
            temp_pdfs = []

            try:
                # 步骤1: 转换每个PPT文件
                for i, ppt_file in enumerate(self.ppt_files):
                    if self.is_cancelled:
                        self.status_updated.emit("转换已取消")
                        return

                    file_name = Path(ppt_file).name
                    self.status_updated.emit(f"转换PPT文件 ({i + 1}/{len(self.ppt_files)}): {file_name}")

                    # 创建临时PDF文件
                    temp_pdf = get_unique_temp_filename("ppt_layout_", ".pdf")
                    temp_pdfs.append(temp_pdf)

                    # 转换PPT为布局PDF
                    success = self.converter.convert_ppt_to_layout_pdf(ppt_file, temp_pdf)
                    if not success:
                        self.error_occurred.emit(f"转换PPT失败: {file_name}")
                        return

                    # 更新进度
                    progress_tracker.next_step()
                    progress = progress_tracker.get_progress_percentage()
                    self.progress_updated.emit(progress, 100)

                    self.status_updated.emit(f"完成: {file_name}")
                    self.msleep(100)

                if self.is_cancelled:
                    return

                # 步骤2: 合并PDF文件
                self.status_updated.emit("合并PDF文件...")

                # 准备合并信息
                merge_info = []
                for i, (ppt_file, pdf_file) in enumerate(zip(self.ppt_files, temp_pdfs)):
                    title = Path(ppt_file).stem
                    merge_info.append({
                        'file': pdf_file,
                        'title': title,
                        'order': i
                    })

                # 执行合并
                success = self.merger.merge_pdfs_with_bookmarks(
                    merge_info, self.output_path
                )

                if not success:
                    self.error_occurred.emit("PDF合并失败")
                    return

                progress_tracker.next_step()
                progress = progress_tracker.get_progress_percentage()
                self.progress_updated.emit(progress, 100)

                if self.is_cancelled:
                    return

                # 步骤3: 后处理
                self.status_updated.emit("优化PDF文件...")

                # PDF优化
                self.merger.optimize_pdf(self.output_path)

                progress_tracker.next_step()
                self.progress_updated.emit(100, 100)

                # 完成
                self.status_updated.emit(f"转换完成！已生成 {len(self.ppt_files)} 个PPT的合并PDF")
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


class SlideLayoutGenerator:
    """修正版幻灯片布局生成器 - 正确比例和页码"""

    def __init__(self):
        self.page_width, self.page_height = A4
        self.margin = 36
        self.gap = 12

        # 布局参数
        self.slides_per_row = 2
        self.slides_per_col = 4
        self.slides_per_page = 8

        # 计算每个幻灯片槽位的尺寸
        available_width = self.page_width - 2 * self.margin - (self.slides_per_row - 1) * self.gap
        available_height = self.page_height - 2 * self.margin - (self.slides_per_col - 1) * self.gap - 40  # 为页码留更多空间

        self.slot_width = available_width / self.slides_per_row
        self.slot_height = available_height / self.slides_per_col

        # 注册中文字体
        self.font_name = self._register_chinese_font()

        logging.info(f"幻灯片槽位尺寸: {self.slot_width:.1f} x {self.slot_height:.1f} points")

    def _register_chinese_font(self) -> str:
        """注册中文字体"""
        try:
            font_paths = [
                r"C:\Windows\Fonts\simhei.ttf",
                r"C:\Windows\Fonts\simsun.ttc",
                r"C:\Windows\Fonts\msyh.ttc",
            ]

            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        if "simhei" in font_path.lower():
                            pdfmetrics.registerFont(TTFont("SimHei", font_path))
                            return "SimHei"
                        elif "simsun" in font_path.lower():
                            pdfmetrics.registerFont(TTFont("SimSun", font_path))
                            return "SimSun"
                        elif "msyh" in font_path.lower():
                            pdfmetrics.registerFont(TTFont("MSYaHei", font_path))
                            return "MSYaHei"
                    except Exception as e:
                        logging.debug(f"注册字体失败 {font_path}: {e}")
                        continue

            logging.warning("未找到中文字体，使用默认字体")
            return "Helvetica"

        except Exception as e:
            logging.error(f"字体注册失败: {e}")
            return "Helvetica"

    def create_layout_pdf(self, slide_images: List[str], output_path: str, title: str = "") -> bool:
        """创建布局PDF - 不添加页码（由合并器统一处理）"""
        try:
            logging.info(f"创建布局PDF: {len(slide_images)} 张幻灯片")

            c = canvas.Canvas(output_path, pagesize=A4)

            # 设置文档信息
            c.setTitle(title if title else "PPT幻灯片")
            c.setAuthor("PPT2Manual")
            c.setSubject("PPT转换PDF")

            page_num = 1

            # 按每页8张幻灯片分组处理
            for page_start in range(0, len(slide_images), self.slides_per_page):
                if page_start > 0:
                    c.showPage()

                # 当前页面的幻灯片
                page_slides = slide_images[page_start:page_start + self.slides_per_page]

                logging.info(f"处理第 {page_num} 页，包含 {len(page_slides)} 张幻灯片")

                # 绘制幻灯片
                for position, slide_path in enumerate(page_slides):
                    if os.path.exists(slide_path):
                        self._draw_slide_with_correct_aspect_ratio(c, slide_path, position)
                    else:
                        logging.warning(f"幻灯片图片不存在: {slide_path}")
                        self._draw_error_placeholder(c, position, f"缺失: {os.path.basename(slide_path)}")

                # 不在这里添加页码，由PDF合并器统一处理
                page_num += 1

            c.save()
            logging.info(f"布局PDF创建完成: {output_path}")
            return True

        except Exception as e:
            logging.error(f"创建布局PDF失败: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False

    def _calculate_slot_position(self, position: int) -> Tuple[float, float]:
        """计算槽位的左下角坐标"""
        row = position // self.slides_per_row
        col = position % self.slides_per_row

        x = self.margin + col * (self.slot_width + self.gap)
        y = self.page_height - self.margin - (row + 1) * (self.slot_height + self.gap)

        return x, y

    def _draw_slide_with_correct_aspect_ratio(self, canvas_obj, slide_path: str, position: int):
        """绘制保持正确宽高比的单张幻灯片"""
        try:
            # 计算槽位位置
            slot_x, slot_y = self._calculate_slot_position(position)

            # 获取图片的真实尺寸
            with Image.open(slide_path) as img:
                if img.mode != 'RGB':
                    img = img.convert('RGB')

                original_width, original_height = img.size
                original_aspect_ratio = original_width / original_height

                logging.debug(
                    f"幻灯片 {position + 1}: 原始尺寸 {original_width}x{original_height}, 宽高比 {original_aspect_ratio:.3f}")

                # 计算槽位的宽高比
                slot_aspect_ratio = self.slot_width / self.slot_height

                # 计算实际显示尺寸，保持宽高比
                if original_aspect_ratio > slot_aspect_ratio:
                    # 图片更宽，以宽度为准
                    display_width = self.slot_width
                    display_height = self.slot_width / original_aspect_ratio
                else:
                    # 图片更高，以高度为准
                    display_height = self.slot_height
                    display_width = self.slot_height * original_aspect_ratio

                # 计算居中位置
                img_x = slot_x + (self.slot_width - display_width) / 2
                img_y = slot_y + (self.slot_height - display_height) / 2

                logging.debug(
                    f"幻灯片 {position + 1}: 显示尺寸 {display_width:.1f}x{display_height:.1f}, 位置 ({img_x:.1f}, {img_y:.1f})")

                # 处理图片尺寸以提高性能
                max_display_dimension = max(display_width, display_height) * 2  # 2倍分辨率确保清晰度
                if max(original_width, original_height) > max_display_dimension * 2:
                    # 只有当原图远大于显示尺寸时才缩放
                    scale_factor = (max_display_dimension * 2) / max(original_width, original_height)
                    new_width = int(original_width * scale_factor)
                    new_height = int(original_height * scale_factor)
                    img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    logging.debug(f"图片预缩放: {original_width}x{original_height} -> {new_width}x{new_height}")

                # 创建ImageReader
                img_buffer = io.BytesIO()
                img.save(img_buffer, format='PNG', quality=95)
                img_buffer.seek(0)
                image_reader = ImageReader(img_buffer)

                # 绘制槽位边框（调试用，可选）
                canvas_obj.setStrokeColor((0.95, 0.95, 0.95))
                canvas_obj.setLineWidth(0.25)
                canvas_obj.rect(slot_x, slot_y, self.slot_width, self.slot_height)

                # 绘制图片
                canvas_obj.drawImage(
                    image_reader,
                    img_x, img_y,
                    width=display_width,
                    height=display_height
                )

                # 绘制图片边框
                canvas_obj.setStrokeColor(lightgrey)
                canvas_obj.setLineWidth(0.5)
                canvas_obj.rect(img_x, img_y, display_width, display_height)

        except Exception as e:
            logging.error(f"绘制幻灯片失败 {slide_path}: {e}")
            import traceback
            logging.error(traceback.format_exc())
            self._draw_error_placeholder(canvas_obj, position, f"错误: {os.path.basename(slide_path)}")

    def _draw_error_placeholder(self, canvas_obj, position: int, error_text: str):
        """绘制错误占位符"""
        try:
            slot_x, slot_y = self._calculate_slot_position(position)

            # 绘制灰色背景
            canvas_obj.setFillColor(lightgrey)
            canvas_obj.rect(slot_x, slot_y, self.slot_width, self.slot_height, fill=1)

            # 绘制错误文本
            canvas_obj.setFillColor(black)
            canvas_obj.setFont(self.font_name, 8)

            # 计算文本位置（居中）
            text_width = canvas_obj.stringWidth(error_text, self.font_name, 8)
            text_x = slot_x + (self.slot_width - text_width) / 2
            text_y = slot_y + self.slot_height / 2

            canvas_obj.drawString(text_x, text_y, error_text)

            # 绘制边框
            canvas_obj.setStrokeColor(black)
            canvas_obj.setLineWidth(0.5)
            canvas_obj.rect(slot_x, slot_y, self.slot_width, self.slot_height)

        except Exception as e:
            logging.error(f"绘制错误占位符失败: {e}")