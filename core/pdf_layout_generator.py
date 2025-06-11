#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF布局生成器 - 将PDF页面转换为8页布局
"""

import os
import logging
import tempfile
from pathlib import Path
from typing import List, Optional

try:
    import fitz  # PyMuPDF
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.colors import lightgrey, black
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from PIL import Image
    import io
except ImportError as e:
    logging.error(f"导入PDF处理库失败: {e}")

from .utils import get_unique_temp_filename


class PDFLayoutGenerator:
    """PDF布局生成器 - 8页布局"""

    def __init__(self):
        self.page_width, self.page_height = A4
        self.margin = 36
        self.gap = 12

        # 布局参数
        self.pages_per_row = 2
        self.pages_per_col = 4
        self.pages_per_layout_page = 8

        # 计算每个页面槽位的尺寸
        available_width = self.page_width - 2 * self.margin - (self.pages_per_row - 1) * self.gap
        available_height = self.page_height - 2 * self.margin - (self.pages_per_col - 1) * self.gap - 40

        self.slot_width = available_width / self.pages_per_row
        self.slot_height = available_height / self.pages_per_col

        # 注册中文字体
        self.font_name = self._register_chinese_font()

        logging.info(f"PDF页面槽位尺寸: {self.slot_width:.1f} x {self.slot_height:.1f} points")

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

    def convert_pdf_to_layout(self, input_pdf_path: str, output_pdf_path: str, title: str = "") -> bool:
        """将PDF转换为8页布局"""
        try:
            logging.info(f"开始转换PDF为布局: {input_pdf_path}")

            # 先将PDF转换为图片
            pdf_images = self._pdf_to_images(input_pdf_path)

            if not pdf_images:
                logging.error("PDF转图片失败")
                return False

            # 创建布局PDF
            success = self._create_layout_pdf_from_images(pdf_images, output_pdf_path, title)

            # 清理临时图片
            for img_path in pdf_images:
                try:
                    if os.path.exists(img_path):
                        os.remove(img_path)
                except Exception as e:
                    logging.warning(f"清理临时图片失败: {e}")

            return success

        except Exception as e:
            logging.error(f"PDF布局转换失败: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False

    def _pdf_to_images(self, pdf_path: str) -> List[str]:
        """将PDF转换为图片列表"""
        try:
            doc = fitz.open(pdf_path)
            image_files = []

            # 创建临时目录
            temp_dir = tempfile.mkdtemp(prefix="pdf_layout_")

            for page_num in range(doc.page_count):
                page = doc[page_num]

                # 使用高分辨率渲染
                zoom = 2.0  # 2倍缩放提高质量
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)

                # 保存为PNG
                image_filename = f"page_{page_num + 1:03d}.png"
                image_path = os.path.join(temp_dir, image_filename)
                pix.save(image_path)

                image_files.append(image_path)
                pix = None

            doc.close()
            logging.info(f"PDF转图片完成: {len(image_files)} 张")
            return image_files

        except Exception as e:
            logging.error(f"PDF转图片失败: {e}")
            return []

    def _create_layout_pdf_from_images(self, image_list: List[str], output_path: str, title: str) -> bool:
        """从图片列表创建布局PDF"""
        try:
            logging.info(f"创建PDF布局: {len(image_list)} 张页面")

            c = canvas.Canvas(output_path, pagesize=A4)

            # 设置文档信息
            c.setTitle(title if title else "PDF布局")
            c.setAuthor("PPT2Manual")
            c.setSubject("PDF转换布局")

            layout_page_num = 1

            # 按每页8张页面分组处理
            for page_start in range(0, len(image_list), self.pages_per_layout_page):
                if page_start > 0:
                    c.showPage()

                # 当前布局页的页面
                layout_pages = image_list[page_start:page_start + self.pages_per_layout_page]

                logging.info(f"处理布局页 {layout_page_num}，包含 {len(layout_pages)} 张页面")

                # 绘制页面
                for position, image_path in enumerate(layout_pages):
                    if os.path.exists(image_path):
                        self._draw_page_with_aspect_ratio(c, image_path, position, page_start + position + 1)
                    else:
                        logging.warning(f"页面图片不存在: {image_path}")
                        self._draw_error_placeholder(c, position, f"缺失页面 {page_start + position + 1}")

                layout_page_num += 1

            c.save()
            logging.info(f"PDF布局创建完成: {output_path}")
            return True

        except Exception as e:
            logging.error(f"创建PDF布局失败: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False

    def _calculate_slot_position(self, position: int):
        """计算槽位的左下角坐标"""
        row = position // self.pages_per_row
        col = position % self.pages_per_row

        x = self.margin + col * (self.slot_width + self.gap)
        y = self.page_height - self.margin - (row + 1) * (self.slot_height + self.gap)

        return x, y

    def _draw_page_with_aspect_ratio(self, canvas_obj, image_path: str, position: int, page_number: int):
        """绘制保持正确宽高比的单张页面"""
        try:
            # 计算槽位位置
            slot_x, slot_y = self._calculate_slot_position(position)

            # 获取图片的真实尺寸
            with Image.open(image_path) as img:
                if img.mode != 'RGB':
                    img = img.convert('RGB')

                original_width, original_height = img.size
                original_aspect_ratio = original_width / original_height

                logging.debug(
                    f"页面 {page_number}: 原始尺寸 {original_width}x{original_height}, 宽高比 {original_aspect_ratio:.3f}")

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
                    f"页面 {page_number}: 显示尺寸 {display_width:.1f}x{display_height:.1f}, 位置 ({img_x:.1f}, {img_y:.1f})")

                # 处理图片尺寸以提高性能
                max_display_dimension = max(display_width, display_height) * 2
                if max(original_width, original_height) > max_display_dimension * 2:
                    scale_factor = (max_display_dimension * 2) / max(original_width, original_height)
                    new_width = int(original_width * scale_factor)
                    new_height = int(original_height * scale_factor)
                    img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    logging.debug(f"图片预缩放: {original_width}x{original_height} -> {new_width}x{new_height}")

                # 创建ImageReader
                from reportlab.lib.utils import ImageReader
                img_buffer = io.BytesIO()
                img.save(img_buffer, format='PNG', quality=95)
                img_buffer.seek(0)
                image_reader = ImageReader(img_buffer)

                # 绘制槽位边框（调试用）
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

                # 添加页码标签（在图片下方）
                self._add_page_label(canvas_obj, slot_x, slot_y, page_number)

        except Exception as e:
            logging.error(f"绘制页面失败 {image_path}: {e}")
            import traceback
            logging.error(traceback.format_exc())
            self._draw_error_placeholder(canvas_obj, position, f"错误: 页面 {page_number}")

    def _add_page_label(self, canvas_obj, slot_x: float, slot_y: float, page_number: int):
        """在页面下方添加页码标签"""
        try:
            canvas_obj.setFont(self.font_name, 8)
            canvas_obj.setFillColor(black)

            label_text = f"p.{page_number}"

            # 计算标签位置（槽位底部中央）
            text_width = canvas_obj.stringWidth(label_text, self.font_name, 8)
            text_x = slot_x + (self.slot_width - text_width) / 2
            text_y = slot_y - 15  # 在槽位下方

            canvas_obj.drawString(text_x, text_y, label_text)

        except Exception as e:
            logging.debug(f"添加页码标签失败: {e}")

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