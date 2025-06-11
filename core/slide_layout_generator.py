#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
修正幻灯片比例和页码的布局生成器
"""

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import lightgrey, black
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image
import os
import logging
import io
from typing import List, Optional, Tuple


class ChineseSlideLayoutGenerator:
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


# 为了向后兼容，创建别名
SlideLayoutGenerator = ChineseSlideLayoutGenerator