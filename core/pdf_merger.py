#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
修正页码位置的PDF合并器
"""

import os
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional
import tempfile

try:
    import fitz  # PyMuPDF
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.colors import black, lightgrey
except ImportError as e:
    logging.error(f"导入PDF处理库失败: {e}")

from .utils import clean_filename, get_unique_temp_filename


class ChineseFontManager:
    """中文字体管理器"""

    def __init__(self):
        self.font_registered = False
        self.font_name = "SimHei"

    def register_chinese_font(self, canvas_obj):
        """注册中文字体"""
        if self.font_registered:
            return self.font_name

        try:
            font_paths = [
                r"C:\Windows\Fonts\simhei.ttf",
                r"C:\Windows\Fonts\SimHei.ttf",
                r"C:\Windows\Fonts\simsun.ttc",
                r"C:\Windows\Fonts\SimSun.ttc",
                r"C:\Windows\Fonts\msyh.ttc",
                r"C:\Windows\Fonts\msyhbd.ttc",
                r"C:\Windows\Fonts\ARIALUNI.TTF",
                r"C:\Program Files\Microsoft Office\root\VFS\Fonts\private\ARIALUNI.TTF",
            ]

            font_found = False
            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        if font_path.endswith('.ttc'):
                            font_names = ["SimSun", "SimHei", "MSYaHei"]
                            for font_name in font_names:
                                try:
                                    pdfmetrics.registerFont(TTFont(font_name, font_path))
                                    self.font_name = font_name
                                    font_found = True
                                    break
                                except:
                                    continue
                        else:
                            if "simhei" in font_path.lower():
                                pdfmetrics.registerFont(TTFont("SimHei", font_path))
                                self.font_name = "SimHei"
                            elif "simsun" in font_path.lower():
                                pdfmetrics.registerFont(TTFont("SimSun", font_path))
                                self.font_name = "SimSun"
                            elif "msyh" in font_path.lower():
                                pdfmetrics.registerFont(TTFont("MSYaHei", font_path))
                                self.font_name = "MSYaHei"
                            elif "arial" in font_path.lower():
                                pdfmetrics.registerFont(TTFont("ArialUnicode", font_path))
                                self.font_name = "ArialUnicode"
                            font_found = True

                        if font_found:
                            logging.info(f"成功注册中文字体: {self.font_name} ({font_path})")
                            break

                    except Exception as e:
                        logging.debug(f"注册字体失败 {font_path}: {e}")
                        continue

            if not font_found:
                logging.warning("未找到中文字体，将使用默认字体")
                self.font_name = "Helvetica"

            self.font_registered = True
            return self.font_name

        except Exception as e:
            logging.error(f"注册中文字体失败: {e}")
            self.font_name = "Helvetica"
            self.font_registered = True
            return self.font_name


class PDFMerger:
    """PDF合并器 - 修正页码版本"""

    def __init__(self):
        self.temp_files = []
        self.font_manager = ChineseFontManager()

    def __del__(self):
        """清理临时文件"""
        self.cleanup_temp_files()

    def cleanup_temp_files(self):
        """清理临时文件"""
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
            except Exception as e:
                logging.warning(f"清理临时文件失败 {temp_file}: {e}")
        self.temp_files.clear()

    def merge_pdfs_with_bookmarks(self, pdf_info_list: List[Dict[str, Any]], output_path: str) -> bool:
        """合并PDF文件并添加书签（统一页码）"""
        try:
            logging.info(f"开始合并PDF文件，共 {len(pdf_info_list)} 个文件")

            # 排序
            pdf_info_list.sort(key=lambda x: x.get('order', 0))

            # 第一步：收集所有PDF信息，计算页码分布
            pdf_page_info = []
            total_content_pages = 0

            for i, pdf_info in enumerate(pdf_info_list):
                pdf_file = pdf_info['file']
                title = pdf_info.get('title', f'文档{i + 1}')

                if os.path.exists(pdf_file):
                    try:
                        temp_doc = fitz.open(pdf_file)
                        page_count = temp_doc.page_count
                        temp_doc.close()

                        pdf_page_info.append({
                            'title': title,
                            'file': pdf_file,
                            'page_count': page_count,
                            'start_page_in_final': total_content_pages + 2,  # +1为目录页，+1因为从1开始计数
                            'order': i
                        })

                        total_content_pages += page_count
                        logging.info(f"PDF {i + 1}: {title} - {page_count} 页")

                    except Exception as e:
                        logging.error(f"无法读取PDF文件 {pdf_file}: {e}")

            if not pdf_page_info:
                logging.error("没有有效的PDF文件可以合并")
                return False

            logging.info(f"总共内容页数: {total_content_pages}")

            # 第二步：创建目录页
            toc_pdf_path = self._create_table_of_contents_pdf_with_accurate_pages(pdf_page_info)

            # 第三步：创建输出PDF文档并合并
            output_doc = fitz.open()
            bookmarks = []
            current_page = 0

            # 添加目录页
            if toc_pdf_path:
                try:
                    toc_doc = fitz.open(toc_pdf_path)
                    output_doc.insert_pdf(toc_doc)
                    current_page += toc_doc.page_count
                    toc_doc.close()
                    logging.info(f"已添加目录页，当前页数: {current_page}")
                except Exception as e:
                    logging.error(f"添加目录页失败: {e}")

            # 合并每个PDF文件
            for pdf_info in pdf_page_info:
                pdf_file = pdf_info['file']
                title = pdf_info['title']

                try:
                    logging.info(f"合并PDF: {title}")

                    src_doc = fitz.open(pdf_file)

                    # 记录书签位置
                    bookmarks.append({
                        'title': title,
                        'page': current_page,
                        'level': 1
                    })

                    # 清除源PDF中的页码（如果有的话）
                    self._remove_existing_page_numbers(src_doc)

                    # 插入页面
                    output_doc.insert_pdf(src_doc)
                    current_page += src_doc.page_count

                    src_doc.close()

                    logging.info(f"完成合并: {title}，当前总页数: {current_page}")

                except Exception as e:
                    logging.error(f"合并PDF文件失败 {pdf_file}: {e}")
                    continue

            # 第四步：添加书签
            self._add_bookmarks(output_doc, bookmarks)

            # 第五步：添加统一的页码到页面底部
            self._add_bottom_page_numbers(output_doc)

            # 第六步：设置文档元数据
            self._set_document_metadata(output_doc, len(pdf_page_info))

            # 第七步：保存文档
            output_doc.save(output_path)
            output_doc.close()

            # 清理临时文件
            if toc_pdf_path and os.path.exists(toc_pdf_path):
                try:
                    os.remove(toc_pdf_path)
                except:
                    pass

            logging.info(f"PDF合并完成: {output_path}，总页数: {current_page}")
            return True

        except Exception as e:
            logging.error(f"PDF合并失败: {e}")
            import traceback
            logging.error(f"错误详情: {traceback.format_exc()}")
            return False

    def _remove_existing_page_numbers(self, doc: fitz.Document):
        """移除现有的页码（尝试清除可能存在的页码）"""
        try:
            for page_num in range(doc.page_count):
                page = doc[page_num]

                # 获取页面底部区域的文本（可能的页码位置）
                bottom_rect = fitz.Rect(0, page.rect.height - 50, page.rect.width, page.rect.height)
                text_instances = page.search_for("第", clip=bottom_rect)

                # 如果找到"第"字，尝试移除相关文本
                for inst in text_instances:
                    try:
                        # 这里可以添加更复杂的页码检测和移除逻辑
                        pass
                    except:
                        pass

        except Exception as e:
            logging.debug(f"移除现有页码时出错: {e}")

    def _create_table_of_contents_pdf_with_accurate_pages(self, pdf_page_info: List[Dict[str, Any]]) -> Optional[str]:
        """创建带有准确页码的目录页"""
        try:
            temp_toc_pdf = get_unique_temp_filename("toc_", ".pdf")
            self.temp_files.append(temp_toc_pdf)

            c = canvas.Canvas(temp_toc_pdf, pagesize=A4)
            page_width, page_height = A4

            font_name = self.font_manager.register_chinese_font(c)

            margin = 50
            line_height = 25
            current_y = page_height - margin - 80

            # 绘制标题
            c.setFont(font_name, 24)
            title_text = "目录"
            title_width = c.stringWidth(title_text, font_name, 24)
            title_x = (page_width - title_width) / 2
            c.drawString(title_x, current_y + 40, title_text)

            # 绘制分割线
            #c.setStrokeColor(lightgrey)
            #c.setLineWidth(1)
            #c.line(margin, current_y + 10, page_width - margin, current_y + 10)

            # 目录项
            c.setFont(font_name, 12)
            c.setFillColor(black)

            for i, info in enumerate(pdf_page_info):
                title = info['title']
                start_page = info['start_page_in_final']

                if len(title) > 40:
                    title = title[:37] + "..."

                toc_text = f"{i + 1}. {title}"
                page_text = f"第 {start_page} 页"

                c.drawString(margin, current_y, toc_text)

                page_text_width = c.stringWidth(page_text, font_name, 12)
                c.drawString(page_width - margin - page_text_width, current_y, page_text)

                dots_start_x = margin + c.stringWidth(toc_text, font_name, 12) + 10
                dots_end_x = page_width - margin - page_text_width - 10
                self._draw_dots_reportlab(c, dots_start_x, dots_end_x, current_y + 3)

                current_y -= line_height

                if current_y < margin + 50:
                    c.showPage()
                    current_y = page_height - margin
                    c.setFont(font_name, 12)
                    c.setFillColor(black)

            c.save()

            if os.path.exists(temp_toc_pdf):
                logging.info(f"目录页创建成功: {temp_toc_pdf}")
                return temp_toc_pdf
            else:
                logging.error("目录页文件未生成")
                return None

        except Exception as e:
            logging.error(f"创建目录页失败: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return None

    def _draw_dots_reportlab(self, canvas_obj, start_x: float, end_x: float, y: float):
        """使用reportlab绘制点线"""
        try:
            if start_x >= end_x:
                return

            canvas_obj.setStrokeColor(lightgrey)
            canvas_obj.setLineWidth(0.5)

            dot_spacing = 4
            x = start_x
            while x < end_x:
                canvas_obj.circle(x, y, 0.5, fill=1)
                x += dot_spacing

        except Exception as e:
            logging.debug(f"绘制点线失败: {e}")

    def _add_bookmarks(self, doc: fitz.Document, bookmarks: List[Dict[str, Any]]):
        """添加书签"""
        try:
            if not bookmarks:
                return

            toc = []
            for bookmark in bookmarks:
                toc.append([
                    bookmark['level'],
                    bookmark['title'],
                    bookmark['page'] + 1
                ])

            doc.set_toc(toc)
            logging.info(f"添加了 {len(bookmarks)} 个书签")

        except Exception as e:
            logging.error(f"添加书签失败: {e}")

    def _add_bottom_page_numbers(self, doc: fitz.Document):
        """在页面底部添加统一的连续页码"""
        try:
            total_pages = doc.page_count

            for page_num in range(total_pages):
                page = doc[page_num]

                # 页码文本
                page_text = f"- {page_num + 1} -"

                # 计算页面底部中央位置
                rect = page.rect
                text_width_estimate = len(page_text) * 5  # 估算文本宽度
                text_x = (rect.width - text_width_estimate) / 2
                text_y = rect.height - 40  # 距离底部25个点

                # 添加页码到页面底部
                page.insert_text(
                    (text_x, text_y),
                    page_text,
                    fontsize=10,
                    fontname="helv",
                    color=(0, 0, 0)
                )

            logging.info(f"添加底部页码完成，总页数: {total_pages}")

        except Exception as e:
            logging.error(f"添加页码失败: {e}")
            import traceback
            logging.error(traceback.format_exc())

    def _set_document_metadata(self, doc: fitz.Document, pdf_count: int):
        """设置文档元数据"""
        try:
            metadata = {
                'title': '合并PDF手册',
                'author': 'PPT2Manual',
                'subject': f'由 {pdf_count} 个PPT文件合并生成的PDF手册',
                'creator': 'PPT2Manual v0.0.1-alpha',
                'producer': 'PyMuPDF',
                'keywords': 'PPT, PDF, 手册, 合并'
            }

            doc.set_metadata(metadata)
            logging.info("设置文档元数据完成")

        except Exception as e:
            logging.error(f"设置文档元数据失败: {e}")

    def optimize_pdf(self, input_path: str, output_path: str = None) -> bool:
        """优化PDF文件"""
        try:
            if output_path is None:
                temp_file = get_unique_temp_filename("optimized_", ".pdf")

                doc = fitz.open(input_path)
                doc.save(
                    temp_file,
                    garbage=4,
                    deflate=True,
                    clean=True
                )
                doc.close()

                import shutil
                shutil.move(temp_file, input_path)

            else:
                doc = fitz.open(input_path)
                doc.save(
                    output_path,
                    garbage=4,
                    deflate=True,
                    clean=True
                )
                doc.close()

            logging.info("PDF优化完成")
            return True

        except Exception as e:
            logging.error(f"PDF优化失败: {e}")
            return False