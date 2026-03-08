"""
文档处理模块
"""

import os
from typing import List, Dict, Tuple
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX
from document_converter import DocumentConverter
from logging_config import get_logger

logger = get_logger("document")


class DocumentProcessor:
    """文档处理器 - 支持.doc和.docx格式"""

    def __init__(self):
        self.document = None
        self.paragraphs = []
        self.text_content = ""
        self.converter = DocumentConverter()
        self.temp_docx_path = None
        self.was_converted = False

    def load_document(self, file_path: str) -> bool:
        """加载文档，自动处理.doc和.docx格式"""
        try:
            # 检测文档格式并确保为.docx
            docx_path, was_converted = self.converter.ensure_docx_format(file_path)

            # 保存转换信息用于清理
            self.temp_docx_path = docx_path if was_converted else None
            self.was_converted = was_converted

            if was_converted:
                logger.info("已将.doc文件转换为.docx格式进行处理")

            # 加载.docx文档
            self.document = Document(docx_path)
            self._extract_text()

            return True

        except Exception as e:
            logger.error(f"加载文档失败: {e}")
            self._cleanup_temp_files()
            return False

    def _cleanup_temp_files(self):
        """清理临时文件"""
        if self.temp_docx_path and self.was_converted:
            self.converter.cleanup_temp_files(self.temp_docx_path, self.was_converted)
            self.temp_docx_path = None
            self.was_converted = False

    def _extract_text(self):
        """提取文档中的文本内容"""
        self.paragraphs = []
        all_text = []

        for paragraph in self.document.paragraphs:
            if paragraph.text.strip():
                self.paragraphs.append({
                    'text': paragraph.text,
                    'paragraph': paragraph,
                    'index': len(self.paragraphs)
                })
                all_text.append(paragraph.text)

        self.text_content = '\n'.join(all_text)

    def highlight_text(self, paragraph_index: int, text: str,
                      color: WD_COLOR_INDEX = WD_COLOR_INDEX.YELLOW):
        """高亮显示文本"""
        try:
            if paragraph_index >= len(self.paragraphs):
                return False

            paragraph = self.paragraphs[paragraph_index]['paragraph']

            # 在段落中查找并高亮文本
            for run in paragraph.runs:
                if text in run.text:
                    run.font.highlight_color = color
                    return True

        except Exception as e:
            logger.error(f"高亮文本失败: {e}")

        return False

    def save_document(self, output_path: str) -> bool:
        """保存文档"""
        try:
            if self.document:
                self.document.save(output_path)

                # 如果原文件是.doc格式，保存后清理临时文件
                if self.was_converted:
                    logger.info(f"已将处理结果保存为.docx格式: {output_path}")

                return True
        except Exception as e:
            logger.error(f"保存文档失败: {e}")
        finally:
            # 清理临时文件
            self._cleanup_temp_files()

        return False

    def get_paragraph_by_text(self, text: str) -> Tuple[int, str]:
        """根据文本内容查找段落"""
        for i, para_info in enumerate(self.paragraphs):
            if text in para_info['text']:
                return i, para_info['text']
        return -1, ""

    def extract_text_content(self, file_path: str) -> List[str]:
        """提取文档文本内容"""
        try:
            document = Document(file_path)
            paragraphs = []

            for paragraph in document.paragraphs:
                text = paragraph.text.strip()
                if text:  # 只保留非空段落
                    paragraphs.append(text)

            return paragraphs
        except Exception as e:
            logger.error(f"提取文档内容失败: {e}")
            return []

    def get_statistics(self) -> Dict:
        """获取文档统计信息"""
        return {
            "paragraph_count": len(self.paragraphs),
            "character_count": len(self.text_content),
            "word_count": len(self.text_content.replace(' ', '')),  # 中文字符数
        }
