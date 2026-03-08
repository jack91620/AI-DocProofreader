#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
视觉修订模块 - 使用删除线和下划线显示修改
"""

from docx.shared import RGBColor
from logging_config import get_logger

logger = get_logger("word_ops.visual_revisions")


class SimpleWordRevisionsManager:
    """简化的Word修订管理器 - 使用直观的修订显示"""

    def __init__(self, document):
        self.document = document
        self.revision_counter = 0
        self.author = "AI校对助手"

    def add_revision(self, paragraph, original_text: str, corrected_text: str, reason: str = ""):
        """添加简化的修订标记"""
        try:
            paragraph_text = paragraph.text
            start_pos = paragraph_text.find(original_text)

            if start_pos == -1:
                logger.warning(f"未找到需要修订的文本: {original_text}")
                return False

            end_pos = start_pos + len(original_text)

            # 清空段落并重建
            paragraph.clear()

            # 添加原文本之前的内容
            if start_pos > 0:
                paragraph.add_run(paragraph_text[:start_pos])

            # 添加删除的文本（删除线样式）
            deleted_run = paragraph.add_run(original_text)
            deleted_run.font.strike = True
            deleted_run.font.color.rgb = None  # 红色

            # 添加插入的文本（下划线样式）
            inserted_run = paragraph.add_run(corrected_text)
            inserted_run.font.underline = True
            inserted_run.font.color.rgb = RGBColor(0, 0, 255)

            # 添加原文本之后的内容
            if end_pos < len(paragraph_text):
                paragraph.add_run(paragraph_text[end_pos:])

            self.revision_counter += 1
            logger.info(f"修订标记已添加: {original_text} -> {corrected_text}")
            return True

        except Exception as e:
            logger.error(f"添加修订标记失败: {e}")
            return False
