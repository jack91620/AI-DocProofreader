#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Word跟踪更改操作模块 - 迁移自word_track_changes.py
实现Microsoft Word审阅中的修订功能
"""

import os
import xml.etree.ElementTree as ET
from datetime import datetime
import uuid

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .docx_packager import modify_docx_package
from logging_config import get_logger

logger = get_logger("word_ops.track_changes")


class WordTrackChangesManager:
    """Word跟踪更改管理器 - 生成真正的Word修订标记"""

    def __init__(self, document):
        self.document = document
        self.revision_counter = 0
        self.author = "AI校对助手"
        self.date = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
        self.revisions_data = []
        self.paragraph_changes = {}  # 跟踪每个段落的修改
        # 使用UUID前缀避免标记碰撞
        self._marker_prefix = uuid.uuid4().hex[:8]

    def add_tracked_change(self, paragraph, original_text: str, corrected_text: str, reason: str = ""):
        """添加跟踪更改（真正的Word修订）"""
        try:
            # 生成修订ID
            self.revision_counter += 1
            del_revision_id = self.revision_counter
            ins_revision_id = self.revision_counter + 1000

            # 存储修订数据
            self.revisions_data.append({
                'paragraph': paragraph,
                'original_text': original_text,
                'corrected_text': corrected_text,
                'del_id': del_revision_id,
                'ins_id': ins_revision_id,
                'author': self.author,
                'date': self.date,
                'reason': reason
            })

            # 跟踪段落修改，批量处理
            para_id = id(paragraph)
            if para_id not in self.paragraph_changes:
                self.paragraph_changes[para_id] = {
                    'paragraph': paragraph,
                    'original_text': paragraph.text,
                    'changes': []
                }

            self.paragraph_changes[para_id]['changes'].append({
                'original_text': original_text,
                'corrected_text': corrected_text,
                'del_id': del_revision_id,
                'ins_id': ins_revision_id,
                'reason': reason
            })

            logger.info(f"已添加跟踪更改: {original_text} -> {corrected_text}")
            return True

        except Exception as e:
            logger.error(f"添加跟踪更改失败: {e}")
            return False

    def apply_all_changes(self):
        """应用所有跟踪更改到文档"""
        for para_id, para_data in self.paragraph_changes.items():
            paragraph = para_data['paragraph']
            original_full_text = para_data['original_text']
            changes = para_data['changes']

            # 应用所有更改到段落
            self._apply_changes_to_paragraph(paragraph, original_full_text, changes)

    def _apply_changes_to_paragraph(self, paragraph, original_text, changes):
        """将多个更改应用到单个段落"""
        try:
            # 收集所有需要更改的位置
            changes_with_pos = []
            processed_positions = set()

            for change in changes:
                original_part = change['original_text']
                start = 0
                while True:
                    pos = original_text.find(original_part, start)
                    if pos == -1:
                        break

                    position_key = (pos, original_part)
                    if position_key not in processed_positions:
                        changes_with_pos.append((pos, change))
                        processed_positions.add(position_key)

                    start = pos + 1

            # 按位置降序排序（从后往前处理）
            changes_with_pos.sort(key=lambda x: x[0], reverse=True)

            # 清空段落
            paragraph.clear()

            if not changes_with_pos:
                paragraph.add_run(original_text)
                return

            # 构建标记文本 - 使用UUID前缀避免碰撞
            current_text = original_text
            pfx = self._marker_prefix

            for pos, change in changes_with_pos:
                original_part = change['original_text']
                corrected_part = change['corrected_text']
                del_id = change['del_id']
                ins_id = change['ins_id']

                if pos + len(original_part) <= len(current_text) and current_text[pos:pos+len(original_part)] == original_part:
                    before_part = current_text[:pos]
                    after_part = current_text[pos + len(original_part):]

                    marker = f"\x00{pfx}DEL:{del_id}:{original_part}\x00\x00{pfx}INS:{ins_id}:{corrected_part}\x00"
                    current_text = before_part + marker + after_part

            # 解析并构建段落
            self._build_paragraph_from_marked_text(paragraph, current_text)

        except Exception as e:
            logger.error(f"应用段落更改失败: {e}")
            paragraph.clear()
            paragraph.add_run(original_text)

    def _build_paragraph_from_marked_text(self, paragraph, marked_text):
        """从标记文本构建段落"""
        try:
            pfx = self._marker_prefix
            del_marker = f"\x00{pfx}DEL:"
            ins_marker = f"\x00{pfx}INS:"

            pos = 0
            while pos < len(marked_text):
                del_start = marked_text.find(del_marker, pos)
                if del_start == -1:
                    remaining = marked_text[pos:]
                    if remaining:
                        paragraph.add_run(remaining)
                    break

                # 添加删除标记之前的文本
                if del_start > pos:
                    before_text = marked_text[pos:del_start]
                    paragraph.add_run(before_text)

                # 解析删除标记
                del_content_start = del_start + len(del_marker)
                del_end = marked_text.find("\x00", del_content_start)
                if del_end == -1:
                    break

                del_content = marked_text[del_content_start:del_end]
                parts = del_content.split(":", 1)
                if len(parts) == 2:
                    del_id = parts[0]
                    del_text = parts[1]

                    del_element = self._create_deletion_element(del_text, del_id)
                    paragraph._element.append(del_element)

                # 查找对应的插入标记
                ins_start = marked_text.find(ins_marker, del_end)
                if ins_start == del_end + 1:  # 紧接着的插入标记
                    ins_content_start = ins_start + len(ins_marker)
                    ins_end = marked_text.find("\x00", ins_content_start)
                    if ins_end != -1:
                        ins_content = marked_text[ins_content_start:ins_end]
                        parts = ins_content.split(":", 1)
                        if len(parts) == 2:
                            ins_id = parts[0]
                            ins_text = parts[1]

                            ins_element = self._create_insertion_element(ins_text, ins_id)
                            paragraph._element.append(ins_element)

                        pos = ins_end + 1
                    else:
                        pos = del_end + 1
                else:
                    pos = del_end + 1

        except Exception as e:
            logger.error(f"构建段落失败: {e}")
            paragraph.clear()
            # Clean up markers from text
            clean_text = marked_text
            for char in ["\x00"]:
                clean_text = clean_text.replace(char, "")
            paragraph.add_run(clean_text)

    def _create_deletion_element(self, text, revision_id):
        """创建删除元素 (w:del)"""
        del_element = OxmlElement('w:del')
        del_element.set(qn('w:id'), str(revision_id))
        del_element.set(qn('w:author'), self.author)
        del_element.set(qn('w:date'), self.date)

        del_run = OxmlElement('w:r')

        del_text = OxmlElement('w:delText')
        del_text.text = text
        del_run.append(del_text)

        del_element.append(del_run)
        return del_element

    def _create_insertion_element(self, text, revision_id):
        """创建插入元素 (w:ins)"""
        ins_element = OxmlElement('w:ins')
        ins_element.set(qn('w:id'), str(revision_id))
        ins_element.set(qn('w:author'), self.author)
        ins_element.set(qn('w:date'), self.date)

        ins_run = OxmlElement('w:r')

        ins_text = OxmlElement('w:t')
        ins_text.text = text
        ins_run.append(ins_text)

        ins_element.append(ins_run)
        return ins_element


def enable_track_changes_in_docx(docx_path, output_path, revisions_data):
    """在Word文档中启用跟踪更改并添加修订"""
    def modifier(temp_dir):
        enable_track_changes_setting(temp_dir)
        verify_document_revisions(temp_dir)
        return True

    return modify_docx_package(docx_path, output_path, modifier)


def enable_track_changes_setting(temp_dir):
    """在settings.xml中启用跟踪更改"""
    try:
        settings_path = os.path.join(temp_dir, 'word', 'settings.xml')

        if os.path.exists(settings_path):
            tree = ET.parse(settings_path)
            root = tree.getroot()
        else:
            ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            root = ET.Element(f'{{{ns_w}}}settings')
            tree = ET.ElementTree(root)

        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        track_revisions = root.find('.//w:trackRevisions', ns)

        if track_revisions is None:
            track_revisions = ET.SubElement(root, f"{{{ns['w']}}}trackRevisions")

        track_revisions.set(f"{{{ns['w']}}}val", "1")

        tree.write(settings_path, encoding='utf-8', xml_declaration=True)
        logger.info("已在settings.xml中启用跟踪更改")

    except Exception as e:
        logger.error(f"启用跟踪更改设置失败: {e}")


def verify_document_revisions(temp_dir):
    """验证document.xml中的修订标记"""
    try:
        document_path = os.path.join(temp_dir, 'word', 'document.xml')

        if os.path.exists(document_path):
            with open(document_path, 'r', encoding='utf-8') as f:
                content = f.read()

            del_count = content.count('<w:del ')
            ins_count = content.count('<w:ins ')
            deltext_count = content.count('<w:delText>')

            logger.info(f"document.xml修订标记统计:")
            logger.info(f"   - w:del (删除标记): {del_count}")
            logger.info(f"   - w:ins (插入标记): {ins_count}")
            logger.info(f"   - w:delText (删除文本): {deltext_count}")

            if del_count > 0 or ins_count > 0:
                logger.info("发现Word修订标记")
            else:
                logger.warning("未发现Word修订标记")

    except Exception as e:
        logger.error(f"验证修订标记失败: {e}")
