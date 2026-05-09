#!/usr/bin/env python
"""
Word批注操作模块 - 合并自word_comments_advanced.py和word_comments_xml.py
负责在document.xml中插入commentRangeStart/End/Reference标记，
以及创建comments.xml文件和更新DOCX元数据。
"""

import os
import xml.etree.ElementTree as ET
from datetime import datetime

from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from logging_config import get_logger

from .docx_packager import modify_docx_package

logger = get_logger("word_ops.comments")


class WordCommentsManager:
    """Word审阅批注管理器"""

    def __init__(self, document):
        self.document = document
        self.comment_counter = 0
        self.comments = []  # 存储批注信息

    def add_comment(
        self, paragraph, target_text: str, comment_text: str, author: str = "AI校对助手"
    ):
        """在段落中添加Word审阅批注"""
        try:
            original_text = paragraph.text
            start_pos = original_text.find(target_text)

            if start_pos == -1:
                logger.warning(f"未找到目标文本: {target_text}")
                return False

            end_pos = start_pos + len(target_text)

            # 生成批注ID
            self.comment_counter += 1
            comment_id = self.comment_counter

            # 存储批注信息
            self.comments.append(
                {
                    "id": comment_id,
                    "text": comment_text,
                    "author": author,
                    "date": datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ"),
                }
            )

            # 重建段落，正确插入批注标记
            self._rebuild_paragraph_with_comment(
                paragraph, original_text, start_pos, end_pos, comment_id
            )

            logger.info(f"Word审阅批注已添加: {comment_text[:50]}...")
            return True

        except Exception as e:
            logger.error(f"添加Word审阅批注失败: {e}")
            return False

    def _rebuild_paragraph_with_comment(
        self, paragraph, original_text, start_pos, end_pos, comment_id
    ):
        """重建段落，正确插入批注标记"""
        try:
            # 1. 添加批注范围开始标记
            self._add_comment_range_start_to_element(paragraph._element, comment_id)

            # 2. 查找包含目标文本的run并高亮
            for run in paragraph.runs:
                if start_pos <= len(run.text) and run.text:
                    run_text = run.text
                    if original_text[start_pos:end_pos] in run_text:
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        break

            # 3. 添加批注范围结束标记
            self._add_comment_range_end_to_element(paragraph._element, comment_id)

            # 4. 添加批注引用标记
            self._add_comment_reference_to_element(paragraph._element, comment_id)

        except Exception as e:
            logger.error(f"重建段落失败: {e}")
            # 如果失败，使用简单的标记方法
            self._add_simple_comment_markers(paragraph, comment_id)

    def _add_comment_range_start_to_element(self, paragraph_element, comment_id):
        """直接在段落元素中添加批注范围开始标记"""
        try:
            element = OxmlElement("w:commentRangeStart")
            element.set(qn("w:id"), str(comment_id))
            paragraph_element.insert(0, element)
        except Exception as e:
            logger.error(f"添加批注范围开始标记到元素失败: {e}")

    def _add_comment_range_end_to_element(self, paragraph_element, comment_id):
        """直接在段落元素中添加批注范围结束标记"""
        try:
            element = OxmlElement("w:commentRangeEnd")
            element.set(qn("w:id"), str(comment_id))
            paragraph_element.append(element)
        except Exception as e:
            logger.error(f"添加批注范围结束标记到元素失败: {e}")

    def _add_comment_reference_to_element(self, paragraph_element, comment_id):
        """直接在段落元素中添加批注引用标记"""
        try:
            run_element = OxmlElement("w:r")
            comment_ref = OxmlElement("w:commentReference")
            comment_ref.set(qn("w:id"), str(comment_id))
            run_element.append(comment_ref)
            paragraph_element.append(run_element)
        except Exception as e:
            logger.error(f"添加批注引用标记到元素失败: {e}")

    def _add_simple_comment_markers(self, paragraph, comment_id):
        """简单的批注标记方法（备用）"""
        try:
            element_start = OxmlElement("w:commentRangeStart")
            element_start.set(qn("w:id"), str(comment_id))
            paragraph._element.append(element_start)

            element_end = OxmlElement("w:commentRangeEnd")
            element_end.set(qn("w:id"), str(comment_id))
            paragraph._element.append(element_end)

            new_run = paragraph.add_run()
            run_element = new_run._element
            comment_ref = OxmlElement("w:commentReference")
            comment_ref.set(qn("w:id"), str(comment_id))
            run_element.append(comment_ref)
        except Exception as e:
            logger.error(f"添加简单批注标记失败: {e}")

    def finalize_document(self):
        """完成文档处理，准备批注数据"""
        try:
            if not self.comments:
                logger.info("没有批注需要处理")
                return True

            logger.info(f"准备 {len(self.comments)} 个批注数据用于XML生成")
            return True

        except Exception as e:
            logger.error(f"完成文档处理失败: {e}")
            return False

    def get_comments_for_xml(self):
        """获取用于生成XML的批注数据"""
        return [
            {
                "id": comment["id"],
                "text": comment["text"],
                "author": comment["author"],
                "date": comment["date"],
            }
            for comment in self.comments
        ]


def add_comments_to_docx(input_docx_path: str, output_docx_path: str, comments_data: list) -> bool:
    """向docx文件添加Word原生批注"""

    def modifier(temp_dir):
        # 创建批注XML文件
        comments_xml_path = os.path.join(temp_dir, "word", "comments.xml")
        create_comments_xml(comments_xml_path, comments_data)

        # 更新文档关系文件
        create_document_rels(temp_dir)

        # 更新内容类型文件
        update_content_types(temp_dir)
        return True

    return modify_docx_package(input_docx_path, output_docx_path, modifier)


def create_comments_xml(comments_xml_path: str, comments_data: list):
    """创建comments.xml文件"""
    # 确保目录存在
    os.makedirs(os.path.dirname(comments_xml_path), exist_ok=True)

    # 创建XML命名空间
    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ET.register_namespace("w", ns_w)

    # 创建根元素
    comments_root = ET.Element(f"{{{ns_w}}}comments")

    for i, comment_data in enumerate(comments_data, 1):
        # 创建批注元素
        comment_elem = ET.SubElement(comments_root, f"{{{ns_w}}}comment")

        # 使用传入的ID，如果没有则使用索引
        comment_id = comment_data.get("id", i)
        comment_elem.set(f"{{{ns_w}}}id", str(comment_id))
        comment_elem.set(f"{{{ns_w}}}author", comment_data.get("author", "AI校对助手"))
        comment_elem.set(
            f"{{{ns_w}}}date",
            comment_data.get("date", datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")),
        )
        comment_elem.set(f"{{{ns_w}}}initials", "AI")

        # 创建段落元素
        p_elem = ET.SubElement(comment_elem, f"{{{ns_w}}}p")

        # 创建文本运行
        r_elem = ET.SubElement(p_elem, f"{{{ns_w}}}r")
        t_elem = ET.SubElement(r_elem, f"{{{ns_w}}}t")
        t_elem.text = comment_data.get("text", "")

        logger.info(f"创建批注XML: ID={comment_id}, 内容={comment_data.get('text', '')[:30]}...")

    # 写入XML文件
    tree = ET.ElementTree(comments_root)
    tree.write(comments_xml_path, encoding="utf-8", xml_declaration=True)
    logger.info(f"批注XML文件已创建: {comments_xml_path}")


_NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def create_document_rels(temp_dir: str):
    """创建或更新document.xml.rels文件，注册批注关系。

    注意：新增 element 必须带命名空间前缀 ``{ns}Tag`` (Clark 表示法)，
    否则被序列化为无命名空间元素，Word 严格 schema 校验会拒收，
    触发"无法读取的内容/恢复此文档"提示。
    """
    rels_dir = os.path.join(temp_dir, "word", "_rels")
    os.makedirs(rels_dir, exist_ok=True)

    rels_path = os.path.join(rels_dir, "document.xml.rels")

    # 检查是否已存在关系文件
    if os.path.exists(rels_path):
        tree = ET.parse(rels_path)
        root = tree.getroot()
    else:
        # 创建新的关系文件
        ET.register_namespace("", _NS_RELS)
        root = ET.Element(f"{{{_NS_RELS}}}Relationships")
        tree = ET.ElementTree(root)

    # 检查是否已有批注关系（findall 要带命名空间）
    ns = {"r": _NS_RELS}
    comment_rels = root.findall(
        ".//r:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments']",
        ns,
    )

    if not comment_rels:
        # 必须用 Clark 表示法 {ns}Tag，否则序列化出来不带命名空间，Word 不认
        rel_elem = ET.SubElement(root, f"{{{_NS_RELS}}}Relationship")
        rel_elem.set("Id", "rId999")
        rel_elem.set(
            "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        )
        rel_elem.set("Target", "comments.xml")

    # 写入文件
    tree.write(rels_path, encoding="utf-8", xml_declaration=True)


def update_content_types(temp_dir: str):
    """更新[Content_Types].xml文件以包含批注内容类型。

    与 create_document_rels 同样的命名空间陷阱：必须用 Clark 表示法。
    """
    content_types_path = os.path.join(temp_dir, "[Content_Types].xml")

    if os.path.exists(content_types_path):
        tree = ET.parse(content_types_path)
        root = tree.getroot()
    else:
        # 创建新的内容类型文件
        ET.register_namespace("", _NS_CT)
        root = ET.Element(f"{{{_NS_CT}}}Types")
        tree = ET.ElementTree(root)

    # 检查是否已有批注内容类型
    ns = {"c": _NS_CT}
    comment_overrides = root.findall(
        ".//c:Override[@PartName='/word/comments.xml']", ns
    )

    if not comment_overrides:
        override_elem = ET.SubElement(root, f"{{{_NS_CT}}}Override")
        override_elem.set("PartName", "/word/comments.xml")
        override_elem.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        )

    # 写入文件
    tree.write(content_types_path, encoding="utf-8", xml_declaration=True)
