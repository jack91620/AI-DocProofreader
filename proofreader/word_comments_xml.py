#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Word文档批注XML处理模块
用于创建Word格式的批注XML文件
"""

import os
import xml.etree.ElementTree as ET
from datetime import datetime
import zipfile
import tempfile


def add_comments_to_docx(input_docx_path: str, output_docx_path: str, comments_data: list) -> bool:
    """向docx文件添加Word原生批注"""
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            # 解压docx文件
            with zipfile.ZipFile(input_docx_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # 创建批注XML文件
            comments_xml_path = os.path.join(temp_dir, 'word', 'comments.xml')
            create_comments_xml(comments_xml_path, comments_data)
            
            # 更新文档关系文件
            create_document_rels(temp_dir)
            
            # 更新内容类型文件
            update_content_types(temp_dir)
            
            # 重新打包为docx
            with zipfile.ZipFile(output_docx_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arc_name = os.path.relpath(file_path, temp_dir)
                        zip_ref.write(file_path, arc_name)
            
            return True
            
    except Exception as e:
        print(f"添加批注失败: {e}")
        return False


def create_comments_xml(comments_xml_path: str, comments_data: list):
    """创建comments.xml文件"""
    # 确保目录存在
    os.makedirs(os.path.dirname(comments_xml_path), exist_ok=True)
    
    # 创建XML命名空间
    ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    ET.register_namespace('w', ns_w)
    
    # 创建根元素
    comments_root = ET.Element(f'{{{ns_w}}}comments')
    
    for i, comment_data in enumerate(comments_data, 1):
        # 创建批注元素
        comment_elem = ET.SubElement(comments_root, f'{{{ns_w}}}comment')
        comment_elem.set(f'{{{ns_w}}}id', str(i))
        comment_elem.set(f'{{{ns_w}}}author', comment_data.get('author', 'AI校对助手'))
        comment_elem.set(f'{{{ns_w}}}date', comment_data.get('date', datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")))
        comment_elem.set(f'{{{ns_w}}}initials', 'AI')
        
        # 创建段落元素
        p_elem = ET.SubElement(comment_elem, f'{{{ns_w}}}p')
        
        # 创建文本运行
        r_elem = ET.SubElement(p_elem, f'{{{ns_w}}}r')
        t_elem = ET.SubElement(r_elem, f'{{{ns_w}}}t')
        t_elem.text = comment_data.get('text', '')
    
    # 写入XML文件
    tree = ET.ElementTree(comments_root)
    tree.write(comments_xml_path, encoding='utf-8', xml_declaration=True)


def create_document_rels(temp_dir: str):
    """创建或更新document.xml.rels文件"""
    rels_dir = os.path.join(temp_dir, 'word', '_rels')
    os.makedirs(rels_dir, exist_ok=True)
    
    rels_path = os.path.join(rels_dir, 'document.xml.rels')
    
    # 检查是否已存在关系文件
    if os.path.exists(rels_path):
        tree = ET.parse(rels_path)
        root = tree.getroot()
    else:
        # 创建新的关系文件
        ns_rels = 'http://schemas.openxmlformats.org/package/2006/relationships'
        ET.register_namespace('', ns_rels)
        root = ET.Element(f'{{{ns_rels}}}Relationships')
        tree = ET.ElementTree(root)
    
    # 检查是否已有批注关系
    ns = {'': 'http://schemas.openxmlformats.org/package/2006/relationships'}
    comment_rels = root.findall(".//Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments']", ns)
    
    if not comment_rels:
        # 添加批注关系
        rel_elem = ET.SubElement(root, 'Relationship')
        rel_elem.set('Id', 'rId999')  # 使用一个不太可能冲突的ID
        rel_elem.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments')
        rel_elem.set('Target', 'comments.xml')
    
    # 写入文件
    tree.write(rels_path, encoding='utf-8', xml_declaration=True)


def update_content_types(temp_dir: str):
    """更新[Content_Types].xml文件以包含批注内容类型"""
    content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
    
    if os.path.exists(content_types_path):
        tree = ET.parse(content_types_path)
        root = tree.getroot()
    else:
        # 创建新的内容类型文件
        ns_ct = 'http://schemas.openxmlformats.org/package/2006/content-types'
        ET.register_namespace('', ns_ct)
        root = ET.Element(f'{{{ns_ct}}}Types')
        tree = ET.ElementTree(root)
    
    # 检查是否已有批注内容类型
    ns = {'': 'http://schemas.openxmlformats.org/package/2006/content-types'}
    comment_overrides = root.findall(".//Override[@PartName='/word/comments.xml']", ns)
    
    if not comment_overrides:
        # 添加批注内容类型
        override_elem = ET.SubElement(root, 'Override')
        override_elem.set('PartName', '/word/comments.xml')
        override_elem.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
    
    # 写入文件
    tree.write(content_types_path, encoding='utf-8', xml_declaration=True) 