#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Word修订XML处理器 - 实现真正的Word跟踪更改功能
"""

import zipfile
import tempfile
import os
from datetime import datetime
import xml.etree.ElementTree as ET


def add_track_changes_to_docx(docx_path, output_path, revisions_data):
    """为Word文档添加跟踪更改功能"""
    try:
        # 创建临时目录
        with tempfile.TemporaryDirectory() as temp_dir:
            # 解压docx文件
            with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # 修改document.xml以添加修订标记
            modify_document_xml(temp_dir, revisions_data)
            
            # 更新settings.xml以启用跟踪更改
            update_settings_xml(temp_dir)
            
            # 重新打包为docx文件
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arc_name = os.path.relpath(file_path, temp_dir)
                        zip_ref.write(file_path, arc_name)
            
            print(f"✅ 成功创建包含跟踪更改的文档: {output_path}")
            return True
            
    except Exception as e:
        print(f"❌ 添加跟踪更改失败: {e}")
        return False


def modify_document_xml(temp_dir, revisions_data):
    """修改document.xml文件以添加修订标记"""
    try:
        document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
        
        if not os.path.exists(document_xml_path):
            print("❌ document.xml文件不存在")
            return False
        
        # 解析document.xml
        tree = ET.parse(document_xml_path)
        root = tree.getroot()
        
        # 定义命名空间
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        
        # 为每个修订处理段落
        for revision in revisions_data:
            paragraph_index = revision.get('paragraph_index', 0)
            original_text = revision.get('original_text', '')
            corrected_text = revision.get('corrected_text', '')
            
            # 查找对应的段落
            paragraphs = root.findall('.//w:p', ns)
            if paragraph_index < len(paragraphs):
                paragraph = paragraphs[paragraph_index]
                
                # 处理段落文本，添加修订标记
                process_paragraph_revisions(paragraph, original_text, corrected_text, ns)
        
        # 保存修改后的document.xml
        tree.write(document_xml_path, encoding='utf-8', xml_declaration=True)
        print("✅ 已更新document.xml文件")
        return True
        
    except Exception as e:
        print(f"❌ 修改document.xml失败: {e}")
        return False


def process_paragraph_revisions(paragraph, original_text, corrected_text, ns):
    """处理段落中的修订标记"""
    try:
        # 获取段落文本
        paragraph_text = get_paragraph_text(paragraph, ns)
        
        if original_text not in paragraph_text:
            return False
        
        # 生成修订ID
        revision_id = generate_revision_id()
        author = "AI校对助手"
        date = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
        
        # 查找并替换文本runs
        runs = paragraph.findall('.//w:r', ns)
        
        for run in runs:
            text_elem = run.find('.//w:t', ns)
            if text_elem is not None and original_text in text_elem.text:
                # 创建修订标记
                create_revision_markup(paragraph, run, original_text, corrected_text, 
                                     revision_id, author, date, ns)
                break
        
        return True
        
    except Exception as e:
        print(f"处理段落修订失败: {e}")
        return False


def get_paragraph_text(paragraph, ns):
    """获取段落的完整文本"""
    text_parts = []
    for text_elem in paragraph.findall('.//w:t', ns):
        if text_elem.text:
            text_parts.append(text_elem.text)
    return ''.join(text_parts)


def create_revision_markup(paragraph, run, original_text, corrected_text, 
                         revision_id, author, date, ns):
    """创建修订标记"""
    try:
        # 获取run的父元素
        parent = run.getparent()
        run_index = list(parent).index(run)
        
        # 创建删除标记
        del_element = ET.Element(f"{{{ns['w']}}}del")
        del_element.set(f"{{{ns['w']}}}id", str(revision_id))
        del_element.set(f"{{{ns['w']}}}author", author)
        del_element.set(f"{{{ns['w']}}}date", date)
        
        # 创建删除的run
        del_run = ET.SubElement(del_element, f"{{{ns['w']}}}r")
        del_run_props = ET.SubElement(del_run, f"{{{ns['w']}}}rPr")
        del_text = ET.SubElement(del_run, f"{{{ns['w']}}}delText")
        del_text.text = original_text
        
        # 创建插入标记
        ins_element = ET.Element(f"{{{ns['w']}}}ins")
        ins_element.set(f"{{{ns['w']}}}id", str(revision_id + 1))
        ins_element.set(f"{{{ns['w']}}}author", author)
        ins_element.set(f"{{{ns['w']}}}date", date)
        
        # 创建插入的run
        ins_run = ET.SubElement(ins_element, f"{{{ns['w']}}}r")
        ins_run_props = ET.SubElement(ins_run, f"{{{ns['w']}}}rPr")
        ins_text = ET.SubElement(ins_run, f"{{{ns['w']}}}t")
        ins_text.text = corrected_text
        
        # 替换原来的run
        parent.remove(run)
        parent.insert(run_index, del_element)
        parent.insert(run_index + 1, ins_element)
        
        print(f"✅ 已创建修订标记: {original_text} -> {corrected_text}")
        
    except Exception as e:
        print(f"创建修订标记失败: {e}")


_revision_counter = 0

def generate_revision_id():
    """生成修订ID"""
    global _revision_counter
    _revision_counter += 1
    return _revision_counter


def update_settings_xml(temp_dir):
    """更新settings.xml文件以启用跟踪更改"""
    try:
        settings_xml_path = os.path.join(temp_dir, 'word', 'settings.xml')
        
        if os.path.exists(settings_xml_path):
            # 解析现有的settings.xml
            tree = ET.parse(settings_xml_path)
            root = tree.getroot()
        else:
            # 创建新的settings.xml
            ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            root = ET.Element(f'{{{ns_w}}}settings')
        
        # 检查是否已存在trackRevisions设置
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        track_revisions = root.find('.//w:trackRevisions', ns)
        
        if track_revisions is None:
            # 添加trackRevisions设置
            track_revisions = ET.SubElement(root, f"{{{ns['w']}}}trackRevisions")
            track_revisions.set(f"{{{ns['w']}}}val", "true")
            
            # 保存settings.xml
            tree = ET.ElementTree(root)
            tree.write(settings_xml_path, encoding='utf-8', xml_declaration=True)
            print("✅ 已更新settings.xml文件")
        
    except Exception as e:
        print(f"❌ 更新settings.xml失败: {e}")


def test_track_changes():
    """测试跟踪更改功能"""
    try:
        # 测试修订数据
        revisions_data = [
            {
                'paragraph_index': 1,
                'original_text': '计算器科学',
                'corrected_text': '计算机科学'
            },
            {
                'paragraph_index': 2,
                'original_text': '程式设计',
                'corrected_text': '程序设计'
            }
        ]
        
        # 使用测试文档
        input_file = 'test_word_revisions.docx'
        output_file = 'test_word_track_changes.docx'
        
        if os.path.exists(input_file):
            if add_track_changes_to_docx(input_file, output_file, revisions_data):
                print(f"✅ 跟踪更改文档已创建: {output_file}")
                print("📝 现在可以在Microsoft Word中查看跟踪更改功能")
            else:
                print("❌ 创建失败")
        else:
            print(f"❌ 输入文件不存在: {input_file}")
            
    except Exception as e:
        print(f"❌ 测试失败: {e}")


if __name__ == "__main__":
    test_track_changes() 