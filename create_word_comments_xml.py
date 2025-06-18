#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
手动创建Word文档的comments.xml文件
解决python-docx无法直接生成审阅批注的问题
"""

import zipfile
import tempfile
import os
from datetime import datetime
import xml.etree.ElementTree as ET


def create_comments_xml(comments_data):
    """创建comments.xml内容"""
    # XML命名空间
    ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    
    # 注册命名空间
    ET.register_namespace('w', ns_w)
    
    # 创建根元素
    root = ET.Element(f'{{{ns_w}}}comments')
    
    # 添加每个批注
    for comment in comments_data:
        comment_elem = ET.SubElement(root, f'{{{ns_w}}}comment')
        comment_elem.set(f'{{{ns_w}}}id', str(comment['id']))
        comment_elem.set(f'{{{ns_w}}}author', comment['author'])
        comment_elem.set(f'{{{ns_w}}}date', comment['date'])
        
        # 添加段落
        p_elem = ET.SubElement(comment_elem, f'{{{ns_w}}}p')
        r_elem = ET.SubElement(p_elem, f'{{{ns_w}}}r')
        t_elem = ET.SubElement(r_elem, f'{{{ns_w}}}t')
        t_elem.text = comment['text']
    
    return ET.tostring(root, encoding='unicode', xml_declaration=True)


def add_comments_to_docx(docx_path, output_path, comments_data):
    """将comments.xml添加到Word文档中"""
    try:
        # 创建临时目录
        with tempfile.TemporaryDirectory() as temp_dir:
            # 解压docx文件
            with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # 创建comments.xml内容
            comments_xml = create_comments_xml(comments_data)
            
            # 保存comments.xml到word目录
            word_dir = os.path.join(temp_dir, 'word')
            if not os.path.exists(word_dir):
                os.makedirs(word_dir)
            
            comments_path = os.path.join(word_dir, 'comments.xml')
            with open(comments_path, 'w', encoding='utf-8') as f:
                f.write(comments_xml)
            
            # 更新document.xml.rels文件
            update_document_rels(temp_dir)
            
            # 更新[Content_Types].xml文件
            update_content_types(temp_dir)
            
            # 重新打包为docx文件
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arc_name = os.path.relpath(file_path, temp_dir)
                        zip_ref.write(file_path, arc_name)
            
            print(f"✅ 成功创建包含Word审阅批注的文档: {output_path}")
            return True
            
    except Exception as e:
        print(f"❌ 添加批注失败: {e}")
        return False


def update_document_rels(temp_dir):
    """更新document.xml.rels文件，添加对comments.xml的引用"""
    try:
        rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
        
        if os.path.exists(rels_path):
            # 解析现有的rels文件
            tree = ET.parse(rels_path)
            root = tree.getroot()
            
            # 检查是否已存在comments关系
            ns_r = 'http://schemas.openxmlformats.org/package/2006/relationships'
            comment_rel_exists = False
            
            for rel in root.findall(f'.//{{{ns_r}}}Relationship'):
                if rel.get('Target') == 'comments.xml':
                    comment_rel_exists = True
                    break
            
            if not comment_rel_exists:
                # 生成新的关系ID
                existing_ids = [rel.get('Id') for rel in root.findall(f'.//{{{ns_r}}}Relationship')]
                new_id = f"rId{len(existing_ids) + 1}"
                
                # 添加comments关系
                ET.register_namespace('', ns_r)
                rel_elem = ET.SubElement(root, f'{{{ns_r}}}Relationship')
                rel_elem.set('Id', new_id)
                rel_elem.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments')
                rel_elem.set('Target', 'comments.xml')
                
                # 保存更新后的rels文件
                tree.write(rels_path, encoding='utf-8', xml_declaration=True)
                print("✅ 已更新document.xml.rels文件")
        else:
            # 创建新的rels文件
            create_document_rels(temp_dir)
            
    except Exception as e:
        print(f"❌ 更新document.xml.rels失败: {e}")


def create_document_rels(temp_dir):
    """创建document.xml.rels文件"""
    try:
        rels_dir = os.path.join(temp_dir, 'word', '_rels')
        if not os.path.exists(rels_dir):
            os.makedirs(rels_dir)
        
        rels_path = os.path.join(rels_dir, 'document.xml.rels')
        
        ns_r = 'http://schemas.openxmlformats.org/package/2006/relationships'
        ET.register_namespace('', ns_r)
        
        root = ET.Element(f'{{{ns_r}}}Relationships')
        rel_elem = ET.SubElement(root, f'{{{ns_r}}}Relationship')
        rel_elem.set('Id', 'rId1')
        rel_elem.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments')
        rel_elem.set('Target', 'comments.xml')
        
        tree = ET.ElementTree(root)
        tree.write(rels_path, encoding='utf-8', xml_declaration=True)
        print("✅ 已创建document.xml.rels文件")
        
    except Exception as e:
        print(f"❌ 创建document.xml.rels失败: {e}")


def update_content_types(temp_dir):
    """更新[Content_Types].xml文件，添加comments.xml的内容类型"""
    try:
        content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
        
        if os.path.exists(content_types_path):
            # 解析现有的Content_Types文件
            tree = ET.parse(content_types_path)
            root = tree.getroot()
            
            # 检查是否已存在comments的Override
            ns_ct = 'http://schemas.openxmlformats.org/package/2006/content-types'
            comment_override_exists = False
            
            for override in root.findall(f'.//{{{ns_ct}}}Override'):
                if override.get('PartName') == '/word/comments.xml':
                    comment_override_exists = True
                    break
            
            if not comment_override_exists:
                # 添加comments的Override
                ET.register_namespace('', ns_ct)
                override_elem = ET.SubElement(root, f'{{{ns_ct}}}Override')
                override_elem.set('PartName', '/word/comments.xml')
                override_elem.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
                
                # 保存更新后的Content_Types文件
                tree.write(content_types_path, encoding='utf-8', xml_declaration=True)
                print("✅ 已更新[Content_Types].xml文件")
        
    except Exception as e:
        print(f"❌ 更新[Content_Types].xml失败: {e}")


def test_create_word_comments():
    """测试创建Word审阅批注"""
    try:
        # 测试数据
        comments_data = [
            {
                'id': 1,
                'text': '错别字：应为"计算机科学"',
                'author': 'AI校对助手',
                'date': datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
            },
            {
                'id': 2,
                'text': '术语问题：应为"程序设计"',
                'author': 'AI校对助手',
                'date': datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
            }
        ]
        
        # 将批注添加到测试文档
        input_file = 'test_word_review_comments.docx'
        output_file = 'test_word_full_comments.docx'
        
        if os.path.exists(input_file):
            if add_comments_to_docx(input_file, output_file, comments_data):
                print(f"✅ 完整的Word审阅批注文档已创建: {output_file}")
                print("📝 现在可以在Microsoft Word中查看完整的审阅批注功能")
            else:
                print("❌ 创建失败")
        else:
            print(f"❌ 输入文件不存在: {input_file}")
            
    except Exception as e:
        print(f"❌ 测试失败: {e}")


if __name__ == "__main__":
    test_create_word_comments() 