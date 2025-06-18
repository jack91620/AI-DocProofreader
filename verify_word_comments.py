#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
验证Word文档是否包含正确的审阅批注结构
"""

import zipfile
import xml.etree.ElementTree as ET
import os


def verify_word_comments(docx_file):
    """验证Word文档的审阅批注结构"""
    print(f"🔍 验证文档: {docx_file}")
    
    if not os.path.exists(docx_file):
        print(f"❌ 文件不存在: {docx_file}")
        return False
    
    try:
        with zipfile.ZipFile(docx_file, 'r') as zip_ref:
            file_list = zip_ref.namelist()
            
            # 检查是否包含comments.xml
            has_comments_xml = 'word/comments.xml' in file_list
            print(f"📝 comments.xml存在: {'✅' if has_comments_xml else '❌'}")
            
            # 检查document.xml.rels
            has_document_rels = 'word/_rels/document.xml.rels' in file_list
            print(f"🔗 document.xml.rels存在: {'✅' if has_document_rels else '❌'}")
            
            # 检查[Content_Types].xml
            has_content_types = '[Content_Types].xml' in file_list
            print(f"📋 [Content_Types].xml存在: {'✅' if has_content_types else '❌'}")
            
            if has_comments_xml:
                # 检查comments.xml内容
                with zip_ref.open('word/comments.xml') as f:
                    comments_content = f.read().decode('utf-8')
                    
                # 解析XML
                root = ET.fromstring(comments_content)
                
                # 统计批注数量
                ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                comments = root.findall(f'.//{{{ns_w}}}comment')
                comment_count = len(comments)
                
                print(f"💬 批注数量: {comment_count}")
                
                # 显示前几个批注的内容
                for i, comment in enumerate(comments[:5]):
                    comment_id = comment.get(f'{{{ns_w}}}id', 'N/A')
                    author = comment.get(f'{{{ns_w}}}author', 'N/A')
                    
                    # 提取批注文本
                    text_elem = comment.find(f'.//{{{ns_w}}}t')
                    text = text_elem.text if text_elem is not None else 'N/A'
                    
                    print(f"  💬 批注 {i+1} (ID:{comment_id}, 作者:{author}): {text[:50]}...")
            
            # 检查document.xml中的批注标记
            if 'word/document.xml' in file_list:
                with zip_ref.open('word/document.xml') as f:
                    document_content = f.read().decode('utf-8')
                
                # 统计批注标记
                comment_range_starts = document_content.count('commentRangeStart')
                comment_range_ends = document_content.count('commentRangeEnd')
                comment_references = document_content.count('commentReference')
                
                print(f"🎯 文档中的批注标记:")
                print(f"  - commentRangeStart: {comment_range_starts}")
                print(f"  - commentRangeEnd: {comment_range_ends}")
                print(f"  - commentReference: {comment_references}")
            
            print(f"✅ 验证完成!")
            return True
            
    except Exception as e:
        print(f"❌ 验证失败: {e}")
        return False


def compare_documents(file1, file2):
    """比较两个文档的审阅批注结构"""
    print(f"\n🔄 比较文档:")
    print(f"  文档1: {file1}")
    print(f"  文档2: {file2}")
    
    verify_word_comments(file1)
    print()
    verify_word_comments(file2)


if __name__ == "__main__":
    # 验证最新生成的文档
    latest_file = "sample_output_word_review.docx"
    
    if os.path.exists(latest_file):
        verify_word_comments(latest_file)
    else:
        print(f"❌ 文件不存在: {latest_file}")
        
        # 查找其他测试文档
        test_files = [
            "test_word_full_comments.docx",
            "sample_output_final_comments.docx"
        ]
        
        for test_file in test_files:
            if os.path.exists(test_file):
                print(f"\n🔍 验证备用文档: {test_file}")
                verify_word_comments(test_file)
                break 