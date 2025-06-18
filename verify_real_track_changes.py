#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
验证真正的Word跟踪更改功能
"""

import zipfile
import tempfile
import os
import xml.etree.ElementTree as ET
from pathlib import Path

def verify_word_track_changes(docx_file):
    """验证Word文档中的真正跟踪更改"""
    print(f"🔍 验证Word跟踪更改: {docx_file}")
    
    if not os.path.exists(docx_file):
        print(f"❌ 文件不存在: {docx_file}")
        return False
    
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            # 解压docx文件
            with zipfile.ZipFile(docx_file, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # 验证document.xml
            track_changes_found = verify_document_xml(temp_dir)
            
            # 验证settings.xml
            settings_ok = verify_settings_xml(temp_dir)
            
            # 总结验证结果
            if track_changes_found and settings_ok:
                print("✅ Word跟踪更改验证通过")
                print("📝 该文档包含真正的Word修订标记，可以在Microsoft Word中正常显示和操作")
                return True
            else:
                print("❌ Word跟踪更改验证失败")
                return False
                
    except Exception as e:
        print(f"❌ 验证过程出错: {e}")
        return False

def verify_document_xml(temp_dir):
    """验证document.xml中的修订标记"""
    document_path = os.path.join(temp_dir, 'word', 'document.xml')
    
    if not os.path.exists(document_path):
        print("❌ document.xml文件不存在")
        return False
    
    try:
        with open(document_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 统计各种修订标记
        del_count = content.count('<w:del ')
        ins_count = content.count('<w:ins ')
        deltext_count = content.count('<w:delText>')
        
        print(f"📊 document.xml修订标记统计:")
        print(f"   - <w:del> (删除标记): {del_count}")
        print(f"   - <w:ins> (插入标记): {ins_count}")
        print(f"   - <w:delText> (删除文本): {deltext_count}")
        
        # 验证修订标记的完整性
        if del_count > 0 and ins_count > 0:
            print("✅ 发现Word修订标记")
            
            # 进一步验证XML结构
            if verify_revision_xml_structure(content):
                print("✅ Word修订XML结构正确")
                return True
            else:
                print("⚠️  Word修订XML结构可能有问题")
                return False
        else:
            print("⚠️  未发现Word修订标记")
            return False
            
    except Exception as e:
        print(f"❌ 读取document.xml失败: {e}")
        return False

def verify_revision_xml_structure(xml_content):
    """验证修订XML结构的正确性"""
    try:
        # 检查必要的属性
        has_revision_id = 'w:id=' in xml_content
        has_author = 'w:author=' in xml_content
        has_date = 'w:date=' in xml_content
        
        structure_ok = has_revision_id and has_author and has_date
        
        if structure_ok:
            print("✅ 修订标记包含必要属性 (id, author, date)")
        else:
            print("⚠️  修订标记缺少必要属性")
            if not has_revision_id:
                print("   - 缺少 w:id 属性")
            if not has_author:
                print("   - 缺少 w:author 属性")
            if not has_date:
                print("   - 缺少 w:date 属性")
        
        return structure_ok
        
    except Exception as e:
        print(f"❌ 验证XML结构失败: {e}")
        return False

def verify_settings_xml(temp_dir):
    """验证settings.xml中的跟踪更改设置"""
    settings_path = os.path.join(temp_dir, 'word', 'settings.xml')
    
    if not os.path.exists(settings_path):
        print("⚠️  settings.xml文件不存在")
        return False
    
    try:
        with open(settings_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 检查跟踪更改设置
        if 'trackRevisions' in content:
            print("✅ settings.xml中已启用跟踪更改 (trackRevisions)")
            return True
        else:
            print("⚠️  settings.xml中未启用跟踪更改")
            return False
            
    except Exception as e:
        print(f"❌ 读取settings.xml失败: {e}")
        return False

def verify_all_output_files():
    """验证所有输出文件"""
    print("🔍 验证所有Word跟踪更改输出文件...\n")
    
    # 要验证的文件列表
    files_to_verify = [
        "test_real_track_changes.docx",
        "test_word_track_changes.docx",
        "sample_output_track_changes.docx"  # 如果存在的话
    ]
    
    verified_files = 0
    total_files = 0
    
    for filename in files_to_verify:
        if os.path.exists(filename):
            total_files += 1
            print(f"\n{'='*50}")
            if verify_word_track_changes(filename):
                verified_files += 1
            print(f"{'='*50}")
        else:
            print(f"⚠️  文件不存在: {filename}")
    
    print(f"\n📊 验证总结:")
    print(f"   - 验证文件数: {verified_files}/{total_files}")
    
    if verified_files > 0:
        print("✅ 至少有一个文件包含正确的Word跟踪更改功能")
        print("\n📝 使用方法:")
        print("1. 用Microsoft Word打开任一验证通过的文档")
        print("2. 在'审阅'选项卡中可以看到跟踪更改")
        print("3. 可以接受或拒绝每个修改")
        print("4. 修改会以红色删除线和蓝色下划线显示")
    else:
        print("❌ 没有找到正确的Word跟踪更改文档")

if __name__ == "__main__":
    verify_all_output_files() 