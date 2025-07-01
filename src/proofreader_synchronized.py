#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
同步校对器 - 真正同步处理跟踪更改和批注
"""

import os
import sys
from typing import Optional, List, Dict, Any
from rich.console import Console
from docx import Document
from datetime import datetime
import zipfile
import tempfile
import xml.etree.ElementTree as ET
import re

from .config import Config
from .document import DocumentProcessor
from .ai_checker import AIChecker, ProofreadingResult
from .word_comments_xml import create_comments_xml, create_document_rels, update_content_types


class SynchronizedProofReader:
    """同步校对器 - 真正同步处理跟踪更改和批注"""
    
    def __init__(self, api_key: str = None):
        """初始化校对器"""
        self.console = Console()
        if api_key:
            # 如果传入了API密钥，创建临时配置
            import os
            os.environ['OPENAI_API_KEY'] = api_key
        self.config = Config()
        self.ai_checker = AIChecker(self.config)
        self.doc_processor = DocumentProcessor()
        
    def proofread_document(self, input_path: str, output_path: str) -> bool:
        """同步校对文档"""
        try:
            self.console.print(f"[blue]开始同步校对：{input_path}[/blue]")
            
            
            
            # 第一步：AI分析
            self.console.print("[yellow]第一步：AI校对分析文档...[/yellow]")
            text_content = self.doc_processor.extract_text_content(input_path)
            self.console.print(f"提取文本内容: {len(text_content)} 个段落")
            
            # 第二步：获取校对结果
            self.console.print("[yellow]第二步：获取AI校对建议...[/yellow]")
            proofreading_result = self.ai_checker.check_document(text_content)
            
            if not proofreading_result or not proofreading_result.suggestions:
                self.console.print("[red]❌ 未发现需要修改的内容[/red]")
                return False
            
            self.console.print(f"✅ 发现 {len(proofreading_result.suggestions)} 个需要修改的问题")
            
            # 第三步：同步应用跟踪更改和批注
            self.console.print("[yellow]第三步：同步应用跟踪更改和批注...[/yellow]")
            success = self.apply_synchronized_changes(input_path, output_path, proofreading_result)
            
            if success:
                self.console.print(f"[green]✅ 同步校对完成：{output_path}[/green]")
                self.console.print("[dim]📝 文档包含：[/dim]")
                self.console.print("[dim]   - 🔄 Word跟踪更改（可接受/拒绝）[/dim]")
                self.console.print("[dim]   - 💬 同步的详细批注（可查看/回复）[/dim]")
                self.console.print("[dim]   - 🔗 正确的批注引用链接[/dim]")
                return True
            else:
                self.console.print("[red]❌ 同步校对失败[/red]")
                return False
                
        except Exception as e:
            self.console.print(f"[red]❌ 校对过程出错: {e}[/red]")
            import traceback
            traceback.print_exc()
            return False
    
    def apply_synchronized_changes(self, input_path: str, output_path: str, result: ProofreadingResult) -> bool:
        """同步应用跟踪更改和批注"""
        try:
            # 使用临时文件处理
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_input = os.path.join(temp_dir, "input.docx")
                temp_output = os.path.join(temp_dir, "output.docx")
                
                # 复制输入文件
                import shutil
                shutil.copy2(input_path, temp_input)
                
                # 解压文档
                with zipfile.ZipFile(temp_input, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                # 读取文档XML
                document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
                with open(document_xml_path, 'r', encoding='utf-8') as f:
                    doc_content = f.read()
                
                # 同步处理每个修改
                comment_data = []
                for i, suggestion in enumerate(result.suggestions, 1):
                    comment_id = str(i)
                    
                    # 同步添加跟踪更改和批注引用
                    doc_content, comment_info = self.add_synchronized_change(
                        doc_content, suggestion, comment_id
                    )
                    
                    if comment_info:
                        comment_data.append(comment_info)
                        self.console.print(f"✅ 同步处理 {i}: {suggestion['original']} -> {suggestion['suggested']}")
                
                # 保存修改后的文档XML
                with open(document_xml_path, 'w', encoding='utf-8') as f:
                    f.write(doc_content)
                
                # 创建批注XML文件
                if comment_data:
                    self.create_comments_system(temp_dir, comment_data)
                
                # 重新打包文档
                self.repackage_document(temp_dir, temp_output)
                
                # 复制到最终输出位置
                shutil.copy2(temp_output, output_path)
                
                self.console.print(f"✅ 成功应用 {len(comment_data)} 个同步更改和批注")
                return True
                
        except Exception as e:
            self.console.print(f"[red]❌ 应用同步更改失败: {e}[/red]")
            import traceback
            traceback.print_exc()
            return False
    
    def add_synchronized_change(self, doc_content: str, suggestion, comment_id: str) -> tuple:
        """同步添加跟踪更改和批注引用"""
        try:
            original_text = suggestion['original']
            corrected_text = suggestion['suggested']
            
            # 查找原文本在文档中的位置
            # 使用更精确的模式匹配
            pattern = f'<w:t[^>]*>([^<]*{re.escape(original_text)}[^<]*)</w:t>'
            match = re.search(pattern, doc_content)
            
            if not match:
                self.console.print(f"⚠️  未找到文本: {original_text}")
                return doc_content, None
            
            full_text = match.group(1)
            original_tag = match.group(0)
            
            # 创建同步的XML结构
            current_time = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
            
            # 构建理想的XML结构：批注范围包围整个修改区域
            synchronized_xml = f'''<w:commentRangeStart w:id="{comment_id}"/>
<w:del w:id="{comment_id}" w:author="AI校对助手" w:date="{current_time}">
    <w:r><w:delText>{original_text}</w:delText></w:r>
</w:del>
<w:ins w:id="{comment_id}" w:author="AI校对助手" w:date="{current_time}">
    <w:r><w:t>{corrected_text}</w:t></w:r>
</w:ins>
<w:commentRangeEnd w:id="{comment_id}"/>
<w:r><w:commentReference w:id="{comment_id}"/></w:r>'''
            
            # 如果原文本是完整的<w:t>标签内容，直接替换
            if full_text == original_text:
                doc_content = doc_content.replace(original_tag, synchronized_xml, 1)
            else:
                # 如果原文本是<w:t>标签内容的一部分，需要分割处理
                before_text = full_text[:full_text.find(original_text)]
                after_text = full_text[full_text.find(original_text) + len(original_text):]
                
                replacement_xml = f'<w:t>{before_text}</w:t>{synchronized_xml}<w:t>{after_text}</w:t>'
                doc_content = doc_content.replace(original_tag, replacement_xml, 1)
            
            # 准备批注数据
            comment_info = {
                'id': comment_id,
                'author': 'AI校对助手',
                'date': current_time,
                'content': f"💡 改进建议: {original_text} → {corrected_text}\n📋 原因: {suggestion['reason']}\n🎯 类型: 改进建议\n⏰ 建议时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            }
            
            return doc_content, comment_info
            
        except Exception as e:
            self.console.print(f"[red]❌ 添加同步更改失败: {e}[/red]")
            return doc_content, None
    
    def create_comments_system(self, temp_dir: str, comment_data: List[Dict]) -> bool:
        """创建完整的批注系统"""
        try:
            self.console.print(f"🔧 创建完整的批注系统，包含 {len(comment_data)} 个批注")
            
            # 创建批注XML文件
            comments_xml_path = os.path.join(temp_dir, 'word', 'comments.xml')
            comments_xml_content = create_comments_xml(comment_data)
            
            with open(comments_xml_path, 'w', encoding='utf-8') as f:
                f.write(comments_xml_content)
            
            # 更新文档关系
            rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
            if os.path.exists(rels_path):
                with open(rels_path, 'r', encoding='utf-8') as f:
                    rels_content = f.read()
                
                updated_rels = create_document_rels(rels_content)
                
                with open(rels_path, 'w', encoding='utf-8') as f:
                    f.write(updated_rels)
            
            # 更新内容类型
            content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
            if os.path.exists(content_types_path):
                with open(content_types_path, 'r', encoding='utf-8') as f:
                    content_types_content = f.read()
                
                updated_content_types = update_content_types(content_types_content)
                
                with open(content_types_path, 'w', encoding='utf-8') as f:
                    f.write(updated_content_types)
            
            self.console.print("✅ 完整的批注系统创建成功")
            return True
            
        except Exception as e:
            self.console.print(f"[red]❌ 创建批注系统失败: {e}[/red]")
            return False
    
    def repackage_document(self, temp_dir: str, output_path: str) -> bool:
        """重新打包文档"""
        try:
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        if file.endswith('.docx'):
                            continue  # 跳过临时docx文件
                        
                        file_path = os.path.join(root, file)
                        arc_path = os.path.relpath(file_path, temp_dir)
                        zip_file.write(file_path, arc_path)
            
            return True
            
        except Exception as e:
            self.console.print(f"[red]❌ 重新打包文档失败: {e}[/red]")
            return False 