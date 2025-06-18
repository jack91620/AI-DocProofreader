#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
增强版校对器 - 同时使用Word跟踪更改和批注功能
"""

import os
import sys
from typing import Optional
from rich.console import Console
from docx import Document
from datetime import datetime

from .config import Config
from .document import DocumentProcessor
from .ai_checker import AIChecker, ProofreadingResult
from .word_track_changes import WordTrackChangesManager, enable_track_changes_in_docx
from .word_comments_advanced import WordCommentsManager
from .create_word_comments_xml import create_comments_xml, create_document_rels, update_content_types
import zipfile
import tempfile


class ProofReaderWithTrackChangesAndComments:
    """增强版校对器 - 同时使用跟踪更改和批注"""
    
    def __init__(self, api_key: str = None):
        """初始化校对器"""
        self.config = Config()
        if api_key:
            self.config.ai.api_key = api_key
        self.ai_checker = AIChecker(self.config)
        self.console = Console()
        
        self.document_processor = DocumentProcessor()
    
    def proofread_with_track_changes_and_comments(self, input_file: str, output_file: str = None) -> bool:
        """使用跟踪更改和批注进行校对"""
        try:
            # 生成输出文件名
            if not output_file:
                output_file = input_file.replace('.docx', '_tracked_with_comments.docx')
            
            self.console.print(f"[green]开始增强校对：{input_file}[/green]")
            
            # 第一步：创建带跟踪更改的版本
            track_changes_file = input_file.replace('.docx', '_temp_track_changes.docx')
            self.console.print("[blue]第一步：生成Word跟踪更改版本...[/blue]")
            
            if not self._create_track_changes_version(input_file, track_changes_file):
                return False
            
            # 第二步：在跟踪更改版本基础上添加批注
            self.console.print("[blue]第二步：添加详细批注说明...[/blue]")
            
            if not self._add_comments_to_track_changes(track_changes_file, output_file):
                return False
            
            # 清理临时文件
            if os.path.exists(track_changes_file):
                os.remove(track_changes_file)
            
            self.console.print(f"[green]✅ 增强校对完成：{output_file}[/green]")
            self.console.print("[blue]📝 文档包含：[/blue]")
            self.console.print("   - 🔄 真正的Word跟踪更改")
            self.console.print("   - 💬 详细的批注说明")
            self.console.print("   - ✅ 可在Word中完整操作")
            
            return True
            
        except Exception as e:
            self.console.print(f"[red]❌ 增强校对失败: {e}[/red]")
            return False
    
    def _create_track_changes_version(self, input_file: str, output_file: str) -> bool:
        """创建带跟踪更改的版本"""
        try:
            # 读取文档
            doc = Document(input_file)
            
            # 创建跟踪更改管理器
            track_changes_manager = WordTrackChangesManager(doc)
            
            # 提取文本内容
            text_content = self.extract_text_content(doc)
            self.console.print(f"[blue]提取文本内容: {len(text_content)} 个段落[/blue]")
            
            # 进行AI校对
            self.console.print("[bold]开始AI校对...")
            ai_result = self.ai_checker.check_text(' '.join(text_content))
            
            # 转换AI校对结果为跟踪更改格式
            changes = self._convert_ai_result_to_track_changes(ai_result, text_content)
            self.console.print(f"[green]✅ AI校对完成，发现 {len(changes)} 个问题[/green]")
            
            # 应用跟踪更改
            change_count = self._apply_track_changes(doc, changes, track_changes_manager)
            
            # 应用所有跟踪更改
            track_changes_manager.apply_all_changes()
            
            # 保存临时文档
            temp_file = output_file.replace('.docx', '_temp.docx')
            doc.save(temp_file)
            
            # 启用Word跟踪更改并生成最终文档
            if enable_track_changes_in_docx(temp_file, output_file, track_changes_manager.revisions_data):
                # 清理临时文件
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                    
                self.console.print(f"[green]✅ 跟踪更改版本创建完成: {change_count} 个修改[/green]")
                return True
            else:
                return False
            
        except Exception as e:
            self.console.print(f"[red]❌ 创建跟踪更改版本失败: {e}[/red]")
            return False
    
    def _add_comments_to_track_changes(self, track_changes_file: str, output_file: str) -> bool:
        """在跟踪更改版本基础上添加批注"""
        try:
            # 重新读取AI校对结果以生成批注
            doc = Document(track_changes_file)
            
            # 创建批注管理器
            comments_manager = WordCommentsManager(doc)
            
            # 重新进行AI校对以获取批注内容
            text_content = self.extract_text_content(doc)
            ai_result = self.ai_checker.check_text(' '.join(text_content))
            
            # 添加批注
            comment_count = self._add_ai_comments(doc, ai_result, text_content, comments_manager)
            
            # 保存带批注的临时文档
            temp_file = output_file.replace('.docx', '_temp.docx')
            doc.save(temp_file)
            
            # 生成最终的带批注文档
            if self._create_final_document_with_comments(temp_file, output_file, comments_manager.comments):
                # 清理临时文件
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                    
                self.console.print(f"[green]✅ 批注添加完成: {comment_count} 个批注[/green]")
                return True
            else:
                return False
            
        except Exception as e:
            self.console.print(f"[red]❌ 添加批注失败: {e}[/red]")
            return False
    
    def _apply_track_changes(self, doc: Document, changes: list, track_changes_manager: WordTrackChangesManager):
        """应用跟踪更改到文档"""
        change_count = 0
        
        for change in changes:
            paragraph_index = change.get('paragraph_index', 0)
            original_text = change.get('original_text', '')
            corrected_text = change.get('corrected_text', '')
            reason = change.get('reason', '')
            
            # 获取对应段落
            if paragraph_index < len(doc.paragraphs):
                paragraph = doc.paragraphs[paragraph_index]
                
                # 应用跟踪更改
                if track_changes_manager.add_tracked_change(paragraph, original_text, corrected_text, reason):
                    change_count += 1
                    self.console.print(f"[green]✅ 跟踪更改 {change_count}: {original_text} -> {corrected_text}[/green]")
        
        return change_count
    
    def _add_ai_comments(self, doc: Document, ai_result: ProofreadingResult, text_content: list, comments_manager: WordCommentsManager):
        """根据AI校对结果添加批注"""
        comment_count = 0
        
        # 处理issues
        for issue in ai_result.issues:
            problem_text = issue.get('text', '')
            suggestion = issue.get('suggestion', '')
            issue_type = issue.get('type', '')
            severity = issue.get('severity', '')
            
            # 找到问题文本在哪个段落
            for i, paragraph_text in enumerate(text_content):
                if problem_text in paragraph_text and i < len(doc.paragraphs):
                    paragraph = doc.paragraphs[i]
                    
                    # 生成批注文本
                    comment_text = f"🔍 发现问题: {issue_type}\n"
                    comment_text += f"📝 建议: {suggestion}\n"
                    comment_text += f"⚠️ 严重程度: {severity}\n"
                    comment_text += f"⏰ 检查时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                    
                    # 添加批注
                    if comments_manager.add_comment(paragraph, problem_text, comment_text):
                        comment_count += 1
                        break
        
        # 处理suggestions
        for suggestion in ai_result.suggestions:
            original_text = suggestion.get('original', '')
            suggested_text = suggestion.get('suggested', '')
            reason = suggestion.get('reason', '')
            
            # 找到原文本在哪个段落
            for i, paragraph_text in enumerate(text_content):
                if original_text in paragraph_text and i < len(doc.paragraphs):
                    paragraph = doc.paragraphs[i]
                    
                    # 生成批注文本
                    comment_text = f"💡 建议修改: '{original_text}' → '{suggested_text}'\n"
                    comment_text += f"📋 原因: {reason}\n"
                    comment_text += f"⏰ 建议时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                    
                    # 添加批注
                    if comments_manager.add_comment(paragraph, original_text, comment_text):
                        comment_count += 1
                        break
        
        return comment_count
    
    def _create_final_document_with_comments(self, temp_file: str, output_file: str, comments_data: list) -> bool:
        """创建最终的带批注文档"""
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                # 解压docx文件
                with zipfile.ZipFile(temp_file, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                # 添加批注相关文件
                word_dir = os.path.join(temp_dir, 'word')
                os.makedirs(word_dir, exist_ok=True)
                
                # 创建comments.xml
                comments_xml_path = os.path.join(word_dir, 'comments.xml')
                create_comments_xml(comments_xml_path, comments_data)
                
                # 创建document.xml.rels
                rels_dir = os.path.join(word_dir, '_rels')
                os.makedirs(rels_dir, exist_ok=True)
                rels_path = os.path.join(rels_dir, 'document.xml.rels')
                create_document_rels(comments_xml_path.replace('/comments.xml', ''))
                
                # 更新Content_Types.xml
                content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
                update_content_types(content_types_path)
                
                # 重新打包
                with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arc_name = os.path.relpath(file_path, temp_dir)
                            zip_ref.write(file_path, arc_name)
                
                print(f"✅ 成功创建带批注的文档: {output_file}")
                return True
                
        except Exception as e:
            print(f"❌ 创建最终文档失败: {e}")
            return False
    
    def _convert_ai_result_to_track_changes(self, ai_result: ProofreadingResult, text_content: list):
        """将AI校对结果转换为跟踪更改格式"""
        changes = []
        
        # 处理issues
        for issue in ai_result.issues:
            problem_text = issue.get('text', '')
            suggestion = issue.get('suggestion', '')
            
            # 提取修正后的文本
            corrected_text = self._extract_corrected_text(suggestion)
            
            # 如果修正文本与原文本不同，才添加跟踪更改
            if corrected_text and corrected_text != problem_text:
                # 找到问题文本在哪个段落
                for i, paragraph_text in enumerate(text_content):
                    if problem_text in paragraph_text:
                        changes.append({
                            'paragraph_index': i,
                            'original_text': problem_text,
                            'corrected_text': corrected_text,
                            'reason': f"{issue.get('type', '')} - {issue.get('severity', '')}"
                        })
                        break
        
        # 处理suggestions
        for suggestion in ai_result.suggestions:
            original_text = suggestion.get('original', '')
            suggested_text = suggestion.get('suggested', '')
            
            # 如果建议文本与原文本不同，才添加跟踪更改
            if suggested_text and suggested_text != original_text:
                # 找到原文本在哪个段落
                for i, paragraph_text in enumerate(text_content):
                    if original_text in paragraph_text:
                        changes.append({
                            'paragraph_index': i,
                            'original_text': original_text,
                            'corrected_text': suggested_text,
                            'reason': suggestion.get('reason', '')
                        })
                        break
        
        return changes
    
    def _extract_corrected_text(self, suggestion: str):
        """从建议中提取修正后的文本"""
        # 尝试从建议中提取修正文本
        if "建议改为：" in suggestion:
            return suggestion.split("建议改为：")[-1].strip()
        elif "应为" in suggestion:
            return suggestion.split("应为")[-1].strip().strip("'\"")
        elif "->" in suggestion:
            return suggestion.split("->")[-1].strip()
        elif "改为" in suggestion:
            return suggestion.split("改为")[-1].strip().strip("'\"")
        else:
            # 如果无法提取，返回空字符串
            return ""
    
    def extract_text_content(self, doc: Document):
        """提取文档的文本内容"""
        text_content = []
        for paragraph in doc.paragraphs:
            text_content.append(paragraph.text)
        return text_content


# 测试函数
def test_enhanced_proofreader():
    """测试增强版校对器"""
    try:
        # 使用测试API密钥
        api_key = "sk-test"  # 替换为真实的API密钥
        
        proofreader = ProofReaderWithTrackChangesAndComments(api_key)
        
        input_file = "sample_input.docx"
        output_file = "sample_output_enhanced_track_changes_comments.docx"
        
        if os.path.exists(input_file):
            success = proofreader.proofread_with_track_changes_and_comments(input_file, output_file)
            if success:
                print(f"✅ 增强校对成功: {output_file}")
            else:
                print("❌ 增强校对失败")
        else:
            print(f"❌ 输入文件不存在: {input_file}")
            
    except Exception as e:
        print(f"❌ 测试失败: {e}")


if __name__ == "__main__":
    test_enhanced_proofreader() 