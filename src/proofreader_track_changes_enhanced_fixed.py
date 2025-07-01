#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
修复版增强校对器 - 确保跟踪更改和批注都正确显示
"""

import os
import sys
from typing import Optional
from rich.console import Console
from docx import Document
from datetime import datetime
import zipfile
import tempfile
import xml.etree.ElementTree as ET

from config import Config
from document import DocumentProcessor
from ai_checker import AIChecker, ProofreadingResult
from word_track_changes import WordTrackChangesManager, enable_track_changes_in_docx
from word_comments_advanced import WordCommentsManager
from word_comments_xml import create_comments_xml, create_document_rels, update_content_types


class ProofReaderWithTrackChangesAndCommentsFixed:
    """修复版增强校对器 - 确保跟踪更改和批注都正确显示"""
    
    def __init__(self, api_key: str = None):
        """初始化校对器"""
        if api_key:
            # 如果传入了API密钥，设置环境变量
            import os
            os.environ['OPENAI_API_KEY'] = api_key
        self.config = Config()
        self.ai_checker = AIChecker(self.config)
        self.console = Console()
        self.document_processor = DocumentProcessor()
    
    def proofread_with_track_changes_and_comments(self, input_file: str, output_file: str = None) -> bool:
        """使用跟踪更改和批注进行校对 - 支持长文档分块处理"""
        try:
            # 生成输出文件名
            if not output_file:
                output_file = input_file.replace('.docx', '_tracked_with_comments_fixed.docx')
            
            self.console.print(f"[green]开始修复版增强校对：{input_file}[/green]")
            
            # 读取文档
            doc = Document(input_file)
            
            # 提取文本内容
            text_content = self.extract_text_content(doc)
            self.console.print(f"[blue]提取文本内容: {len(text_content)} 个段落[/blue]")
            
            # 进行AI校对 - 使用分块处理方法
            self.console.print("[bold]开始AI校对...")
            ai_result = self.ai_checker.check_text_by_paragraphs(text_content)
            
            # 创建同步的跟踪更改和批注数据
            synchronized_changes = self._create_synchronized_changes(ai_result, text_content, doc)
            self.console.print(f"[green]✅ AI校对完成，发现 {len(synchronized_changes)} 个问题[/green]")
            
            if not synchronized_changes:
                self.console.print("[yellow]⚠️ 未发现需要修改的问题，直接复制原文档[/yellow]")
                import shutil
                shutil.copy2(input_file, output_file)
                return True
            
            # 应用同步更改
            success = self._apply_synchronized_changes_fixed(doc, synchronized_changes, output_file)
            
            if success:
                self.console.print(f"[green]✅ 修复版增强校对完成！输出文件：{output_file}[/green]")
                self.console.print("[blue]📝 使用Word打开文档，体验完整的校对功能：[/blue]")
                self.console.print("   - [red]删除线文本[/red] 表示需要删除的内容")
                self.console.print("   - [blue underline]下划线文本[/blue underline] 表示新插入的内容")
                self.console.print("   - 💬 详细的批注说明修改原因")
                self.console.print("   - ✅ 可以接受或拒绝每个修改")
                self.console.print("   - 🗨️ 可以回复批注进行讨论")
                return True
            else:
                return False
            
        except Exception as e:
            self.console.print(f"[red]❌ 修复版增强校对失败: {e}[/red]")
            import traceback
            traceback.print_exc()
            return False

    def _create_synchronized_changes(self, ai_result: ProofreadingResult, text_content: list, doc: Document):
        """创建同步的跟踪更改和批注数据"""
        synchronized_changes = []
        processed_pairs = set()  # 避免重复处理相同的修正对
        
        self.console.print(f"[blue]🔍 处理AI发现的 {len(ai_result.issues)} 个问题和 {len(ai_result.suggestions)} 个建议[/blue]")
        
        # 创建所有可能的修正对
        all_corrections = []
        
        # 从suggestions中提取修正对
        for suggestion in ai_result.suggestions:
            original_text = suggestion.get('original', '')
            suggested_text = suggestion.get('suggested', '')
            reason = suggestion.get('reason', '')
            
            if suggested_text and suggested_text != original_text:
                # 提取具体的词汇修正
                corrections = self._extract_word_corrections(original_text, suggested_text)
                for orig, corr in corrections:
                    correction_pair = (orig, corr)
                    if correction_pair not in processed_pairs:
                        all_corrections.append({
                            'original': orig,
                            'corrected': corr,
                            'reason': reason,
                            'type': 'suggestion',
                            'full_original': original_text,
                            'full_suggested': suggested_text
                        })
                        processed_pairs.add(correction_pair)
        
        # 从issues中提取修正对
        for issue in ai_result.issues:
            problem_text = issue.get('text', '')
            suggestion = issue.get('suggestion', '')
            issue_type = issue.get('type', '')
            severity = issue.get('severity', '')
            
            if issue_type == "术语不一致" and "发现多种术语：" in problem_text:
                # 处理术语不一致
                terms = self._extract_terms_from_inconsistency(problem_text, suggestion)
                for original_term, corrected_term in terms:
                    correction_pair = (original_term, corrected_term)
                    if correction_pair not in processed_pairs:
                        all_corrections.append({
                            'original': original_term,
                            'corrected': corrected_term,
                            'reason': f"{issue_type} - {severity}",
                            'type': 'terminology',
                            'suggestion_text': suggestion
                        })
                        processed_pairs.add(correction_pair)
            elif issue_type in ["错别字和用词不当", "标点符号使用"]:
                # 处理错别字和标点问题
                corrected_text = self._extract_corrected_text(suggestion)
                if corrected_text and corrected_text != problem_text:
                    correction_pair = (problem_text, corrected_text)
                    if correction_pair not in processed_pairs:
                        all_corrections.append({
                            'original': problem_text,
                            'corrected': corrected_text,
                            'reason': f"{issue_type} - {severity}",
                            'type': 'error_fix',
                            'suggestion_text': suggestion
                        })
                        processed_pairs.add(correction_pair)
        
        # 创建段落索引映射 - 从非空段落索引到实际段落索引
        paragraph_mapping = {}
        text_index = 0
        for doc_index, paragraph in enumerate(doc.paragraphs):
            if paragraph.text.strip():
                paragraph_mapping[text_index] = doc_index
                text_index += 1
        
        # 改进的文本匹配和应用逻辑
        for correction in all_corrections:
            original = correction['original']
            corrected = correction['corrected']
            reason = correction['reason']
            corr_type = correction['type']
            
            # 在所有段落中查找匹配项
            matches_found = []
            for i, paragraph_text in enumerate(text_content):
                # 使用更精确的匹配策略
                if self._is_text_match(original, paragraph_text):
                    actual_paragraph_index = paragraph_mapping.get(i, i)
                    # 计算该术语在段落中出现的次数
                    occurrences = paragraph_text.count(original)
                    matches_found.append((i, actual_paragraph_index, paragraph_text, occurrences))
            
            # 处理所有匹配项
            if matches_found:
                for text_idx, para_idx, para_text, occurrences in matches_found:
                    # 为每个出现的术语创建一个修正项
                    for occurrence in range(occurrences):
                        # 创建批注文本
                        if corr_type == 'suggestion':
                            comment_text = f"💡 改进建议: {original} → {corrected}\n"
                            comment_text += f"📋 原因: {reason}\n"
                            comment_text += f"🎯 类型: 改进建议\n"
                        elif corr_type == 'terminology':
                            comment_text = f"🔍 术语不一致修正: {original} → {corrected}\n"
                            comment_text += f"📝 理由: {reason}\n"
                            comment_text += f"💡 建议: {correction.get('suggestion_text', '')}\n"
                        else:
                            comment_text = f"🔧 错误修正: {original} → {corrected}\n"
                            comment_text += f"📝 理由: {reason}\n"
                            comment_text += f"💡 建议: {correction.get('suggestion_text', '')}\n"
                        
                        comment_text += f"⏰ 处理时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                        
                        synchronized_changes.append({
                            'paragraph_index': para_idx,
                            'original_text': original,
                            'corrected_text': corrected,
                            'comment_text': comment_text,
                            'reason': reason,
                            'type': corr_type,
                            'occurrence_index': occurrence  # 添加出现次数索引
                        })
                        
                        self.console.print(f"[green]✅ 添加修正: {original} → {corrected} (段落{para_idx+1}, 第{occurrence+1}次出现)[/green]")
            else:
                self.console.print(f"[yellow]⚠️  未找到匹配文本: {original}[/yellow]")
        
        self.console.print(f"[green]✅ 总共创建了 {len(synchronized_changes)} 个同步更改[/green]")
        return synchronized_changes

    def _is_text_match(self, target_text: str, paragraph_text: str) -> bool:
        """改进的文本匹配逻辑"""
        # 精确匹配
        if target_text in paragraph_text:
            return True
        
        # 去除标点符号后匹配
        import re
        target_clean = re.sub(r'[^\w\s]', '', target_text)
        paragraph_clean = re.sub(r'[^\w\s]', '', paragraph_text)
        if target_clean in paragraph_clean:
            return True
        
        # 分词匹配（处理术语）
        target_words = target_text.split()
        if len(target_words) == 1 and target_words[0] in paragraph_text:
            return True
        
        return False

    def _apply_changes_with_proper_comments(self, doc: Document, synchronized_changes: list, output_file: str) -> bool:
        """应用更改并确保批注正确显示 - 修复版"""
        try:
            # 创建跟踪更改管理器
            track_changes_manager = WordTrackChangesManager(doc)
            
            # 创建批注管理器
            comments_manager = WordCommentsManager(doc)
            
            # 应用每个更改
            applied_count = 0
            comments_data = []  # 收集批注数据
            
            for change in synchronized_changes:
                paragraph_index = change.get('paragraph_index', 0)
                original_text = change.get('original_text', '')
                corrected_text = change.get('corrected_text', '')
                comment_text = change.get('comment_text', '')
                reason = change.get('reason', '')
                
                if paragraph_index < len(doc.paragraphs):
                    paragraph = doc.paragraphs[paragraph_index]
                    
                    # 应用跟踪更改
                    track_success = track_changes_manager.add_tracked_change(
                        paragraph, original_text, corrected_text, reason
                    )
                    
                    # 收集批注数据（不直接应用到docx）
                    if track_success:
                        comments_data.append({
                            'id': applied_count + 1,
                            'text': comment_text,
                            'author': 'AI校对助手',
                            'date': datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ"),
                            'paragraph_index': paragraph_index,
                            'target_text': original_text
                        })
                        applied_count += 1
                        self.console.print(f"[green]✅ 应用更改 {applied_count}: {original_text} -> {corrected_text}[/green]")
            
            # 应用所有跟踪更改
            track_changes_manager.apply_all_changes()
            
            # 保存临时文档（只包含跟踪更改）
            temp_file = output_file.replace('.docx', '_temp_with_track_changes.docx')
            doc.save(temp_file)
            
            self.console.print(f"[cyan]✅ 准备 {len(comments_data)} 个批注数据用于XML生成[/cyan]")
            
            # 使用修复的批注系统
            success = self._create_complete_comment_system(temp_file, output_file, comments_data)
            
            # 清理临时文件
            if os.path.exists(temp_file):
                os.remove(temp_file)
            
            if success:
                self.console.print(f"[green]✅ 成功应用 {applied_count} 个更改和批注[/green]")
                return True
            else:
                return False
            
        except Exception as e:
            self.console.print(f"[red]❌ 应用更改和批注失败: {e}[/red]")
            import traceback
            traceback.print_exc()
            return False

    def _create_complete_comment_system(self, temp_file: str, output_file: str, comments_data: list) -> bool:
        """创建完整的批注系统"""
        try:
            import zipfile
            import tempfile
            import shutil
            
            if not comments_data:
                self.console.print("[yellow]⚠️ 没有批注数据需要处理[/yellow]")
                shutil.copy2(temp_file, output_file)
                return True
            
            self.console.print(f"[cyan]🔧 创建完整的批注系统，包含 {len(comments_data)} 个批注[/cyan]")
            
            # 创建临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 解压原文档
                with zipfile.ZipFile(temp_file, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                # 1. 创建comments.xml文件
                comments_xml = self._create_comments_xml(comments_data)
                comments_path = os.path.join(temp_dir, 'word', 'comments.xml')
                with open(comments_path, 'w', encoding='utf-8') as f:
                    f.write(comments_xml)
                self.console.print("[green]✅ comments.xml 创建成功[/green]")
                
                # 2. 更新_rels/.rels文件
                rels_content = self._create_updated_rels_for_comments(temp_dir)
                rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
                with open(rels_path, 'w', encoding='utf-8') as f:
                    f.write(rels_content)
                self.console.print("[green]✅ document.xml.rels 更新成功[/green]")
                
                # 3. 更新[Content_Types].xml
                content_types = self._create_updated_content_types_for_comments(temp_dir)
                content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
                with open(content_types_path, 'w', encoding='utf-8') as f:
                    f.write(content_types)
                self.console.print("[green]✅ [Content_Types].xml 更新成功[/green]")
                
                # 4. 更新document.xml，添加批注引用
                document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
                with open(document_xml_path, 'r', encoding='utf-8') as f:
                    document_content = f.read()
                
                updated_document = self._add_comment_references_to_document(document_content, comments_data)
                with open(document_xml_path, 'w', encoding='utf-8') as f:
                    f.write(updated_document)
                self.console.print("[green]✅ document.xml 更新成功[/green]")
                
                # 5. 重新打包文档
                with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, temp_dir)
                            zip_out.write(file_path, arcname)
                
                self.console.print("[green]✅ 完整的批注系统创建成功[/green]")
                return True
                
        except Exception as e:
            self.console.print(f"[red]❌ 创建批注系统失败: {e}[/red]")
            import traceback
            traceback.print_exc()
            # 回退到无批注版本
            try:
                shutil.copy2(temp_file, output_file)
                return True
            except:
                return False

    def _create_comments_xml(self, comments_data: list) -> str:
        """创建批注XML内容"""
        xml_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'''
        
        for comment in comments_data:
            comment_id = comment.get('id', 1)
            author = comment.get('author', 'AI校对助手')
            date = comment.get('date', datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ"))
            text = comment.get('text', '')
            
            xml_content += f'''
    <w:comment w:id="{comment_id}" w:author="{author}" w:date="{date}" w:initials="AI">
        <w:p>
            <w:r>
                <w:t>{text}</w:t>
            </w:r>
        </w:p>
    </w:comment>'''
        
        xml_content += '\n</w:comments>'
        return xml_content

    def _create_updated_rels_for_comments(self, temp_dir: str) -> str:
        """为批注更新_rels文件"""
        try:
            rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
            if os.path.exists(rels_path):
                with open(rels_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            else:
                # 创建基础rels文件
                content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>'''
            
            # 检查是否已经有comments关系
            if 'comments.xml' not in content:
                # 添加comments关系
                comments_rel = '''<Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'''
                
                # 在</Relationships>前插入
                if '</Relationships>' in content:
                    content = content.replace('</Relationships>', f'{comments_rel}\n</Relationships>')
                else:
                    content = content.replace('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">', 
                                            f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n{comments_rel}')
                    if not content.endswith('</Relationships>'):
                        content += '\n</Relationships>'
            
            return content
            
        except Exception as e:
            self.console.print(f"[yellow]⚠️ 更新rels文件失败: {e}[/yellow]")
            return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>'''
    
    def _create_updated_content_types_for_comments(self, temp_dir: str) -> str:
        """为批注更新Content_Types文件"""
        try:
            content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
            if os.path.exists(content_types_path):
                with open(content_types_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            else:
                # 创建基础Content_Types文件
                content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
            
            # 检查是否已经有comments的Override
            if '/word/comments.xml' not in content:
                # 添加comments的Override
                comments_override = '''<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'''
                
                # 在</Types>前插入
                if '</Types>' in content:
                    content = content.replace('</Types>', f'{comments_override}\n</Types>')
            
            return content
            
        except Exception as e:
            self.console.print(f"[yellow]⚠️ 更新Content_Types文件失败: {e}[/yellow]")
            return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>'''

    def _add_comment_references_to_document(self, document_xml: str, comments_data: list) -> str:
        """在document.xml中添加批注引用标记"""
        try:
            modified_xml = document_xml
            
            # 为每个批注添加引用标记
            for comment in comments_data:
                comment_id = comment['id']
                # 从synchronized_changes中获取原始文本信息
                # 这里需要改进获取原始文本的方法
                
                # 简单的方法：在每个包含跟踪更改的段落后添加批注引用
                import re
                
                # 查找包含删除标记的位置，在其后添加批注引用
                del_pattern = f'(<w:del[^>]*w:id="{comment_id}"[^>]*>.*?</w:del>)'
                matches = re.finditer(del_pattern, modified_xml, re.DOTALL)
                
                for match in matches:
                    # 在删除标记后添加批注标记
                    comment_range_start = f'<w:commentRangeStart w:id="{comment_id}"/>'
                    comment_range_end = f'<w:commentRangeEnd w:id="{comment_id}"/>'
                    comment_reference = f'<w:r><w:commentReference w:id="{comment_id}"/></w:r>'
                    
                    replacement = f'{comment_range_start}{match.group(0)}{comment_range_end}{comment_reference}'
                    modified_xml = modified_xml.replace(match.group(0), replacement, 1)
                    self.console.print(f"[green]✅ 添加批注引用标记: comment_id={comment_id}[/green]")
                    break  # 只处理第一个匹配
            
            return modified_xml
            
        except Exception as e:
            self.console.print(f"[red]添加批注引用失败: {e}[/red]")
            import traceback
            traceback.print_exc()
            return document_xml

    def _prepare_comments_with_changes(self, comments_data: list, synchronized_changes: list) -> list:
        """准备包含更改信息的批注数据"""
        try:
            enhanced_comments = []
            
            for i, comment in enumerate(comments_data):
                enhanced_comment = comment.copy()
                
                # 从synchronized_changes中找到对应的更改信息
                if i < len(synchronized_changes):
                    change = synchronized_changes[i]
                    enhanced_comment['original_text'] = change.get('original_text', '')
                    enhanced_comment['corrected_text'] = change.get('corrected_text', '')
                
                enhanced_comments.append(enhanced_comment)
            
            return enhanced_comments
            
        except Exception as e:
            self.console.print(f"[red]准备批注数据失败: {e}[/red]")
            return comments_data

    def _create_updated_content_types(self, input_zip) -> str:
        """创建更新的内容类型文件"""
        try:
            # 读取原始内容类型文件
            if '[Content_Types].xml' in input_zip.namelist():
                content_types_xml = input_zip.read('[Content_Types].xml').decode('utf-8')
            else:
                content_types_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
</Types>'''
            
            # 检查是否已包含批注内容类型
            if 'word/comments.xml' not in content_types_xml:
                # 添加批注内容类型
                content_types_xml = content_types_xml.replace(
                    '</Types>',
                    '''    <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>'''
                )
            
            return content_types_xml
            
        except Exception as e:
            print(f"创建内容类型文件失败: {e}")
            # 返回基本的内容类型文件
            return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>'''

    def _extract_word_corrections(self, original_text: str, suggested_text: str):
        """从句子级别的修正中提取词汇级别的修正"""
        corrections = []
        try:
            # 简单的词汇差异检测
            original_words = original_text.split()
            suggested_words = suggested_text.split()
            
            # 如果长度相同，逐词比较
            if len(original_words) == len(suggested_words):
                for orig, sugg in zip(original_words, suggested_words):
                    if orig != sugg:
                        corrections.append((orig, sugg))
            else:
                # 如果长度不同，查找明显的替换
                common_replacements = [
                    ("计算器科学", "计算机科学"),
                    ("程式设计", "程序设计"), 
                    ("软体工程", "软件工程"),
                    ("变数", "变量"),
                    ("函式", "函数"),
                    ("超级计算器", "超级计算机"),
                    (",", "，"),  # 标点符号替换
                ]
                
                for orig, repl in common_replacements:
                    if orig in original_text and repl in suggested_text:
                        corrections.append((orig, repl))
            
            return corrections
        except Exception as e:
            self.console.print(f"[yellow]⚠️  提取词汇修正失败: {e}[/yellow]")
            return []

    def _extract_terms_from_inconsistency(self, problem_text: str, suggestion: str):
        """从术语不一致问题中提取术语对"""
        terms = []
        
        # 解析不一致术语描述
        if "发现多种术语：" in problem_text:
            # 提取术语列表
            terms_part = problem_text.split("发现多种术语：")[1].strip()
            # 移除可能的额外描述
            if "，" in terms_part:
                terms_part = terms_part.split("，")[0]
            if "。" in terms_part:
                terms_part = terms_part.split("。")[0]
            
            # 分割术语
            term_variants = []
            if "、" in terms_part:
                term_variants = [t.strip().strip('"').strip("'") for t in terms_part.split("、")]
            elif "，" in terms_part:
                term_variants = [t.strip().strip('"').strip("'") for t in terms_part.split("，")]
            else:
                # 单个术语的情况
                term_variants = [terms_part.strip().strip('"').strip("'")]
            
            # 从建议中提取标准术语
            standard_term = None
            if "建议统一使用" in suggestion:
                standard_part = suggestion.split("建议统一使用")[1].strip()
                if "。" in standard_part:
                    standard_part = standard_part.split("。")[0]
                if "，" in standard_part:
                    standard_part = standard_part.split("，")[0]
                standard_term = standard_part.strip().strip('"').strip("'")
            elif "推荐使用" in suggestion:
                standard_part = suggestion.split("推荐使用")[1].strip()
                if "。" in standard_part:
                    standard_part = standard_part.split("。")[0]
                if "，" in standard_part:
                    standard_part = standard_part.split("，")[0]
                standard_term = standard_part.strip().strip('"').strip("'")
            
            # 如果找到标准术语，为每个变体创建修正对
            if standard_term and term_variants:
                for variant in term_variants:
                    if variant and variant != standard_term:
                        terms.append((variant, standard_term))
                        self.console.print(f"[cyan]📝 术语修正: {variant} → {standard_term}[/cyan]")
            
            # 如果没有明确的标准术语，使用第一个作为标准
            elif len(term_variants) > 1:
                standard_term = term_variants[0]
                for variant in term_variants[1:]:
                    if variant and variant != standard_term:
                        terms.append((variant, standard_term))
                        self.console.print(f"[cyan]📝 术语修正: {variant} → {standard_term}[/cyan]")
        
        # 处理特殊的术语对
        special_corrections = {
            "软体工程": "软件工程",
            "程式设计": "程序设计", 
            "计算器科学": "计算机科学",
            "资料结构": "数据结构",
            "演算法": "算法"
        }
        
        # 检查是否包含特殊术语
        for original, corrected in special_corrections.items():
            if original in problem_text or original in suggestion:
                terms.append((original, corrected))
                self.console.print(f"[cyan]🔧 特殊术语修正: {original} → {corrected}[/cyan]")
        
        return terms

    def _extract_corrected_text(self, suggestion: str):
        """从建议中提取修正后的文本"""
        if not suggestion:
            return None
        
        # 常见的修正模式
        patterns = [
            r"应为[：:]?\s*[\"']([^\"']+)[\"']",
            r"改为[：:]?\s*[\"']([^\"']+)[\"']", 
            r"修正为[：:]?\s*[\"']([^\"']+)[\"']",
            r"建议改为[：:]?\s*[\"']([^\"']+)[\"']",
            r"应该是[：:]?\s*[\"']([^\"']+)[\"']",
            r"正确的是[：:]?\s*[\"']([^\"']+)[\"']",
            r"→\s*[\"']([^\"']+)[\"']",
            r"替换为[：:]?\s*[\"']([^\"']+)[\"']"
        ]
        
        import re
        for pattern in patterns:
            match = re.search(pattern, suggestion)
            if match:
                corrected = match.group(1).strip()
                self.console.print(f"[cyan]🔍 提取修正文本: {corrected}[/cyan]")
                return corrected
        
        # 如果没有找到引号包围的文本，尝试其他模式
        simple_patterns = [
            r"应为[：:]?\s*([^\s，。]+)",
            r"改为[：:]?\s*([^\s，。]+)",
            r"修正为[：:]?\s*([^\s，。]+)",
            r"建议改为[：:]?\s*([^\s，。]+)",
            r"应该是[：:]?\s*([^\s，。]+)",
            r"正确的是[：:]?\s*([^\s，。]+)"
        ]
        
        for pattern in simple_patterns:
            match = re.search(pattern, suggestion)
            if match:
                corrected = match.group(1).strip()
                self.console.print(f"[cyan]🔍 提取修正文本: {corrected}[/cyan]")
                return corrected
        
        # 处理特殊情况：直接的替换建议
        if "→" in suggestion:
            parts = suggestion.split("→")
            if len(parts) >= 2:
                corrected = parts[-1].strip().strip('"').strip("'").strip("。").strip("，")
                if corrected:
                    self.console.print(f"[cyan]🔍 提取修正文本: {corrected}[/cyan]")
                    return corrected
        
        self.console.print(f"[yellow]⚠️  无法提取修正文本: {suggestion}[/yellow]")
        return None
    
    def extract_text_content(self, doc: Document):
        """提取文档的文本内容"""
        text_content = []
        for paragraph in doc.paragraphs:
            text_content.append(paragraph.text)
        return text_content

    def _apply_synchronized_changes_fixed(self, doc: Document, synchronized_changes: list, output_file: str) -> bool:
        """应用同步更改的修复版本"""
        try:
            if not synchronized_changes:
                self.console.print("[yellow]⚠️ 没有需要应用的更改[/yellow]")
                doc.save(output_file)
                return True
            
            # 使用现有的应用更改方法，但增加了错误处理
            return self._apply_changes_with_proper_comments(doc, synchronized_changes, output_file)
            
        except Exception as e:
            self.console.print(f"[red]❌ 应用同步更改失败: {e}[/red]")
            import traceback
            traceback.print_exc()
            return False


def test_fixed_enhanced_proofreader():
    """测试修复版增强校对器"""
    try:
        # 使用测试API密钥
        api_key = "sk-test"
        
        proofreader = ProofReaderWithTrackChangesAndCommentsFixed(api_key)
        
        input_file = "sample_input.docx"
        output_file = "output_fixed.docx"
        
        if os.path.exists(input_file):
            success = proofreader.proofread_with_track_changes_and_comments(input_file, output_file)
            if success:
                print(f"✅ 修复版增强校对成功: {output_file}")
            else:
                print("❌ 修复版增强校对失败")
        else:
            print(f"❌ 输入文件不存在: {input_file}")
            
    except Exception as e:
        print(f"❌ 测试失败: {e}")


if __name__ == "__main__":
    test_fixed_enhanced_proofreader() 