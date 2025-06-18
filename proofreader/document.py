"""
文档处理模块
"""

import re
from typing import List, Dict, Tuple
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.shared import qn
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml


class DocumentProcessor:
    """文档处理器"""
    
    def __init__(self):
        self.document = None
        self.paragraphs = []
        self.text_content = ""
    
    def load_document(self, file_path: str) -> bool:
        """加载docx文档"""
        try:
            self.document = Document(file_path)
            self._extract_text()
            return True
        except Exception as e:
            print(f"加载文档失败: {e}")
            return False
    
    def _extract_text(self):
        """提取文档中的文本内容"""
        self.paragraphs = []
        all_text = []
        
        for paragraph in self.document.paragraphs:
            if paragraph.text.strip():
                self.paragraphs.append({
                    'text': paragraph.text,
                    'paragraph': paragraph,
                    'index': len(self.paragraphs)
                })
                all_text.append(paragraph.text)
        
        self.text_content = '\n'.join(all_text)
    
    def get_text_segments(self, max_length: int = 2000) -> List[str]:
        """将文档文本分割成适合AI处理的段落"""
        segments = []
        current_segment = ""
        
        for para_info in self.paragraphs:
            text = para_info['text']
            
            # 如果当前段落加上新文本超过最大长度，保存当前段落并开始新段落
            if len(current_segment) + len(text) > max_length and current_segment:
                segments.append(current_segment.strip())
                current_segment = text
            else:
                current_segment += "\n" + text if current_segment else text
        
        # 添加最后一个段落
        if current_segment:
            segments.append(current_segment.strip())
        
        return segments
    
    def add_comment(self, paragraph_index: int, text: str, comment: str, 
                   author: str = "AI校对助手", color: str = "red"):
        """在指定段落添加批注"""
        try:
            if paragraph_index >= len(self.paragraphs):
                return False
            
            paragraph = self.paragraphs[paragraph_index]['paragraph']
            
            # 查找要批注的文本
            if text in paragraph.text:
                # 创建批注
                self._add_comment_to_paragraph(paragraph, text, comment, author, color)
                return True
            
        except Exception as e:
            print(f"添加批注失败: {e}")
        
        return False
    
    def _add_comment_to_paragraph(self, paragraph, target_text: str, 
                                comment: str, author: str, color: str):
        """在段落中添加批注"""
        # 清空段落内容
        paragraph.clear()
        
        # 重新构建段落，添加批注
        original_text = paragraph.text if hasattr(paragraph, '_original_text') else target_text
        
        # 查找目标文本位置
        start_pos = original_text.find(target_text)
        if start_pos == -1:
            # 如果找不到，直接添加到段落末尾
            run = paragraph.add_run(original_text)
            comment_run = paragraph.add_run(f"【批注：{comment}】")
            comment_run.font.color.rgb = RGBColor(255, 0, 0)  # 红色
            comment_run.font.bold = True
        else:
            # 添加目标文本之前的内容
            if start_pos > 0:
                paragraph.add_run(original_text[:start_pos])
            
            # 添加目标文本（高亮显示）
            highlighted_run = paragraph.add_run(target_text)
            highlighted_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            
            # 添加批注
            comment_run = paragraph.add_run(f"【批注：{comment}】")
            comment_run.font.color.rgb = RGBColor(255, 0, 0)  # 红色
            comment_run.font.bold = True
            
            # 添加目标文本之后的内容
            end_pos = start_pos + len(target_text)
            if end_pos < len(original_text):
                paragraph.add_run(original_text[end_pos:])
    
    def highlight_text(self, paragraph_index: int, text: str, 
                      color: WD_COLOR_INDEX = WD_COLOR_INDEX.YELLOW):
        """高亮显示文本"""
        try:
            if paragraph_index >= len(self.paragraphs):
                return False
            
            paragraph = self.paragraphs[paragraph_index]['paragraph']
            
            # 在段落中查找并高亮文本
            for run in paragraph.runs:
                if text in run.text:
                    run.font.highlight_color = color
                    return True
            
        except Exception as e:
            print(f"高亮文本失败: {e}")
        
        return False
    
    def save_document(self, output_path: str) -> bool:
        """保存文档"""
        try:
            if self.document:
                self.document.save(output_path)
                return True
        except Exception as e:
            print(f"保存文档失败: {e}")
        
        return False
    
    def get_paragraph_by_text(self, text: str) -> Tuple[int, str]:
        """根据文本内容查找段落"""
        for i, para_info in enumerate(self.paragraphs):
            if text in para_info['text']:
                return i, para_info['text']
        return -1, ""
    
    def get_statistics(self) -> Dict:
        """获取文档统计信息"""
        return {
            "paragraph_count": len(self.paragraphs),
            "character_count": len(self.text_content),
            "word_count": len(self.text_content.replace(' ', '')),  # 中文字符数
        } 