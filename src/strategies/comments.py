"""
批注模式策略 - 在Word审阅窗格中显示批注建议
"""

import os
from typing import List
from datetime import datetime
from docx import Document
from rich.console import Console

from .base import OutputStrategy, Correction
from word_ops.comments import WordCommentsManager, add_comments_to_docx


class CommentsStrategy(OutputStrategy):
    """批注模式输出策略"""

    def __init__(self):
        self.console = Console()

    def get_output_suffix(self) -> str:
        return "proofread"

    def apply_corrections(self, doc: Document, corrections: List[Correction], text_content: List[str],
                          input_file: str, output_file: str) -> bool:
        try:
            # 创建Word批注管理器
            comments_manager = WordCommentsManager(doc)

            # 添加批注和修正
            comments_data = []
            comment_count = 0

            for correction in corrections:
                # 构建批注内容
                comment_text = correction.corrected_text
                if correction.reason:
                    comment_text += f"\n理由: {correction.reason}"

                # 获取对应段落
                if correction.paragraph_index < len(doc.paragraphs):
                    paragraph = doc.paragraphs[correction.paragraph_index]

                    if comments_manager.add_comment(paragraph, correction.original_text, comment_text):
                        comment_count += 1
                        comments_data.append({
                            'id': comment_count,
                            'text': comment_text,
                            'author': 'AI校对助手',
                            'date': datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
                        })
                        self.console.print(f"[green]✅ 添加Word审阅批注 {comment_count}: {correction.original_text} -> {correction.corrected_text}[/green]")

            self.console.print(f"[blue]📝 总共添加了 {comment_count} 个Word审阅批注[/blue]")

            # 完成文档处理
            comments_manager.finalize_document()

            # 保存临时文档
            temp_file = output_file.replace('.docx', '_temp.docx')
            doc.save(temp_file)

            # 使用完整的Word审阅批注功能
            if add_comments_to_docx(temp_file, output_file, comments_data):
                os.remove(temp_file)
                self.console.print(f"[green]✅ 校对完成，输出文件: {output_file}[/green]")
                return True
            else:
                os.rename(temp_file, output_file)
                self.console.print(f"[yellow]⚠️ 审阅批注添加失败，使用基础版本: {output_file}[/yellow]")
                return True

        except Exception as e:
            self.console.print(f"[red]❌ 批注模式校对失败: {e}[/red]")
            return False
