#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
检查Word批注内容的脚本
"""

from docx import Document
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
import re

console = Console()

def extract_comments_from_text(text):
    """从文本中提取批注内容"""
    # 查找 [批注: ...] 格式的批注
    comment_pattern = r'\[批注:\s*([^\]]+)\]'
    comments = re.findall(comment_pattern, text)
    return comments

def analyze_document_comments(file_path):
    """分析文档中的批注内容"""
    try:
        doc = Document(file_path)
        console.print(f"[bold blue]📄 分析文档: {file_path}[/bold blue]\n")
        
        total_paragraphs = 0
        paragraphs_with_comments = 0
        total_comments = 0
        all_comments = []
        
        for i, paragraph in enumerate(doc.paragraphs, 1):
            if paragraph.text.strip():
                total_paragraphs += 1
                
                # 检查段落中的批注
                comments_in_paragraph = extract_comments_from_text(paragraph.text)
                
                if comments_in_paragraph:
                    paragraphs_with_comments += 1
                    total_comments += len(comments_in_paragraph)
                    
                    console.print(f"[cyan]段落 {i}:[/cyan]")
                    console.print(f"  原文: {paragraph.text[:100]}{'...' if len(paragraph.text) > 100 else ''}")
                    
                    for j, comment in enumerate(comments_in_paragraph, 1):
                        console.print(f"  [yellow]批注 {j}:[/yellow] {comment}")
                        all_comments.append(comment)
                    
                    console.print()
        
        # 显示统计信息
        table = Table(title="批注内容分析结果")
        table.add_column("统计项目", style="cyan")
        table.add_column("数量", style="green")
        
        table.add_row("总段落数", str(total_paragraphs))
        table.add_row("包含批注的段落", str(paragraphs_with_comments))
        table.add_row("批注总数", str(total_comments))
        
        console.print(table)
        
        if total_comments > 0:
            console.print(f"\n[bold green]✅ 成功找到 {total_comments} 个批注内容！[/bold green]")
            
            # 显示所有批注的分类
            typo_comments = [c for c in all_comments if "错别字" in c or "用词不当" in c]
            term_comments = [c for c in all_comments if "建议修改" in c]
            punct_comments = [c for c in all_comments if "标点符号" in c]
            
            if typo_comments:
                console.print(f"\n[bold red]🔍 错别字和用词问题 ({len(typo_comments)} 个):[/bold red]")
                for comment in typo_comments[:3]:  # 显示前3个
                    console.print(f"  • {comment[:80]}{'...' if len(comment) > 80 else ''}")
            
            if term_comments:
                console.print(f"\n[bold yellow]📝 修改建议 ({len(term_comments)} 个):[/bold yellow]")
                for comment in term_comments[:3]:  # 显示前3个
                    console.print(f"  • {comment[:80]}{'...' if len(comment) > 80 else ''}")
            
            if punct_comments:
                console.print(f"\n[bold blue]🔤 标点符号问题 ({len(punct_comments)} 个):[/bold blue]")
                for comment in punct_comments:
                    console.print(f"  • {comment[:80]}{'...' if len(comment) > 80 else ''}")
        else:
            console.print(f"\n[bold red]❌ 未找到批注内容[/bold red]")
        
        return total_comments > 0
        
    except Exception as e:
        console.print(f"[red]分析失败: {e}[/red]")
        return False

def compare_comment_versions():
    """对比不同版本的批注效果"""
    console.print(Panel.fit("[bold blue]批注内容对比分析[/bold blue]"))
    
    files_to_check = [
        ("sample_input.docx", "原始输入文档"),
        ("sample_output_with_word_comments.docx", "旧版批注系统"),
        ("sample_output_with_full_comments.docx", "新版完整批注系统")
    ]
    
    for filename, description in files_to_check:
        console.print(f"\n{'='*60}")
        console.print(f"[bold green]{description}[/bold green]")
        
        try:
            has_comments = analyze_document_comments(filename)
            
            if filename == "sample_output_with_full_comments.docx" and has_comments:
                console.print("[green]🎉 新版批注系统工作正常！[/green]")
            elif filename == "sample_input.docx":
                console.print("[blue]ℹ️ 这是原始文档，无批注[/blue]")
            elif not has_comments:
                console.print("[yellow]⚠️ 未检测到批注内容[/yellow]")
                
        except FileNotFoundError:
            console.print(f"[red]❌ 文件不存在: {filename}[/red]")

if __name__ == "__main__":
    compare_comment_versions()
    
    console.print(f"\n[bold cyan]💡 使用建议:[/bold cyan]")
    console.print("1. 用Microsoft Word打开 sample_output_with_full_comments.docx")
    console.print("2. 查看高亮文本和紧跟其后的红色批注内容")
    console.print("3. 批注内容格式为：[批注: 具体问题和建议]")
    console.print("4. 高亮部分表示有问题的文本，红色斜体部分是AI的修改建议") 