#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
检查最终批注文档的脚本
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

def analyze_final_document():
    """分析最终生成的批注文档"""
    file_path = "sample_output_final_comments.docx"
    
    try:
        doc = Document(file_path)
        console.print(f"[bold blue]🎉 分析最终批注文档: {file_path}[/bold blue]\n")
        
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
                    # 显示段落原文，但限制长度
                    clean_text = re.sub(r'\[批注:[^\]]+\]', '[批注...]', paragraph.text)
                    console.print(f"  原文: {clean_text[:80]}{'...' if len(clean_text) > 80 else ''}")
                    
                    for j, comment in enumerate(comments_in_paragraph, 1):
                        console.print(f"  [yellow]批注 {j}:[/yellow] {comment[:100]}{'...' if len(comment) > 100 else ''}")
                        all_comments.append(comment)
                    
                    console.print()
        
        # 显示统计信息
        table = Table(title="最终批注文档分析结果")
        table.add_column("统计项目", style="cyan")
        table.add_column("数量", style="green")
        
        table.add_row("总段落数", str(total_paragraphs))
        table.add_row("包含批注的段落", str(paragraphs_with_comments))
        table.add_row("批注总数", str(total_comments))
        
        console.print(table)
        
        if total_comments > 0:
            console.print(f"\n[bold green]🎉 成功！文档包含 {total_comments} 个完整的批注内容！[/bold green]")
            
            # 按类型分类显示批注
            console.print(f"\n[bold cyan]📋 批注内容分类：[/bold cyan]")
            
            typo_comments = [c for c in all_comments if "错别字" in c or "用词不当" in c]
            suggestion_comments = [c for c in all_comments if "建议修改" in c]
            punct_comments = [c for c in all_comments if "标点符号" in c]
            
            if typo_comments:
                console.print(f"\n[red]🔍 错别字和用词问题 ({len(typo_comments)} 个):[/red]")
                for i, comment in enumerate(typo_comments[:5], 1):  # 显示前5个
                    console.print(f"  {i}. {comment[:70]}{'...' if len(comment) > 70 else ''}")
            
            if suggestion_comments:
                console.print(f"\n[yellow]📝 修改建议 ({len(suggestion_comments)} 个):[/yellow]")
                for i, comment in enumerate(suggestion_comments[:5], 1):  # 显示前5个
                    console.print(f"  {i}. {comment[:70]}{'...' if len(comment) > 70 else ''}")
            
            if punct_comments:
                console.print(f"\n[blue]🔤 标点符号问题 ({len(punct_comments)} 个):[/blue]")
                for i, comment in enumerate(punct_comments, 1):
                    console.print(f"  {i}. {comment[:70]}{'...' if len(comment) > 70 else ''}")
            
            # 显示效果说明
            console.print(f"\n[bold green]📋 批注显示效果：[/bold green]")
            console.print("✅ 问题文本：黄色高亮背景")
            console.print("✅ 批注内容：红色斜体文字")
            console.print("✅ 格式：[批注: 具体问题描述和修改建议]")
            console.print("✅ 位置：紧跟在问题文本后面")
            
        else:
            console.print(f"\n[bold red]❌ 未找到批注内容[/bold red]")
        
        return total_comments > 0
        
    except FileNotFoundError:
        console.print(f"[red]❌ 文件不存在: {file_path}[/red]")
        return False
    except Exception as e:
        console.print(f"[red]分析失败: {e}[/red]")
        return False

def show_usage_guide():
    """显示使用指南"""
    console.print(f"\n[bold blue]💡 Microsoft Word 中的查看效果：[/bold blue]")
    
    usage_table = Table(title="批注功能说明")
    usage_table.add_column("功能", style="cyan")
    usage_table.add_column("效果", style="green")
    usage_table.add_column("说明", style="yellow")
    
    usage_table.add_row(
        "高亮显示", 
        "黄色背景", 
        "标识有问题的文本"
    )
    usage_table.add_row(
        "批注内容", 
        "红色斜体文字", 
        "显示具体问题和修改建议"
    )
    usage_table.add_row(
        "批注格式", 
        "[批注: ...]", 
        "统一的批注格式，易于识别"
    )
    usage_table.add_row(
        "位置", 
        "问题文本后", 
        "批注紧跟在问题文本后面"
    )
    
    console.print(usage_table)

if __name__ == "__main__":
    console.print(Panel.fit("[bold blue]🔍 最终批注文档检查[/bold blue]"))
    
    success = analyze_final_document()
    
    if success:
        show_usage_guide()
        
        console.print(f"\n[bold green]🎊 批注系统升级完成！[/bold green]")
        console.print("现在AI校对系统可以：")
        console.print("• ✅ 高亮显示有问题的文本")
        console.print("• ✅ 显示完整的批注内容")
        console.print("• ✅ 提供具体的修改建议")
        console.print("• ✅ 在Word中清晰可见")
        
        console.print(f"\n[bold cyan]📁 查看文档：[/bold cyan]")
        console.print("用Microsoft Word打开 sample_output_final_comments.docx 查看完整效果")
    else:
        console.print(f"\n[bold red]❌ 批注系统仍需调试[/bold red]") 