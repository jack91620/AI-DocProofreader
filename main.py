#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
AI校对助手 - 主程序入口
"""

import click
import os
import sys
from rich.console import Console

from proofreader import ProofReader, Config


console = Console()


@click.group()
@click.version_option(version="1.0.0")
def cli():
    """AI校对助手 - 专业的中文计算机教材校对工具"""
    pass


@cli.command()
@click.option('--input', '-i', required=True, help='输入的docx文件路径')
@click.option('--output', '-o', required=True, help='输出的docx文件路径')
@click.option('--config', '-c', help='配置文件路径')
def proofread(input, output, config):
    """校对单个docx文档"""
    try:
        # 检查输入文件是否存在
        if not os.path.exists(input):
            console.print(f"[red]错误：输入文件不存在 - {input}[/red]")
            sys.exit(1)
        
        # 检查文件格式
        if not input.lower().endswith('.docx'):
            console.print("[red]错误：只支持.docx格式的文件[/red]")
            sys.exit(1)
        
        # 创建校对器
        proofreader_config = Config()
        proofreader = ProofReader(proofreader_config)
        
        # 执行校对
        if proofreader.proofread_document(input, output):
            console.print(f"[green]✅ 校对完成！输出文件：{output}[/green]")
        else:
            console.print("[red]❌ 校对失败[/red]")
            sys.exit(1)
            
    except Exception as e:
        console.print(f"[red]程序执行错误：{e}[/red]")
        sys.exit(1)


@cli.command()
@click.option('--input-dir', '-i', required=True, help='输入文件夹路径')
@click.option('--output-dir', '-o', required=True, help='输出文件夹路径')
def batch(input_dir, output_dir):
    """批量校对文件夹中的所有docx文档"""
    try:
        # 检查输入目录是否存在
        if not os.path.exists(input_dir):
            console.print(f"[red]错误：输入目录不存在 - {input_dir}[/red]")
            sys.exit(1)
        
        # 创建校对器
        proofreader_config = Config()
        proofreader = ProofReader(proofreader_config)
        
        # 执行批量校对
        if proofreader.batch_proofread(input_dir, output_dir):
            console.print(f"[green]✅ 批量校对完成！输出目录：{output_dir}[/green]")
        else:
            console.print("[red]❌ 批量校对失败[/red]")
            sys.exit(1)
            
    except Exception as e:
        console.print(f"[red]程序执行错误：{e}[/red]")
        sys.exit(1)


@cli.command()
@click.option('--text', '-t', help='要检查的文本内容')
@click.option('--file', '-f', help='包含文本的文件路径')
def check(text, file):
    """快速检查文本片段"""
    try:
        # 获取文本内容
        if file:
            if not os.path.exists(file):
                console.print(f"[red]错误：文件不存在 - {file}[/red]")
                sys.exit(1)
            
            with open(file, 'r', encoding='utf-8') as f:
                text_content = f.read()
        elif text:
            text_content = text
        else:
            console.print("[red]错误：请提供要检查的文本内容或文件路径[/red]")
            sys.exit(1)
        
        # 创建校对器
        proofreader_config = Config()
        proofreader = ProofReader(proofreader_config)
        
        # 执行快速检查
        console.print("[blue]正在检查文本...[/blue]")
        result = proofreader.quick_check(text_content)
        
        # 显示结果
        if result.issues:
            console.print(f"[yellow]发现 {len(result.issues)} 个问题：[/yellow]")
            for i, issue in enumerate(result.issues, 1):
                console.print(f"{i}. {issue['type']}: {issue['text']}")
                console.print(f"   建议: {issue['suggestion']}")
                console.print()
        else:
            console.print("[green]✅ 未发现明显问题[/green]")
        
        if result.suggestions:
            console.print(f"[blue]改进建议 ({len(result.suggestions)} 条)：[/blue]")
            for i, suggestion in enumerate(result.suggestions, 1):
                console.print(f"{i}. 原文: {suggestion['original']}")
                console.print(f"   建议: {suggestion['suggested']}")
                console.print(f"   理由: {suggestion['reason']}")
                console.print()
        
    except Exception as e:
        console.print(f"[red]程序执行错误：{e}[/red]")
        sys.exit(1)


@cli.command()
def setup():
    """设置配置和环境检查"""
    console.print("[blue]AI校对助手环境检查[/blue]")
    
    try:
        config = Config()
        config.validate()
        console.print("[green]✅ 配置验证通过[/green]")
        
        console.print(f"[cyan]OpenAI模型: {config.ai.model}[/cyan]")
        console.print(f"[cyan]最大Token数: {config.ai.max_tokens}[/cyan]")
        console.print(f"[cyan]温度参数: {config.ai.temperature}[/cyan]")
        
        console.print("[green]环境配置正常，可以开始使用校对功能！[/green]")
        
    except Exception as e:
        console.print(f"[red]❌ 配置检查失败: {e}[/red]")
        console.print("[yellow]请检查以下项目：[/yellow]")
        console.print("1. 是否设置了OPENAI_API_KEY环境变量")
        console.print("2. 网络连接是否正常")
        console.print("3. API密钥是否有效")
        sys.exit(1)


@cli.command()
def demo():
    """运行演示示例"""
    demo_text = """
Python是一种解释型、面向对象、动态数据类型的高级程序设计语言。Python由Guido van Rossum于1989年底发明，第一个公开发行版发行于1991年。
Python的设计理念是优雅、明确、简单。Python开发者的哲学是"用一种方法，最好是只有一种方法来做一件事"。
在设计Python语言时，如果面临多种选择，Python开发者一般会拒绝花俏的语法，而选择明确没有或者很少有歧义的语法。
"""
    
    console.print("[blue]运行演示示例...[/blue]")
    console.print(f"[cyan]示例文本:[/cyan]\n{demo_text}")
    
    try:
        proofreader_config = Config()
        proofreader = ProofReader(proofreader_config)
        
        result = proofreader.quick_check(demo_text)
        
        if result.issues:
            console.print(f"\n[yellow]发现问题 ({len(result.issues)} 个):[/yellow]")
            for issue in result.issues:
                console.print(f"• {issue['type']}: {issue['suggestion']}")
        else:
            console.print("\n[green]✅ 文本质量良好，未发现明显问题[/green]")
        
        if result.suggestions:
            console.print(f"\n[blue]改进建议 ({len(result.suggestions)} 条):[/blue]")
            for suggestion in result.suggestions:
                console.print(f"• {suggestion['reason']}")
        
        console.print("\n[green]演示完成！[/green]")
        
    except Exception as e:
        console.print(f"[red]演示失败: {e}[/red]")
        sys.exit(1)


if __name__ == '__main__':
    cli() 