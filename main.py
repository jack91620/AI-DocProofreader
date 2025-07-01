#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
AI校对助手 - 主程序入口
"""

import click
import os
import sys
from rich.console import Console
import configparser
from pathlib import Path

# 添加src目录到Python路径
src_path = Path(__file__).parent / "src" 
sys.path.insert(0, str(src_path))

from src.proofreader import ProofReader
from src.config import Config
from src.proofreader_revisions import ProofReaderWithRevisions
from src.proofreader_track_changes import ProofReaderWithTrackChanges
from src.proofreader_track_changes_enhanced import ProofReaderWithTrackChangesAndComments
from src.proofreader_track_changes_enhanced_fixed import ProofReaderWithTrackChangesAndCommentsFixed
from src.proofreader_synchronized import SynchronizedProofReader
# from proofreader.proofreader_track_changes_enhanced_v2 import ProofReaderWithTrackChangesAndCommentsV2


console = Console()


def load_config():
    """加载配置文件"""
    config = configparser.ConfigParser()
    config_files = ['config.ini', 'settings.ini', '.env']
    
    for config_file in config_files:
        if os.path.exists(config_file):
            config.read(config_file, encoding='utf-8')
            return config
    
    return None


def get_api_key():
    """获取API密钥"""
    # 优先从环境变量获取
    api_key = os.getenv('OPENAI_API_KEY') or os.getenv('API_KEY')
    if api_key:
        return api_key
    
    # 从配置文件获取
    config = load_config()
    if config:
        if config.has_section('ai') and config.has_option('ai', 'api_key'):
            return config.get('ai', 'api_key')
        if config.has_section('openai') and config.has_option('openai', 'api_key'):
            return config.get('openai', 'api_key')
    
    return None


@click.group()
@click.version_option(version="1.0.0")
def cli():
    """AI校对助手 - 专业的中文计算机教材校对工具"""
    pass


@cli.command()
@click.option('-i', '--input', 'input_file', required=True, help='输入Word文档路径')
@click.option('-o', '--output', 'output_file', help='输出Word文档路径')
@click.option('-m', '--mode', default='comments', type=click.Choice(['comments', 'revisions', 'track_changes', 'enhanced']), 
              help='校对模式：comments（批注模式）、revisions（修订模式）、track_changes（跟踪更改）、enhanced（跟踪更改+批注）')
@click.option('--show-chunks', is_flag=True, help='显示文档分块处理信息')
def proofread(input_file: str, output_file: str, mode: str, show_chunks: bool):
    """校对Word文档 - 支持长文档自动分块处理"""
    try:
        api_key = get_api_key()
        if not api_key:
            console.print("[red]❌ 错误：未找到API密钥。请设置环境变量OPENAI_API_KEY或在配置文件中设置。[/red]")
            return
        
        if show_chunks:
            console.print("[blue]📊 分块处理功能已启用，将显示详细的分块信息[/blue]")
        
        console.print(f"[blue]🤖 使用模型: {os.getenv('OPENAI_MODEL', 'gpt-4')}[/blue]")
        console.print(f"[blue]📏 Token限制: {os.getenv('MAX_TOKENS', '8192')}[/blue]")
        console.print("[blue]💡 如果文档超过token限制，将自动分块处理[/blue]")
        
        if mode == 'enhanced':
            console.print("[blue]🌟 使用增强模式（跟踪更改+批注）进行校对...[/blue]")
            console.print("[dim]增强模式同时提供跟踪更改和详细批注，提供最完整的校对体验[/dim]")
            proofreader = ProofReaderWithTrackChangesAndCommentsFixed(api_key)
        elif mode == 'track_changes':
            console.print("[blue]🔄 使用真正的Word跟踪更改功能进行校对...[/blue]")
            console.print("[dim]真正的Word跟踪更改功能将直接在文档中显示修改，使用Word的跟踪更改功能[/dim]")
            proofreader = ProofReaderWithTrackChanges(api_key)
        elif mode == 'revisions':
            console.print("[blue]🔄 使用修订模式进行校对...[/blue]")
            console.print("[dim]修订模式将直接在文档中显示修改，使用Word的跟踪更改功能[/dim]")
            proofreader = ProofReaderWithRevisions(api_key)
        else:
            console.print("[blue]💬 使用批注模式进行校对...[/blue]")
            console.print("[dim]批注模式将在Word审阅窗格中显示建议[/dim]")
            proofreader = ProofReader(api_key)
        
        # 根据不同模式调用相应的校对方法
        if mode == 'enhanced':
            success = proofreader.proofread_with_track_changes_and_comments(input_file, output_file)
        else:
            success = proofreader.proofread_document(input_file, output_file, mode)
        
        if success:
            console.print(f"[green]✅ 校对完成！输出文件：{output_file or input_file.replace('.docx', f'_{mode}.docx')}[/green]")
            
            if mode == 'enhanced':
                console.print("[blue]📝 使用Word打开文档，体验完整的校对功能：[/blue]")
                console.print("   - [red]删除线文本[/red] 表示需要删除的内容")
                console.print("   - [blue underline]下划线文本[/blue underline] 表示新插入的内容")
                console.print("   - 💬 详细的批注说明修改原因")
                console.print("   - ✅ 可以接受或拒绝每个修改")
                console.print("   - 🗨️ 可以回复批注进行讨论")
            elif mode == 'revisions':
                console.print("[blue]📝 使用Word打开文档，可以看到跟踪更改：[/blue]")
                console.print("   - [red]删除线文本[/red] 表示需要删除的内容")
                console.print("   - [blue underline]下划线文本[/blue underline] 表示新插入的内容")
                console.print("   - 可以在Word中接受或拒绝这些修改")
            elif mode == 'track_changes':
                console.print("[blue]📝 使用Word打开文档，可以看到跟踪更改：[/blue]")
                console.print("   - [red]删除线文本[/red] 表示需要删除的内容")
                console.print("   - [blue underline]下划线文本[/blue underline] 表示新插入的内容")
                console.print("   - 可以在Word中接受或拒绝这些修改")
            else:
                console.print("[blue]📝 使用Word打开文档，在审阅窗格中查看批注建议[/blue]")
                
            console.print("[green]🎉 校对完成！长文档分块处理功能已自动处理Token限制问题[/green]")
        else:
            console.print("[red]❌ 校对失败[/red]")
    
    except Exception as e:
        console.print(f"[red]错误：{e}[/red]")
        import traceback
        traceback.print_exc()


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


@cli.command()
@click.option('-i', '--input', 'input_file', required=True, help='输入Word文档路径')
@click.option('-o', '--output', 'output_file', help='输出Word文档路径')
def revise(input_file: str, output_file: str):
    """使用修订模式校对Word文档（快捷命令）"""
    try:
        api_key = get_api_key()
        if not api_key:
            console.print("[red]❌ 错误：未找到API密钥。请设置环境变量OPENAI_API_KEY或在配置文件中设置。[/red]")
            return
        
        from src.proofreader_revisions import ProofReaderWithRevisions
        
        proofreader = ProofReaderWithRevisions(api_key)
        
        console.print("[blue]🔄 使用修订模式进行校对...[/blue]")
        console.print("[dim]修订模式将直接在文档中显示修改，使用Word的跟踪更改功能[/dim]")
        
        success = proofreader.proofread_document_with_revisions(input_file, output_file)
        
        if success:
            console.print(f"[green]✅ 修订校对完成！输出文件：{output_file or input_file.replace('.docx', '_revised.docx')}[/green]")
            console.print("[blue]📝 使用Word打开文档，可以看到跟踪更改：[/blue]")
            console.print("   - [red]删除线文本[/red] 表示需要删除的内容")
            console.print("   - [blue underline]下划线文本[/blue underline] 表示新插入的内容")
            console.print("   - 可以在Word中接受或拒绝这些修改")
        else:
            console.print("[red]❌ 修订校对失败[/red]")
    
    except Exception as e:
        console.print(f"[red]错误：{e}[/red]")


# 文档转换能力检查已在CLI中实现


if __name__ == '__main__':
    cli() 