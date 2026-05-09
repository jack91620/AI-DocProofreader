#!/usr/bin/env python
"""
AI校对助手 - 主程序入口
"""

import configparser
import os
import sys
from pathlib import Path

import click
from rich.console import Console

# 添加src目录到Python路径
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

from src.config import Config
from src.proofreader import ProofReader
from src.proofreader_factory import create_proofreader
from version import __version__

console = Console()


def load_config():
    """加载配置文件"""
    config = configparser.ConfigParser()
    config_files = ["config.ini", "settings.ini", ".env"]

    for config_file in config_files:
        if os.path.exists(config_file):
            config.read(config_file, encoding="utf-8")
            return config

    return None


def get_api_key():
    """获取API密钥 - 支持licloud、OpenAI和Claude"""
    # licloud: 优先检查内部LLM网关Token
    api_key = os.getenv("LICLOUD_TOKEN") or os.getenv("LLM_API_TOKEN")
    if api_key:
        return api_key

    # Claude: OAuth Session Token
    api_key = os.getenv("CLAUDE_SESSION_TOKEN") or os.getenv("ANTHROPIC_API_KEY")
    if api_key:
        return api_key

    # OpenAI
    api_key = os.getenv("OPENAI_API_KEY") or os.getenv("API_KEY")
    if api_key:
        return api_key

    # 从配置文件获取
    config = load_config()
    if config:
        # licloud配置
        if config.has_section("licloud") and config.has_option("licloud", "token"):
            return config.get("licloud", "token")
        # Claude配置
        if config.has_section("claude") and config.has_option("claude", "session_token"):
            return config.get("claude", "session_token")
        if config.has_section("claude") and config.has_option("claude", "api_key"):
            return config.get("claude", "api_key")
        # OpenAI配置
        if config.has_section("ai") and config.has_option("ai", "api_key"):
            return config.get("ai", "api_key")
        if config.has_section("openai") and config.has_option("openai", "api_key"):
            return config.get("openai", "api_key")

    return None


@click.group()
@click.version_option(version=__version__)
def cli():
    """AI校对助手 - 专业的中文计算机教材校对工具"""
    pass


@cli.command()
@click.option("-i", "--input", "input_file", required=True, help="输入Word文档路径")
@click.option("-o", "--output", "output_file", help="输出Word文档路径")
@click.option(
    "-m",
    "--mode",
    default="comments",
    type=click.Choice(["comments", "revisions", "track_changes", "enhanced"]),
    help="校对模式：comments（批注模式）、revisions（修订模式）、track_changes（跟踪更改）、enhanced（跟踪更改+批注）",
)
@click.option(
    "-c",
    "--category",
    "categories",
    multiple=True,
    help="校对维度（可重复 -c 多次）。例: -c typo -c terminology。运行 list-categories 查看全部。",
)
@click.option("--show-chunks", is_flag=True, help="显示文档分块处理信息")
def proofread(
    input_file: str,
    output_file: str,
    mode: str,
    categories: tuple[str, ...],
    show_chunks: bool,
):
    """校对Word文档 - 支持长文档自动分块处理"""
    try:
        api_key = get_api_key()
        if not api_key:
            console.print(
                "[red]❌ 错误：未找到API密钥。请设置环境变量 LICLOUD_TOKEN、CLAUDE_SESSION_TOKEN 或 OPENAI_API_KEY。[/red]"
            )
            return

        if show_chunks:
            console.print("[blue]📊 分块处理功能已启用，将显示详细的分块信息[/blue]")

        # 检测提供商（优先级: licloud > claude > openai）
        provider = os.getenv("AI_PROVIDER", "").lower()
        if not provider:
            if os.getenv("LICLOUD_TOKEN") or os.getenv("LLM_API_TOKEN"):
                provider = "licloud"
            elif os.getenv("CLAUDE_SESSION_TOKEN") or os.getenv("ANTHROPIC_API_KEY"):
                provider = "claude"
            else:
                provider = "openai"

        if provider == "licloud":
            model_name = os.getenv("LICLOUD_MODEL", "aws-claude-sonnet-4-6")
            console.print(f"[blue]🤖 使用licloud网关模型: {model_name}[/blue]")
        elif provider == "claude":
            model_name = os.getenv(
                "CLAUDE_MODEL", os.getenv("ANTHROPIC_MODEL", "claude-sonnet-4-20250514")
            )
            console.print(f"[blue]🤖 使用Claude模型: {model_name}[/blue]")
        else:
            model_name = os.getenv("OPENAI_MODEL", "gpt-4")
            console.print(f"[blue]🤖 使用OpenAI模型: {model_name}[/blue]")

        console.print(f"[blue]📏 Token限制: {os.getenv('MAX_TOKENS', 'auto')}[/blue]")
        console.print("[blue]💡 如果文档超过token限制，将自动分块处理[/blue]")

        mode_labels = {
            "enhanced": (
                "[blue]🌟 使用增强模式（跟踪更改+批注）进行校对...[/blue]",
                "[dim]增强模式同时提供跟踪更改和详细批注，提供最完整的校对体验[/dim]",
            ),
            "track_changes": (
                "[blue]🔄 使用真正的Word跟踪更改功能进行校对...[/blue]",
                "[dim]真正的Word跟踪更改功能将直接在文档中显示修改[/dim]",
            ),
            "revisions": (
                "[blue]🔄 使用修订模式进行校对...[/blue]",
                "[dim]修订模式将直接在文档中显示修改，使用Word的跟踪更改功能[/dim]",
            ),
            "comments": (
                "[blue]💬 使用批注模式进行校对...[/blue]",
                "[dim]批注模式将在Word审阅窗格中显示建议[/dim]",
            ),
        }
        label, desc = mode_labels[mode]
        console.print(label)
        console.print(desc)

        cat_list = list(categories) if categories else None
        proofreader = create_proofreader(api_key, mode)
        success = proofreader.proofread_document(input_file, output_file, categories=cat_list)

        if success:
            console.print(
                f"[green]✅ 校对完成！输出文件：{output_file or input_file.replace('.docx', f'_{mode}.docx')}[/green]"
            )

            if mode in ("enhanced",):
                console.print("[blue]📝 使用Word打开文档，体验完整的校对功能：[/blue]")
                console.print("   - [red]删除线文本[/red] 表示需要删除的内容")
                console.print("   - [blue underline]下划线文本[/blue underline] 表示新插入的内容")
                console.print("   - 💬 详细的批注说明修改原因")
                console.print("   - ✅ 可以接受或拒绝每个修改")
                console.print("   - 🗨️ 可以回复批注进行讨论")
            elif mode in ("revisions", "track_changes"):
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
@click.option("--input-dir", "-i", required=True, help="输入文件夹路径")
@click.option("--output-dir", "-o", required=True, help="输出文件夹路径")
def batch(input_dir, output_dir):
    """批量校对文件夹中的所有docx文档"""
    try:
        # 检查输入目录是否存在
        if not os.path.exists(input_dir):
            console.print(f"[red]错误：输入目录不存在 - {input_dir}[/red]")
            sys.exit(1)

        api_key = get_api_key()
        if not api_key:
            console.print("[red]❌ 错误：未找到API密钥。[/red]")
            sys.exit(1)

        proofreader = ProofReader(api_key)

        if proofreader.batch_proofread(input_dir, output_dir):
            console.print(f"[green]✅ 批量校对完成！输出目录：{output_dir}[/green]")
        else:
            console.print("[red]❌ 批量校对失败[/red]")
            sys.exit(1)

    except Exception as e:
        console.print(f"[red]程序执行错误：{e}[/red]")
        sys.exit(1)


@cli.command()
@click.option("--text", "-t", help="要检查的文本内容")
@click.option("--file", "-f", help="包含文本的文件路径")
def check(text, file):
    """快速检查文本片段"""
    try:
        # 获取文本内容
        if file:
            if not os.path.exists(file):
                console.print(f"[red]错误：文件不存在 - {file}[/red]")
                sys.exit(1)

            with open(file, encoding="utf-8") as f:
                text_content = f.read()
        elif text:
            text_content = text
        else:
            console.print("[red]错误：请提供要检查的文本内容或文件路径[/red]")
            sys.exit(1)

        api_key = get_api_key()
        if not api_key:
            console.print("[red]❌ 错误：未找到API密钥。[/red]")
            sys.exit(1)

        proofreader = ProofReader(api_key)

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

        console.print(f"[cyan]AI提供商: {config.ai.provider}[/cyan]")
        console.print(f"[cyan]模型: {config.ai.model}[/cyan]")
        console.print(f"[cyan]最大Token数: {config.ai.max_tokens}[/cyan]")
        console.print(f"[cyan]温度参数: {config.ai.temperature}[/cyan]")

        console.print("[green]环境配置正常，可以开始使用校对功能！[/green]")

    except Exception as e:
        console.print(f"[red]❌ 配置检查失败: {e}[/red]")
        console.print("[yellow]请检查以下项目：[/yellow]")
        console.print(
            "1. 是否设置了 LICLOUD_TOKEN、CLAUDE_SESSION_TOKEN 或 OPENAI_API_KEY 环境变量"
        )
        console.print("2. 网络连接是否正常（licloud需要在公司内网/VPN环境下使用）")
        console.print("3. API密钥/Token是否有效")
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
        api_key = get_api_key()
        if not api_key:
            console.print("[red]❌ 错误：未找到API密钥。[/red]")
            sys.exit(1)

        proofreader = ProofReader(api_key)

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
@click.option("-i", "--input", "input_file", required=True, help="输入Word文档路径")
@click.option("-o", "--output", "output_file", help="输出Word文档路径")
def revise(input_file: str, output_file: str):
    """使用修订模式校对Word文档（快捷命令）"""
    try:
        api_key = get_api_key()
        if not api_key:
            console.print(
                "[red]❌ 错误：未找到API密钥。请设置环境变量OPENAI_API_KEY或在配置文件中设置。[/red]"
            )
            return

        proofreader = create_proofreader(api_key, "revisions")

        console.print("[blue]🔄 使用修订模式进行校对...[/blue]")
        console.print("[dim]修订模式将直接在文档中显示修改，使用Word的跟踪更改功能[/dim]")

        success = proofreader.proofread_document(input_file, output_file)

        if success:
            console.print(
                f"[green]✅ 修订校对完成！输出文件：{output_file or input_file.replace('.docx', '_revised.docx')}[/green]"
            )
            console.print("[blue]📝 使用Word打开文档，可以看到跟踪更改：[/blue]")
            console.print("   - [red]删除线文本[/red] 表示需要删除的内容")
            console.print("   - [blue underline]下划线文本[/blue underline] 表示新插入的内容")
            console.print("   - 可以在Word中接受或拒绝这些修改")
        else:
            console.print("[red]❌ 修订校对失败[/red]")

    except Exception as e:
        console.print(f"[red]错误：{e}[/red]")


@cli.command("list-categories")
def list_categories():
    """列出所有校对维度。"""
    from src.categories import CATEGORIES, GROUPS

    by_group: dict[str, list] = {}
    for c in CATEGORIES:
        by_group.setdefault(c.group, []).append(c)

    console.print("\n[bold]📋 可用校对维度（CY/T 266—2023 + 计算机教材专属）[/bold]\n")
    for g in sorted(GROUPS, key=lambda x: x.order):
        cats = by_group.get(g.id, [])
        if not cats:
            continue
        console.print(f"[cyan bold]{g.label}[/cyan bold] ({g.id}) — {g.description}")
        for c in cats:
            mark = "[green]●[/green]" if c.default else "[dim]○[/dim]"
            risk = " [yellow]⚠️ 高风险[/yellow]" if c.high_risk else ""
            console.print(f"  {mark} [bold]{c.id:<14}[/bold] {c.label}{risk}")
            console.print(f"      [dim]{c.description}[/dim]")
        console.print()
    console.print("[dim]● = 默认开启；⚠️ = 高风险（建议人审兜底，政治性默认关闭）[/dim]")
    console.print(
        "\n[dim]用法: python main.py proofread -i x.docx -c typo -c punctuation -c terminology[/dim]\n"
    )


@cli.command()
@click.option("--host", default="127.0.0.1", show_default=True, help="监听主机")
@click.option("--port", default=8000, show_default=True, type=int, help="监听端口")
@click.option("--reload", is_flag=True, help="开启代码热重载（开发模式）")
def serve(host: str, port: int, reload: bool):
    """启动 Web UI（FastAPI + HTMX），浏览器访问 http://HOST:PORT"""
    try:
        import uvicorn
    except ImportError:
        console.print("[red]❌ 缺少 web 依赖：请先 `uv sync --extra web` 或 `--extra all`[/red]")
        sys.exit(1)

    api_key = get_api_key()
    if not api_key:
        console.print("[yellow]⚠️ 未检测到 API 密钥，Web 页面可访问但上传后会失败。[/yellow]")
    console.print(f"[blue]🌐 Web UI 启动: http://{host}:{port}[/blue]")
    uvicorn.run("src.web.app:app", host=host, port=port, reload=reload)


if __name__ == "__main__":
    cli()
