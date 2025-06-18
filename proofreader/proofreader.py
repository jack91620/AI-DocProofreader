"""
主校对器模块
"""

import os
from typing import Optional
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich.table import Table

from .config import Config
from .document import DocumentProcessor
from .ai_checker import AIChecker, ProofreadingResult


class ProofReader:
    """主校对器类"""
    
    def __init__(self, config: Optional[Config] = None):
        self.config = config or Config()
        self.config.validate()
        
        self.document_processor = DocumentProcessor()
        self.ai_checker = AIChecker(self.config)
        self.console = Console()
    
    def proofread_document(self, input_path: str, output_path: str) -> bool:
        """校对文档的主方法"""
        try:
            self.console.print(f"[green]开始校对文档：{input_path}[/green]")
            
            # 1. 加载文档
            if not self.document_processor.load_document(input_path):
                self.console.print("[red]文档加载失败[/red]")
                return False
            
            # 2. 获取文档统计信息
            stats = self.document_processor.get_statistics()
            self.console.print(f"[blue]文档统计：{stats['paragraph_count']}段落，{stats['character_count']}字符[/blue]")
            
            # 3. 分段处理文档
            segments = self.document_processor.get_text_segments()
            
            total_issues = 0
            all_results = []
            
            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                console=self.console
            ) as progress:
                task = progress.add_task("正在校对...", total=len(segments))
                
                for i, segment in enumerate(segments):
                    progress.update(task, description=f"校对第 {i+1}/{len(segments)} 段")
                    
                    # AI校对
                    result = self.ai_checker.check_text(segment)
                    all_results.append((segment, result))
                    
                    # 添加批注到文档
                    self._add_comments_to_document(segment, result)
                    
                    total_issues += len(result.issues)
                    progress.advance(task)
            
            # 4. 保存文档
            if self.document_processor.save_document(output_path):
                self.console.print(f"[green]校对完成！发现 {total_issues} 个问题[/green]")
                self.console.print(f"[green]输出文件：{output_path}[/green]")
                
                # 5. 显示校对报告
                self._show_report(all_results)
                
                return True
            else:
                self.console.print("[red]文档保存失败[/red]")
                return False
                
        except Exception as e:
            self.console.print(f"[red]校对过程中出现错误：{e}[/red]")
            return False
    
    def _add_comments_to_document(self, segment: str, result: ProofreadingResult):
        """将校对结果添加为文档批注"""
        for issue in result.issues:
            # 查找包含问题文本的段落
            para_index, para_text = self.document_processor.get_paragraph_by_text(issue["text"])
            
            if para_index >= 0:
                comment = f"{issue['type']}: {issue['suggestion']}"
                self.document_processor.add_comment(
                    para_index, 
                    issue["text"], 
                    comment,
                    self.config.comment_style.author
                )
        
        # 添加改进建议
        for suggestion in result.suggestions:
            para_index, para_text = self.document_processor.get_paragraph_by_text(suggestion["original"])
            
            if para_index >= 0:
                comment = f"建议修改: {suggestion['suggested']} (理由: {suggestion['reason']})"
                self.document_processor.add_comment(
                    para_index,
                    suggestion["original"],
                    comment,
                    self.config.comment_style.author
                )
    
    def _show_report(self, results: list):
        """显示校对报告"""
        self.console.print("\n[bold blue]校对报告[/bold blue]")
        
        # 统计各类问题
        issue_counts = {}
        severity_counts = {"high": 0, "medium": 0, "low": 0}
        
        for segment, result in results:
            for issue in result.issues:
                issue_type = issue["type"]
                severity = issue["severity"]
                
                issue_counts[issue_type] = issue_counts.get(issue_type, 0) + 1
                severity_counts[severity] = severity_counts.get(severity, 0) + 1
        
        # 创建问题类型统计表
        if issue_counts:
            table = Table(title="问题类型统计")
            table.add_column("问题类型", style="cyan")
            table.add_column("数量", style="magenta")
            
            for issue_type, count in sorted(issue_counts.items(), key=lambda x: x[1], reverse=True):
                table.add_row(issue_type, str(count))
            
            self.console.print(table)
        
        # 创建严重程度统计表
        severity_table = Table(title="问题严重程度统计")
        severity_table.add_column("严重程度", style="cyan")
        severity_table.add_column("数量", style="magenta")
        severity_table.add_column("颜色标识", style="white")
        
        colors = {"high": "[red]高[/red]", "medium": "[yellow]中[/yellow]", "low": "[green]低[/green]"}
        for severity, count in severity_counts.items():
            if count > 0:
                severity_table.add_row(severity, str(count), colors[severity])
        
        self.console.print(severity_table)
        
        # 显示详细问题列表
        self._show_detailed_issues(results)
    
    def _show_detailed_issues(self, results: list):
        """显示详细问题列表"""
        self.console.print("\n[bold blue]详细问题列表[/bold blue]")
        
        issue_num = 1
        for segment, result in results:
            if result.issues:
                self.console.print(f"\n[bold cyan]段落内容:[/bold cyan] {segment[:100]}...")
                
                for issue in result.issues:
                    severity_color = {"high": "red", "medium": "yellow", "low": "green"}
                    color = severity_color.get(issue["severity"], "white")
                    
                    self.console.print(f"[{color}]{issue_num}. {issue['type']}[/{color}]")
                    self.console.print(f"   问题文本: {issue['text']}")
                    self.console.print(f"   修改建议: {issue['suggestion']}")
                    self.console.print(f"   严重程度: {issue['severity']}")
                    
                    issue_num += 1
                    
                    if issue_num > 20:  # 限制显示数量
                        self.console.print(f"[yellow]... 还有更多问题，请查看输出文档中的批注[/yellow]")
                        return
    
    def batch_proofread(self, input_dir: str, output_dir: str) -> bool:
        """批量校对目录下的所有docx文件"""
        try:
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            docx_files = [f for f in os.listdir(input_dir) if f.endswith('.docx')]
            
            if not docx_files:
                self.console.print("[yellow]未找到docx文件[/yellow]")
                return False
            
            self.console.print(f"[green]找到 {len(docx_files)} 个文档待校对[/green]")
            
            success_count = 0
            for filename in docx_files:
                input_path = os.path.join(input_dir, filename)
                output_path = os.path.join(output_dir, f"校对_{filename}")
                
                self.console.print(f"\n[blue]处理文件：{filename}[/blue]")
                
                if self.proofread_document(input_path, output_path):
                    success_count += 1
                else:
                    self.console.print(f"[red]文件 {filename} 校对失败[/red]")
            
            self.console.print(f"\n[green]批量校对完成！成功处理 {success_count}/{len(docx_files)} 个文件[/green]")
            return success_count == len(docx_files)
            
        except Exception as e:
            self.console.print(f"[red]批量校对失败：{e}[/red]")
            return False
    
    def quick_check(self, text: str) -> ProofreadingResult:
        """快速检查文本片段"""
        return self.ai_checker.check_text(text) 