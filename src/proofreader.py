"""
主校对器模块 - 统一校对器 + 策略注入
"""

import os

from docx import Document

from ai_checker import ProofreadingResult
from proofreader_base import ProofReaderBase
from strategies.base import Correction, OutputStrategy


class ProofReader(ProofReaderBase):
    """主校对器类 - 支持策略模式"""

    def __init__(self, api_key: str = None, strategy: OutputStrategy = None):
        """初始化校对器"""
        super().__init__(api_key)
        self.strategy = strategy

    def proofread_document(
        self, input_file: str, output_file: str = None, mode: str = "comments"
    ) -> bool:
        """校对文档并生成输出"""
        try:
            # 如果没有策略，使用mode参数创建默认策略
            strategy = self.strategy
            if strategy is None:
                from proofreader_factory import STRATEGY_MAP

                strategy_cls = STRATEGY_MAP.get(mode)
                if strategy_cls is None:
                    self.console.print(f"[red]❌ 未知的校对模式: {mode}[/red]")
                    return False
                strategy = strategy_cls()

            # 生成输出文件名
            if not output_file:
                output_file = self._generate_output_path(input_file, strategy.get_output_suffix())

            self.console.print(f"[green]开始校对文档：{input_file}[/green]")

            # 读取文档
            doc = Document(input_file)

            # 提取文本内容
            text_content = self.extract_text_content(doc)
            self.console.print(f"[blue]提取文本内容: {len(text_content)} 个段落[/blue]")

            # 进行AI校对
            self.console.print("[bold]开始AI校对...")
            ai_result = self.ai_checker.check_text_by_paragraphs(text_content)

            # 转换AI结果为统一的Correction列表
            corrections = self._convert_ai_result(ai_result, text_content)
            self.console.print(f"[green]✅ AI校对完成，发现 {len(corrections)} 个问题[/green]")

            if not corrections:
                self.console.print("[yellow]⚠️ 未发现需要修改的问题，直接复制原文档[/yellow]")
                import shutil

                shutil.copy2(input_file, output_file)
                return True

            # 使用策略应用修正
            return strategy.apply_corrections(
                doc, corrections, text_content, input_file, output_file
            )

        except Exception as e:
            self.console.print(f"[red]❌ 校对失败: {e}[/red]")
            return False

    def batch_proofread(self, input_dir: str, output_dir: str) -> bool:
        """批量校对目录下的所有docx文件"""
        try:
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            docx_files = [f for f in os.listdir(input_dir) if f.endswith(".docx")]

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

            self.console.print(
                f"\n[green]批量校对完成！成功处理 {success_count}/{len(docx_files)} 个文件[/green]"
            )
            return success_count == len(docx_files)

        except Exception as e:
            self.console.print(f"[red]批量校对失败：{e}[/red]")
            return False

    def quick_check(self, text: str) -> ProofreadingResult:
        """快速检查文本片段"""
        return self.ai_checker.check_text(text)

    def _convert_ai_result(
        self, ai_result: ProofreadingResult, text_content: list
    ) -> list[Correction]:
        """转换AI校对结果为统一的Correction列表"""
        corrections = []

        # 处理问题
        for issue in ai_result.issues:
            problem_text = issue.get("text", "")
            suggestion = issue.get("suggestion", "")
            issue_type = issue.get("type", "其他问题")
            severity = issue.get("severity", "medium")
            chunk_info = issue.get("chunk_info", None)

            corrected_text = self._extract_corrected_text(suggestion)
            if not corrected_text:
                corrected_text = suggestion

            paragraph_index = self._find_text_in_paragraphs(problem_text, text_content, chunk_info)

            corrections.append(
                Correction(
                    paragraph_index=paragraph_index,
                    original_text=problem_text,
                    corrected_text=corrected_text,
                    reason=f"{issue_type}: {suggestion}",
                    correction_type="issue",
                    severity=severity,
                    chunk_info=chunk_info,
                )
            )

        # 处理建议
        for suggestion in ai_result.suggestions:
            original_text = suggestion.get("original", "")
            suggested_text = suggestion.get("suggested", "")
            reason = suggestion.get("reason", "")
            chunk_info = suggestion.get("chunk_info", None)

            if original_text and suggested_text:
                paragraph_index = self._find_text_in_paragraphs(
                    original_text, text_content, chunk_info
                )

                corrections.append(
                    Correction(
                        paragraph_index=paragraph_index,
                        original_text=original_text,
                        corrected_text=suggested_text,
                        reason=reason,
                        correction_type="suggestion",
                        severity="low",
                        chunk_info=chunk_info,
                    )
                )

        return corrections
