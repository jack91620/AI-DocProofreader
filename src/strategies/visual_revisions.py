"""
视觉修订模式策略 - 使用删除线和下划线显示修改
"""

from docx import Document
from rich.console import Console

from word_ops.visual_revisions import SimpleWordRevisionsManager

from .base import Correction, OutputStrategy


class VisualRevisionsStrategy(OutputStrategy):
    """视觉修订模式输出策略"""

    def __init__(self):
        self.console = Console()

    def get_output_suffix(self) -> str:
        return "revised"

    def apply_corrections(
        self,
        doc: Document,
        corrections: list[Correction],
        text_content: list[str],
        input_file: str,
        output_file: str,
    ) -> bool:
        try:
            revisions_manager = SimpleWordRevisionsManager(doc)
            revision_count = 0

            for correction in corrections:
                if (
                    not correction.corrected_text
                    or correction.corrected_text == correction.original_text
                ):
                    continue

                if correction.paragraph_index < len(doc.paragraphs):
                    paragraph = doc.paragraphs[correction.paragraph_index]

                    if revisions_manager.add_revision(
                        paragraph,
                        correction.original_text,
                        correction.corrected_text,
                        correction.reason,
                    ):
                        revision_count += 1
                        self.console.print(
                            f"[green]✅ 修订 {revision_count}: {correction.original_text} -> {correction.corrected_text}[/green]"
                        )

            doc.save(output_file)
            self.console.print(
                f"[green]✅ 修订模式校对完成！成功应用了 {revision_count} 个修订[/green]"
            )
            return True

        except Exception as e:
            self.console.print(f"[red]❌ 修订模式校对失败: {e}[/red]")
            return False
