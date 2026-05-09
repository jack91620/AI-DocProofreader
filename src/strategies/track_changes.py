"""
跟踪更改模式策略 - 使用真正的Word跟踪更改功能
"""

import os

from docx import Document
from rich.console import Console

from word_ops.track_changes import WordTrackChangesManager, enable_track_changes_in_docx

from .base import Correction, OutputStrategy


class TrackChangesStrategy(OutputStrategy):
    """跟踪更改模式输出策略"""

    def __init__(self):
        self.console = Console()

    def get_output_suffix(self) -> str:
        return "tracked"

    def apply_corrections(
        self,
        doc: Document,
        corrections: list[Correction],
        text_content: list[str],
        input_file: str,
        output_file: str,
    ) -> bool:
        try:
            track_changes_manager = WordTrackChangesManager(doc)
            change_count = 0

            for correction in corrections:
                if (
                    not correction.corrected_text
                    or correction.corrected_text == correction.original_text
                ):
                    continue

                if correction.paragraph_index < len(doc.paragraphs):
                    paragraph = doc.paragraphs[correction.paragraph_index]

                    if track_changes_manager.add_tracked_change(
                        paragraph,
                        correction.original_text,
                        correction.corrected_text,
                        correction.reason,
                    ):
                        change_count += 1
                        self.console.print(
                            f"[green]✅ 跟踪更改 {change_count}: {correction.original_text} -> {correction.corrected_text}[/green]"
                        )

            # 应用所有跟踪更改到文档
            track_changes_manager.apply_all_changes()

            # 保存临时文档
            temp_file = output_file.replace(".docx", "_temp.docx")
            doc.save(temp_file)

            # 启用Word跟踪更改并生成最终文档
            if enable_track_changes_in_docx(
                temp_file, output_file, track_changes_manager.revisions_data
            ):
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                self.console.print(
                    f"[green]✅ Word跟踪更改校对完成，已应用 {change_count} 个跟踪更改[/green]"
                )
                return True
            self.console.print("[red]❌ 启用跟踪更改失败[/red]")
            return False

        except Exception as e:
            self.console.print(f"[red]❌ 跟踪更改模式校对失败: {e}[/red]")
            return False
