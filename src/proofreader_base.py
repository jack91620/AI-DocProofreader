"""
校对器基类 - 提取所有校对器的公共逻辑
"""

from docx import Document
from rich.console import Console

from ai_checker import AIChecker
from config import Config
from document import DocumentProcessor


class ProofReaderBase:
    """校对器基类，包含所有校对器的公共方法"""

    def __init__(self, api_key: str = None):
        self.config = Config()
        if api_key:
            self.config.ai.api_key = api_key
        self.ai_checker = AIChecker(self.config)
        self.console = Console()
        self.document_processor = DocumentProcessor()

    def extract_text_content(self, doc: Document) -> list[str]:
        """提取文档的文本内容"""
        return [paragraph.text for paragraph in doc.paragraphs]

    def _find_text_in_paragraphs(
        self, target_text: str, text_content: list, chunk_info: dict = None
    ) -> int:
        """在段落中查找目标文本的位置 - 支持分块信息"""
        if not target_text.strip():
            return 0

        # 如果有块信息，优先在指定范围内查找
        if chunk_info:
            start_para = chunk_info.get("start_paragraph", 0)
            end_para = chunk_info.get("end_paragraph", len(text_content) - 1)
            for i in range(start_para, min(end_para + 1, len(text_content))):
                if target_text in text_content[i]:
                    return i

        # 在全文中查找
        for i, paragraph_text in enumerate(text_content):
            if target_text in paragraph_text:
                return i

        return 0

    def _extract_corrected_text(self, suggestion: str) -> str:
        """从建议中提取修正后的文本"""
        if not suggestion:
            return ""
        if "建议改为：" in suggestion:
            return suggestion.rsplit("建议改为：", maxsplit=1)[-1].strip()
        if "应为" in suggestion:
            return suggestion.rsplit("应为", maxsplit=1)[-1].strip().strip("'\"")
        if "->" in suggestion:
            return suggestion.rsplit("->", maxsplit=1)[-1].strip()
        if "改为" in suggestion:
            return suggestion.rsplit("改为", maxsplit=1)[-1].strip().strip("'\"")
        return ""

    def _generate_output_path(self, input_file: str, suffix: str) -> str:
        """生成输出文件路径"""
        return input_file.replace(".docx", f"_{suffix}.docx")
