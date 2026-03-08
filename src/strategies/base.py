"""
输出策略基类
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import List, Dict, Optional


@dataclass
class Correction:
    """统一的校对修正数据类"""
    paragraph_index: int
    original_text: str
    corrected_text: str
    reason: str
    correction_type: str  # 'issue', 'suggestion', 'terminology'
    severity: str = "medium"
    chunk_info: Optional[Dict] = None


class OutputStrategy(ABC):
    """输出策略抽象基类"""

    @abstractmethod
    def apply_corrections(self, doc, corrections: List[Correction], text_content: List[str],
                          input_file: str, output_file: str) -> bool:
        """应用修正到文档"""
        ...

    @abstractmethod
    def get_output_suffix(self) -> str:
        """获取输出文件后缀"""
        ...
