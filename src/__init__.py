"""
AI校对助手
========

一个专为中文计算机教材设计的AI校对工具。

主要功能：
- docx文档处理
- AI智能校对
- 批注系统
- 配置管理
"""

from .proofreader import ProofReader
from .config import Config

__version__ = "1.0.0"
__author__ = "AI Assistant"

__all__ = ["ProofReader", "Config"] 