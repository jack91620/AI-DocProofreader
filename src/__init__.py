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

import os
import sys

from .config import Config
from .proofreader import ProofReader

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from version import __version__

__author__ = "AI Assistant"

__all__ = ["Config", "ProofReader"]
