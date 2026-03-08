"""
输出策略模块包
"""

from .base import OutputStrategy
from .comments import CommentsStrategy
from .track_changes import TrackChangesStrategy
from .enhanced import EnhancedStrategy
from .visual_revisions import VisualRevisionsStrategy

__all__ = [
    "OutputStrategy",
    "CommentsStrategy",
    "TrackChangesStrategy",
    "EnhancedStrategy",
    "VisualRevisionsStrategy",
]
