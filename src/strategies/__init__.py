"""
输出策略模块包
"""

from .base import OutputStrategy
from .comments import CommentsStrategy
from .enhanced import EnhancedStrategy
from .track_changes import TrackChangesStrategy
from .visual_revisions import VisualRevisionsStrategy

__all__ = [
    "CommentsStrategy",
    "EnhancedStrategy",
    "OutputStrategy",
    "TrackChangesStrategy",
    "VisualRevisionsStrategy",
]
