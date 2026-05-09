"""
Word操作模块包 - 统一的Word文档XML操作
"""

from .comments import (
    WordCommentsManager,
    add_comments_to_docx,
    create_comments_xml,
    create_document_rels,
    update_content_types,
)
from .docx_packager import modify_docx_package
from .track_changes import WordTrackChangesManager, enable_track_changes_in_docx
from .visual_revisions import SimpleWordRevisionsManager

__all__ = [
    "SimpleWordRevisionsManager",
    "WordCommentsManager",
    "WordTrackChangesManager",
    "add_comments_to_docx",
    "create_comments_xml",
    "create_document_rels",
    "enable_track_changes_in_docx",
    "modify_docx_package",
    "update_content_types",
]
