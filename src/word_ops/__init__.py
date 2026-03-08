"""
Word操作模块包 - 统一的Word文档XML操作
"""

from .comments import WordCommentsManager, add_comments_to_docx, create_comments_xml, create_document_rels, update_content_types
from .track_changes import WordTrackChangesManager, enable_track_changes_in_docx
from .visual_revisions import SimpleWordRevisionsManager
from .docx_packager import modify_docx_package

__all__ = [
    "WordCommentsManager",
    "add_comments_to_docx",
    "create_comments_xml",
    "create_document_rels",
    "update_content_types",
    "WordTrackChangesManager",
    "enable_track_changes_in_docx",
    "SimpleWordRevisionsManager",
    "modify_docx_package",
]
