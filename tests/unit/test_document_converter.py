"""DocumentConverter模块单元测试"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "src"))

from document_converter import DocumentConverter


class TestDocumentConverter:
    """文档转换器测试"""

    def test_detect_docx_format(self):
        converter = DocumentConverter()
        assert converter.detect_document_format("test.docx") == "unknown"  # file doesn't exist

    def test_detect_unknown_format(self):
        converter = DocumentConverter()
        assert converter.detect_document_format("nonexistent.xyz") == "unknown"

    def test_conversion_method_detection(self):
        converter = DocumentConverter()
        assert converter.conversion_method in ["libreoffice", "pandoc", "com", "online", "none"]

    def test_can_convert_doc_returns_tuple(self):
        converter = DocumentConverter()
        result = converter.can_convert_doc()
        assert isinstance(result, tuple)
        assert len(result) == 2
        assert isinstance(result[0], bool)
        assert isinstance(result[1], str)

    def test_convert_nonexistent_file(self):
        converter = DocumentConverter()
        result = converter.convert_doc_to_docx("nonexistent.doc")
        assert result is None

    def test_convert_non_doc_file(self):
        converter = DocumentConverter()
        # Create a temp file with wrong extension
        result = converter.convert_doc_to_docx("test.txt")
        assert result is None
