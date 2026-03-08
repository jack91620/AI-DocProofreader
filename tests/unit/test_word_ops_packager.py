"""word_ops.docx_packager模块单元测试"""

import os
import sys
import tempfile
import zipfile

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', 'src'))

from word_ops.docx_packager import modify_docx_package


def _create_test_zip(path, files: dict):
    """创建一个包含指定文件的测试zip（模拟docx）。

    Args:
        path: zip文件路径
        files: {arcname: content_bytes} 字典
    """
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for arcname, content in files.items():
            zf.writestr(arcname, content)


class TestDocxPackager:
    """DOCX打包器测试"""

    def test_roundtrip_identity(self, tmp_path):
        """Round-trip: identity modifier preserves contents."""
        input_path = str(tmp_path / "input.docx")
        output_path = str(tmp_path / "output.docx")
        original_files = {
            "word/document.xml": "<w:document>hello</w:document>",
            "[Content_Types].xml": "<Types></Types>",
        }
        _create_test_zip(input_path, original_files)

        def identity_modifier(temp_dir):
            return True

        result = modify_docx_package(input_path, output_path, identity_modifier)

        assert result is True
        assert os.path.exists(output_path)
        with zipfile.ZipFile(output_path, 'r') as zf:
            names = set(zf.namelist())
            assert names == set(original_files.keys())
            for arcname, expected_content in original_files.items():
                assert zf.read(arcname).decode('utf-8') == expected_content

    def test_modifier_adds_file(self, tmp_path):
        """Modifier that adds a new file - verify it appears in output."""
        input_path = str(tmp_path / "input.docx")
        output_path = str(tmp_path / "output.docx")
        _create_test_zip(input_path, {"existing.txt": "original"})

        def add_file_modifier(temp_dir):
            new_file = os.path.join(temp_dir, "added.txt")
            with open(new_file, 'w', encoding='utf-8') as f:
                f.write("new content")
            return True

        result = modify_docx_package(input_path, output_path, add_file_modifier)

        assert result is True
        with zipfile.ZipFile(output_path, 'r') as zf:
            names = set(zf.namelist())
            assert "existing.txt" in names
            assert "added.txt" in names
            assert zf.read("added.txt").decode('utf-8') == "new content"

    def test_modifier_modifies_existing_file(self, tmp_path):
        """Modifier that changes an existing file - verify modification persists."""
        input_path = str(tmp_path / "input.docx")
        output_path = str(tmp_path / "output.docx")
        _create_test_zip(input_path, {"data.txt": "before"})

        def modify_modifier(temp_dir):
            target = os.path.join(temp_dir, "data.txt")
            with open(target, 'w', encoding='utf-8') as f:
                f.write("after")
            return True

        result = modify_docx_package(input_path, output_path, modify_modifier)

        assert result is True
        with zipfile.ZipFile(output_path, 'r') as zf:
            assert zf.read("data.txt").decode('utf-8') == "after"

    def test_modifier_returns_false(self, tmp_path):
        """Modifier returning False - function returns False, no output created."""
        input_path = str(tmp_path / "input.docx")
        output_path = str(tmp_path / "output.docx")
        _create_test_zip(input_path, {"file.txt": "content"})

        def failing_modifier(temp_dir):
            return False

        result = modify_docx_package(input_path, output_path, failing_modifier)

        assert result is False
        assert not os.path.exists(output_path)

    def test_invalid_input_path(self, tmp_path):
        """Invalid input path - returns False."""
        input_path = str(tmp_path / "nonexistent.docx")
        output_path = str(tmp_path / "output.docx")

        result = modify_docx_package(input_path, output_path, lambda d: True)

        assert result is False
        assert not os.path.exists(output_path)

    def test_modifier_raises_exception(self, tmp_path):
        """Modifier raising an exception - returns False."""
        input_path = str(tmp_path / "input.docx")
        output_path = str(tmp_path / "output.docx")
        _create_test_zip(input_path, {"file.txt": "content"})

        def exploding_modifier(temp_dir):
            raise RuntimeError("something went wrong")

        result = modify_docx_package(input_path, output_path, exploding_modifier)

        assert result is False
