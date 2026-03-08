"""Correction转换与辅助方法单元测试"""

import os
import sys
import pytest
from unittest.mock import patch, MagicMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', 'src'))

from ai_checker import ProofreadingResult
from strategies.base import Correction


def _make_proofreader():
    """创建一个不连接真实API的ProofReader实例"""
    with patch('ai_checker.openai.OpenAI'):
        from proofreader import ProofReader
        proofreader = ProofReader("test-key")
    return proofreader


def _make_base():
    """创建一个不连接真实API的ProofReaderBase实例"""
    with patch('ai_checker.openai.OpenAI'):
        from proofreader_base import ProofReaderBase
        base = ProofReaderBase("test-key")
    return base


class TestFindTextInParagraphs:
    """_find_text_in_paragraphs 段落查找测试"""

    def test_finds_exact_match(self):
        base = _make_base()
        text_content = ["第一段文字", "第二段包含目标文本在这里", "第三段文字"]
        idx = base._find_text_in_paragraphs("目标文本", text_content)
        assert idx == 1

    def test_returns_zero_for_empty(self):
        base = _make_base()
        text_content = ["段落一", "段落二"]
        idx = base._find_text_in_paragraphs("", text_content)
        assert idx == 0

    def test_returns_zero_for_whitespace(self):
        base = _make_base()
        text_content = ["段落一", "段落二"]
        idx = base._find_text_in_paragraphs("   ", text_content)
        assert idx == 0

    def test_uses_chunk_info(self):
        base = _make_base()
        # "关键词" appears in paragraphs 1 and 3; chunk_info restricts to range [2, 4]
        text_content = ["段落零", "含有关键词", "段落二", "也含有关键词", "段落四"]
        chunk_info = {"start_paragraph": 2, "end_paragraph": 4}
        idx = base._find_text_in_paragraphs("关键词", text_content, chunk_info)
        assert idx == 3

    def test_falls_back_to_full_search(self):
        base = _make_base()
        # chunk_info range does not contain the text, but full scan does
        text_content = ["段落零", "含有目标", "段落二", "段落三"]
        chunk_info = {"start_paragraph": 2, "end_paragraph": 3}
        idx = base._find_text_in_paragraphs("目标", text_content, chunk_info)
        assert idx == 1

    def test_returns_zero_when_not_found(self):
        base = _make_base()
        text_content = ["段落一", "段落二"]
        idx = base._find_text_in_paragraphs("不存在的文本", text_content)
        assert idx == 0


class TestExtractCorrectedText:
    """_extract_corrected_text 建议文本提取测试"""

    def test_extract_with_建议改为(self):
        base = _make_base()
        result = base._extract_corrected_text("建议改为：正确文本")
        assert result == "正确文本"

    def test_extract_with_应为(self):
        base = _make_base()
        result = base._extract_corrected_text("应为'correct'")
        assert result == "correct"

    def test_extract_with_arrow(self):
        base = _make_base()
        result = base._extract_corrected_text("old -> new")
        assert result == "new"

    def test_extract_with_改为(self):
        base = _make_base()
        result = base._extract_corrected_text("改为正确")
        assert result == "正确"

    def test_returns_empty_for_no_match(self):
        base = _make_base()
        result = base._extract_corrected_text("一般描述")
        assert result == ""

    def test_returns_empty_for_none(self):
        base = _make_base()
        result = base._extract_corrected_text(None)
        assert result == ""

    def test_returns_empty_for_empty_string(self):
        base = _make_base()
        result = base._extract_corrected_text("")
        assert result == ""


class TestGenerateOutputPath:
    """_generate_output_path 输出路径生成测试"""

    def test_adds_suffix(self):
        base = _make_base()
        result = base._generate_output_path("doc.docx", "proofread")
        assert result == "doc_proofread.docx"

    def test_adds_suffix_with_path(self):
        base = _make_base()
        result = base._generate_output_path("/path/to/doc.docx", "comments")
        assert result == "/path/to/doc_comments.docx"


class TestConvertAiResult:
    """_convert_ai_result 结果转换测试"""

    def test_converts_issues(self):
        pr = _make_proofreader()
        ai_result = ProofreadingResult()
        ai_result.issues.append({
            "type": "拼写错误",
            "text": "计算器",
            "suggestion": "建议改为：计算机",
            "severity": "high",
        })

        text_content = ["这台计算器很好用", "第二段"]
        corrections = pr._convert_ai_result(ai_result, text_content)

        assert len(corrections) == 1
        c = corrections[0]
        assert isinstance(c, Correction)
        assert c.original_text == "计算器"
        assert c.corrected_text == "计算机"
        assert c.correction_type == "issue"
        assert c.severity == "high"
        assert c.paragraph_index == 0

    def test_converts_suggestions(self):
        pr = _make_proofreader()
        ai_result = ProofreadingResult()
        ai_result.suggestions.append({
            "original": "程序代码",
            "suggested": "源代码",
            "reason": "术语更准确",
        })

        text_content = ["第一段", "这段有程序代码"]
        corrections = pr._convert_ai_result(ai_result, text_content)

        assert len(corrections) == 1
        c = corrections[0]
        assert isinstance(c, Correction)
        assert c.original_text == "程序代码"
        assert c.corrected_text == "源代码"
        assert c.correction_type == "suggestion"
        assert c.severity == "low"
        assert c.paragraph_index == 1

    def test_empty_result(self):
        pr = _make_proofreader()
        ai_result = ProofreadingResult()
        corrections = pr._convert_ai_result(ai_result, ["段落一"])
        assert corrections == []

    def test_skips_suggestion_with_empty_original(self):
        pr = _make_proofreader()
        ai_result = ProofreadingResult()
        ai_result.suggestions.append({
            "original": "",
            "suggested": "something",
            "reason": "reason",
        })
        corrections = pr._convert_ai_result(ai_result, ["段落一"])
        assert corrections == []

    def test_issue_with_chunk_info(self):
        pr = _make_proofreader()
        ai_result = ProofreadingResult()
        chunk_info = {"start_paragraph": 1, "end_paragraph": 2, "chunk_index": 0}
        ai_result.issues.append({
            "type": "语法",
            "text": "错误文本",
            "suggestion": "建议改为：正确文本",
            "severity": "medium",
            "chunk_info": chunk_info,
        })

        text_content = ["第零段", "含有错误文本的段落", "第二段"]
        corrections = pr._convert_ai_result(ai_result, text_content)

        assert len(corrections) == 1
        c = corrections[0]
        assert c.paragraph_index == 1
        assert c.chunk_info == chunk_info
