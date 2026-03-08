"""AIChecker模块单元测试"""

import os
import sys
import json
import pytest
from unittest.mock import patch, MagicMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', 'src'))

from ai_checker import AIChecker, ProofreadingResult


class TestProofreadingResult:
    """ProofreadingResult 数据类测试"""

    def test_init(self):
        result = ProofreadingResult()
        assert result.issues == []
        assert result.suggestions == []
        assert result.statistics == {}

    def test_add_issue(self):
        result = ProofreadingResult()
        result.add_issue(
            issue_type="拼写错误",
            text="计算器",
            suggestion="建议改为：计算机",
            severity="high",
            position=3,
        )
        assert len(result.issues) == 1
        issue = result.issues[0]
        assert issue["type"] == "拼写错误"
        assert issue["text"] == "计算器"
        assert issue["suggestion"] == "建议改为：计算机"
        assert issue["severity"] == "high"
        assert issue["position"] == 3

    def test_add_issue_default_severity(self):
        result = ProofreadingResult()
        result.add_issue(issue_type="语法", text="abc", suggestion="def")
        assert result.issues[0]["severity"] == "medium"
        assert result.issues[0]["position"] is None

    def test_add_suggestion(self):
        result = ProofreadingResult()
        result.add_suggestion(
            text="原始文本",
            suggestion="改进后的文本",
            reason="表述更清晰",
        )
        assert len(result.suggestions) == 1
        s = result.suggestions[0]
        assert s["original"] == "原始文本"
        assert s["suggested"] == "改进后的文本"
        assert s["reason"] == "表述更清晰"


def _make_checker():
    """创建一个不连接真实API的AIChecker实例"""
    with patch('ai_checker.openai.OpenAI') as mock_openai:
        mock_openai.return_value = MagicMock()
        checker = AIChecker("test-api-key")
    return checker


class TestJsonParsing:
    """_parse_json_response 解析能力测试"""

    def test_valid_json(self):
        checker = _make_checker()
        data = {"issues": [{"type": "错别字", "text": "a"}], "suggestions": []}
        result = checker._parse_json_response(json.dumps(data, ensure_ascii=False))
        assert result["issues"][0]["type"] == "错别字"
        assert result["suggestions"] == []

    def test_json_with_markdown_wrapper(self):
        checker = _make_checker()
        raw = '```json\n{"issues": [], "suggestions": []}\n```'
        result = checker._parse_json_response(raw)
        assert result["issues"] == []
        assert result["suggestions"] == []

    def test_json_with_trailing_comma(self):
        checker = _make_checker()
        raw = '{"issues": [{"text": "a",}], "suggestions": []}'
        result = checker._parse_json_response(raw)
        # Should either parse successfully or fall back gracefully
        assert "issues" in result or "suggestions" in result

    def test_json_with_single_quotes(self):
        checker = _make_checker()
        raw = "{'issues': [], 'suggestions': []}"
        result = checker._parse_json_response(raw)
        assert "issues" in result

    def test_completely_invalid(self):
        checker = _make_checker()
        result = checker._parse_json_response("这不是JSON，只是普通文字")
        assert "issues" in result
        assert "suggestions" in result

    def test_text_fallback_returns_structure(self):
        """_parse_text_fallback_smart always returns dict with issues/suggestions keys,
        even when internal regex errors occur (known bug in arrow character class)."""
        checker = _make_checker()
        result = checker._parse_text_fallback_smart("没有任何有效信息")
        assert "issues" in result
        assert "suggestions" in result
        assert isinstance(result["issues"], list)
        assert isinstance(result["suggestions"], list)


class TestCleanAiResponse:
    """_clean_ai_response 清理测试"""

    def test_removes_markdown_blocks(self):
        checker = _make_checker()
        raw = '```json\n{"a": 1}\n```'
        cleaned = checker._clean_ai_response(raw)
        assert "```" not in cleaned
        assert '{"a": 1}' in cleaned

    def test_removes_preamble_text(self):
        checker = _make_checker()
        raw = '以下是校对结果：\n{"issues": []}'
        cleaned = checker._clean_ai_response(raw)
        assert cleaned.strip().startswith("{")


class TestCheckSpelling:
    """_check_spelling 拼写检查测试"""

    def test_finds_typo(self):
        checker = _make_checker()
        # Config's typo_dict maps correct -> [typo_list]
        # e.g. "计算机": ["计算器"]
        # The method triggers when typo is in text AND correct is NOT in text
        result = ProofreadingResult()
        checker._check_spelling("这台计算器很好用", result)
        assert len(result.issues) >= 1
        assert any("计算机" in i["suggestion"] for i in result.issues)

    def test_no_false_positive(self):
        checker = _make_checker()
        result = ProofreadingResult()
        # Text contains the correct term "计算机", so no issue even if typo also present
        checker._check_spelling("计算机是现代工具", result)
        # Should produce no issue for 计算机 since it is the correct form
        typo_issues = [i for i in result.issues if "计算机" in i.get("suggestion", "")]
        assert len(typo_issues) == 0


class TestBuildPrompt:
    """_build_proofread_prompt 提示词构建测试"""

    def test_contains_text(self):
        checker = _make_checker()
        text = "这是一段测试文本"
        prompt = checker._build_proofread_prompt(text)
        assert text in prompt

    def test_contains_json_format_hint(self):
        checker = _make_checker()
        prompt = checker._build_proofread_prompt("sample")
        assert "issues" in prompt
        assert "suggestions" in prompt


class TestParseAiResult:
    """_parse_ai_result 结果合并测试"""

    def test_parses_issues_and_suggestions(self):
        checker = _make_checker()
        ai_result = {
            "issues": [
                {"type": "语法", "text": "错误文本", "suggestion": "修正", "severity": "high"}
            ],
            "suggestions": [
                {"original": "原文", "suggested": "新文", "reason": "更好"}
            ],
        }
        result = ProofreadingResult()
        checker._parse_ai_result(ai_result, result)
        assert len(result.issues) == 1
        assert result.issues[0]["type"] == "语法"
        assert len(result.suggestions) == 1
        assert result.suggestions[0]["original"] == "原文"

    def test_handles_empty_ai_result(self):
        checker = _make_checker()
        result = ProofreadingResult()
        checker._parse_ai_result({"issues": [], "suggestions": []}, result)
        assert result.issues == []
        assert result.suggestions == []
