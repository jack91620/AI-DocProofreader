"""JSON修复链路的快照测试。

为什么这一层值得快照？
=====================

`AIChecker._parse_json_response` 通过4层修复+文本fallback处理AI返回的脏JSON。
直接assert具体输出会让测试脆弱（修复算法迭代频繁，但形状稳定）。
syrupy快照适合：输入扰动 → 输出形状稳定不变。

如何更新快照
============

新增/修改测试后第一次跑 → 用 ``uv run pytest --snapshot-update``。
review syrupy 输出文件（``__snapshots__/*.ambr``）的 diff 后入仓。
"""

from __future__ import annotations

import pytest

from ai_checker import AIChecker
from config import Config


@pytest.fixture
def parser():
    """构造一个仅用于触发解析方法的checker；不会发任何HTTP。"""
    cfg = Config()
    cfg.ai.api_key = "test-key"
    return AIChecker(cfg)


class TestCleanAiResponse:
    def test_strips_markdown_fence(self, parser, snapshot):
        raw = '```json\n{"issues": []}\n```'
        assert parser._clean_ai_response(raw) == snapshot

    def test_strips_chatty_preamble(self, parser, snapshot):
        raw = '好的，下面是校对结果：\n\n{"issues": [{"type": "错别字"}]}'
        # _clean_ai_response 不一定剥preamble，看实现行为
        assert parser._clean_ai_response(raw) == snapshot


class TestParseJsonResponse:
    """快照固化几条关键修复路径。"""

    def test_well_formed_json(self, parser, snapshot):
        raw = '{"issues": [{"type": "错别字", "text": "操做"}], "suggestions": []}'
        assert parser._parse_json_response(raw) == snapshot

    def test_missing_comma_between_objects(self, parser, snapshot):
        # 实际AI常见错误：相邻对象之间漏逗号
        raw = '{"issues": [{"type": "A"} {"type": "B"}], "suggestions": []}'
        assert parser._parse_json_response(raw) == snapshot

    def test_smart_quotes_replaced(self, parser, snapshot):
        # AI偶尔输出中文全角引号
        raw = '{"issues": [{"type": "错别字", "text": "操做"}], "suggestions": []}'
        assert parser._parse_json_response(raw) == snapshot

    def test_unparseable_falls_back_to_text(self, parser, snapshot):
        # 完全不是JSON，应走文本fallback并返回dict形状
        raw = "原文：操做系统 → 操作系统"
        result = parser._parse_json_response(raw)
        # 形状要稳：dict with issues/suggestions keys
        assert set(result.keys()) >= {"issues", "suggestions"}
        assert isinstance(result["issues"], list)
        assert isinstance(result["suggestions"], list)
        # 整体输出快照
        assert result == snapshot

    def test_empty_string_returns_empty_shape(self, parser, snapshot):
        assert parser._parse_json_response("") == snapshot
