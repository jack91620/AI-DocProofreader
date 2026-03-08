"""TextSplitter模块单元测试"""

import os
import sys
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', 'src'))

from text_splitter import TokenCounter, SmartTextSplitter, TextChunk


class TestTokenCounter:
    """Token计数器测试"""

    def test_empty_text(self):
        assert TokenCounter.estimate_tokens("") == 0

    def test_chinese_text(self):
        text = "计算机科学是一门研究"
        tokens = TokenCounter.estimate_tokens(text)
        assert tokens > 0
        # 10个中文字符 * 0.6 = 6, max(6, 10//2=5) = 6
        assert tokens == 6

    def test_english_text(self):
        text = "hello world"
        tokens = TokenCounter.estimate_tokens(text)
        assert tokens > 0

    def test_mixed_text(self):
        text = "Python是一种编程语言"
        tokens = TokenCounter.estimate_tokens(text)
        assert tokens > 0

    def test_full_request_tokens_includes_overhead(self):
        text = "测试文本"
        content_tokens = TokenCounter.estimate_tokens(text)
        full_tokens = TokenCounter.estimate_full_request_tokens(text)
        assert full_tokens > content_tokens
        # Should include system prompt, user template, response format, and overhead
        expected_overhead = (
            TokenCounter.SYSTEM_PROMPT_TOKENS +
            TokenCounter.USER_PROMPT_TEMPLATE_TOKENS +
            TokenCounter.RESPONSE_FORMAT_TOKENS +
            TokenCounter.API_OVERHEAD_TOKENS
        )
        assert full_tokens == content_tokens + expected_overhead


class TestSmartTextSplitter:
    """智能文本分割器测试"""

    def test_single_chunk_for_short_text(self):
        splitter = SmartTextSplitter(max_tokens=8192)
        paragraphs = ["这是一个短文本。", "这是第二段。"]
        chunks = splitter.split_by_paragraphs(paragraphs)
        assert len(chunks) == 1

    def test_multiple_chunks_for_long_text(self):
        splitter = SmartTextSplitter(max_tokens=3000, overlap_tokens=50)
        # Create long paragraphs that will exceed token limit
        paragraphs = ["这是一段非常长的文本内容。" * 100 for _ in range(10)]
        chunks = splitter.split_by_paragraphs(paragraphs)
        assert len(chunks) > 1

    def test_chunk_fields(self):
        splitter = SmartTextSplitter(max_tokens=8192)
        paragraphs = ["第一段", "第二段"]
        chunks = splitter.split_by_paragraphs(paragraphs)
        assert len(chunks) == 1
        chunk = chunks[0]
        assert chunk.chunk_index == 0
        assert chunk.start_paragraph == 0
        assert chunk.content is not None
        assert chunk.token_count > 0

    def test_empty_paragraphs(self):
        splitter = SmartTextSplitter(max_tokens=8192)
        chunks = splitter.split_by_paragraphs([])
        assert len(chunks) == 0

    def test_get_split_summary_empty(self):
        splitter = SmartTextSplitter(max_tokens=8192)
        summary = splitter.get_split_summary([])
        assert summary["total_chunks"] == 0
        assert summary["total_tokens"] == 0

    def test_get_split_summary(self):
        splitter = SmartTextSplitter(max_tokens=8192)
        paragraphs = ["测试段落一", "测试段落二"]
        chunks = splitter.split_by_paragraphs(paragraphs)
        summary = splitter.get_split_summary(chunks)
        assert summary["total_chunks"] > 0
        assert "chunks_info" in summary

    def test_content_max_tokens_minimum(self):
        # Even with very small max_tokens, content_max_tokens should be at least 500
        splitter = SmartTextSplitter(max_tokens=100)
        assert splitter.content_max_tokens == 500
