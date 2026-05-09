"""Cassette 风格的 AIChecker 集成测试。

如何重录 cassette
=================

1. 设好真实 token::

       export LICLOUD_TOKEN=...   # 或 OPENAI_API_KEY / ANTHROPIC_API_KEY

2. 删 ``tests/cassettes/<test_name>.yaml`` (或全清 ``tests/cassettes/``)

3. ``uv run pytest tests/integration/test_ai_checker_recorded.py --record-mode=rewrite``

cassette 入仓后，CI/普通开发者无需 token 也能跑这些测试，HTTP 全部回放。
token/敏感 header 在录制时已被 conftest 脱敏。

一条测试对应一个 yaml；新增测试如果有新HTTP路径，需要重录。
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING

import pytest

if TYPE_CHECKING:
    pass


# pytest-recording 的 vcr marker — 自动用 conftest 里的 vcr_config
pytestmark = [pytest.mark.integration, pytest.mark.vcr]


# ---------------------------------------------------------------------------
# licloud (理想内部网关) — 默认 provider
# ---------------------------------------------------------------------------
class TestLicloudCheckerRecorded:
    """licloud + Sonnet 4.6 链路的 cassette 测试。"""

    @pytest.fixture
    def licloud_checker(self, monkeypatch: pytest.MonkeyPatch):
        """构造 licloud 后端的 AIChecker；token 用占位符（cassette已脱敏）。"""
        # 强制走 licloud
        monkeypatch.setenv("AI_PROVIDER", "licloud")
        # 没有真token时给一个占位，pytest-recording 会从 cassette 回放
        monkeypatch.setenv("LICLOUD_TOKEN", os.getenv("LICLOUD_TOKEN", "test-licloud-token"))
        monkeypatch.setenv("LICLOUD_MODEL", os.getenv("LICLOUD_MODEL", "aws-claude-sonnet-4-6"))

        # 重要：licloud 后端依赖 src.ai_checker，需要在 env 设置后才 import
        from ai_checker import AIChecker
        from config import Config

        return AIChecker(Config())

    def test_call_llm_simple_chinese_passes_through(self, licloud_checker):
        """链路最小冒烟：发一个明确无误的句子，验证返回非空字符串。"""
        result = licloud_checker._call_llm(
            system_prompt="你是中文校对助手，直接回答。",
            user_prompt="请确认这句话是否需要修改：今天天气真好。",
            max_tokens=100,
            temperature=0.1,
        )
        assert isinstance(result, str)
        assert len(result) > 0

    def test_quick_check_finds_typo(self, licloud_checker):
        """端到端：明确错别字应被识别。"""
        result = licloud_checker.check_text("计算器科学是研究算法的学科。")
        # AI返回内容受温度影响有抖动，做模糊断言
        all_text = "".join(
            issue.get("text", "") + issue.get("suggestion", "") for issue in result.issues
        ) + "".join(s.get("original", "") + s.get("suggested", "") for s in result.suggestions)
        assert "计算机" in all_text or "计算器" in all_text, (
            f"期望识别'计算器'→'计算机'错别字，实际issues={result.issues}, suggestions={result.suggestions}"
        )


# ---------------------------------------------------------------------------
# 演示：标记成 live 的测试默认跳过
# ---------------------------------------------------------------------------
class TestLiveOnly:
    """这些测试需要 --run-live 才执行（用于真打 API 录制 cassette）。"""

    @pytest.mark.live
    def test_real_licloud_round_trip(self):
        """只有 --run-live 才执行；用真 token 跑一轮，便于人工 sanity check。"""
        if not os.getenv("LICLOUD_TOKEN"):
            pytest.skip("需要 LICLOUD_TOKEN 环境变量")

        from ai_checker import AIChecker
        from config import Config

        checker = AIChecker(Config())
        out = checker._call_llm(
            system_prompt="只回答 OK",
            user_prompt="say OK",
            max_tokens=10,
            temperature=0.0,
        )
        assert isinstance(out, str) and out.strip()
