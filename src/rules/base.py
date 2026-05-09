"""规则引擎核心抽象。"""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass, field
from typing import Protocol


@dataclass
class RuleHit:
    """规则匹配命中——一条具体的"问题片段"。"""

    category_id: str  # 校对维度 id (typo / punctuation / ...)
    text: str  # 原文中匹配到的片段（用于在 Word 中定位）
    suggestion: str  # 建议替换文本
    reason: str  # 给用户看的原因
    rule_id: str  # 规则唯一 id (如 "punctuation.cn_ctx_ascii_comma")
    severity: str = "medium"  # low / medium / high
    confidence: float = 0.95  # 0~1，规则一般高置信度
    standard: str = ""  # 引用的国标（如 "GB/T 15834-2011"）

    def as_issue_dict(self) -> dict:
        """转为 ProofreadingResult.issues 可吞下的 dict 形状。"""
        ref = f"（{self.standard}）" if self.standard else ""
        return {
            "type": self.category_id,
            "text": self.text,
            "suggestion": self.suggestion,
            "severity": self.severity,
            "reason": f"{self.reason}{ref}",
            "rule_id": self.rule_id,
            "confidence": self.confidence,
        }


class Rule(Protocol):
    """规则协议。每条规则只关注一个 category_id。"""

    rule_id: str
    category_id: str
    enabled: bool

    def scan(self, text: str) -> Iterable[RuleHit]:
        """扫描文本，yield 每一处命中。"""
        ...


@dataclass
class RuleEngine:
    """规则引擎：编排多条 Rule，按 category 过滤，去重。"""

    rules: list[Rule] = field(default_factory=list)

    def register(self, rule: Rule) -> None:
        self.rules.append(rule)

    def scan(self, text: str, category_ids: set[str] | None = None) -> list[RuleHit]:
        """扫描文本，仅返回选中 category_id 的规则结果。

        category_ids=None 表示运行全部启用的规则。
        """
        hits: list[RuleHit] = []
        seen: set[tuple[str, str, str]] = set()  # 简单去重: (category, text, rule_id)

        for rule in self.rules:
            if not getattr(rule, "enabled", True):
                continue
            if category_ids is not None and rule.category_id not in category_ids:
                continue
            try:
                for hit in rule.scan(text):
                    key = (hit.category_id, hit.text, hit.rule_id)
                    if key in seen:
                        continue
                    seen.add(key)
                    hits.append(hit)
            except Exception:
                # 单条规则崩了不能拖累整轮校对
                import logging

                logging.getLogger(__name__).exception(
                    "规则 %s 扫描失败", getattr(rule, "rule_id", "?")
                )
        return hits

    def covered_categories(self) -> set[str]:
        """有规则覆盖的 category_id 集合（供 LLM prompt 提示）。"""
        return {r.category_id for r in self.rules if getattr(r, "enabled", True)}


def build_default_engine() -> RuleEngine:
    """组装默认规则集。延迟 import 避免循环依赖。"""
    from rules.numerals import build_numerals_rules
    from rules.punctuation import build_punctuation_rules
    from rules.terminology import build_terminology_rules

    engine = RuleEngine()
    for rule in (
        *build_punctuation_rules(),
        *build_numerals_rules(),
        *build_terminology_rules(),
    ):
        engine.register(rule)
    return engine
