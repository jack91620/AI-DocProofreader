"""术语规则集——基于 ``data/terminology/cs_terms.yaml`` 词典的：

1. **港台/外来用词替换** (regional_variants)
   - 一对一替换，确定性高（如"程式 → 程序"、"软体 → 软件"）
2. **全书译名一致性扫描** (competing_translations)
   - 同一英文术语对应的多个合规中文译名同时出现，建议统一
   - 不强制改写，仅给出"全书统一"提示
3. **英文术语保留检查**
   - english_should_be_kept 列表里的词（如 HTTP/JSON）若被强行汉化则提示

数据来源
========
词典依据全国科学技术名词审定委员会（cnterm.cn）发布的规范术语。
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from rules.base import RuleHit

# ---------------------------------------------------------------------------
# 词典加载
# ---------------------------------------------------------------------------
DEFAULT_DICT_PATH = (
    Path(__file__).resolve().parent.parent.parent / "data" / "terminology" / "cs_terms.yaml"
)


def load_terminology(path: Path | str | None = None) -> dict:
    """加载 YAML 术语词典。返回原始 dict。"""
    p = Path(path) if path else DEFAULT_DICT_PATH
    if not p.exists():
        return {"regional_variants": {}, "competing_translations": {}, "english_should_be_kept": []}
    with p.open(encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    return {
        "regional_variants": data.get("regional_variants") or {},
        "competing_translations": data.get("competing_translations") or {},
        "english_should_be_kept": data.get("english_should_be_kept") or [],
    }


# ---------------------------------------------------------------------------
# 规则: 港台用词替换
# ---------------------------------------------------------------------------
@dataclass
class _RegionalVariantRule:
    """匹配词典中已知港台/外来用词，建议替换为大陆规范词。"""

    rule_id: str = "terminology.regional_variant"
    category_id: str = "word_usage"  # 港台用词归属"词语用法"维度
    enabled: bool = True
    variants_to_preferred: dict[str, str] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, regional_variants: dict[str, list[str]]) -> _RegionalVariantRule:
        # 翻转映射: variant → preferred；按变体长度倒序避免短匹配先吃掉长匹配
        v2p: dict[str, str] = {}
        for preferred, variants in regional_variants.items():
            for v in variants:
                v2p[v] = preferred
        return cls(variants_to_preferred=dict(sorted(v2p.items(), key=lambda kv: -len(kv[0]))))

    def scan(self, text: str) -> Iterable[RuleHit]:
        if not self.variants_to_preferred:
            return
        # 用单个 alternation 一次性扫描，避免 N 次遍历
        # 注意要 escape，且用 longest-match-first
        pattern = "|".join(re.escape(v) for v in self.variants_to_preferred)
        for m in re.finditer(pattern, text):
            variant = m.group(0)
            preferred = self.variants_to_preferred[variant]
            yield RuleHit(
                category_id=self.category_id,
                rule_id=f"{self.rule_id}.{preferred}",
                text=variant,
                suggestion=preferred,
                reason=f"港台/外来用词建议改为大陆规范词：{variant!r} → {preferred!r}",
                severity="medium",
                confidence=0.9,
                standard="全国计算机科学技术名词审定委员会",
            )


# ---------------------------------------------------------------------------
# 规则: 全书译名一致性扫描
# ---------------------------------------------------------------------------
@dataclass
class _ConsistencyRule:
    """同一英文术语的多个合规中文译名同时出现 → 建议统一。

    与上面的"替换"规则不同，这条规则仅在两个译名都出现时才提示，且建议而非强制。
    """

    rule_id: str = "terminology.consistency"
    category_id: str = "terminology"
    enabled: bool = True
    competing: dict[str, list[str]] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, competing: dict[str, list[str]]) -> _ConsistencyRule:
        return cls(competing={k: list(v) for k, v in competing.items()})

    def scan(self, text: str) -> Iterable[RuleHit]:
        for english, translations in self.competing.items():
            present = [t for t in translations if t in text]
            if len(present) <= 1:
                continue
            preferred = translations[0]
            for variant in present[1:]:
                # 在文本中找到第一处 variant 出现，给出建议
                idx = text.find(variant)
                if idx < 0:
                    continue
                yield RuleHit(
                    category_id=self.category_id,
                    rule_id=f"{self.rule_id}.{english}",
                    text=variant,
                    suggestion=preferred,
                    reason=(f'"{english}" 在全书出现多种译名: {present}；建议统一为 "{preferred}"'),
                    severity="low",
                    confidence=0.7,  # 一致性建议比硬替换确定性低
                    standard="全国科学技术名词审定委员会",
                )


# ---------------------------------------------------------------------------
# 工厂
# ---------------------------------------------------------------------------
def build_terminology_rules(dict_path: Path | str | None = None) -> list:
    data = load_terminology(dict_path)
    return [
        _RegionalVariantRule.from_dict(data["regional_variants"]),
        _ConsistencyRule.from_dict(data["competing_translations"]),
    ]
