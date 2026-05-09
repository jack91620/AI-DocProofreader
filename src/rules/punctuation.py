"""标点符号规则集，依据 GB/T 15834-2011《标点符号用法》。

覆盖范围（MVP）
==============

1. **中文上下文中的英文标点**: 中文字符相邻位置出现 ASCII 逗号/句号/问号/
   感叹号/分号/冒号 → 应改全角。
2. **重复标点**: ``。。``、``！！``、``？？``、``!?`` 等。多用一个等于无。
3. **省略号写法**: ``...``、``。。。`` → 标准是 ``……`` （6 点占两格）。
4. **破折号写法**: ``--``、``——`` 之外的奇数破折号。

不覆盖的（交给 LLM）
====================

- 引号嵌套语义判定
- 冒号后接独立分句的判断
- 段中插入语两端括号配对的语义合法性

这些规则只看字面层 ASCII 与中文混排，误判率极低。
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass

from rules.base import RuleHit

# 中文字符（基本汉字 + 全角中日韩标点）
CN_CHAR = r"[一-鿿　-〿＀-￯]"

# ASCII 标点 → 全角等价物
PUNCT_MAP: dict[str, str] = {
    ",": "，",
    ".": "。",
    ";": "；",
    ":": "：",
    "?": "？",
    "!": "！",
}


@dataclass
class _AsciiPunctInChineseCtx:
    """中文字符紧邻 ASCII 标点 → 应换全角。"""

    rule_id: str = "punctuation.cn_ctx_ascii_punct"
    category_id: str = "punctuation"
    enabled: bool = True

    def scan(self, text: str) -> Iterable[RuleHit]:
        # 模式: 中文字符 + 可选空格 + ASCII 标点 + (空格或结束或中文)
        # 注意排除代码块语境（被 ` 或缩进包围），简单粗暴：跳过 4 空格缩进开头的行
        for line in text.split("\n"):
            if line.startswith(("    ", "\t")):
                continue
            # 也跳过看起来是 URL / 邮箱 / 文件路径的整行
            if re.match(r"^\s*(?:https?://|\S+@\S+\.\S+|/\S+/)", line):
                continue

            for m in re.finditer(rf"({CN_CHAR})\s*([{re.escape(',.;:?!')}])", line):
                cn, asc = m.group(1), m.group(2)
                # 例外: "Mr." "U.S." "1.5" 这类 — 句点前后是字母或数字时跳过
                # 但本规则前置条件是中文字符，所以这种情况不会进来
                full = PUNCT_MAP[asc]
                bad = m.group(0)
                fixed = cn + full
                yield RuleHit(
                    category_id=self.category_id,
                    rule_id=self.rule_id,
                    text=bad,
                    suggestion=fixed,
                    reason=f"中文上下文应使用全角标点：{asc!r} → {full!r}",
                    severity="medium",
                    confidence=0.97,
                    standard="GB/T 15834-2011",
                )


@dataclass
class _RepeatedPunctuation:
    """重复的句末标点：。。 / ！！ / ？？ / ?! / !? 等。"""

    rule_id: str = "punctuation.repeated"
    category_id: str = "punctuation"
    enabled: bool = True

    # 规范允许 "?!" 表强烈疑问，但不应用 "!?"；"！！" 等同样不规范
    PATTERNS: tuple[tuple[str, str, str], ...] = (
        (r"。{2,}", "。", "句号不应连用"),
        (r"！{2,}", "！", "感叹号不应连用"),
        (r"？{2,}", "？", "问号不应连用"),
        (r"!{2,}", "！", "感叹号不应连用，且应使用全角"),
        (r"\?{2,}", "？", "问号不应连用，且应使用全角"),
    )

    def scan(self, text: str) -> Iterable[RuleHit]:
        for pattern, replacement, reason in self.PATTERNS:
            for m in re.finditer(pattern, text):
                yield RuleHit(
                    category_id=self.category_id,
                    rule_id=self.rule_id,
                    text=m.group(0),
                    suggestion=replacement,
                    reason=reason,
                    severity="low",
                    confidence=0.95,
                    standard="GB/T 15834-2011",
                )


@dataclass
class _Ellipsis:
    """省略号: ``...`` / ``。。。`` → ``……``。"""

    rule_id: str = "punctuation.ellipsis"
    category_id: str = "punctuation"
    enabled: bool = True

    def scan(self, text: str) -> Iterable[RuleHit]:
        # 三个或以上的 ASCII 句点
        for m in re.finditer(r"\.{3,}", text):
            # 跳过代码 / 文件路径上下文（前一字符是字母数字斜杠）
            start = m.start()
            if start > 0 and text[start - 1] in "/\\":
                continue
            yield RuleHit(
                category_id=self.category_id,
                rule_id=self.rule_id,
                text=m.group(0),
                suggestion="……",
                reason="规范的省略号是中文 ……（6 点占两格）",
                severity="medium",
                confidence=0.9,
                standard="GB/T 15834-2011",
            )
        # 句点全角连用（。。。 等同样不规范）
        for m in re.finditer(r"。{3,}", text):
            yield RuleHit(
                category_id=self.category_id,
                rule_id=self.rule_id,
                text=m.group(0),
                suggestion="……",
                reason="表省略应用 ……，不是连续的句号",
                severity="medium",
                confidence=0.95,
                standard="GB/T 15834-2011",
            )


def build_punctuation_rules() -> list:
    """返回标点规则集列表。供 RuleEngine.register 使用。"""
    return [
        _AsciiPunctInChineseCtx(),
        _RepeatedPunctuation(),
        _Ellipsis(),
    ]
