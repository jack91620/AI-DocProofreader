"""数字与单位规则集，依据 GB/T 15835-2011《出版物上数字用法》、
GB 3100/3101/3102《量和单位》。

覆盖范围（MVP）
==============

1. **数值-单位间空格**: ``5ms`` / ``128MB`` → ``5 ms`` / ``128 MB``
   （GB 3101 §6.5.6: 数字与单位之间应留一个空格，但 ``%``、``°``、``′``、``″`` 例外）
2. **范围号**: ``5~10`` / ``2024-2025`` → ``5—10`` / ``2024—2025``
   （GB/T 15834: 范围号用一字线 "—"，不是波浪号或减号）
3. **千分位**: ``1,000,000`` 在中文教材中应改为千分位空格 ``1 000 000``
   （GB/T 15835 §5.1.2）

不覆盖
======

- 阿拉伯数字 vs 汉字数字的语境选择（"几十年" vs "30 年"）需上下文，留给 LLM
- 序数词的用法（"第一" vs "第 1"）同样需上下文判断
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass

from rules.base import RuleHit

# 常见物理 / 计算机单位（数值后应加空格再写）
COMMON_UNITS = (
    # 时间
    "ns",
    "us",
    "μs",
    "ms",
    "s",
    "min",
    "h",
    "d",
    # 长度
    "nm",
    "μm",
    "mm",
    "cm",
    "dm",
    "m",
    "km",
    # 质量
    "mg",
    "g",
    "kg",
    "t",
    # 字节 / 数据
    "B",
    "KB",
    "MB",
    "GB",
    "TB",
    "PB",
    "Kb",
    "Mb",
    "Gb",
    "Tb",
    "KiB",
    "MiB",
    "GiB",
    "TiB",
    # 频率/速率
    "Hz",
    "kHz",
    "MHz",
    "GHz",
    "bps",
    "kbps",
    "Mbps",
    "Gbps",
    # 电学
    "mA",
    "A",
    "mV",
    "V",
    "kV",
    "W",
    "kW",
    "Wh",
    "kWh",
    # 计算
    "fps",
    "rpm",
    "dpi",
    "ppi",
)
# 不需要空格的单位（百分号、角度等紧贴数字写）
NO_SPACE_UNITS = ("%", "°", "′", "″", "℃", "℉", "‰")


@dataclass
class _NumberUnitSpacing:
    """数字与单位之间的空格规则。"""

    rule_id: str = "numerals.unit_spacing"
    category_id: str = "numerals"
    enabled: bool = True

    def scan(self, text: str) -> Iterable[RuleHit]:
        # 跳过可能是变量名的：连续字母数字 + 单位
        # 我们要找的是：纯数字（含小数）+ 紧贴单位
        # 排除 "1ms" 但保留 "TCP1" 这类
        units_alt = "|".join(re.escape(u) for u in sorted(COMMON_UNITS, key=len, reverse=True))
        # 前置否定回顾: 前面不能是字母（避免匹配 var1ms 这类标识符）
        pattern = rf"(?<![A-Za-z_])(\d+(?:\.\d+)?)({units_alt})(?![A-Za-z0-9_])"
        for m in re.finditer(pattern, text):
            num, unit = m.group(1), m.group(2)
            bad = m.group(0)
            fixed = f"{num} {unit}"
            yield RuleHit(
                category_id=self.category_id,
                rule_id=self.rule_id,
                text=bad,
                suggestion=fixed,
                reason=f"数值与单位之间应保留一个空格：{bad!r} → {fixed!r}",
                severity="low",
                confidence=0.93,
                standard="GB 3101-1993 §6.5.6",
            )


@dataclass
class _RangeSeparator:
    """范围号: 数字之间用 ~ / - 应改为 —（一字线）。"""

    rule_id: str = "numerals.range_separator"
    category_id: str = "numerals"
    enabled: bool = True

    def scan(self, text: str) -> Iterable[RuleHit]:
        # 数值范围用 "~" / "～"（应改为一字线）
        # 注意必须两边都是数字，避免误伤 "约~大约"（不会出现）
        for m in re.finditer(r"(\d+(?:\.\d+)?)\s*[~～]\s*(\d+(?:\.\d+)?)", text):
            a, b = m.group(1), m.group(2)
            bad = m.group(0)
            fixed = f"{a}—{b}"
            yield RuleHit(
                category_id=self.category_id,
                rule_id=self.rule_id,
                text=bad,
                suggestion=fixed,
                reason='数值范围应使用一字线 "—"，不是波浪号 "~"',
                severity="low",
                confidence=0.95,
                standard="GB/T 15834-2011 §4.13",
            )

        # ASCII 减号 "-" 在两个数字之间也常被错用作范围号（如年代）
        # 但 "-" 也用作连字符，要谨慎：只匹配 4位年-4位年 这种典型场景
        for m in re.finditer(r"(\b(?:19|20)\d{2})\s*-\s*((?:19|20)\d{2}\b)", text):
            a, b = m.group(1), m.group(2)
            bad = m.group(0)
            fixed = f"{a}—{b}"
            yield RuleHit(
                category_id=self.category_id,
                rule_id=self.rule_id,
                text=bad,
                suggestion=fixed,
                reason='年份范围应使用一字线 "—"',
                severity="low",
                confidence=0.9,
                standard="GB/T 15834-2011 §4.13",
            )


@dataclass
class _ThousandsCommaSeparator:
    """千分位: ``1,000,000`` 在中文教材中应改为千分位空格 ``1 000 000``。"""

    rule_id: str = "numerals.thousands_comma"
    category_id: str = "numerals"
    enabled: bool = True

    def scan(self, text: str) -> Iterable[RuleHit]:
        # 至少两个 ,000 才算（避免误伤 "1,2" 这类列表）
        for m in re.finditer(r"\b(\d{1,3}(?:,\d{3}){2,})\b", text):
            num = m.group(1)
            fixed = num.replace(",", " ")
            yield RuleHit(
                category_id=self.category_id,
                rule_id=self.rule_id,
                text=num,
                suggestion=fixed,
                reason="出版物中四位以上整数应每三位用空格分节，不用千分位逗号",
                severity="low",
                confidence=0.85,
                standard="GB/T 15835-2011 §5.1.2",
            )


def build_numerals_rules() -> list:
    return [
        _NumberUnitSpacing(),
        _RangeSeparator(),
        _ThousandsCommaSeparator(),
    ]
