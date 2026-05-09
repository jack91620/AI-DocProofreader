"""体例 (style) 规则集。

依据
====
- GB/T 7714-2015《信息与文献 参考文献著录规则》
- GB/T 7713.2-2022《学术论文编写规则》

覆盖范围 (MVP)
==============

1. **图表/公式编号一致性**
   全文扫描 "图 1-1" / "图1.1" / "图1·1" 等不同写法，若混用则提示统一。
   编号格式无对错，但同一书必须一致。
2. **章节标题层级标识混用**
   "第一章" 与 "第1章" 同时出现 → 建议统一。
3. **参考文献类型标识符**
   GB/T 7714 要求文献类型标识：[M] 专著、[J] 期刊、[D] 学位论文、
   [EB/OL] 电子文档等。检测出现 "[m]" / "[j]" 等小写或缺失方括号的常见错误。
"""

from __future__ import annotations

import re
from collections import Counter
from collections.abc import Iterable
from dataclasses import dataclass

from rules.base import RuleHit


# ---------------------------------------------------------------------------
# 图表/公式编号一致性
# ---------------------------------------------------------------------------
@dataclass
class _NumberingConsistency:
    """扫描全文图/表/公式的编号格式，混用则提示统一。

    匹配三种常见格式：
        图 1-1 / 图1-1   → 章节-序号
        图 1.1 / 图1.1   → 章节.序号
        图 1·1 / 图1·1   → 章节·序号（少见但有教材在用）

    检测策略：把每种格式出现次数计数，若 ≥2 种格式都出现且任一非主导
    格式占比 ≥10%，就给出"统一"提示，标在第一处非主导格式上。
    """

    rule_id: str = "style.numbering_consistency"
    category_id: str = "style"
    enabled: bool = True

    LABELS: tuple[str, ...] = ("图", "表", "公式")
    # 三种分隔符的正则  → name
    SEPARATORS: tuple[tuple[str, str], ...] = (
        (r"-", "短横线"),
        (r"\.", "句点"),
        (r"·", "间隔号"),
    )

    def scan(self, text: str) -> Iterable[RuleHit]:
        for label in self.LABELS:
            # 收集每种分隔符在该 label 下的出现位置
            buckets: dict[str, list[tuple[int, str]]] = {}
            for sep, sep_name in self.SEPARATORS:
                pattern = rf"{label}\s?(\d+){sep}(\d+)"
                for m in re.finditer(pattern, text):
                    buckets.setdefault(sep_name, []).append((m.start(), m.group(0)))

            if len(buckets) < 2:
                continue

            counts = Counter({name: len(hits) for name, hits in buckets.items()})
            dominant_name, _ = counts.most_common(1)[0]
            # 对每种非主导格式提示一次
            for sep_name, hits in buckets.items():
                if sep_name == dominant_name:
                    continue
                # 只在第一处给提示，避免刷屏
                _, sample = hits[0]
                yield RuleHit(
                    category_id=self.category_id,
                    rule_id=f"{self.rule_id}.{label}",
                    text=sample,
                    suggestion=f"全书统一为「{label}X.Y」形式（"
                    f"现混用 {dict(counts)}，建议保留主导格式）",
                    reason=(
                        f"{label}编号在全文混用了 {len(buckets)} 种分隔符 "
                        f"（{dict(counts)}）；同一书应统一编号格式"
                    ),
                    severity="low",
                    confidence=0.85,
                    standard="GB/T 7713.2-2022 §7",
                )


# ---------------------------------------------------------------------------
# 章节标题层级标识混用
# ---------------------------------------------------------------------------
@dataclass
class _ChapterNumberingConsistency:
    """同一稿内"第一章"（汉字数字）与"第1章"（阿拉伯数字）混用。"""

    rule_id: str = "style.chapter_numbering"
    category_id: str = "style"
    enabled: bool = True

    def scan(self, text: str) -> Iterable[RuleHit]:
        cn_chapters = re.findall(r"第([一二三四五六七八九十百零]+)章", text)
        ar_chapters = re.findall(r"第(\d+)章", text)
        if not cn_chapters or not ar_chapters:
            return
        # 两种都用了，给出提示（标在第一个出现的次要形式上）
        first_cn = re.search(r"第[一二三四五六七八九十百零]+章", text)
        first_ar = re.search(r"第\d+章", text)
        if not (first_cn and first_ar):
            return
        # 谁少标谁
        secondary = first_ar if len(ar_chapters) <= len(cn_chapters) else first_cn
        yield RuleHit(
            category_id=self.category_id,
            rule_id=self.rule_id,
            text=secondary.group(0),
            suggestion="全书章节统一使用同一种数字（汉字 OR 阿拉伯）",
            reason=(
                f"章节编号混用汉字 {len(cn_chapters)} 处与阿拉伯数字 "
                f"{len(ar_chapters)} 处；建议统一"
            ),
            severity="low",
            confidence=0.9,
            standard="GB/T 7713.2-2022",
        )


# ---------------------------------------------------------------------------
# 参考文献类型标识符
# ---------------------------------------------------------------------------
@dataclass
class _ReferenceTypeMarker:
    """检测参考文献中的类型标识符（[M]/[J]/[D] 等）的常见错误。

    GB/T 7714-2015 规定的类型标识符（部分常见的）:
        [M]  专著 monograph
        [J]  期刊文章 journal
        [D]  学位论文 dissertation
        [C]  论文集 conference
        [N]  报纸 newspaper
        [R]  报告 report
        [S]  标准 standard
        [P]  专利 patent
        [EB/OL] 电子文档 electronic
        [DB/OL] 在线数据库
        [CP/DK] 光盘软件
    """

    rule_id: str = "style.reference_type_marker"
    category_id: str = "style"
    enabled: bool = True

    VALID_MARKERS = frozenset(
        {"M", "J", "D", "C", "N", "R", "S", "P", "EB/OL", "DB/OL", "CP/DK", "Z", "G", "A"}
    )

    def scan(self, text: str) -> Iterable[RuleHit]:
        # 1. 小写型号：[m] / [j]
        for m in re.finditer(r"\[([a-z]{1,2}(?:/[a-z]{1,3})?)\]", text):
            lower = m.group(1)
            upper = lower.upper()
            if upper in self.VALID_MARKERS:
                yield RuleHit(
                    category_id=self.category_id,
                    rule_id=f"{self.rule_id}.lowercase",
                    text=m.group(0),
                    suggestion=f"[{upper}]",
                    reason=f'参考文献类型标识符应大写："{m.group(0)}" → "[{upper}]"',
                    severity="medium",
                    confidence=0.95,
                    standard="GB/T 7714-2015",
                )

        # 2. 全角方括号：［M］
        for m in re.finditer(r"［([A-Z]{1,2}(?:/[A-Z]{1,3})?)］", text):
            marker = m.group(1)
            if marker in self.VALID_MARKERS:
                yield RuleHit(
                    category_id=self.category_id,
                    rule_id=f"{self.rule_id}.full_width_bracket",
                    text=m.group(0),
                    suggestion=f"[{marker}]",
                    reason=f'参考文献类型标识符应使用半角方括号："{m.group(0)}" → "[{marker}]"',
                    severity="medium",
                    confidence=0.95,
                    standard="GB/T 7714-2015",
                )


def build_style_rules() -> list:
    return [
        _NumberingConsistency(),
        _ChapterNumberingConsistency(),
        _ReferenceTypeMarker(),
    ]
