"""代码段 (code) 规则集。

针对计算机教材最常见的两类代码错误：

1. **代码块内全角符号污染**
   编辑误把 `（）；，：` 全角符号写进代码，编译器拒收。检测代码块（4 空格
   缩进或 \\`\\`\\` 围栏）内的全角 ASCII 等价符号。
2. **中英括号混用**
   中文段落中 ``Python 库 (numpy)`` 与 ``Python 库（numpy）`` 混用。
   GB/T 15834 推荐：在中文上下文中括号内若纯英文/数字，可用半角括号
   且首字母前后留空格；若括号内含中文，必须用全角括号。
   规则只标记**明显问题**：括号内包含中文却用了半角，或括号内纯英文/数字
   却用了全角。

不覆盖
======

- 代码语法本身（应交给真正的编译器/lint，而非 LLM 或正则）
- 代码风格（缩进、命名约定）—— 教材中的代码示范风格不一，强制易误判
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass

from rules.base import RuleHit

# 全角 → 半角 ASCII 映射，用于在代码块里"修复"
FULL_TO_HALF_PUNCT: dict[str, str] = {
    "（": "(",
    "）": ")",
    "，": ",",
    "；": ";",
    "：": ":",
    "！": "!",
    "？": "?",
    "“": '"',
    "”": '"',
    "‘": "'",
    "’": "'",
    "【": "[",
    "】": "]",
    "｛": "{",
    "｝": "}",
    "．": ".",
    "／": "/",
}


# ---------------------------------------------------------------------------
# 代码块识别工具
# ---------------------------------------------------------------------------
def _iter_code_blocks(text: str) -> Iterable[tuple[int, int, str]]:
    """yield (start_offset, end_offset, code_text) for each code block.

    识别两种代码块：
    - 三反引号围栏 ```...```（含语言标记的）
    - 4 空格缩进的连续行
    """
    # 1. 三反引号围栏
    for m in re.finditer(r"```[a-zA-Z]*\s*\n(.*?)```", text, flags=re.DOTALL):
        yield m.start(1), m.end(1), m.group(1)

    # 2. 4 空格缩进连续段落（简单识别：连续 ≥1 行以 4 空格或 tab 开头）
    lines = text.splitlines(keepends=True)
    pos = 0
    in_block = False
    block_start = 0
    block_lines: list[str] = []
    for line in lines:
        is_indented = line.startswith(("    ", "\t"))
        if is_indented:
            if not in_block:
                in_block = True
                block_start = pos
                block_lines = []
            block_lines.append(line)
        else:
            if in_block and block_lines:
                code = "".join(block_lines)
                yield block_start, block_start + len(code), code
            in_block = False
            block_lines = []
        pos += len(line)
    # 文末可能还在一个块里
    if in_block and block_lines:
        code = "".join(block_lines)
        yield block_start, block_start + len(code), code


# ---------------------------------------------------------------------------
# 规则：代码块全角符号污染
# ---------------------------------------------------------------------------
@dataclass
class _FullWidthInCode:
    """代码块中出现全角标点 → 应改半角，否则编译错。"""

    rule_id: str = "code.fullwidth_in_block"
    category_id: str = "code"
    enabled: bool = True

    def scan(self, text: str) -> Iterable[RuleHit]:
        full_chars = "".join(FULL_TO_HALF_PUNCT.keys())
        pattern = f"[{re.escape(full_chars)}]"
        for _start, _end, code in _iter_code_blocks(text):
            for m in re.finditer(pattern, code):
                full = m.group(0)
                half = FULL_TO_HALF_PUNCT[full]
                yield RuleHit(
                    category_id=self.category_id,
                    rule_id=self.rule_id,
                    text=full,
                    suggestion=half,
                    reason=(
                        f"代码块中混入全角字符 {full!r}，应改为半角 {half!r}（否则编译器无法识别）"
                    ),
                    severity="high",  # 代码不能跑就是 high
                    confidence=0.98,
                    standard="出版社编校规程",
                )


# ---------------------------------------------------------------------------
# 规则：中文段落中的中英括号混用
# ---------------------------------------------------------------------------
@dataclass
class _BracketMixed:
    """中文段落中 `()` 内含中文 → 应用全角；`（）` 内纯英文 → 可用半角。

    规则只标"明确错误"：括号内含中文但用了半角括号。
    （括号内纯英文却用了全角括号，规范允许两种，不强制）
    """

    rule_id: str = "code.bracket_mixed"
    category_id: str = "code"
    enabled: bool = True

    def scan(self, text: str) -> Iterable[RuleHit]:
        # 跳过代码块
        code_ranges: list[tuple[int, int]] = [(s, e) for s, e, _ in _iter_code_blocks(text)]

        def in_code_block(pos: int) -> bool:
            return any(s <= pos < e for s, e in code_ranges)

        # 半角括号包裹含中文内容 → 提示改为全角
        # 至少 1 个中文字符在括号内
        for m in re.finditer(r"\(([^()\n]{0,30})\)", text):
            inner = m.group(1)
            if not inner:
                continue
            if not any("一" <= c <= "鿿" for c in inner):
                continue
            if in_code_block(m.start()):
                continue
            # 检查前后是否中文上下文
            start = m.start()
            ctx_before = text[max(0, start - 5) : start]
            if not any("一" <= c <= "鿿" for c in ctx_before):
                continue
            yield RuleHit(
                category_id=self.category_id,
                rule_id=self.rule_id,
                text=m.group(0),
                suggestion=f"（{inner}）",
                reason="中文段落中括号内含中文时应使用全角括号「（）」",
                severity="medium",
                confidence=0.9,
                standard="GB/T 15834-2011",
            )


def build_code_rules() -> list:
    return [
        _FullWidthInCode(),
        _BracketMixed(),
    ]
