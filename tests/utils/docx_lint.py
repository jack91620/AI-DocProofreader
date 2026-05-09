"""docx OOXML 业务规则 lint。

为什么不用现成工具？
====================

- Office Open XML SDK Validator 只有 .NET 版，跨平台不便
- ECMA-376 XSD 校验只查 XML 层，不查业务规则（如 pPr 必须第一）
- python-docx round-trip 太宽松（能解析 ≠ Word 接受）
- LibreOffice headless 能转 PDF 但比 Word 宽松，warning 不一定暴露

本工具硬编码我们已经踩过的 Word 拒收坑，每条规则带来源（哪次 commit 引入）。
跑得快、零依赖、CI 友好。

使用
====

    from tests.utils.docx_lint import lint_docx

    issues = lint_docx("/path/to/file.docx")
    assert not issues, f"docx 不合规: {issues}"

新增规则
========

每出一个 "Word 报无法读取" 的实例，就在这里加一条规则钉死它。
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def Q(tag: str, ns: str = W_NS) -> str:
    return f"{{{ns}}}{tag}"


@dataclass(frozen=True)
class LintIssue:
    rule: str
    severity: str  # error / warn
    location: str  # part name + 简短定位
    detail: str

    def __str__(self) -> str:
        return f"[{self.severity}/{self.rule}] {self.location}: {self.detail}"


def lint_docx(path: str | Path) -> list[LintIssue]:
    """打开 docx，跑全部业务规则。返回所有 issue 列表（空表示合规）。"""
    issues: list[LintIssue] = []
    p = Path(path)
    with zipfile.ZipFile(p) as z:
        names = set(z.namelist())

        # 各 part 解析（缺失则记 issue 后跳过）
        parts: dict[str, etree._Element] = {}
        for required in ("[Content_Types].xml", "word/document.xml"):
            if required not in names:
                issues.append(
                    LintIssue("structure.missing_part", "error", required, "必需 part 不存在")
                )
            else:
                try:
                    parts[required] = etree.fromstring(z.read(required))
                except etree.XMLSyntaxError as e:
                    issues.append(LintIssue("xml.syntax", "error", required, str(e)))

        for optional in ("word/comments.xml", "word/_rels/document.xml.rels"):
            if optional in names:
                try:
                    parts[optional] = etree.fromstring(z.read(optional))
                except etree.XMLSyntaxError as e:
                    issues.append(LintIssue("xml.syntax", "error", optional, str(e)))

    # === 业务规则 ===
    if "word/document.xml" in parts:
        issues.extend(_check_paragraph_order(parts["word/document.xml"]))
        issues.extend(
            _check_comment_id_consistency(
                parts["word/document.xml"], parts.get("word/comments.xml")
            )
        )

    if "[Content_Types].xml" in parts:
        issues.extend(_check_namespace_uniformity(parts["[Content_Types].xml"], CT_NS))

    if "word/_rels/document.xml.rels" in parts:
        issues.extend(_check_namespace_uniformity(parts["word/_rels/document.xml.rels"], RELS_NS))

    return issues


# ---------------------------------------------------------------------------
# 规则: <w:pPr> 必须是 <w:p> 的第一个子元素
# ---------------------------------------------------------------------------
def _check_paragraph_order(doc: etree._Element) -> list[LintIssue]:
    """OOXML §17.3.1.22: 若 ``<w:p>`` 有 ``<w:pPr>``，必须是第一个子元素。

    历史 bug: WordCommentsManager._add_comment_range_start_to_element
    用 ``insert(0, ...)`` 把 commentRangeStart 插到 pPr 之前，触发
    "Word 在 X 中发现无法读取的内容" — commit f2f44f3 / D2 修复。
    """
    issues: list[LintIssue] = []
    for i, p in enumerate(doc.iter(Q("p"))):
        children = list(p)
        if not children:
            continue
        # 找 pPr 出现位置
        pPr_idx = next((idx for idx, c in enumerate(children) if c.tag == Q("pPr")), None)
        if pPr_idx is None:
            continue
        if pPr_idx != 0:
            # pPr 不在第一位 → 取第一个非 pPr 元素的 tag 报告
            offender = etree.QName(children[0].tag).localname
            issues.append(
                LintIssue(
                    rule="ooxml.pPr_must_be_first",
                    severity="error",
                    location=f"document.xml:p[{i}]",
                    detail=f"<w:pPr> 不是 <w:p> 的第一个子元素（前面是 <w:{offender}>）",
                )
            )
    return issues


# ---------------------------------------------------------------------------
# 规则: 批注 id 在 document.xml 与 comments.xml 对齐
# ---------------------------------------------------------------------------
def _check_comment_id_consistency(
    doc: etree._Element, comments: etree._Element | None
) -> list[LintIssue]:
    issues: list[LintIssue] = []

    starts = {e.get(Q("id")) for e in doc.iter(Q("commentRangeStart")) if e.get(Q("id"))}
    ends = {e.get(Q("id")) for e in doc.iter(Q("commentRangeEnd")) if e.get(Q("id"))}
    refs = {e.get(Q("id")) for e in doc.iter(Q("commentReference")) if e.get(Q("id"))}

    # start / end 配对
    if starts != ends:
        diff = (starts - ends) | (ends - starts)
        issues.append(
            LintIssue(
                rule="ooxml.comment_range_pairing",
                severity="error",
                location="document.xml",
                detail=f"commentRangeStart 与 commentRangeEnd id 不匹配，差异: {diff}",
            )
        )

    # commentReference 必须在 <w:r> 内
    for ref in doc.iter(Q("commentReference")):
        parent = ref.getparent()
        if parent is None or parent.tag != Q("r"):
            ptag = etree.QName(parent.tag).localname if parent is not None else "None"
            issues.append(
                LintIssue(
                    rule="ooxml.commentReference_parent",
                    severity="error",
                    location="document.xml",
                    detail=f"<w:commentReference id={ref.get(Q('id'))}> 必须在 <w:r> 内，"
                    f"实际父元素 <w:{ptag}>",
                )
            )

    # 引用的 id 必须在 comments.xml 中存在
    if comments is not None:
        defined = {c.get(Q("id")) for c in comments.iter(Q("comment")) if c.get(Q("id"))}
        missing = refs - defined
        if missing:
            issues.append(
                LintIssue(
                    rule="ooxml.comment_id_undefined",
                    severity="error",
                    location="comments.xml",
                    detail=f"document.xml 引用的 comment id 在 comments.xml 中不存在: {sorted(missing)}",
                )
            )
        orphan = defined - refs
        if orphan:
            issues.append(
                LintIssue(
                    rule="ooxml.comment_id_orphan",
                    severity="warn",
                    location="comments.xml",
                    detail=f"comments.xml 中未被引用的 comment id: {sorted(orphan)}",
                )
            )
    elif refs:
        issues.append(
            LintIssue(
                rule="ooxml.comments_part_missing",
                severity="error",
                location="comments.xml",
                detail="document.xml 有 commentReference 但不存在 word/comments.xml",
            )
        )

    return issues


# ---------------------------------------------------------------------------
# 规则: 命名空间一致性 (Override / Relationship 必须与父节点同 ns)
# ---------------------------------------------------------------------------
def _check_namespace_uniformity(root: etree._Element, expected_ns: str) -> list[LintIssue]:
    """所有子元素必须与 root 同命名空间。

    历史 bug: ``ET.SubElement(root, "Override")`` 字面字符串名创建无命名空间
    元素，与父节点 ``<ns0:Types xmlns:ns0=...>`` 不同 ns，Word 拒收。
    commit f2f44f3 / D1 修复。
    """
    issues: list[LintIssue] = []
    root_qname = etree.QName(root.tag)
    if root_qname.namespace != expected_ns:
        issues.append(
            LintIssue(
                rule="ooxml.root_namespace",
                severity="error",
                location="root",
                detail=f"root 元素 namespace {root_qname.namespace!r} ≠ 期望 {expected_ns!r}",
            )
        )
        return issues  # root 都错了，子元素就别看了

    for child in root:
        cqname = etree.QName(child.tag)
        if cqname.namespace != expected_ns:
            issues.append(
                LintIssue(
                    rule="ooxml.child_namespace",
                    severity="error",
                    location=f"<{cqname.localname}>",
                    detail=f"namespace {cqname.namespace!r} ≠ 父节点 {expected_ns!r}（"
                    "Word 严格 schema 校验会拒收）",
                )
            )
    return issues


# ---------------------------------------------------------------------------
# CLI 接口（便于本地手工跑）
# ---------------------------------------------------------------------------
def main() -> int:
    import sys

    if len(sys.argv) < 2:
        print("Usage: python -m tests.utils.docx_lint <docx-path>", file=sys.stderr)
        return 2
    issues = lint_docx(sys.argv[1])
    if not issues:
        print("✅ docx OOXML lint 全部通过")
        return 0
    print(f"❌ 发现 {len(issues)} 处问题:")
    for it in issues:
        print(f"  {it}")
    return 1 if any(it.severity == "error" for it in issues) else 0


if __name__ == "__main__":
    raise SystemExit(main())
