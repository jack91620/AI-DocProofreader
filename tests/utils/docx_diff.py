"""docx 文档结构性 diff 工具。

.docx 是 zip + XML，二进制对比无意义；本工具的策略：

1. 解压 zip
2. 抽取 ``word/document.xml`` 与 ``word/comments.xml`` (如有)
3. 用 lxml 把可读文本/批注/修订标记规范化（去掉 rsidR、自增ID、时间戳）
4. 输出 dict 形式的"摘要" — 段落文本、批注内容、修订插入/删除文本，
   便于 syrupy 快照或 assert 比对

不追求完整还原XML结构；只要"内容相同"的两个docx输出相同摘要。
"""

from __future__ import annotations

import re
import zipfile
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

# Word XML 命名空间
NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


@dataclass
class DocxSummary:
    """docx结构化摘要。"""

    paragraphs: list[str]  # 段落纯文本（按顺序）
    comments: list[str]  # 批注内容（按id顺序）
    insertions: list[str]  # <w:ins> 插入文本
    deletions: list[str]  # <w:del> 删除文本
    has_track_changes_setting: bool  # settings.xml 是否启用了 trackChanges

    def to_dict(self) -> dict:
        """供syrupy快照用。"""
        return {
            "paragraphs": self.paragraphs,
            "comments": self.comments,
            "insertions": self.insertions,
            "deletions": self.deletions,
            "has_track_changes_setting": self.has_track_changes_setting,
        }


def summarize_docx(path: str | Path) -> DocxSummary:
    """从 .docx 文件提取结构化摘要。"""
    path = Path(path)
    with zipfile.ZipFile(path) as z:
        document_xml = z.read("word/document.xml")
        comments_xml = _try_read(z, "word/comments.xml")
        settings_xml = _try_read(z, "word/settings.xml")

    return DocxSummary(
        paragraphs=_extract_paragraphs(document_xml),
        comments=_extract_comments(comments_xml) if comments_xml else [],
        insertions=_extract_runs_by_tag(document_xml, "ins"),
        deletions=_extract_runs_by_tag(document_xml, "del"),
        has_track_changes_setting=_has_track_changes_setting(settings_xml)
        if settings_xml
        else False,
    )


def _try_read(z: zipfile.ZipFile, name: str) -> bytes | None:
    try:
        return z.read(name)
    except KeyError:
        return None


def _extract_paragraphs(xml_bytes: bytes) -> list[str]:
    root = etree.fromstring(xml_bytes)
    paragraphs = []
    for p in root.iter(f"{{{NSMAP['w']}}}p"):
        # 跳过批注内的段落（comments.xml里的p）— 这里只看document.xml的p
        text_parts = [t.text for t in p.iter(f"{{{NSMAP['w']}}}t") if t.text]
        text = "".join(text_parts)
        if text:
            paragraphs.append(text)
    return paragraphs


def _extract_comments(xml_bytes: bytes) -> list[str]:
    root = etree.fromstring(xml_bytes)
    comments = []
    for c in root.iter(f"{{{NSMAP['w']}}}comment"):
        text_parts = [t.text for t in c.iter(f"{{{NSMAP['w']}}}t") if t.text]
        comments.append("".join(text_parts))
    return comments


def _extract_runs_by_tag(xml_bytes: bytes, tag: str) -> list[str]:
    """提取所有 <w:ins> 或 <w:del> 中的可见文本。"""
    root = etree.fromstring(xml_bytes)
    out = []
    for elem in root.iter(f"{{{NSMAP['w']}}}{tag}"):
        # delText 用于 <w:del>，t 用于 <w:ins>
        text_parts = [
            t.text
            for t in elem.iter()
            if t.tag in (f"{{{NSMAP['w']}}}t", f"{{{NSMAP['w']}}}delText") and t.text
        ]
        text = "".join(text_parts)
        if text:
            out.append(text)
    return out


def _has_track_changes_setting(xml_bytes: bytes) -> bool:
    root = etree.fromstring(xml_bytes)
    return root.find(f"{{{NSMAP['w']}}}trackChanges") is not None


def diff_summary(a: DocxSummary, b: DocxSummary) -> Iterator[str]:
    """逐字段比 a vs b，逐条 yield 差异说明（人类可读）。"""
    if a.paragraphs != b.paragraphs:
        yield f"paragraphs differ: {len(a.paragraphs)} vs {len(b.paragraphs)} blocks"
    if a.comments != b.comments:
        yield f"comments differ: {len(a.comments)} vs {len(b.comments)} entries"
    if a.insertions != b.insertions:
        yield f"insertions differ: {len(a.insertions)} vs {len(b.insertions)}"
    if a.deletions != b.deletions:
        yield f"deletions differ: {len(a.deletions)} vs {len(b.deletions)}"
    if a.has_track_changes_setting != b.has_track_changes_setting:
        yield "trackChanges setting flipped"


def normalize_for_snapshot(summary: DocxSummary) -> dict:
    """为 syrupy 快照规范化（剥noise）。

    - 去掉空字符串
    - 去掉常见噪音前缀（如 "💡 改进建议："这种）
    """
    noise_pattern = re.compile(r"^[💡🔄⚠️✅❌📝💬🌟]+\s*")

    def clean(items: list[str]) -> list[str]:
        return [noise_pattern.sub("", s) for s in items if s.strip()]

    return {
        "paragraphs": clean(summary.paragraphs),
        "comments": clean(summary.comments),
        "insertions": clean(summary.insertions),
        "deletions": clean(summary.deletions),
        "has_track_changes_setting": summary.has_track_changes_setting,
    }
