"""docx_lint 工具自身的单元测试。

测试策略
========

构造**故意违规**的最小 docx，验证 lint 工具能准确报错。
反向验证：构造合规 docx 应当 0 issue。
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

# 让 tests/utils 可被 import
TESTS_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(TESTS_ROOT))

from utils.docx_lint import lint_docx

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


# ---------------------------------------------------------------------------
# 帮助函数
# ---------------------------------------------------------------------------
def _make_docx(tmp_path: Path, parts: dict[str, bytes]) -> Path:
    """把 dict<part_name, content_bytes> 打包成 docx zip。"""
    p = tmp_path / "test.docx"
    with zipfile.ZipFile(p, "w") as z:
        for name, content in parts.items():
            z.writestr(name, content)
    return p


def _content_types_ok() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8"?>'
        b'<Types xmlns="' + CT_NS.encode() + b'">'
        b'<Override PartName="/word/document.xml" '
        b'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        b"</Types>"
    )


def _document_ok(extra_paragraphs: bytes = b"") -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8"?>'
        b'<w:document xmlns:w="' + W_NS.encode() + b'">'
        b"<w:body>"
        b"<w:p><w:r><w:t>hello</w:t></w:r></w:p>" + extra_paragraphs + b"</w:body></w:document>"
    )


# ---------------------------------------------------------------------------
# 正例：合规 docx 应当 0 issue
# ---------------------------------------------------------------------------
class TestCleanDocx:
    def test_minimal_clean_docx_passes(self, tmp_path: Path):
        path = _make_docx(
            tmp_path,
            {
                "[Content_Types].xml": _content_types_ok(),
                "word/document.xml": _document_ok(),
            },
        )
        assert lint_docx(path) == []


# ---------------------------------------------------------------------------
# pPr 必须第一
# ---------------------------------------------------------------------------
class TestPPrOrder:
    def test_pPr_not_first_is_caught(self, tmp_path: Path):
        # commentRangeStart 出现在 pPr 之前 → 这就是 D2 修前的 bug
        bad = b"""<w:p>
            <w:commentRangeStart w:id="1"/>
            <w:pPr><w:pStyle w:val="Title"/></w:pPr>
            <w:r><w:t>x</w:t></w:r>
        </w:p>"""
        path = _make_docx(
            tmp_path,
            {
                "[Content_Types].xml": _content_types_ok(),
                "word/document.xml": _document_ok(extra_paragraphs=bad),
            },
        )
        issues = lint_docx(path)
        rules = {i.rule for i in issues}
        assert "ooxml.pPr_must_be_first" in rules, f"未捕获 pPr 顺序违规: {issues}"

    def test_pPr_first_is_clean(self, tmp_path: Path):
        ok = b"""<w:p>
            <w:pPr><w:pStyle w:val="Title"/></w:pPr>
            <w:commentRangeStart w:id="1"/>
            <w:r><w:t>x</w:t></w:r>
            <w:commentRangeEnd w:id="1"/>
            <w:r><w:commentReference w:id="1"/></w:r>
        </w:p>"""
        comments = (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<w:comments xmlns:w="' + W_NS.encode() + b'">'
            b'<w:comment w:id="1"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:comment>'
            b"</w:comments>"
        )
        path = _make_docx(
            tmp_path,
            {
                "[Content_Types].xml": _content_types_ok(),
                "word/document.xml": _document_ok(extra_paragraphs=ok),
                "word/comments.xml": comments,
            },
        )
        issues = [i for i in lint_docx(path) if i.severity == "error"]
        assert not issues, f"合规 docx 不该报错: {issues}"


# ---------------------------------------------------------------------------
# Comment id 配对
# ---------------------------------------------------------------------------
class TestCommentIdConsistency:
    def test_range_start_end_mismatch(self, tmp_path: Path):
        bad = b"""<w:p>
            <w:commentRangeStart w:id="1"/>
            <w:r><w:t>x</w:t></w:r>
            <w:commentRangeEnd w:id="2"/>
            <w:r><w:commentReference w:id="1"/></w:r>
        </w:p>"""
        path = _make_docx(
            tmp_path,
            {
                "[Content_Types].xml": _content_types_ok(),
                "word/document.xml": _document_ok(extra_paragraphs=bad),
            },
        )
        rules = {i.rule for i in lint_docx(path)}
        assert "ooxml.comment_range_pairing" in rules

    def test_reference_outside_run(self, tmp_path: Path):
        # commentReference 直接放在 <w:p> 下而非 <w:r> 下
        bad = b"""<w:p>
            <w:commentRangeStart w:id="1"/>
            <w:r><w:t>x</w:t></w:r>
            <w:commentRangeEnd w:id="1"/>
            <w:commentReference w:id="1"/>
        </w:p>"""
        path = _make_docx(
            tmp_path,
            {
                "[Content_Types].xml": _content_types_ok(),
                "word/document.xml": _document_ok(extra_paragraphs=bad),
            },
        )
        rules = {i.rule for i in lint_docx(path)}
        assert "ooxml.commentReference_parent" in rules

    def test_undefined_comment_id(self, tmp_path: Path):
        bad = b"""<w:p>
            <w:commentRangeStart w:id="99"/>
            <w:r><w:t>x</w:t></w:r>
            <w:commentRangeEnd w:id="99"/>
            <w:r><w:commentReference w:id="99"/></w:r>
        </w:p>"""
        comments = (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<w:comments xmlns:w="' + W_NS.encode() + b'">'
            b'<w:comment w:id="1"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:comment>'
            b"</w:comments>"
        )
        path = _make_docx(
            tmp_path,
            {
                "[Content_Types].xml": _content_types_ok(),
                "word/document.xml": _document_ok(extra_paragraphs=bad),
                "word/comments.xml": comments,
            },
        )
        rules = {i.rule for i in lint_docx(path)}
        assert "ooxml.comment_id_undefined" in rules


# ---------------------------------------------------------------------------
# revision id 全局唯一
# ---------------------------------------------------------------------------
class TestRevisionIdUnique:
    """复现 D3: <w:del>/<w:ins> 的 w:id 必须文档内唯一，重复则 Word 拒收。"""

    def test_duplicate_del_id_caught(self, tmp_path: Path):
        bad = b"""<w:p>
            <w:del w:id="3" w:author="x" w:date="2024-01-01T00:00:00Z">
                <w:r><w:delText>old1</w:delText></w:r>
            </w:del>
            <w:del w:id="3" w:author="x" w:date="2024-01-01T00:00:00Z">
                <w:r><w:delText>old2</w:delText></w:r>
            </w:del>
        </w:p>"""
        path = _make_docx(
            tmp_path,
            {
                "[Content_Types].xml": _content_types_ok(),
                "word/document.xml": _document_ok(extra_paragraphs=bad),
            },
        )
        rules = {i.rule for i in lint_docx(path)}
        assert "ooxml.revision_id_unique" in rules

    def test_duplicate_ins_id_caught(self, tmp_path: Path):
        bad = b"""<w:p>
            <w:ins w:id="1003" w:author="x" w:date="2024-01-01T00:00:00Z">
                <w:r><w:t>new1</w:t></w:r>
            </w:ins>
            <w:ins w:id="1003" w:author="x" w:date="2024-01-01T00:00:00Z">
                <w:r><w:t>new2</w:t></w:r>
            </w:ins>
        </w:p>"""
        path = _make_docx(
            tmp_path,
            {
                "[Content_Types].xml": _content_types_ok(),
                "word/document.xml": _document_ok(extra_paragraphs=bad),
            },
        )
        rules = {i.rule for i in lint_docx(path)}
        assert "ooxml.revision_id_unique" in rules

    def test_unique_revision_ids_no_warn(self, tmp_path: Path):
        ok = b"""<w:p>
            <w:del w:id="1" w:author="x" w:date="2024-01-01T00:00:00Z">
                <w:r><w:delText>old1</w:delText></w:r>
            </w:del>
            <w:del w:id="2" w:author="x" w:date="2024-01-01T00:00:00Z">
                <w:r><w:delText>old2</w:delText></w:r>
            </w:del>
            <w:ins w:id="1001" w:author="x" w:date="2024-01-01T00:00:00Z">
                <w:r><w:t>new1</w:t></w:r>
            </w:ins>
        </w:p>"""
        path = _make_docx(
            tmp_path,
            {
                "[Content_Types].xml": _content_types_ok(),
                "word/document.xml": _document_ok(extra_paragraphs=ok),
            },
        )
        errors = [i for i in lint_docx(path) if i.severity == "error"]
        assert not errors, f"unique revision id 不该报错: {errors}"


# ---------------------------------------------------------------------------
# 命名空间一致性
# ---------------------------------------------------------------------------
class TestNamespaceUniformity:
    def test_override_without_namespace_prefix_caught(self, tmp_path: Path):
        # 历史 D1 bug 现场：父用 ns0 前缀绑定，子元素无前缀
        bad_ct = (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<ns0:Types xmlns:ns0="' + CT_NS.encode() + b'">'
            b'<ns0:Override PartName="/word/document.xml" ContentType="x"/>'
            b'<Override PartName="/word/comments.xml" ContentType="y"/>'  # ← 无前缀
            b"</ns0:Types>"
        )
        path = _make_docx(
            tmp_path,
            {
                "[Content_Types].xml": bad_ct,
                "word/document.xml": _document_ok(),
            },
        )
        rules = {i.rule for i in lint_docx(path)}
        assert "ooxml.child_namespace" in rules


# ---------------------------------------------------------------------------
# mc:Ignorable 引用的前缀必须声明
# ---------------------------------------------------------------------------
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


class TestMcIgnorableDeclared:
    """复现 D5: settings.xml 的 ``mc:Ignorable="w14"`` 必须有对应
    ``xmlns:w14=...`` 声明，否则 Word 拒收。"""

    def test_undeclared_prefix_caught(self, tmp_path: Path):
        # 故意在 settings.xml 里有 mc:Ignorable="w14"，但未声明 w14 前缀
        bad_settings = (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<ns0:settings '
            b'xmlns:ns0="' + W_NS.encode() + b'" '
            b'xmlns:ns1="' + MC_NS.encode() + b'" '
            b'ns1:Ignorable="w14"/>'  # ← w14 未声明
        )
        path = _make_docx(
            tmp_path,
            {
                "[Content_Types].xml": _content_types_ok(),
                "word/document.xml": _document_ok(),
                "word/settings.xml": bad_settings,
            },
        )
        rules = {i.rule for i in lint_docx(path)}
        assert "ooxml.mc_ignorable_undeclared" in rules

    def test_declared_prefix_passes(self, tmp_path: Path):
        ok_settings = (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<w:settings '
            b'xmlns:w="' + W_NS.encode() + b'" '
            b'xmlns:mc="' + MC_NS.encode() + b'" '
            b'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
            b'mc:Ignorable="w14"/>'
        )
        path = _make_docx(
            tmp_path,
            {
                "[Content_Types].xml": _content_types_ok(),
                "word/document.xml": _document_ok(),
                "word/settings.xml": ok_settings,
            },
        )
        errors = [i for i in lint_docx(path) if i.severity == "error"]
        assert not errors, f"已声明前缀不该报错: {errors}"


# ---------------------------------------------------------------------------
# CLI smoke
# ---------------------------------------------------------------------------
class TestCli:
    def test_cli_exits_nonzero_on_error(self, tmp_path: Path):
        bad = b"""<w:p>
            <w:commentRangeStart w:id="1"/>
            <w:pPr/>
        </w:p>"""
        path = _make_docx(
            tmp_path,
            {
                "[Content_Types].xml": _content_types_ok(),
                "word/document.xml": _document_ok(extra_paragraphs=bad),
            },
        )
        from utils.docx_lint import main as lint_main

        old_argv = sys.argv
        sys.argv = ["docx_lint", str(path)]
        try:
            rc = lint_main()
        finally:
            sys.argv = old_argv
        assert rc == 1, "CLI 在有 error 时应当返回 1"
