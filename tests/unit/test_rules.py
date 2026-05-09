"""规则引擎单元测试。"""

from __future__ import annotations

from rules.base import Rule, RuleEngine, RuleHit, build_default_engine
from rules.numerals import build_numerals_rules
from rules.punctuation import build_punctuation_rules
from rules.terminology import build_terminology_rules


# ===========================================================================
# 标点规则
# ===========================================================================
class TestPunctuationRules:
    def setup_method(self):
        self.engine = RuleEngine(rules=build_punctuation_rules())

    def test_chinese_context_ascii_comma(self):
        hits = self.engine.scan("你好,世界")
        assert any(h.rule_id.startswith("punctuation.cn_ctx") for h in hits)
        assert any("，" in h.suggestion for h in hits)

    def test_chinese_context_ascii_period(self):
        hits = self.engine.scan("结束了.")
        assert any("。" in h.suggestion for h in hits)

    def test_chinese_context_ascii_question(self):
        hits = self.engine.scan("好吗?")
        assert any("？" in h.suggestion for h in hits)

    def test_skips_code_block_indented_lines(self):
        # 4-空格缩进开头的行视为代码，不报英文标点
        hits = self.engine.scan("    if x == 1:\n        return True")
        # 不应有"中文上下文 ASCII 标点"的命中
        cn_ctx_hits = [h for h in hits if h.rule_id.startswith("punctuation.cn_ctx")]
        assert not cn_ctx_hits

    def test_does_not_flag_pure_english(self):
        hits = self.engine.scan("Hello, world!")
        # 没有中文字符，不该触发 cn_ctx 规则
        cn_ctx_hits = [h for h in hits if h.rule_id.startswith("punctuation.cn_ctx")]
        assert not cn_ctx_hits

    def test_repeated_question_marks(self):
        hits = self.engine.scan("真的吗？？？")
        repeated = [h for h in hits if h.rule_id == "punctuation.repeated"]
        assert repeated, "应识别重复问号"

    def test_ascii_ellipsis_to_chinese(self):
        hits = self.engine.scan("等等...")
        ellipsis = [h for h in hits if h.rule_id == "punctuation.ellipsis"]
        assert ellipsis
        assert ellipsis[0].suggestion == "……"


# ===========================================================================
# 数字单位规则
# ===========================================================================
class TestNumeralsRules:
    def setup_method(self):
        self.engine = RuleEngine(rules=build_numerals_rules())

    def test_unit_spacing_ms(self):
        hits = self.engine.scan("延迟 5ms 完成")
        sp = [h for h in hits if h.rule_id == "numerals.unit_spacing"]
        assert sp, f"未识别 5ms 间距问题: {hits}"
        assert sp[0].suggestion == "5 ms"

    def test_unit_spacing_mb(self):
        hits = self.engine.scan("文件 128MB")
        sp = [h for h in hits if h.rule_id == "numerals.unit_spacing"]
        assert sp and sp[0].suggestion == "128 MB"

    def test_unit_spacing_doesnt_match_identifier(self):
        # var1ms 不该被识别（前面是字母）
        hits = self.engine.scan("var1ms = 0")
        assert not any(h.rule_id == "numerals.unit_spacing" for h in hits)

    def test_range_separator_tilde(self):
        hits = self.engine.scan("范围是 5~10 之间")
        rs = [h for h in hits if h.rule_id == "numerals.range_separator"]
        assert rs
        assert rs[0].suggestion == "5—10"

    def test_range_separator_year(self):
        hits = self.engine.scan("2023-2024 年")
        rs = [h for h in hits if h.rule_id == "numerals.range_separator"]
        assert rs
        assert "2023—2024" in rs[0].suggestion

    def test_thousands_comma(self):
        hits = self.engine.scan("一共 1,234,567 条记录")
        tc = [h for h in hits if h.rule_id == "numerals.thousands_comma"]
        assert tc
        assert tc[0].suggestion == "1 234 567"

    def test_thousands_comma_skip_short(self):
        # 只有一个 ,000 不算（避免误伤 "1,2"）
        hits = self.engine.scan("文件 1,000 字节")
        tc = [h for h in hits if h.rule_id == "numerals.thousands_comma"]
        assert not tc


# ===========================================================================
# 术语规则
# ===========================================================================
class TestTerminologyRules:
    def setup_method(self):
        self.engine = RuleEngine(rules=build_terminology_rules())

    def test_regional_variant_basic(self):
        hits = self.engine.scan("软体工程师写程式")
        # 应同时识别 软体 → 软件 / 程式 → 程序
        suggestions = {(h.text, h.suggestion) for h in hits}
        assert ("软体", "软件") in suggestions or ("程式", "程序") in suggestions

    def test_regional_variant_long_match_wins(self):
        # "操做系统" 应整体匹配而非"操做"
        hits = self.engine.scan("操做系统是核心")
        assert any(h.text == "操做系统" and h.suggestion == "操作系统" for h in hits)

    def test_consistency_robustness(self):
        text = "本书使用鲁棒性概念，下文用了健壮性来描述同一性质。"
        hits = self.engine.scan(text)
        cons = [h for h in hits if h.rule_id.startswith("terminology.consistency")]
        assert cons
        # 推荐译名（第一个）是"鲁棒性"
        assert any("鲁棒性" in h.reason for h in cons)

    def test_consistency_only_when_multiple_present(self):
        # 仅出现一种译名时不应报一致性
        hits = self.engine.scan("性能是核心指标。")
        cons = [h for h in hits if h.rule_id.startswith("terminology.consistency")]
        assert not cons


# ===========================================================================
# 引擎组合
# ===========================================================================
class TestRuleEngineComposition:
    def test_default_engine_loads(self):
        engine = build_default_engine()
        assert len(engine.rules) > 0
        assert {
            "punctuation",
            "numerals",
            "terminology",
            "word_usage",
        } <= engine.covered_categories()

    def test_filter_by_category(self):
        engine = build_default_engine()
        # 只跑 punctuation 维度
        text = "你好,世界。程式 5ms"
        hits = engine.scan(text, category_ids={"punctuation"})
        # 应只有 punctuation 命中
        assert hits
        assert all(h.category_id == "punctuation" for h in hits)

    def test_dedup_by_category_text_rule(self):
        engine = RuleEngine()

        class _Dup(Rule):
            rule_id = "test.dup"
            category_id = "punctuation"
            enabled = True

            def scan(self, text):
                yield RuleHit(
                    category_id="punctuation",
                    text="，",
                    suggestion="，",
                    reason="x",
                    rule_id="test.dup",
                )
                yield RuleHit(
                    category_id="punctuation",
                    text="，",
                    suggestion="，",
                    reason="x",
                    rule_id="test.dup",
                )

        engine.register(_Dup())
        hits = engine.scan("，")
        assert len(hits) == 1, "相同 (category, text, rule_id) 应去重"

    def test_disabled_rule_skipped(self):
        engine = RuleEngine(rules=build_punctuation_rules())
        for r in engine.rules:
            r.enabled = False
        assert engine.scan("你好,世界") == []

    def test_rule_exception_does_not_break_engine(self):
        """单条规则崩了不能影响其它规则。"""
        engine = RuleEngine()

        class _Bad:
            rule_id = "test.bad"
            category_id = "punctuation"
            enabled = True

            def scan(self, text):
                raise RuntimeError("intentional")

        class _Good:
            rule_id = "test.good"
            category_id = "punctuation"
            enabled = True

            def scan(self, text):
                yield RuleHit(
                    category_id="punctuation",
                    text=",",
                    suggestion="，",
                    reason="x",
                    rule_id="test.good",
                )

        engine.register(_Bad())
        engine.register(_Good())
        # 不应抛
        hits = engine.scan("你好,世界")
        assert any(h.rule_id == "test.good" for h in hits)


# ===========================================================================
# RuleHit 序列化
# ===========================================================================
class TestRuleHit:
    def test_as_issue_dict(self):
        hit = RuleHit(
            category_id="punctuation",
            text=",",
            suggestion="，",
            reason="ASCII 标点",
            rule_id="x",
            standard="GB/T 15834-2011",
        )
        d = hit.as_issue_dict()
        assert d["type"] == "punctuation"
        assert d["text"] == ","
        assert d["suggestion"] == "，"
        assert "GB/T 15834" in d["reason"]
