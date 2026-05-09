"""校对维度分类体系测试。"""

from __future__ import annotations

from categories import (
    CATEGORIES,
    CATEGORY_BY_ID,
    DEFAULT_CATEGORY_IDS,
    GROUP_BY_ID,
    GROUPS,
    categories_for_ui,
    group_categories,
    resolve_categories,
)


class TestCategoryConstants:
    def test_categories_have_unique_ids(self):
        ids = [c.id for c in CATEGORIES]
        assert len(ids) == len(set(ids)), "category id 必须唯一"

    def test_every_category_belongs_to_known_group(self):
        for c in CATEGORIES:
            assert c.group in GROUP_BY_ID, f"{c.id} 引用未知 group: {c.group}"

    def test_default_categories_cover_text_group(self):
        # 默认勾选项至少要包含"文字层面"全部 4 项 + 通用术语 (与改造前行为一致)
        defaults = set(DEFAULT_CATEGORY_IDS)
        assert "typo" in defaults
        assert "word_usage" in defaults
        assert "grammar" in defaults
        assert "logic" in defaults
        assert "terminology" in defaults

    def test_political_default_off(self):
        political = CATEGORY_BY_ID.get("political")
        assert political is not None
        assert political.default is False, "政治性差错必须默认关闭"
        assert political.high_risk is True

    def test_groups_have_unique_orders(self):
        orders = [g.order for g in GROUPS]
        assert len(orders) == len(set(orders))


class TestResolveCategories:
    def test_none_returns_defaults(self):
        cats = resolve_categories(None)
        ids = {c.id for c in cats}
        assert ids == set(DEFAULT_CATEGORY_IDS)

    def test_empty_list_returns_empty(self):
        assert resolve_categories([]) == []

    def test_filters_unknown_silently(self, caplog):
        cats = resolve_categories(["typo", "totally-fake", "punctuation"])
        ids = [c.id for c in cats]
        assert ids == ["typo", "punctuation"]

    def test_preserves_order(self):
        cats = resolve_categories(["punctuation", "typo", "logic"])
        ids = [c.id for c in cats]
        assert ids == ["punctuation", "typo", "logic"]


class TestGroupCategories:
    def test_grouping_orders_by_groups_definition(self):
        # 即使输入顺序乱序，输出也按 GROUPS.order
        cats = resolve_categories(["code", "typo", "punctuation", "terminology"])
        grouped = group_categories(cats)
        keys = list(grouped.keys())
        # text < norm < term < cs（根据 GROUPS.order）
        assert keys == ["text", "norm", "term", "cs"]

    def test_skips_empty_groups(self):
        cats = resolve_categories(["typo"])
        grouped = group_categories(cats)
        assert list(grouped.keys()) == ["text"]
        assert len(grouped["text"]) == 1


class TestUiHelper:
    def test_categories_for_ui_in_group_order(self):
        ui = categories_for_ui()
        # 5 组，按 order 排序
        order_seen = [g.order for g, _ in ui]
        assert order_seen == sorted(order_seen)

    def test_each_group_has_categories(self):
        ui = categories_for_ui()
        for g, cats in ui:
            assert cats, f"group {g.id} 没有 category"
