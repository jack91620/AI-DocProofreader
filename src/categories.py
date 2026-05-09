"""校对维度分类体系。

依据
====
- 出版业行业标准 CY/T 266—2023《图书编校质量差错判定和计算方法》
- 《图书编校质量差错认定细则》(2005/2018) 的 9 大类差错口径
- GB/T 15834《标点符号用法》、GB/T 15835《数字用法》、GB 3100/3101/3102《量和单位》
- 全国科学技术名词审定委员会（术语在线 cnterm.cn）

设计原则
========
- **分类对齐质检口径**：12 维度 = 9 类标准差错 + 3 个计算机教材专属（代码/公式/中英术语）
- **组级 LLM 调用**：12 维度按 5 组打包，避免 N 倍 token 消耗
- **默认行为兼容**：未传 categories 时只跑"文字层"组，等价于改造前默认行为
- **政治性单独处理**：默认禁用，红线问题不让 LLM 独立判定
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class Category:
    """单个校对维度的元信息。"""

    id: str  # 英文 snake_case
    label: str  # 中文标签
    group: str  # 所属分组 id
    description: str  # 一句话说明（用于 UI tooltip）
    prompt_focus: list[str] = field(default_factory=list)  # 写入 system prompt 的检查要点
    default: bool = False  # 是否默认勾选
    high_risk: bool = False  # 高风险维度（需要警示用户）


@dataclass(frozen=True)
class CategoryGroup:
    """分组元信息。"""

    id: str
    label: str
    description: str
    order: int


# ---------------------------------------------------------------------------
# 5 个分组（按 LLM 调用粒度）
# ---------------------------------------------------------------------------
GROUPS: list[CategoryGroup] = [
    CategoryGroup(
        id="text",
        label="文字层面",
        description="错别字、词语、语法、表达逻辑",
        order=1,
    ),
    CategoryGroup(
        id="norm",
        label="规范层面",
        description="标点、数字、单位、体例（建议规则引擎兜底）",
        order=2,
    ),
    CategoryGroup(
        id="term",
        label="术语层面",
        description="专业术语统一与中英对照",
        order=3,
    ),
    CategoryGroup(
        id="cs",
        label="计算机教材专属",
        description="代码、公式、符号",
        order=4,
    ),
    CategoryGroup(
        id="content",
        label="内容层面",
        description="知识性差错（高风险），政治性默认关闭",
        order=5,
    ),
]
GROUP_BY_ID: dict[str, CategoryGroup] = {g.id: g for g in GROUPS}


# ---------------------------------------------------------------------------
# 12 个校对维度
# ---------------------------------------------------------------------------
CATEGORIES: list[Category] = [
    # --- 文字层面 (默认全选，与改造前行为一致) ---
    Category(
        id="typo",
        label="文字差错",
        group="text",
        description="错别字、多字、漏字、颠倒字、异体字、繁简、拼音、外文拼写",
        prompt_focus=[
            "错别字（同音字、形近字误用）",
            "多字、漏字、颠倒字",
            "繁简混用、异体字、旧字形",
            "汉语拼音和外文拼写错误",
        ],
        default=True,
    ),
    Category(
        id="word_usage",
        label="词语用法",
        group="text",
        description="异形词、词义误用、成语误用、生造词、搭配不当",
        prompt_focus=[
            "异形词（依《第一批异形词整理表》取规范形）",
            "词义误用、近义词混淆",
            "成语误用、生造词、搭配不当",
            '港台/外来用语规范化（如"程式→程序""软体→软件"）',
        ],
        default=True,
    ),
    Category(
        id="grammar",
        label="语法表达",
        group="text",
        description="成分残缺、搭配不当、语序混乱、关联词、指代不明",
        prompt_focus=[
            "成分残缺（主语、谓语、宾语缺失）",
            "搭配不当、语序混乱",
            "关联词误用、指代不明、句子杂糅",
        ],
        default=True,
    ),
    Category(
        id="logic",
        label="逻辑表达",
        group="text",
        description="表达冗余、歧义、前后矛盾、表达清晰度",
        prompt_focus=[
            "表达冗余、口语化、不必要的修饰",
            "歧义、表意不清",
            "前后矛盾或自相打架",
            "上下文人名/数字/术语不一致",
        ],
        default=True,
    ),
    # --- 规范层面 (默认不开，需要时勾选) ---
    Category(
        id="punctuation",
        label="标点符号",
        group="norm",
        description="GB/T 15834-2011；中文标点规范、引号嵌套、半角全角",
        prompt_focus=[
            "中文应使用全角标点（。，；：？！等）",
            "顿号、书名号、引号嵌套规则（如外双内单 \"…'…'…\"）",
            "标点位置（行首禁则、句末标点不出现在行首）",
            "破折号占两字、省略号占两字、连接号「-」「—」「～」选用",
            "依据 GB/T 15834-2011《标点符号用法》",
        ],
    ),
    Category(
        id="numerals",
        label="数字与单位",
        group="norm",
        description="GB/T 15835、GB 3100/3101/3102；阿拉伯/汉字、千分位、单位",
        prompt_focus=[
            "公历世纪/年代/年/月/日用阿拉伯数字（如 2024 年）",
            '概数、惯用语用汉字数字（如"几十年""二三十人"）',
            "数值范围用「—」（连接号）而非「~」「-」",
            "三位以上整数加千分位空格而非逗号（GB/T 15835）",
            "量和单位用国际标准符号（kg、MB、ms），数值与单位间保留空格",
            "依据 GB/T 15835-2011《出版物上数字用法》、GB 3100 系列",
        ],
    ),
    Category(
        id="style",
        label="体例统一",
        group="norm",
        description="章节标题层级、图表/公式编号、参考文献格式",
        prompt_focus=[
            "章节标题层级（一级标题、二级标题）格式与字体一致",
            "图表/公式编号统一（如全书统一用「图 1-1」或「图1.1」，不混用）",
            "图题、表题位置（图下题、表上题）",
            "参考文献依 GB/T 7714-2015 著录（[M]/[J]/[D]/[EB/OL]）",
            "目录与正文标题完全一致",
        ],
    ),
    # --- 术语层面 ---
    Category(
        id="terminology",
        label="通用专业术语",
        group="term",
        description="同一概念全书统一译名、统一写法",
        prompt_focus=[
            '同一概念全书使用同一译名（如统一用"队列"不混用"伫列"）',
            "标准术语优先（参照全国科学技术名词审定委员会发布的规范术语）",
            "术语首次出现是否给出英文原文或定义",
        ],
        default=True,  # 现有行为保留
    ),
    Category(
        id="cn_en_terms",
        label="中英术语对照",
        group="term",
        description='形如"队列(queue)"的中英对照规范',
        prompt_focus=[
            '中英对照格式统一（推荐"中文（English）"，括号用全角）',
            "英文术语首次出现给中文译名，重复出现可省略",
            "推荐译名以全国计算机科学技术名词审定委员会的发布为准",
            '如"鲁棒性/健壮性/稳健性 (robustness)"应择一统一',
        ],
    ),
    # --- 计算机教材专属 ---
    Category(
        id="code",
        label="代码片段",
        group="cs",
        description="代码语法、变量名一致性、注释、全角符号污染",
        prompt_focus=[
            '代码片段中是否混入全角字符（"（）""，""；"等会导致编译错）',
            "变量、函数命名风格一致（snake_case / camelCase 不混用）",
            "代码缩进一致、空格/Tab 不混用",
            "注释与代码保持同步、注释行使用语言对应注释符（# / // / /* */）",
            "字符串引号风格一致",
        ],
    ),
    Category(
        id="formula",
        label="公式与符号",
        group="cs",
        description="GB 3102.11；变量斜体、函数名正体、单位与数值空格",
        prompt_focus=[
            "数学变量用斜体（如 x、y、n）",
            "函数名、log、ln、sin、cos 等用正体",
            "矢量加粗或带箭头表示",
            '数值与单位之间保留空格（如 "5 ms" 而非 "5ms"）',
            "依据 GB 3102.11《物理科学和技术中使用的数学符号》",
        ],
    ),
    # --- 内容层面 (高风险) ---
    Category(
        id="knowledge",
        label="知识性差错",
        group="content",
        description="计算机学科事实、人物年代、引文出处。LLM 在冷门事实上易幻觉",
        prompt_focus=[
            "计算机学科基本事实（算法复杂度、协议规范、历史事件）",
            "人物姓名/年代/事件准确性",
            "引文与原文一致",
            '如不确定请标注"建议核实"而非武断修改',
        ],
        high_risk=True,
    ),
    Category(
        id="political",
        label="政治性差错",
        group="content",
        description="⚠️ 红线问题，LLM 不得独立判定，建议人审兜底；默认关闭",
        prompt_focus=[
            "国家、地区、领土表述（不得出现影响主权、领土完整的表述）",
            "民族、宗教用语规范",
            "领导人姓名职务准确",
            '**仅做提示性扫描**：发现可疑表述时标注"⚠️建议人工审阅"，不做修改',
        ],
        high_risk=True,
    ),
]

CATEGORY_BY_ID: dict[str, Category] = {c.id: c for c in CATEGORIES}
DEFAULT_CATEGORY_IDS: list[str] = [c.id for c in CATEGORIES if c.default]


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------
def resolve_categories(ids: list[str] | None) -> list[Category]:
    """把分类 id 列表解析为 Category 对象列表；None 表示用默认。

    未知 id 直接忽略并打日志（不抛错，避免 UI/CLI 调用异常）。
    """
    import logging

    logger = logging.getLogger(__name__)

    if ids is None:
        return [CATEGORY_BY_ID[i] for i in DEFAULT_CATEGORY_IDS]

    resolved: list[Category] = []
    for cid in ids:
        cat = CATEGORY_BY_ID.get(cid)
        if cat is None:
            logger.warning("未知校对维度: %s（已忽略）", cid)
            continue
        resolved.append(cat)
    return resolved


def group_categories(cats: list[Category]) -> dict[str, list[Category]]:
    """按 group_id 分组。返回保持 GROUPS 中的顺序，跳过空组。"""
    groups: dict[str, list[Category]] = {}
    for cat in cats:
        groups.setdefault(cat.group, []).append(cat)
    # 按 GROUPS 顺序重新排列
    return {g.id: groups[g.id] for g in GROUPS if g.id in groups}


def categories_for_ui() -> list[tuple[CategoryGroup, list[Category]]]:
    """供 UI 渲染用的（组, [分类]）列表，按 group.order 排序。"""
    by_group: dict[str, list[Category]] = {}
    for cat in CATEGORIES:
        by_group.setdefault(cat.group, []).append(cat)
    return [(g, by_group.get(g.id, [])) for g in sorted(GROUPS, key=lambda x: x.order)]
