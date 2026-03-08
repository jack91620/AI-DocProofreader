"""
校对器工厂 - mode → strategy 映射
"""

from strategies import (
    CommentsStrategy,
    VisualRevisionsStrategy,
    TrackChangesStrategy,
    EnhancedStrategy,
)


STRATEGY_MAP = {
    'comments': CommentsStrategy,
    'revisions': VisualRevisionsStrategy,
    'track_changes': TrackChangesStrategy,
    'enhanced': EnhancedStrategy,
}


def create_proofreader(api_key: str, mode: str):
    """创建带有指定策略的校对器"""
    from proofreader import ProofReader

    strategy_cls = STRATEGY_MAP.get(mode)
    if strategy_cls is None:
        raise ValueError(f"未知的校对模式: {mode}. 可用模式: {list(STRATEGY_MAP.keys())}")

    return ProofReader(api_key, strategy_cls())
