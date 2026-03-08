"""
统一日志配置模块
"""

import logging
import sys


def setup_logging(verbose: bool = False):
    """配置统一日志系统"""
    level = logging.DEBUG if verbose else logging.INFO

    # 创建根日志记录器
    logger = logging.getLogger("ai_proofreader")
    logger.setLevel(level)

    # 避免重复添加handler
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stderr)
        handler.setLevel(level)

        formatter = logging.Formatter(
            "%(levelname)s | %(name)s | %(message)s"
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    return logger


def get_logger(name: str) -> logging.Logger:
    """获取指定名称的日志记录器"""
    return logging.getLogger(f"ai_proofreader.{name}")
