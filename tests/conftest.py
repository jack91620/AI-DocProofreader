"""项目级 pytest 共享配置。

主要职责：
1. 注册 ``--run-live`` 选项控制 ``@pytest.mark.live`` 测试是否真打 LLM
2. 提供 ``vcr_config`` 默认配置：cassette 路径、脱敏 header、忽略动态字段
3. 提供工程根/dotenv 加载等共用 fixture
"""

from __future__ import annotations

import os
from pathlib import Path

import pytest
from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).parent.parent
CASSETTE_DIR = Path(__file__).parent / "cassettes"

# 项目根的 .env 在 src/config.py 里也会 load_dotenv，但测试 fixture
# 早于 src/config 导入时就要读 token，所以这里先 load 一次。
load_dotenv(PROJECT_ROOT / ".env")

# 真实 token 类的环境变量名 → 录制时脱敏为占位
SENSITIVE_HEADERS = {
    "authorization",  # OpenAI/Anthropic
    "x-api-key",  # Anthropic 备用
    "x-chj-gwtoken",  # licloud
    "bcs-apihub-requestid",
    "x-chj-gw-source",
    "cookie",
    "set-cookie",
}

# 录制时把 token 内容替换成此占位符
TOKEN_PLACEHOLDER = "REDACTED"


# ---------------------------------------------------------------------------
# CLI 选项 + marker handling
# ---------------------------------------------------------------------------
def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        "--run-live",
        action="store_true",
        default=False,
        help="运行 @pytest.mark.live 测试（会真实调用LLM API、需要token）",
    )


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    """默认跳过 live 测试，除非显式 --run-live。"""
    if config.getoption("--run-live"):
        return
    skip_live = pytest.mark.skip(reason="需要 --run-live 启用 (会真实调用LLM)")
    for item in items:
        if "live" in item.keywords:
            item.add_marker(skip_live)


# ---------------------------------------------------------------------------
# VCR 配置 (pytest-recording)
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def vcr_config() -> dict:
    """pytest-recording 默认配置 — cassette 脱敏 + 比较粒度。"""
    return {
        "cassette_library_dir": str(CASSETTE_DIR),
        # 录制策略: once = 没cassette就录，有cassette就replay；不会写覆盖
        # 用 `pytest --record-mode=rewrite` 强制重录
        "record_mode": "once",
        # 比较请求时忽略动态字段
        "match_on": ["method", "scheme", "host", "port", "path", "query"],
        # 录制 → cassette 写入前脱敏请求头
        "filter_headers": [(h, TOKEN_PLACEHOLDER) for h in SENSITIVE_HEADERS],
        # 脱敏查询参数中的token（OpenAI/Azure风格）
        "filter_query_parameters": [
            ("api-key", TOKEN_PLACEHOLDER),
            ("api_key", TOKEN_PLACEHOLDER),
        ],
        # 脱敏请求/响应 body 内的 token 字面量
        "before_record_request": _redact_request,
        "before_record_response": _redact_response,
    }


def _redact_request(request):
    """录制前抹掉请求 body 中可能的 token。"""
    for env_name in (
        "LICLOUD_TOKEN",
        "OPENAI_API_KEY",
        "ANTHROPIC_API_KEY",
        "CLAUDE_SESSION_TOKEN",
    ):
        token = os.environ.get(env_name)
        if token and request.body:
            try:
                body_str = (
                    request.body.decode() if isinstance(request.body, bytes) else str(request.body)
                )
                if token in body_str:
                    body_str = body_str.replace(token, TOKEN_PLACEHOLDER)
                    request.body = (
                        body_str.encode() if isinstance(request.body, bytes) else body_str
                    )
            except (UnicodeDecodeError, AttributeError):
                pass
    return request


def _redact_response(response):
    """response 中通常没有 token，但保留hook以便扩展。"""
    return response


# ---------------------------------------------------------------------------
# 共用 fixture
# ---------------------------------------------------------------------------
@pytest.fixture
def project_root() -> Path:
    return PROJECT_ROOT


@pytest.fixture
def sample_input_docx(project_root: Path) -> Path:
    return project_root / "sample_input.docx"
