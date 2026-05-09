# AI文档校对器 - Makefile (uv-driven)
.PHONY: help sync test test-fast test-cov test-live lint format type-check security clean all dev check-release git-hooks

help:
	@echo "AI文档校对器开发工具"
	@echo ""
	@echo "环境:"
	@echo "  sync         uv sync 安装项目+dev依赖+全部可选后端"
	@echo "  clean        清理缓存与构建产物"
	@echo ""
	@echo "测试:"
	@echo "  test         运行所有非live测试"
	@echo "  test-fast    并行跑单元测试 (-n auto)"
	@echo "  test-cov     运行+覆盖率报告"
	@echo "  test-live    跑 @pytest.mark.live (需要真实API token)"
	@echo "  test-update-snapshots  syrupy快照重录"
	@echo "  test-record-cassettes  pytest-recording 重录HTTP"
	@echo ""
	@echo "代码质量:"
	@echo "  lint         ruff lint"
	@echo "  format       ruff format + lint --fix"
	@echo "  type-check   mypy"
	@echo "  security     bandit"
	@echo "  all          format + lint + type + security + test-cov"
	@echo ""
	@echo "Git:"
	@echo "  git-hooks    安装pre-commit钩子"

# === 环境 ===
sync:
	uv sync --extra all --group dev

# === 测试 ===
test:
	uv run pytest

test-fast:
	uv run pytest -n auto -m "not slow and not live"

test-cov:
	uv run pytest --cov=src --cov-report=term-missing --cov-report=html --cov-report=xml

test-live:
	uv run pytest -m live --run-live

test-update-snapshots:
	uv run pytest --snapshot-update

test-record-cassettes:
	uv run pytest --record-mode=once

# === 代码质量 ===
lint:
	uv run ruff check .

format:
	uv run ruff format .
	uv run ruff check --fix .

type-check:
	uv run mypy src

security:
	uv run bandit -c pyproject.toml -r src

all: format lint type-check security test-cov

# === 清理 ===
clean:
	find . -type f -name "*.pyc" -delete
	find . -type d -name "__pycache__" -exec rm -rf {} + 2>/dev/null || true
	find . -type d -name "*.egg-info" -exec rm -rf {} + 2>/dev/null || true
	rm -rf build/ dist/ .coverage htmlcov/ coverage.xml
	rm -rf .pytest_cache/ .mypy_cache/ .ruff_cache/

# === Git ===
git-hooks:
	uv run pre-commit install
	uv run pre-commit install --hook-type commit-msg

# === 应用入口 ===
dev:
	uv run python main.py --help
