# AI文档校对器 Makefile

.PHONY: help install test lint format clean build docs

# 默认目标
help:
	@echo "AI文档校对器开发工具"
	@echo ""
	@echo "可用命令:"
	@echo "  install     安装依赖"
	@echo "  install-dev 安装开发依赖"
	@echo "  test        运行测试"
	@echo "  test-cov    运行测试并生成覆盖率报告"
	@echo "  lint        代码检查"
	@echo "  format      代码格式化"
	@echo "  type-check  类型检查"
	@echo "  security    安全检查"
	@echo "  clean       清理缓存文件"
	@echo "  build       构建项目"
	@echo "  docs        生成文档"
	@echo "  all         运行所有检查"

# 安装依赖
install:
	pip install -r requirements.txt

install-dev:
	pip install -r requirements.txt
	pip install -r requirements-dev.txt

# 测试相关
test:
	pytest tests/ -v

test-cov:
	pytest tests/ --cov=. --cov-report=html --cov-report=term-missing --cov-report=xml

test-unit:
	pytest tests/ -v -m unit

test-integration:
	pytest tests/ -v -m integration

# 代码质量
lint:
	flake8 . --count --select=E9,F63,F7,F82 --show-source --statistics
	flake8 . --count --exit-zero --max-complexity=10 --max-line-length=127 --statistics

format:
	black .
	isort .

type-check:
	mypy . --ignore-missing-imports

security:
	bandit -r . -f json -o bandit-report.json
	safety check

# 清理
clean:
	find . -type f -name "*.pyc" -delete
	find . -type d -name "__pycache__" -delete
	find . -type d -name "*.egg-info" -exec rm -rf {} +
	rm -rf build/
	rm -rf dist/
	rm -rf .coverage
	rm -rf htmlcov/
	rm -rf .pytest_cache/
	rm -rf .mypy_cache/

# 构建
build: clean
	python -m build

# 文档
docs:
	cd docs && make html

# 运行所有检查
all: format lint type-check security test-cov

# 开发服务器 (如果需要)
dev:
	python main.py --help

# 发布检查
check-release: all
	python -m twine check dist/*
	@echo "发布检查完成！"

# Git相关
git-hooks:
	pre-commit install
	pre-commit install --hook-type commit-msg

# 初始化开发环境
init-dev: install-dev git-hooks
	@echo "开发环境初始化完成！"
	@echo "运行 'make test' 验证环境" 