# AI文档校对器 - 开发指南

## 环境准备

### 系统要求
- Python 3.8+
- Git
- 推荐使用虚拟环境

### 开发环境搭建

```bash
# 克隆项目
git clone https://github.com/jack91620/AI-DocProofreader.git
cd AI-DocProofreader

# 创建虚拟环境
python -m venv venv
source venv/bin/activate  # Linux/Mac
# 或
venv\Scripts\activate     # Windows

# 安装依赖
pip install -r requirements.txt
pip install -r requirements-dev.txt  # 开发依赖

# 安装预提交钩子
pre-commit install
```

## 开发工作流

### 分支管理
- `master`: 主分支，用于生产环境
- `develop`: 开发分支，用于集成开发
- `feature/*`: 功能分支
- `bugfix/*`: 修复分支
- `hotfix/*`: 热修复分支

### 提交规范
使用 Conventional Commits 规范：

```
<type>[optional scope]: <description>

[optional body]

[optional footer(s)]
```

类型说明：
- `feat`: 新功能
- `fix`: 修复bug
- `docs`: 文档更新
- `style`: 代码格式化
- `refactor`: 代码重构
- `test`: 测试相关
- `chore`: 构建过程或辅助工具的变动

示例：
```
feat(proofreader): 添加批量校对功能

添加了批量处理多个Word文档的功能，
支持并发处理以提高效率。

Closes #123
```

## 代码质量

### 代码风格
- 使用 Black 进行代码格式化
- 使用 isort 整理导入
- 使用 flake8 进行代码检查
- 使用 mypy 进行类型检查

```bash
# 运行代码质量检查
black .
isort .
flake8 .
mypy .
```

### 测试要求
- 单元测试覆盖率 >= 80%
- 所有新功能必须有对应测试
- 测试命名规范：`test_功能描述`

```bash
# 运行测试
pytest

# 运行测试并生成覆盖率报告
pytest --cov=. --cov-report=html

# 运行特定测试
pytest tests/test_main.py::TestMainModule::test_cli_version
```

## 测试策略

### 测试分类
1. **单元测试** (`@pytest.mark.unit`)
   - 测试单个函数或方法
   - 快速执行，无外部依赖

2. **集成测试** (`@pytest.mark.integration`)
   - 测试模块间交互
   - 可能涉及文件系统、网络等

3. **端到端测试** (`@pytest.mark.e2e`)
   - 测试完整用户场景
   - 较慢，用于关键路径验证

### 测试数据管理
- 使用 `tests/fixtures/` 目录存放测试数据
- 使用 `conftest.py` 定义共享的 fixtures
- 敏感数据使用环境变量或模拟

### Mock 使用原则
- 外部API调用必须Mock
- 文件系统操作建议Mock
- 数据库操作必须Mock

## CI/CD 流程

### 自动化检查
每次提交会自动运行：
1. 代码风格检查
2. 单元测试
3. 覆盖率检查
4. 安全性扫描
5. 类型检查

### 发布流程
1. 创建发布分支 `release/v1.x.x`
2. 更新版本号和变更日志
3. 运行完整测试套件
4. 合并到 master 分支
5. 创建 Git 标签
6. 发布到 PyPI（如果需要）

## 调试指南

### 本地调试
```bash
# 启用详细日志
export LOG_LEVEL=DEBUG

# 运行特定命令
python main.py proofread --input sample_input.docx --mode comments
```

### 测试调试
```bash
# 使用 pdb 调试
pytest --pdb

# 只运行失败的测试
pytest --lf

# 运行到第一个失败就停止
pytest -x
```

## 性能优化

### 性能测试
```bash
# 使用 pytest-benchmark
pytest tests/test_performance.py --benchmark-only

# 生成性能报告
pytest --benchmark-only --benchmark-save=baseline
```

### 内存分析
```bash
# 使用 memory_profiler
pip install memory_profiler
python -m memory_profiler main.py
```

## 文档编写

### 代码文档
- 使用 Google 风格的 docstring
- 重要函数必须有文档
- 复杂逻辑添加内联注释

```python
def proofread_document(self, input_file: str, output_file: str, mode: str) -> bool:
    """校对Word文档
    
    Args:
        input_file: 输入文档路径
        output_file: 输出文档路径
        mode: 校对模式 ('comments', 'revisions', 'track_changes', 'enhanced')
    
    Returns:
        bool: 校对是否成功
        
    Raises:
        FileNotFoundError: 输入文件不存在
        ValueError: 不支持的校对模式
    """
```

### 用户文档
- README.md: 项目介绍和快速开始
- docs/: 详细文档目录
- 示例代码要能直接运行

## 依赖管理

### 生产依赖
```bash
# 添加新依赖
pip install new-package
pip freeze > requirements.txt

# 或使用 pip-tools
pip-compile requirements.in
```

### 开发依赖
开发专用依赖放在 `requirements-dev.txt`:
- pytest
- pytest-cov
- black
- flake8
- mypy
- pre-commit

### 安全更新
定期检查依赖安全性：
```bash
pip install safety
safety check
```

## 发布检查清单

### 版本发布前
- [ ] 所有测试通过
- [ ] 代码覆盖率达标
- [ ] 文档更新完整
- [ ] 变更日志记录
- [ ] 版本号正确更新
- [ ] 依赖版本锁定

### 发布后
- [ ] GitHub Release 创建
- [ ] 标签正确推送
- [ ] 文档站点更新
- [ ] 发布公告发送

## 常见问题

### Q: 测试运行很慢怎么办？
A: 
1. 使用 `pytest -x` 快速失败
2. 运行特定测试文件
3. 使用 `pytest --lf` 只运行失败的测试
4. 考虑并行测试 `pytest -n auto`

### Q: 如何处理测试中的API调用？
A: 
1. 使用 Mock 替代真实API调用
2. 使用 `responses` 库模拟HTTP响应
3. 创建测试专用的配置

### Q: 代码覆盖率不达标怎么办？
A:
1. 添加遗漏的测试用例
2. 删除不必要的代码
3. 使用 `# pragma: no cover` 排除特定行

## 工具配置

### VSCode 配置
```json
{
    "python.defaultInterpreterPath": "./venv/bin/python",
    "python.testing.pytestEnabled": true,
    "python.linting.enabled": true,
    "python.linting.flake8Enabled": true,
    "python.formatting.provider": "black"
}
```

### PyCharm 配置
1. 设置 Python 解释器为虚拟环境
2. 配置 pytest 作为默认测试运行器
3. 启用代码质量检查插件

## 贡献指南

1. Fork 项目到个人仓库
2. 创建功能分支
3. 提交代码和测试
4. 确保CI通过
5. 创建 Pull Request
6. 等待代码审查
7. 合并后删除分支

欢迎贡献代码！🚀 