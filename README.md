# AI校对助手

一个专为中文计算机教材设计的AI校对工具，能够自动检查docx文档并生成带批注的校对结果。

## 功能特性

- 📖 **docx文档处理**: 支持读取和写入Microsoft Word文档
- 🤖 **AI智能校对**: 使用大语言模型进行语法、错别字、术语检查
- 📝 **批注系统**: 在原文档中直接添加校对建议和批注
- ⚙️ **配置管理**: 支持自定义校对规则和术语词典
- 🎯 **专业领域**: 针对计算机领域术语优化

## 安装

1. 克隆项目
```bash
git clone <your-repo>
cd Editor
```

2. 创建conda虚拟环境
```bash
conda env create -f environment.yml
conda activate ai-proofreader
```

3. 配置环境变量
创建 `.env` 文件并添加你的 OpenAI API Key：
```bash
OPENAI_API_KEY=your_openai_api_key_here
OPENAI_BASE_URL=https://api.openai.com/v1
OPENAI_MODEL=gpt-4
MAX_TOKENS=4000
TEMPERATURE=0.1
```

## 使用方法

### 命令行使用

**基本校对：**
```bash
python main.py proofread -i input.docx -o output.docx
```

**批量校对：**
```bash
python main.py batch -i ./input_folder -o ./output_folder
```

**快速文本检查：**
```bash
python main.py check -t "要检查的文本内容"
```

**环境检查：**
```bash
python main.py setup
```

**运行演示：**
```bash
python main.py demo
```

### 程序化使用
```python
from proofreader import ProofReader

# 创建校对器实例
proofreader = ProofReader()

# 校对文档
proofreader.proofread_document("input.docx", "output.docx")
```

## 配置

通过环境变量配置AI参数：
- `OPENAI_API_KEY`: OpenAI API密钥
- `OPENAI_MODEL`: 使用的模型（默认：gpt-4）
- `MAX_TOKENS`: 最大token数（默认：4000）
- `TEMPERATURE`: 温度参数（默认：0.1）

程序内置了计算机领域的术语词典和校对规则，可以在代码中自定义修改。

## 项目结构

```
Editor/
├── main.py              # 主程序入口
├── proofreader/         # 核心校对模块
│   ├── __init__.py
│   ├── document.py      # 文档处理
│   ├── ai_checker.py    # AI校对引擎
│   ├── annotator.py     # 批注系统
│   └── config.py        # 配置管理
├── environment.yml       # conda环境配置文件
├── requirements.txt     # 依赖列表
└── README.md           # 项目说明
``` 