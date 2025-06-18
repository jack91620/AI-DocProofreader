# AI校对助手

一个专为中文计算机教材设计的AI校对工具，支持多种校对模式，能够自动检查docx文档并生成专业的校对结果。

## 🌟 功能特性

### 📖 核心功能
- **docx文档处理**: 支持读取和写入Microsoft Word文档
- **AI智能校对**: 使用大语言模型进行语法、错别字、术语检查
- **多种校对模式**: 支持批注、修订、跟踪更改和增强模式
- **专业领域优化**: 针对计算机领域术语特别优化
- **配置管理**: 支持自定义校对规则和术语词典

### 🔧 四种校对模式

#### 1. 💬 批注模式 (comments)
- 使用Word审阅批注功能
- 在Word审阅窗格中显示修改建议
- 保持原文档不变，通过批注提供建议
- 适合需要保留原文档完整性的场景

#### 2. ✏️ 修订模式 (revisions)
- 使用Word修订标记功能
- 通过文本样式显示修改内容
- 原始错误文本显示为删除线
- 修正后文本显示为下划线
- 适合快速查看修改对比

#### 3. 🔄 跟踪更改模式 (track_changes)
- 使用真正的Microsoft Word跟踪更改功能
- 生成标准的Word XML修订标记（`<w:del>`, `<w:ins>`, `<w:delText>`）
- 包含完整的修订元数据（ID、作者、时间戳）
- 完全兼容Microsoft Word的审阅功能
- 可在Word中完整操作（接受/拒绝修改）

#### 4. 🌟 增强模式 (enhanced) - 推荐
- **同时使用跟踪更改和批注功能**
- 真正的Word修订标记 + 详细的批注说明
- 为每个修订位置添加详细的批注
- 说明修订原因、类型和对比信息
- 提供最完整和专业的校对体验

## 🚀 安装与配置

### 1. 环境准备
```bash
# 克隆项目
git clone <your-repo>
cd Editor

# 创建conda虚拟环境
conda env create -f environment.yml
conda activate ai-proofreader

# 安装依赖
pip install -r requirements.txt
```

### 2. API配置
创建 `.env` 文件并添加你的 OpenAI API Key：
```bash
OPENAI_API_KEY=your_openai_api_key_here
OPENAI_BASE_URL=https://api.openai.com/v1
OPENAI_MODEL=gpt-4
MAX_TOKENS=4000
TEMPERATURE=0.1
```

或通过环境变量设置：
```bash
export OPENAI_API_KEY=your_openai_api_key_here
```

## 🚀 快速开始

### 环境准备
```bash
# 激活conda环境
conda activate ai-proofreader

# 设置API密钥
export OPENAI_API_KEY=your_api_key_here
```

### 推荐使用（增强模式）
```bash
# 使用增强模式（跟踪更改+批注）- 推荐
python main.py proofread -i sample_input.docx -o output.docx -m enhanced
```

## 📚 详细使用方法

### 命令行使用

#### 基本校对命令
```bash
# 增强模式（推荐）- 跟踪更改+批注
python main.py proofread -i sample_input.docx -o output.docx -m enhanced

# 跟踪更改模式
python main.py proofread -i sample_input.docx -o output.docx -m track_changes

# 修订模式
python main.py proofread -i sample_input.docx -o output.docx -m revisions

# 批注模式（默认）
python main.py proofread -i sample_input.docx -o output.docx -m comments
```

#### 快捷命令
```bash
# 修订模式快捷命令
python main.py revise -i sample_input.docx -o output.docx
```

#### 其他功能
```bash
# 批量校对
python main.py batch -i ./input_folder -o ./output_folder

# 快速文本检查
python main.py check -t "要检查的文本内容"
python main.py check -f text_file.txt

# 环境检查
python main.py setup

# 运行演示
python main.py demo
```

### 程序化使用
```python
from proofreader import ProofReader
from proofreader.proofreader_track_changes_enhanced import ProofReaderWithTrackChangesAndComments

# 基本批注模式
proofreader = ProofReader(api_key)
proofreader.proofread_document("sample_input.docx", "output.docx")

# 增强模式（推荐）
enhanced_proofreader = ProofReaderWithTrackChangesAndComments(api_key)
enhanced_proofreader.proofread_with_track_changes_and_comments("sample_input.docx", "output.docx")
```

## 📝 在Microsoft Word中的效果

### 增强模式显示效果
- **红色删除线文本** - 需要删除的原始错误内容
- **蓝色下划线文本** - 新插入的正确内容  
- **审阅窗格** - 显示修订的详细信息和元数据
- **批注说明** - 每个修订的详细原因和类型说明
- **作者标识** - "AI校对助手"
- **时间戳** - 精确到秒的修订时间

### 批注内容示例
```
🔄 修订说明：
原文：'计算器科学'
修正：'计算机科学'
原因：错别字修正，'器'应为'机'
类型：错别字纠正
时间：2024-01-01 12:00:00
```

### Word中的操作功能
- ✅ **接受修改** - 应用AI建议的修正
- ❌ **拒绝修改** - 保留原始文本
- 💬 **回复批注** - 与AI校对助手"对话"
- 📝 **查看详情** - 完整的修订历史
- 🔄 **批量操作** - 批量接受或拒绝修改

## 🔍 校对能力

### 检测类型
- ✅ **错别字检测** - 识别常见的错别字和笔误
- ✅ **术语一致性检查** - 确保专业术语使用统一
- ✅ **标点符号规范** - 检查标点符号使用规范
- ✅ **语法错误识别** - 发现语法结构问题
- ✅ **专业术语校正** - 计算机领域术语优化
- ✅ **逻辑表达优化** - 改善句子表达逻辑

### 实际测试案例

#### 输入文本示例
```
计算器科学是一门非常重要的学科，涉及到程式设计和筭法等内容。
```

#### 校对结果
- `计算器科学` → `计算机科学` ✅ 错别字修正
- `程式设计` → `程序设计` ✅ 术语统一
- `筭法` → `算法` ✅ 错别字修正

## 🔧 技术架构

### 核心模块
```
proofreader/
├── __init__.py                           # 模块初始化
├── config.py                            # 配置管理
├── document.py                          # 文档处理
├── ai_checker.py                        # AI校对引擎
├── proofreader.py                       # 基础校对器（批注模式）
├── proofreader_revisions.py             # 修订模式校对器
├── proofreader_track_changes.py         # 跟踪更改校对器
├── proofreader_track_changes_enhanced.py # 增强版校对器
├── word_comments.py                     # Word批注功能
├── word_comments_advanced.py            # 高级批注功能
├── word_revisions.py                    # Word修订功能
├── word_track_changes.py                # Word跟踪更改功能
├── word_track_changes_with_comments.py  # 跟踪更改+批注
├── create_word_comments_xml.py          # 批注XML生成
└── create_word_revisions_xml.py         # 修订XML生成
```

### 关键类
- `ProofReader` - 基础校对器（批注模式）
- `ProofReaderWithRevisions` - 修订模式校对器
- `ProofReaderWithTrackChanges` - 跟踪更改校对器
- `ProofReaderWithTrackChangesAndComments` - 增强版校对器
- `WordTrackChangesManager` - 跟踪更改管理器
- `WordCommentsManager` - 批注管理器
- `AIChecker` - AI校对引擎

### XML处理技术
- **document.xml** - 修订标记生成
- **comments.xml** - 批注内容管理
- **settings.xml** - 跟踪更改启用
- **document.xml.rels** - 关系映射
- **[Content_Types].xml** - 内容类型定义

## 📊 模式对比

| 功能特性 | 批注模式 | 修订模式 | 跟踪更改模式 | **增强模式** ⭐ |
|----------|----------|----------|--------------|----------------|
| Word兼容性 | ✅ 标准批注 | ⚠️ 文本样式 | ✅ 原生支持 | ✅ **完全兼容** |
| 可操作性 | ✅ 可回复批注 | ❌ 纯显示 | ✅ 可接受/拒绝 | ✅ **完全可操作** |
| 详细说明 | ✅ 批注说明 | ❌ 无说明 | ❌ 无说明 | ✅ **详细批注** |
| 专业程度 | ⭐⭐⭐ | ⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| 协作友好 | ✅ 审阅窗格 | ❌ 混乱显示 | ✅ 标准审阅 | ✅ **完美协作** |
| 推荐场景 | 简单审查 | 快速对比 | 标准校对 | **专业校对** |

## 📁 文件说明

### 核心文件
- **输入文件**: `sample_input.docx` - 包含各种需要校对的文本
- **输出文件**: `output.docx` - 统一的输出文件名（避免多个测试文件）

### 辅助工具
- `create_sample.py` - 创建包含各种问题的示例文档
- `example.py` - 使用示例代码，展示校对器功能
- `show_word_comments.py` - 显示Word文档批注内容
- `compare_documents.py` - 对比两个文档的差异

## ⚙️ 配置选项

### 环境变量配置
```bash
# OpenAI API配置
OPENAI_API_KEY=your_api_key
OPENAI_BASE_URL=https://api.openai.com/v1
OPENAI_MODEL=gpt-4

# 校对参数
MAX_TOKENS=4000
TEMPERATURE=0.1
```

### 配置文件 (config.ini)
```ini
[ai]
api_key = your_api_key
model = gpt-4
max_tokens = 4000
temperature = 0.1

[proofreading]
check_grammar = true
check_spelling = true
check_terminology = true
```

## 💡 使用建议

### 推荐工作流程
1. **备份原文档** - 处理前务必备份原始文档
2. **选择增强模式** - 获得最完整的校对体验
3. **逐一审查修订** - 仔细检查每个AI建议
4. **利用批注互动** - 通过批注了解修改原因
5. **批量操作** - 确认无误后批量接受修改

### 实用建议
- **首选增强模式**: 同时包含跟踪更改和批注，提供最完整的校对体验
- **统一输出**: 所有测试都输出到`output.docx`，避免文件混乱
- **Word查看**: 使用Microsoft Word打开`output.docx`查看完整效果
- **首次使用**: 建议先用demo模式熟悉功能
- **大型文档**: 建议分段处理，避免API调用超时
- **协作场景**: 跟踪更改模式便于团队协作

### 🎯 核心优势

✅ **同时支持跟踪更改和批注**  
✅ **完全兼容Microsoft Word**  
✅ **详细的修改说明**  
✅ **可接受/拒绝每个修改**

### 注意事项
1. **Word版本** - 建议使用Microsoft Word 2016及以上版本
2. **文档格式** - 输入必须是.docx格式文件
3. **API配置** - 确保OpenAI API密钥配置正确
4. **网络连接** - 需要稳定的网络连接调用API
5. **文档备份** - 重要文档建议先备份再处理

## 🔧 故障排除

### 常见问题

#### Q: 增强模式不工作？
A: 检查是否正确配置了API密钥，运行 `python main.py setup` 验证环境

#### Q: 跟踪更改在Word中不显示？
A: 确保Word的"审阅"选项卡中"跟踪更改"功能已启用

#### Q: 修订标记显示异常？
A: 检查Word文档是否正确启用了跟踪更改功能，或重新运行校对命令

#### Q: API调用失败？
A: 检查API密钥配置、网络连接和API配额

### 技术支持
- 查看命令行输出了解详细错误信息
- 使用`python main.py setup`检查环境配置
- 确保所有依赖正确安装
- 运行`python example.py`测试基本功能

## 📈 项目结构

```
Editor/
├── main.py                              # 主程序入口
├── proofreader/                         # 核心校对模块
│   ├── __init__.py                      # 模块初始化
│   ├── config.py                        # 配置管理
│   ├── document.py                      # 文档处理
│   ├── ai_checker.py                    # AI校对引擎
│   └── ...                              # 其他核心模块
├── sample_input.docx                    # 测试输入文件
├── output.docx                          # 统一输出文件
├── environment.yml                      # conda环境配置
├── requirements.txt                     # 依赖列表
├── README.md                           # 项目说明（本文件）
├── create_sample.py                     # 创建示例文档
├── example.py                           # 使用示例代码
├── show_word_comments.py                # 显示Word批注
└── compare_documents.py                 # 文档对比工具
```

## 🎉 总结

AI校对助手提供了业界最完整的Word文档校对解决方案：

- ✅ **四种校对模式** - 满足不同场景需求
- ✅ **增强模式** - 跟踪更改+批注双重支持
- ✅ **专业级体验** - 完全兼容Microsoft Word
- ✅ **AI智能校对** - 准确识别各类问题
- ✅ **开源免费** - 完全开源，可自由定制

现在您可以享受**专业级的AI文档校对体验**！

---

## 🔄 Git版本管理

### 版本控制最佳实践

#### 当前版本标签
- `v1.2.0` - 增强模式修复版：完全支持Word跟踪更改+批注功能
- `v1.1.0` - 基础功能完整版  
- `v1.0.0` - 初始版本

#### Git工作流程
```bash
# 查看当前状态
git status

# 添加所有更改
git add .

# 提交更改（使用规范的提交消息）
git commit -m "feat: 新增功能描述" 
git commit -m "fix: 修复问题描述"
git commit -m "refactor: 重构描述"

# 创建版本标签
git tag v1.2.0 -m "版本说明"

# 查看提交历史
git log --oneline -10
```

#### 提交消息规范
- `feat:` 新功能
- `fix:` 错误修复  
- `refactor:` 重构代码
- `docs:` 文档更新
- `test:` 测试相关
- `chore:` 构建过程或辅助工具的变动

#### .gitignore配置
```
# 忽略生成的输出文件
*_output*.docx
*_temp*.docx  
*_test*.docx
output.docx

# 保留示例文件
!sample_input.docx

# Python运行时文件
__pycache__/
*.pyc
.env
```

#### 文件管理策略
- ✅ **核心代码** - 全部受版本控制
- ✅ **示例文件** - `sample_input.docx` 受版本控制
- ❌ **输出文件** - 生成的校对结果不受版本控制
- ❌ **临时文件** - 测试和临时文件被忽略

---

*版本: v1.2.0 增强版*  
*功能: 四种校对模式全支持*  
*兼容性: Microsoft Word 2016+*  
*推荐: 增强模式（跟踪更改+批注）* 