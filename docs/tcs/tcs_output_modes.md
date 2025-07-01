# 测试用例规格 - 输出模式模块
## TCS-MODE: Output Modes Test Cases

**版本**: v1.0.0  
**创建日期**: 2024-12-20  
**最后更新**: 2024-12-20  
**状态**: 定义完成

---

## TCS-MODE-001: 批注模式测试
**需求ID**: REQ-MODE-001  
**测试类别**: 功能测试  
**优先级**: 高

### 测试用例 TCS-MODE-001-01: 基本批注插入
**目的**: 验证系统能够正确插入Word批注  
**前置条件**: 
- 有校对结果数据
- 批注模式配置正确

**测试步骤**:
1. 准备包含错误的测试文档
2. 执行批注模式校对
3. 在Word中验证批注显示

**测试数据**:
```python
test_comment_data = {
    "original_text": "计算器科学是一门重要学科",
    "error_position": 0,
    "error_length": 4,
    "suggested_correction": "计算机科学",
    "comment_text": "错别字修正：'器'应为'机'",
    "comment_author": "AI校对助手"
}
```

**预期结果**:
- 批注正确插入到错误位置
- 批注内容包含修正建议和说明
- 批注作者显示为"AI校对助手"

### 测试用例 TCS-MODE-001-02: 多处批注处理
**目的**: 验证同一文档多处错误的批注处理  

**测试数据**:
```python
multiple_comments_test = [
    {"position": 10, "text": "程式设计", "suggestion": "程序设计"},
    {"position": 25, "text": "筭法", "suggestion": "算法"},
    {"position": 40, "text": "数据结够", "suggestion": "数据结构"}
]
```

**预期结果**:
- 所有错误位置都有相应批注
- 批注之间不互相干扰
- 批注顺序与错误位置一致

### 测试用例 TCS-MODE-001-03: 批注格式兼容性
**目的**: 验证批注在不同Word版本中的兼容性  

**测试环境**:
- Microsoft Word 2016
- Microsoft Word 2019  
- Microsoft Word 365

**预期结果**:
- 批注在所有版本中正确显示
- 批注可以正常回复和删除
- 格式保持一致

---

## TCS-MODE-002: 修订模式测试
**需求ID**: REQ-MODE-002  
**测试类别**: 功能测试  
**优先级**: 中

### 测试用例 TCS-MODE-002-01: 删除线和下划线样式
**目的**: 验证修订模式的文本样式显示  

**测试步骤**:
1. 准备带错误的文档
2. 执行修订模式校对
3. 检查样式标记

**预期结果**:
- 错误文本显示删除线样式
- 修正文本显示下划线样式
- 颜色区分明显（红色删除，蓝色添加）

### 测试用例 TCS-MODE-002-02: 复杂修订处理
**目的**: 验证复杂文本修订的处理  

**测试数据**:
```python
complex_revision_test = {
    "original": "这个筭法的时间复杂度很好",
    "corrections": [
        {"start": 2, "end": 4, "old": "筭法", "new": "算法"},
        {"start": 12, "end": 14, "old": "很好", "new": "很低"}
    ]
}
```

**预期结果**:
- 多处修订正确标记
- 样式不重叠或冲突
- 原文档格式保持

---

## TCS-MODE-003: 跟踪更改模式测试
**需求ID**: REQ-MODE-003  
**测试类别**: 功能测试  
**优先级**: 高

### 测试用例 TCS-MODE-003-01: Word标准跟踪更改
**目的**: 验证生成标准的Word跟踪更改标记  

**测试步骤**:
1. 执行跟踪更改模式校对
2. 在Word中打开结果文档
3. 验证跟踪更改功能

**预期结果**:
- 跟踪更改在Word中正确显示
- 可以接受/拒绝修改
- 修订历史完整记录

**测试实现**:
```python
# 测试用例实现位置：tests/integration/test_output_modes.py
def test_track_changes_xml_generation():
    """测试跟踪更改XML标记生成"""
    pass

def test_track_changes_word_compatibility():
    """测试Word兼容性"""
    pass
```

### 测试用例 TCS-MODE-003-02: 修订元数据验证
**目的**: 验证修订元数据的完整性  

**预期元数据**:
```xml
<w:ins w:id="1" w:author="AI校对助手" w:date="2024-12-20T10:30:00Z">
    <w:t>修正文本</w:t>
</w:ins>
<w:del w:id="2" w:author="AI校对助手" w:date="2024-12-20T10:30:00Z">
    <w:delText>错误文本</w:delText>
</w:del>
```

**验证项目**:
- 修订ID唯一性
- 作者信息正确
- 时间戳准确
- XML格式符合标准

### 测试用例 TCS-MODE-003-03: 批量修订处理
**目的**: 验证大量修订的处理性能  

**性能指标**:
- 100处修订处理时间 < 10秒
- 内存使用 < 500MB
- 输出文档大小合理

---

## TCS-MODE-004: 增强模式测试
**需求ID**: REQ-MODE-004  
**测试类别**: 功能测试  
**优先级**: 高

### 测试用例 TCS-MODE-004-01: 跟踪更改与批注结合
**目的**: 验证跟踪更改和批注同时工作  

**测试步骤**:
1. 执行增强模式校对
2. 验证跟踪更改标记
3. 验证批注说明
4. 检查两者关联性

**预期结果**:
- 跟踪更改标记正确生成
- 每个修订都有对应批注
- 批注内容详细说明修订原因

**测试数据**:
```python
enhanced_mode_test = {
    "input_text": "计算器科学的筭法很复杂",
    "expected_revisions": [
        {
            "position": 0,
            "old_text": "计算器科学",
            "new_text": "计算机科学",
            "comment": "错别字修正：'器'应为'机'，确保术语准确性"
        },
        {
            "position": 5,
            "old_text": "筭法",
            "new_text": "算法", 
            "comment": "错别字修正：'筭'应为'算'，专业术语规范化"
        }
    ]
}
```

### 测试用例 TCS-MODE-004-02: 详细批注说明
**目的**: 验证增强模式批注的详细程度  

**批注内容要求**:
```
🔄 修订说明：
原文：'计算器科学'
修正：'计算机科学'
原因：错别字修正，'器'应为'机'
类型：错别字纠正
时间：2024-12-20 10:30:00
```

**验证项目**:
- 批注包含原文和修正文本
- 说明修正原因和类型
- 包含时间戳信息
- 格式统一美观

### 测试用例 TCS-MODE-004-03: Word操作功能测试
**目的**: 验证在Word中的完整操作体验  

**操作测试项目**:
- 接受单个修改
- 拒绝单个修改
- 批量接受修改
- 批量拒绝修改
- 回复批注
- 删除批注

**预期结果**:
- 所有Word标准操作正常工作
- 修改接受后批注自动处理
- 修改拒绝后恢复原文

---

## 集成测试用例

### TCS-MODE-INT-001: 模式切换测试
**目的**: 验证不同输出模式之间的切换  

**测试步骤**:
1. 同一文档分别用四种模式处理
2. 比较输出结果差异
3. 验证模式特性

**预期结果**:
- 每种模式都有独特特征
- 修正建议内容一致
- 适用场景明确

### TCS-MODE-INT-002: 性能对比测试
**目的**: 比较不同模式的性能表现  

**测试指标**:
```python
performance_comparison = {
    "comments_mode": {"time": "< 30s", "memory": "< 200MB"},
    "revisions_mode": {"time": "< 25s", "memory": "< 180MB"},
    "track_changes_mode": {"time": "< 35s", "memory": "< 250MB"},
    "enhanced_mode": {"time": "< 45s", "memory": "< 300MB"}
}
```

---

## 用户体验测试

### TCS-MODE-UX-001: 易用性测试
**目的**: 验证不同模式的用户友好性  

**评估标准**:
- 学习曲线：新用户5分钟内掌握
- 操作便利：点击3次以内完成主要操作
- 视觉清晰：修订标记清晰可见
- 错误处理：友好的错误提示

### TCS-MODE-UX-002: 文档可读性测试
**目的**: 验证校对后文档的可读性  

**测试方法**:
- 打印文档可读性测试
- 屏幕阅读器兼容性
- 不同字体大小下的显示效果
- 颜色盲用户友好性

---

## 自动化测试框架

### 测试数据生成
```python
# tests/test_data/output_modes_test_data.py
class OutputModeTestData:
    """输出模式测试数据生成器"""
    
    @staticmethod
    def generate_comment_test_data():
        """生成批注模式测试数据"""
        pass
    
    @staticmethod
    def generate_revision_test_data():
        """生成修订模式测试数据"""
        pass
    
    @staticmethod
    def generate_track_changes_test_data():
        """生成跟踪更改测试数据"""
        pass
    
    @staticmethod
    def generate_enhanced_test_data():
        """生成增强模式测试数据"""
        pass
```

### 验证辅助函数
```python
# tests/utils/output_verification.py
class OutputVerification:
    """输出结果验证工具"""
    
    @staticmethod
    def verify_comments(docx_file, expected_comments):
        """验证批注正确性"""
        pass
    
    @staticmethod
    def verify_revisions(docx_file, expected_revisions):
        """验证修订样式"""
        pass
    
    @staticmethod  
    def verify_track_changes(docx_file):
        """验证跟踪更改XML"""
        pass
    
    @staticmethod
    def verify_enhanced_mode(docx_file, expected_data):
        """验证增强模式完整性"""
        pass
```

---

## 测试执行计划

### 执行顺序
1. 单元测试：各模式独立功能
2. 集成测试：模式间交互
3. 性能测试：效率对比
4. 兼容性测试：Word版本兼容
5. 用户体验测试：易用性评估

### 自动化执行
```bash
# 运行输出模式测试套件
pytest tests/integration/test_output_modes.py -v --html=reports/output_modes_test_report.html

# 运行性能对比测试
pytest tests/performance/test_output_modes_performance.py -v

# 运行兼容性测试
pytest tests/compatibility/test_word_compatibility.py -v
```

---

## 测试通过准则
- 所有四种输出模式功能正常
- Word兼容性测试100%通过
- 性能指标满足需求规格
- 用户体验评分≥8.0/10.0
- 自动化测试覆盖率≥95% 