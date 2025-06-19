# 变更日志

所有项目的重要变更都会记录在此文件中。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)，
项目遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

### 新增
- 完善的需求管理体系
- 版本规划文档
- 完整的单元测试框架
- CI/CD 自动化流水线
- 代码质量检查工具集成
- 开发指南和贡献指南

### 变更
- 项目结构重组，更清晰的模块划分
- 测试覆盖率要求提升至80%

### 修复
- 配置文件加载的优先级问题

## [1.0.0] - 2024-06-19

### 新增
- 基础的Word文档校对功能
- 多种校对模式支持：
  - 批注模式 (comments)
  - 修订模式 (revisions) 
  - 跟踪更改模式 (track_changes)
  - 增强模式 (enhanced)
- 快速文本检查功能
- 命令行界面 (CLI)
- 批量处理功能
- 配置文件支持
- OpenAI API 集成
- 丰富的输出格式

### 特性
- 专业的中文计算机教材校对
- 智能错误检测和建议
- 友好的用户界面
- 详细的帮助文档

## 版本说明

### 版本号格式
- 主版本号：不兼容的API修改
- 次版本号：向后兼容的功能性新增  
- 修订号：向后兼容的问题修正

### 标签说明
- **新增** (Added): 新功能
- **变更** (Changed): 现有功能的变更
- **弃用** (Deprecated): 即将删除的功能
- **移除** (Removed): 已删除的功能
- **修复** (Fixed): 错误修复
- **安全** (Security): 安全相关修复

[Unreleased]: https://github.com/jack91620/AI-DocProofreader/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/jack91620/AI-DocProofreader/releases/tag/v1.0.0 