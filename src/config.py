"""
配置管理模块
"""

import os
from typing import Dict, List, Optional
from pydantic import BaseModel
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()


class AIConfig(BaseModel):
    """AI配置"""
    provider: str = "openai"  # "openai" 或 "claude"
    api_key: str
    base_url: str = "https://api.openai.com/v1"
    model: str = "gpt-4"
    max_tokens: int = 4000
    temperature: float = 0.1
    # 分块处理配置
    chunk_max_tokens: int = 6000  # 每个块的最大token数
    chunk_overlap_tokens: int = 200  # 块间重叠token数
    safe_token_buffer: int = 2000  # 安全缓冲token数


class ProofreadingRules(BaseModel):
    """校对规则配置"""
    check_grammar: bool = True
    check_spelling: bool = True
    check_terminology: bool = True
    check_punctuation: bool = True
    check_consistency: bool = True


class CommentStyle(BaseModel):
    """批注样式配置"""
    author: str = "AI校对助手"
    color: str = "red"
    highlight_color: str = "yellow"


class Config:
    """主配置类"""
    
    def __init__(self):
        # 检测AI提供商
        provider = os.getenv("AI_PROVIDER", "").lower()

        # 如果未显式设置provider，自动检测
        if not provider:
            if os.getenv("CLAUDE_SESSION_TOKEN") or os.getenv("ANTHROPIC_API_KEY"):
                provider = "claude"
            else:
                provider = "openai"

        # 模型token限制
        openai_token_limits = {
            "gpt-4": 8192,
            "gpt-4-turbo": 128000,
            "gpt-4o": 128000,
            "gpt-3.5-turbo": 4096,
            "gpt-3.5-turbo-16k": 16384,
        }
        claude_token_limits = {
            "claude-sonnet-4-20250514": 200000,
            "claude-opus-4-20250514": 200000,
            "claude-haiku-4-5-20251001": 200000,
            "claude-3-5-sonnet-20241022": 200000,
            "claude-3-5-haiku-20241022": 200000,
            "claude-3-opus-20240229": 200000,
        }

        if provider == "claude":
            # Claude配置
            api_key = os.getenv("CLAUDE_SESSION_TOKEN") or os.getenv("ANTHROPIC_API_KEY", "")
            base_url = os.getenv("ANTHROPIC_BASE_URL", "https://api.anthropic.com")
            default_model = "claude-sonnet-4-20250514"
            model_name = os.getenv("CLAUDE_MODEL", os.getenv("ANTHROPIC_MODEL", default_model))
            model_max_tokens = claude_token_limits.get(model_name, 200000)
        else:
            # OpenAI配置
            api_key = os.getenv("OPENAI_API_KEY", "")
            base_url = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
            model_name = os.getenv("OPENAI_MODEL", "gpt-4")
            model_max_tokens = openai_token_limits.get(model_name, 8192)

        # 用户设置的token限制不能超过模型限制
        user_max_tokens = int(os.getenv("MAX_TOKENS", str(model_max_tokens)))
        safe_max_tokens = min(user_max_tokens, model_max_tokens)

        self.ai = AIConfig(
            provider=provider,
            api_key=api_key,
            base_url=base_url,
            model=model_name,
            max_tokens=safe_max_tokens,
            temperature=float(os.getenv("TEMPERATURE", "0.1")),
            chunk_max_tokens=safe_max_tokens,
            chunk_overlap_tokens=int(os.getenv("CHUNK_OVERLAP_TOKENS", "200")),
            safe_token_buffer=int(os.getenv("SAFE_TOKEN_BUFFER", "1000")),
        )
        
        self.rules = ProofreadingRules()
        self.comment_style = CommentStyle()
        
        # 计算机术语词典
        self.terminology_dict = {
            "数据库": ["数据库", "资料库"],
            "算法": ["算法", "演算法"],
            "程序": ["程序", "程式"],
            "软件": ["软件", "软体"],
            "硬件": ["硬件", "硬体"],
            "网络": ["网络", "网路"],
            "编程": ["编程", "程式设计"],
            "调试": ["调试", "除错"],
            "变量": ["变量", "变数"],
            "函数": ["函数", "函式"],
            "对象": ["对象", "物件"],
            "类": ["类", "类别"],
            "接口": ["接口", "介面"],
            "框架": ["框架", "架构"],
            "系统": ["系统", "系统"],
        }
        
        # 常见错别字
        self.typo_dict = {
            "计算机": ["计算器"],
            "程序": ["程式"],
            "软件": ["软体"],
            "网络": ["网路"],
            "数据": ["资料"],
            "文件": ["档案"],
            "目录": ["目录"],
            "删除": ["刪除"],
            "复制": ["复制"],
            "粘贴": ["贴上"],
        }
    
    def validate(self) -> bool:
        """验证配置"""
        if not self.ai.api_key:
            if self.ai.provider == "claude":
                raise ValueError("Claude API Key 未设置，请设置 CLAUDE_SESSION_TOKEN 或 ANTHROPIC_API_KEY 环境变量")
            else:
                raise ValueError("OpenAI API Key 未设置")

        if self.ai.max_tokens <= 0:
            raise ValueError(f"最大token数必须为正数，当前值: {self.ai.max_tokens}")

        if self.ai.chunk_overlap_tokens >= self.ai.max_tokens:
            raise ValueError(f"重叠token数 ({self.ai.chunk_overlap_tokens}) 不能大于等于最大token数 ({self.ai.max_tokens})")

        return True 