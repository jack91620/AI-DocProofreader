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
    api_key: str
    base_url: str = "https://api.openai.com/v1"
    model: str = "gpt-4"
    max_tokens: int = 4000
    temperature: float = 0.1
    # 新增分块处理配置
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
        # 获取模型的实际token限制
        model_token_limits = {
            "gpt-4": 8192,
            "gpt-4-turbo": 128000,
            "gpt-4o": 128000,
            "gpt-3.5-turbo": 4096,
            "gpt-3.5-turbo-16k": 16384
        }
        
        model_name = os.getenv("OPENAI_MODEL", "gpt-4")
        model_max_tokens = model_token_limits.get(model_name, 8192)
        
        # 用户设置的token限制不能超过模型限制
        user_max_tokens = int(os.getenv("MAX_TOKENS", str(model_max_tokens)))
        safe_max_tokens = min(user_max_tokens, model_max_tokens)
        
        self.ai = AIConfig(
            api_key=os.getenv("OPENAI_API_KEY", ""),
            base_url=os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1"),
            model=model_name,
            max_tokens=safe_max_tokens,  # 直接使用模型限制，分割器内部会处理API开销
            temperature=float(os.getenv("TEMPERATURE", "0.1")),
            chunk_max_tokens=safe_max_tokens,  # 与max_tokens保持一致
            chunk_overlap_tokens=int(os.getenv("CHUNK_OVERLAP_TOKENS", "200")),
            safe_token_buffer=int(os.getenv("SAFE_TOKEN_BUFFER", "1000"))  # 减少缓冲区，因为分割器内部已处理
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
            raise ValueError("OpenAI API Key 未设置")
        
        # 验证token配置的合理性
        if self.ai.chunk_max_tokens >= self.ai.max_tokens:
            raise ValueError(f"分块token数 ({self.ai.chunk_max_tokens}) 不能大于等于最大token数 ({self.ai.max_tokens})")
        
        return True 