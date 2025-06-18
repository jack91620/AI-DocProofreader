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
        self.ai = AIConfig(
            api_key=os.getenv("OPENAI_API_KEY", ""),
            base_url=os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1"),
            model=os.getenv("OPENAI_MODEL", "gpt-4"),
            max_tokens=int(os.getenv("MAX_TOKENS", "4000")),
            temperature=float(os.getenv("TEMPERATURE", "0.1"))
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
        return True 