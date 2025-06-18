"""
AI校对引擎模块
"""

import json
import re
from typing import List, Dict, Any, Optional
import openai
from .config import Config


class ProofreadingResult:
    """校对结果"""
    
    def __init__(self):
        self.issues = []
        self.suggestions = []
        self.statistics = {}
    
    def add_issue(self, issue_type: str, text: str, suggestion: str, 
                 severity: str = "medium", position: Optional[int] = None):
        """添加校对问题"""
        self.issues.append({
            "type": issue_type,
            "text": text,
            "suggestion": suggestion,
            "severity": severity,
            "position": position
        })
    
    def add_suggestion(self, text: str, suggestion: str, reason: str):
        """添加改进建议"""
        self.suggestions.append({
            "original": text,
            "suggested": suggestion,
            "reason": reason
        })


class AIChecker:
    """AI校对检查器"""
    
    def __init__(self, config_or_api_key):
        if isinstance(config_or_api_key, str):
            # 如果传入的是API密钥字符串，创建临时配置
            import os
            os.environ['OPENAI_API_KEY'] = config_or_api_key
            self.config = Config()
        else:
            # 如果传入的是Config对象
            self.config = config_or_api_key
        try:
            self.client = openai.OpenAI(
                api_key=self.config.ai.api_key,
                base_url=self.config.ai.base_url
            )
        except Exception as e:
            # 兼容不同版本的openai库
            openai.api_key = self.config.ai.api_key
            if hasattr(openai, 'api_base'):
                openai.api_base = self.config.ai.base_url
            self.client = None
    
    def check_text(self, text: str) -> ProofreadingResult:
        """检查文本"""
        result = ProofreadingResult()
        
        # 1. 基础检查
        if self.config.rules.check_spelling:
            self._check_spelling(text, result)
        
        if self.config.rules.check_terminology:
            self._check_terminology(text, result)
        
        # 2. AI深度检查
        ai_result = self._ai_proofread(text)
        self._parse_ai_result(ai_result, result)
        
        return result
    
    def _check_spelling(self, text: str, result: ProofreadingResult):
        """检查拼写错误"""
        for correct, typos in self.config.typo_dict.items():
            for typo in typos:
                if typo in text and correct not in text:
                    result.add_issue(
                        issue_type="拼写错误",
                        text=typo,
                        suggestion=f"建议改为：{correct}",
                        severity="high"
                    )
    
    def _check_terminology(self, text: str, result: ProofreadingResult):
        """检查术语一致性"""
        for standard_term, variants in self.config.terminology_dict.items():
            # 检查是否同时出现不同的术语变体
            found_variants = [var for var in variants if var in text]
            if len(found_variants) > 1:
                result.add_issue(
                    issue_type="术语不一致",
                    text=f"发现多种术语：{', '.join(found_variants)}",
                    suggestion=f"建议统一使用：{standard_term}",
                    severity="medium"
                )
    
    def _ai_proofread(self, text: str) -> Dict[str, Any]:
        """使用AI进行深度校对"""
        try:
            prompt = self._build_proofread_prompt(text)
            
            if self.client:
                # 使用新版openai客户端
                response = self.client.chat.completions.create(
                    model=self.config.ai.model,
                    messages=[
                        {
                            "role": "system",
                            "content": "你是一个专业的中文计算机教材校对专家。请仔细检查文本中的语法、用词、逻辑和专业术语问题。"
                        },
                        {
                            "role": "user",
                            "content": prompt
                        }
                    ],
                    max_tokens=self.config.ai.max_tokens,
                    temperature=self.config.ai.temperature
                )
                result_text = response.choices[0].message.content
            else:
                # 使用旧版openai接口
                response = openai.ChatCompletion.create(
                    model=self.config.ai.model,
                    messages=[
                        {
                            "role": "system",
                            "content": "你是一个专业的中文计算机教材校对专家。请仔细检查文本中的语法、用词、逻辑和专业术语问题。"
                        },
                        {
                            "role": "user",
                            "content": prompt
                        }
                    ],
                    max_tokens=self.config.ai.max_tokens,
                    temperature=self.config.ai.temperature
                )
                result_text = response.choices[0].message.content
            
            return self._parse_json_response(result_text)
            
        except Exception as e:
            print(f"AI校对失败: {e}")
            return {"issues": [], "suggestions": []}
    
    def _build_proofread_prompt(self, text: str) -> str:
        """构建校对提示词"""
        return f"""
请对以下中文计算机教材文本进行专业校对，检查以下方面：

1. 语法错误和句式问题
2. 错别字和用词不当
3. 标点符号使用
4. 计算机专业术语的准确性和一致性
5. 逻辑表达和上下文连贯性
6. 格式和排版问题

请按照以下JSON格式返回结果：

{{
    "issues": [
        {{
            "type": "问题类型",
            "text": "问题文本",
            "suggestion": "修改建议",
            "severity": "严重程度(high/medium/low)",
            "reason": "问题原因"
        }}
    ],
    "suggestions": [
        {{
            "original": "原文",
            "suggested": "建议修改",
            "reason": "修改理由"
        }}
    ],
    "overall_assessment": "整体评价"
}}

待校对文本：
{text}
"""
    
    def _parse_json_response(self, response_text: str) -> Dict[str, Any]:
        """解析AI返回的JSON响应"""
        try:
            # 尝试提取JSON部分
            json_start = response_text.find('{')
            json_end = response_text.rfind('}') + 1
            
            if json_start != -1 and json_end != -1:
                json_text = response_text[json_start:json_end]
                return json.loads(json_text)
        except Exception as e:
            print(f"解析AI响应失败: {e}")
        
        return {"issues": [], "suggestions": []}
    
    def _parse_ai_result(self, ai_result: Dict[str, Any], result: ProofreadingResult):
        """解析AI校对结果"""
        # 添加AI发现的问题
        for issue in ai_result.get("issues", []):
            result.add_issue(
                issue_type=issue.get("type", "其他"),
                text=issue.get("text", ""),
                suggestion=issue.get("suggestion", ""),
                severity=issue.get("severity", "medium")
            )
        
        # 添加AI的建议
        for suggestion in ai_result.get("suggestions", []):
            result.add_suggestion(
                text=suggestion.get("original", ""),
                suggestion=suggestion.get("suggested", ""),
                reason=suggestion.get("reason", "")
            )
        
        # 保存整体评价
        if "overall_assessment" in ai_result:
            result.statistics["overall_assessment"] = ai_result["overall_assessment"]
    
    def check_grammar(self, text: str) -> List[Dict[str, str]]:
        """专门检查语法问题"""
        try:
            prompt = f"""
请检查以下中文文本的语法问题，包括：
1. 句子结构问题
2. 主谓宾搭配不当
3. 语序错误
4. 时态不一致
5. 修饰语位置不当

请返回JSON格式的结果：
{{
    "grammar_issues": [
        {{
            "text": "有问题的文本",
            "issue": "问题描述",
            "correction": "修正建议"
        }}
    ]
}}

文本：{text}
"""
            
            if self.client:
                response = self.client.chat.completions.create(
                    model=self.config.ai.model,
                    messages=[
                        {"role": "system", "content": "你是一个中文语法专家。"},
                        {"role": "user", "content": prompt}
                    ],
                    max_tokens=2000,
                    temperature=0.1
                )
                result_text = response.choices[0].message.content
            else:
                response = openai.ChatCompletion.create(
                    model=self.config.ai.model,
                    messages=[
                        {"role": "system", "content": "你是一个中文语法专家。"},
                        {"role": "user", "content": prompt}
                    ],
                    max_tokens=2000,
                    temperature=0.1
                )
                result_text = response.choices[0].message.content
            
            result = self._parse_json_response(result_text)
            return result.get("grammar_issues", [])
            
        except Exception as e:
            print(f"语法检查失败: {e}")
            return []
    
    def check_document(self, text_content: List[str]):
        """检查整个文档"""
        try:
            # 将段落合并为完整文本
            full_text = '\n'.join(text_content)
            
            # 使用AI进行深度校对
            ai_result = self._ai_proofread(full_text)
            
            # 创建校对结果对象
            result = ProofreadingResult()
            
            # 解析AI结果
            self._parse_ai_result(ai_result, result)
            
            # 添加基础检查
            if self.config.rules.check_spelling:
                self._check_spelling(full_text, result)
            
            if self.config.rules.check_terminology:
                self._check_terminology(full_text, result)
            
            return result
            
        except Exception as e:
            print(f"文档检查失败: {e}")
            # 返回空结果
            result = ProofreadingResult()
            return result 