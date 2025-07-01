"""
AI校对引擎模块
"""

import json
import re
from typing import List, Dict, Any, Optional
import openai
from config import Config
from text_splitter import SmartTextSplitter, TextChunk
from rich.console import Console
from rich.progress import Progress, TaskID


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
        
        # 初始化文本分割器 - 使用改进的token计算
        # 使用模型的实际token限制，分割器内部会计算API开销
        self.text_splitter = SmartTextSplitter(
            max_tokens=self.config.ai.max_tokens,  # 使用模型的真实限制
            overlap_tokens=self.config.ai.chunk_overlap_tokens
        )
        self.console = Console()
    
    def check_text(self, text: str) -> ProofreadingResult:
        """检查文本 - 支持自动分块处理长文档"""
        # 使用改进的token估算
        estimated_tokens = self.text_splitter.token_counter.estimate_full_request_tokens(text)
        safe_limit = self.config.ai.max_tokens
        
        if estimated_tokens <= safe_limit:
            # 文本不长，直接处理
            return self._check_single_text(text)
        else:
            # 文本太长，需要分块处理
            self.console.print(f"[yellow]⚠️ 文本过长 (估算{estimated_tokens} tokens > {safe_limit} tokens)，启用分块处理[/yellow]")
            return self._check_text_chunks(text)
    
    def check_text_by_paragraphs(self, paragraphs: List[str]) -> ProofreadingResult:
        """按段落检查文本 - 支持自动分块处理"""
        # 使用改进的token估算
        total_text = '\n'.join(paragraphs)
        estimated_tokens = self.text_splitter.token_counter.estimate_full_request_tokens(total_text)
        safe_limit = self.config.ai.max_tokens
        
        if estimated_tokens <= safe_limit:
            # 文本不长，直接处理
            return self._check_single_text(total_text)
        else:
            # 文本太长，按段落分块处理
            self.console.print(f"[yellow]⚠️ 文档过长 (估算{estimated_tokens} tokens > {safe_limit} tokens)，启用分块处理[/yellow]")
            return self._check_paragraphs_chunks(paragraphs)
    
    def _check_single_text(self, text: str) -> ProofreadingResult:
        """检查单个文本块（原有逻辑）"""
        result = ProofreadingResult()
        
        # 1. 基础检查
        if self.config.rules.check_spelling:
            self._check_spelling(text, result)
        
        if self.config.rules.check_terminology:
            try:
                terminology_issues = self._check_terminology(text)
                if isinstance(terminology_issues, list):
                    for issue in terminology_issues:
                        if isinstance(issue, dict):
                            result.add_issue(
                                issue_type=issue.get("type", "术语"),
                                text=issue.get("original", ""),
                                suggestion=issue.get("suggested", ""),
                                severity="medium"
                            )
                        else:
                            # 如果issue不是字典，跳过或记录错误
                            print(f"⚠️ 术语检查返回了无效格式: {type(issue)}")
            except Exception as e:
                print(f"⚠️ 术语检查处理失败: {e}")
        
        # 2. AI深度检查
        ai_result = self._ai_proofread(text)
        self._parse_ai_result(ai_result, result)
        
        return result
    
    def _check_text_chunks(self, text: str) -> ProofreadingResult:
        """分块检查长文本"""
        # 将文本按行分割为段落
        paragraphs = [line.strip() for line in text.split('\n') if line.strip()]
        return self._check_paragraphs_chunks(paragraphs)
    
    def _check_paragraphs_chunks(self, paragraphs: List[str]) -> ProofreadingResult:
        """按段落分块检查文档"""
        # 分割文本
        chunks = self.text_splitter.split_by_paragraphs(paragraphs)
        
        # 显示分割摘要
        summary = self.text_splitter.get_split_summary(chunks)
        self.console.print(f"[blue]📄 文档分析结果：[/blue]")
        self.console.print(f"   - 总块数: {summary['total_chunks']}")
        self.console.print(f"   - 总token数: {summary['total_tokens']}")
        self.console.print(f"   - 平均每块: {summary['average_tokens_per_chunk']} tokens")
        self.console.print(f"   - 最大块: {summary['max_tokens_per_chunk']} tokens")
        
        # 创建总结果对象
        final_result = ProofreadingResult()
        
        # 使用进度条处理各个块
        with Progress() as progress:
            task = progress.add_task(
                "[green]处理文档块...", 
                total=len(chunks)
            )
            
            for i, chunk in enumerate(chunks):
                progress.update(
                    task, 
                    advance=1, 
                    description=f"[green]处理块 {i+1}/{len(chunks)} ({chunk.token_count} tokens)..."
                )
                
                # 处理当前块
                chunk_result = self._check_single_text(chunk.content)
                
                # 合并结果，调整位置信息
                self._merge_chunk_result(final_result, chunk_result, chunk)
        
        # 显示最终统计
        self.console.print(f"[green]✅ 分块处理完成！[/green]")
        self.console.print(f"   - 发现问题: {len(final_result.issues)} 个")
        self.console.print(f"   - 改进建议: {len(final_result.suggestions)} 个")
        
        return final_result
    
    def _merge_chunk_result(self, final_result: ProofreadingResult, chunk_result: ProofreadingResult, chunk: TextChunk):
        """合并块处理结果到最终结果"""
        # 合并问题，调整段落位置
        for issue in chunk_result.issues:
            # 为块中的问题添加块信息
            issue_copy = issue.copy()
            issue_copy['chunk_info'] = {
                'chunk_index': chunk.chunk_index,
                'start_paragraph': chunk.start_paragraph,
                'end_paragraph': chunk.end_paragraph
            }
            final_result.issues.append(issue_copy)
        
        # 合并建议
        for suggestion in chunk_result.suggestions:
            # 为块中的建议添加块信息
            suggestion_copy = suggestion.copy()
            suggestion_copy['chunk_info'] = {
                'chunk_index': chunk.chunk_index,
                'start_paragraph': chunk.start_paragraph,
                'end_paragraph': chunk.end_paragraph
            }
            final_result.suggestions.append(suggestion_copy)
        
        # 合并统计信息
        chunk_stats = chunk_result.statistics
        if chunk_stats:
            if 'chunk_statistics' not in final_result.statistics:
                final_result.statistics['chunk_statistics'] = []
            
            final_result.statistics['chunk_statistics'].append({
                'chunk_index': chunk.chunk_index,
                'token_count': chunk.token_count,
                'paragraph_range': f"{chunk.start_paragraph}-{chunk.end_paragraph}",
                'issues_count': len(chunk_result.issues),
                'suggestions_count': len(chunk_result.suggestions),
                'stats': chunk_stats
            })
    
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
    
    def _check_terminology(self, text: str) -> List[Dict]:
        """检查术语一致性 - 带重试机制"""
        max_retries = 3
        base_delay = 1
        
        for attempt in range(max_retries):
            try:
                # 术语检查的回复通常较短，设置较小的max_tokens
                max_completion_tokens = 800  # 减少token数，降低服务器负担
                
                # 简化提示词，减少token消耗
                prompt = f"""
检查以下文本的术语一致性问题：

文本：{text[:1000]}...

返回JSON格式：{{"problems": [{{"type": "terminology", "original": "原文", "suggested": "建议", "reason": "原因"}}]}}
"""
                
                if self.client:
                    response = self.client.chat.completions.create(
                        model=self.config.ai.model,
                        messages=[
                            {"role": "system", "content": "你是术语检查专家，检查技术文档术语一致性。"},
                            {"role": "user", "content": prompt}
                        ],
                        max_tokens=max_completion_tokens,
                        temperature=0.1,
                        timeout=30  # 添加超时设置
                    )
                    content = response.choices[0].message.content
                else:
                    response = openai.ChatCompletion.create(
                        model=self.config.ai.model,
                        messages=[
                            {"role": "system", "content": "你是术语检查专家，检查技术文档术语一致性。"},
                            {"role": "user", "content": prompt}
                        ],
                        max_tokens=max_completion_tokens,
                        temperature=0.1,
                        request_timeout=30
                    )
                    content = response.choices[0].message.content
                
                # 解析AI返回的结果
                result = self._parse_json_response(content)
                
                # 确保返回正确格式的列表
                if isinstance(result, dict) and "problems" in result:
                    problems = result["problems"]
                    if isinstance(problems, list):
                        return problems
                
                # 如果解析失败，返回空列表
                return []
                
            except Exception as e:
                error_msg = str(e).lower()
                
                # 检查是否是服务器错误或超时
                if any(keyword in error_msg for keyword in ['server error', '500', 'timeout', 'connection']):
                    if attempt < max_retries - 1:
                        delay = base_delay * (2 ** attempt)  # 指数退避
                        self.console.print(f"[yellow]⚠️ 术语检查遇到服务器错误，{delay}秒后重试 (尝试 {attempt + 1}/{max_retries})[/yellow]")
                        import time
                        time.sleep(delay)
                        continue
                    else:
                        self.console.print(f"[yellow]⚠️ 术语检查最终失败，跳过此检查: {e}[/yellow]")
                        return []
                else:
                    # 其他错误直接失败
                    self.console.print(f"[yellow]⚠️ 术语检查失败: {e}[/yellow]")
                    return []
        
        # 如果所有重试都失败，返回空列表
        return []
    
    def _ai_proofread(self, text: str) -> Dict[str, Any]:
        """AI深度校对"""
        try:
            prompt = self._build_proofread_prompt(text)
            
            # 计算合理的max_tokens：避免输入+输出超过模型总限制
            # 估算输入的token数（包括系统提示和用户提示）
            input_tokens = self.text_splitter.token_counter.estimate_full_request_tokens(text)
            # 为输出保留合理的token数，通常校对结果不会比原文长太多
            max_completion_tokens = min(2000, max(500, self.config.ai.max_tokens - input_tokens - 500))
            
            if max_completion_tokens < 500:
                # 如果计算出的完成tokens太少，说明输入太长，应该进行分块
                print(f"⚠️ 输入过长 ({input_tokens} tokens)，建议使用分块处理")
                max_completion_tokens = 500
            
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
                    max_tokens=max_completion_tokens,  # 使用计算出的合理完成长度
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
                    max_tokens=max_completion_tokens,  # 使用计算出的合理完成长度
                    temperature=self.config.ai.temperature
                )
                result_text = response.choices[0].message.content
            
            return self._parse_json_response(result_text)
            
        except Exception as e:
            print(f"AI校对失败: {e}")
            return {"issues": [], "suggestions": []}
    
    def _build_proofread_prompt(self, text: str) -> str:
        """构建校对提示词 - 优化JSON输出格式"""
        return f"""
你是一个专业的中文计算机教材校对专家。请对以下文本进行校对，检查语法、用词、术语和逻辑问题。

**重要：请严格按照以下JSON格式返回结果，不要添加任何其他文字或说明**

{{
  "issues": [
    {{
      "type": "问题类型",
      "text": "问题文本",
      "suggestion": "修改建议",
      "severity": "medium",
      "reason": "问题原因"
    }}
  ],
  "suggestions": [
    {{
      "original": "原文",
      "suggested": "建议修改",
      "reason": "修改理由"
    }}
  ]
}}

**检查要点：**
1. 错别字和用词不当
2. 语法错误和句式问题  
3. 计算机术语的准确性
4. 标点符号使用
5. 逻辑表达问题

**待校对文本：**
{text}

**请只返回上述JSON格式的结果，不要包含其他内容。**
"""
    
    def _parse_json_response(self, response_text: str) -> Dict[str, Any]:
        """解析AI返回的JSON响应 - 超强容错版本"""
        try:
            # 先尝试清理和预处理响应文本
            cleaned_text = self._clean_ai_response(response_text)
            
            # 尝试提取JSON部分
            json_start = cleaned_text.find('{')
            json_end = cleaned_text.rfind('}') + 1
            
            if json_start != -1 and json_end != -1:
                json_text = cleaned_text[json_start:json_end]
                
                # 首先尝试直接解析
                try:
                    return json.loads(json_text)
                except json.JSONDecodeError as e:
                    # 尝试多种修复方法
                    for fix_method in [
                        self._fix_common_json_issues,
                        self._fix_nested_json_issues,
                        self._fix_quote_issues,
                        self._fix_aggressive_json_issues  # 新增激进修复
                    ]:
                        try:
                            fixed_json = fix_method(json_text)
                            return json.loads(fixed_json)
                        except json.JSONDecodeError:
                            continue
                    
                    # 如果所有修复都失败，尝试智能文本解析
                    self.console.print("[yellow]JSON修复失败，使用智能文本解析[/yellow]")
                    return self._parse_text_fallback_smart(response_text)
                        
        except Exception as e:
            self.console.print(f"[yellow]解析AI响应失败: {e}[/yellow]")
        
        return {"issues": [], "suggestions": []}
    
    def _clean_ai_response(self, response_text: str) -> str:
        """清理AI响应文本"""
        import re
        
        # 移除markdown代码块标记
        response_text = re.sub(r'```json\s*', '', response_text)
        response_text = re.sub(r'```\s*$', '', response_text)
        
        # 移除多余的空行和空格
        response_text = re.sub(r'\n\s*\n', '\n', response_text)
        
        # 移除AI可能添加的说明文字
        lines = response_text.split('\n')
        cleaned_lines = []
        json_started = False
        
        for line in lines:
            line = line.strip()
            if line.startswith('{') or json_started:
                json_started = True
                cleaned_lines.append(line)
            elif line.endswith('}') and json_started:
                cleaned_lines.append(line)
                break
        
        return '\n'.join(cleaned_lines)
    
    def _fix_nested_json_issues(self, json_text: str) -> str:
        """修复嵌套JSON的问题"""
        import re
        
        # 修复嵌套对象中缺失的逗号
        json_text = re.sub(r'}\s*{', '},{', json_text)
        
        # 修复数组中缺失的逗号
        json_text = re.sub(r']\s*\[', '],[', json_text)
        
        # 修复对象和数组混合的问题
        json_text = re.sub(r'}\s*\[', '},[', json_text)
        json_text = re.sub(r']\s*{', '],{', json_text)
        
        return json_text
    
    def _fix_quote_issues(self, json_text: str) -> str:
        """修复引号问题"""
        import re
        
        # 修复单引号为双引号
        json_text = json_text.replace("'", '"')
        
        # 修复值中包含换行符的问题
        json_text = re.sub(r'"([^"]*?)\n([^"]*?)"', r'"\1 \2"', json_text)
        
        # 修复值中包含特殊字符的问题
        json_text = re.sub(r'"([^"]*?)"([^,}\]]*?)"', r'"\1\2"', json_text)
        
        return json_text
    
    def _fix_common_json_issues(self, json_text: str) -> str:
        """修复常见的JSON格式问题"""
        import re
        
        # 移除多余的逗号（在}或]前的逗号）
        json_text = re.sub(r',(\s*[}\]])', r'\1', json_text)
        
        # 修复缺失的逗号（在"后面直接跟"的情况）
        json_text = re.sub(r'"\s*\n\s*"', '",\n"', json_text)
        
        # 修复缺失的逗号（在}后面直接跟{的情况）
        json_text = re.sub(r'}\s*\n\s*{', '},\n{', json_text)
        
        # 移除JSON中的注释
        json_text = re.sub(r'//.*?\n', '\n', json_text)
        
        # 确保字符串值被正确引用
        json_text = re.sub(r':\s*([^",{}\[\]\n]+)(?=\s*[,}\]])', r': "\1"', json_text)
        
        return json_text
    
    def _fix_aggressive_json_issues(self, json_text: str) -> str:
        """激进的JSON修复方法"""
        import re
        
        # 移除所有换行符并重新格式化
        json_text = re.sub(r'\s*\n\s*', ' ', json_text)
        
        # 修复属性名没有引号的问题
        json_text = re.sub(r'(\w+):', r'"\1":', json_text)
        
        # 修复值没有引号的问题（排除数字、布尔值、null）
        json_text = re.sub(r':\s*([^",{\[\]}\s][^,}\]]*?)(?=[,}\]])', r': "\1"', json_text)
        
        # 修复多余的逗号
        json_text = re.sub(r',\s*([}\]])', r'\1', json_text)
        
        # 修复缺失的逗号
        json_text = re.sub(r'([}\]"])\s*([{\[""])', r'\1,\2', json_text)
        
        # 修复嵌套问题
        json_text = re.sub(r'}\s*{', '},{', json_text)
        
        return json_text
    
    def _parse_text_fallback_smart(self, response_text: str) -> Dict[str, Any]:
        """智能文本解析 - 改进版"""
        result = {"issues": [], "suggestions": []}
        
        try:
            # 使用正则表达式提取结构化信息
            import re
            
            # 查找可能的问题描述
            issue_patterns = [
                r'(?:问题|错误|error|issue)[:：]\s*([^\n]+)',
                r'(?:原文|原始|original)[:：]\s*([^\n]+)',
                r'(?:建议|修改|suggested?)[:：]\s*([^\n]+)',
                r'(?:类型|type)[:：]\s*([^\n]+)',
            ]
            
            suggestions = []
            issues = []
            
            # 尝试从原始文本中提取修改建议
            lines = response_text.split('\n')
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # 查找 "原文 -> 修改" 模式
                arrow_match = re.search(r'(.+?)\s*[→->]\s*(.+)', line)
                if arrow_match:
                    original = arrow_match.group(1).strip()
                    suggested = arrow_match.group(2).strip()
                    
                    # 清理引号和标点
                    original = re.sub(r'^["\'"]*|["\'"]*$', '', original)
                    suggested = re.sub(r'^["\'"]*|["\'"]*$', '', suggested)
                    
                    if original and suggested and original != suggested:
                        suggestions.append({
                            "original": original,
                            "suggested": suggested,
                            "reason": "AI自动解析的建议"
                        })
                        continue
                
                # 查找问题描述
                for pattern in issue_patterns:
                    match = re.search(pattern, line, re.IGNORECASE)
                    if match:
                        text = match.group(1).strip()
                        if text:
                            issues.append({
                                "type": "解析问题",
                                "text": text,
                                "suggestion": "请参考原文进行修改",
                                "severity": "medium"
                            })
                        break
            
            # 如果没有找到结构化信息，尝试提取关键词
            if not suggestions and not issues:
                # 查找可能的术语或修改建议
                potential_corrections = re.findall(r'([^\s]+)\s*[→->]\s*([^\s]+)', response_text)
                for orig, sugg in potential_corrections:
                    if orig != sugg and len(orig) > 1 and len(sugg) > 1:
                        suggestions.append({
                            "original": orig,
                            "suggested": sugg,
                            "reason": "文本解析提取的建议"
                        })
            
            result["suggestions"] = suggestions
            result["issues"] = issues
            
            # 记录解析结果
            if suggestions or issues:
                self.console.print(f"[green]✅ 智能文本解析成功: {len(suggestions)} 个建议, {len(issues)} 个问题[/green]")
            else:
                self.console.print("[yellow]⚠️ 智能文本解析未找到有效信息[/yellow]")
                
        except Exception as e:
            self.console.print(f"[yellow]智能文本解析失败: {e}[/yellow]")
        
        return result
    
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
                terminology_issues = self._check_terminology(full_text)
                for issue in terminology_issues:
                    if isinstance(issue, dict):
                        result.add_issue(
                            issue_type=issue.get("type", "术语"),
                            text=issue.get("original", ""),
                            suggestion=issue.get("suggested", ""),
                            severity="medium"
                        )
            
            return result
            
        except Exception as e:
            print(f"文档检查失败: {e}")
            # 返回空结果
            return ProofreadingResult() 