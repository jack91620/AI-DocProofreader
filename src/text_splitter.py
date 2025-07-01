#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
智能文本分割器
用于将长文档分割成适合AI模型处理的小块
"""

import re
import math
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass


@dataclass
class TextChunk:
    """文本块数据类"""
    content: str
    start_paragraph: int
    end_paragraph: int
    token_count: int
    chunk_index: int


class TokenCounter:
    """Token计数器 - 改进版，更准确地估算实际API调用消耗"""
    
    # 系统提示词的大致token数（更保守的估算）
    SYSTEM_PROMPT_TOKENS = 300
    
    # 用户提示模板的大致token数（包含大量JSON格式说明等）
    USER_PROMPT_TEMPLATE_TOKENS = 3500
    
    # JSON响应格式的预期token数
    RESPONSE_FORMAT_TOKENS = 800
    
    # 其他API开销（请求结构、格式化等）
    API_OVERHEAD_TOKENS = 500
    
    @staticmethod
    def estimate_tokens(text: str) -> int:
        """
        估算文本的token数量 - 非常保守的估算
        对于中文，约1个字符 = 1个token（非常保守）
        对于英文，约1个单词 = 1个token
        """
        if not text:
            return 0
        
        # 计算中文字符数量
        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
        # 计算英文单词数量
        english_words = len(re.findall(r'[a-zA-Z]+', text))
        # 计算其他字符（数字、标点、空格等）
        other_chars = len(text) - chinese_chars - sum(len(word) for word in re.findall(r'[a-zA-Z]+', text))
        
        # 非常保守的估算token数量
        estimated_tokens = (
            chinese_chars +          # 中文1个字符约1个token（非常保守）
            english_words +          # 英文单词约1个单词1个token
            other_chars // 2         # 其他字符约2个字符1个token（更保守）
        )
        
        # 至少保证不少于字符数/2（极保守的下限）
        return max(int(estimated_tokens), len(text) // 2)
    
    @classmethod
    def estimate_full_request_tokens(cls, content_text: str) -> int:
        """
        估算完整API请求的token消耗
        包括系统提示、用户提示模板、内容文本和响应格式
        """
        content_tokens = cls.estimate_tokens(content_text)
        
        total_tokens = (
            cls.SYSTEM_PROMPT_TOKENS +        # 系统提示词
            cls.USER_PROMPT_TEMPLATE_TOKENS + # 用户提示模板
            content_tokens +                   # 实际内容
            cls.RESPONSE_FORMAT_TOKENS +      # 预期响应
            cls.API_OVERHEAD_TOKENS           # API开销
        )
        
        return total_tokens


class SmartTextSplitter:
    """智能文本分割器 - 改进版，使用更准确的token估算"""
    
    def __init__(self, max_tokens: int = 6000, overlap_tokens: int = 200):
        """
        初始化分割器
        
        Args:
            max_tokens: 每个文本块的最大token数（这是模型的总限制）
            overlap_tokens: 块之间的重叠token数，用于保持上下文连贯性
        """
        self.max_tokens = max_tokens
        self.overlap_tokens = overlap_tokens
        self.token_counter = TokenCounter()
        
        # 计算内容文本的实际可用token数
        # 需要减去系统提示、用户模板、响应和API开销
        overhead_tokens = (
            TokenCounter.SYSTEM_PROMPT_TOKENS +
            TokenCounter.USER_PROMPT_TEMPLATE_TOKENS +
            TokenCounter.RESPONSE_FORMAT_TOKENS +
            TokenCounter.API_OVERHEAD_TOKENS
        )
        
        # 为内容留出的实际token数，再减去更大的安全边距
        self.content_max_tokens = max(500, max_tokens - overhead_tokens - 1000)  # 增加安全边距到1000
        
        print(f"[DEBUG] 模型总限制: {max_tokens} tokens")
        print(f"[DEBUG] API开销估算: {overhead_tokens} tokens") 
        print(f"[DEBUG] 安全边距: 1000 tokens")
        print(f"[DEBUG] 内容可用token: {self.content_max_tokens} tokens")
    
    def split_by_paragraphs(self, paragraphs: List[str]) -> List[TextChunk]:
        """
        按段落智能分割文本
        
        Args:
            paragraphs: 段落列表
            
        Returns:
            分割后的文本块列表
        """
        chunks = []
        current_chunk_content = ""
        current_chunk_start = 0
        current_paragraph_index = 0
        chunk_index = 0
        
        for i, paragraph in enumerate(paragraphs):
            # 使用内容token限制而不是总限制
            paragraph_tokens = self.token_counter.estimate_tokens(paragraph)
            current_tokens = self.token_counter.estimate_tokens(current_chunk_content)
            
            # 如果单个段落就超过内容token限制，需要进一步分割
            if paragraph_tokens > self.content_max_tokens:
                # 先保存当前块（如果有内容）
                if current_chunk_content.strip():
                    # 计算完整请求的token数用于记录
                    full_request_tokens = self.token_counter.estimate_full_request_tokens(current_chunk_content)
                    chunks.append(TextChunk(
                        content=current_chunk_content.strip(),
                        start_paragraph=current_chunk_start,
                        end_paragraph=i-1,
                        token_count=full_request_tokens,  # 记录完整请求的token数
                        chunk_index=chunk_index
                    ))
                    chunk_index += 1
                
                # 分割超长段落
                long_paragraph_chunks = self._split_long_paragraph(
                    paragraph, i, chunk_index
                )
                chunks.extend(long_paragraph_chunks)
                chunk_index += len(long_paragraph_chunks)
                
                # 重置当前块
                current_chunk_content = ""
                current_chunk_start = i + 1
                current_paragraph_index = i + 1
                
            # 如果添加这个段落会超过内容限制
            elif current_tokens + paragraph_tokens > self.content_max_tokens:
                # 保存当前块
                if current_chunk_content.strip():
                    full_request_tokens = self.token_counter.estimate_full_request_tokens(current_chunk_content)
                    chunks.append(TextChunk(
                        content=current_chunk_content.strip(),
                        start_paragraph=current_chunk_start,
                        end_paragraph=i-1,
                        token_count=full_request_tokens,
                        chunk_index=chunk_index
                    ))
                    chunk_index += 1
                
                # 开始新块，可能包含一些重叠内容
                overlap_content = self._get_overlap_content(paragraphs, i, current_chunk_start)
                current_chunk_content = overlap_content + paragraph
                current_chunk_start = max(0, i - 1) if overlap_content else i
                
            else:
                # 添加段落到当前块
                if current_chunk_content:
                    current_chunk_content += "\n" + paragraph
                else:
                    current_chunk_content = paragraph
                    current_chunk_start = i
        
        # 添加最后一个块
        if current_chunk_content.strip():
            final_full_tokens = self.token_counter.estimate_full_request_tokens(current_chunk_content)
            chunks.append(TextChunk(
                content=current_chunk_content.strip(),
                start_paragraph=current_chunk_start,
                end_paragraph=len(paragraphs)-1,
                token_count=final_full_tokens,
                chunk_index=chunk_index
            ))
        
        return chunks
    
    def _split_long_paragraph(self, paragraph: str, paragraph_index: int, start_chunk_index: int) -> List[TextChunk]:
        """分割超长段落"""
        chunks = []
        sentences = self._split_into_sentences(paragraph)
        current_content = ""
        chunk_index = start_chunk_index
        
        for sentence in sentences:
            sentence_tokens = self.token_counter.estimate_tokens(sentence)
            current_tokens = self.token_counter.estimate_tokens(current_content)
            
            if current_tokens + sentence_tokens > self.content_max_tokens and current_content:
                # 保存当前块
                full_request_tokens = self.token_counter.estimate_full_request_tokens(current_content)
                chunks.append(TextChunk(
                    content=current_content.strip(),
                    start_paragraph=paragraph_index,
                    end_paragraph=paragraph_index,
                    token_count=full_request_tokens,
                    chunk_index=chunk_index
                ))
                chunk_index += 1
                current_content = sentence
            else:
                current_content += sentence
        
        # 添加最后一块
        if current_content.strip():
            final_full_tokens = self.token_counter.estimate_full_request_tokens(current_content)
            chunks.append(TextChunk(
                content=current_content.strip(),
                start_paragraph=paragraph_index,
                end_paragraph=paragraph_index,
                token_count=final_full_tokens,
                chunk_index=chunk_index
            ))
        
        return chunks
    
    def _split_into_sentences(self, text: str) -> List[str]:
        """将文本分割成句子"""
        # 中文句子分割符
        sentence_endings = r'[。！？；]'
        sentences = re.split(sentence_endings, text)
        
        # 重新添加标点符号
        result = []
        for i, sentence in enumerate(sentences):
            if sentence.strip():
                if i < len(sentences) - 1:  # 不是最后一个
                    # 查找对应的标点符号
                    original_pos = sum(len(s) for s in sentences[:i+1]) + i
                    if original_pos < len(text):
                        punctuation = text[original_pos]
                        result.append(sentence + punctuation)
                    else:
                        result.append(sentence)
                else:
                    result.append(sentence)
        
        return result
    
    def _get_overlap_content(self, paragraphs: List[str], current_index: int, chunk_start: int) -> str:
        """获取重叠内容以保持上下文连贯性"""
        if self.overlap_tokens <= 0 or current_index <= chunk_start:
            return ""
        
        overlap_content = ""
        overlap_tokens = 0
        
        # 从当前位置向前查找，直到达到重叠token限制
        for i in range(current_index - 1, chunk_start - 1, -1):
            paragraph_tokens = self.token_counter.estimate_tokens(paragraphs[i])
            if overlap_tokens + paragraph_tokens <= self.overlap_tokens:
                overlap_content = paragraphs[i] + "\n" + overlap_content
                overlap_tokens += paragraph_tokens
            else:
                break
        
        return overlap_content.strip()
    
    def get_split_summary(self, chunks: List[TextChunk]) -> Dict:
        """获取分割摘要信息"""
        if not chunks:
            return {
                "total_chunks": 0,
                "total_tokens": 0,
                "average_tokens_per_chunk": 0,
                "max_tokens_per_chunk": 0,
                "min_tokens_per_chunk": 0
            }
        
        total_tokens = sum(chunk.token_count for chunk in chunks)
        token_counts = [chunk.token_count for chunk in chunks]
        
        return {
            "total_chunks": len(chunks),
            "total_tokens": total_tokens,
            "average_tokens_per_chunk": total_tokens // len(chunks),
            "max_tokens_per_chunk": max(token_counts),
            "min_tokens_per_chunk": min(token_counts),
            "chunks_info": [
                {
                    "chunk_index": chunk.chunk_index,
                    "paragraphs": f"{chunk.start_paragraph}-{chunk.end_paragraph}",
                    "tokens": chunk.token_count,
                    "content_preview": chunk.content[:100] + "..." if len(chunk.content) > 100 else chunk.content
                }
                for chunk in chunks
            ]
        }


def test_improved_text_splitter():
    """测试改进的文本分割器"""
    # 创建测试段落
    test_paragraphs = [
        "第一章 计算机基础知识。计算机科学是一门研究设计计算系统以及利用计算系统解决问题方法的学科。",
        "计算机科学包含很多分支领域，有些强调特定结果的计算，比如计算机图形学；而有些是探讨计算问题的性质。",
        "程序设计是计算机科学的核心技能之一。良好的程序设计能力需要掌握多种编程语言和开发工具。",
        "数据结构与算法是程序设计的基础。常见的数据结构包括数组、链表、栈、队列、树、图等。",
        "软件工程强调系统化的软件开发方法。它包括需求分析、系统设计、编码实现、测试验证等阶段。"
    ]
    
    print("=== 改进的Token计数器测试 ===")
    counter = TokenCounter()
    
    test_content = "\n".join(test_paragraphs)
    content_tokens = counter.estimate_tokens(test_content)
    full_request_tokens = counter.estimate_full_request_tokens(test_content)
    
    print(f"内容文本: {test_content[:100]}...")
    print(f"字符数: {len(test_content)}")
    print(f"内容token估算: {content_tokens}")
    print(f"完整请求token估算: {full_request_tokens}")
    print(f"API开销: {full_request_tokens - content_tokens}")
    
    print("\n=== 改进的文本分割器测试 ===")
    
    # 模拟真实的gpt-4限制
    splitter = SmartTextSplitter(max_tokens=8192, overlap_tokens=100)
    chunks = splitter.split_by_paragraphs(test_paragraphs)
    
    summary = splitter.get_split_summary(chunks)
    print(f"总块数: {summary['total_chunks']}")
    print(f"总token数: {summary['total_tokens']}")
    print(f"平均每块: {summary['average_tokens_per_chunk']} tokens")
    print(f"最大块: {summary['max_tokens_per_chunk']} tokens")
    
    for i, chunk in enumerate(chunks):
        print(f"\n--- 块 {i+1} ---")
        print(f"段落范围: {chunk.start_paragraph}-{chunk.end_paragraph}")
        print(f"完整请求Token数: {chunk.token_count}")
        print(f"内容预览: {chunk.content[:100]}...")


if __name__ == "__main__":
    test_improved_text_splitter() 