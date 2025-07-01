#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
pytest配置文件
"""

import pytest
import os
import tempfile
import shutil


@pytest.fixture
def temp_dir():
    """创建临时目录的fixture"""
    temp_path = tempfile.mkdtemp()
    yield temp_path
    shutil.rmtree(temp_path)


@pytest.fixture
def sample_text():
    """提供示例文本的fixture"""
    return """
    这是一个示例文档，用于测试AI校对功能。
    文档中可能包含一些语法错误、拼写错误或者逻辑问题。
    校对系统应能够识别并提供修改建议。
    """


@pytest.fixture
def mock_api_key():
    """提供模拟API密钥的fixture"""
    return "test_openai_api_key"


@pytest.fixture
def sample_config():
    """提供示例配置的fixture"""
    return {
        'openai_api_key': 'test_key',
        'model_name': 'gpt-3.5-turbo',
        'max_tokens': 2000,
        'temperature': 0.7
    } 