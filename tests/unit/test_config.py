"""Config模块单元测试"""

import os
import sys
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', 'src'))

from config import Config, AIConfig, ProofreadingRules, CommentStyle


class TestAIConfig:
    """AIConfig配置测试"""

    def test_default_values(self):
        config = AIConfig(api_key="test-key")
        assert config.api_key == "test-key"
        assert config.model == "gpt-4"
        assert config.temperature == 0.1
        assert config.chunk_overlap_tokens == 200

    def test_custom_values(self):
        config = AIConfig(api_key="key", model="gpt-4o", max_tokens=16000)
        assert config.model == "gpt-4o"
        assert config.max_tokens == 16000


class TestProofreadingRules:
    """校对规则测试"""

    def test_defaults_all_enabled(self):
        rules = ProofreadingRules()
        assert rules.check_grammar is True
        assert rules.check_spelling is True
        assert rules.check_terminology is True
        assert rules.check_punctuation is True
        assert rules.check_consistency is True


class TestCommentStyle:
    """批注样式测试"""

    def test_default_author(self):
        style = CommentStyle()
        assert style.author == "AI校对助手"


class TestConfig:
    """主配置类测试"""

    def test_config_creates_with_defaults(self):
        config = Config()
        assert config.ai is not None
        assert config.rules is not None
        assert config.comment_style is not None

    def test_terminology_dict_not_empty(self):
        config = Config()
        assert len(config.terminology_dict) > 0

    def test_typo_dict_not_empty(self):
        config = Config()
        assert len(config.typo_dict) > 0

    def test_validate_fails_without_api_key(self):
        config = Config()
        config.ai.api_key = ""
        with pytest.raises(ValueError, match="API Key"):
            config.validate()

    def test_validate_fails_with_bad_overlap(self):
        config = Config()
        config.ai.api_key = "test-key"
        config.ai.chunk_overlap_tokens = 99999
        config.ai.max_tokens = 100
        with pytest.raises(ValueError, match="重叠token数"):
            config.validate()

    def test_validate_passes_with_valid_config(self):
        config = Config()
        config.ai.api_key = "test-key"
        assert config.validate() is True
