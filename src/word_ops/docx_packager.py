#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
DOCX打包器 - 统一的DOCX解包-修改-重打包逻辑
"""

import os
import zipfile
import tempfile
from logging_config import get_logger

logger = get_logger("word_ops.docx_packager")


def modify_docx_package(input_path: str, output_path: str, modifier_fn) -> bool:
    """
    解压DOCX → 调用modifier_fn(temp_dir) → 重新打包

    Args:
        input_path: 输入DOCX文件路径
        output_path: 输出DOCX文件路径
        modifier_fn: 修改函数，接收temp_dir参数，返回bool表示是否成功

    Returns:
        bool: 是否成功
    """
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            # 解压docx文件
            with zipfile.ZipFile(input_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            # 调用修改函数
            if not modifier_fn(temp_dir):
                return False

            # 重新打包为docx
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arc_name = os.path.relpath(file_path, temp_dir)
                        zip_ref.write(file_path, arc_name)

            return True

    except Exception as e:
        logger.error(f"DOCX打包操作失败: {e}")
        return False
