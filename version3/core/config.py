#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置文件 - 仅保留 Qwen 配置
"""

import os

# VLM API 配置（只保留 Qwen）
VLM_CONFIG = {
    "qwen": {
        "api_key": "sk-feb52aae598c4043865273ada1635897",
        "model": "qwen3.5-plus",  # 改为最新模型
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "max_tokens": 4096,
        "temperature": 0.2,
        "timeout": 300,
        "max_retries": 3
    }
}

# 处理逻辑配置（保留，但可能不被使用）
PROCESSING_CONFIG = {
    "similarity_threshold": 0.65
}

# 路径配置（保留）
PATHS = {
    "input_folder": "input",
    "output_folder": "output",
    "images_folder": "images",
    "prompts_folder": "prompts"
}

def get_api_key(provider: str) -> str:
    """获取指定提供商的 API 密钥，优先级：环境变量 > 配置文件"""
    if provider == "qwen":
        api_key = os.getenv("DASHSCOPE_API_KEY")
        if api_key:
            return api_key
        return VLM_CONFIG["qwen"]["api_key"]
    return ""