#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
统一的 VLM 客户端 - 精简版
仅支持 Qwen VL 和 Mock
"""

import json
import base64
import os
import time
import re
import requests
from typing import List, Dict

try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

try:
    from core.config import get_api_key, VLM_CONFIG
except ImportError:
    get_api_key = None
    VLM_CONFIG = None


class BaseVLMClient:
    """VLM 客户端基类（仅保留单图分析核心）"""
    def __init__(self, api_key: str, model: str, base_url: str, vlm_type: str = "unknown"):
        self.api_key = api_key
        self.model = model
        self.base_url = base_url
        self.vlm_type = vlm_type
        self.client = None
        if OPENAI_AVAILABLE:
            self.client = OpenAI(api_key=self.api_key, base_url=self.base_url)

    def encode_image(self, image_path: str) -> str:
        with open(image_path, "rb") as f:
            return base64.b64encode(f.read()).decode('utf-8')

    def analyze_image(self, image_path: str, prompt: str) -> List[Dict]:
        """分析单张图片，返回结构化结果（数组格式）"""
        base64_image = self.encode_image(image_path)
        if self.client:
            return self._analyze_with_openai_sdk(base64_image, prompt, image_path)
        else:
            return self._analyze_with_requests(base64_image, prompt, image_path)

    def analyze_image_raw(self, image_path: str, prompt: str) -> str:
        """分析单张图片，返回模型原始文本（不进行JSON解析）"""
        base64_image = self.encode_image(image_path)
        if self.client:
            return self._analyze_raw_with_openai_sdk(base64_image, prompt)
        else:
            return self._analyze_raw_with_requests(base64_image, prompt)

    def _build_content(self, base64_image: str, prompt: str) -> list:
        return [
            {"type": "text", "text": prompt},
            {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}}
        ]

    # ---------- 原有分析逻辑 ----------
    def _analyze_with_openai_sdk(self, base64_image: str, prompt: str, image_path: str) -> List[Dict]:
        max_retries = 5
        retry_delay = 3
        for attempt in range(max_retries):
            try:
                print(f"🔄 调用 {self.model} 分析图片 (第 {attempt+1}/{max_retries} 次)...")
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[{"role": "user", "content": self._build_content(base64_image, prompt)}],
                    max_tokens=4000,
                    temperature=0.05,
                    timeout=300,
                )
                content = response.choices[0].message.content
                result = self._parse_json_response(content)
                if result:
                    return result
                elif attempt < max_retries - 1:
                    print("⚠️ VLM返回无效结果，重试...")
                    time.sleep(retry_delay)
                    retry_delay *= 2
                else:
                    print("❌ 所有重试均失败")
                    return []
            except Exception as e:
                print(f"❌ 调用失败 (第 {attempt+1} 次): {e}")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                    retry_delay *= 2
        return []

    def _analyze_with_requests(self, base64_image: str, prompt: str, image_path: str) -> List[Dict]:
        max_retries = 5
        retry_delay = 3
        headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
        payload = {
            "model": self.model,
            "messages": [{"role": "user", "content": self._build_content(base64_image, prompt)}],
            "max_tokens": 4000,
            "temperature": 0.05
        }
        for attempt in range(max_retries):
            try:
                print(f"🔄 调用 {self.model} (requests) 第 {attempt+1} 次...")
                resp = requests.post(f"{self.base_url}/chat/completions", headers=headers, json=payload, timeout=300)
                resp.raise_for_status()
                content = resp.json()["choices"][0]["message"]["content"]
                result = self._parse_json_response(content)
                if result:
                    return result
                elif attempt < max_retries - 1:
                    print("⚠️ 无效结果，重试...")
                    time.sleep(retry_delay)
                    retry_delay *= 2
            except Exception as e:
                print(f"❌ 请求失败: {e}")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                    retry_delay *= 2
        return []

    # ---------- 新增：原始文本返回 ----------
    def _analyze_raw_with_openai_sdk(self, base64_image: str, prompt: str) -> str:
        max_retries = 5
        retry_delay = 3
        for attempt in range(max_retries):
            try:
                print(f"🔄 调用 {self.model} 分析图片 (第 {attempt+1}/{max_retries} 次)...")
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[{"role": "user", "content": self._build_content(base64_image, prompt)}],
                    max_tokens=16384,
                    temperature=0.05,
                    timeout=300,
                )
                return response.choices[0].message.content
            except Exception as e:
                print(f"❌ 调用失败 (第 {attempt+1} 次): {e}")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                    retry_delay *= 2
        return ""

    def _analyze_raw_with_requests(self, base64_image: str, prompt: str) -> str:
        max_retries = 5
        retry_delay = 3
        headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
        payload = {
            "model": self.model,
            "messages": [{"role": "user", "content": self._build_content(base64_image, prompt)}],
            "max_tokens": 16384,
            "temperature": 0.05
        }
        for attempt in range(max_retries):
            try:
                print(f"🔄 调用 {self.model} (requests) 第 {attempt+1} 次...")
                resp = requests.post(f"{self.base_url}/chat/completions", headers=headers, json=payload, timeout=300)
                resp.raise_for_status()
                return resp.json()["choices"][0]["message"]["content"]
            except Exception as e:
                print(f"❌ 请求失败: {e}")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                    retry_delay *= 2
        return ""

    # ---------- JSON解析 ----------
    def _parse_json_response(self, content: str) -> List[Dict]:
        """解析 JSON 响应（只保留核心逻辑）"""
        if not content or not content.strip():
            return []
        content = content.strip()
        try:
            data = json.loads(content)
            if isinstance(data, list):
                return self._remove_duplicates(data)
            elif isinstance(data, dict):
                return [data]
        except:
            pass
        match = re.search(r'\[[\s\S]*\]', content)
        if match:
            try:
                data = json.loads(match.group())
                if isinstance(data, list):
                    return self._remove_duplicates(data)
            except:
                pass
        return []

    def _remove_duplicates(self, items: List[Dict]) -> List[Dict]:
        seen = set()
        unique = []
        for item in items:
            if not isinstance(item, dict) or 'words' not in item:
                continue
            w = item['words'].strip()
            if w and w not in seen:
                seen.add(w)
                unique.append(item)
        return unique


class QwenVLMClient(BaseVLMClient):
    def __init__(self, api_key: str = None, model: str = None, base_url: str = None):
        api_key = api_key or os.getenv("DASHSCOPE_API_KEY") or (VLM_CONFIG and VLM_CONFIG.get("qwen", {}).get("api_key"))
        model = model or (VLM_CONFIG and VLM_CONFIG.get("qwen", {}).get("model", "qwen3-vl-plus"))
        base_url = base_url or "https://dashscope.aliyuncs.com/compatible-mode/v1"
        if not api_key:
            raise ValueError("需要设置阿里云 API Key")
        super().__init__(api_key, model, base_url, vlm_type="qwen")
        if self.client:
            print(f"✅ 使用 OpenAI SDK 调用 Qwen VL ({self.model})")
        else:
            print("⚠️ 未安装 openai，将使用 requests 方式")


class MockVLMClient:
    """模拟客户端，用于测试"""
    def __init__(self, *args, **kwargs):
        self.vlm_type = "mock"

    def analyze_image(self, image_path: str, prompt: str) -> List[Dict]:
        print(f"🤖 [模拟] 分析图片: {os.path.basename(image_path)}")
        return [
            {"words": "模拟表头1", "type": "column_header"},
            {"words": "模拟表头2", "type": "row_header"}
        ]


def create_vlm_client(provider: str = "qwen", api_key: str = None, model: str = None, base_url: str = None):
    if provider.lower() == "qwen":
        return QwenVLMClient(api_key, model, base_url)
    elif provider.lower() == "mock":
        return MockVLMClient()
    else:
        raise ValueError(f"不支持的提供商: {provider}")