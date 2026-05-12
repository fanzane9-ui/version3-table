#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
表格子表拆分器（基于 VLM + 最终 JSON）
- Prompt 强化拆分粒度与背景色判断（通用规则，不针对具体案例）
- 可选利用首次 VLM 的表头 JSON 过滤掉误识别的非表头词汇
"""

import json
import os
import re
from typing import List, Dict, Optional
from fuzzywuzzy import fuzz

from core.vlm_client import create_vlm_client
from core.annotator import normalize_text


class TableSplitter:
    def __init__(self, vlm_provider: str = "qwen", api_key: str = None,
                 match_threshold: int = 80,
                 trusted_headers_path: Optional[str] = None):
        self.vlm_client = create_vlm_client(vlm_provider, api_key)
        self.match_threshold = match_threshold
        # 可信表头白名单（用于后处理过滤）
        self.trusted_header_words = set()
        if trusted_headers_path and os.path.exists(trusted_headers_path):
            with open(trusted_headers_path, 'r', encoding='utf-8-sig') as f:
                header_data = json.load(f)
            headers = header_data.get('headers', [])
            for h in headers:
                w = h.get('words', '').strip()
                if w:
                    self.trusted_header_words.add(normalize_text(w))

    def _build_split_prompt(self) -> str:
        """
        构建拆分 Prompt，始终不注入表头列表，
        但包含背景色判断、强制横向拆分等通用规则。
        """
        prompt = '''你是一个表格理解专家，现在有一张复杂表格需要你处理。
为了帮助你从视觉上理解表格，我做了一些预处理：
表格中标注为红色的单元格为表格的列表头
表格中标注为绿色的单元格为表格的行表头
表格中其余颜色的单元格均视为数据
现在你基于颜色标注的表头信息进行复杂表格的拆分，一定要严格按以下规则拆分：
1. 我们认为拆分后的子表必然也是规则的矩形区域，存在于原始表格中。
2. 可能存在多级表头的情况，例如一级列表头下可能存在很多二级列表头，再下面才可能是数据。
3. 充分利用表头的位置信息。人类看表的顺序必然是从左到右，从上到下。那么对于表格的拆分也可以利用这个方法，我们可以从数据单元格入手。对于任意一个数据单元格（其余颜色），它上方或者左方的单元格只存在三种情况：要么是数据单元格（其余颜色），要么是表头单元格（行表头或者列表头被标注为绿色或者红色），要么是整张表格的边界。基于这种启发我们可以很容易圈定一个矩形的数据区域。
举个简单的例子（仅为说明方法，并非实际表格）：
| 姓名 | 年龄 | 城市 |
| 张三 | 25 | 北京 |
| 李四 | 30 | 上海 |
在这个表格中，你可以先找到一个矩形数据区域：张三、25、北京、李四、30、上海。然后基于该数据区域向上方查找发现存在列表头的对应，向左查找发现是整张表格的边界。那么该表就可以作为一个完整的子表被拆分出来。
4. 一定要基于位置信息的前提下，再进行语义上的理解，辅助子表拆分更加合理完善。
5. 如果某个区域内，左侧有一组列表头（如某些字段）和对应的数据行，右侧有另一组完全不同的列表头（如另一些字段）和对应的数据行，且这两组列表头之间没有共享数据，必须拆分为两个独立的子表，哪怕它们物理上在同一个高度。不要强行合并成一个子表。
6. ⚠️ 判断表头的唯一标准是单元格的背景填充颜色，不是文字颜色！只有那些单元格填充色为红色的才是列表头，填充色为绿色的才是行表头。如果某个单元格背景是白色/无色，即使它的文字是红色或绿色，也请不要列为表头。请忽略文字颜色，只关注单元格底色。
7. ⚠️ 多级表头必须全部列出！对于每一个子表，请将每一个红色单元格（不论它是几级表头）都列入 column_headers，每一个绿色单元格都列入 row_headers。绝对不要遗漏任何一级表头。
现在请你先简要说明你的拆分逻辑（200字以内），然后严格按照以下 JSON 格式输出拆分结果（不要包含任何其他文字，JSON 前后不要有代码块标记）：
{
  "sub_tables": [
    {
      "sub_table_id": 0,
      "name": "辅助信息",
      "description": "表格顶部的标题、日期等说明性文字",
      "headers": [{"words": "...", "type": "column_header"}]
    },
    {
      "sub_table_id": 1,
      "name": "子表名称",
      "description": "对该子表内容的简要描述",
      "row_headers": [{"words": "...", "type": "row_header"}],
      "column_headers": [{"words": "...", "type": "column_header"}]
    }
  ]
}
注意：
- sub_table_id 为 0 的子表只放表格顶部或四周的说明性文字。
- 其余每个子表严格区分 row_headers（绿色行头）和 column_headers（红色列头）。
- 请勿遗漏任何表头单元格。'''
        return prompt

    def _extract_json_object(self, content: str) -> Optional[Dict]:
        """从可能包含额外文字的响应中提取最外层的 JSON 对象"""
        if not content:
            return None
        content = re.sub(r'```(?:json)?\s*', '', content)
        content = content.strip()
        start = content.find('{')
        end = content.rfind('}')
        if start != -1 and end != -1 and end > start:
            json_str = content[start:end+1]
            try:
                return json.loads(json_str)
            except json.JSONDecodeError:
                pass
        match = re.search(r'\{[^{}]*"sub_tables"\s*:\s*\[[^}]*?\}\s*\]\s*\}', content, re.DOTALL)
        if match:
            try:
                return json.loads(match.group())
            except json.JSONDecodeError:
                pass
        return None

    def _parse_split_response(self, content: str) -> Dict:
        """解析 VLM 返回的拆分 JSON"""
        parsed = self._extract_json_object(content)
        if parsed and 'sub_tables' in parsed:
            return parsed
        match = re.search(r'"sub_tables"\s*:\s*(\[[\s\S]*?\])', content)
        if match:
            try:
                sub_tables = json.loads(match.group(1))
                return {"sub_tables": sub_tables}
            except json.JSONDecodeError:
                pass
        return {}

    def _filter_headers(self, sub_tables: list) -> None:
        """
        使用可信表头白名单过滤掉 VLM 错误识别的表头词汇（如红色文字误判）。
        如果未提供白名单，则不进行过滤。
        """
        if not self.trusted_header_words:
            return
        for sub in sub_tables:
            # 过滤 row_headers
            if 'row_headers' in sub:
                sub['row_headers'] = [
                    h for h in sub['row_headers']
                    if normalize_text(h.get('words', '')) in self.trusted_header_words
                ]
            # 过滤 column_headers
            if 'column_headers' in sub:
                sub['column_headers'] = [
                    h for h in sub['column_headers']
                    if normalize_text(h.get('words', '')) in self.trusted_header_words
                ]
            # 兼容旧格式
            if 'headers' in sub:
                sub['headers'] = [
                    h for h in sub['headers']
                    if normalize_text(h.get('words', '')) in self.trusted_header_words
                ]

    def get_sub_tables_raw(self, marked_image_path: str) -> Dict:
        """
        调用 VLM 进行拆分，返回原始 sub_tables 结构（不进行单元格切割）。
        如果初始化时提供了 trusted_headers_path，会在返回前自动过滤。
        """
        prompt = self._build_split_prompt()
        raw_text = self.vlm_client.analyze_image_raw(marked_image_path, prompt)
        if not raw_text:
            raise RuntimeError("VLM 未返回任何内容")
        parsed = self._parse_split_response(raw_text)
        if not parsed or 'sub_tables' not in parsed:
            print("⚠️ VLM 返回内容为：")
            print(raw_text[:500])
            raise RuntimeError("无法解析 VLM 返回的拆分子表 JSON")

        sub_tables = parsed.get('sub_tables', [])
        # 后处理过滤
        self._filter_headers(sub_tables)
        parsed['sub_tables'] = sub_tables
        return parsed

    def split_table(self, marked_image_path: str,
                    header_json_path: Optional[str],
                    final_json_path: str) -> List[Dict]:
        """
        拆分并基于 final JSON 切割出每个子表的单元格列表。
        注意：header_json_path 仅用于切割时的表头匹配，不参与拆分 Prompt。
        """
        raw_split = self.get_sub_tables_raw(marked_image_path)
        sub_tables_info = raw_split.get('sub_tables', [])

        with open(final_json_path, 'r', encoding='utf-8') as f:
            all_cells = json.load(f)

        result_sub_tables = []
        for sub in sub_tables_info:
            sub_id = sub.get('sub_table_id', -1)
            if sub_id == 0:
                continue

            row_headers = sub.get('row_headers', [])
            col_headers = sub.get('column_headers', [])
            if not row_headers and not col_headers:
                headers = sub.get('headers', [])
                for h in headers:
                    if h.get('type') == 'row_header':
                        row_headers.append(h)
                    else:
                        col_headers.append(h)

            if not row_headers and not col_headers:
                continue

            # 匹配表头单元格
            header_cells = []
            for h in row_headers + col_headers:
                words = h.get('words', '')
                htype = h.get('type', '')
                if not words or htype not in ('column_header', 'row_header'):
                    continue
                target_norm = normalize_text(words)
                best_sim = 0
                best_cell = None
                for cell in all_cells:
                    if cell.get('type') != htype:
                        continue
                    cell_text = normalize_text(cell.get('words', ''))
                    sim = fuzz.ratio(target_norm, cell_text)
                    if sim > best_sim:
                        best_sim = sim
                        best_cell = cell
                if best_sim >= self.match_threshold and best_cell:
                    header_cells.append(best_cell)

            if not header_cells:
                continue

            data_row_set = set()
            data_col_set = set()
            for hcell in header_cells:
                if hcell['type'] == 'row_header':
                    for r in hcell['rows']:
                        data_row_set.add(r)
                elif hcell['type'] == 'column_header':
                    for c in hcell['columns']:
                        data_col_set.add(c)

            sub_cells = []
            added = set()
            for hcell in header_cells:
                key = (tuple(hcell['rows']), tuple(hcell['columns']))
                if key not in added:
                    sub_cells.append(hcell)
                    added.add(key)

            for cell in all_cells:
                if cell['type'] != 'data':
                    continue
                cell_rows = set(cell['rows'])
                cell_cols = set(cell['columns'])
                row_ok = (not data_row_set) or (not cell_rows.isdisjoint(data_row_set))
                col_ok = (not data_col_set) or (not cell_cols.isdisjoint(data_col_set))
                if row_ok and col_ok:
                    sub_cells.append(cell)

            for cell in all_cells:
                if cell['type'] not in ('column_header', 'row_header'):
                    continue
                key = (tuple(cell['rows']), tuple(cell['columns']))
                if key in added:
                    continue
                cell_rows = set(cell['rows'])
                cell_cols = set(cell['columns'])
                if cell['type'] == 'column_header':
                    if not data_col_set or not cell_cols.isdisjoint(data_col_set):
                        sub_cells.append(cell)
                        added.add(key)
                else:
                    if not data_row_set or not cell_rows.isdisjoint(data_row_set):
                        sub_cells.append(cell)
                        added.add(key)

            sub_cells.sort(key=lambda x: (min(x['rows']), min(x['columns'])))
            result_sub_tables.append({
                "sub_table_id": sub_id,
                "name": sub.get('name', ''),
                "description": sub.get('description', ''),
                "cells": sub_cells
            })

        return result_sub_tables