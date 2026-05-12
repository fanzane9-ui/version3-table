#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Step 1: 颜色标注 + Step 2: 提取 JSON + Step 3: 结构化存储
"""

import sys
import os
from pathlib import Path

# 将 version3 目录加入 sys.path
current_file = Path(__file__).resolve()
version3_dir = current_file.parent.parent
if str(version3_dir) not in sys.path:
    sys.path.insert(0, str(version3_dir))

from core.annotator import ExcelAnnotator
from core.extractor import batch_extract
from core.excel_generator import convert_json_to_structured_excel


def main():
    # 定义文件夹路径
    table_folder = os.path.join(version3_dir, "data", "table")
    vlm_folder = os.path.join(version3_dir, "data", "output", "vlm")
    output_folder = os.path.join(version3_dir, "data", "output", "table_mark")
    final_folder = os.path.join(version3_dir, "data", "output", "final")
    structured_folder = os.path.join(version3_dir, "data", "output", "structured")

    print("=" * 50)
    print("Step 1: 颜色标注")
    print("=" * 50)
    print(f"原始表格文件夹: {table_folder}")
    print(f"VLM 结果文件夹: {vlm_folder}")
    print(f"标注输出文件夹: {output_folder}")
    print("-" * 50)

    annotator = ExcelAnnotator(threshold=80, method='ratio', strict=True)
    annotator.batch_annotate(table_folder, vlm_folder, output_folder)

    print("\n" + "=" * 50)
    print("✅ 颜色标注完成！")
    print(f"标注后的表格已保存在: {output_folder}")
    print("\n请手动打开这些文件，检查并矫正颜色标注。")
    print("确保每个表头单元格颜色正确（红色=列头，绿色=行头），并保存文件。")
    print("\n完成人工矫正后，按 Enter 键继续 Step 2（提取 JSON）...")
    input()

    print("\n" + "=" * 50)
    print("Step 2: 从矫正后表格提取 JSON")
    print("=" * 50)
    print(f"输入文件夹: {output_folder}")
    print(f"输出文件夹: {final_folder}")
    print("-" * 50)
    batch_extract(output_folder, final_folder, debug=False)

    # ========== 新增 Step 3: 结构化存储 ==========
    print("\n" + "=" * 50)
    print("Step 3: 结构化存储（生成扁平化 Excel）")
    print("=" * 50)
    os.makedirs(structured_folder, exist_ok=True)

    # 查找所有 final JSON 文件
    json_files = [f for f in os.listdir(final_folder) if f.endswith('.json')]
    if not json_files:
        print("❌ 未找到 final JSON 文件，跳过结构化步骤")
    else:
        for json_file in json_files:
            json_path = os.path.join(final_folder, json_file)
            # 输出文件名：原文件名_structured.xlsx
            base_name = os.path.splitext(json_file)[0]
            output_excel = os.path.join(structured_folder, f"{base_name}_structured.xlsx")
            print(f"\n📄 处理: {json_file}")
            try:
                convert_json_to_structured_excel(json_path, output_excel, use_qwen=False)
            except Exception as e:
                print(f"❌ 结构化失败: {e}")
        print(f"\n✅ 结构化 Excel 已保存至: {structured_folder}")

    print("\n" + "=" * 50)
    print("🎉 全部流程完成！")
    print(f"最终 JSON 文件保存在: {final_folder}")
    print(f"结构化 Excel 文件保存在: {structured_folder}")


if __name__ == "__main__":
    main()