#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 表格分析工具 - 主入口
功能：
  1. 仅 VLM 识别（默认）
  2. VLM 识别 + 自动标注（加 --annotate 参数）
"""

import sys
import os
import argparse
from pathlib import Path

# 将 version3 目录加入 sys.path（必须在所有 core 导入之前）
current_file = Path(__file__).resolve()
version3_dir = current_file.parent.parent   # 指向 version3 目录
if str(version3_dir) not in sys.path:
    sys.path.insert(0, str(version3_dir))

# 现在可以安全导入 core 模块
from core.processor import HeaderVLMProcessor
from core.annotator import ExcelAnnotator
from core.config import get_api_key


def find_excel_files(folder_path):
    """查找文件夹中的 Excel 文件"""
    if not os.path.exists(folder_path):
        return []
    excel_extensions = {'.xlsx', '.xls', '.xlsm'}
    return [os.path.join(folder_path, f) for f in os.listdir(folder_path)
            if any(f.lower().endswith(ext) for ext in excel_extensions)]


def main():
    parser = argparse.ArgumentParser(description="Excel 表格分析工具 - 主入口")
    parser.add_argument("--input", nargs="?", default="data/table", help="Excel 文件或文件夹（相对于项目根目录）")
    parser.add_argument("--vlm", choices=["qwen", "mock"], default="qwen", help="VLM 提供商")
    parser.add_argument("--api-key", help="API 密钥")
    parser.add_argument("--output", default="data/output", help="输出目录（相对于项目根目录）")
    parser.add_argument("--annotate", action="store_true", help="VLM 识别后自动进行颜色标注，并暂停等待人工矫正")
    args = parser.parse_args()

    # 将相对路径转换为基于项目根目录的绝对路径
    input_path_abs = Path(version3_dir) / args.input
    output_path_abs = Path(version3_dir) / args.output

    if not input_path_abs.exists():
        print(f"❌ 路径不存在: {input_path_abs}")
        return

    # API Key 处理
    api_key = args.api_key
    if args.vlm == "qwen" and not api_key:
        api_key = os.getenv("DASHSCOPE_API_KEY")
        if not api_key and get_api_key:
            api_key = get_api_key("qwen")
        if not api_key:
            print("❌ 请设置阿里云 API Key")
            return

    # 创建 VLM 处理器
    processor = HeaderVLMProcessor(vlm_provider=args.vlm, api_key=api_key)

    # 获取 Excel 文件列表
    if input_path_abs.is_file():
        excel_files = [str(input_path_abs)]
    else:
        excel_files = find_excel_files(str(input_path_abs))
        if not excel_files:
            print(f"❌ 文件夹中没有 Excel 文件: {input_path_abs}")
            return

    print(f"📊 找到 {len(excel_files)} 个 Excel 文件")
    print(f"🔧 使用 VLM: {args.vlm}")
    print(f"📁 输出目录: {output_path_abs}")
    if args.annotate:
        print("🎨 将执行颜色标注并暂停等待人工矫正")
    print("-" * 50)

    # ------------------ VLM 识别阶段 ------------------
    vlm_success_count = 0
    if len(excel_files) == 1:
        print("🚀 开始 VLM 识别...")
        result = processor.process_excel_file(excel_files[0], str(output_path_abs))
        if result["success"]:
            vlm_success_count = 1
            print(f"✅ 识别完成，结果: {result['output_path']}")
        else:
            print(f"❌ 识别失败: {result.get('error')}")
    else:
        print("🚀 开始批量 VLM 识别...")
        def on_complete(filename, res):
            if res["success"]:
                print(f"   ✅ {filename}: {res['headers_count']} 个表头")
            else:
                print(f"   ❌ {filename}: {res.get('error')}")

        batch_result = processor.batch_process(str(input_path_abs), str(output_path_abs), on_file_complete=on_complete)
        if batch_result["success"]:
            vlm_success_count = batch_result["success_count"]
            print(f"\n✅ VLM 识别完成: {vlm_success_count}/{batch_result['total']} 个文件成功")
        else:
            print(f"❌ 批量处理失败: {batch_result.get('error')}")

    # ------------------ 标注阶段（如果启用） ------------------
    if args.annotate and vlm_success_count > 0:
        print("\n" + "=" * 50)
        print("开始颜色标注...")
        print("=" * 50)
        if ExcelAnnotator is None:
            print("❌ 无法导入 ExcelAnnotator，跳过标注")
            return

        table_folder = str(input_path_abs) if input_path_abs.is_dir() else str(input_path_abs.parent)
        vlm_folder = str(output_path_abs / "vlm")
        mark_folder = str(output_path_abs / "table_mark")

        annotator = ExcelAnnotator(threshold=80, method='ratio', strict=True)
        annotator.batch_annotate(table_folder, vlm_folder, mark_folder)

        print("\n" + "=" * 50)
        print("✅ 颜色标注完成！")
        print(f"标注后的表格已保存在: {mark_folder}")
        print("\n请手动打开这些文件，检查并矫正颜色标注。")
        print("确保每个表头单元格颜色正确（红色=列头，绿色=行头），并保存文件。")
        input("完成人工矫正后，按 Enter 键继续...")

        # ========== 新增：自动提取 JSON ==========
        from core.extractor import batch_extract
        print("\n" + "=" * 50)
        print("Step 2: 从矫正后表格提取 JSON")
        print("=" * 50)
        final_folder = str(output_path_abs / "final")
        batch_extract(mark_folder, final_folder, debug=False)
        print(f"\n🎉 提取完成！最终 JSON 保存在: {final_folder}")

    print("\n所有操作完成。")


if __name__ == "__main__":
    main()