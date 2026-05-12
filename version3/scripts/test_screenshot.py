#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
截图测试脚本 - 批量生成 Excel 区域图片，支持分离模式
"""

import sys
import os
import argparse
import time
from pathlib import Path

# 将 version3 目录加入 sys.path（确保运行时能找到 core 模块）
current_file = Path(__file__).resolve()
version3_dir = current_file.parent.parent   # 指向 version3 目录
if str(version3_dir) not in sys.path:
    sys.path.insert(0, str(version3_dir))

from core.converter import DirectHeaderConverter


def find_excel_files(folder_path):
    """查找文件夹中的 Excel 文件"""
    if not os.path.exists(folder_path):
        return []
    excel_extensions = {'.xlsx', '.xls', '.xlsm'}
    return [os.path.join(folder_path, f) for f in os.listdir(folder_path)
            if any(f.lower().endswith(ext) for ext in excel_extensions)]


def test_single_file_normal(excel_path, output_dir, apply_formatting, verbose):
    """普通模式：生成单张截图到 output_dir"""
    os.makedirs(output_dir, exist_ok=True)
    try:
        with DirectHeaderConverter(apply_formatting=apply_formatting, verbose=verbose) as converter:
            if not converter.app:
                print(f"  ❌ Excel 启动失败")
                return False
            img_path = converter.convert_full_table(excel_path, output_dir)
            print(f"  ✅ 截图已保存: {os.path.basename(img_path)}")
            return True
    except Exception as e:
        print(f"  ❌ 截图失败: {e}")
        return False


def test_single_file_split(excel_path, dir_formatted, dir_raw, verbose):
    """分离模式：生成格式化和未格式化两张截图"""
    os.makedirs(dir_formatted, exist_ok=True)
    os.makedirs(dir_raw, exist_ok=True)
    try:
        with DirectHeaderConverter(apply_formatting=True, verbose=verbose) as converter:
            if not converter.app:
                print(f"  ❌ Excel 启动失败")
                return None
            fmt_path = converter.convert_full_table(excel_path, dir_formatted, suffix="")
            converter.apply_formatting = False
            raw_path = converter.convert_full_table(excel_path, dir_raw, suffix="")
        return {"filename": os.path.basename(excel_path), "fmt_path": fmt_path, "raw_path": raw_path}
    except Exception as e:
        print(f"  ❌ 截图失败: {e}")
        return None


def main():
    parser = argparse.ArgumentParser(description="Excel 截图测试工具（无筛选依赖版）")
    parser.add_argument("input", help="Excel 文件或包含 Excel 文件的文件夹")
    parser.add_argument("--output", "-o", default="test_images", help="输出根目录 (默认: test_images)")
    parser.add_argument("--no-format", action="store_true", help="禁用格式化（普通模式有效）")
    parser.add_argument("--split", action="store_true", help="分离模式：分别生成格式化和未格式化截图")
    parser.add_argument("--verbose", "-v", action="store_true", help="显示详细调试信息")
    args = parser.parse_args()

    input_path = Path(args.input)
    if input_path.is_file():
        files = [str(input_path)]
    elif input_path.is_dir():
        files = find_excel_files(str(input_path))
        if not files:
            print(f"❌ 文件夹中没有 Excel 文件")
            return
    else:
        print(f"❌ 路径不存在")
        return

    print(f"📁 找到 {len(files)} 个 Excel 文件")
    if args.split:
        print("🔄 模式: 分离模式（生成格式化和未格式化截图）")
        dir_fmt = os.path.join(args.output, "formatted")
        dir_raw = os.path.join(args.output, "raw")
        print(f"📁 格式化截图 → {dir_fmt}")
        print(f"📁 未格式化截图 → {dir_raw}")
    else:
        print("🔄 模式: 普通模式")
        print(f"🎨 格式化: {'启用' if not args.no_format else '禁用'}")
    print("-" * 50)

    start_time = time.time()
    if args.split:
        success = 0
        for i, file_path in enumerate(files, 1):
            print(f"[{i}/{len(files)}] 处理: {os.path.basename(file_path)}")
            if test_single_file_split(file_path, dir_fmt, dir_raw, args.verbose):
                success += 1
            print()
        print(f"✅ 成功生成 {success}/{len(files)} 组截图")
    else:
        success = 0
        for i, file_path in enumerate(files, 1):
            print(f"[{i}/{len(files)}] 处理: {os.path.basename(file_path)}")
            if test_single_file_normal(file_path, args.output, not args.no_format, args.verbose):
                success += 1
            print()
        print(f"✅ 成功生成 {success}/{len(files)} 张截图")

    elapsed = time.time() - start_time
    print(f"总耗时 {elapsed:.2f} 秒")


if __name__ == "__main__":
    main()