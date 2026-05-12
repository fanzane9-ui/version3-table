#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
子表拆分测试模块
- 默认仅使用标注截图 + Prompt 进行拆分
- 可选利用第一次 VLM 表头 JSON 过滤非表头词汇（如红色文字误判）
- 输出 split_result.json 和 metadata.txt
- 使用 --generate 生成子表结构化 Excel
"""

import os
import sys
import argparse
import json
from pathlib import Path

current_file = Path(__file__).resolve()
version3_dir = current_file.parent.parent
if str(version3_dir) not in sys.path:
    sys.path.insert(0, str(version3_dir))

from core.table_splitter import TableSplitter
from core.excel_generator import convert_json_to_structured_excel
from core.converter import DirectHeaderConverter
from core.config import get_api_key


def find_file_by_suffix(directory: str, suffix: str) -> str:
    if not os.path.isdir(directory):
        return ""
    for f in os.listdir(directory):
        if f.endswith(suffix):
            return os.path.join(directory, f)
    return ""


def main():
    parser = argparse.ArgumentParser(description="子表拆分独立测试")
    parser.add_argument("--session", help="会话目录名 (例如 abc123)")
    parser.add_argument("--final-json", help="直接指定最终 JSON 路径")
    parser.add_argument("--marked-excel", help="标注后的 Excel 路径")
    parser.add_argument("--vlm-json", default=None, help="第一次 VLM 表头 JSON 路径（用于过滤误识别的表头词汇，可选）")
    parser.add_argument("--api-key", help="Qwen API Key")
    parser.add_argument("--threshold", type=int, default=80, help="表头匹配阈值")
    parser.add_argument("--generate", action="store_true", help="生成子表结构化 Excel")
    parser.add_argument("--output-dir", default="data/output", help="输出根目录")
    args = parser.parse_args()

    # ---------- 确定文件 ----------
    if args.final_json:
        final_json = args.final_json
        session_dir = os.path.dirname(os.path.dirname(final_json))
    elif args.session:
        base_output = Path(args.output_dir)
        session_dir = str(base_output / args.session)
        if not os.path.exists(session_dir):
            print(f"❌ 会话目录不存在: {session_dir}")
            return
        final_dir = os.path.join(session_dir, "final")
        final_json = find_file_by_suffix(final_dir, ".json")
        if not final_json:
            print(f"❌ 未找到最终 JSON")
            return
    else:
        print("❌ 请指定 --session 或 --final-json")
        return

    if not os.path.exists(final_json):
        print(f"❌ 最终 JSON 不存在: {final_json}")
        return

    base_name = os.path.splitext(os.path.basename(final_json))[0]

    # 标注 Excel
    if args.marked_excel:
        marked_excel = args.marked_excel
    else:
        mark_dir = os.path.join(session_dir, "table_mark")
        possible_names = [base_name + ".xlsx", base_name.replace("_final", "") + ".xlsx"]
        marked_excel = ""
        for name in possible_names:
            path = os.path.join(mark_dir, name)
            if os.path.exists(path):
                marked_excel = path
                break
        if not marked_excel:
            print(f"❌ 找不到标注后 Excel，请用 --marked-excel 指定")
            return

    if not os.path.exists(marked_excel):
        print(f"❌ 标注 Excel 不存在: {marked_excel}")
        return

    # API Key
    api_key = args.api_key or (get_api_key("qwen") if get_api_key else os.getenv("DASHSCOPE_API_KEY"))
    if not api_key:
        print("❌ 缺少 API Key，请设置")
        return

    # ---------- 标注截图 ----------
    images_dir = os.path.join(session_dir, "images")
    os.makedirs(images_dir, exist_ok=True)
    excel_stem = os.path.splitext(os.path.basename(marked_excel))[0]
    marked_img = os.path.join(images_dir, f"{excel_stem}_marked.png")
    if not os.path.exists(marked_img):
        print("📸 正在生成标注截图...")
        try:
            with DirectHeaderConverter(apply_formatting=True) as converter:
                if not converter.app:
                    print("❌ Excel 启动失败")
                    return
                marked_img = converter.convert_full_table(marked_excel, images_dir, suffix="marked")
                print(f"✅ 标注截图已生成: {marked_img}")
        except Exception as e:
            print(f"❌ 生成截图失败: {e}")
            return
    else:
        print(f"📸 标注截图已存在: {marked_img}")

    # ---------- 拆分 ----------
    # 如果提供了第一次 VLM 的表头 JSON，则用于过滤误识别的表头词汇
    trusted_path = args.vlm_json if args.vlm_json else None
    splitter = TableSplitter(vlm_provider="qwen", api_key=api_key, match_threshold=args.threshold,
                             trusted_headers_path=trusted_path)

    try:
        raw_split = splitter.get_sub_tables_raw(marked_img)
    except Exception as e:
        print(f"❌ 拆分失败: {e}")
        return

    # 保存拆分方案
    split_json_path = os.path.join(session_dir, "split_result.json")
    with open(split_json_path, 'w', encoding='utf-8') as f:
        json.dump(raw_split, f, ensure_ascii=False, indent=2)
    print(f"💾 拆分方案已保存: {split_json_path}")

    sub_tables = raw_split.get('sub_tables', [])
    print("\n📋 拆分结果：")
    for sub in sub_tables:
        sid = sub.get('sub_table_id')
        if sid == 0:
            print(f"  [辅助信息] {sub.get('name')}")
            print(f"    {sub.get('description')}")
            for h in sub.get('headers', []):
                print(f"    - {h.get('words')} ({h.get('type')})")
        else:
            print(f"  子表 {sid}: {sub.get('name')}")
            print(f"    描述: {sub.get('description')}")
            rheaders = sub.get('row_headers', [])
            cheaders = sub.get('column_headers', [])
            headers = sub.get('headers', [])
            if headers:
                print(f"    表头数量: {len(headers)}")
            if rheaders or cheaders:
                print(f"    行表头: {len(rheaders)} 个，列表头: {len(cheaders)} 个")

    # ---------- 辅助信息 ----------
    meta_lines = []
    for sub in sub_tables:
        if sub.get('sub_table_id') == 0:
            meta_lines.append(f"[{sub.get('name')}]")
            meta_lines.append(sub.get('description', ''))
            for h in sub.get('headers', []):
                meta_lines.append(f"  - {h.get('words')} ({h.get('type')})")
            break
    if meta_lines:
        meta_path = os.path.join(session_dir, "metadata.txt")
        with open(meta_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(meta_lines))
        print(f"📝 辅助信息已保存: {meta_path}")

    # ---------- 可选：生成结构化 ----------
    if args.generate:
        print("\n🔧 按拆分结果生成子表结构化 Excel...")
        try:
            result_sub_tables = splitter.split_table(marked_img, trusted_path, final_json)
        except Exception as e:
            print(f"❌ 切割子表失败: {e}")
            return

        if not result_sub_tables:
            print("⚠️ 没有可生成的子表")
            return

        structured_dir = os.path.join(session_dir, "structured")
        os.makedirs(structured_dir, exist_ok=True)

        for sub in result_sub_tables:
            sub_id = sub['sub_table_id']
            cells = sub['cells']
            if not cells:
                continue
            # 保存子表 JSON
            sub_json_name = f"{excel_stem}_sub_{sub_id}.json"
            sub_json_path = os.path.join(session_dir, "final", sub_json_name)
            with open(sub_json_path, 'w', encoding='utf-8') as f:
                json.dump(cells, f, ensure_ascii=False, indent=2)

            # 生成结构化 Excel
            tmp_json = os.path.join(structured_dir, f"_temp_sub_{sub_id}.json")
            with open(tmp_json, 'w', encoding='utf-8') as f:
                json.dump(cells, f, ensure_ascii=False, indent=2)
            structured_excel = os.path.join(structured_dir, f"{excel_stem}_sub_{sub_id}_structured.xlsx")
            try:
                convert_json_to_structured_excel(tmp_json, structured_excel, use_qwen=False)
                print(f"  ✅ 子表 {sub_id} 结构化输出: {os.path.basename(structured_excel)}")
            except Exception as e:
                print(f"  ❌ 子表 {sub_id} 结构化失败: {e}")
            finally:
                if os.path.exists(tmp_json):
                    os.remove(tmp_json)

        print("🎉 结构化生成完成。")
    else:
        print("\n🔍 已输出拆分方案，如需生成结构化表格，请加 --generate 参数")


if __name__ == "__main__":
    main()