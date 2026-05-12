#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
表头VLM处理器 - 固定使用格式化截图
单文件处理直接使用格式化图片，批量处理跳过筛选步骤
输出结构：output/images/{formatted, raw}
"""

import os
import json
import time
import tempfile
import shutil
import traceback
from typing import Dict
from tqdm import tqdm
from core.converter import DirectHeaderConverter
from core.vlm_client import create_vlm_client


class HeaderVLMProcessor:
    def __init__(self, vlm_provider: str = "qwen", api_key: str = None):
        self.vlm_client = create_vlm_client(vlm_provider, api_key)
        self.vlm_provider = vlm_provider

    def _get_end_to_end_prompt(self) -> str:
        """
        端到端专用 Prompt (V4 终极融合版) - 无位置信息，仅输出表头文本和类型
        """
        return '''# Role: 资深表格结构解析专家

    ## 🎯 任务目标
    分析这张完整的表格截图，精准提取所有**表头单元格**，并严格区分其类型（列头/行头）。
    **核心原则**：只关注语义分类信息，忽略具体数值数据。

    ## 🧠 第一步：全局结构分析与逻辑推理 (CoT)
    在输出结果前，请先在内心执行以下推理步骤（不要输出思考过程，仅作为判断依据）：

    1. **【灵魂三问 - 左侧列判定】**
       - 扫描左侧第一列：内容是 "1,2,3..." (序号)? "周一/周二" (时间)? 还是 "产品A/产品B" (分类)?
       - 如果是序号或简单时间序列 -> **视为数据索引**，标记为 column_header (若位于顶部) 或忽略 (若纯数据)。
       - 如果左侧有明确的分类层级 (如 "销售部", "生产部") -> **视为行表头 (row_header)**。

    2. **【顶部行判定】**
       - 扫描顶部区域：是否有 "2023年", "Q1", "销售额" 等分类标签?
       - 如果有 -> **视为列头 (column_header)**。
       - 如果顶部全是具体数值 -> **忽略**。

    3. **【矩阵完整性检查】**
       - 观察数据区：是否形成行列交叉的矩阵？
       - 左上角总标题 (如 "2023 年报") -> **视为 column_header** (统一归类到列头)，不要单独标记为 title。

    ## ⚠️ 关键指令 (必须严格遵守)
    1. **只识别表头**：忽略所有具体的数值数据行（如销售额、数量、日期值、人名）。
    2. **区分类型**：
       - 位于顶部的标题行 (含左上角总标题) -> type: "column_header"
       - 位于左侧的分类列表 (排除纯序号/时间) -> type: "row_header"
    3. **处理合并单元格**：
       - 跨越多行/多列的单元格视为**一个整体对象**。
       - 单元格内若有换行文字，用空格连接 (例如："生产\n统计" -> "生产 统计")。
    4. **多级表头**：
       - 必须**逐行逐格**全部列出，不要合并层级。
       - 例如：第一行"2023年"，第二行"Q1", "Q2" -> 分别输出两个对象。
    5. **负向约束**：
       - ❌ 严禁识别数据行 (看到纯数字行立即停止)。
       - ❌ 严禁拆分合并单元格。
       - ❌ 严禁臆造内容。

    ## 💡 Few-Shot 学习示例 (参考模式)

    **输入场景 A (混合表头)**:
    [Row 1] [2023 年报] [------ 销售统计 ------]
    [Row 2] [      ] [Q1] [Q2] [Q3]
    [Row 3] [销售部] [100] [200] [150]
    [Row 4] [生产部] [300] [400] [500]

    **正确输出**:
    [
      {"words": "2023 年报", "type": "column_header"},
      {"words": "销售统计", "type": "column_header"},
      {"words": "Q1", "type": "column_header"},
      {"words": "Q2", "type": "column_header"},
      {"words": "Q3", "type": "column_header"},
      {"words": "销售部", "type": "row_header"},
      {"words": "生产部", "type": "row_header"}
    ]
    *(注意：忽略了数据行 100, 200...)*

    **输入场景 B (含序号列)**:
    [Row 1] [序号] [产品名称] [单价]
    [Row 2] [1] [手机] [100]
    [Row 3] [2] [电脑] [5000]

    **正确输出**:
    [
      {"words": "序号", "type": "column_header"}, 
      {"words": "产品名称", "type": "column_header"},
      {"words": "单价", "type": "column_header"}
    ]
    *(注意：左侧序号列被视为列头的一部分，因为它是顶部的一行；若左侧是纯数据行则忽略)*

    ## 📝 输出格式要求
    1. **仅输出一个标准的 JSON 列表**。
    2. **不要包含** Markdown 代码块标记 (```json)、不要包含思考过程、不要有任何解释性文字。
    3. **确保 type 字段** 只能是 "column_header" 或 "row_header"。

    [
      {"words": "表头文字", "type": "column_header"},
      {"words": "表头文字", "type": "row_header"}
    ]

    现在请分析图片：'''

    def process_excel_file(self, excel_path: str, output_dir: str = "output") -> Dict:
        """
        单文件处理：只生成格式化截图，直接用于识别
        """
        base_name = os.path.splitext(os.path.basename(excel_path))[0]
        images_dir = os.path.join(output_dir, "images")
        vlm_dir = os.path.join(output_dir, "vlm")
        os.makedirs(images_dir, exist_ok=True)
        os.makedirs(vlm_dir, exist_ok=True)

        try:
            # 只生成格式化截图
            with DirectHeaderConverter(apply_formatting=True) as converter:
                if not converter.app:
                    return {"success": False, "error": "Excel 应用程序未启动"}
                best_path = converter.convert_full_table(excel_path, images_dir, suffix="")
                fmt_path = best_path
                raw_path = ""   # 不再生成原始截图

            print(f"🏆 使用图片: {os.path.basename(best_path)}")

            prompt = self._get_end_to_end_prompt()
            print(f"🤖 调用 VLM ({self.vlm_provider}) 进行识别...")
            headers = self.vlm_client.analyze_image(best_path, prompt)
            if not headers:
                return {"success": False, "error": "VLM 未返回任何表头"}

            result = {
                "file": os.path.basename(excel_path),
                "mode": "end_to_end_full_image",
                "image_source": best_path,
                "formatted_image": fmt_path,
                "raw_image": raw_path,
                "best_score": 0.0,
                "headers": headers,
                "vlm_provider": self.vlm_provider,
                "total_headers_found": len(headers)
            }
            out_path = os.path.join(vlm_dir, f"{base_name}_e2e_vlm.json")
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(result, f, ensure_ascii=False, indent=2)

            print(f"🎉 完成，共 {len(headers)} 个表头")
            return {"success": True, "output_path": out_path, "headers_count": len(headers)}
        except Exception as e:
            traceback.print_exc()
            return {"success": False, "error": str(e)}

    def batch_process(self, input_folder: str = "table", output_dir: str = "output",
                      on_file_complete=None) -> Dict:
        """
        批量处理：先生成所有图片（仅格式化截图），然后直接进行VLM识别
        输出结构：output/images/ 下直接存放格式化截图
        """
        if not os.path.exists(input_folder):
            return {"success": False, "error": "文件夹不存在"}
        excel_extensions = {'.xlsx', '.xls', '.xlsm'}
        excel_files = [f for f in os.listdir(input_folder)
                       if any(f.lower().endswith(ext) for ext in excel_extensions)]
        if not excel_files:
            return {"success": False, "error": "未找到 Excel 文件"}

        # 创建临时文件夹（仅用于格式化截图）
        temp_dir = tempfile.mkdtemp(prefix="excel_temp_")
        dir_formatted = os.path.join(temp_dir, "formatted")
        os.makedirs(dir_formatted, exist_ok=True)

        # 第一步：为所有文件生成格式化截图
        file_records = []  # 存储 (base_name, original_filename, excel_path)
        print("🖼️ 正在生成格式化截图...")
        for filename in tqdm(excel_files, desc="生成截图"):
            excel_path = os.path.join(input_folder, filename)
            base_name = os.path.splitext(filename)[0]
            try:
                # 只生成格式化截图
                with DirectHeaderConverter(apply_formatting=True) as converter:
                    fmt_path = converter.convert_full_table(excel_path, dir_formatted)
                file_records.append((base_name, filename, excel_path))
            except Exception as e:
                print(f"❌ 生成截图失败 {filename}: {e}")

        if not file_records:
            shutil.rmtree(temp_dir)
            return {"success": False, "error": "没有成功生成任何图片"}

        # 将图片复制到 output/images 目录下
        images_dir = os.path.join(output_dir, "images")
        os.makedirs(images_dir, exist_ok=True)
        for filename in os.listdir(dir_formatted):
            shutil.copy2(os.path.join(dir_formatted, filename), os.path.join(images_dir, filename))

        # 第二步：对格式化图片进行 VLM 识别
        results = {}
        vlm_dir = os.path.join(output_dir, "vlm")
        os.makedirs(vlm_dir, exist_ok=True)

        with tqdm(file_records, desc="识别进度", unit="文件") as pbar:
            for base_name, original_filename, excel_path in pbar:
                start_time = time.time()
                try:
                    best_path = os.path.join(images_dir, f"{base_name}_full_table.png")
                    if not os.path.exists(best_path):
                        raise FileNotFoundError(f"图片不存在: {best_path}")

                    prompt = self._get_end_to_end_prompt()
                    headers = self.vlm_client.analyze_image(best_path, prompt)
                    if not headers:
                        res = {"success": False, "error": "VLM 未返回任何表头"}
                    else:
                        result = {
                            "file": original_filename,
                            "mode": "end_to_end_full_image",
                            "image_source": best_path,
                            "formatted_image": best_path,  # 与 image_source 相同
                            "best_score": 0.0,
                            "headers": headers,
                            "vlm_provider": self.vlm_provider,
                            "total_headers_found": len(headers)
                        }
                        out_path = os.path.join(vlm_dir, f"{base_name}_e2e_vlm.json")
                        with open(out_path, "w", encoding="utf-8") as f:
                            json.dump(result, f, ensure_ascii=False, indent=2)
                        res = {"success": True, "output_path": out_path, "headers_count": len(headers)}

                    elapsed = time.time() - start_time
                    results[original_filename] = res
                    if res["success"]:
                        pbar.write(f"✅ {original_filename}: {res['headers_count']} 个表头 (耗时 {elapsed:.1f}s)")
                    else:
                        pbar.write(f"❌ {original_filename}: {res.get('error')} (耗时 {elapsed:.1f}s)")
                    if on_file_complete:
                        on_file_complete(original_filename, res)
                except Exception as e:
                    elapsed = time.time() - start_time
                    error_msg = f"{e}\n{traceback.format_exc()}"
                    results[original_filename] = {"success": False, "error": str(e)}
                    pbar.write(f"❌ {original_filename}: 异常 ({elapsed:.1f}s) - {error_msg[:200]}")

        # 清理临时文件夹
        shutil.rmtree(temp_dir)

        success_count = sum(1 for r in results.values() if r.get("success"))
        return {"success": True, "total": len(results), "success_count": success_count, "results": results}