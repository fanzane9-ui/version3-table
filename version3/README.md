# 半结构化表格智能分析系统

> 基于视觉语言模型（VLM）的 Excel 表格分析与结构化工具  
> 从原始半结构化表格（多级表头、合并单元格）到标准化 JSON / 扁平化 Excel 的完整流程

## ✨ 核心功能

- **智能截图预处理**  
  自动调整列宽、合并单元格行高、居中、增强边框，并自动跳过顶部嵌入图片。

- **VLM 表头识别**  
  调用 Qwen VL（或 Mock）模型，精准提取所有表头单元格，区分行头/列头，并输出每个表头的视觉行列位置。

- **自动颜色标注**  
  基于 VLM 结果，在原始 Excel 中通过模糊匹配 + 密度过滤动态定位表头，红色标记列头，绿色标记行头。

- **人工矫正与最终提取**  
  用户可在 Excel 中手动调整颜色，系统自动提取所有单元格的文本、位置和类型，生成结构化 JSON。

- **多级表头扁平化**  
  将复杂表头展开为单级路径（例如 `"外圈 > 内径报废"`），输出规整的二维表，便于入库或 SQL 查询。

- **交互式前端界面（Streamlit）**  
  5 步向导式操作：上传文件 → VLM 配置 → 自动处理 → 人工矫正 → 下载结果。支持会话隔离、日志查看、未来功能预告。

## 🧭 系统流程

```mermaid
graph LR
    A[原始 Excel] --> B[截图预处理]
    B --> C[VLM 表头识别]
    C --> D[颜色标注]
    D --> E[人工矫正 Excel]
    E --> F[提取最终 JSON]
    F --> G[结构化 Excel]
📁 项目结构
text
version3/
├── core/                         # 后端核心模块
│   ├── converter.py              # 截图转换器（VBA 增强，支持图片跳过）
│   ├── vlm_client.py             # VLM 客户端（Qwen / Mock）
│   ├── processor.py              # VLM 识别主流程（含 Prompt）
│   ├── annotator.py              # 颜色标注器（模糊匹配 + 密度过滤）
│   ├── extractor.py              # 最终 JSON 提取器
│   ├── excel_generator.py        # 结构化 Excel 生成器（多级表头扁平化）
│   └── config.py                 # API 密钥、模型配置
├── scripts/                      # 命令行工具
│   ├── run.py                    # 完整流程入口
│   ├── test_annotate.py          # 独立标注 + 提取测试
│   └── test_screenshot.py        # 截图功能测试
├── data/                         # 数据目录
│   ├── table/                    # 原始 Excel 文件（可选）
│   └── output/                   # 输出根目录（自动生成会话子目录）
├── app_st.py                     # Streamlit 前端主程序
└── requirements.txt
🚀 快速开始
环境要求
操作系统：Windows（需本地 Microsoft Excel，支持 COM 操作）

Python：3.8+

依赖安装：

bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install fuzzywuzzy python-Levenshtein -i https://pypi.tuna.tsinghua.edu.cn/simple
API 密钥配置（Qwen 模型）
powershell
# Windows PowerShell
$env:DASHSCOPE_API_KEY="你的阿里云DashScope API Key"
或在 core/config.py 中直接填写（不推荐提交）。

启动前端界面（推荐）
bash
cd version3
streamlit run app_st.py
浏览器打开 http://localhost:8501，按照向导完成上传、配置、自动处理、矫正、下载。

命令行使用（批量/脚本）
1. 仅 VLM 识别（不标注）
bash
python scripts/run.py --input data/table/test.xlsx --vlm qwen --output data/output
2. 完整流程（识别 + 标注 + 提取）
bash
python scripts/run.py --input data/table/test.xlsx --vlm qwen --annotate --output data/output
3. 截图测试（验证预处理效果）
bash
python scripts/test_screenshot.py data/table/test.xlsx --split --output test_images
4. 独立标注（基于已有 VLM JSON）
bash
python scripts/test_annotate.py
📂 输出目录说明
每次上传的 Excel 会生成一个独立会话目录（基于文件名哈希），所有中间文件保存在该目录下：

text
data/output/<session_id>/
├── images/                  # 格式化截图、标注后截图
├── vlm/                     # VLM 识别结果 JSON
├── table_mark/              # 颜色标注后的 Excel（供矫正）
├── final/                   # 最终 JSON 和结构化 Excel
└── corrected.xlsx           # 用户上传的矫正文件
⚙️ 关键模块参数
标注器（annotator.py）
threshold：模糊匹配相似度阈值（默认 80）

method：匹配方式（ratio / partial_ratio / token_sort_ratio）

density_min_count：表头行最小匹配数（默认 1）

density_min_ratio：表头行密度比例（默认 0.1）

max_gap：允许跳过非表头行的最大数量（默认 2）

截图转换器（converter.py）
apply_formatting：是否应用预处理（居中、边框、列宽自适应）

字体增强：默认关闭（若需开启，取消 converter.py 中字体设置代码的注释）

🧪 常见问题
1. 前端上传后日志无显示？
检查侧边栏是否展开，点击「清空日志」后重试。

确保 Streamlit 版本 ≥ 1.28，并使用 st.code 显示日志。

2. VLM 识别出的表头有错别字（如“外圈”→“外圆”）？
依赖人工矫正环节修正，或修改 annotator.py 添加后处理纠正映射。

3. 截图生成时提示 Excel 启动失败？
确保本地已安装 Microsoft Excel，且文件未被其他程序占用。

首次运行可能弹出 COM 缓存提示，可忽略。

4. 标注时大量数值被误标为表头？
标注器已内置密度过滤，可自动排除数据区域。若仍有误标，可适当降低 density_min_count 或提高 threshold。

5. 结构化 Excel 中多级表头路径不完整？
已实现基于几何包含的路径重建，若仍缺失，请检查原始 VLM JSON 中是否包含中间层级表头（type 为 column_header）。

🔮 未来拓展功能（规划中）
多子表结构化：自动识别并拆分嵌入的多个子表格

智能问答：基于表格内容的自然语言问答

数据可视化：自动生成统计图表和报告

📄 许可证
MIT License