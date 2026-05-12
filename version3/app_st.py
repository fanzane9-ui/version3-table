import streamlit as st
import os
import json
import time
import shutil
import hashlib
from pathlib import Path
import sys
import uuid

sys.path.insert(0, str(Path(__file__).parent))

from core.converter import DirectHeaderConverter
from core.vlm_client import create_vlm_client
from core.annotator import ExcelAnnotator
from core.extractor import ExcelExtractor
from core.processor import HeaderVLMProcessor
from core.excel_generator import convert_json_to_structured_excel

st.set_page_config(page_title="复杂表格智能分析系统", layout="wide")
st.title("📊 复杂表格智能分析系统")


# 自定义CSS
st.markdown("""
<style>
    .reportview-container .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    .guidance-text {
        font-size: 0.9rem;
        color: #6c757d;
        margin-top: 0.2rem;
    }
    .info-box {
        background-color: #e9ecef;
        padding: 1rem;
        border-radius: 0.5rem;
        margin-bottom: 1rem;
    }
    .future-card {
        background-color: #f8f9fa;
        border-left: 5px solid #4caf50;
        padding: 0.8rem;
        margin-bottom: 0.8rem;
        border-radius: 0.3rem;
    }
</style>
""", unsafe_allow_html=True)

PROJECT_ROOT = Path(__file__).parent
DATA_OUTPUT = PROJECT_ROOT / "data" / "output"
DATA_OUTPUT.mkdir(parents=True, exist_ok=True)


def clean_old_sessions(keep=5):
    sessions = sorted([d for d in DATA_OUTPUT.iterdir() if d.is_dir()], key=lambda d: d.stat().st_ctime, reverse=True)
    for old in sessions[keep:]:
        shutil.rmtree(old)


# 初始化 session state
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'session_id' not in st.session_state:
    st.session_state.session_id = None
if 'uploaded_file_path' not in st.session_state:
    st.session_state.uploaded_file_path = None
if 'output_dir' not in st.session_state:
    st.session_state.output_dir = None
if 'vlm_provider' not in st.session_state:
    st.session_state.vlm_provider = "qwen"
if 'api_key' not in st.session_state:
    st.session_state.api_key = ""
if 'model_name' not in st.session_state:
    st.session_state.model_name = "qwen3.5-plus"
if 'img_path' not in st.session_state:
    st.session_state.img_path = None
if 'marked_img_path' not in st.session_state:
    st.session_state.marked_img_path = None
if 'marked_excel' not in st.session_state:
    st.session_state.marked_excel = None
if 'final_json_path' not in st.session_state:
    st.session_state.final_json_path = None
if 'structured_excel_path' not in st.session_state:
    st.session_state.structured_excel_path = None
if 'log' not in st.session_state:
    st.session_state.log = []
if 'processing' not in st.session_state:
    st.session_state.processing = False
if 'auto_done' not in st.session_state:
    st.session_state.auto_done = False
if 'corrected_processed' not in st.session_state:
    st.session_state.corrected_processed = False


def add_log(msg):
    st.session_state.log.append(msg)


# 侧边栏日志
with st.sidebar:
    st.subheader("处理日志")
    if st.session_state.log:
        st.code("\n".join(st.session_state.log[-20:]), language="text")
    else:
        st.info("暂无日志")
    if st.button("清空日志"):
        st.session_state.log = []
        st.rerun()

# 步骤导航
steps = ["1. 上传文件", "2. VLM配置", "3. 自动识别与标注", "4. 人工矫正", "5. 下载结果"]
current_step = st.session_state.step
cols = st.columns(len(steps))
for i, step_name in enumerate(steps):
    cols[i].button(step_name, disabled=(current_step == i + 1),
                   on_click=lambda s=i + 1: setattr(st.session_state, 'step', s))
st.divider()

# 步骤1：上传文件
if st.session_state.step == 1:
    st.subheader("第一步：上传 Excel 文件")
    st.markdown("请上传需要分析的 Excel 表格文件（支持 .xlsx, .xls, .xlsm）。")
    uploaded_file = st.file_uploader("选择 Excel 文件", type=["xlsx", "xls", "xlsm"])
    if uploaded_file is not None:
        file_stem = Path(uploaded_file.name).stem
        session_id = hashlib.md5(file_stem.encode()).hexdigest()[:8]
        session_dir = DATA_OUTPUT / session_id
        session_dir.mkdir(parents=True, exist_ok=True)
        clean_old_sessions(keep=5)
        file_path = session_dir / uploaded_file.name
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.session_state.session_id = session_id
        st.session_state.uploaded_file_path = str(file_path)
        st.session_state.output_dir = str(session_dir)
        st.session_state.img_path = None
        st.session_state.marked_img_path = None
        st.session_state.marked_excel = None
        st.session_state.final_json_path = None
        st.session_state.structured_excel_path = None
        st.session_state.auto_done = False
        st.session_state.corrected_processed = False
        add_log(f"已上传文件: {uploaded_file.name} 保存至 {session_dir}")
        st.success(f"已上传: {uploaded_file.name} (会话ID: {session_id})")
        if st.button("下一步", key="step1_next"):
            st.session_state.step = 2
            st.rerun()

# 步骤2：VLM配置
elif st.session_state.step == 2:
    if not st.session_state.uploaded_file_path:
        st.warning("请先上传文件")
        st.session_state.step = 1
        st.rerun()
    st.subheader("第二步：VLM 配置")
    st.markdown("请选择视觉语言模型（VLM）提供商，并填写必要的 API 密钥（如使用 Qwen）。")
    provider = st.radio("VLM 提供商", ["qwen", "mock"], index=0 if st.session_state.vlm_provider == "qwen" else 1)
    st.session_state.vlm_provider = provider
    if provider == "qwen":
        api_key = st.text_input("API Key", type="password", value=st.session_state.api_key,
                                help="请从阿里云百炼平台获取 DashScope API Key")
        model_name = st.text_input("模型名称", value=st.session_state.model_name,
                                   help="例如 qwen-vl-plus 或 qwen3.5-plus")
        st.session_state.api_key = api_key
        st.session_state.model_name = model_name
    if st.button("下一步", key="step2_next"):
        st.session_state.step = 3
        st.rerun()

# 步骤3：自动识别与标注
elif st.session_state.step == 3:
    if not st.session_state.uploaded_file_path:
        st.warning("请先上传文件")
        st.session_state.step = 1
        st.rerun()
    st.subheader("第三步：自动识别与标注")
    st.markdown("点击下方按钮，系统将自动生成截图、调用 VLM 识别表头、并对 Excel 进行颜色标注。")

    preview_placeholder = st.empty()
    if st.session_state.img_path and os.path.exists(st.session_state.img_path):
        preview_placeholder.image(st.session_state.img_path, caption="格式化截图", width='stretch')
    else:
        preview_placeholder.info("截图将在此处显示")

    if st.button("开始自动处理", disabled=st.session_state.processing):
        st.session_state.processing = True
        progress_bar = st.progress(0, text="准备就绪")
        status_text = st.empty()
        try:
            # 1. 截图
            status_text.info("📸 生成格式化截图...")
            progress_bar.progress(10)
            with DirectHeaderConverter(apply_formatting=True) as converter:
                if not converter.app:
                    st.error("Excel 启动失败")
                    st.stop()
                img_path = converter.convert_full_table(st.session_state.uploaded_file_path,
                                                        st.session_state.output_dir, suffix="")
            st.session_state.img_path = img_path
            add_log("✅ 格式化截图已生成")
            preview_placeholder.image(img_path, caption="格式化截图", width='stretch')
            progress_bar.progress(25)
            status_text.success("截图完成")
            time.sleep(0.5)

            # 2. VLM识别
            status_text.info("🤖 调用 VLM 识别表头...")
            progress_bar.progress(30)
            vlm_client = create_vlm_client(st.session_state.vlm_provider,
                                           st.session_state.api_key if st.session_state.api_key else None)
            if st.session_state.vlm_provider == 'qwen' and st.session_state.model_name:
                vlm_client.model = st.session_state.model_name
            proc = HeaderVLMProcessor(vlm_provider=st.session_state.vlm_provider, api_key=st.session_state.api_key)
            prompt = proc._get_end_to_end_prompt()
            for i in range(10, 90, 10):
                progress_bar.progress(30 + i // 2, text=f"识别中... {i}%")
                time.sleep(0.2)
            headers = vlm_client.analyze_image(img_path, prompt)
            if not headers:
                st.error("VLM 未返回任何表头")
                st.stop()
            add_log(f"✅ 识别成功，共 {len(headers)} 个表头")
            progress_bar.progress(70)
            status_text.success("识别完成")
            vlm_result = {
                "file": os.path.basename(st.session_state.uploaded_file_path),
                "headers": headers,
                "total_headers_found": len(headers)
            }
            vlm_dir = os.path.join(st.session_state.output_dir, "vlm")
            os.makedirs(vlm_dir, exist_ok=True)
            base_name = os.path.splitext(os.path.basename(st.session_state.uploaded_file_path))[0]
            vlm_json_path = os.path.join(vlm_dir, f"{base_name}_e2e_vlm.json")
            with open(vlm_json_path, 'w', encoding='utf-8') as f:
                json.dump(vlm_result, f, ensure_ascii=False, indent=2)
            time.sleep(0.5)

            # 3. 颜色标注
            status_text.info("🎨 开始颜色标注...")
            progress_bar.progress(75)
            annotator = ExcelAnnotator(threshold=80, method='ratio', strict=True)
            mark_folder = os.path.join(st.session_state.output_dir, "table_mark")
            os.makedirs(mark_folder, exist_ok=True)
            marked_excel = os.path.join(mark_folder, os.path.basename(st.session_state.uploaded_file_path))
            annotator.annotate_excel(st.session_state.uploaded_file_path, vlm_json_path, marked_excel)
            st.session_state.marked_excel = marked_excel
            add_log("✅ 标注完成")
            progress_bar.progress(85)
            status_text.success("标注完成")
            time.sleep(0.5)

            # 4. 生成标注后截图
            status_text.info("📸 生成标注后截图...")
            with DirectHeaderConverter(apply_formatting=True) as converter:
                marked_img_path = converter.convert_full_table(marked_excel, st.session_state.output_dir,
                                                               suffix="marked")
            st.session_state.marked_img_path = marked_img_path
            add_log("✅ 标注后截图已生成")
            preview_placeholder.image(marked_img_path, caption="标注后的表格（红色列头/绿色行头）", width='stretch')

            progress_bar.progress(100)
            status_text.success("自动处理完成！")
            st.balloons()
            st.session_state.auto_done = True
            st.session_state.processing = False
            st.rerun()
        except Exception as e:
            add_log(f"❌ 错误: {str(e)}")
            st.error(str(e))
            st.session_state.processing = False

    if st.session_state.auto_done:
        st.success("自动标注已完成，请点击下一步进入人工矫正环节。")
        # 新增：多子表结构化预告（放在第三步末尾）
        st.markdown("""
        <div class="info-box">
        <strong>🔮 后续功能预告：多子表结构化</strong><br>
        对于包含多个独立子表的复杂表格，未来版本将支持自动识别并拆分，分别生成结构化数据。
        </div>
        """, unsafe_allow_html=True)
        if st.button("下一步", key="step3_next"):
            st.session_state.step = 4
            st.rerun()
    elif not st.session_state.processing:
        st.info("点击「开始自动处理」按钮执行完整流程。")

# 步骤4：人工矫正
elif st.session_state.step == 4:
    if not st.session_state.auto_done:
        st.warning("请先完成自动处理步骤")
        st.session_state.step = 3
        st.rerun()
    st.subheader("第四步：人工矫正")

    if st.session_state.marked_img_path and os.path.exists(st.session_state.marked_img_path):
        st.image(st.session_state.marked_img_path, caption="标注后的表格预览", width='stretch')
    else:
        st.image(st.session_state.img_path, caption="格式化截图", width='stretch')

    st.markdown("""
    <div class="info-box">
    <strong>📌 矫正说明：</strong><br>
    • 下载下方的“标注后的 Excel”文件，使用 Microsoft Excel 打开。<br>
    • <span style="color:red;">红色背景</span>的单元格表示 VLM 识别为<strong>列头（column header）</strong>，通常位于表格顶部。<br>
    • <span style="color:green;">绿色背景</span>的单元格表示 VLM 识别为<strong>行头（row header）</strong>，通常位于表格左侧。<br>
    • 请检查颜色是否正确：<br>
      - 如果某个表头颜色错误，可以手动修改单元格背景色（右键 → 设置单元格格式 → 填充）。<br>
      - 如果 VLM 遗漏了某个表头，请手动将其背景色设置为对应颜色。<br>
      - 如果某个非表头单元格被错误标色，请清除其背景色。<br>
    • 保存修改后的 Excel 文件（建议使用原文件名或任意名称）。<br>
    • 然后回到本页面，在下方上传矫正后的 Excel 文件。
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 1. 下载标注后的 Excel")
    if st.session_state.marked_excel:
        with open(st.session_state.marked_excel, "rb") as f:
            st.download_button("📥 下载标注后的 Excel", f, file_name=os.path.basename(st.session_state.marked_excel))

    st.markdown("---")
    st.markdown("### 2. 上传矫正后的 Excel 文件")
    corrected_file = st.file_uploader("选择矫正后的 Excel 文件", type=["xlsx", "xls", "xlsm"], key="corrected_upload")

    if corrected_file is not None and not st.session_state.corrected_processed:
        corrected_path = Path(st.session_state.output_dir) / "corrected.xlsx"
        with open(corrected_path, "wb") as f:
            f.write(corrected_file.getbuffer())
        add_log(f"已上传矫正文件: {corrected_file.name}")
        with st.spinner("正在提取最终 JSON 和结构化 Excel..."):
            try:
                final_folder = Path(st.session_state.output_dir) / "final"
                final_folder.mkdir(exist_ok=True)
                extractor = ExcelExtractor(debug=False)
                cells_info = extractor.extract_from_workbook(str(corrected_path))
                final_json_path = final_folder / f"{Path(corrected_file.name).stem}.json"
                extractor.save_to_json(cells_info, str(final_json_path))
                st.session_state.final_json_path = str(final_json_path)
                add_log(f"✅ 最终 JSON 已生成")

                structured_excel_path = final_folder / f"{Path(corrected_file.name).stem}_structured.xlsx"
                convert_json_to_structured_excel(str(final_json_path), str(structured_excel_path), use_qwen=False)
                st.session_state.structured_excel_path = str(structured_excel_path)
                add_log(f"✅ 结构化 Excel 已生成")
                st.session_state.corrected_processed = True
                st.success("处理完成！")
                st.rerun()
            except Exception as e:
                add_log(f"❌ 处理失败: {str(e)}")
                st.error(f"处理失败: {str(e)}")

    if st.session_state.corrected_processed:
        st.success("矫正文件已处理，点击下一步下载结果。")
        if st.button("下一步", key="step4_next"):
            st.session_state.step = 5
            st.rerun()
    elif not st.session_state.corrected_processed:
        st.info("请上传矫正后的 Excel 文件以继续。")

# 步骤5：下载结果
elif st.session_state.step == 5:
    if not st.session_state.corrected_processed:
        st.warning("请先完成人工矫正步骤")
        st.session_state.step = 4
        st.rerun()
    st.subheader("第五步：下载最终结果")
    st.markdown("以下是根据您矫正后的表格生成的结构化数据文件。")

    col1, col2 = st.columns(2)
    if st.session_state.final_json_path and os.path.exists(st.session_state.final_json_path):
        with col1:
            with open(st.session_state.final_json_path, "rb") as f:
                st.download_button("📄 下载最终 JSON", f, file_name=os.path.basename(st.session_state.final_json_path))
            st.markdown("<div class='guidance-text'>包含每个单元格的文本、行列位置和类型。</div>", unsafe_allow_html=True)
    if st.session_state.structured_excel_path and os.path.exists(st.session_state.structured_excel_path):
        with col2:
            with open(st.session_state.structured_excel_path, "rb") as f:
                st.download_button("📊 下载结构化 Excel", f,
                                   file_name=os.path.basename(st.session_state.structured_excel_path))
            st.markdown("<div class='guidance-text'>将多级表头展开为单级路径，数据规整，便于入库或查询。</div>",
                        unsafe_allow_html=True)

    # 未来功能预告（仅保留智能问答和数据可视化）
    st.markdown("---")
    st.markdown("### 🚀 后续拓展功能（敬请期待）")
    st.markdown("以下功能正在开发中，未来将集成到本系统：")

    future_cols = st.columns(2)
    with future_cols[0]:
        st.markdown("""
        <div class="future-card">
        <strong>💬 智能问答</strong><br>
        <span style="font-size:0.85rem;">基于表格内容进行自然语言问答，支持统计、查询、推理等。</span>
        </div>
        """, unsafe_allow_html=True)
        st.button("💬 智能问答", disabled=True, key="future_qa")
    with future_cols[1]:
        st.markdown("""
        <div class="future-card">
        <strong>📊 数据可视化</strong><br>
        <span style="font-size:0.85rem;">自动生成图表和统计报告，直观展示数据分布与趋势。</span>
        </div>
        """, unsafe_allow_html=True)
        st.button("📊 数据可视化", disabled=True, key="future_viz")

    st.caption("以上功能正在开发中，欢迎关注后续更新。")

    st.markdown("---")
    if st.button("重新开始"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()