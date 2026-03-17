import streamlit as st
import pandas as pd
import tempfile
import os
import importlib

st.set_page_config(page_title="患者服务部数据分析工具", layout="wide")

# 自定义CSS（可自定义配色、按钮样式等）
st.markdown("""
<style>
    /* 主标题颜色 */
    .main-title {
        color: #2c3e50;
        text-align: center;
        font-size: 2.5rem;
        margin-bottom: 1rem;
    }
    /* 副标题样式 */
    .sub-header {
        color: #34495e;
        font-size: 1.2rem;
        margin-bottom: 1.5rem;
    }
    /* 复选框容器 */
    .checkbox-container {
        background-color: #f8f9fa;
        padding: 1.5rem;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    /* 结果卡片 */
    .result-card {
        background-color: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 2rem;
    }
    /* 下载按钮样式 */
    .stDownloadButton button {
        background-color: #27ae60;
        color: white;
        border: none;
        border-radius: 5px;
        padding: 0.5rem 1rem;
        font-weight: bold;
    }
    .stDownloadButton button:hover {
        background-color: #2ecc71;
    }
</style>
""", unsafe_allow_html=True)

st.title("患者服务部数据分析工具")
st.markdown("上传Excel文件，包含TIME\Medic\Quantity\ID四个字段，选择分析类型，即可获得计算结果。")

# 定义可用的分析类型及其对应的模块名和函数名
ANALYSIS_TYPES = {
    "复购率分析": {
        "module": "repurchase_analysis",
        "function": "calculate_repurchase_rate"
    },
    "脱落分析": {
        "module": "dropout_logic",
        "function": "calculate_dropout_rate"
    },
    "DOT分析": {
        "module": "dot_logic",
        "function": "calculate_dot"
    },
    "新患分析": {
        "module": "new_patient_logic",
        "function": "calculate_new_patient_rate"
    }
}

# ========== 复选框区域（增大文字和底色） ==========
st.markdown('<div class="checkbox-container">', unsafe_allow_html=True)
st.markdown("**请选择需要分析的类型（可多选）**")

# 使用两列布局让复选框更紧凑
col1, col2 = st.columns(2)
selected_analyses = []
items = list(ANALYSIS_TYPES.items())
half = len(items) // 2
for i, (name, info) in enumerate(items):
    with col1 if i < half else col2:
        if st.checkbox(name, key=f"chk_{name}"):
            selected_analyses.append(name)
st.markdown('</div>', unsafe_allow_html=True)

# ========== 文件上传 ==========
uploaded_file = st.file_uploader("请上传Excel文件", type=["xlsx", "xls"])

# 处理上传文件：保存到临时文件，并清理旧的临时文件
if uploaded_file is not None:
    # 如果之前有临时文件，先删除
    if 'input_path' in st.session_state and os.path.exists(st.session_state['input_path']):
        os.unlink(st.session_state['input_path'])
    # 保存新上传的文件
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
        tmp_input.write(uploaded_file.getvalue())
        st.session_state['input_path'] = tmp_input.name
    st.success(f"已上传文件：{uploaded_file.name}")
    # 清除之前的结果，因为上传了新文件
    if 'results' in st.session_state:
        del st.session_state['results']

# ========== 开始分析按钮 ==========
if st.button("🚀 开始分析", type="primary"):
    # 验证条件
    if 'input_path' not in st.session_state:
        st.warning("请先上传文件")
    elif not selected_analyses:
        st.warning("请至少选择一种分析类型")
    else:
        # 执行所有选中的分析
        results = {}
        input_path = st.session_state['input_path']
        
        # 显示加载提示（可选）
        with st.spinner("正在计算中，请稍候..."):
            for analysis_name in selected_analyses:
                try:
                    module_info = ANALYSIS_TYPES[analysis_name]
                    module = importlib.import_module(module_info["module"])
                    func = getattr(module, module_info["function"])
                    
                    # 调用函数，传入输入文件路径
                    result_df = func(input_path)
                    
                    if result_df is not None and not result_df.empty:
                        results[analysis_name] = result_df
                    else:
                        # 如果结果为空，也存入一个空值，以便显示提示
                        results[analysis_name] = None
                except Exception as e:
                    st.error(f"{analysis_name} 计算失败：{e}")
                    # 可选择将异常信息存入结果
                    results[analysis_name] = None
        
        # 将结果存入 session_state
        st.session_state['results'] = results

# ========== 显示结果 ==========
if 'results' in st.session_state and st.session_state['results']:
    for analysis_name, result_df in st.session_state['results'].items():
        with st.container():
            st.markdown('<div class="result-card">', unsafe_allow_html=True)
            st.subheader(f"📈 {analysis_name} 结果")
            
            if result_df is not None and not result_df.empty:
                st.dataframe(result_df)
                
                # 提供下载
                output_filename = f"{analysis_name}_结果.xlsx"
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_output:
                    result_df.to_excel(tmp_output.name, index=False)
                    tmp_output_path = tmp_output.name
                
                with open(tmp_output_path, "rb") as f:
                    st.download_button(
                        label=f"📥 下载 {analysis_name} 结果",
                        data=f,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{analysis_name}"
                    )
                
                # 清理临时输出文件
                os.unlink(tmp_output_path)
            else:
                st.warning(f"{analysis_name} 计算完成，但结果为空。请检查输入数据格式。")
            
            st.markdown('</div>', unsafe_allow_html=True)

# ========== 页面底部清理（可选） ==========
# 注意：输入临时文件的清理在上传新文件时已经处理，但为了防止程序退出时残留，
# 可以添加一个清理函数，但Streamlit无直接支持，我们依靠下次上传时清理。