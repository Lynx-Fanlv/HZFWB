import streamlit as st
import pandas as pd
import tempfile
import os
from repurchase_analysis import calculate_medic_repurchase  # 从你的脚本导入函数

st.set_page_config(page_title="药物复购率计算器", layout="wide")
st.title("💊 药物复购率分析工具")
st.markdown("上传包含购药记录的Excel文件，自动计算每种药物在每个月的复购率。")

# 文件上传组件
uploaded_file = st.file_uploader("请上传Excel文件", type=["xlsx", "xls"])

if uploaded_file is not None:
    # 显示上传成功
    st.success(f"已上传文件：{uploaded_file.name}")
    
    # 将上传的文件保存到临时文件（因为pandas可以直接读取BytesIO，但你的函数需要文件路径）
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
        tmp_input.write(uploaded_file.getvalue())
        tmp_input_path = tmp_input.name
    
    # 准备一个临时输出文件路径（可选，如果你的函数需要保存）
    tmp_output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
    
    try:
        # 调用你的处理函数，传入临时文件路径和临时输出路径
        # 函数会返回结果DataFrame，同时也会保存一份到tmp_output_path
        result_df = calculate_medic_repurchase(tmp_input_path, tmp_output_path)
        
        # 显示结果
        if result_df is not None and not result_df.empty:
            st.subheader("📊 计算结果预览")
            st.dataframe(result_df)
            
            # 提供下载按钮
            # 方法1：使用函数生成的临时文件
            with open(tmp_output_path, "rb") as f:
                st.download_button(
                    label="📥 下载完整结果 (Excel)",
                    data=f,
                    file_name="复购率结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # 方法2：也可以直接从result_df生成下载（无需临时文件），这里不重复
        else:
            st.warning("处理完成，但结果为空。请检查输入数据格式。")
            
    except Exception as e:
        st.error(f"处理过程中出现错误：{e}")
        st.exception(e)  # 显示详细错误信息（便于调试）
    
    finally:
        # 清理临时文件
        os.unlink(tmp_input_path)
        if os.path.exists(tmp_output_path):
            os.unlink(tmp_output_path)