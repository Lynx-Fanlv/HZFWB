import pandas as pd
import numpy as np

def calculate_medic_repurchase(file_path, output_path):
    """
    计算每种药物在每个月的复购率。
    逻辑：
    - 基准月：当前月 - 2
    - 观察窗口：(当前月 - 1) 和 (当前月)
    - 复购率 = 在基准月出现且在观察窗口也出现的ID数 / 基准月总ID数
    """
    # 1. 读取数据
    # 假设 Excel 中列名为 ID, TIME, Medic
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        return f"读取文件失败: {e}"

    # 2. 预处理：确保时间格式正确并提取年月
    df['TIME'] = pd.to_datetime(df['TIME'])
    df['year'] = df['TIME'].dt.year
    df['month'] = df['TIME'].dt.month
    
    # 为了方便计算月份偏移，创建一个 Period 格式
    df['period'] = df['TIME'].dt.to_period('M')

    unique_medics = df['Medic'].unique()
    results = []

    # 3. 按药物循环处理
    for medic in unique_medics:
        medic_df = df[df['Medic'] == medic].copy()
        
        # 获取该药物涉及的所有月份范围
        start_period = medic_df['period'].min()
        end_period = medic_df['period'].max()
        # 生成完整的月份序列（类似 MATLAB 的 dateRange）
        all_months = pd.period_range(start=start_period, end=end_period, freq='M')

        # 循环月份计算（从第3个月开始，对应 MATLAB 的 i=3）
        for i in range(2, len(all_months)):
            current_month = all_months[i]
            follow_window = [all_months[i-1], all_months[i]] # 前推1月和当月
            base_month = all_months[i-2]                     # 前推2月

            # 筛选基准月用户
            base_patients = medic_df[medic_df['period'] == base_month]['ID'].unique()
            total_base_count = len(base_patients)

            # 筛选观察窗口用户
            follow_patients = medic_df[medic_df['period'].isin(follow_window)]['ID'].unique()

            # 计算复购
            repurchase_count = 0
            repurchase_rate = 0.0

            if total_base_count > 0:
                # 计算交集：在基准月且在观察窗口出现的人
                repurchase_patients = np.intersect1d(base_patients, follow_patients)
                repurchase_count = len(repurchase_patients)
                repurchase_rate = repurchase_count / total_base_count

            # 存储结果
            results.append({
                'Medic': medic,
                '年': current_month.year,
                '月': current_month.month,
                '复购人数': repurchase_count,
                '基准月购药人数': total_base_count,
                '复购率': repurchase_rate
            })

    # 4. 转换为 DataFrame 并导出
    result_table = pd.DataFrame(results)
    
    if not result_table.empty:
        result_table.to_excel(output_path, index=False)
        print(f"文件成功写入: {output_path}")
    else:
        print("未生成任何计算数据。")
        
    return result_table

# --- 使用示例 ---
#input_file = r'C:\Users\yym\Desktop\matlabFOR_fg\data1.xlsx'
#output_file = r'C:\Users\yym\Desktop\matlabFOR_fg\all_medic_repurchase_data_python.xlsx'

# 记得调用函数
#calculate_medic_repurchase(input_file, output_file)