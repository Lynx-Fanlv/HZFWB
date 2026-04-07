import pandas as pd

def calculate_dot(file_path):
    # 1. 读取并标准化字段
    df = pd.read_excel(file_path)
    column_mapping = {
        '药店': 'Store', 'Store': 'Store',
        '药品名称': 'Medic', 'Medic': 'Medic',
        '患者ID': 'ID', 'ID': 'ID',
        '销售时间': 'TIME', '时间': 'TIME',
        '数量': 'Quantity', '销量': 'Quantity', 'Quantity': 'Quantity'
    }
    df = df.rename(columns={c: column_mapping[c] for c in df.columns if c in column_mapping})
    
    # 清洗：只删除关键信息缺失的行，ID和数量强制转换
    df = df.dropna(subset=['Store', 'Medic', 'ID', 'TIME'])
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
    df[['Store', 'Medic', 'ID']] = df[['Store', 'Medic', 'ID']].astype(str)
    
    # 转换月份
    df['TIME'] = pd.to_datetime(df['TIME'], errors='coerce')
    df = df.dropna(subset=['TIME'])
    df['period'] = df['TIME'].dt.to_period('M')

    # --- 核心改进：分步计算，杜绝分组交叉 ---

    # 第一步：按 药店+药品+月份 计算每个月的【销量总和】和【去重患者名单】
    # 注意：这里我们存的是患者名单(set)，为了后面计算12个月去重人数
    monthly_agg = df.groupby(['Store', 'Medic', 'period']).agg({
        'Quantity': 'sum',
        'ID': lambda x: set(x.unique())
    }).reset_index()

    results = []

    # 第二步：对每一组 (药店+药品) 独立进行12个月滑动窗口计算
    unique_groups = monthly_agg.groupby(['Store', 'Medic'])

    for (store_name, medic_name), group_df in unique_groups:
        # 确保该组内按月份严格排序
        group_df = group_df.sort_values('period')
        periods = group_df['period'].tolist()
        
        # 遍历该药店该药品的所有月份
        for i in range(len(periods)):
            curr_p = periods[i]
            start_p = curr_p - 11 # 倒推12个月的起点
            
            # 仅在【当前组内】筛选过去12个月的数据行
            window_data = group_df[(group_df['period'] >= start_p) & (group_df['period'] <= curr_p)]
            
            # 计算12个月销量总和
            total_qty = window_data['Quantity'].sum()
            
            # 计算12个月去重患者总数
            # 逻辑：把过去12个月每个月的 set 合并在一起求长度
            all_window_patients = set()
            for patient_set in window_data['ID']:
                all_window_patients.update(patient_set)
            
            unique_patients_count = len(all_window_patients)
            
            # 计算 DOT
            dot = total_qty / unique_patients_count if unique_patients_count > 0 else 0
            
            results.append({
                '药店': store_name,
                '药品名称': medic_name,
                '月份': str(curr_p),
                'DOT': dot,
                '倒推12个月销量': total_qty,
                '倒推12个月去重患者数': unique_patients_count
            })

    # 3. 最终输出
    if not results:
        return pd.DataFrame(columns=['药店', '药品名称', '月份', 'DOT', '倒推12个月销量', '倒推12个月去重患者数'])
        
    return pd.DataFrame(results).sort_values(['药店', '药品名称', '月份'])

# --- 本地验证代码 ---
if __name__ == "__main__":
    # 请根据实际文件名修改
    input_p = r'C:\Users\yym\Desktop\matlabFOR_fg\data1.xlsx'
    output_p = r'C:\Users\yym\Desktop\matlabFOR_fg\验证_DOT_分组修正版.xlsx'
    
    print("正在计算 DOT (通过预聚合确保分组唯一性)...")
    result_df = calculate_dot(input_p)
    result_df.to_excel(output_p, index=False)
    print(f"验证完成！结果已保存至: {output_p}")
