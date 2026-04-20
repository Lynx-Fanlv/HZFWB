import pandas as pd
import numpy as np

# 将函数名改回网页程序调用的名称：calculate_medic_repurchase
def calculate_repurchase_rate(file_path):
    """
    计算复购率（适配网页调用名）。
    输入：Excel 文件路径
    返回：包含‘药店’和‘药品名称’维度的 DataFrame
    """
    # 1. 读取数据
    df = pd.read_excel(file_path)
    
    # 2. 字段映射 (支持 Store 维度及中文表头)
    column_mapping = {
        '药店': 'Store', 'Store': 'Store',
        '药品名称': 'Medic', 'Medic': 'Medic',
        '患者ID': 'ID', 'ID': 'ID',
        '销售时间': 'TIME', '时间': 'TIME'
    }
    df = df.rename(columns={c: column_mapping[c] for c in df.columns if c in column_mapping})
    
    # 3. 数据清洗
    df = df.dropna(subset=['Store', 'Medic', 'ID', 'TIME'])
    df[['Store', 'Medic', 'ID']] = df[['Store', 'Medic', 'ID']].astype(str)
    df['TIME'] = pd.to_datetime(df['TIME'], errors='coerce')
    df = df.dropna(subset=['TIME'])
    df['period'] = df['TIME'].dt.to_period('M')

    results = []
    
    # 4. 按 药店 + 药品 双重分组
    groups = df.groupby(['Store', 'Medic'])

    for (store_name, medic_name), group_df in groups:
        all_months = sorted(group_df['period'].unique())
        
        if len(all_months) < 3:
            continue

        for i in range(2, len(all_months)):
            curr_m = all_months[i]       
            prev_1 = all_months[i-1]    
            prev_2 = all_months[i-2]    

            # 基准月患者 (T-2)
            base_pts = set(group_df[group_df['period'] == prev_2]['ID'])
            # 观察窗口患者 (T-1 和 T)
            follow_pts = set(group_df[group_df['period'].isin([prev_1, curr_m])]['ID'])

            total_base = len(base_pts)
            repurchase_pts = base_pts & follow_pts 
            repurchase_count = len(repurchase_pts)
            
            rate = repurchase_count / total_base if total_base > 0 else 0

            results.append({
                '药店': store_name,
                '药品名称': medic_name,
                '月份': str(curr_m),
                '复购人数': repurchase_count,
                '基准月(T-2)购药人数': total_base,
                '复购率': rate
            })

    if not results:
        return pd.DataFrame(columns=['药店', '药品名称', '月份', '复购人数', '基准月(T-2)购药人数', '复购率'])
        
    return pd.DataFrame(results).sort_values(['药店', '药品名称', '月份'])

# --- 本地验证启动代码 ---
if __name__ == "__main__":
    # 验证时请确保路径正确
    input_file = r'C:\Users\yym\Desktop\matlabFOR_fg\data1.xlsx'
    output_file = r'C:\Users\yym\Desktop\matlabFOR_fg\验证_复购率_Store版.xlsx'
    
    print("正在开始复购率本地验证...")
    # 这里也同步修改调用名
    final_df = calculate_medic_repurchase(input_file)
    final_df.to_excel(output_file, index=False)
    print(f"本地验证成功！结果已保存至: {output_file}")
