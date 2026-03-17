import pandas as pd
import numpy as np

def calculate_repurchase_rate(file_path):
    df = pd.read_excel(file_path)
    if 'MEDIC' in df.columns: df = df.rename(columns={'MEDIC': 'Medic'})
    df = df.dropna(subset=['ID', 'TIME', 'Medic'])
    df['ID'] = df['ID'].astype(str)
    df['Medic'] = df['Medic'].astype(str)
    df['TIME'] = pd.to_datetime(df['TIME'], errors='coerce')
    df = df.dropna(subset=['TIME'])
    df['period'] = df['TIME'].dt.to_period('M')

    results = []
    # 【关键修改】：对药品进行排序，确保输出顺序一致
    unique_medics = sorted(df['Medic'].unique())

    for medic in unique_medics:
        medic_df = df[df['Medic'] == medic].copy()
        all_months = sorted(medic_df['period'].unique())
        if len(all_months) < 3: continue

        for i in range(2, len(all_months)):
            curr_month = all_months[i]
            follow_window = [all_months[i-1], all_months[i]]
            base_month = all_months[i-2]
            
            base_patients = medic_df[medic_df['period'] == base_month]['ID'].unique()
            total_base_count = len(base_patients)
            follow_patients = medic_df[medic_df['period'].isin(follow_window)]['ID'].unique()

            repurchase_count = 0
            repurchase_rate = 0.0
            if total_base_count > 0:
                repurchase_patients = set(base_patients) & set(follow_patients)
                repurchase_count = len(repurchase_patients)
                repurchase_rate = repurchase_count / total_base_count

            results.append({
                '药品名称': medic,
                '月份': str(curr_month),
                '复购人数': repurchase_count,
                '基准月(T-2)购药人数': total_base_count,
                '复购率': repurchase_rate
            })

    return pd.DataFrame(results)