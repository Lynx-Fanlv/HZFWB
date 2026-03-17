import pandas as pd

def calculate_dot(file_path):
    df = pd.read_excel(file_path)
    if 'MEDIC' in df.columns: df = df.rename(columns={'MEDIC': 'Medic'})
    df = df.dropna(subset=['ID', 'TIME', 'Medic', 'Quantity'])
    df['Medic'] = df['Medic'].astype(str)
    df['TIME'] = pd.to_datetime(df['TIME'], errors='coerce')
    df = df.dropna(subset=['TIME'])
    df['period'] = df['TIME'].dt.to_period('M')

    results = []
    # 【关键修改】：排序
    unique_medics = sorted(df['Medic'].unique())

    for medic in unique_medics:
        medic_df = df[df['Medic'] == medic].copy()
        all_months = sorted(medic_df['period'].unique())
        for current_month in all_months:
            start_month = current_month - 11
            mask = (medic_df['period'] >= start_month) & (medic_df['period'] <= current_month)
            window_data = medic_df[mask]
            total_qty = window_data['Quantity'].sum()
            unique_patients = window_data['ID'].nunique()
            dot = total_qty / unique_patients if unique_patients > 0 else 0
            results.append({
                '药品名称': medic,
                '月份': str(current_month),
                'DOT': dot,
                '倒推12个月销量': total_qty,
                '倒推12个月去重患者数': unique_patients
            })
    return pd.DataFrame(results)