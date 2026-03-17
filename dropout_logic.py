import pandas as pd

def calculate_dropout_rate(file_path):
    df = pd.read_excel(file_path)
    if 'MEDIC' in df.columns: df = df.rename(columns={'MEDIC': 'Medic'})
    df = df.dropna(subset=['ID', 'TIME', 'Medic'])
    df['ID'] = df['ID'].astype(str)
    df['Medic'] = df['Medic'].astype(str)
    df['TIME'] = pd.to_datetime(df['TIME'], errors='coerce')
    df = df.dropna(subset=['TIME'])
    df['period'] = df['TIME'].dt.to_period('M')

    results = []
    # 【关键修改】：排序
    unique_medics = sorted(df['Medic'].unique())

    for medic in unique_medics:
        medic_df = df[df['Medic'] == medic]
        all_months = sorted(medic_df['period'].unique())
        for i in range(2, len(all_months)):
            curr_m = all_months[i]
            prev_1 = all_months[i-1]
            prev_2 = all_months[i-2]
            patients_t2 = set(medic_df[medic_df['period'] == prev_2]['ID'])
            base_count_t2 = len(patients_t2)
            patients_t1_t0 = set(medic_df[medic_df['period'].isin([prev_1, curr_m])]['ID'])
            dropout_count = len(patients_t2 - patients_t1_t0)
            dropout_rate = dropout_count / base_count_t2 if base_count_t2 > 0 else 0.0
            results.append({
                '药品名称': medic,
                '月份': str(curr_m),
                '基准月(T-2)购药人数': base_count_t2,
                '脱落人数': dropout_count,
                '脱落率': dropout_rate
            })
    return pd.DataFrame(results)