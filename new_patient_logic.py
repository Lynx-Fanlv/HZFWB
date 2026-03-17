import pandas as pd

def calculate_new_patient_rate(file_path):
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
        medic_df = df[df['Medic'] == medic].copy()
        all_months = sorted(medic_df['period'].unique())
        seen_patients = set()
        for curr_m in all_months:
            curr_patients = set(medic_df[medic_df['period'] == curr_m]['ID'])
            new_patients = curr_patients - seen_patients
            new_count = len(new_patients)
            total_count = len(curr_patients)
            rate = new_count / total_count if total_count > 0 else 0.0
            seen_patients.update(curr_patients)
            results.append({
                '药品名称': medic,
                '月份': str(curr_m),
                '购药总人数': total_count,
                '新患人数(历史首购)': new_count,
                '新患率': rate
            })
    return pd.DataFrame(results)