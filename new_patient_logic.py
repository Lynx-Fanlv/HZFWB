import pandas as pd

def calculate_new_patient_rate(file_path):
    df = pd.read_excel(file_path)
    col_map = {'药店':'Store','Store':'Store','药品名称':'Medic','Medic':'Medic','患者ID':'ID','ID':'ID','销售时间':'TIME','时间':'TIME'}
    df = df.rename(columns={c: col_map[c] for c in df.columns if c in col_map})
    
    df = df.dropna(subset=['Store', 'Medic', 'ID', 'TIME'])
    df[['Store', 'Medic', 'ID']] = df[['Store', 'Medic', 'ID']].astype(str)
    df['TIME'] = pd.to_datetime(df['TIME'], errors='coerce')
    df = df.dropna(subset=['TIME']).copy()
    df['period'] = df['TIME'].dt.to_period('M')

    results = []
    groups = df.groupby(['Store', 'Medic'])

    for (store_name, medic_name), group_df in groups:
        # 确保按时间排序，这对“历史首次”逻辑至关重要
        group_df = group_df.sort_values('period')
        all_months = group_df['period'].unique()
        seen_patients = set()

        for curr_m in all_months:
            curr_pts = set(group_df[group_df['period'] == curr_m]['ID'])
            new_pts = curr_pts - seen_patients
            
            total_cnt = len(curr_pts)
            new_cnt = len(new_pts)
            rate = new_cnt / total_cnt if total_cnt > 0 else 0
            
            results.append({
                '药店': store_name,
                '药品名称': medic_name,
                '月份': str(curr_m),
                '购药总人数': total_cnt,
                '新患人数(历史首购)': new_cnt,
                '新患率': rate
            })
            seen_patients.update(curr_pts)

    return pd.DataFrame(results).sort_values(['药店', '药品名称', '月份'])
