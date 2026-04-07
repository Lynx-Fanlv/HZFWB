import pandas as pd

def calculate_dropout_rate(file_path):
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
        all_months = sorted(group_df['period'].unique())
        for i in range(2, len(all_months)):
            curr_m = all_months[i]
            p1, p2 = all_months[i-1], all_months[i-2]

            pts_t2 = set(group_df[group_df['period'] == p2]['ID'])
            pts_recent = set(group_df[group_df['period'].isin([p1, curr_m])]['ID'])

            dropout_pts = pts_t2 - pts_recent
            base_cnt = len(pts_t2)
            rate = len(dropout_pts) / base_cnt if base_cnt > 0 else 0

            results.append({
                '药店': store_name,
                '药品名称': medic_name,
                '月份': str(curr_m),
                '基准月(T-2)购药人数': base_cnt,
                '脱落人数': len(dropout_pts),
                '脱落率': rate
            })

    return pd.DataFrame(results).sort_values(['药店', '药品名称', '月份'])
