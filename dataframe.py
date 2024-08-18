import pandas as pd
import file_save
import graph

param_value = None

def set_param_value(value):
    global param_value
    param_value = value
    
def code_comp(df_code, df_col):
    df_code = df_code.drop_duplicates(subset=['Code'], keep='first')
    df_comp = df_col.merge(df_code[['Code', '분류']], left_on='Code_3', right_on='Code', how='left').drop(columns=['Code'])
    return df_comp

def dataframe_read(file):
    try:
        df = pd.read_excel(file, skiprows=19)
        required_columns = ['Part No', 'Desc.', 'Spec.', 'Unit Price (USD)']
        if not all(col in df.columns for col in required_columns):
            raise ValueError("BOM_Cost 파일이 아닙니다.")
    except Exception as e:
        print(f"Error reading the file: {e}")
        return None, None

    df_comp = sort_col(df, 0)
    if df_comp is None:
        return None, None
    
    return df_comp, None  # Ensure two values are returned

def sort_col(df, count):
    selected_columns = ['Model.Suffix', 'Seq.', 'Lvl', 'Part No', 'Supply Type', 'Desc.', 'Spec.', 'Class Code', 'Class Name', 'Material Cost (KRW)', 'Material Cost (USD)', 'Curr (All)', 'Exchange Rate (All)']
    df_col = df[selected_columns]
    df_col['Code_3'] = df_col['Class Code'].str[:3]

    if count == 0:
        global model_name, rounded_usd_exchange_rate
        model_name = df_col['Model.Suffix'].str[0:11].iloc[3]
        usd_exchange_rate = df_col.loc[df_col['Curr (All)'] == 'USD', 'Exchange Rate (All)'].iloc[0]
        rounded_usd_exchange_rate = round(usd_exchange_rate, 2)
        print('환율:', rounded_usd_exchange_rate)

    df_code = pd.read_excel('code.xlsx')
    if 'Code' not in df_code.columns:
        pass

    df_comp = code_comp(df_code, df_col)
    df_comp = add_cols(df_comp, count)
    return df_comp

def add_cols(df_comp, count):
    df_comp['Lvl'] = df_comp['Lvl'].apply(lambda x: str(x).replace('.', '') if pd.notnull(x) else x)
    df_comp['Lvl'] = df_comp['Lvl'].fillna(0).astype(int)

    if count == 0:
        level_sum(df_comp)
        return None, None  # Ensure two values are returned
    else:
        return df_comp[['분류', 'Part No', 'Desc.', 'Spec.', 'Material Cost (USD)']], None

def level_sum(df_level1):   
    df_level1['Cost[$]'] = 0
    df_level1['Material Cost (USD)'] = df_level1['Material Cost (USD)'].fillna(0)

    i = 0
    while i < len(df_level1):
        if df_level1.at[i, 'Lvl'] == 1:
            start_idx = i
            while i + 1 < len(df_level1) and df_level1.at[i + 1, 'Lvl'] != 1:
                i += 1
            end_idx = i
            
            if df_level1.at[start_idx, 'Material Cost (USD)'] == 0 and df_level1.at[start_idx, 'Part No'].startswith('ACQ'):
                for j in range(start_idx + 1, end_idx):
                    if df_level1.at[j, 'Part No'].startswith('E'):
                        cost_sum = df_level1.loc[j:end_idx, 'Material Cost (USD)'].sum()
                        df_level1.at[j, 'Cost[$]'] = cost_sum
                        break
                    else:
                        cost_sum = df_level1.loc[start_idx:end_idx, 'Material Cost (USD)'].sum()
                        df_level1.at[start_idx, 'Cost[$]'] = cost_sum
            else:
                cost_sum = df_level1.loc[start_idx:end_idx, 'Material Cost (USD)'].sum()
                df_level1.at[start_idx, 'Cost[$]'] = cost_sum        

        i += 1

    df_level1['Cost[$]'] = round(df_level1['Cost[$]'], 2)
    sort_class(df_level1)

def sort_class(df_result):
    df_result = df_result.groupby(['분류', 'Part No', 'Desc.', 'Spec.'], as_index=False)['Cost[$]'].sum()
    last_grouping(df_result)

def last_grouping(df_result):
    sorted_df = df_result.sort_values(by=['Cost[$]'], ascending=False)
    global param_value, html_table
    top_df = sorted_df.groupby('분류').head(param_value)
    other_df = sorted_df[~sorted_df.index.isin(top_df.index)]
    other_grouped = other_df.groupby('분류').agg({'Cost[$]': 'sum'}).reset_index()
    other_grouped['분류'] = other_grouped['분류'] + "_Etc"
    summary_df = sorted_df.groupby('분류').agg({'Cost[$]': 'sum'}).reset_index()
    summary_df['분류'] = summary_df['분류'] + "_Sum"
    final_df = pd.concat([top_df, other_grouped, summary_df], ignore_index=True).sort_values(by=['분류', 'Cost[$]'], ascending=[True, False]).reset_index(drop=True)
    cost_sum = df_result['Cost[$]'].sum()
    total_sum_row = pd.DataFrame({'분류': ['Total_Sum'], 'Cost[$]': [cost_sum]})
    file_df = pd.concat([final_df, total_sum_row], ignore_index=True)
    file_name = file_save.save_today(file_df, model_name)
    file_save.auto_fit_columns(file_name)
    
    # 데이터프레임을 HTML 테이블로 변환
    html_df = file_df.fillna('')
    html_table = html_df.to_html(classes='table table-striped', index=False)
    total_sum(final_df)

def total_sum(final_df):
    sum_col = ['분류', 'Cost[$]']
    df_sum = final_df[sum_col]
    df_sum = df_sum.loc[df_sum['분류'].str.contains('Sum')].sort_values(by=['Cost[$]'])
    df_sum = df_sum.replace(['Packing_Sum', 'ME_Sum', 'HW_Sum', '기타_Sum'],['Packing', 'ME', 'HW', '기타'])
    df_sum = df_sum.rename(columns={'분류':'구분', 'Cost[$]':'Sum[$]'})
    df_sum['Sum[$]'] = round(df_sum['Sum[$]'], 2)
    total_sum = df_sum['Sum[$]'].sum()
    total_row = pd.DataFrame({'구분': ['Sum'], 'Sum[$]': [total_sum]})
    df_sum = pd.concat([df_sum, total_row], ignore_index=True)
    df_sum['is_sum'] = df_sum['구분'].apply(lambda x: 'Sum' in x )
    sorted_df = df_sum.sort_values(by=['구분'])
    df2 = sorted_df.sort_values(by=['is_sum', '구분'], ascending=[True, True]).drop(columns='is_sum')
    df2.set_index('구분', inplace=True)
    df2 = df2.sort_values(by='Sum[$]', ascending=True)
    df3 = df2.T
    graph.stack_graph(df2, df3)
    
# 데이터프레임을 HTML로 변환하고 스타일 적용
def highlight_sum_rows(row):
    return ['background-color: lightblue' if 'Sum' in str(cell) else '' for cell in row]


