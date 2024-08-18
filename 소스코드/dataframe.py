import pandas as pd
import file_save
import graph
import main
import input_window

def code_comp(df_code, df_col):
    # 중복 제거
    df_code = df_code.drop_duplicates(subset=['Code'], keep='first')
    # 코드와 컬럼 병합
    df_comp = df_col.merge(df_code[['Code', '분류']], left_on='Code_3', right_on='Code', how='left').drop(columns=['Code'])
    return df_comp

def dataframe_read(file):
    try:
        df = pd.read_excel(file, skiprows=19)
        if 'Part No' and 'Desc.' and 'Spec.' and 'Unit Price (USD)' not in df.columns:
            main.message_box("BOM_Cost 파일이 아닙니다. 종료합니다.")
            exit()
    except Exception as e:
        main.message_box("BOM_Cost 파일이 아닙니다. 프로그램을 종료합니다.")
        exit()
    sort_col(df)

def sort_col(df):
    # 선택 컬럼 정의
    selected_columns = ['Model.Suffix', 'Seq.', 'Lvl', 'Part No', 'Supply Type', 'Desc.', 'Spec.', 'Class Code','Class Name',
                        'Material Cost (KRW)','Material Cost (USD)', 'Curr (All)', 'Exchange Rate (All)',]
    df_col = df[selected_columns]
    # Code_3 생성
    df_col['Code_3'] = df_col['Class Code'].str[:3]

    # 모델명 가져오기
    global model_name
    model_name = df_col['Model.Suffix'].str[0:11].loc[3,]

    # 환욜 가져오기
    global rounded_usd_exchange_rate
    usd_exchange_rate = df_col.loc[df_col['Curr (All)'] == 'USD', 'Exchange Rate (All)'].iloc[0]
    rounded_usd_exchange_rate = round(usd_exchange_rate, 2)
    print('환율:', rounded_usd_exchange_rate)

    # 코드 파일 읽기
    df_code = pd.read_excel('code.xlsx')
    if 'Code' not in df_code.columns:
        main.message_box("code 파일이 아닙니다. 종료합니다.")
        exit()

    df_comp = code_comp(df_code, df_col)
    add_cols(df_comp)

def add_cols(df_comp):
    # Lvl 처리
    df_comp['Lvl'] = df_comp['Lvl'].apply(lambda x: str(x).replace('.', '') if pd.notnull(x) else x)
    cost_sum(df_comp)

def cost_sum(df_comp):
    lvl_sum = 0
    start_idx = None

    for idx, lvl in enumerate(df_comp['Lvl']):
        if lvl == '1' and start_idx is None:
            start_idx = idx
            lvl_sum = float(df_comp.iloc[idx]['Material Cost (USD)']) if pd.notna(df_comp.iloc[idx]['Material Cost (USD)']) else 0
        elif lvl == '1' and start_idx is not None:
            df_comp.loc[start_idx, 'Cost[$]'] = lvl_sum
            start_idx = idx
            lvl_sum = float(df_comp.iloc[idx]['Material Cost (USD)']) if pd.notna(df_comp.iloc[idx]['Material Cost (USD)']) else 0
        elif pd.notna(lvl_sum) and lvl != '1':
            lvl_sum += float(df_comp.iloc[idx]['Material Cost (USD)']) if pd.notna(df_comp.iloc[idx]['Material Cost (USD)']) else 0

    if start_idx is not None:
        df_comp.loc[start_idx, 'Cost[$]'] = lvl_sum

    sort_class(df_comp)

def sort_class(df_comp):
    result = df_comp[df_comp['Lvl'] == '1'].groupby(['분류', 'Part No', 'Desc.', 'Spec.'], as_index=False)['Cost[$]'].sum()
    result['Cost[$]'] = round(result['Cost[$]'], 2)
    last_grouping(result)

def last_grouping(result):
    sorted_df = result.sort_values(by=['Cost[$]'], ascending=False)
    data = input_window.create_input_window()
    print(data)
    top_df = sorted_df.groupby('분류').head(data) # 상위 5개 데이터만 가져오기
    other_df = sorted_df[~sorted_df.index.isin(top_df.index)]
    other_grouped = other_df.groupby('분류').agg({'Cost[$]': 'sum'}).reset_index()
    other_grouped['분류'] = other_grouped['분류'] + "_Etc"
    summary_df = sorted_df.groupby('분류').agg({'Cost[$]': 'sum'}).reset_index()
    summary_df['분류'] = summary_df['분류'] + "_Sum"
    final_df = pd.concat([top_df, other_grouped, summary_df], ignore_index=True)
    final_df = final_df.sort_values(by=['분류', 'Cost[$]'], ascending=[True, False]).reset_index(drop=True)
    cost_sum = result['Cost[$]'].sum()
    total_sum_row = pd.DataFrame({'분류': ['Total_Sum'], 'Cost[$]': [cost_sum]})
    file_df = pd.concat([final_df, total_sum_row], ignore_index=True)
    file_name = file_save.save_today(file_df, model_name)
    file_save.auto_fit_columns(file_name)
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
    transfrom_row_col(df2)

def transfrom_row_col(df2):
    df3 = df2.T
    graph.stack_graph(df2, df3)

