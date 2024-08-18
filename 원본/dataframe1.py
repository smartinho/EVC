import pandas as pd
import file_save
import graph
import main

def code_comp(df_code, df_col):
    df_code = df_code.drop_duplicates(subset=['Code'], keep='first')
    df_comp = df_col.merge(df_code[['Code', '분류']], left_on='Code_3', right_on='Code', how='left').drop(columns=['Code'])
    return df_comp

def dataframe_read(file):
    try:
        # 데이터프레임 만들기   
        df = pd.read_excel(file, skiprows=19)
        if 'Desc.' not in df.columns:
            main.message_box("BOM_Cost 파일이 아닙니다. 종료합니다.")
            exit()
    except Exception as e:
            main.message_box("BOM_Cost파일이 아닙니다. 프로그램을 종료합니다.")
            exit()
            print(f"파일을 읽는 중 오류가 발생했습니다: {e}")
    sort_col(df)

#  컬럼 정리
def sort_col(df):
    selected_columns = ['Lvl', 'Part No','Desc.','Spec.', 'Class Code','Class Name','Material Cost (KRW)','Material Cost (USD)']
    df_col = df[selected_columns]
    df_col['Code_3'] = df_col['Class Code'].str[:3]

    # 분류를 위한 code 파일 읽기
    df_code = pd.read_excel('code.xlsx')
    if 'Code' not in df_code.columns:
        main.message_box("code 파일이 아닙니다. 종료합니다.")
        exit()

    # Code 비교하기
    df_comp = code_comp(df_code, df_col)
    # file_save.Save_Today(df_comp)
    add_cols(df_comp)

# '분류' 열 추가
def add_cols(df_comp):
    df_comp['Lvl'] = df_comp['Lvl'].apply(lambda x: str(x).replace('.', '') if pd.notnull(x) else x)
    cost_sum(df_comp)

# Level 1 기준으로 Cost 합산
def cost_sum(df_comp):
    lvl_sum = 0  # 합산할 변수 초기화
    start_idx = None  # 첫 번째 '1' 값의 위치를 기억할 변수 초기화

    for idx, lvl in enumerate(df_comp['Lvl']):
        if lvl == '1' and start_idx is None:
            # 첫 번째 '1' 값의 위치를 기억
            start_idx = idx
            lvl_sum = float(df_comp.iloc[idx]['Material Cost (USD)']) if pd.notna(df_comp.iloc[idx]['Material Cost (USD)']) else 0
        elif lvl == '1' and start_idx is not None:
            # 이전 구간의 합산을 마무리하고 새로운 구간의 합산을 시작
            df_comp.loc[start_idx, 'Cost[$]'] = lvl_sum
            start_idx = idx
            lvl_sum = float(df_comp.iloc[idx]['Material Cost (USD)']) if pd.notna(df_comp.iloc[idx]['Material Cost (USD)']) else 0
        elif pd.notna(lvl_sum) and lvl != '1':
            # '1'이 아닌 경우 합산을 진행
            lvl_sum += float(df_comp.iloc[idx]['Material Cost (USD)']) if pd.notna(df_comp.iloc[idx]['Material Cost (USD)']) else 0

    # 마지막 구간의 합산 결과를 반영
    if start_idx is not None:
        df_comp.loc[start_idx, 'Cost[$]'] = lvl_sum

    sort_class(df_comp)   

# 분류 기준으로 정리
def sort_class(df_comp):
    # 'Lvl' 값이 '1'인 행만 선택하여 '분류' 기준으로 그룹화하고 'Cost[$]' 기준으로 내림차순 정렬
    result = df_comp[df_comp['Lvl'] == '1'].groupby(['분류', 'Part No', 'Desc.', 'Spec.'], as_index=False)['Cost[$]'].sum()

    # 소수점 2자리 반올림
    result['Cost[$]'] = round(result['Cost[$]'], 2)

    last_grouping(result)

# 최종 그룹핑
def last_grouping(result):

    # 분류 및 내림차순으로 정렬된 데이터프레임 생성
    sorted_df = result.sort_values(by=['Cost[$]'], ascending=False)

    # 분류 별 상위 5개 항목 선택
    top_df = sorted_df.groupby('분류').head(5)

    # 나머지 항목을 기타로 묶기
    other_df = sorted_df[~sorted_df.index.isin(top_df.index)]
    other_grouped = other_df.groupby('분류').agg({'Cost[$]': 'sum'}).reset_index()
    other_grouped['분류'] = other_grouped['분류'] + "_Etc"

    # 분류 Sum 추가
    summary_df = sorted_df.groupby('분류').agg({'Cost[$]': 'sum'}).reset_index()
    summary_df['분류'] = summary_df['분류'] + "_Sum"

    # 최종 데이터프레임 결합
    final_df = pd.concat([top_df, other_grouped, summary_df], ignore_index=True)

    # 최종 정렬
    final_df = final_df.sort_values(by=['분류', 'Cost[$]'], ascending=[True, False]).reset_index(drop=True)
    
    # Cost[$]열의 합 total_sum 추가
    cost_sum = result['Cost[$]'].sum()

    # #'Total_Sum' 행 추가
    total_sum_row = pd.DataFrame({'분류': ['Total_Sum'], 'Cost[$]': [cost_sum]})

    # 원본 데이터프레임에 'Total_Sum' 행 추가
    file_df = pd.concat([final_df, total_sum_row], ignore_index=True)

    # 오늘 날짜로 엑셀 저장
    file_name = file_save.Save_Today(file_df)
    file_save.auto_fit_columns(file_name)

    total_sum(final_df)

# 영역별 Sum
def total_sum(final_df):
    sum_col = ['분류', 'Cost[$]']
    df_sum = final_df[sum_col]
    df_sum = df_sum.loc[df_sum['분류'].str.contains('Sum')].sort_values(by=['Cost[$]'])
    df_sum = df_sum.replace(['Packing_Sum', 'ME_Sum', 'HW_Sum', '기타_Sum'],['Packing', 'ME', 'HW', '기타'])
    df_sum = df_sum.rename(columns={'분류':'구분', 'Cost[$]':'Sum[$]'})
    df_sum['Sum[$]'] = round(df_sum['Sum[$]'], 2)
    total_sum = df_sum['Sum[$]'].sum()

    # Sum 행 추가
    total_row = pd.DataFrame({'구분': ['Sum'], 'Sum[$]': [total_sum]})
    df_sum = pd.concat([df_sum, total_row], ignore_index=True)

    # 'Sum'이 포함된 행을 식별하는 새로운 열 추가
    df_sum['is_sum'] = df_sum['구분'].apply(lambda x: 'Sum' in x )

    # 우선 '구분' 열을 기준으로 오름차순 정렬
    sorted_df = df_sum.sort_values(by=['구분'])

    # 'is_sum' 열을 기준으로 최종 정렬
    df2 = sorted_df.sort_values(by=['is_sum', '구분'], ascending=[True, True]).drop(columns='is_sum')
    df2.set_index('구분', inplace=True)

    # 오름차순 정렬
    df2 = df2.sort_values(by='Sum[$]', ascending=True)
    transfrom_row_col(df2)

# 막대그래프를 위한 행열 변경
def transfrom_row_col(df2):
    df3 = df2.T
    graph.stack_graph(df2, df3)


