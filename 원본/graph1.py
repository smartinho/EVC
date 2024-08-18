import matplotlib.pyplot as plt
import matplotlib
from datetime import datetime
import re
matplotlib.rcParams['font.family'] = 'Malgun Gothic' # Windows
matplotlib.rcParams['font.size'] = 15 # 글자 크기
matplotlib.rcParams['axes.unicode_minus'] = False # 한글 폰트 사용 시, 마이너스 글자가 깨지는 현상을 해결


def stack_graph(df2, df3):

    # 누적 막대 그래프 그리기
    fig, ax = plt.subplots(figsize=(7, 10))

    for i, col in enumerate(df3.columns):
        if col == 'Sum': continue
        p = ax.bar(0, df3[col], bottom=df3.iloc[:, :i].sum(axis=1), label=col, color='#B0BEC5', alpha=1 - 0.4 * i, width=0.7)
        
        # 레이블 포맷 설정 함수
        def format_label(value):
            return f'{value:,}'
    
        labels = []
        for rect in p:
            height = rect.get_height()
            formatted_number = format_label(height)
            labels.append(formatted_number)

        ax.bar_label(p, labels=labels, label_type='center')
       
        
    # Sum 값 표시
    for i, col in enumerate(df3['Sum']):
        col = round(col, 2)
        plt.text(i, col+(col * 0.03), '$'+str( "{:,}".format(col)), ha='center')

    # label 표시
    currnet = 0
    height = []
    bottom = 0
    for i, col in enumerate(df2['Sum[$]']):
        col = round(col, 2)
        currnet += col
        if i == 0 :
            height.append(currnet/2)
        elif i == len(df2['Sum[$]']) - 1:
            height.append(col + (col * 0.03))
        else:
            height.append(currnet - (col/2))
                            
    for i, idx in enumerate(df2.index):
        plt.text(bottom - 0.47, height[i], idx)
        # print([height[i], idx])
        
    plt.axis('off')

    today_date = datetime.today().strftime('%Y%m%d')
    # 파일명 생성
    file_name = f'BOM_Cost_{today_date}.png'
    # 저장
    plt.savefig(file_name)
    # 그래프 표시
    # plt.show()