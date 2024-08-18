import matplotlib.pyplot as plt
import matplotlib
from datetime import datetime

matplotlib.rcParams['font.family'] = 'Malgun Gothic'
matplotlib.rcParams['font.size'] = 20
matplotlib.rcParams['axes.unicode_minus'] = False


def stack_graph(df2, df3):
    # 누적 막대 그래프 그리기
    fig, ax = plt.subplots(figsize=(8, 12))

    for i, col in enumerate(df3.columns):
        if col == 'Sum':
            continue
        p = ax.bar(0, df3[col], bottom=df3.iloc[:, :i].sum(axis=1), label=col, color='#B0BEC5', alpha=1 - 0.4 * i, width=0.7)
        
        # 막대 레이블 포맷 설정
        def format_label(value):
            return f'{value:,}'
        
        labels = [format_label(rect.get_height()) for rect in p]
        ax.bar_label(p, labels=labels, label_type='center')
        
    # Sum 값 표시
    for i, col in enumerate(df3['Sum']):
        col = round(col, 2)
        plt.text(i, col + (col * 0.03), f'${col:,.2f}', ha='center')

    # Label 표시
    current = 0
    heights = []
    for i, col in enumerate(df2['Sum[$]']):
        col = round(col, 2)
        current += col
        if i == 0:
            heights.append(current / 2)
        elif i == len(df2['Sum[$]']) - 1:
            heights.append(col + (col * 0.03))
        else:
            heights.append(current - (col / 2))
                            
    for i, idx in enumerate(df2.index):
        plt.text(-0.47, heights[i], idx)
        
    plt.axis('off')

    today_date = datetime.today().strftime('%Y%m%d')
    plt.savefig('BOM_Cost_Graph_'+ f'{today_date}.png')

