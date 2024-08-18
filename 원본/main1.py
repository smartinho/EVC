import tkinter as tk
from tkinter import messagebox
from tkinter import Label
from tkinter import filedialog
import time
import win32com.client
import dataframe
import file_save

# 프로그램 시작 시간 기록
start_time = time.time()

def open_file_explorer():
    # Tkinter 루트 윈도우 생성 및 숨기기
    root = tk.Tk()
    root.withdraw()

    # 파일 다이얼로그 열기
    file_path = filedialog.askopenfilename(
        title="파일 선택",
        filetypes=[("All Files", "*.*")]
    )

    # 파일이 선택된 경우
    if file_path:    
        # 파일 경로를 변수에 저장
        selected_file_path = file_path
        return(selected_file_path)
    else:
        print("파일이 선택되지 않았습니다.")
        message_box("파일이 선택되지 않았습니다. 프로그램을 종료합니다.")
        # on_closing()

def save_and_close_excel():
    # 엑셀 애플리케이션 객체 생성
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False  # 엑셀 창을 보이지 않게 설정

    try:
        # 열려 있는 모든 워크북 확인
        for workbook in excel_app.Workbooks:
            workbook.Close(SaveChanges=True)
            excel_app.Quit()
            return
    except Exception as e:
        print(f"Error: {e}")
    finally:
        # 엑셀 애플리케이션 종료
        excel_app.Quit()

def message_box(text):
    messagebox.showinfo("알림", f"{text}")

# 스플래시 화면
def create_splash_screen(root):

    splash = tk.Toplevel()
    splash.overrideredirect(True)
    
    # 스플래시 화면의 크기
    splash_width = 300
    splash_height = 200
    
    # 화면의 크기
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    # 스플래시 화면을 중앙에 위치시키기 위한 좌표 계산
    x = int((screen_width / 2) - (splash_width / 2))
    y = int((screen_height / 2) - (splash_height / 2))
    
    splash.geometry(f"{splash_width}x{splash_height}+{x}+{y}")
    
    label = Label(splash, text="Loading, please wait...", font=("Arial", 16))
    label.pack(expand=True)

    return splash

# 메인 함수
def main():
    root = tk.Tk()
    root.withdraw()  # 메인 윈도우 숨기기

    splash = create_splash_screen(root)
    root.update()  # 스플래시 화면 업데이트

    time.sleep(2)  # 초기 로딩 시간을 시뮬레이션

    splash.destroy()  # 스플래시 화면 닫기
    root.destroy()

if __name__ == "__main__":
    main()

    # 파일 저장/닫기
    save_and_close_excel()

    # 파일 탐색기 열기 함수 호출
    selected_file = open_file_explorer()
    
    # Dataframe 만들기
    dataframe.dataframe_read(selected_file)

    # Dataframe 만들기
    file_save.TBD_Cost(selected_file)
    
    # 프로그램 종료 시간 기록
    end_time = time.time()
    elapsed_time = end_time - start_time
    # print(f"프로그램 경과 시간: {elapsed_time:.2f} 초")
    message_box(f"완료 되었습니다. 경과 시간: {elapsed_time:.2f} 초")


    
