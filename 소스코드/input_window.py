import tkinter as tk
from tkinter import messagebox
from tkinter import font as tkfont
import sys

valid_number = None

def is_integer(s):
    try:
        int(s)
        return True
    except ValueError:
        return False

def submit(event=None):
    input_value = entry.get()
    if is_integer(input_value):
        number = int(input_value)
        if number <= 100:
            global valid_number
            valid_number = number
            root.quit()  # root.quit()을 사용하여 메인 루프를 종료
        else:
            messagebox.showerror("Invalid input", "잘 못 입력했습니다. 1에서 100 사이의 정수를 입력하세요.")
            entry.delete(0, tk.END)  # 입력창 초기화
    else:
        messagebox.showerror("Invalid input", "잘 못 입력했습니다. 1에서 100 사이의 정수를 입력하세요.")
        entry.delete(0, tk.END)  # 입력창 초기화

def on_closing():
    if entry.get().strip() == "":
        messagebox.showerror("No Input", "입력 값이 없어 프로그램 종료합니다.")
        root.destroy()  # root.destroy()를 호출하여 창을 완전히 닫기
        sys.exit()  # 프로그램 종료
    else:
        root.quit()  # root.quit()을 사용하여 메인 루프를 종료

def create_input_window():
    global root, entry
    root = tk.Tk()
    root.title("숫자를 입력하세요.")

    # 폰트 설정
    custom_font = tkfont.Font(family="Helvetica", size=12)

    # 화면의 너비와 높이를 가져옵니다.
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # 창의 너비와 높이를 설정합니다.
    window_width = 550
    window_height = 150

    # 창의 위치를 계산합니다.
    position_top = int(screen_height / 2 - window_height / 2)
    position_right = int(screen_width / 2 - window_width / 2)

    # 창의 크기와 위치를 설정합니다.
    root.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")

    label = tk.Label(root, text="HW, ME, Packing Cost상위부터 정렬할 갯수(1~100)를 입력하세요.", font=custom_font)
    label.pack(padx=20, pady=5)

    entry = tk.Entry(root, font=custom_font)
    entry.pack(padx=30, pady=10)
    entry.bind("<Return>", submit)  # 엔터 키 이벤트를 submit 함수에 연결

    # 입력창에 커서를 활성화
    entry.focus_set()

    submit_button = tk.Button(root, text="OK", command=submit, font=custom_font)
    submit_button.pack(padx=20, pady=20)

    root.protocol("WM_DELETE_WINDOW", on_closing)  # 창 닫기 이벤트 핸들러 등록

    root.mainloop()
    root.destroy()  # root.destroy()를 호출하여 창을 완전히 닫기

    return valid_number
