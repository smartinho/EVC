import tkinter as tk
from tkinter import messagebox, Label, filedialog
import time
import win32com.client
import dataframe
import file_save

# 프로그램 시작 시간 기록
start_time = time.time()

def open_file_explorer():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="파일 선택",
        filetypes=[("All Files", "*.*")]
    )

    if file_path:
        return file_path
    else:
        message_box("파일이 선택되지 않았습니다. 프로그램을 종료합니다.")
        exit()

def save_and_close_excel():
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False

    try:
        for workbook in excel_app.Workbooks:
            workbook.Close(SaveChanges=False)
        excel_app.Quit()
    except Exception as e:
        print(f"Error: {e}")
    finally:
        excel_app.Quit()

def message_box(text):
    messagebox.showinfo("알림", text)

def create_splash_screen(root):
    splash = tk.Toplevel()
    splash.overrideredirect(True)
    splash_width, splash_height = 300, 200
    screen_width, screen_height = root.winfo_screenwidth(), root.winfo_screenheight()
    x = int((screen_width / 2) - (splash_width / 2))
    y = int((screen_height / 2) - (splash_height / 2))
    splash.geometry(f"{splash_width}x{splash_height}+{x}+{y}")
    label = Label(splash, text="Loading, please wait...", font=("Arial", 16))
    label.pack(expand=True)
    return splash

def main():
    root = tk.Tk()
    root.withdraw()

    splash = create_splash_screen(root)
    root.update()

    time.sleep(2)

    splash.destroy()
    root.destroy()

    save_and_close_excel()

    selected_file = open_file_explorer()

    dataframe.dataframe_read(selected_file)
    file_save.tbd_cost(selected_file)

    end_time = time.time()
    elapsed_time = end_time - start_time
    message_box(f"완료 되었습니다. 경과 시간: {elapsed_time:.2f} 초")

if __name__ == "__main__":
    main()
    