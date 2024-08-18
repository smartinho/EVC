import sys
import os
import time
import warnings
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QLabel, QFileDialog, QMessageBox, QProgressBar, QTextEdit
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import dataframe
import main
import file_save

# SettingWithCopyWarning 경고 숨기기
warnings.simplefilter(action='ignore', category=UserWarning)

class Worker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(float)

    def __init__(self, selected_file, param_value):
        super().__init__()
        self.selected_file = selected_file
        self.param_value = param_value

    def run(self):
        start_time = time.time()
        try:
            # 각 함수의 실행 시간을 측정
            total_time = (
                self.measure_time(main.save_and_close_excel) +
                self.measure_time(lambda: dataframe.dataframe_read(self.selected_file)) +
                self.measure_time(lambda: file_save.tbd_cost(self.selected_file))
            )

            # 상태바를 1부터 100까지 순차적으로 증가
            self.update_progress(total_time)
        except Exception as e:
            print(f"Error: {e}")  # 표준 출력을 UI로 리디렉션한 경우, 이 메시지가 UI에 표시됨
        end_time = time.time()
        elapsed_time = end_time - start_time
        self.finished.emit(elapsed_time)

    def measure_time(self, func):
        func_start_time = time.time()
        func()
        func_end_time = time.time()
        return func_end_time - func_start_time

    def update_progress(self, total_time):
        # 상태바를 1부터 100까지 순차적으로 증가
        for i in range(1, 101):
            self.progress.emit(i)
            time.sleep(total_time / 100)

class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # 메인 레이아웃 설정
        main_layout = QVBoxLayout()

        # 파일 선택 레이아웃 설정
        file_layout = QHBoxLayout()
        
        # 파일 선택 버튼과 라벨
        self.file_label = QLabel('1.파일 선택: BOM Cost 파일(xlsm)을 선택해 주세요!', self)
        self.file_label.setStyleSheet("QLabel { padding-top: 4px; }") 
        self.file_btn = QPushButton('파일 선택', self)
        self.file_btn.clicked.connect(self.showFileDialog)

        # 파일 선택 레이아웃에 위젯 추가
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.file_btn, alignment=Qt.AlignRight)

        # 파라미터 입력 레이아웃 설정
        param_layout = QHBoxLayout()
        self.param_label = QLabel('2.정렬 갯수:', self)
        self.param_input = QLineEdit(self)
        self.param_input.returnPressed.connect(self.validateAndRunProgram)  # 엔터 키를 누르면 프로그램 실행

        # 파라미터 입력 레이아웃에 위젯 추가
        param_layout.addWidget(self.param_label)
        param_layout.addWidget(self.param_input)

        # OK 버튼
        self.ok_btn = QPushButton('OK', self)
        self.ok_btn.clicked.connect(self.validateAndRunProgram)
        
        # OK 버튼을 중앙에 배치하기 위한 레이아웃
        ok_btn_layout = QHBoxLayout()
        ok_btn_layout.addStretch(1)
        ok_btn_layout.addWidget(self.ok_btn)
        ok_btn_layout.addStretch(1)
        
        # 상태바와 진행율 텍스트를 위한 레이아웃
        progress_layout = QHBoxLayout()
        self.progress_label = QLabel('진행율:', self)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        
        # 상태바 레이아웃에 위젯 추가
        progress_layout.addWidget(self.progress_label)
        progress_layout.addWidget(self.progress_bar)

        # 콘솔 출력용 텍스트 영역 추가
        self.console_output = QTextEdit(self)
        self.console_output.setReadOnly(True)

        # 메인 레이아웃에 위젯 추가
        main_layout.addLayout(file_layout)
        main_layout.addLayout(param_layout)
        main_layout.addLayout(ok_btn_layout)
        main_layout.addLayout(progress_layout)
        main_layout.addWidget(self.console_output)
    
        # 레이아웃 설정
        self.setLayout(main_layout)

        # 윈도우 설정
        self.setWindowTitle('BOM Cost 분석')
        self.resize(600, 300)
        self.center()
        self.show()
        self.activateWindow()
        self.raise_()
        
        # 스타일 시트 설정
        self.setStyleSheet("""
            QWidget {
                background-color: #333333;
                color: #FFFFFF;
            }
            QPushButton {
                background-color: #666666;
                color: #FFFFFF;
            }
            QLineEdit {
                background-color: #333333;
                color: #FFFFFF;
            }
            QProgressBar {
                background-color: #333333;
                color: #FFFFFF;
            }
        """)

        # 표준 출력을 텍스트 영역으로 리디렉션
        sys.stdout = self
        sys.stderr = self

    def center(self):
        # 화면 중앙에 윈도우 배치
        qr = self.frameGeometry()
        cp = QApplication.desktop().screen().rect().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def showFileDialog(self):
        # 파일 다이얼로그 열기
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "파일 선택", "", "All Files (*);;Python Files (*.py)", options=options)
        if file_name:
            base_name = os.path.basename(file_name)  # 파일명만 추출
            self.file_label.setText(f'선택된 파일: {base_name}')
        else:
            QMessageBox.warning(self, '경고', 'BOM Cost 엑셀(xlsm) 파일을 선택해 주세요.')

    def validateAndRunProgram(self):
        # 파일 선택 여부 확인
        file_name = self.file_label.text().replace('선택된 파일: ', '')
        if file_name == 'BOM Cost 파일(xlsm)을 선택해 주세요!':
            QMessageBox.warning(self, '경고', 'BOM Cost 엑셀(xlsm) 파일을 선택해 주세요.')
            return

        # 파라미터 값 검증 및 프로그램 실행
        param = self.param_input.text()
        try:
            param_value = int(param)
            if 1 <= param_value <= 100:
                self.runProgram(param_value)
            else:
                raise ValueError
        except ValueError:
            QMessageBox.warning(self, '경고', '1~100까지의 정수로만 입력하세요.')
            self.param_input.clear()
            
    def runProgram(self, param_value):
        # OK 버튼 클릭 시 실행할 동작
        file_name = self.file_label.text().replace('선택된 파일: ', '')
        param = self.param_input.text()
        if file_name == '1.파일 선택: BOM Cost 파일(xlsm)을 선택해 주세요!' or not param:
            QMessageBox.warning(self, '경고', '파일 선택과 정렬할 숫자 모두 입력해주세요.')
        else:
            dataframe.set_param_value(param_value)
            self.progress_bar.setValue(0)
            self.worker = Worker(file_name, int(param))
            self.worker.progress.connect(self.updateProgressBar)
            self.worker.finished.connect(self.finishProgram)
            self.worker.start()

    def updateProgressBar(self, value):
        self.progress_bar.setValue(value)

    def finishProgram(self, elapsed_time):
        QMessageBox.information(self, '정보', f'경과 시간: {elapsed_time:.2f}초, 정리가 완료되었습니다.')
        self.progress_bar.setValue(100)
        # QApplication.quit()  # 프로그램 종료
    
    def gui_kill(self):
        QApplication.quit()  # 프로그램 종료

    def write(self, text):
        if text.strip() and not text.startswith('d:\\'):  # 빈 줄은 무시하고, 경고 메시지 제외
            # QTextEdit에 텍스트를 추가하고 스크롤을 최하단으로 이동
            self.console_output.append(text.strip())
            self.console_output.verticalScrollBar().setValue(self.console_output.verticalScrollBar().maximum())

    def flush(self):
        pass  # 표준 출력 리디렉션을 위해 필요하지만 아무 작업도 하지 않음

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
