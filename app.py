from flask import Flask, request, render_template, jsonify, send_from_directory, url_for
import io, os
from datetime import datetime
import webbrowser
from threading import Timer
import dataframe, file_save

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Flask의 세션을 위한 비밀키 설정

# 그림 파일을 저장할 디렉토리 설정
IMAGE_FOLDER = '.'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    try:
        param_value = request.form.get('param_value')
        file_data = request.files.get('file_data')

        if file_data is None:
            return jsonify({'result': 'No file uploaded'}), 400

        param_value = int(param_value)

        # 파일을 메모리에서 처리
        file_stream = io.BytesIO(file_data.read())
        dataframe.set_param_value(param_value)
        
        # 파일 스트림을 데이터프레임으로 읽기 및 파일 저장
        dataframe.dataframe_read(file_stream) # 정렬 갯수 파일
        file_save.tbd_cost(file_stream) # 미정단가 정리
        
        # 저장된 이미지 불러오기
        today_date = datetime.today().strftime('%Y%m%d')
        image_filename = f"{dataframe.model_name}_BOM_Cost_Graph_{today_date}.png"
        image_path = os.path.join(IMAGE_FOLDER, image_filename)

        # 이미지 파일이 실제로 존재하는지 확인
        if not os.path.isfile(image_path):
            return jsonify({'result': 'Error', 'message': 'File not saved'}), 500

        return jsonify({
            'result': 'Success',
            'html_table': dataframe.html_table,
            'html_table_tbd': file_save.html_table_tbd,
            'image_url': url_for('uploaded_file', filename=image_filename)
        })
    except Exception as e:
        import traceback
        traceback.print_exc()  # 콘솔에 자세한 오류 출력
        return jsonify({'result': f'Error processing the request: {str(e)}'}), 500

@app.route('/images/<filename>')
def uploaded_file(filename):
    return send_from_directory(IMAGE_FOLDER, filename)

def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000/")

if __name__ == '__main__':
    # 애플리케이션이 시작된 후 브라우저를 여는 방식으로 수정
    Timer(1, open_browser).start()
    app.run(debug=True, use_reloader=False)  # use_reloader=False로 설정하여 브라우저 중복 열림 방지
