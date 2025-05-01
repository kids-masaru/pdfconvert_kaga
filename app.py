from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import os
import tempfile

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_files():
    excel_file = request.files.get('excel_file')
    if not excel_file:
        return 'Excelファイルがアップロードされていません。', 400

    # アップロードされたファイルを保存
    excel_filename = secure_filename(excel_file.filename)
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
    excel_file.save(excel_path)

    # 処理（今回はそのまま返す例）
    # 実際にはここでExcelファイルを処理・変換する
    result_path = excel_path  # 実際には処理後のファイルパスに変更

    return send_file(result_path, as_attachment=True, download_name='Processed_Result.xlsm')

if __name__ == '__main__':
    app.run(debug=True)
