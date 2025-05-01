import os
import tempfile
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from openpyxl.writer.excel import save_vba

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.secret_key = os.urandom(24)

ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_to_xlsm(input_path, output_path):
    """xlsxをxlsmに変換（ダミーマクロ付加）"""
    wb = load_workbook(input_path)
    
    # ダミーのVBAプロジェクト（実際にマクロが必要な場合は中身を実装）
    vba_hex = '000000'  # 空のプロジェクト
    
    # マクロ有効ブックとして保存
    save_vba(wb, vba_hex)
    wb.save(output_path)
    wb.close()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_files():
    if 'excel_file' not in request.files:
        return render_template('error.html', message="ファイルが選択されていません"), 400
    
    excel_file = request.files['excel_file']
    
    if excel_file.filename == '':
        return render_template('error.html', message="ファイル名が不正です"), 400
    
    if not (excel_file and allowed_file(excel_file.filename)):
        return render_template('error.html', message="許可されていないファイル形式です"), 400
    
    try:
        filename = secure_filename(excel_file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{filename}")
        excel_file.save(input_path)
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"processed_{filename}.xlsm")
        convert_to_xlsm(input_path, output_path)
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name='Processed_Result.xlsm',
            mimetype='application/vnd.ms-excel.sheet.macroEnabled.12'
        )
        
    except Exception as e:
        app.logger.error(f"処理エラー: {str(e)}")
        return render_template('error.html', message=f"処理エラー: {str(e)}"), 500
    finally:
        # 一時ファイル削除
        if os.path.exists(input_path):
            os.remove(input_path)
        if os.path.exists(output_path):
            os.remove(output_path)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    host = os.environ.get('HOST', '0.0.0.0')
    app.run(host=host, port=port)
