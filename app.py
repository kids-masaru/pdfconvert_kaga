import io
import os
import traceback
from flask import (
    Flask,
    request,
    render_template,
    send_file,
    send_from_directory,
    url_for
)
from openpyxl import load_workbook

app = Flask(__name__, template_folder='templates', static_folder='static')

# --- 定数設定 ---
TEMPLATE_FILE_PATH = "template.xlsm"
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB

# --- 初期設定 ---
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

def allowed_file(filename: str) -> bool:
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def process_excel_template(upload_stream):
    """Excelデータをテンプレートに転送する"""
    # テンプレートをマクロ有効で読み込み
    wb = load_workbook(TEMPLATE_FILE_PATH, keep_vba=True)
    
    # アップロードされたExcelを読み込み
    upload_wb = load_workbook(upload_stream, data_only=True)
    upload_ws = upload_wb.active
    template_ws = wb[wb.sheetnames[0]]
    
    # データ転送処理
    for row in upload_ws.iter_rows(values_only=True):
        template_ws.append(row)
    
    # 列幅の転送
    for col, dim in upload_ws.column_dimensions.items():
        if dim.width:
            template_ws.column_dimensions[col].width = dim.width
    
    # メモリに保存
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- ルート定義 ---
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/favicon.ico')
def favicon():
    return send_from_directory(
        os.path.join(app.static_folder, 'images'),
        'icon.png',
        mimetype='image/vnd.microsoft.icon'
    )

@app.route('/process', methods=['POST'])
def handle_processing():
    try:
        # 単一ファイルアップロードに変更
        if 'excel_file' not in request.files:
            return render_template('error.html', message="Excelファイルを選択してください"), 400
            
        excel_file = request.files['excel_file']
        
        # ファイル検証
        if excel_file.filename == '':
            return render_template('error.html', message="ファイルが選択されていません"), 400
            
        if not allowed_file(excel_file.filename):
            return render_template('error.html', message="Excelファイル(xls/xlsx)のみアップロード可能です"), 400
        
        # ファイル処理
        processed_file = process_excel_template(excel_file.stream)
        
        return send_file(
            processed_file,
            mimetype='application/vnd.ms-excel.sheet.macroEnabled.12',
            download_name='processed_result.xlsm',
            as_attachment=True
        )
        
    except Exception as e:
        app.logger.error(f"処理エラー: {traceback.format_exc()}")
        return render_template('error.html', message=f"処理エラー: {str(e)}"), 500

@app.route('/service-worker.js')
def service_worker():
    return send_from_directory(app.static_folder, 'service-worker.js')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    host = os.environ.get('HOST', '0.0.0.0')
    app.run(host=host, port=port)
