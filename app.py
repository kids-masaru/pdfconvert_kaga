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
from openpyxl.styles import Border, Side
from werkzeug.utils import secure_filename

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
    """Excelデータをテンプレートに転送する（完全修正版）"""
    # テンプレートをマクロ有効で読み込み
    wb = load_workbook(TEMPLATE_FILE_PATH, keep_vba=True)
    
    # アップロードされたExcelを読み込み
    upload_wb = load_workbook(upload_stream, data_only=True)
    upload_ws = upload_wb.active
    template_ws = wb[wb.sheetnames[0]]
    
    # 既存データのクリア
    template_ws.delete_rows(1, template_ws.max_row)
    
    # データ転送処理
    for r_idx, row in enumerate(upload_ws.iter_rows(values_only=True), start=1):
        for c_idx, value in enumerate(row, start=1):
            template_ws.cell(row=r_idx, column=c_idx, value=value)
    
    # 列幅の転送
    for col_letter, dim in upload_ws.column_dimensions.items():
        template_col = template_ws.column_dimensions[col_letter]
        template_col.width = dim.width
        if dim.hidden:
            template_col.hidden = True
    
    # スタイルの転送
    for row in upload_ws.iter_rows():
        for cell in row:
            template_cell = template_ws.cell(row=cell.row, column=cell.column)
            template_cell.font = cell.font.copy()
            template_cell.border = cell.border.copy()
            template_cell.fill = cell.fill.copy()
            template_cell.number_format = cell.number_format
            template_cell.alignment = cell.alignment.copy()
    
    # メモリに保存
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- ルート定義（以下変更なし）---
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
        if 'excel_file' not in request.files:
            return render_template('error.html', message="Excelファイルを選択してください"), 400
            
        excel_file = request.files['excel_file']
        
        if excel_file.filename == '':
            return render_template('error.html', message="ファイルが選択されていません"), 400
            
        if not allowed_file(excel_file.filename):
            return render_template('error.html', message="Excelファイル(xls/xlsx)のみアップロード可能です"), 400
        
        # ファイル名加工処理
        original_name = os.path.splitext(excel_file.filename)[0]
        
        # 正規表現で「○○保育園」パターンを検索（例: 東京保育園 または 大阪保育園）
        import re
        pattern = r'([^_]+保育園)'  # アンダースコアを含まない「○○保育園」を抽出
        match = re.search(pattern, original_name)
        
        if match:
            school_name = match.group(1)
            new_name = f"{school_name}_Processed.xlsm"
        else:
            new_name = "Processed.xlsm"
        
        safe_name = secure_filename(new_name)
        
        processed_file = process_excel_template(excel_file.stream)
        
        return send_file(
            processed_file,
            mimetype='application/vnd.ms-excel.sheet.macroEnabled.12',
            download_name=safe_name,
            as_attachment=True
        )
        
    except Exception as e:
        app.logger.error(f"処理エラー: {traceback.format_exc()}")
        return render_template('error.html', message=f"処理エラー: {str(e)}"), 500

@app.route('/service-worker.js')
def service_worker():
    return send_from_directory(app.static_folder, 'service-worker.js')

@app.route('/manifest.json')  # 追加
def manifest():
    return send_from_directory(app.static_folder, 'manifest.json')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    host = os.environ.get('HOST', '0.0.0.0')
    app.run(host=host, port=port)
