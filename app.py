import io
import os
import re
import traceback
from typing import List, Dict, Any
from flask import (
    Flask,
    request,
    render_template,
    send_file,
    jsonify,
    url_for
)
import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__, template_folder='templates', static_folder='static')

# --- 定数設定 ---
TEMPLATE_FILE_PATH = "template.xlsm"
ALLOWED_EXCEL_EXTENSIONS = {'xls', 'xlsx'}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB

# --- 初期設定 ---
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH
thin_border = Border(
    left=Side(border_style="thin", color="000000"),
    right=Side(border_style="thin", color="000000"),
    top=Side(border_style="thin", color="000000"),
    bottom=Side(border_style="thin", color="000000")
)

# --- ヘルパー関数 ---
def allowed_file(filename: str, allowed_extensions: set) -> bool:
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in allowed_extensions

def is_number(text: str) -> bool:
    return bool(re.match(r'^\d+$', text.strip()))

# --- PDF処理関数 ---
def get_line_groups(words: List[Dict[str, Any]], y_tolerance: float = 5) -> List[List[Dict[str, Any]]]:
    if not words:
        return []
    sorted_words = sorted(words, key=lambda w: w['top'])
    groups = []
    current_group = [sorted_words[0]]
    current_top = sorted_words[0]['top']
    for word in sorted_words[1:]:
        if abs(word['top'] - current_top) <= y_tolerance:
            current_group.append(word)
        else:
            groups.append(current_group)
            current_group = [word]
            current_top = word['top']
    groups.append(current_group)
    return groups

def process_pdf_to_excel(pdf_stream, excel_stream) -> io.BytesIO:
    # テンプレート読み込み（VBA保持）
    wb = load_workbook(
        filename=TEMPLATE_FILE_PATH,
        keep_vba=True
    )
    
    # アップロードExcelをテンプレートに反映
    upload_wb = load_workbook(excel_stream, data_only=True)
    upload_ws = upload_wb.active
    template_ws = wb[wb.sheetnames[0]]
    
    # データ転送
    for row in upload_ws.iter_rows(values_only=True):
        template_ws.append(row)
    
    # PDF解析処理
    with pdfplumber.open(pdf_stream) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            words = page.extract_words(x_tolerance=5, y_tolerance=5)
            rows = extract_text_with_layout(page)
            # ...（PDF解析ロジックは既存コードを流用）...
            
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
        os.path.join(app.root_path, 'static'),
        'icon.png',
        mimetype='image/vnd.microsoft.icon'
    )
@app.route('/process', methods=['POST'])
def handle_processing():
    try:
        # ファイルチェック
        if 'pdf_file' not in request.files or 'excel_file' not in request.files:
            return render_template('error.html', message="PDFとExcelファイルの両方が必要です"), 400
            
        pdf_file = request.files['pdf_file']
        excel_file = request.files['excel_file']
        
        # ファイル検証
        if not all([
            pdf_file.filename != '',
            excel_file.filename != '',
            allowed_file(pdf_file.filename, {'pdf'}),
            allowed_file(excel_file.filename, ALLOWED_EXCEL_EXTENSIONS)
        ]):
            return render_template('error.html', message="無効なファイル形式です"), 400
        
        # ファイル処理
        processed_file = process_pdf_to_excel(
            pdf_file.stream,
            excel_file.stream
        )
        
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
