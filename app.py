import io
import os
import traceback
from datetime import datetime
from flask import (
    Flask,
    request,
    render_template,
    send_file,
    url_for
)
from openpyxl import load_workbook
# from openpyxl.styles import Border, Side # 今回の要件では直接使用しないためコメントアウト
from werkzeug.utils import secure_filename

# PyPDF2はテキストベースのPDFからの抽出に利用
try:
    from PyPDF2 import PdfReader # PyPDF2 3.0.0以降の推奨
except ImportError:
    from PyPDF2 import PdfFileReader # 以前のバージョン向け

app = Flask(__name__, template_folder='templates', static_folder='static')

# --- 定数設定 ---
# TEMPLATE_FILE_PATHは今回の要件では不要
ALLOWED_EXTENSIONS = {'xls', 'xlsx', 'pdf'} # PDFも許可する
MAX_CONTENT_LENGTH = 32 * 1024 * 1024  # 32MBに増量（PDFサイズ考慮）

# --- 初期設定 ---
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

def allowed_file(filename: str) -> bool:
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# process_excel_template 関数は今回の要件では直接使用しないため削除またはコメントアウト
# 現在のコードのその部分も削除してください。

# --- ルーティング ---

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload_and_process', methods=['POST'])
def upload_and_process():
    try:
        # 1. ファイルの存在確認
        if 'excel_file' not in request.files or 'pdf_files' not in request.files:
            return render_template('error.html', message="ExcelファイルとPDFファイルの両方をアップロードしてください。"), 400

        excel_file_storage = request.files['excel_file']
        pdf_files_storage = request.files.getlist('pdf_files') # 複数PDFを受け取る

        if not excel_file_storage.filename:
            return render_template('error.html', message="Excelファイルが選択されていません。"), 400
        if not all(f.filename for f in pdf_files_storage):
            return render_template('error.html', message="PDFファイルが選択されていません。"), 400
        
        # 2. Excelファイルの読み込み（アップロードされたExcelをベースにする）
        # アップロードされたExcelの最初のシートがシフト表データとして扱われる想定
        excel_workbook = load_workbook(io.BytesIO(excel_file_storage.read()))

        # 3. PDFファイルの処理
        for pdf_file_storage in pdf_files_storage:
            pdf_bytes = pdf_file_storage.read()
            
            extracted_texts = []
            try:
                # テキストベースPDFとして抽出を試みる (OCRは不要)
                pdf_reader = PdfReader(io.BytesIO(pdf_bytes))
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    text = page.extract_text()
                    if text and text.strip(): # テキストが空でなければ追加
                        extracted_texts.append(text)
                    else:
                        # テキストが抽出できなかった場合（例: 空白ページ）、空のテキストとして追加
                        extracted_texts.append("") 
                
            except Exception as e:
                # PDFの読み込みやテキスト抽出に失敗した場合
                print(f"PDF '{pdf_file_storage.filename}' のテキスト抽出エラー: {e}")
                return render_template('error.html', message=f"PDF '{pdf_file_storage.filename}' の読み込みまたはテキスト抽出に失敗しました。詳細: {e}"), 500

            if not extracted_texts:
                return render_template('error.html', message=f"PDF '{pdf_file_storage.filename}' から有効な内容を抽出できませんでした。"), 500

            # 抽出したテキストをExcelの新しいシートに貼り付ける
            for i, page_content in enumerate(extracted_texts):
                # シート名を Page_1, Page_2 の形式で生成
                # ユニーク性を保つために、PDFファイル名の一部を含める
                base_filename = os.path.splitext(secure_filename(pdf_file_storage.filename))[0]
                sheet_name_base = base_filename[:20] if len(base_filename) > 20 else base_filename
                
                # シート名には使えない文字があるので、安全な名前に変換
                # Excelのシート名に使えない文字: \ / ? * [ ] :
                safe_sheet_name_base = sheet_name_base.replace('[', '').replace(']', '').replace('*', '').replace('?', '').replace(':', '').replace('/', '').replace('\\', '')

                # 最終的なシート名 (最大31文字)
                sheet_name = f"{safe_sheet_name_base}_Page_{i+1}"
                if len(sheet_name) > 31:
                    # 長すぎる場合は末尾を切り詰める (Page_X の部分は残すように調整)
                    suffix = f"_Page_{i+1}"
                    sheet_name = sheet_name[:31 - len(suffix)] + suffix


                new_sheet = excel_workbook.create_sheet(title=sheet_name)
                
                # 抽出したテキストをExcelのセルに貼り付け
                rows = page_content.split('\n')
                for r_idx, row_text in enumerate(rows):
                    if row_text.strip(): # 空行はスキップ
                        new_sheet.cell(row=r_idx + 1, column=1, value=row_text.strip())

        # 4. 処理済みExcelファイルをメモリに保存
        output_excel_stream = io.BytesIO()
        excel_workbook.save(output_excel_stream)
        output_excel_stream.seek(0) # ストリームの先頭に戻す

        # 5. 処理済みExcelファイルをダウンロード用に返す
        return send_file(
            output_excel_stream,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'processed_data_{datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx'
        )

    except Exception as e:
        # エラーログ出力
        print(f"An error occurred: {e}")
        traceback.print_exc() # 詳細なトレースバックを出力

        return render_template('error.html', message=f"ファイル処理中に予期せぬエラーが発生しました: {e}"), 500

if __name__ == '__main__':
    app.run(debug=True)
