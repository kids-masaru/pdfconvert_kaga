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
from werkzeug.utils import secure_filename
import copy

# PyPDF2はテキストベースのPDFからの抽出に利用
try:
    from PyPDF2 import PdfReader # PyPDF2 3.0.0以降の推奨
except ImportError:
    from PyPDF2 import PdfFileReader # 以前のバージョン向け

app = Flask(__name__, template_folder='templates', static_folder='static')

# --- 定数設定 ---
TEMPLATE_FILE_PATH = 'template.xlsm'  # テンプレートファイルのパス
ALLOWED_EXTENSIONS = {'xls', 'xlsx', 'xlsm', 'pdf'} # xlsmも追加
MAX_CONTENT_LENGTH = 32 * 1024 * 1024  # 32MBに増量（PDFサイズ考慮）

# --- 初期設定 ---
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

def allowed_file(filename: str) -> bool:
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def copy_worksheet_data(source_sheet, target_sheet, start_row=1, start_col=1):
    """
    ソースシートからターゲットシートにデータをコピーする関数
    既存のデータがある場合は指定した位置から貼り付ける
    """
    # ソースシートの使用範囲を取得
    if source_sheet.max_row == 1 and source_sheet.max_column == 1:
        # 空のシートの場合はスキップ
        return
    
    # データをコピー（値のみ、書式は保持しない）
    for row in range(1, source_sheet.max_row + 1):
        for col in range(1, source_sheet.max_column + 1):
            source_cell = source_sheet.cell(row=row, column=col)
            if source_cell.value is not None:  # 値がある場合のみコピー
                target_cell = target_sheet.cell(
                    row=start_row + row - 1, 
                    column=start_col + col - 1
                )
                target_cell.value = source_cell.value

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
        
        # 2. テンプレートファイルの存在確認と読み込み
        if not os.path.exists(TEMPLATE_FILE_PATH):
            return render_template('error.html', message=f"テンプレートファイル '{TEMPLATE_FILE_PATH}' が見つかりません。"), 500
        
        # テンプレートファイルを読み込み（元のファイルは変更されない）
        template_workbook = load_workbook(TEMPLATE_FILE_PATH, keep_vba=True)
        
        # 3. アップロードされたExcelファイルの読み込み
        uploaded_excel_workbook = load_workbook(io.BytesIO(excel_file_storage.read()))
        
        # アップロードされたExcelの1ページ目を取得
        if len(uploaded_excel_workbook.worksheets) == 0:
            return render_template('error.html', message="アップロードされたExcelファイルにシートがありません。"), 400
        
        uploaded_first_sheet = uploaded_excel_workbook.worksheets[0]
        
        # 4. テンプレートの1枚目のシートにアップロードされたデータを貼り付け
        if len(template_workbook.worksheets) == 0:
            return render_template('error.html', message="テンプレートファイルにシートがありません。"), 500
        
        template_first_sheet = template_workbook.worksheets[0]
        
        # テンプレートの1枚目にデータを貼り付け
        # 既存のデータがある場合は、適切な位置から貼り付ける
        # ここでは既存データの下に貼り付けるか、上書きするかを決める必要があります
        # 要件に応じて調整してください
        
        # 既存データの最後の行を取得
        last_row = template_first_sheet.max_row if template_first_sheet.max_row > 1 else 1
        
        # 空の行があるかチェック
        is_empty_template = True
        for row in range(1, last_row + 1):
            for col in range(1, template_first_sheet.max_column + 1):
                if template_first_sheet.cell(row=row, column=col).value is not None:
                    is_empty_template = False
                    break
            if not is_empty_template:
                break
        
        # データを貼り付ける開始位置を決定
        if is_empty_template:
            start_row = 1  # テンプレートが空の場合は1行目から
        else:
            start_row = last_row + 2  # 既存データがある場合は2行空けて貼り付け
        
        copy_worksheet_data(uploaded_first_sheet, template_first_sheet, start_row, 1)

        # 5. PDFファイルの処理
        for pdf_file_storage in pdf_files_storage:
            pdf_bytes = pdf_file_storage.read()
            
            extracted_texts = []
            try:
                # テキストベースPDFとして抽出を試みる
                pdf_reader = PdfReader(io.BytesIO(pdf_bytes))
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    text = page.extract_text()
                    
                    # 文字化け対策：テキストが正常に抽出されているかチェック
                    if text and text.strip(): 
                        # 改行コードを統一し、不要な空白を削除
                        cleaned_text = text.replace('\r\n', '\n').replace('\r', '\n')
                        extracted_texts.append(cleaned_text)
                    else:
                        # テキストが抽出できなかった場合
                        extracted_texts.append(f"[ページ {page_num + 1}: テキストを抽出できませんでした]") 
                
            except Exception as e:
                # PDFの読み込みやテキスト抽出に失敗した場合
                print(f"PDF '{pdf_file_storage.filename}' のテキスト抽出エラー: {e}")
                return render_template('error.html', message=f"PDF '{pdf_file_storage.filename}' の読み込みまたはテキスト抽出に失敗しました。詳細: {e}"), 500

            if not extracted_texts:
                return render_template('error.html', message=f"PDF '{pdf_file_storage.filename}' から有効な内容を抽出できませんでした。"), 500

            # 抽出したテキストをテンプレートの新しいシートに貼り付ける
            for i, page_content in enumerate(extracted_texts):
                # シート名を生成（PDFファイル名 + ページ番号）
                base_filename = os.path.splitext(secure_filename(pdf_file_storage.filename))[0]
                sheet_name_base = base_filename[:20] if len(base_filename) > 20 else base_filename
                
                # シート名に使えない文字を安全な文字に変換
                safe_sheet_name_base = (sheet_name_base
                                      .replace('[', '').replace(']', '')
                                      .replace('*', '').replace('?', '')
                                      .replace(':', '').replace('/', '')
                                      .replace('\\', ''))

                # 最終的なシート名 (最大31文字制限)
                sheet_name = f"{safe_sheet_name_base}_P{i+1}"
                if len(sheet_name) > 31:
                    suffix = f"_P{i+1}"
                    sheet_name = sheet_name[:31 - len(suffix)] + suffix

                # 同名のシートが既に存在する場合は番号を付ける
                original_sheet_name = sheet_name
                counter = 1
                while sheet_name in [ws.title for ws in template_workbook.worksheets]:
                    sheet_name = f"{original_sheet_name}_{counter}"
                    if len(sheet_name) > 31:
                        # 31文字を超える場合は調整
                        base_part = original_sheet_name[:25]
                        sheet_name = f"{base_part}_{counter}"
                    counter += 1

                # 新しいシートを作成
                new_sheet = template_workbook.create_sheet(title=sheet_name)
                
                # 抽出したテキストをExcelのセルに貼り付け
                # 改行で分割してセルに配置
                rows = page_content.split('\n')
                for r_idx, row_text in enumerate(rows):
                    if row_text.strip():  # 空行はスキップ
                        # 長すぎるテキストは複数のセルに分割することも可能
                        cell_text = row_text.strip()
                        # Excelのセルの文字数制限（32,767文字）を考慮
                        if len(cell_text) > 32767:
                            cell_text = cell_text[:32767]
                        
                        new_sheet.cell(row=r_idx + 1, column=1, value=cell_text)
                
                # ヘッダー情報を追加（オプション）
                header_info = f"PDF: {pdf_file_storage.filename} - ページ {i+1}"
                new_sheet.cell(row=1, column=2, value=header_info)

        # 6. 処理済みファイルをメモリに保存
        output_excel_stream = io.BytesIO()
        
        # xlsmファイルとして保存（VBAマクロも保持）
        template_workbook.save(output_excel_stream)
        output_excel_stream.seek(0) # ストリームの先頭に戻す

        # 7. 処理済みExcelファイルをダウンロード用に返す
        download_filename = f'processed_template_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsm'
        
        return send_file(
            output_excel_stream,
            mimetype='application/vnd.ms-excel.sheet.macroEnabled.12',  # xlsmのMIMEタイプ
            as_attachment=True,
            download_name=download_filename
        )

    except Exception as e:
        # エラーログ出力
        print(f"An error occurred: {e}")
        traceback.print_exc() # 詳細なトレースバックを出力

        return render_template('error.html', message=f"ファイル処理中に予期せぬエラーが発生しました: {e}"), 500

if __name__ == '__main__':
    app.run(debug=True)