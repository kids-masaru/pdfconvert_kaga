import io
import re
import os
from typing import List, Dict, Any
from flask import Flask, request, render_template, send_file, jsonify, url_for
import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
import traceback # エラーハンドリング用

# --- Flask アプリケーションの初期化 ---
# templates フォルダと static フォルダを Flask が認識するように設定
app = Flask(__name__, template_folder='templates', static_folder='static')

# --- 定数 ---
TEMPLATE_FILE_PATH = "template.xlsm" # app.py と同じディレクトリにあると仮定

# --- PDF抽出関連の関数 (元のコードから流用) ---
def is_number(text: str) -> bool:
    return bool(re.match(r'^\d+$', text.strip()))

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

def get_vertical_boundaries(page, tolerance: float = 2, min_gap: float = 20.0) -> List[float]:
    vertical_lines = []
    for line in page.lines:
        if abs(line['x0'] - line['x1']) < tolerance:
            vertical_lines.append((line['x0'] + line['x1']) / 2)
    vertical_lines = [round(x, 1) for x in vertical_lines]
    words = page.extract_words()
    if not words:
        return sorted(set(vertical_lines))
    left_boundary = round(min(word['x0'] for word in words), 1)
    right_boundary = round(max(word['x1'] for word in words), 1)
    boundaries = sorted(set([left_boundary] + vertical_lines + [right_boundary]))
    merged_boundaries = [boundaries[0]]
    for b in boundaries[1:]:
        if b - merged_boundaries[-1] >= min_gap:
            merged_boundaries.append(b)
    return merged_boundaries

def split_line_using_boundaries(sorted_words: List[Dict[str, Any]], boundaries: List[float]) -> List[str]:
    columns = []
    for i in range(len(boundaries) - 1):
        left = boundaries[i]
        right = boundaries[i+1]
        col_words = [word['text'] for word in sorted_words
                     if (word['x0'] + word['x1'])/2 >= left and (word['x0'] + word['x1'])/2 < right]
        cell_text = " ".join(col_words)
        columns.append(cell_text)
    return columns

def extract_text_with_layout(page) -> List[List[str]]:
    words = page.extract_words(x_tolerance=5, y_tolerance=5)
    if not words:
        return []
    boundaries = get_vertical_boundaries(page)
    row_groups = get_line_groups(words, y_tolerance=5)
    result_rows = []
    for group in row_groups:
        sorted_group = sorted(group, key=lambda w: w['x0'])
        if boundaries:
            columns = split_line_using_boundaries(sorted_group, boundaries)
        else:
            columns = [" ".join(word['text'] for word in sorted_group)]
        result_rows.append(columns)
    return result_rows

def remove_extra_empty_columns(rows: List[List[str]]) -> List[List[str]]:
    if not rows:
        return rows
    num_cols = max(len(row) for row in rows) if rows else 0 # 行がない場合の考慮
    if num_cols == 0:
        return rows
    keep_indices = []
    for col in range(num_cols):
        if any(col < len(row) and row[col].strip() for row in rows): # インデックス範囲チェックを追加
            keep_indices.append(col)
    new_rows = []
    for row in rows:
        new_row = [row[i] if i < len(row) else "" for i in keep_indices]
        new_rows.append(new_row)
    return new_rows

# --- Excelへの書き込み関数 (元のコードから流用) ---
thin_border_side = Side(border_style="thin", color="000000")
thin_border = Border(
    left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side
)

def append_pdf_to_template(pdf_file_stream, excel_file_stream):
    """
    PDFとExcelのファイルストリームを受け取り、処理して結合したExcelデータを返す関数
    """
    # テンプレートファイルを読み込む (VBAマクロ保持)
    wb = load_workbook(TEMPLATE_FILE_PATH, keep_vba=True)

    # 計算モードを自動に設定（必要に応じて）
    if hasattr(wb, "calculation_properties"):
        wb.calculation_properties.calcMode = 'auto'
        wb.calculation_properties.fullCalcOnLoad = True

    # --- アップロードされたExcelデータをテンプレートの最初のシートにコピー ---
    if excel_file_stream:
        wb_uploaded = load_workbook(excel_file_stream, data_only=True) # data_only=Trueで数式ではなく値を取得
        ws_uploaded = wb_uploaded.active
        ws_template = wb[wb.sheetnames[0]] # テンプレートの最初のシートを取得

        # 既存のテンプレートシートの内容をクリア (オプション)
        # for row in ws_template.iter_rows():
        #     for cell in row:
        #         cell.value = None

        # アップロードされたExcelの内容をコピー
        for row_idx, row in enumerate(ws_uploaded.iter_rows(values_only=True), start=1):
            for col_idx, value in enumerate(row, start=1):
                ws_template.cell(row=row_idx, column=col_idx, value=value)

        # 列幅をコピー (元のExcelの書式を維持するため)
        for col_letter, dim in ws_uploaded.column_dimensions.items():
            if dim.width: # 幅が設定されている場合のみコピー
                 ws_template.column_dimensions[col_letter].width = dim.width

    # --- PDFからデータを抽出し、新しいシートに追加 ---
    with pdfplumber.open(pdf_file_stream) as pdf:
        for idx, page in enumerate(pdf.pages, start=1):
            rows = extract_text_with_layout(page)
            rows = [row for row in rows if any(cell.strip() for cell in row)] # 空行を除去
            if not rows:
                continue # 抽出データがなければスキップ

            rows = remove_extra_empty_columns(rows) # 空の列を除去
            max_cols = max(len(row) for row in rows) if rows else 0

            # 新しいシートを作成
            sheet_name = f"Page_{idx}"
            # 同じ名前のシートが既に存在する場合の処理 (例: 末尾に番号を追加)
            original_sheet_name = sheet_name
            counter = 1
            while sheet_name in wb.sheetnames:
                sheet_name = f"{original_sheet_name}_{counter}"
                counter += 1
            ws = wb.create_sheet(title=sheet_name)

            # データを書き込み、罫線を設定
            for r_idx, row in enumerate(rows, start=1):
                for c_idx, cell_value in enumerate(row, start=1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=cell_value)
                    cell.border = thin_border # 罫線を適用

            # 列幅を設定 (デフォルト値)
            for col in range(1, max_cols + 1):
                col_letter = get_column_letter(col)
                ws.column_dimensions[col_letter].width = 12 # デフォルト幅

    # --- 処理結果をメモリ上のバイナリデータとして保存 ---
    output = io.BytesIO()
    wb.save(output)
    output.seek(0) # ストリームの先頭に戻す
    return output.read() # バイナリデータを返す

# --- Flask ルート定義 ---

@app.route('/')
def index():
    """
    メインページを表示するルート
    """
    # index.html をレンダリングして返す
    # ファビコンのパスを渡す (url_forを使用)
    favicon_url = url_for('static', filename='icon.png')
    return render_template('index.html', favicon_url=favicon_url)

@app.route('/process', methods=['POST'])
def process_files():
    """
    ファイルアップロードを受け付け、処理を実行し、結果を返すAPIエンドポイント
    """
    if 'pdf_file' not in request.files or 'excel_file' not in request.files:
        return jsonify({"error": "PDFファイルとExcelファイルの両方が必要です。"}), 400

    pdf_file = request.files['pdf_file']
    excel_file = request.files['excel_file']

    if pdf_file.filename == '' or excel_file.filename == '':
        return jsonify({"error": "ファイルが選択されていません。"}), 400

    # ファイル拡張子のチェック (念のため)
    if not pdf_file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "PDFファイルをアップロードしてください。"}), 400
    if not excel_file.filename.lower().endswith(('.xls', '.xlsx')):
         return jsonify({"error": "Excelファイル (.xls または .xlsx) をアップロードしてください。"}), 400

    try:
        # ファイルストリームを append_pdf_to_template 関数に渡す
        # request.files[key] は FileStorage オブジェクトで、ファイルライクなオブジェクトとして扱える
        combined_excel_data = append_pdf_to_template(pdf_file.stream, excel_file.stream)

        # 元のExcelファイル名から拡張子を除いた部分を取得
        excel_base_filename = os.path.splitext(excel_file.filename)[0]
        output_filename = f"Combined_{excel_base_filename}.xlsm"

        # send_file を使ってファイルをダウンロードさせる
        return send_file(
            io.BytesIO(combined_excel_data), # BytesIOでラップ
            mimetype='application/vnd.ms-excel.sheet.macroEnabled.12', # .xlsm の MIME タイプ
            download_name=output_filename, # ダウンロード時のファイル名
            as_attachment=True # 添付ファイルとして扱う
        )

    except Exception as e:
        # エラーハンドリング: 詳細なエラーログを出力し、ユーザーには汎用的なメッセージを返す
        print("Error during file processing:")
        traceback.print_exc() # コンソールにスタックトレースを出力
        return jsonify({"error": f"ファイルの処理中にエラーが発生しました: {e}"}), 500

# --- アプリケーションの実行 ---
if __name__ == '__main__':
    # debug=True にすると、コード変更時に自動リロードされ、デバッグ情報が表示される
    # 本番環境では debug=False にする
    app.run(debug=True)
