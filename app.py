import io
import re
import os
import traceback
import zipfile
from typing import List, Dict, Any
from flask import (
    Flask,
    request,
    render_template,
    send_file,
    send_from_directory,
    jsonify,
    url_for,
)
import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
from win32com import client as win32client  # ExcelからPDF変換用

app = Flask(__name__, template_folder='templates', static_folder='static')

# --- 定数 ---
TEMPLATE_FILE_PATH = "template.xlsm"  # app.py と同じディレクトリにあると仮定

# --- PDF抽出関連の関数 ---
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
        columns.append(" ".join(col_words))
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
            cols = split_line_using_boundaries(sorted_group, boundaries)
        else:
            cols = [" ".join(w['text'] for w in sorted_group)]
        result_rows.append(cols)
    return result_rows

def remove_extra_empty_columns(rows: List[List[str]]) -> List[List[str]]:
    if not rows:
        return rows
    num_cols = max(len(r) for r in rows)
    keep = [i for i in range(num_cols) if any(i < len(r) and r[i].strip() for r in rows)]
    return [[r[i] if i < len(r) else "" for i in keep] for r in rows]

# --- ExcelからPDF変換関数 ---
def excel_to_pdf(excel_data, sheet_name):
    # 一時ファイルとして保存
    temp_excel_path = "temp_output.xlsm"
    with open(temp_excel_path, "wb") as f:
        f.write(excel_data)
    
    # ExcelからPDF変換
    excel = win32client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        workbook = excel.Workbooks.Open(os.path.abspath(temp_excel_path))
        sheet = workbook.Worksheets(sheet_name)
        temp_pdf_path = "temp_output.pdf"
        sheet.ExportAsFixedFormat(0, os.path.abspath(temp_pdf_path))
        workbook.Close(False)
        
        # PDFデータを読み込み
        with open(temp_pdf_path, "rb") as f:
            pdf_data = f.read()
        return pdf_data
    finally:
        excel.Quit()
        # 一時ファイル削除
        if os.path.exists(temp_excel_path):
            os.remove(temp_excel_path)
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)

# --- Excel書き込み関連 ---
thin_border_side = Side(border_style="thin", color="000000")
thin_border = Border(left=thin_border_side, right=thin_border_side,
                     top=thin_border_side, bottom=thin_border_side)

def append_pdf_to_template(pdf_stream, excel_stream):
    wb = load_workbook(TEMPLATE_FILE_PATH, keep_vba=True)
    if hasattr(wb, "calculation_properties"):
        wb.calculation_properties.calcMode = 'auto'
        wb.calculation_properties.fullCalcOnLoad = True

    # アップロードExcelをテンプレートに反映
    wb_up = load_workbook(excel_stream, data_only=True)
    ws_up = wb_up.active
    ws_tpl = wb[wb.sheetnames[0]]
    for r_idx, row in enumerate(ws_up.iter_rows(values_only=True), start=1):
        for c_idx, v in enumerate(row, start=1):
            ws_tpl.cell(row=r_idx, column=c_idx, value=v)
    for col, dim in ws_up.column_dimensions.items():
        if dim.width:
            ws_tpl.column_dimensions[col].width = dim.width

    # PDF抽出 → 新規シートに書き込み
    with pdfplumber.open(pdf_stream) as pdf:
        for idx, page in enumerate(pdf.pages, start=1):
            rows = extract_text_with_layout(page)
            rows = [r for r in rows if any(c.strip() for c in r)]
            rows = remove_extra_empty_columns(rows)
            if not rows:
                continue
            name = f"Page_{idx}"
            base = name; c=1
            while name in wb.sheetnames:
                name = f"{base}_{c}"; c+=1
            ws = wb.create_sheet(title=name)
            max_cols = max(len(r) for r in rows)
            for r_i, row in enumerate(rows, start=1):
                for c_i, val in enumerate(row, start=1):
                    cell = ws.cell(row=r_i, column=c_i, value=val)
                    cell.border = thin_border
            for col_i in range(1, max_cols+1):
                ws.column_dimensions[get_column_letter(col_i)].width = 12

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()

# --- service-worker.js 配信ルート ---
@app.route('/service-worker.js')
def sw():
    return send_from_directory(app.static_folder, 'service-worker.js')

# --- ルート定義 ---
@app.route('/')
def index():
    favicon_url = url_for('static', filename='icon.png')
    return render_template('index.html', favicon_url=favicon_url)

@app.route('/process', methods=['POST'])
def process_files():
    if 'pdf_file' not in request.files or 'excel_file' not in request.files:
        return jsonify({"error": "PDFファイルとExcelファイルの両方が必要です。"}), 400
    pdf = request.files['pdf_file']
    xlsx = request.files['excel_file']
    if pdf.filename == '' or xlsx.filename == '':
        return jsonify({"error": "ファイルが選択されていません。"}), 400
    if not pdf.filename.lower().endswith('.pdf'):
        return jsonify({"error": "PDFファイルをアップロードしてください。"}), 400
    if not xlsx.filename.lower().endswith(('.xls', '.xlsx')):
        return jsonify({"error": "Excelファイル (.xls または .xlsx) をアップロードしてください。"}), 400
    try:
        # Excelファイルを処理
        excel_data = append_pdf_to_template(pdf.stream, xlsx.stream)
        base = os.path.splitext(xlsx.filename)[0]
        excel_name = f"Combined_{base}.xlsm"
        
        # 「提出用」シートをPDFに変換
        pdf_data = excel_to_pdf(excel_data, "提出用")
        pdf_name = f"Combined_{base}_提出用.pdf"
        
        # ZIPファイルにまとめてダウンロード
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            zip_file.writestr(excel_name, excel_data)
            zip_file.writestr(pdf_name, pdf_data)
        
        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            download_name=f"Combined_{base}.zip",
            as_attachment=True
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"ファイル処理中にエラーが発生しました: {e}"}), 500

if __name__ == '__main__':
    app.run(debug=True)
