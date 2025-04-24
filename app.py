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
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table
from reportlab.lib import colors
from reportlab.platypus import TableStyle

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
    merged = [boundaries[0]]
    for b in boundaries[1:]:
        if b - merged[-1] >= min_gap:
            merged.append(b)
    return merged

def split_line_using_boundaries(sorted_words: List[Dict[str, Any]], boundaries: List[float]) -> List[str]:
    cols = []
    for i in range(len(boundaries) - 1):
        left, right = boundaries[i], boundaries[i+1]
        texts = [w['text'] for w in sorted_words
                 if (w['x0'] + w['x1'])/2 >= left and (w['x0'] + w['x1'])/2 < right]
        cols.append(" ".join(texts))
    return cols

def extract_text_with_layout(page) -> List[List[str]]:
    words = page.extract_words(x_tolerance=5, y_tolerance=5)
    if not words:
        return []
    boundaries = get_vertical_boundaries(page)
    groups = get_line_groups(words)
    rows = []
    for grp in groups:
        sorted_grp = sorted(grp, key=lambda w: w['x0'])
        if boundaries:
            cols = split_line_using_boundaries(sorted_grp, boundaries)
        else:
            cols = [" ".join(w['text'] for w in sorted_grp)]
        rows.append(cols)
    return rows

def remove_extra_empty_columns(rows: List[List[str]]) -> List[List[str]]:
    if not rows:
        return rows
    num_cols = max(len(r) for r in rows)
    keep = [i for i in range(num_cols) if any(i < len(r) and r[i].strip() for r in rows)]
    return [[r[i] if i < len(r) else "" for i in keep] for r in rows]

# --- Excel書き込み関連 ---
thin = Side(border_style="thin", color="000000")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def append_pdf_to_template(pdf_stream, excel_stream) -> bytes:
    wb = load_workbook(TEMPLATE_FILE_PATH, keep_vba=True)
    if hasattr(wb, 'calculation_properties'):
        wb.calculation_properties.calcMode = 'auto'
        wb.calculation_properties.fullCalcOnLoad = True
    # Excel反映
    up = load_workbook(excel_stream, data_only=True)
    ws_up, ws_tpl = up.active, wb[wb.sheetnames[0]]
    for r, row in enumerate(ws_up.iter_rows(values_only=True), start=1):
        for c, v in enumerate(row, start=1):
            ws_tpl.cell(row=r, column=c, value=v)
    for col, dim in ws_up.column_dimensions.items():
        if dim.width:
            ws_tpl.column_dimensions[col].width = dim.width
    # PDF抽出
    with pdfplumber.open(pdf_stream) as pdf:
        for idx, page in enumerate(pdf.pages, start=1):
            rows = extract_text_with_layout(page)
            rows = [r for r in rows if any(c.strip() for c in r)]
            rows = remove_extra_empty_columns(rows)
            if not rows:
                continue
            name, base, cnt = f'Page_{idx}', f'Page_{idx}', 1
            while name in wb.sheetnames:
                name = f'{base}_{cnt}'; cnt += 1
            ws = wb.create_sheet(title=name)
            maxc = max(len(r) for r in rows)
            for r_i, row in enumerate(rows, 1):
                for c_i, val in enumerate(row, 1):
                    cell = ws.cell(row=r_i, column=c_i, value=val)
                    cell.border = border
            for c_i in range(1, maxc+1):
                ws.column_dimensions[get_column_letter(c_i)].width = 12
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()

# PDF作成 & ZIP化 ルート
@app.route('/process', methods=['POST'])
def process_files():
    try:
        pdf_f = request.files['pdf_file']; xlsx_f = request.files['excel_file']
        # 既存処理
        combined = append_pdf_to_template(pdf_f.stream, xlsx_f.stream)
        # Excel書き出し名
        base = os.path.splitext(xlsx_f.filename)[0]
        excel_name = f'Combined_{base}.xlsm'
        # 提出用シートをPDF化
        wb2 = load_workbook(io.BytesIO(combined), data_only=True)
        if '提出用' not in wb2.sheetnames:
            raise Exception('提出用シートが見つかりません')
        ws_pdf = wb2['提出用']
        data = [[cell.value or '' for cell in row] for row in ws_pdf.iter_rows()]
        pdf_buf = io.BytesIO()
        doc = SimpleDocTemplate(pdf_buf, pagesize=A4)
        tbl = Table(data)
        tbl.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold')
        ]))
        doc.build([tbl])
        pdf_buf.seek(0)
        pdf_name = f'{base}_提出用.pdf'
        # ZIPにまとめる
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w') as zf:
            zf.writestr(excel_name, combined)
            zf.writestr(pdf_name, pdf_buf.getvalue())
        zip_buf.seek(0)
        # 送信
        return send_file(
            zip_buf,
            mimetype='application/zip',
            download_name=f'{base}_results.zip',
            as_attachment=True
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

# その他のルートはそのまま
@app.route('/')
def index():
    return render_template('index.html', favicon_url=url_for('static', filename='icon.png'))

@app.route('/service-worker.js')
def sw():
    return send_from_directory(app.static_folder, 'service-worker.js')

if __name__ == '__main__':
    app.run(debug=True)
