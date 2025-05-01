# app.py
import os
import tempfile
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB制限
app.secret_key = os.urandom(24)  # セッション用のシークレットキー

# 許可するファイル拡張子
ALLOWED_EXTENSIONS = {'xls', 'xlsx', 'xlsm'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_files():
    # ファイルの存在チェック
    if 'excel_file' not in request.files:
        return render_template('error.html', message="ファイルが選択されていません"), 400
    
    excel_file = request.files['excel_file']
    
    # ファイル名の空チェック
    if excel_file.filename == '':
        return render_template('error.html', message="ファイル名が不正です"), 400
    
    # ファイル形式チェック
    if not (excel_file and allowed_file(excel_file.filename)):
        return render_template('error.html', message="許可されていないファイル形式です"), 400
    
    try:
        # 安全なファイル名で保存
        filename = secure_filename(excel_file.filename)
        save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        excel_file.save(save_path)

        # ★ここに実際の処理ロジックを追加★
        # 例: Pandasを使ったExcel処理など
        processed_file_path = save_path  # 現状は同じファイルを返す

        # 処理済みファイルをダウンロード
        return send_file(
            processed_file_path,
            as_attachment=True,
            download_name='Processed_Result.xlsm',
            mimetype='application/vnd.ms-excel'
        )

    except Exception as e:
        app.logger.error(f"処理エラー: {str(e)}")
        return render_template('error.html', message=f"処理エラーが発生しました: {str(e)}"), 500

@app.errorhandler(413)
def request_entity_too_large(error):
    return render_template('error.html', message="ファイルサイズが大きすぎます（最大16MB）"), 413

if __name__ == '__main__':
    # 本番環境用設定
    port = int(os.environ.get('PORT', 5000))
    host = os.environ.get('HOST', '0.0.0.0')
    app.run(host=host, port=port, debug=os.environ.get('DEBUG', False))
